"""After a live run, read back from Microsoft Graph what the assistant really wrote, as evidence.

    python scripts/live_readback.py examples/live-run "graph-assistant run output file"

Reads the Excel table rows, the "Hold:" events, the Drafts it wrote and the attachment folders, and
writes them to the folder given (excel-rows.json, calendar.json, drafts.json, files.txt). Email
addresses are masked (a***@example) so nothing personal lands in the repo.
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from graph_assistant import auth  # noqa: E402
from graph_assistant.assistant import COLUMNS  # noqa: E402
from graph_assistant.config import load  # noqa: E402
from graph_assistant.graph import GraphClient  # noqa: E402

EMAIL = re.compile(r"([A-Za-z0-9._%+-])[A-Za-z0-9._%+-]*@([A-Za-z0-9.-]+\.[A-Za-z]{2,})")


def mask(obj):
    s = json.dumps(obj, indent=2, ensure_ascii=False)
    return EMAIL.sub(lambda m: f"{m.group(1)}***@{'example' if not m.group(2).endswith('.example') else m.group(2)}", s)


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "examples" / "live-run")
    out.mkdir(parents=True, exist_ok=True)
    cfg = load(ROOT / "config" / "assistant.yaml")
    token = auth.device_code_token(os.environ["GRAPH_CLIENT_ID"], os.environ.get("GRAPH_TENANT", "consumers"))
    g = GraphClient(requests.Session(), token, mailbox=cfg.mailbox, drive=cfg.drive)

    book = g.item_by_path(cfg.workbook_path)
    # The displayed text, as a person sees it in Excel Online: Excel turns ISO date text into real
    # dates, so the raw values come back as date serials (46356 = 2026-11-30).
    body = g.request("GET", f"{cfg.drive}/items/{book['id']}/workbook/tables/{cfg.table}/dataBodyRange",
                     params={"$select": "text,values"})
    rows = [dict(zip(COLUMNS, r)) for r in body["text"] if any(r)]
    raw = [dict(zip(COLUMNS, r)) for r in body["values"] if any(r)]
    (out / "excel-raw-values.json").write_text(mask(raw))
    (out / "excel-rows.json").write_text(mask(rows))

    now = datetime.now()
    evs = g.calendar_view(now.strftime("%Y-%m-%dT00:00:00"), (now + timedelta(days=21)).strftime("%Y-%m-%dT23:59:59"),
                          cfg.timezone)
    holds = []
    for e in evs:
        if e.get("subject", "").startswith("Hold:"):
            full = g.request("GET", f"{cfg.mailbox}/events/{e['id']}",
                             params={"$select": "subject,start,end,showAs,attendees,categories,transactionId"})
            full = {k: v for k, v in full.items() if "@odata" not in k and k != "id"}
            holds.append(full)
    (out / "calendar.json").write_text(mask(holds))

    drafts = g.request("GET", f"{cfg.mailbox}/mailFolders/drafts/messages",
                       params={"$select": "subject,isDraft,body,toRecipients", "$top": "20"},
                       headers={"Prefer": 'outlook.body-content-type="text"'})["value"]
    drafts = [{"subject": d["subject"], "isDraft": d["isDraft"], "to": d["toRecipients"],
               "body": {"content": d["body"]["content"].split("\n________________________________")[0].strip()}}
              for d in drafts if d["subject"].startswith("RE:")]
    (out / "drafts.json").write_text(mask(drafts))

    files = []

    def walk(path):
        item = g.item_by_path(path)
        for child in g.pages(f"{cfg.drive}/items/{item['id']}/children", params={"$select": "name,folder,size"}):
            p = f"{path}/{child['name']}"
            if "folder" in child:
                walk(p)
            else:
                files.append(f"{p}  ({child.get('size', 0)} bytes)")
    walk(cfg.attachments_root)
    (out / "files.txt").write_text("\n".join(files) + "\n")
    print(f"{len(rows)} rows, {len(holds)} holds, {len(drafts)} drafts, {len(files)} files -> {out}")


if __name__ == "__main__":
    main()
