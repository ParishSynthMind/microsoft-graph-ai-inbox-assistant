"""Runs the whole assistant on the sample inbox with REAL Claude and the in-memory Microsoft Graph fake.

    ANTHROPIC_API_KEY=... python scripts/local_demo.py

Writes examples/local-run/: run.txt (terminal output), extractions.json (what Claude said, with the
checks), requests.xlsx (the Excel table as it ended up), calendar.json and drafts.json.
"""
from __future__ import annotations

import io
import json
import sys
from datetime import datetime
from pathlib import Path

import yaml
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from graph_assistant.assistant import COLUMNS, Assistant, summarise  # noqa: E402
from graph_assistant.config import load  # noqa: E402
from graph_assistant.extract import ClaudeExtractor  # noqa: E402
from graph_assistant.fake_graph import FakeGraph  # noqa: E402
from graph_assistant.graph import GraphClient  # noqa: E402
from graph_assistant.store import Store  # noqa: E402

OUT = ROOT / "examples" / "local-run"


def tiny_pdf(text: str) -> bytes:
    return f"%PDF-1.4\n% sample attachment: {text}\n%%EOF\n".encode()


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    fake = FakeGraph()
    for m in yaml.safe_load((ROOT / "samples" / "inbox.yaml").read_text()):
        atts = [{"name": n, "content": tiny_pdf(n)} for n in m.get("attachments", [])]
        atts += [{"name": "signature-logo.png", "content": b"PNG", "isInline": True}] if atts else []
        fake.add_message(m["subject"], tuple(m["from"]), m["body"], m["received"], attachments=atts)
    # Monday's diary: a standup and a full Tuesday lunchtime
    fake.add_event("Team standup", "2026-11-10T09:30:00", "2026-11-10T10:00:00")
    fake.add_event("Site visit", "2026-11-10T12:00:00", "2026-11-10T14:00:00")

    cfg = load(ROOT / "config" / "assistant.yaml")
    cfg.since_days = 3650
    graph = GraphClient(fake, lambda: fake.token)
    store_path = OUT / "state.db"
    store_path.unlink(missing_ok=True)
    a = Assistant(graph, ClaudeExtractor(cfg.model), Store(store_path), cfg, now=datetime(2026, 11, 9, 9, 0))
    a.setup()
    outcomes = a.run()
    text = summarise(outcomes)
    second = summarise(a.run())
    report = f"$ graph-assistant run   (Monday 9 Nov 2026 09:00, Europe/London)\n{text}\n\n$ graph-assistant run   (again)\n{second}\n"
    print(report)
    (OUT / "run.txt").write_text(report)

    book = fake.find_path(cfg.workbook_path).tables[cfg.table]
    wb = Workbook()
    ws = wb.active
    ws.title = "Requests"
    ws.append(COLUMNS)
    for r in book["rows"]:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    (OUT / "requests.xlsx").write_bytes(buf.getvalue())
    (OUT / "excel-rows.json").write_text(json.dumps([dict(zip(COLUMNS, r)) for r in book["rows"]], indent=2))
    ex = {r["subject"]: json.loads(r["extraction"]) for r in Store(store_path).all() if r["extraction"]}
    (OUT / "extractions.json").write_text(json.dumps(ex, indent=2))
    holds = [e for e in fake.events if e["subject"].startswith("Hold:")]
    (OUT / "calendar.json").write_text(json.dumps(holds, indent=2))
    (OUT / "drafts.json").write_text(json.dumps(list(fake.drafts.values()), indent=2))
    files = sorted(fake.path_of(i) for i in fake.items.values() if not i.folder)
    (OUT / "files.txt").write_text("\n".join(files) + "\n")
    store_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
