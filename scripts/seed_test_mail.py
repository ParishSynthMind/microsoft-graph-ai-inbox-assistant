"""Put the sample emails into a real Outlook Inbox through Graph, for a live test.

    python scripts/seed_test_mail.py samples/live-inbox.yaml

Used for the live test because the test account has no email address that can receive mail from
outside. Each message is created in the Inbox with the MAPI message-flags property set to 1 (read,
not a draft), which is the usual way to import mail through Graph, so it looks like received mail
to the assistant. Attachments named in the YAML are read from samples/attachments/.
"""
from __future__ import annotations

import base64
import os
import sys
from pathlib import Path

import requests
import yaml

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from graph_assistant import auth  # noqa: E402
from graph_assistant.graph import GraphClient  # noqa: E402


def main():
    mails = yaml.safe_load(Path(sys.argv[1]).read_text())
    token = auth.device_code_token(os.environ["GRAPH_CLIENT_ID"], os.environ.get("GRAPH_TENANT", "consumers"))
    g = GraphClient(requests.Session(), token)
    for m in mails:
        atts = []
        for name in m.get("attachments", []):
            data = (ROOT / "samples" / "attachments" / name).read_bytes()
            atts.append({"@odata.type": "#microsoft.graph.fileAttachment", "name": name,
                         "contentType": "application/pdf" if name.endswith(".pdf") else "application/octet-stream",
                         "contentBytes": base64.b64encode(data).decode()})
        g.request("POST", "/me/mailFolders/inbox/messages", json={
            "subject": m["subject"],
            "body": {"contentType": "text", "content": m["body"]},
            "from": {"emailAddress": {"name": m["from"][0], "address": m["from"][1]}},
            "sender": {"emailAddress": {"name": m["from"][0], "address": m["from"][1]}},
            "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
            "attachments": atts,
            # PR_MESSAGE_FLAGS = 1 (read): created as received mail, not a draft
            "singleValueExtendedProperties": [{"id": "Integer 0x0E07", "value": "1"}],
        })
        print("placed:", m["subject"], f"({len(m.get('attachments', []))} attachment(s))")


if __name__ == "__main__":
    main()
