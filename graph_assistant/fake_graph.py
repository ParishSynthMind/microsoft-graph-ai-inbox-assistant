"""An in-memory stand-in for the parts of Microsoft Graph this assistant uses.

It answers with Graph's own shapes and error codes (itemNotFound,
nameAlreadyExists, ErrorItemNotFound, InvalidAuthenticationToken, 429 with
Retry-After), pages results with @odata.nextLink, ends a delta round with
@odata.deltaLink, and refuses an Authorization header on upload-session URLs
the way the real service does. Tests and the local demo run against it, so
they need no Microsoft account.
"""
from __future__ import annotations

import base64
import itertools
import json
import re
from dataclasses import dataclass, field
from urllib.parse import parse_qs, unquote, urlparse

GRAPH = "https://graph.microsoft.com/v1.0"
UPLOAD_HOST = "https://upload.fake-sharepoint.example"


class FakeResponse:
    def __init__(self, status: int, body=None, headers=None, content: bytes | None = None):
        self.status_code = status
        self.headers = headers or {}
        if content is not None:
            self.content = content
        else:
            self.content = b"" if body is None else json.dumps(body).encode()
        self._body = body

    def json(self):
        return json.loads(self.content)

    @property
    def text(self):
        return self.content.decode(errors="replace")


def err(status: int, code: str, message: str = "", headers=None) -> FakeResponse:
    return FakeResponse(status, {"error": {"code": code, "message": message or code}}, headers)


@dataclass
class DriveItem:
    id: str
    name: str
    parent: str | None
    folder: bool
    content: bytes = b""
    tables: dict = field(default_factory=dict)  # workbook: table name -> {"id", "columns", "rows"}


class FakeGraph:
    def __init__(self, me: dict | None = None, page_size: int = 2):
        self.me = me or {"id": "u1", "displayName": "Alex Morgan", "mail": "alex@contoso-demo.example",
                         "userPrincipalName": "alex@contoso-demo.example"}
        self.page_size = page_size
        self.ids = (f"id{n:04d}" for n in itertools.count(1))
        self.messages: dict[str, dict] = {}
        self.attachments: dict[str, list[dict]] = {}
        self.drafts: dict[str, dict] = {}
        self.events: list[dict] = []
        self.items: dict[str, DriveItem] = {"root": DriveItem("root", "root", None, True)}
        self.uploads: dict[str, dict] = {}
        self.sessions: set[str] = set()
        self.calls: list[tuple[str, str]] = []
        self.fail_next: list[FakeResponse] = []
        self.delta_rounds: list[set] = []  # message ids already returned, per issued delta link
        self.token = "fake-token"

    # --- seeding -------------------------------------------------------
    def add_message(self, subject: str, sender: tuple[str, str], body: str, received: str,
                    attachments: list[dict] | None = None, categories=None, internet_id=None) -> str:
        mid = "AAMk" + next(self.ids)
        self.messages[mid] = {
            "id": mid, "internetMessageId": internet_id or f"<{mid}@mail.example>", "subject": subject,
            "from": {"emailAddress": {"name": sender[0], "address": sender[1]}},
            "toRecipients": [{"emailAddress": {"address": self.me["mail"]}}],
            "receivedDateTime": received, "hasAttachments": bool(attachments),
            "body": {"contentType": "text", "content": body}, "categories": list(categories or []),
            "conversationId": "conv-" + mid,
        }
        self.attachments[mid] = []
        for a in attachments or []:
            aid = "att-" + next(self.ids)
            raw = a.get("content", b"")
            self.attachments[mid].append({
                "@odata.type": a.get("type", "#microsoft.graph.fileAttachment"), "id": aid, "name": a["name"],
                "contentType": a.get("contentType", "application/octet-stream"), "size": a.get("size", len(raw)),
                "isInline": a.get("isInline", False), "contentBytes": base64.b64encode(raw).decode()})
        return mid

    def add_event(self, subject: str, start: str, end: str, show_as: str = "busy"):
        self.events.append({"id": "ev-" + next(self.ids), "subject": subject, "start": {"dateTime": start},
                            "end": {"dateTime": end}, "showAs": show_as, "isCancelled": False})

    # --- drive helpers ---------------------------------------------------
    def children(self, parent_id: str) -> list[DriveItem]:
        return [i for i in self.items.values() if i.parent == parent_id]

    def find_path(self, path: str) -> DriveItem | None:
        cur = self.items["root"]
        for part in [p for p in path.split("/") if p]:
            nxt = next((c for c in self.children(cur.id) if c.name.lower() == part.lower()), None)
            if nxt is None:
                return None
            cur = nxt
        return cur

    def path_of(self, item: DriveItem) -> str:
        parts = []
        while item.parent:
            parts.append(item.name)
            item = self.items[item.parent]
        return "/".join(reversed(parts))

    def item_json(self, it: DriveItem) -> dict:
        d = {"id": it.id, "name": it.name, "webUrl": f"https://onedrive.example/{self.path_of(it)}",
             "parentReference": {"id": it.parent}}
        d["folder" if it.folder else "file"] = {"childCount": len(self.children(it.id))} if it.folder else {}
        if not it.folder:
            d["size"] = len(it.content)
        return d

    def _put_file(self, folder_path: str, name: str, content: bytes, conflict: str) -> FakeResponse:
        parent = self.find_path(folder_path)
        if parent is None:
            # Graph creates missing parent folders on a path PUT
            parent = self.items["root"]
            for part in [p for p in folder_path.split("/") if p]:
                nxt = next((c for c in self.children(parent.id) if c.name.lower() == part.lower()), None)
                if nxt is None:
                    nxt = DriveItem(next(self.ids), part, parent.id, True)
                    self.items[nxt.id] = nxt
                parent = nxt
        existing = next((c for c in self.children(parent.id) if c.name.lower() == name.lower()), None)
        if existing and conflict == "fail":
            return err(409, "nameAlreadyExists", "The specified item name already exists.")
        if existing and conflict == "replace":
            existing.content = content
            return FakeResponse(200, self.item_json(existing))
        if existing and conflict == "rename":
            stem, dot, ext = name.rpartition(".")
            n = 1
            while any(c.name.lower() == f"{stem} {n}.{ext}".lower() for c in self.children(parent.id)):
                n += 1
            name = f"{stem} {n}.{ext}" if dot else f"{name} {n}"
        it = DriveItem(next(self.ids), name, parent.id, False, content)
        if name.lower().endswith(".xlsx"):
            it.tables = {}
        self.items[it.id] = it
        return FakeResponse(201, self.item_json(it))

    # --- the HTTP surface -------------------------------------------------
    def request(self, method, url, json=None, data=None, headers=None, params=None, timeout=None):
        headers = headers or {}
        self.calls.append((method, url))
        if url.startswith(UPLOAD_HOST):
            return self._upload_chunk(url, data, headers)
        if headers.get("Authorization") != f"Bearer {self.token}":
            return err(401, "InvalidAuthenticationToken", "Access token is empty or invalid.")
        if self.fail_next:
            return self.fail_next.pop(0)
        u = urlparse(url)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        q.update(params or {})
        path = u.path[len("/v1.0"):] if u.path.startswith("/v1.0") else u.path
        try:
            return self._route(method, path, q, json, data, headers)
        except KeyError as e:
            return err(404, "ErrorItemNotFound", f"not found: {e}")

    def _route(self, method, path, q, body, data, headers):
        m = re.fullmatch(r"/(?:me|users/[^/]+)(/.*)?", path)
        if m and m.group(1) is None and method == "GET":
            return FakeResponse(200, self.me)
        rest = m.group(1) if m else path

        # mail
        if rest == "/mailFolders/inbox/messages/delta" and method == "GET":
            return self._delta(q)
        mm = re.fullmatch(r"/messages/([^/]+)/attachments", rest)
        if mm and method == "GET":
            atts = [{k: v for k, v in a.items() if k != "contentBytes"} for a in self.attachments[mm.group(1)]]
            return FakeResponse(200, {"value": atts})
        mm = re.fullmatch(r"/messages/([^/]+)/attachments/([^/]+)", rest)
        if mm and method == "GET":
            a = next(a for a in self.attachments[mm.group(1)] if a["id"] == mm.group(2))
            return FakeResponse(200, a)
        mm = re.fullmatch(r"/messages/([^/]+)/createReply", rest)
        if mm and method == "POST":
            orig = self.messages[mm.group(1)]
            did = "draft-" + next(self.ids)
            self.drafts[did] = {"id": did, "subject": "RE: " + (orig["subject"] or ""), "isDraft": True,
                                "replyTo": mm.group(1), "body": {"content": ""}}
            return FakeResponse(201, self.drafts[did])
        mm = re.fullmatch(r"/messages/([^/]+)", rest)
        if mm and method == "PATCH":
            target = self.drafts.get(mm.group(1)) or self.messages[mm.group(1)]
            target.update(body or {})
            return FakeResponse(200, target)

        # calendar
        if rest == "/calendarView" and method == "GET":
            if "startDateTime" not in q or "endDateTime" not in q:
                return err(400, "ErrorInvalidParameter", "startDateTime and endDateTime are required")
            lo, hi = q["startDateTime"][:19], q["endDateTime"][:19]
            evs = [e for e in self.events if e["end"]["dateTime"][:19] > lo and e["start"]["dateTime"][:19] < hi]
            return FakeResponse(200, {"value": evs})
        if rest == "/events" and method == "POST":
            tid = (body or {}).get("transactionId")
            if tid and any(e.get("transactionId") == tid for e in self.events):
                return err(409, "ErrorDuplicateTransactionId", "An event with this transactionId already exists.")
            ev = dict(body)
            ev["id"] = "ev-" + next(self.ids)
            ev.setdefault("isCancelled", False)
            self.events.append(ev)
            return FakeResponse(201, ev)

        # drive
        dm = re.fullmatch(r"/drive(/.*)", rest) or re.fullmatch(r"/(?:sites/[^/]+|drives/[^/]+)/drive(/.*)", path) \
            or re.fullmatch(r"/drives/[^/]+(/.*)", path)
        if dm:
            return self._drive(method, dm.group(1), q, body, data, headers)
        return err(400, "BadRequest", f"fake has no route for {method} {path}")

    def _delta(self, q):
        token = q.get("$deltatoken")
        skip = int(q.get("$skiptoken", "0"))
        seen = self.delta_rounds[int(token)] if token is not None else set()
        new = [m for mid, m in sorted(self.messages.items()) if mid not in seen]
        page = new[skip:skip + self.page_size]
        body = {"value": [dict(m) for m in page]}
        if skip + self.page_size < len(new):
            tq = f"$deltatoken={token}&" if token is not None else ""
            body["@odata.nextLink"] = f"{GRAPH}/me/mailFolders/inbox/messages/delta?{tq}$skiptoken={skip + self.page_size}"
        else:
            self.delta_rounds.append(set(self.messages))
            body["@odata.deltaLink"] = f"{GRAPH}/me/mailFolders/inbox/messages/delta?$deltatoken={len(self.delta_rounds) - 1}"
        return FakeResponse(200, body)

    def _drive(self, method, rest, q, body, data, headers):
        pm = re.fullmatch(r"/root:/(.+?)(?::/(content|createUploadSession))?", rest)
        if pm:
            full = unquote(pm.group(1))
            action = pm.group(2)
            if action is None and method == "GET":
                it = self.find_path(full)
                if it is None:
                    return err(404, "itemNotFound", "The resource could not be found.")
                return FakeResponse(200, self.item_json(it))
            folder, _, name = full.rpartition("/")
            if action == "content" and method == "PUT":
                if len(data or b"") > 4 * 1024 * 1024:
                    return err(413, "requestTooLarge", "Use an upload session for files over 4 MB.")
                return self._put_file(folder, name, data or b"", q.get("@microsoft.graph.conflictBehavior", "fail"))
            if action == "createUploadSession" and method == "POST":
                uid = next(self.ids)
                conflict = ((body or {}).get("item") or {}).get("@microsoft.graph.conflictBehavior", "fail")
                self.uploads[uid] = {"folder": folder, "name": name, "conflict": conflict, "buf": b""}
                return FakeResponse(200, {"uploadUrl": f"{UPLOAD_HOST}/up/{uid}", "nextExpectedRanges": ["0-"]})
        cm = re.fullmatch(r"/(?:root|items/([^/]+))/children", rest)
        if cm and method == "POST":
            parent = self.items[cm.group(1) or "root"]
            name = body["name"]
            if any(c.name.lower() == name.lower() for c in self.children(parent.id)):
                return err(409, "nameAlreadyExists", "The specified item name already exists.")
            it = DriveItem(next(self.ids), name, parent.id, True)
            self.items[it.id] = it
            return FakeResponse(201, self.item_json(it))
        wm = re.fullmatch(r"/items/([^/]+)/workbook(/.*)", rest)
        if wm:
            return self._workbook(method, self.items[wm.group(1)], wm.group(2), body, headers)
        return err(400, "BadRequest", f"fake drive has no route for {method} {rest}")

    def _upload_chunk(self, url, data, headers):
        if "Authorization" in headers:
            return err(401, "unauthenticated", "Do not send an Authorization header to the upload URL.")
        uid = url.rsplit("/", 1)[1]
        up = self.uploads[uid]
        m = re.fullmatch(r"bytes (\d+)-(\d+)/(\d+)", headers.get("Content-Range", ""))
        if not m or int(m.group(1)) != len(up["buf"]):
            return err(416, "invalidRange", "Content-Range does not continue the upload.")
        start, end, total = map(int, m.groups())
        if (end - start + 1) % (320 * 1024) and end + 1 != total:
            return err(400, "invalidRange", "Chunks must be a multiple of 320 KiB.")
        up["buf"] += data
        if len(up["buf"]) < total:
            return FakeResponse(202, {"nextExpectedRanges": [f"{len(up['buf'])}-"]})
        return self._put_file(up["folder"], up["name"], up["buf"], up["conflict"])

    def _workbook(self, method, item, rest, body, headers):
        if not item.name.lower().endswith(".xlsx"):
            return err(400, "InvalidRequest", "Not a workbook.")
        sid = headers.get("workbook-session-id")
        if sid and sid not in self.sessions:
            return err(404, "InvalidSessionReCreatable", "The session is no longer valid.")
        if rest == "/createSession" and method == "POST":
            s = "sess-" + next(self.ids)
            self.sessions.add(s)
            return FakeResponse(201, {"id": s, "persistChanges": True})
        if rest == "/closeSession" and method == "POST":
            self.sessions.discard(sid)
            return FakeResponse(204)
        if rest == "/tables" and method == "GET":
            return FakeResponse(200, {"value": [{"id": t["id"], "name": n} for n, t in item.tables.items()]})
        if re.fullmatch(r"/worksheets/[^/]+/tables/add", rest) and method == "POST":
            tid = "{" + next(self.ids) + "}"
            name = f"Table{len(item.tables) + 1}"
            cols = _header_from_xlsx(item.content)
            item.tables[name] = {"id": tid, "columns": cols, "rows": [], "address": body["address"]}
            return FakeResponse(201, {"id": tid, "name": name})
        tm = re.fullmatch(r"/tables/([^/]+)", rest)
        if tm and method == "PATCH":
            key = unquote(tm.group(1))
            name, t = next((n, t) for n, t in item.tables.items() if t["id"] == key or n == key)
            new = body.get("name", name)
            item.tables[new] = item.tables.pop(name)
            return FakeResponse(200, {"id": t["id"], "name": new})
        rm = re.fullmatch(r"/tables/([^/]+)/rows", rest)
        if rm:
            name = unquote(rm.group(1))
            if name not in item.tables:
                return err(404, "ItemNotFound", f"The requested resource doesn't exist: table {name}")
            t = item.tables[name]
            if method == "GET":
                return FakeResponse(200, {"value": [{"index": i, "values": [r]} for i, r in enumerate(t["rows"])]})
            if method == "POST":
                for r in body["values"]:
                    if len(r) != len(t["columns"]):
                        return err(400, "InvalidArgument", "The number of values does not match the table columns.")
                t["rows"].extend(body["values"])
                return FakeResponse(201, {"index": len(t["rows"]) - 1, "values": body["values"]})
        return err(400, "BadRequest", f"fake workbook has no route for {method} {rest}")


def _header_from_xlsx(content: bytes) -> list[str]:
    import io

    from openpyxl import load_workbook

    ws = load_workbook(io.BytesIO(content)).active
    return [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
