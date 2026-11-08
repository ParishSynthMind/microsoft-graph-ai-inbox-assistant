"""A small Microsoft Graph client: only the calls this assistant needs, done carefully.

Graph says "slow down" with HTTP 429 (and sometimes 503/504), usually with a
Retry-After header in seconds. Those are retried, honouring Retry-After, with
exponential backoff and jitter when it is missing. Everything else is raised
at once as a GraphError carrying Graph's own error code (ErrorAccessDenied,
itemNotFound, InvalidAuthenticationToken ...), so the caller can tell a
missing permission from a typo in a folder path.

Paths are relative to https://graph.microsoft.com/v1.0. The mailbox and the
drive are configurable, so the same code runs as:
  - delegated (a person signed in):   mailbox "/me",            drive "/me/drive"
  - app-only (a background service):  mailbox "/users/{upn}",   drive "/sites/{site-id}/drive"
"""
from __future__ import annotations

import base64
import random
import time
from typing import Callable, Iterator, Optional
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"
RETRY_STATUS = {429, 503, 504}
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # Graph's simple PUT upload is for files up to 4 MB
CHUNK = 320 * 1024 * 10  # upload-session chunks must be multiples of 320 KiB


class GraphError(RuntimeError):
    def __init__(self, status: int, code: str, message: str, url: str = ""):
        super().__init__(f"Graph {status} {code}: {message}")
        self.status = status
        self.code = code
        self.message = message
        self.url = url


def error_code(resp) -> tuple[str, str]:
    try:
        err = resp.json().get("error", {})
    except Exception:
        return "unknown", (getattr(resp, "text", "") or "")[:200]
    if isinstance(err, str):  # token endpoint style
        return err, str(resp.json().get("error_description", ""))
    return str(err.get("code") or "unknown"), str(err.get("message") or "")


def drive_path(path: str) -> str:
    """'Attachments/2026-11/a b.pdf' -> URL-safe path segment for /root:/{path}:"""
    return "/".join(quote(p, safe="") for p in path.strip("/").split("/") if p)


class GraphClient:
    """`session` is anything with a requests-style .request() (requests.Session, or FakeGraph in tests).
    `token` is a callable returning a current access token (MSAL refreshes it)."""

    def __init__(self, session, token: Callable[[], str], mailbox: str = "/me", drive: str = "/me/drive",
                 max_retries: int = 6, base_delay: float = 1.0, max_delay: float = 60.0,
                 sleep: Callable[[float], None] = time.sleep, timeout: float = 60.0,
                 rng: Optional[random.Random] = None):
        self.session = session
        self.token = token
        self.mailbox = mailbox.rstrip("/")
        self.drive = drive.rstrip("/")
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.max_delay = max_delay
        self.sleep = sleep
        self.timeout = timeout
        self.rng = rng or random.Random()
        self.retries = 0

    # --- transport -----------------------------------------------------
    def request(self, method: str, url: str, *, json=None, data=None, headers=None, params=None, raw=False):
        if not url.startswith("http"):
            url = GRAPH + url
        hdrs = {"Authorization": f"Bearer {self.token()}"}
        hdrs.update(headers or {})
        for attempt in range(self.max_retries + 1):
            resp = self.session.request(method, url, json=json, data=data, headers=hdrs, params=params,
                                        timeout=self.timeout)
            if resp.status_code in RETRY_STATUS and attempt < self.max_retries:
                self.retries += 1
                self.sleep(self._delay(resp, attempt))
                continue
            if resp.status_code >= 400:
                code, msg = error_code(resp)
                raise GraphError(resp.status_code, code, msg, url)
            if raw:
                return resp
            if resp.status_code == 204 or not resp.content:
                return {}
            return resp.json()
        raise AssertionError("unreachable")

    def _delay(self, resp, attempt: int) -> float:
        ra = resp.headers.get("Retry-After") if getattr(resp, "headers", None) else None
        if ra:
            try:
                return min(float(ra), self.max_delay)
            except ValueError:
                pass
        return min(self.base_delay * (2 ** attempt), self.max_delay) * (0.5 + self.rng.random() / 2)

    def pages(self, url: str, params=None) -> Iterator[dict]:
        """Yields every item across @odata.nextLink pages. The last page's @odata.deltaLink is kept on self."""
        self.last_delta_link = None
        while url:
            body = self.request("GET", url, params=params)
            params = None  # nextLink already carries the query
            yield from body.get("value", [])
            url = body.get("@odata.nextLink")
            if "@odata.deltaLink" in body:
                self.last_delta_link = body["@odata.deltaLink"]

    # --- who am I ------------------------------------------------------
    def whoami(self) -> dict:
        return self.request("GET", self.mailbox, params={"$select": "displayName,mail,userPrincipalName,id"})

    # --- mail ----------------------------------------------------------
    MESSAGE_FIELDS = "id,internetMessageId,subject,from,toRecipients,receivedDateTime,hasAttachments,body,categories,conversationId"

    def inbox_changes(self, delta_link: Optional[str] = None) -> tuple[list[dict], Optional[str]]:
        """New or changed Inbox messages since the last call (delta query). Returns (messages, next delta link)."""
        url = delta_link or f"{self.mailbox}/mailFolders/inbox/messages/delta"
        params = None if delta_link else {"$select": self.MESSAGE_FIELDS}
        headers = {"Prefer": 'outlook.body-content-type="text", odata.maxpagesize=50'}
        out = []
        self.last_delta_link = None
        while url:
            body = self.request("GET", url, params=params, headers=headers)
            params = None
            out.extend(m for m in body.get("value", []) if "@removed" not in m)
            url = body.get("@odata.nextLink")
            if "@odata.deltaLink" in body:
                self.last_delta_link = body["@odata.deltaLink"]
        return out, self.last_delta_link

    def attachments(self, message_id: str) -> list[dict]:
        url = f"{self.mailbox}/messages/{message_id}/attachments"
        return list(self.pages(url, params={"$select": "id,name,contentType,size,isInline"}))

    def attachment_bytes(self, message_id: str, attachment_id: str) -> bytes:
        att = self.request("GET", f"{self.mailbox}/messages/{message_id}/attachments/{attachment_id}")
        if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
            raise GraphError(400, "notAFile", f"attachment {att.get('name')!r} is an Outlook item or link, not a file")
        return base64.b64decode(att["contentBytes"])

    def add_categories(self, message_id: str, existing: list[str], new: list[str]) -> None:
        cats = list(dict.fromkeys([*existing, *new]))
        self.request("PATCH", f"{self.mailbox}/messages/{message_id}", json={"categories": cats})

    def draft_reply(self, message_id: str, comment: str) -> dict:
        """Creates a reply in Drafts. It is NOT sent; a person reviews and sends it."""
        draft = self.request("POST", f"{self.mailbox}/messages/{message_id}/createReply")
        self.request("PATCH", f"{self.mailbox}/messages/{draft['id']}",
                     json={"body": {"contentType": "text", "content": comment}})
        return draft

    # --- files (OneDrive or a SharePoint document library) ---------------
    def item_by_path(self, path: str) -> Optional[dict]:
        try:
            return self.request("GET", f"{self.drive}/root:/{drive_path(path)}")
        except GraphError as e:
            if e.status == 404:
                return None
            raise

    def ensure_folder(self, path: str) -> dict:
        """Creates each missing folder along the path. Safe to run again."""
        parent_url = f"{self.drive}/root"
        item = None
        done = []
        for part in [p for p in path.strip("/").split("/") if p]:
            done.append(part)
            item = self.item_by_path("/".join(done))
            if item is None:
                try:
                    item = self.request("POST", f"{parent_url}/children", json={
                        "name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
                except GraphError as e:
                    if e.code != "nameAlreadyExists":
                        raise
                    item = self.item_by_path("/".join(done))
            parent_url = f"{self.drive}/items/{item['id']}"
        return item

    def upload(self, folder: str, name: str, content: bytes, conflict: str = "replace") -> dict:
        """Uploads a file into folder. Small files in one PUT, bigger ones through an upload session."""
        path = drive_path(f"{folder}/{name}")
        if len(content) <= SIMPLE_UPLOAD_LIMIT:
            return self.request("PUT", f"{self.drive}/root:/{path}:/content",
                                params={"@microsoft.graph.conflictBehavior": conflict}, data=content,
                                headers={"Content-Type": "application/octet-stream"})
        sess = self.request("POST", f"{self.drive}/root:/{path}:/createUploadSession",
                            json={"item": {"@microsoft.graph.conflictBehavior": conflict}})
        upload_url = sess["uploadUrl"]
        result = {}
        for start in range(0, len(content), CHUNK):
            chunk = content[start:start + CHUNK]
            end = start + len(chunk) - 1
            # The upload URL is pre-authorised: Graph rejects it if an Authorization header is sent.
            resp = self.session.request("PUT", upload_url, data=chunk, timeout=self.timeout, headers={
                "Content-Length": str(len(chunk)), "Content-Range": f"bytes {start}-{end}/{len(content)}"})
            if resp.status_code >= 400:
                code, msg = error_code(resp)
                raise GraphError(resp.status_code, code, msg, upload_url)
            result = resp.json() if resp.content else {}
        return result

    # --- Excel Online (workbook API) -----------------------------------
    def workbook_session(self, item_id: str) -> str:
        return self.request("POST", f"{self.drive}/items/{item_id}/workbook/createSession",
                            json={"persistChanges": True})["id"]

    def close_workbook_session(self, item_id: str, session_id: str) -> None:
        try:
            self.request("POST", f"{self.drive}/items/{item_id}/workbook/closeSession",
                         headers={"workbook-session-id": session_id})
        except GraphError:
            pass  # sessions expire by themselves; never fail a run over this

    def table_rows(self, item_id: str, table: str, session_id: Optional[str] = None) -> list[list]:
        hdr = {"workbook-session-id": session_id} if session_id else None
        body = self.request("GET", f"{self.drive}/items/{item_id}/workbook/tables/{quote(table)}/rows",
                            headers=hdr)
        return [r["values"][0] for r in body.get("value", [])]

    def add_table_rows(self, item_id: str, table: str, rows: list[list], session_id: Optional[str] = None) -> dict:
        hdr = {"workbook-session-id": session_id} if session_id else None
        return self.request("POST", f"{self.drive}/items/{item_id}/workbook/tables/{quote(table)}/rows",
                            json={"values": rows}, headers=hdr)

    def create_table(self, item_id: str, worksheet: str, address: str, name: str,
                     session_id: Optional[str] = None) -> dict:
        hdr = {"workbook-session-id": session_id} if session_id else None
        t = self.request("POST", f"{self.drive}/items/{item_id}/workbook/worksheets/{quote(worksheet)}/tables/add",
                         json={"address": address, "hasHeaders": True}, headers=hdr)
        return self.request("PATCH", f"{self.drive}/items/{item_id}/workbook/tables/{t['id']}",
                            json={"name": name}, headers=hdr)

    def list_tables(self, item_id: str) -> list[dict]:
        return self.request("GET", f"{self.drive}/items/{item_id}/workbook/tables").get("value", [])

    # --- calendar ------------------------------------------------------
    def calendar_view(self, start_iso: str, end_iso: str, tz: str) -> list[dict]:
        """Every event between start and end, with times given in `tz` (an IANA or Windows zone name)."""
        url = f"{self.mailbox}/calendarView"
        params = {"startDateTime": start_iso, "endDateTime": end_iso,
                  "$select": "subject,start,end,showAs,isCancelled,transactionId", "$top": "100"}
        headers = {"Prefer": f'outlook.timezone="{tz}"'}
        out = []
        while url:
            body = self.request("GET", url, params=params, headers=headers)
            params = None
            out.extend(body.get("value", []))
            url = body.get("@odata.nextLink")
        return out

    def create_event(self, event: dict) -> dict:
        return self.request("POST", f"{self.mailbox}/events", json=event)
