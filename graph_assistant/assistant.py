"""The assistant: new Outlook mail in, a logged request (and maybe a filed attachment and a calendar hold) out.

For each new Inbox message:
  1. skip it if it was already handled (local state, or the "AI logged" category on the message itself)
  2. Claude reads it and says what is being asked for (extract.py), checked against the email text
  3. attachments are saved into a dated folder in OneDrive or a SharePoint library
  4. if they asked to meet and the details check out, a tentative hold goes in the calendar (slots.py)
  5. a row goes into the Excel Online table, with a link to the files and the hold
  6. a reply proposing the slot is written into Drafts (never sent)
  7. the message gets a category, so a person can see in Outlook what was done

Nothing is ever deleted or sent. Every step is recorded in the store as it
succeeds, so a run that stops halfway carries on next time without doing
anything twice.
"""
from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from typing import Optional

from .config import Config
from .extract import CostCapReached, CostTracker, email_text
from .graph import GraphClient, GraphError
from .slots import first_free
from .store import Store

COLUMNS = ["Received", "From", "Email", "Company", "Type", "Priority", "Summary", "Due", "Meeting hold",
           "Files", "Status", "Message ID", "AI cost USD"]
MAX_ERRORS = 3


@dataclass
class Outcome:
    subject: str
    sender: str
    status: str  # logged | review | skipped | duplicate | error | ignored | planned
    detail: str = ""
    files: list[str] = field(default_factory=list)
    hold: str = ""
    cost: float = 0.0


def safe_name(s: str, limit: int = 60) -> str:
    """OneDrive and SharePoint refuse " * : < > ? / \\ | and names ending in a dot or space."""
    s = re.sub(r'[\"*:<>?/\\|#%]', " ", s or "")
    s = re.sub(r"\s+", " ", s).strip().strip(".")
    return (s[:limit].rstrip(" .") or "untitled")


def message_key(msg: dict) -> str:
    return msg.get("internetMessageId") or msg["id"]


def make_workbook_bytes(columns: list[str] = COLUMNS) -> bytes:
    """A blank .xlsx with just the header row; setup then turns it into an Excel table through Graph."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Requests"
    ws.append(columns)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Assistant:
    def __init__(self, graph: GraphClient, extractor, store: Store, cfg: Config,
                 now: Optional[datetime] = None, log=print):
        self.g = graph
        self.x = extractor
        self.store = store
        self.cfg = cfg
        self._now = now
        self.log = log
        self.me: dict = {}
        self.workbook_id: Optional[str] = None
        self.logged_ids: set[str] = set()

    def now(self) -> datetime:
        """Local wall-clock time in the configured zone (naive), matching calendarView times."""
        if self._now:
            return self._now
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo(self.cfg.timezone)).replace(tzinfo=None)

    # --- setup / check -------------------------------------------------
    def setup(self) -> str:
        """Creates the folders, the workbook and the Excel table if they are missing. Safe to run again."""
        folder, _, name = self.cfg.workbook_path.rpartition("/")
        if folder:
            self.g.ensure_folder(folder)
        self.g.ensure_folder(self.cfg.attachments_root)
        item = self.g.item_by_path(self.cfg.workbook_path)
        if item is None:
            item = self.g.upload(folder, name, make_workbook_bytes(), conflict="fail")
        tables = [t["name"] for t in self.g.list_tables(item["id"])]
        if self.cfg.table not in tables:
            address = f"A1:{_col_letter(len(COLUMNS))}1"
            self.g.create_table(item["id"], "Requests", address, self.cfg.table)
        return item.get("webUrl", item["id"])

    def check(self) -> dict:
        me = self.g.whoami()
        item = self.g.item_by_path(self.cfg.workbook_path)
        tables = [t["name"] for t in self.g.list_tables(item["id"])] if item else []
        return {"signed_in_as": me.get("mail") or me.get("userPrincipalName"), "workbook": bool(item),
                "table": self.cfg.table in tables, "workbook_url": (item or {}).get("webUrl")}

    # --- run -----------------------------------------------------------
    def run(self, dry_run: bool = False) -> list[Outcome]:
        self.me = self.g.whoami()
        my_address = (self.me.get("mail") or self.me.get("userPrincipalName") or "").lower()
        item = self.g.item_by_path(self.cfg.workbook_path)
        if item is None:
            raise RuntimeError(f"workbook {self.cfg.workbook_path!r} not found; run `graph-assistant setup` first")
        self.workbook_id = item["id"]
        id_col = COLUMNS.index("Message ID")
        self.logged_ids = {str(r[id_col]) for r in self.g.table_rows(self.workbook_id, self.cfg.table) if len(r) > id_col}

        delta = self.store.get_meta("inbox_delta")
        messages, next_delta = self.g.inbox_changes(delta)
        since = datetime.now(timezone.utc) - timedelta(days=self.cfg.since_days)
        costs = CostTracker(self.cfg.cost_cap_usd, self.cfg.model)
        outcomes = []
        stopped = False
        for msg in sorted(messages, key=lambda m: m.get("receivedDateTime", "")):
            if stopped:
                break
            try:
                received = datetime.fromisoformat(msg["receivedDateTime"].replace("Z", "+00:00"))
            except (KeyError, ValueError):
                received = datetime.now(timezone.utc)
            if delta is None and received < since:
                continue
            try:
                outcomes.append(self.handle(msg, my_address, costs, dry_run))
            except CostCapReached as e:
                self.log(f"stopping: {e}")
                stopped = True
            except Exception as e:  # one bad email never stops the run
                key = message_key(msg)
                n = self.store.record_error(key, f"{type(e).__name__}: {e}") if not dry_run else 0
                outcomes.append(Outcome(msg.get("subject") or "", _sender(msg)[1], "error",
                                        f"{type(e).__name__}: {e} (attempt {n})"))
        if not dry_run and any(o.status in ("logged", "review") for o in outcomes):
            self.tidy_sheet()
        # Only move the delta link on when every message was dealt with, so nothing is ever skipped.
        if not dry_run and not stopped and next_delta and all(o.status != "error" for o in outcomes):
            self.store.set_meta("inbox_delta", next_delta)
        return outcomes

    def handle(self, msg: dict, my_address: str, costs: CostTracker, dry_run: bool) -> Outcome:
        key = message_key(msg)
        name, address = _sender(msg)
        subject = msg.get("subject") or "(no subject)"
        rec = self.store.get(key)
        if rec and rec["status"] in ("logged", "review", "skipped"):
            return Outcome(subject, address, "duplicate", "already handled")
        if rec and rec["status"] == "error" and rec["errors"] >= MAX_ERRORS:
            return Outcome(subject, address, "error", f"skipped after {MAX_ERRORS} failures: {rec['last_error']}")
        cats = msg.get("categories") or []
        if self.cfg.cat_logged in cats or self.cfg.cat_skipped in cats or self.cfg.cat_review in cats:
            return Outcome(subject, address, "duplicate", "category already on the message")
        if key in self.logged_ids:
            return Outcome(subject, address, "duplicate", "already in the Excel table")
        if not address or address.lower() == my_address or any(p in address.lower() for p in self.cfg.ignore_senders):
            return Outcome(subject, address, "ignored", "own address or ignored sender")

        received_utc = datetime.fromisoformat(msg["receivedDateTime"].replace("Z", "+00:00"))
        received_local = self._to_local(received_utc)
        text = email_text(msg)
        steps = (rec or {}).get("steps", {})
        ex = self.x.extract(text, received_local.date(), self.cfg.timezone, costs)
        out = Outcome(subject, address, "planned", cost=ex.cost_usd)
        if not dry_run:
            self.store.upsert(key, graph_id=msg["id"], subject=subject, sender=address,
                              received=msg["receivedDateTime"], extraction=ex.to_dict(),
                              cost_usd=((rec or {}).get("cost_usd") or 0) + ex.cost_usd)

        rtype = ex.value("request_type")
        if rtype == "not_a_request" and ex.conf("request_type") == "high":
            out.status, out.detail = "skipped", "not a request"
            if not dry_run:
                self.g.add_categories(msg["id"], cats, [self.cfg.cat_skipped])
                self.store.upsert(key, status="skipped")
            return out

        problems = self._problems(ex)

        # attachments
        files_link = steps.get("files_link", "")
        if msg.get("hasAttachments") and "files" not in steps:
            folder = (f"{self.cfg.attachments_root}/{received_local:%Y-%m}/"
                      f"{received_local:%Y-%m-%d} {safe_name(name or address, 30)} - {safe_name(subject, 50)}")
            saved = []
            for att in self.g.attachments(msg["id"]):
                if att.get("isInline"):
                    continue  # signature logos and pasted pictures
                if att.get("@odata.type") not in (None, "#microsoft.graph.fileAttachment"):
                    problems.append(f"attachment {att.get('name')!r} is not a file, not saved")
                    continue
                if (att.get("size") or 0) > self.cfg.max_attachment_mb * 1024 * 1024:
                    problems.append(f"attachment {att.get('name')!r} over {self.cfg.max_attachment_mb} MB, not saved")
                    continue
                fname = safe_name(att.get("name") or "attachment", 120)
                if dry_run:
                    saved.append(fname)
                    continue
                if not saved:
                    folder_item = self.g.ensure_folder(folder)
                    files_link = folder_item.get("webUrl", "")
                self.g.upload(folder, fname, self.g.attachment_bytes(msg["id"], att["id"]))
                saved.append(fname)
            out.files = saved
            if not dry_run:
                self.store.step(key, "files", saved)
                self.store.step(key, "files_link", files_link)
        elif "files" in steps:
            out.files = steps["files"]

        # calendar hold
        hold_text = steps.get("hold", "")
        slot = None
        wants = ex.value("meeting_wanted") is True and ex.conf("meeting_wanted") != "low"
        if wants and self.cfg.book_meetings and not problems and "event_id" not in steps:
            minutes = ex.value("meeting_minutes") if ex.conf("meeting_minutes") != "low" else None
            minutes = minutes or self.cfg.default_meeting_minutes
            pref = ex.value("meeting_date") if ex.conf("meeting_date") != "low" else None
            window = ex.value("meeting_window") or "any"
            now = self.now()
            horizon_end = now + timedelta(days=self.cfg.hours.horizon_days + 3)
            events = self.g.calendar_view(now.strftime("%Y-%m-%dT00:00:00"), horizon_end.strftime("%Y-%m-%dT23:59:59"),
                                          self.cfg.timezone)
            slot = first_free(events, minutes, now, self.cfg.hours,
                              date.fromisoformat(pref) if pref else None, window)
            if slot is None:
                problems.append("no free slot in the next two weeks")
            else:
                hold_text = f"{slot[0]:%a %d %b %H:%M}-{slot[1]:%H:%M}"
                if not dry_run:
                    try:
                        ev = self.g.create_event(self._event(msg, ex, name, address, slot))
                    except GraphError as e:
                        if e.status != 409:
                            raise
                        ev = {"id": "already-booked"}  # same transactionId: an earlier attempt booked it
                    self.store.step(key, "event_id", ev.get("id"))
                    self.store.step(key, "hold", hold_text)
        elif wants and problems and "event_id" not in steps:
            problems.append("meeting asked for but not booked until a person checks it")
        out.hold = hold_text

        # Excel row
        status = "New" if not problems else "Check: " + "; ".join(problems)
        if not steps.get("row"):
            row = [received_local.strftime("%Y-%m-%d %H:%M"), ex.value("requester_name") or name, address,
                   ex.value("requester_company") or "", rtype or "", ex.value("priority") or "normal",
                   ex.value("summary") or subject, ex.value("due_date") or "", hold_text, files_link, status, key,
                   round(ex.cost_usd, 4)]
            if not dry_run:
                self.g.add_table_rows(self.workbook_id, self.cfg.table, [row])
                self.store.step(key, "row", True)
                self.logged_ids.add(key)

        # draft reply proposing the slot (never sent)
        if slot and self.cfg.draft_replies and not steps.get("draft_id") and not dry_run:
            first = (ex.value("requester_name") or name or "").split(" ")[0] or "there"
            comment = (f"Hi {first},\n\nThanks for getting in touch. I've pencilled in "
                       f"{slot[0]:%A} {slot[0].day} {slot[0]:%B} at {slot[0]:%H:%M} to {slot[1]:%H:%M} ({self.cfg.timezone}). "
                       f"Does that work for you? If not, send me a couple of times that suit and I'll move it.\n")
            draft = self.g.draft_reply(msg["id"], comment)
            self.store.step(key, "draft_id", draft.get("id"))

        out.status = "logged" if not problems else "review"
        out.detail = status
        if not dry_run:
            cat = self.cfg.cat_logged if not problems else self.cfg.cat_review
            self.g.add_categories(msg["id"], cats, [cat])
            self.store.upsert(key, status=out.status)
        else:
            out.status = "planned" if not problems else "planned-review"
        return out

    def tidy_sheet(self) -> None:
        """Cosmetic: real dates shown as dates, columns wide enough to read. Never fails a run.
        (Excel Online turns ISO date text into real dates, which then show as #### in a narrow column.)"""
        try:
            n = len(self.g.table_rows(self.workbook_id, self.cfg.table))
            ws = "Requests"
            for name, fmt in (("Received", "yyyy-mm-dd hh:mm"), ("Due", "yyyy-mm-dd")):
                col = _col_letter(COLUMNS.index(name) + 1)
                self.g.set_number_format(self.workbook_id, ws, col, 2, n + 1, fmt)
            summary = _col_letter(COLUMNS.index("Summary") + 1)
            self.g.autofit_columns(self.workbook_id, ws, f"A:{_col_letter(len(COLUMNS))}")
            self.g.set_column_format(self.workbook_id, ws, summary, columnWidth=330, wrapText=True)
            self.g.autofit_rows(self.workbook_id, ws, f"A1:{_col_letter(len(COLUMNS))}{n + 1}")
        except GraphError as e:
            self.log(f"could not tidy the sheet (rows are saved): {e}")

    # --- helpers -------------------------------------------------------
    def _problems(self, ex) -> list[str]:
        out = []
        for f in ("request_type", "summary"):
            if ex.conf(f) == "low":
                out.append(f"{f.replace('_', ' ')} unclear" + (f" ({ex.fields[f].checks[0]})" if ex.fields[f].checks else ""))
        for f in ("due_date", "meeting_date"):
            if ex.value(f) is not None and ex.conf(f) == "low":
                out.append(f"{f.replace('_', ' ')} doubtful ({'; '.join(ex.fields[f].checks) or 'low confidence'})")
        return out

    def _to_local(self, dt_utc: datetime) -> datetime:
        from zoneinfo import ZoneInfo

        return dt_utc.astimezone(ZoneInfo(self.cfg.timezone)).replace(tzinfo=None)

    def _event(self, msg, ex, name, address, slot) -> dict:
        who = ex.value("requester_name") or name or address
        body = (f"Held by the AI inbox assistant from the email \"{msg.get('subject') or ''}\" from {who} <{address}>.\n"
                f"What they want: {ex.value('summary') or ''}\n\nThis is a tentative hold: confirm or move it.")
        ev = {
            "subject": f"Hold: {who} - {safe_name(msg.get('subject') or 'meeting', 60)}",
            "start": {"dateTime": slot[0].strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": self.cfg.timezone},
            "end": {"dateTime": slot[1].strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": self.cfg.timezone},
            "showAs": "tentative",
            "body": {"contentType": "text", "content": body},
            "categories": [self.cfg.cat_logged],
            # Graph uses transactionId to refuse a second POST of the same event, so a retry never double-books.
            "transactionId": "graph-assistant-" + hashlib.sha256(message_key(msg).encode()).hexdigest()[:32],
        }
        if self.cfg.invite_sender:
            ev["attendees"] = [{"emailAddress": {"address": address, "name": who}, "type": "required"}]
        return ev


def _sender(msg: dict) -> tuple[str, str]:
    e = (msg.get("from") or {}).get("emailAddress") or {}
    return e.get("name") or "", e.get("address") or ""


def summarise(outcomes: list[Outcome]) -> str:
    lines = []
    for o in outcomes:
        extra = []
        if o.files:
            extra.append(f"{len(o.files)} file(s)")
        if o.hold:
            extra.append(f"hold {o.hold}")
        lines.append(f"[{o.status:<9}] {o.sender:<32} {o.subject[:48]:<48} {', '.join(extra)}"
                     + (f"  {o.detail}" if o.detail and o.status not in ("logged", "planned") else ""))
    spent = sum(o.cost for o in outcomes)
    lines.append(f"{len(outcomes)} message(s), AI spend ${spent:.4f}")
    return "\n".join(lines)


__all__ = ["Assistant", "COLUMNS", "GraphError", "Outcome", "make_workbook_bytes", "safe_name", "summarise"]
