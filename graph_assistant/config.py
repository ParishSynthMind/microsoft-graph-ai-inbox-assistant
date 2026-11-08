"""Settings from config/assistant.yaml, with safe defaults for anything left out."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import time
from pathlib import Path

import yaml

from .slots import WorkingHours


@dataclass
class Config:
    mailbox: str = "/me"
    drive: str = "/me/drive"
    workbook_path: str = "AI Inbox/Requests.xlsx"
    table: str = "Requests"
    attachments_root: str = "AI Inbox/Attachments"
    timezone: str = "Europe/London"
    since_days: int = 7
    ignore_senders: list[str] = field(default_factory=lambda: ["noreply@", "no-reply@", "mailer-daemon@"])
    book_meetings: bool = True
    invite_sender: bool = False
    draft_replies: bool = True
    hours: WorkingHours = field(default_factory=WorkingHours)
    default_meeting_minutes: int = 30
    model: str = "claude-sonnet-5-5"
    cost_cap_usd: float = 0.50
    max_attachment_mb: int = 150
    cat_logged: str = "AI logged"
    cat_skipped: str = "AI skipped"
    cat_review: str = "AI check"


def _t(s: str) -> time:
    h, m = str(s).split(":")
    return time(int(h), int(m))


def load(path: str | Path) -> Config:
    raw = yaml.safe_load(Path(path).read_text()) or {}
    wh = raw.pop("working_hours", {}) or {}
    cats = raw.pop("categories", {}) or {}
    cfg = Config(**{k: v for k, v in raw.items() if k in Config.__dataclass_fields__})
    cfg.hours = WorkingHours(
        start=_t(wh.get("start", "09:00")), end=_t(wh.get("end", "17:30")),
        days=tuple(wh.get("days", WorkingHours.days)), buffer_minutes=int(wh.get("buffer_minutes", 15)),
        lead_hours=int(wh.get("lead_hours", 2)))
    cfg.default_meeting_minutes = int(wh.get("default_meeting_minutes", 30))
    cfg.cat_logged = cats.get("logged", cfg.cat_logged)
    cfg.cat_skipped = cats.get("skipped", cfg.cat_skipped)
    cfg.cat_review = cats.get("review", cfg.cat_review)
    return cfg
