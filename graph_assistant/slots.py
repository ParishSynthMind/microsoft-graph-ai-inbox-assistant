"""Finding a free calendar slot. Pure code, no model: Claude says what was asked for, this decides when.

Busy time comes from the calendar (calendarView, times already in the
assistant's time zone). Anything not marked free counts as busy, with a
buffer either side. The search starts on the day the sender asked for, in
the part of the day they asked for, and widens from there, never outside
working hours or on non-working days.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Optional

DAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


@dataclass
class WorkingHours:
    start: time = time(9, 0)
    end: time = time(17, 30)
    days: tuple[str, ...] = ("mon", "tue", "wed", "thu", "fri")
    buffer_minutes: int = 15
    step_minutes: int = 30
    lead_hours: int = 2  # never offer a slot sooner than this
    horizon_days: int = 14

    def works_on(self, d: date) -> bool:
        return DAYS[d.weekday()] in self.days


def parse_local(s: str) -> datetime:
    """Graph returns '2026-11-10T14:00:00.0000000' in the Prefer'd time zone. Keep it naive local time."""
    s = s.split(".")[0].rstrip("Z")
    return datetime.fromisoformat(s)


def busy_blocks(events: list[dict], buffer: timedelta) -> list[tuple[datetime, datetime]]:
    out = []
    for e in events:
        if e.get("isCancelled") or e.get("showAs") == "free":
            continue
        out.append((parse_local(e["start"]["dateTime"]) - buffer, parse_local(e["end"]["dateTime"]) + buffer))
    return sorted(out)


def _round_up(dt: datetime, step: int) -> datetime:
    dt = dt.replace(second=0, microsecond=0)
    extra = (-dt.minute) % step
    return dt + timedelta(minutes=extra)


def _windows(d: date, wh: WorkingHours, window: str) -> list[tuple[datetime, datetime]]:
    day_start, day_end = datetime.combine(d, wh.start), datetime.combine(d, wh.end)
    noon = datetime.combine(d, time(12, 0))
    if window == "morning":
        return [(day_start, min(noon, day_end))]
    if window == "afternoon":
        return [(max(noon, day_start), day_end)]
    return [(day_start, day_end)]


def first_free(events: list[dict], minutes: int, now: datetime, wh: WorkingHours,
               preferred_date: Optional[date] = None, window: str = "any") -> Optional[tuple[datetime, datetime]]:
    """Earliest free slot of `minutes`. Tries the preferred day and window first, then any time that day,
    then the following working days (preferred window first each day)."""
    busy = busy_blocks(events, timedelta(minutes=wh.buffer_minutes))
    length = timedelta(minutes=minutes)
    earliest = _round_up(now + timedelta(hours=wh.lead_hours), wh.step_minutes)
    first_day = max(preferred_date or earliest.date(), earliest.date())

    def fits(start: datetime) -> bool:
        end = start + length
        return start >= earliest and all(end <= b0 or start >= b1 for b0, b1 in busy)

    for offset in range(wh.horizon_days + 1):
        d = first_day + timedelta(days=offset)
        if not wh.works_on(d):
            continue
        tries = _windows(d, wh, window) + (_windows(d, wh, "any") if window != "any" else [])
        for w0, w1 in tries:
            t = _round_up(max(w0, earliest), wh.step_minutes)
            while t + length <= w1:
                if fits(t):
                    return t, t + length
                t += timedelta(minutes=wh.step_minutes)
    return None
