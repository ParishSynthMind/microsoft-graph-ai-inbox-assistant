from datetime import date, datetime, time

from graph_assistant.slots import WorkingHours, first_free

MON = datetime(2026, 11, 9, 9, 0)
WH = WorkingHours()


def ev(s, e, show="busy"):
    return {"start": {"dateTime": s + ".0000000"}, "end": {"dateTime": e + ".0000000"}, "showAs": show}


def test_respects_lead_time():
    assert first_free([], 30, MON, WH) == (datetime(2026, 11, 9, 11, 0), datetime(2026, 11, 9, 11, 30))


def test_skips_busy_time_with_buffer():
    events = [ev("2026-11-09T11:00:00", "2026-11-09T12:00:00")]
    start, _ = first_free(events, 30, MON, WH)
    assert start == datetime(2026, 11, 9, 12, 30)  # 12:00 end + 15 min buffer, rounded to the half hour


def test_free_and_cancelled_events_do_not_block():
    events = [ev("2026-11-09T11:00:00", "2026-11-09T12:00:00", "free"),
              {**ev("2026-11-09T11:00:00", "2026-11-09T12:00:00"), "isCancelled": True}]
    assert first_free(events, 30, MON, WH)[0] == datetime(2026, 11, 9, 11, 0)


def test_preferred_day_and_afternoon():
    start, end = first_free([], 45, MON, WH, date(2026, 11, 10), "afternoon")
    assert start == datetime(2026, 11, 10, 12, 0) and end == datetime(2026, 11, 10, 12, 45)


def test_morning_full_falls_back_to_same_day_afternoon():
    events = [ev("2026-11-10T08:00:00", "2026-11-10T12:30:00")]
    start, _ = first_free(events, 30, MON, WH, date(2026, 11, 10), "morning")
    assert start.date() == date(2026, 11, 10) and start.time() >= time(12, 45)


def test_weekend_is_skipped():
    fri_late = datetime(2026, 11, 13, 17, 0)
    start, _ = first_free([], 30, fri_late, WH)
    assert start == datetime(2026, 11, 16, 9, 0)  # Monday


def test_preferred_day_in_the_past_starts_today():
    start, _ = first_free([], 30, MON, WH, date(2026, 11, 2))
    assert start.date() == MON.date()


def test_meeting_longer_than_the_day_finds_nothing():
    assert first_free([], 600, MON, WH) is None


def test_fully_booked_fortnight_finds_nothing():
    events = [ev("2026-11-09T00:00:00", "2026-11-30T00:00:00")]
    assert first_free(events, 30, MON, WH) is None
