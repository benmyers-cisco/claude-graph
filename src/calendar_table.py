"""Plain-text calendar rendering shared by /morning-coffee, /today, /wrap-up, /cos, /1on1s.

The raw `calendar` JSON carries every event body (411 KB for 14 days), which overflows a
tool result and used to leave the Week-ahead section blank. This renders the standard
Time / Dur / Subject table, the RSVP markers, and the `Open:` line in one place so every
skill shows the same thing.

RSVP rules (Ben, 2026-11-18):
  accepted / organizer / no status  -> confirmed: busy, no marker
  tentativelyAccepted               -> busy, marked "(tentative)"
  notResponded                      -> NOT busy (counts as open), marked "(no RSVP)"
  declined                          -> dropped entirely
"""

from datetime import datetime, timedelta

DAY_START_HOUR = 8
MIN_OPEN_MINUTES = 30


def _parse(dt_str: str) -> datetime:
    return datetime.fromisoformat(dt_str)


def _clock(dt: datetime) -> str:
    return dt.strftime("%-I:%M %p")


def _is_declined(event: dict) -> bool:
    return event.get("response_status") == "declined" or (
        (event.get("subject") or "").startswith("Declined:")
    )


def _marker(event: dict) -> str:
    status = event.get("response_status")
    if status == "tentativelyAccepted":
        return " (tentative)"
    if status == "notResponded":
        return " (no RSVP)"
    return ""


def _counts_as_busy(event: dict) -> bool:
    return event.get("response_status") != "notResponded"


def _open_blocks(busy: list[tuple[datetime, datetime]], day_start: datetime) -> str:
    """Free gaps of at least MIN_OPEN_MINUTES, with 'onward' after the last meeting."""
    merged: list[list[datetime]] = []
    for start, end in sorted(busy):
        if merged and start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])

    gaps: list[str] = []
    cursor = day_start
    for start, end in merged:
        if start - cursor >= timedelta(minutes=MIN_OPEN_MINUTES):
            gaps.append(_range(cursor, start))
        cursor = max(cursor, end)
    if not merged:
        return "all day"
    gaps.append(f"{_clock(cursor)} onward")
    return ", ".join(gaps)


def _range(a: datetime, b: datetime) -> str:
    """`10:00–11:00 AM`, repeating the meridiem only when it changes."""
    if a.strftime("%p") == b.strftime("%p"):
        return f"{a.strftime('%-I:%M')}–{_clock(b)}"
    return f"{_clock(a)}–{_clock(b)}"


def _by_day(events: list[dict]) -> dict[str, list[dict]]:
    days: dict[str, list[dict]] = {}
    for event in sorted(events, key=lambda e: e["start"]["date_time"]):
        if _is_declined(event):
            continue
        days.setdefault(event["start"]["date"], []).append(event)
    return days


def _heading(day: str) -> str:
    return _parse(day).strftime("%A %-m/%-d")


def render_tables(events: list[dict], remaining_from: datetime | None = None) -> str:
    """One table + Open line per day. `remaining_from` drops ended meetings (for /today)."""
    days = _by_day(events)
    if not days:
        return "Calendar's clear.\n"

    multi = len(days) > 1
    out: list[str] = []
    for day, day_events in days.items():
        timed = [e for e in day_events if not e.get("is_all_day")]
        all_day = [e for e in day_events if e.get("is_all_day")]
        start_of_day = _parse(day).replace(hour=DAY_START_HOUR)

        if remaining_from is not None and day == remaining_from.date().isoformat():
            timed = [e for e in timed if _parse(e["end"]["date_time"]) > remaining_from]
            start_of_day = max(start_of_day, remaining_from.replace(second=0, microsecond=0))

        if multi:
            out.append(f"### {_heading(day)}\n")

        if timed:
            out.append("| Time | Dur | Subject |")
            out.append("|------|-----|---------|")
            for e in timed:
                out.append(
                    f"| {_clock(_parse(e['start']['date_time']))} "
                    f"| {e['duration_minutes']}m | {e['subject'].strip()}{_marker(e)} |"
                )
            busy = [
                (_parse(e["start"]["date_time"]), _parse(e["end"]["date_time"]))
                for e in timed
                if _counts_as_busy(e)
            ]
            out.append("")
            out.append(f"Open: {_open_blocks(busy, start_of_day)}")
        else:
            out.append("No meetings." if remaining_from is None else "Calendar's clear for the rest of the day.")

        if all_day:
            out.append("All day: " + " · ".join(f"{e['subject'].strip()}{_marker(e)}" for e in all_day))
        out.append("")
    return "\n".join(out)


def render_bullets(events: list[dict]) -> str:
    """One bullet per day for Week-ahead sections. Days with nothing are skipped."""
    lines = []
    for day, day_events in _by_day(events).items():
        parts = [
            f"{e['subject'].strip()}{_marker(e)}, {_clock(_parse(e['start']['date_time']))}"
            for e in day_events
            if not e.get("is_all_day")
        ]
        if parts:
            lines.append(f"- **{_heading(day)}** — " + " · ".join(parts))
    return "\n".join(lines) + "\n" if lines else "Nothing on the calendar.\n"
