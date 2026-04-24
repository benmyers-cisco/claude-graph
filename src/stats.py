"""Meeting statistics computation module.

Pure computation — no API calls. Takes event dicts (same format as
calendar_tools.get_events_in_range) and computes meeting load stats.
"""

import re
from datetime import datetime, date, timedelta

# Subjects matching these patterns are not real meetings (case-insensitive).
_SKIP_PATTERNS = [
    "hold",
    "working session",
    "busy",
    "lunch",
    "focus time",
    "block",
    "off hours",
    "ooo",
]

_SKIP_RE = re.compile("|".join(re.escape(p) for p in _SKIP_PATTERNS), re.IGNORECASE)


def _is_real_meeting(event: dict, user_email: str | None = None) -> bool:
    """Return True if the event is a real meeting (not a hold/block/solo)."""
    if event.get("is_cancelled"):
        return False
    if event.get("is_all_day"):
        return False
    subject = event.get("subject") or ""
    if _SKIP_RE.search(subject):
        return False
    # Solo event: organizer is the only attendee (no other non-resource attendees)
    attendees = event.get("attendees") or []
    non_resource = [a for a in attendees if a.get("type") != "resource"]
    if len(non_resource) <= 1:
        return False
    return True


def _parse_time(dt_dict: dict) -> datetime:
    """Parse a time dict from calendar_tools into a datetime."""
    return datetime.fromisoformat(dt_dict["date_time"])


def _merge_intervals(
    intervals: list[tuple[datetime, datetime]],
    gap_minutes: int = 0,
) -> list[tuple[datetime, datetime]]:
    """Merge overlapping time intervals.

    Args:
        intervals: List of (start, end) tuples.
        gap_minutes: Also merge intervals separated by gaps this short.
            E.g., gap_minutes=10 merges meetings with <=10 min between them
            (transition time, not real free time).
    """
    if not intervals:
        return []
    sorted_intervals = sorted(intervals, key=lambda x: x[0])
    gap = timedelta(minutes=gap_minutes)
    merged = [sorted_intervals[0]]
    for start, end in sorted_intervals[1:]:
        if start <= merged[-1][1] + gap:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _clamp(dt: datetime, lo: datetime, hi: datetime) -> datetime:
    """Clamp a datetime between lo and hi."""
    if dt < lo:
        return lo
    if dt > hi:
        return hi
    return dt


def compute_day_stats(
    events: list[dict],
    event_date: date | None = None,
    work_start: str = "09:00",
    work_end: str = "17:00",
    user_email: str | None = None,
    min_free_block_minutes: int = 10,
) -> dict:
    """Compute meeting stats for a single day's events.

    Args:
        events: List of event dicts from get_events_in_range.
        event_date: The date these events are for (used in output).
        work_start: Core hours start (HH:MM).
        work_end: Core hours end (HH:MM).
        user_email: Current user email for filtering solo events.
        min_free_block_minutes: Gaps shorter than this between meetings are
            absorbed as transition time, not counted as free blocks (default: 10).

    Returns:
        Dict with meetings, free_time, and off_hours stats.
    """
    meetings = [e for e in events if _is_real_meeting(e, user_email)]

    # Parse meeting intervals
    intervals = []
    meeting_items = []
    for m in meetings:
        start_dt = _parse_time(m["start"])
        end_dt = _parse_time(m["end"])
        intervals.append((start_dt, end_dt))
        meeting_items.append({
            "subject": m.get("subject"),
            "start": start_dt.strftime("%H:%M"),
            "end": end_dt.strftime("%H:%M"),
            "duration_minutes": int((end_dt - start_dt).total_seconds() / 60),
        })

    # Sort meeting items by start time
    meeting_items.sort(key=lambda x: x["start"])

    # Determine the day for core hours boundaries
    if event_date is None and intervals:
        event_date = intervals[0][0].date()
    elif event_date is None:
        event_date = date.today()

    ws_h, ws_m = map(int, work_start.split(":"))
    we_h, we_m = map(int, work_end.split(":"))
    # Build naive datetimes (same as event times from the CLI)
    core_start = datetime(event_date.year, event_date.month, event_date.day, ws_h, ws_m)
    core_end = datetime(event_date.year, event_date.month, event_date.day, we_h, we_m)
    core_minutes = int((core_end - core_start).total_seconds() / 60)

    # Merge overlapping intervals (and absorb short gaps as transition time)
    merged = _merge_intervals(intervals, gap_minutes=min_free_block_minutes)

    # Split into core-hours and off-hours
    core_blocks = []
    off_hours_items = []
    for start, end in merged:
        # Off-hours portion before core start
        if start < core_start:
            off_end = min(end, core_start)
            off_hours_items.append({
                "start": start.strftime("%H:%M"),
                "end": off_end.strftime("%H:%M"),
                "duration_minutes": int((off_end - start).total_seconds() / 60),
            })
        # Off-hours portion after core end
        if end > core_end:
            off_start = max(start, core_end)
            off_hours_items.append({
                "start": off_start.strftime("%H:%M"),
                "end": end.strftime("%H:%M"),
                "duration_minutes": int((end - off_start).total_seconds() / 60),
            })
        # Core-hours portion
        clamped_start = _clamp(start, core_start, core_end)
        clamped_end = _clamp(end, core_start, core_end)
        if clamped_start < clamped_end:
            core_blocks.append((clamped_start, clamped_end))

    # Compute core meeting minutes
    core_meeting_minutes = sum(
        int((e - s).total_seconds() / 60) for s, e in core_blocks
    )

    # Compute free blocks within core hours
    free_blocks = []
    if not core_blocks:
        free_blocks.append({
            "start": core_start.strftime("%H:%M"),
            "end": core_end.strftime("%H:%M"),
            "duration_minutes": core_minutes,
        })
    else:
        # Gap before first meeting
        if core_blocks[0][0] > core_start:
            gap = int((core_blocks[0][0] - core_start).total_seconds() / 60)
            free_blocks.append({
                "start": core_start.strftime("%H:%M"),
                "end": core_blocks[0][0].strftime("%H:%M"),
                "duration_minutes": gap,
            })
        # Gaps between meetings
        for i in range(len(core_blocks) - 1):
            gap_start = core_blocks[i][1]
            gap_end = core_blocks[i + 1][0]
            gap = int((gap_end - gap_start).total_seconds() / 60)
            if gap > 0:
                free_blocks.append({
                    "start": gap_start.strftime("%H:%M"),
                    "end": gap_end.strftime("%H:%M"),
                    "duration_minutes": gap,
                })
        # Gap after last meeting
        if core_blocks[-1][1] < core_end:
            gap = int((core_end - core_blocks[-1][1]).total_seconds() / 60)
            free_blocks.append({
                "start": core_blocks[-1][1].strftime("%H:%M"),
                "end": core_end.strftime("%H:%M"),
                "duration_minutes": gap,
            })

    total_free_minutes = sum(b["duration_minutes"] for b in free_blocks)
    longest_free = max((b["duration_minutes"] for b in free_blocks), default=0)
    blocks_under_30 = sum(1 for b in free_blocks if b["duration_minutes"] < 30)
    blocks_30_plus = sum(1 for b in free_blocks if b["duration_minutes"] >= 30)

    off_hours_total = sum(item["duration_minutes"] for item in off_hours_items)
    # Attach subjects to off-hours items from the original meetings
    off_hours_with_subjects = []
    for item in off_hours_items:
        # Find the original meeting that covers this off-hours slot
        subject = None
        for m in meeting_items:
            if m["start"] <= item["start"] and m["end"] >= item["end"]:
                subject = m["subject"]
                break
            # Also check if the meeting overlaps at all
            if m["start"] < item["end"] and m["end"] > item["start"]:
                subject = m["subject"]
                break
        off_hours_with_subjects.append({**item, "subject": subject})

    return {
        "date": event_date.isoformat(),
        "day_name": event_date.strftime("%A"),
        "meetings": {
            "count": len(meetings),
            "total_minutes": core_meeting_minutes,
            "items": meeting_items,
        },
        "free_time": {
            "total_minutes": total_free_minutes,
            "blocks": free_blocks,
            "longest_block_minutes": longest_free,
            "blocks_under_30min": blocks_under_30,
            "blocks_30min_plus": blocks_30_plus,
        },
        "off_hours": {
            "count": len(off_hours_with_subjects),
            "total_minutes": off_hours_total,
            "items": off_hours_with_subjects,
        },
    }


def compute_range_stats(daily_stats: list[dict]) -> dict:
    """Aggregate stats across multiple days.

    Args:
        daily_stats: List of dicts from compute_day_stats.

    Returns:
        Dict with totals and averages.
    """
    if not daily_stats:
        return {
            "total_days": 0,
            "total_meetings": 0,
            "total_meeting_minutes": 0,
            "total_free_minutes": 0,
            "avg_meetings_per_day": 0,
            "avg_meeting_minutes_per_day": 0,
            "avg_free_minutes_per_day": 0,
            "busiest_day": None,
            "lightest_day": None,
            "off_hours_total_count": 0,
            "off_hours_total_minutes": 0,
        }

    n = len(daily_stats)
    total_meetings = sum(d["meetings"]["count"] for d in daily_stats)
    total_mtg_min = sum(d["meetings"]["total_minutes"] for d in daily_stats)
    total_free_min = sum(d["free_time"]["total_minutes"] for d in daily_stats)
    off_count = sum(d["off_hours"]["count"] for d in daily_stats)
    off_min = sum(d["off_hours"]["total_minutes"] for d in daily_stats)

    busiest = max(daily_stats, key=lambda d: d["meetings"]["total_minutes"])
    lightest = min(daily_stats, key=lambda d: d["meetings"]["total_minutes"])

    return {
        "total_days": n,
        "total_meetings": total_meetings,
        "total_meeting_minutes": total_mtg_min,
        "total_free_minutes": total_free_min,
        "avg_meetings_per_day": round(total_meetings / n, 1),
        "avg_meeting_minutes_per_day": round(total_mtg_min / n, 1),
        "avg_free_minutes_per_day": round(total_free_min / n, 1),
        "busiest_day": {
            "date": busiest["date"],
            "day_name": busiest["day_name"],
            "meeting_minutes": busiest["meetings"]["total_minutes"],
        },
        "lightest_day": {
            "date": lightest["date"],
            "day_name": lightest["day_name"],
            "meeting_minutes": lightest["meetings"]["total_minutes"],
        },
        "off_hours_total_count": off_count,
        "off_hours_total_minutes": off_min,
    }
