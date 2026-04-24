"""Meeting time finder tools for Microsoft Graph API.

Find optimal meeting times for multiple participants using the findMeetingTimes
API with fallback to manual slot finding via getSchedule.
"""

from datetime import datetime, date, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.attendee_base import AttendeeBase
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.time_constraint import TimeConstraint
from msgraph.generated.models.time_slot import TimeSlot
from msgraph.generated.models.location_constraint import LocationConstraint
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.users.item.find_meeting_times.find_meeting_times_post_request_body import (
    FindMeetingTimesPostRequestBody,
)
from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
    GetSchedulePostRequestBody,
)

from .graph_client import get_graph_client, run_async
from .calendar_tools import get_local_timezone, _to_datetime, windows_tz_to_iana


# Duo scheduling policy: meetings start 5 minutes after each half-hour boundary
# (e.g., :05 or :35) to ensure buffer time between back-to-back meetings.
# A "30-minute" meeting is really 25 minutes; a "60-minute" meeting is really 55.
DUO_BUFFER_MINUTES = 5


def _snap_to_duo_boundary(dt: datetime) -> datetime:
    """Snap to the next :05 or :35 at or after dt."""
    base = dt.replace(second=0, microsecond=0)
    minute = base.minute
    if minute <= 5:
        return base.replace(minute=5)
    elif minute <= 35:
        return base.replace(minute=35)
    else:
        return (base + timedelta(hours=1)).replace(minute=5)


def _overlaps_any_lunch(
    start: datetime, end: datetime, participant_tzs: dict[str, ZoneInfo]
) -> bool:
    """Check if a meeting slot overlaps noon-1pm in any participant's timezone."""
    for tz in participant_tzs.values():
        local_start = start.astimezone(tz)
        local_end = end.astimezone(tz)
        lunch_start = local_start.replace(hour=12, minute=0, second=0, microsecond=0)
        lunch_end = local_start.replace(hour=13, minute=0, second=0, microsecond=0)
        if local_start < lunch_end and local_end > lunch_start:
            return True
    return False


async def _get_participant_tzs(
    participants: list[str], client
) -> dict[str, ZoneInfo]:
    """Fetch IANA timezones for participants (best-effort)."""
    from .people_tools import async_get_user_timezones

    windows_tzs = await async_get_user_timezones(participants, client=client)
    result = {}
    for email, win_tz in windows_tzs.items():
        if win_tz:
            iana = windows_tz_to_iana(win_tz)
            if iana:
                try:
                    result[email] = ZoneInfo(iana)
                except Exception:
                    pass
    return result


def _apply_duo_policy_to_results(
    results: list[dict[str, Any]],
    duration_minutes: int,
    local_tz: ZoneInfo,
    participant_tzs: dict[str, ZoneInfo] | None = None,
) -> list[dict[str, Any]]:
    """Apply Duo 5-minute buffer and lunch avoidance to API suggestions."""
    duo_duration = duration_minutes - DUO_BUFFER_MINUTES
    non_lunch = []
    lunch_slots = []
    seen: set[str] = set()

    for result in results:
        start_info = result.get("start")
        end_info = result.get("end")
        if not start_info or not end_info:
            continue

        try:
            raw_start = datetime.fromisoformat(start_info["date_time"])
            raw_end = datetime.fromisoformat(end_info["date_time"])
        except (KeyError, ValueError):
            continue

        if raw_start.tzinfo is None:
            raw_start = raw_start.replace(tzinfo=local_tz)
        if raw_end.tzinfo is None:
            raw_end = raw_end.replace(tzinfo=local_tz)

        snapped_start = _snap_to_duo_boundary(raw_start)
        duo_end = snapped_start + timedelta(minutes=duo_duration)

        # Verify Duo meeting fits within the original free window
        if duo_end > raw_end:
            continue

        # Deduplicate
        key = snapped_start.strftime("%Y-%m-%dT%H:%M")
        if key in seen:
            continue
        seen.add(key)

        adjusted = {
            **result,
            "start": {
                "date_time": snapped_start.strftime("%Y-%m-%dT%H:%M:%S"),
                "time_zone": str(local_tz),
                "formatted": snapped_start.strftime("%I:%M %p"),
                "date": snapped_start.strftime("%Y-%m-%d"),
            },
            "end": {
                "date_time": duo_end.strftime("%Y-%m-%dT%H:%M:%S"),
                "time_zone": str(local_tz),
                "formatted": duo_end.strftime("%I:%M %p"),
                "date": duo_end.strftime("%Y-%m-%d"),
            },
        }

        if participant_tzs and _overlaps_any_lunch(snapped_start, duo_end, participant_tzs):
            adjusted["lunch_overlap"] = True
            lunch_slots.append(adjusted)
        else:
            non_lunch.append(adjusted)

    return non_lunch + lunch_slots


def find_meeting_times(
    participants: list[str],
    *,
    duration_minutes: int = 30,
    start: datetime | date | str | None = None,
    end: datetime | date | str | None = None,
    days_ahead: int = 7,
    business_hours_start: int = 9,
    business_hours_end: int = 17,
    include_weekends: bool = False,
    max_suggestions: int = 10,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Find available meeting times. Synchronous wrapper.

    See async_find_meeting_times() for full documentation.
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_find_meeting_times(
        participants, duration_minutes=duration_minutes, start=start, end=end,
        days_ahead=days_ahead, business_hours_start=business_hours_start,
        business_hours_end=business_hours_end, include_weekends=include_weekends,
        max_suggestions=max_suggestions, timezone=timezone, client=client
    ))


async def async_find_meeting_times(
    participants: list[str],
    *,
    duration_minutes: int = 30,
    start: datetime | date | str | None = None,
    end: datetime | date | str | None = None,
    days_ahead: int = 7,
    business_hours_start: int = 9,
    business_hours_end: int = 17,
    include_weekends: bool = False,
    max_suggestions: int = 10,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Find available meeting times for all participants.

    Attempts to use the findMeetingTimes API first for intelligent suggestions
    with confidence scores. Falls back to manual slot finding using getSchedule
    if the API fails or returns no results.

    Args:
        participants: List of email addresses for meeting attendees.
        duration_minutes: Meeting duration in minutes (default: 30).
        start: Start of search range. Defaults to now.
        end: End of search range. Defaults to start + days_ahead.
        days_ahead: Days to search ahead if end not specified (default: 7).
        business_hours_start: Start hour for business hours (default: 9).
        business_hours_end: End hour for business hours (default: 17).
        include_weekends: Whether to include weekend slots (default: False).
        max_suggestions: Maximum number of time suggestions (default: 10).
        timezone: Timezone for times. Auto-detects if None.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of time slot dictionaries sorted by preference.
    """
    if not participants:
        return []

    if client is None:
        client = get_graph_client()

    # Resolve timezone
    if timezone is None:
        local_tz = get_local_timezone()
    elif isinstance(timezone, str):
        local_tz = ZoneInfo(timezone)
    else:
        local_tz = timezone

    if duration_minutes <= DUO_BUFFER_MINUTES:
        raise ValueError(
            f"Duration must be greater than {DUO_BUFFER_MINUTES} minutes "
            f"(the Duo buffer). Got {duration_minutes}."
        )

    # Set up time range
    now = datetime.now(local_tz)
    if start is None:
        start_dt = now
    else:
        start_dt = _to_datetime(start, local_tz)

    if end is None:
        end_dt = start_dt + timedelta(days=days_ahead)
    else:
        end_dt = _to_datetime(end, local_tz)

    # Fetch participant timezones for lunch avoidance (best-effort)
    participant_tzs: dict[str, ZoneInfo] = {}
    try:
        participant_tzs = await _get_participant_tzs(participants, client)
    except Exception:
        pass

    # Try findMeetingTimes API first
    try:
        results = await _async_find_meeting_times_api(
            participants=participants,
            duration_minutes=duration_minutes,
            start=start_dt,
            end=end_dt,
            max_suggestions=max_suggestions,
            local_tz=local_tz,
            client=client
        )
        if results:
            filtered = _apply_duo_policy_to_results(
                results, duration_minutes, local_tz, participant_tzs
            )[:max_suggestions]
            if filtered:
                return filtered
    except ODataError:
        pass  # Fall through to manual method

    # Fallback: manual slot finding using getSchedule
    return await _async_find_meeting_times_manual(
        participants=participants,
        duration_minutes=duration_minutes,
        start=start_dt,
        end=end_dt,
        business_hours_start=business_hours_start,
        business_hours_end=business_hours_end,
        include_weekends=include_weekends,
        max_suggestions=max_suggestions,
        local_tz=local_tz,
        client=client,
        participant_tzs=participant_tzs,
    )


def check_participant_availability(
    participants: list[str],
    start: datetime | date | str,
    end: datetime | date | str,
    *,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Check raw availability for participants. Synchronous wrapper.

    See async_check_participant_availability() for full documentation.
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_check_participant_availability(
        participants, start, end, timezone=timezone, client=client
    ))


async def async_check_participant_availability(
    participants: list[str],
    start: datetime | date | str,
    end: datetime | date | str,
    *,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Check raw availability for participants in a time range.

    Uses the getSchedule API to check free/busy status for multiple
    participants in a single request.

    Args:
        participants: List of email addresses to check.
        start: Start of time range.
        end: End of time range.
        timezone: Timezone for times. Auto-detects if None.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of availability dictionaries, one per participant.
    """
    if not participants:
        return []

    if client is None:
        client = get_graph_client()

    # Resolve timezone
    if timezone is None:
        local_tz = get_local_timezone()
    elif isinstance(timezone, str):
        local_tz = ZoneInfo(timezone)
    else:
        local_tz = timezone

    # Convert start/end to datetime in local timezone
    start_dt = _to_datetime(start, local_tz)
    end_dt = _to_datetime(end, local_tz)

    # Build request body for getSchedule
    request_body = GetSchedulePostRequestBody(
        schedules=participants,
        start_time=DateTimeTimeZone(
            date_time=start_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=str(local_tz),
        ),
        end_time=DateTimeTimeZone(
            date_time=end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=str(local_tz),
        ),
        availability_view_interval=15,  # 15-minute intervals
    )

    # Call getSchedule
    response = await client.me.calendar.get_schedule.post(request_body)

    results = []
    for schedule in response.value:
        # Parse availability view - all 0s means fully available
        availability_view = schedule.availability_view or ""
        is_available = all(c == '0' for c in availability_view) if availability_view else True

        # Extract busy times
        busy_times = []
        if schedule.schedule_items:
            for item in schedule.schedule_items:
                busy_times.append({
                    "start": _schedule_datetime_to_dict(item.start, local_tz),
                    "end": _schedule_datetime_to_dict(item.end, local_tz),
                    "status": item.status.value if item.status else None,
                    "subject": item.subject,
                    "location": item.location,
                    "is_private": item.is_private,
                })

        results.append({
            "email": schedule.schedule_id,
            "is_available": is_available,
            "availability_view": availability_view,
            "busy_times": busy_times,
            "working_hours": {
                "days_of_week": [d.value for d in schedule.working_hours.days_of_week] if schedule.working_hours and schedule.working_hours.days_of_week else [],
                "start_time": str(schedule.working_hours.start_time) if schedule.working_hours and schedule.working_hours.start_time else None,
                "end_time": str(schedule.working_hours.end_time) if schedule.working_hours and schedule.working_hours.end_time else None,
                "time_zone": schedule.working_hours.time_zone.name if schedule.working_hours and schedule.working_hours.time_zone else None,
            } if schedule.working_hours else None,
        })

    return results


# Private helper functions

async def _async_find_meeting_times_api(
    participants: list[str],
    duration_minutes: int,
    start: datetime,
    end: datetime,
    max_suggestions: int,
    local_tz: ZoneInfo,
    client
) -> list[dict[str, Any]]:
    """Use the findMeetingTimes API to get intelligent meeting time suggestions."""
    attendees = [
        AttendeeBase(
            email_address=EmailAddress(address=email),
            type=AttendeeType.Required
        )
        for email in participants
    ]

    time_constraint = TimeConstraint(
        time_slots=[
            TimeSlot(
                start=DateTimeTimeZone(
                    date_time=start.strftime("%Y-%m-%dT%H:%M:%S"),
                    time_zone=str(local_tz)
                ),
                end=DateTimeTimeZone(
                    date_time=end.strftime("%Y-%m-%dT%H:%M:%S"),
                    time_zone=str(local_tz)
                )
            )
        ]
    )

    request_body = FindMeetingTimesPostRequestBody(
        attendees=attendees,
        time_constraint=time_constraint,
        meeting_duration=timedelta(minutes=duration_minutes),
        max_candidates=max_suggestions,
        return_suggestion_reasons=True,
        minimum_attendee_percentage=100.0,
    )

    response = await client.me.find_meeting_times.post(request_body)

    results = []
    if response.meeting_time_suggestions:
        for suggestion in response.meeting_time_suggestions:
            slot = _suggestion_to_dict(suggestion, local_tz, participants)
            results.append(slot)

    return results


async def _async_find_meeting_times_manual(
    participants: list[str],
    duration_minutes: int,
    start: datetime,
    end: datetime,
    business_hours_start: int,
    business_hours_end: int,
    include_weekends: bool,
    max_suggestions: int,
    local_tz: ZoneInfo,
    client,
    participant_tzs: dict[str, ZoneInfo] | None = None,
) -> list[dict[str, Any]]:
    """Find meeting times manually using getSchedule API.

    This fallback method:
    1. Gets availability for all participants
    2. Parses availability views to find common free slots
    3. Filters by business hours, weekends, and Duo scheduling policy
    """
    # Align start to hour boundary so 15-min intervals land on :00/:15/:30/:45
    aligned_start = start.replace(minute=0, second=0, microsecond=0)

    request_body = GetSchedulePostRequestBody(
        schedules=participants,
        start_time=DateTimeTimeZone(
            date_time=aligned_start.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=str(local_tz),
        ),
        end_time=DateTimeTimeZone(
            date_time=end.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=str(local_tz),
        ),
        availability_view_interval=15,
    )

    response = await client.me.calendar.get_schedule.post(request_body)

    schedules = []
    for schedule in response.value:
        schedules.append({
            "email": schedule.schedule_id,
            "availability_view": schedule.availability_view or "",
        })

    if not schedules:
        return []

    common_free = _find_common_free_slots(
        schedules=schedules,
        duration_minutes=duration_minutes,
        aligned_start=aligned_start,
        min_start=start,
        business_hours_start=business_hours_start,
        business_hours_end=business_hours_end,
        include_weekends=include_weekends,
        local_tz=local_tz,
        participant_tzs=participant_tzs,
    )

    results = []
    for slot in common_free[:max_suggestions]:
        slot_start = slot["start"]
        slot_end = slot["end"]

        result = {
            "start": {
                "date_time": slot_start.strftime("%Y-%m-%dT%H:%M:%S"),
                "time_zone": str(local_tz),
                "formatted": slot_start.strftime("%I:%M %p"),
                "date": slot_start.strftime("%Y-%m-%d"),
            },
            "end": {
                "date_time": slot_end.strftime("%Y-%m-%dT%H:%M:%S"),
                "time_zone": str(local_tz),
                "formatted": slot_end.strftime("%I:%M %p"),
                "date": slot_end.strftime("%Y-%m-%d"),
            },
            "confidence": 1.0,
            "all_available": True,
            "attendee_availability": [
                {"email": s["email"], "available": True}
                for s in schedules
            ],
        }
        if slot.get("lunch_overlap"):
            result["lunch_overlap"] = True
        results.append(result)

    return results


def _find_common_free_slots(
    schedules: list[dict[str, Any]],
    duration_minutes: int,
    aligned_start: datetime,
    min_start: datetime,
    business_hours_start: int,
    business_hours_end: int,
    include_weekends: bool,
    local_tz: ZoneInfo,
    participant_tzs: dict[str, ZoneInfo] | None = None,
) -> list[dict[str, datetime]]:
    """Find time slots where all participants are free, using Duo scheduling.

    Parses the availability_view strings (15-min intervals, 0=free) and finds
    free blocks aligned to Duo boundaries (:05 and :35 starts). Meetings are
    5 minutes shorter than the requested duration to provide buffer time.
    Slots overlapping any participant's lunch (noon-1pm local) are sorted last.
    """
    if not schedules:
        return []

    # Get the shortest availability view length
    min_len = min(len(s["availability_view"]) for s in schedules)
    if min_len == 0:
        return []

    # Combine availability - free only if ALL participants are free
    combined = []
    for i in range(min_len):
        all_free = all(
            s["availability_view"][i] == '0' if i < len(s["availability_view"]) else False
            for s in schedules
        )
        combined.append(all_free)

    # Find consecutive free slots of required duration
    slots_needed = -(-duration_minutes // 15)  # ceiling division
    if slots_needed < 1:
        slots_needed = 1

    duo_duration = duration_minutes - DUO_BUFFER_MINUTES
    non_lunch = []
    lunch_slots = []

    for i in range(len(combined) - slots_needed + 1):
        # Only consider starts at :00 or :30 (Duo meetings start at :05 or :35)
        interval_time = aligned_start + timedelta(minutes=i * 15)
        if interval_time.minute not in (0, 30):
            continue

        # Check all intervals for the requested duration are free
        if not all(combined[i:i + slots_needed]):
            continue

        # Duo-adjusted times: start 5 min after boundary, shorter duration
        slot_start = interval_time + timedelta(minutes=DUO_BUFFER_MINUTES)
        slot_end = slot_start + timedelta(minutes=duo_duration)

        # Skip slots before the minimum start time (e.g., in the past)
        if slot_start < min_start:
            continue

        # Check business hours and weekends
        if not _is_valid_meeting_time(
            slot_start, slot_end,
            business_hours_start, business_hours_end,
            include_weekends
        ):
            continue

        slot: dict[str, Any] = {"start": slot_start, "end": slot_end}

        # Check lunch overlap — prefer non-lunch, but keep lunch as fallback
        if participant_tzs and _overlaps_any_lunch(slot_start, slot_end, participant_tzs):
            slot["lunch_overlap"] = True
            lunch_slots.append(slot)
        else:
            non_lunch.append(slot)

    return non_lunch + lunch_slots


def _is_valid_meeting_time(
    start: datetime,
    end: datetime,
    business_hours_start: int,
    business_hours_end: int,
    include_weekends: bool
) -> bool:
    """Check if a time slot is within valid meeting hours."""
    # Check weekends
    if not include_weekends:
        if start.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return False

    # Check business hours (meeting must fit entirely within business hours)
    if start.hour < business_hours_start:
        return False
    if end.hour > business_hours_end:
        return False
    if end.hour == business_hours_end and end.minute > 0:
        return False

    return True


def _suggestion_to_dict(
    suggestion,
    local_tz: ZoneInfo,
    participants: list[str]
) -> dict[str, Any]:
    """Convert findMeetingTimes suggestion to result dictionary."""
    # Parse meeting time slot
    meeting_time = suggestion.meeting_time_slot
    start_dict = None
    end_dict = None

    if meeting_time:
        if meeting_time.start:
            start_dict = _datetime_zone_to_dict(meeting_time.start, local_tz)
        if meeting_time.end:
            end_dict = _datetime_zone_to_dict(meeting_time.end, local_tz)

    # Calculate confidence from suggestion confidence
    # API returns confidence as percentage (0-100), normalize to 0-1 for consistency
    confidence = 1.0
    if suggestion.confidence is not None:
        # API returns values like 100.0 for 100% confidence
        confidence = suggestion.confidence / 100.0 if suggestion.confidence > 1 else suggestion.confidence

    # Build attendee availability
    attendee_availability = []
    if suggestion.attendee_availability:
        for att_avail in suggestion.attendee_availability:
            email = att_avail.attendee.email_address.address if att_avail.attendee and att_avail.attendee.email_address else None
            available = att_avail.availability.value == "free" if att_avail.availability else True
            attendee_availability.append({
                "email": email,
                "available": available,
                "status": att_avail.availability.value if att_avail.availability else None,
            })
    else:
        # If no attendee availability info, assume all free
        attendee_availability = [{"email": p, "available": True} for p in participants]

    all_available = all(a["available"] for a in attendee_availability)

    return {
        "start": start_dict,
        "end": end_dict,
        "confidence": confidence,
        "all_available": all_available,
        "attendee_availability": attendee_availability,
        "suggestion_reason": suggestion.suggestion_reason,
    }


def _datetime_zone_to_dict(dt_zone, local_tz: ZoneInfo) -> dict[str, str] | None:
    """Convert DateTimeTimeZone to dictionary with local time."""
    if dt_zone is None or dt_zone.date_time is None:
        return None

    # Parse the datetime string
    dt_str = dt_zone.date_time
    # Handle various formats
    dt_str = dt_str.rstrip('0').rstrip('.')

    try:
        dt = datetime.fromisoformat(dt_str)
    except ValueError:
        return None

    # If the response includes timezone info, use it; otherwise assume local
    if dt.tzinfo is None:
        # Check if timezone was specified in the response
        if dt_zone.time_zone:
            try:
                response_tz = ZoneInfo(dt_zone.time_zone)
                dt = dt.replace(tzinfo=response_tz)
            except KeyError:
                dt = dt.replace(tzinfo=local_tz)
        else:
            dt = dt.replace(tzinfo=local_tz)

    # Convert to local timezone
    local_dt = dt.astimezone(local_tz)

    return {
        "date_time": local_dt.strftime("%Y-%m-%dT%H:%M:%S"),
        "time_zone": str(local_tz),
        "formatted": local_dt.strftime("%I:%M %p"),
        "date": local_dt.strftime("%Y-%m-%d"),
    }


def _schedule_datetime_to_dict(dt_zone, local_tz: ZoneInfo) -> dict[str, str] | None:
    """Convert schedule DateTimeTimeZone to dictionary with local time.

    The getSchedule API returns times in the requested timezone.
    """
    if dt_zone is None:
        return None

    # Parse the datetime string
    dt_str = dt_zone.date_time
    if not dt_str:
        return None

    # Remove any trailing zeros from microseconds
    dt_str = dt_str.rstrip('0').rstrip('.')
    dt = datetime.fromisoformat(dt_str)

    # The response is in the requested timezone, so apply it
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=local_tz)

    return {
        "date_time": dt.strftime("%Y-%m-%dT%H:%M:%S"),
        "time_zone": str(local_tz),
        "formatted": dt.strftime("%I:%M %p"),
        "date": dt.strftime("%Y-%m-%d"),
    }
