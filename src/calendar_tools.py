"""Calendar tools for Microsoft Graph API."""

import subprocess
from datetime import datetime, date, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
    CalendarViewRequestBuilder,
)
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from .graph_client import get_graph_client, run_async


# Mapping of Windows timezone names (from Graph API mailboxSettings.timeZone)
# to IANA timezone names (used by Python zoneinfo).
WINDOWS_TO_IANA: dict[str, str] = {
    "Dateline Standard Time": "Etc/GMT+12",
    "UTC-11": "Etc/GMT+11",
    "Aleutian Standard Time": "America/Adak",
    "Hawaiian Standard Time": "Pacific/Honolulu",
    "Marquesas Standard Time": "Pacific/Marquesas",
    "Alaskan Standard Time": "America/Anchorage",
    "UTC-09": "Etc/GMT+9",
    "Pacific Standard Time (Mexico)": "America/Tijuana",
    "UTC-08": "Etc/GMT+8",
    "Pacific Standard Time": "America/Los_Angeles",
    "US Mountain Standard Time": "America/Phoenix",
    "Mountain Standard Time (Mexico)": "America/Chihuahua",
    "Mountain Standard Time": "America/Denver",
    "Central America Standard Time": "America/Guatemala",
    "Central Standard Time": "America/Chicago",
    "Easter Island Standard Time": "Pacific/Easter",
    "Central Standard Time (Mexico)": "America/Mexico_City",
    "Canada Central Standard Time": "America/Regina",
    "SA Pacific Standard Time": "America/Bogota",
    "Eastern Standard Time (Mexico)": "America/Cancun",
    "Eastern Standard Time": "America/New_York",
    "Haiti Standard Time": "America/Port-au-Prince",
    "Cuba Standard Time": "America/Havana",
    "US Eastern Standard Time": "America/Indianapolis",
    "Turks And Caicos Standard Time": "America/Grand_Turk",
    "Paraguay Standard Time": "America/Asuncion",
    "Atlantic Standard Time": "America/Halifax",
    "Venezuela Standard Time": "America/Caracas",
    "Central Brazilian Standard Time": "America/Cuiaba",
    "SA Western Standard Time": "America/La_Paz",
    "Pacific SA Standard Time": "America/Santiago",
    "Newfoundland Standard Time": "America/St_Johns",
    "Tocantins Standard Time": "America/Araguaina",
    "E. South America Standard Time": "America/Sao_Paulo",
    "SA Eastern Standard Time": "America/Cayenne",
    "Argentina Standard Time": "America/Buenos_Aires",
    "Greenland Standard Time": "America/Godthab",
    "Montevideo Standard Time": "America/Montevideo",
    "Bahia Standard Time": "America/Bahia",
    "UTC-02": "Etc/GMT+2",
    "Azores Standard Time": "Atlantic/Azores",
    "Cape Verde Standard Time": "Atlantic/Cape_Verde",
    "UTC": "Etc/UTC",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central Europe Standard Time": "Europe/Budapest",
    "Romance Standard Time": "Europe/Paris",
    "Central European Standard Time": "Europe/Warsaw",
    "W. Central Africa Standard Time": "Africa/Lagos",
    "Jordan Standard Time": "Asia/Amman",
    "GTB Standard Time": "Europe/Bucharest",
    "Middle East Standard Time": "Asia/Beirut",
    "Egypt Standard Time": "Africa/Cairo",
    "E. Europe Standard Time": "Europe/Chisinau",
    "Syria Standard Time": "Asia/Damascus",
    "West Bank Standard Time": "Asia/Hebron",
    "South Africa Standard Time": "Africa/Johannesburg",
    "FLE Standard Time": "Europe/Kiev",
    "Israel Standard Time": "Asia/Jerusalem",
    "Kaliningrad Standard Time": "Europe/Kaliningrad",
    "Sudan Standard Time": "Africa/Khartoum",
    "Libya Standard Time": "Africa/Tripoli",
    "Namibia Standard Time": "Africa/Windhoek",
    "Arabic Standard Time": "Asia/Baghdad",
    "Turkey Standard Time": "Europe/Istanbul",
    "Arab Standard Time": "Asia/Riyadh",
    "Belarus Standard Time": "Europe/Minsk",
    "Russian Standard Time": "Europe/Moscow",
    "E. Africa Standard Time": "Africa/Nairobi",
    "Iran Standard Time": "Asia/Tehran",
    "Arabian Standard Time": "Asia/Dubai",
    "Astrakhan Standard Time": "Europe/Astrakhan",
    "Azerbaijan Standard Time": "Asia/Baku",
    "Russia Time Zone 3": "Europe/Samara",
    "Mauritius Standard Time": "Indian/Mauritius",
    "Saratov Standard Time": "Europe/Saratov",
    "Georgian Standard Time": "Asia/Tbilisi",
    "Volgograd Standard Time": "Europe/Volgograd",
    "Caucasus Standard Time": "Asia/Yerevan",
    "Afghanistan Standard Time": "Asia/Kabul",
    "West Asia Standard Time": "Asia/Tashkent",
    "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
    "Pakistan Standard Time": "Asia/Karachi",
    "India Standard Time": "Asia/Calcutta",
    "Sri Lanka Standard Time": "Asia/Colombo",
    "Nepal Standard Time": "Asia/Katmandu",
    "Central Asia Standard Time": "Asia/Almaty",
    "Bangladesh Standard Time": "Asia/Dhaka",
    "Omsk Standard Time": "Asia/Omsk",
    "Myanmar Standard Time": "Asia/Rangoon",
    "SE Asia Standard Time": "Asia/Bangkok",
    "Altai Standard Time": "Asia/Barnaul",
    "W. Mongolia Standard Time": "Asia/Hovd",
    "North Asia Standard Time": "Asia/Krasnoyarsk",
    "N. Central Asia Standard Time": "Asia/Novosibirsk",
    "Tomsk Standard Time": "Asia/Tomsk",
    "China Standard Time": "Asia/Shanghai",
    "North Asia East Standard Time": "Asia/Irkutsk",
    "Singapore Standard Time": "Asia/Singapore",
    "W. Australia Standard Time": "Australia/Perth",
    "Taipei Standard Time": "Asia/Taipei",
    "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
    "Aus Central W. Standard Time": "Australia/Eucla",
    "Transbaikal Standard Time": "Asia/Chita",
    "Tokyo Standard Time": "Asia/Tokyo",
    "North Korea Standard Time": "Asia/Pyongyang",
    "Korea Standard Time": "Asia/Seoul",
    "Yakutsk Standard Time": "Asia/Yakutsk",
    "Cen. Australia Standard Time": "Australia/Adelaide",
    "AUS Central Standard Time": "Australia/Darwin",
    "E. Australia Standard Time": "Australia/Brisbane",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "West Pacific Standard Time": "Pacific/Port_Moresby",
    "Tasmania Standard Time": "Australia/Hobart",
    "Vladivostok Standard Time": "Asia/Vladivostok",
    "Lord Howe Standard Time": "Australia/Lord_Howe",
    "Bougainville Standard Time": "Pacific/Bougainville",
    "Russia Time Zone 10": "Asia/Srednekolymsk",
    "Magadan Standard Time": "Asia/Magadan",
    "Norfolk Standard Time": "Pacific/Norfolk",
    "Sakhalin Standard Time": "Asia/Sakhalin",
    "Central Pacific Standard Time": "Pacific/Guadalcanal",
    "Russia Time Zone 11": "Asia/Kamchatka",
    "New Zealand Standard Time": "Pacific/Auckland",
    "UTC+12": "Etc/GMT-12",
    "Fiji Standard Time": "Pacific/Fiji",
    "Chatham Islands Standard Time": "Pacific/Chatham",
    "UTC+13": "Etc/GMT-13",
    "Tonga Standard Time": "Pacific/Tongatapu",
    "Samoa Standard Time": "Pacific/Apia",
    "Line Islands Standard Time": "Pacific/Kiritimati",
}


# Reverse mapping: IANA -> Windows (built from WINDOWS_TO_IANA)
IANA_TO_WINDOWS: dict[str, str] = {v: k for k, v in WINDOWS_TO_IANA.items()}


def iana_tz_to_windows(iana_name: str) -> str | None:
    """Convert an IANA timezone name to a Windows timezone name.

    Args:
        iana_name: IANA timezone name (e.g., "America/New_York").

    Returns:
        Windows timezone name (e.g., "Eastern Standard Time"), or None if not found.
    """
    return IANA_TO_WINDOWS.get(iana_name)


def windows_tz_to_iana(windows_name: str) -> str | None:
    """Convert a Windows timezone name to an IANA timezone name.

    Args:
        windows_name: Windows timezone name (e.g., "Eastern Standard Time").

    Returns:
        IANA timezone name (e.g., "America/New_York"), or None if not found.
    """
    return WINDOWS_TO_IANA.get(windows_name)


def get_timezone_abbr(tz: ZoneInfo, dt: datetime) -> str:
    """Get the timezone abbreviation for a given timezone at a specific datetime.

    Args:
        tz: ZoneInfo timezone object.
        dt: A datetime to determine the abbreviation (accounts for DST).

    Returns:
        Timezone abbreviation string (e.g., "EST", "PST", "CDT").
    """
    return dt.astimezone(tz).strftime("%Z")


def get_local_timezone() -> ZoneInfo:
    """Detect the system's local timezone.

    Returns:
        ZoneInfo object for the local timezone.

    Raises:
        RuntimeError: If timezone cannot be detected.
    """
    # macOS: read symlink target of /etc/localtime
    try:
        result = subprocess.run(
            ['readlink', '/etc/localtime'],
            capture_output=True,
            text=True,
            check=True
        )
        tz_path = result.stdout.strip()
        # Extract timezone name from path like /var/db/timezone/zoneinfo/America/Detroit
        parts = tz_path.split('/')
        # Find 'zoneinfo' and take everything after it
        if 'zoneinfo' in parts:
            idx = parts.index('zoneinfo')
            tz_name = '/'.join(parts[idx + 1:])
        else:
            # Fallback: take last two parts
            tz_name = '/'.join(parts[-2:])
        return ZoneInfo(tz_name)
    except (subprocess.CalledProcessError, KeyError, IndexError) as e:
        raise RuntimeError(f"Could not detect local timezone: {e}")


def get_events_in_range(
    start: datetime | date | str,
    end: datetime | date | str,
    timezone: ZoneInfo | str | None = None,
    client=None,
) -> list[dict[str, Any]]:
    """Get all calendar events overlapping with the specified time range.

    Synchronous wrapper around async_get_events_in_range(). For use in scripts
    and direct Python usage. NOT safe to call from an async context (e.g., MCP server).

    See async_get_events_in_range() for full documentation.
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_get_events_in_range(start, end, timezone=timezone, client=client))


async def async_get_events_in_range(
    start: datetime | date | str,
    end: datetime | date | str,
    timezone: ZoneInfo | str | None = None,
    client=None,
) -> list[dict[str, Any]]:
    """Get all calendar events overlapping with the specified time range.

    This uses the calendarView endpoint which automatically handles recurring
    events and returns all event instances that overlap with the time range,
    including events that started before the range but extend into it.

    Args:
        start: Start of time range. Can be:
            - date object: Uses midnight in the specified timezone
            - datetime object: Uses the datetime as-is (timezone-aware recommended)
            - ISO 8601 string: Parsed as local time if no timezone specified
        end: End of time range (same format options as start).
        timezone: Timezone for interpreting start/end and for output times.
            Can be a ZoneInfo object or timezone string (e.g., "America/Detroit").
            If None, auto-detects from system. All returned times will be in this timezone.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of event dictionaries with full metadata. Times are converted to
        the specified timezone.

    Raises:
        ODataError: If the Graph API request fails.
        ValueError: If no valid token is available.
        RuntimeError: If timezone cannot be detected.
    """
    if client is None:
        client = get_graph_client()

    # Resolve timezone
    if timezone is None:
        local_tz = get_local_timezone()
    elif isinstance(timezone, str):
        local_tz = ZoneInfo(timezone)
    else:
        local_tz = timezone

    # Convert start to datetime in local timezone
    start_dt = _to_datetime(start, local_tz)
    end_dt = _to_datetime(end, local_tz)

    # Convert to UTC for API call
    utc = ZoneInfo('UTC')
    start_utc = start_dt.astimezone(utc)
    end_utc = end_dt.astimezone(utc)

    start_str = start_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end_utc.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Use calendarView which handles recurring events and overlap automatically
    config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
        query_parameters=CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
            start_date_time=start_str,
            end_date_time=end_str,
            select=[
                "id",
                "subject",
                "start",
                "end",
                "location",
                "locations",
                "attendees",
                "organizer",
                "isAllDay",
                "isCancelled",
                "isOnlineMeeting",
                "onlineMeetingUrl",
                "onlineMeeting",
                "bodyPreview",
                "body",
                "importance",
                "sensitivity",
                "showAs",
                "responseStatus",
                "categories",
                "webLink",
                "seriesMasterId",
                "recurrence",
                "reminderMinutesBeforeStart",
                "isReminderOn",
                "hasAttachments",
                "createdDateTime",
                "lastModifiedDateTime",
            ],
            orderby=["start/dateTime"],
            top=100,
        )
    )

    events_response = await client.me.calendar_view.get(request_configuration=config)

    events = []
    for event in events_response.value:
        events.append(_event_to_dict(event, local_tz))

    # Handle pagination if there are more events
    next_link = events_response.odata_next_link
    while next_link:
        # Fetch next page using the next link
        next_response = await client.me.calendar_view.with_url(next_link).get()
        for event in next_response.value:
            events.append(_event_to_dict(event, local_tz))
        next_link = next_response.odata_next_link

    return events


def _to_datetime(value: datetime | date | str, tz: ZoneInfo) -> datetime:
    """Convert various input types to a timezone-aware datetime."""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            # Naive datetime - assume it's in the specified timezone
            return value.replace(tzinfo=tz)
        return value
    elif isinstance(value, date):
        # Date object - use midnight in the specified timezone
        return datetime(value.year, value.month, value.day, 0, 0, 0, tzinfo=tz)
    elif isinstance(value, str):
        # ISO string - parse it
        # Remove trailing Z if present and treat as UTC
        if value.endswith('Z'):
            dt = datetime.fromisoformat(value[:-1])
            return dt.replace(tzinfo=ZoneInfo('UTC'))
        # Try parsing as ISO format
        dt = datetime.fromisoformat(value)
        if dt.tzinfo is None:
            return dt.replace(tzinfo=tz)
        return dt
    else:
        raise TypeError(f"Cannot convert {type(value)} to datetime")


def _event_to_dict(event, local_tz: ZoneInfo) -> dict[str, Any]:
    """Convert a Graph API Event object to a dictionary with local times."""
    result = {
        "id": event.id,
        "subject": event.subject,
        "start": _datetime_zone_to_dict(event.start, local_tz),
        "end": _datetime_zone_to_dict(event.end, local_tz),
        "is_all_day": event.is_all_day,
        "is_cancelled": event.is_cancelled,
        "body_preview": event.body_preview,
        "importance": event.importance.value if event.importance else None,
        "sensitivity": event.sensitivity.value if event.sensitivity else None,
        "show_as": event.show_as.value if event.show_as else None,
        "web_link": event.web_link,
        "has_attachments": event.has_attachments,
        "created_date_time": _datetime_to_str(event.created_date_time),
        "last_modified_date_time": _datetime_to_str(event.last_modified_date_time),
        "is_reminder_on": event.is_reminder_on,
        "reminder_minutes_before_start": event.reminder_minutes_before_start,
        "series_master_id": event.series_master_id,
        "categories": event.categories or [],
    }

    # Calculate duration
    if event.start and event.end:
        start_dt = _parse_graph_datetime(event.start)
        end_dt = _parse_graph_datetime(event.end)
        if start_dt and end_dt:
            duration = end_dt - start_dt
            result["duration_minutes"] = int(duration.total_seconds() / 60)

    # Location
    if event.location:
        result["location"] = {
            "display_name": event.location.display_name,
            "location_type": event.location.location_type.value if event.location.location_type else None,
        }
    else:
        result["location"] = None

    # Multiple locations
    if event.locations:
        result["locations"] = [
            {
                "display_name": loc.display_name,
                "location_type": loc.location_type.value if loc.location_type else None,
            }
            for loc in event.locations
        ]
    else:
        result["locations"] = []

    # Organizer
    if event.organizer and event.organizer.email_address:
        result["organizer"] = {
            "name": event.organizer.email_address.name,
            "address": event.organizer.email_address.address,
        }
    else:
        result["organizer"] = None

    # Attendees
    if event.attendees:
        result["attendees"] = [
            {
                "name": att.email_address.name if att.email_address else None,
                "address": att.email_address.address if att.email_address else None,
                "type": att.type.value if att.type else None,
                "response_status": att.status.response.value if att.status and att.status.response else None,
            }
            for att in event.attendees
        ]
    else:
        result["attendees"] = []

    # Response status
    if event.response_status:
        result["response_status"] = {
            "response": event.response_status.response.value if event.response_status.response else None,
            "time": _datetime_to_str(event.response_status.time),
        }
    else:
        result["response_status"] = None

    # Online meeting
    result["is_online_meeting"] = event.is_online_meeting
    result["online_meeting_url"] = event.online_meeting_url
    if event.online_meeting:
        result["online_meeting"] = {
            "join_url": event.online_meeting.join_url,
            "conference_id": event.online_meeting.conference_id,
            "toll_number": event.online_meeting.toll_number,
        }
    else:
        result["online_meeting"] = None

    # Recurrence
    if event.recurrence:
        result["recurrence"] = {
            "pattern": _recurrence_pattern_to_dict(event.recurrence.pattern) if event.recurrence.pattern else None,
            "range": _recurrence_range_to_dict(event.recurrence.range) if event.recurrence.range else None,
        }
    else:
        result["recurrence"] = None

    # Body (HTML content)
    if event.body:
        result["body"] = {
            "content_type": event.body.content_type.value if event.body.content_type else None,
            "content": event.body.content,
        }
    else:
        result["body"] = None

    return result


def _parse_graph_datetime(dt_zone) -> datetime | None:
    """Parse a Graph API DateTimeTimeZone to a datetime object."""
    if dt_zone is None or dt_zone.date_time is None:
        return None
    # Graph returns times in UTC when using calendarView
    dt_str = dt_zone.date_time.rstrip('0').rstrip('.')
    dt = datetime.fromisoformat(dt_str)
    return dt.replace(tzinfo=ZoneInfo('UTC'))


def _datetime_zone_to_dict(dt_zone, local_tz: ZoneInfo) -> dict[str, str] | None:
    """Convert DateTimeTimeZone to dictionary with local time."""
    if dt_zone is None:
        return None

    # Parse the UTC time from Graph
    dt_str = dt_zone.date_time.rstrip('0').rstrip('.')
    utc_dt = datetime.fromisoformat(dt_str).replace(tzinfo=ZoneInfo('UTC'))

    # Convert to local timezone
    local_dt = utc_dt.astimezone(local_tz)

    return {
        "date_time": local_dt.strftime("%Y-%m-%dT%H:%M:%S"),
        "time_zone": str(local_tz),
        "formatted": local_dt.strftime("%I:%M %p"),
        "date": local_dt.strftime("%Y-%m-%d"),
    }


def _datetime_to_str(dt) -> str | None:
    """Convert datetime to ISO string."""
    if dt is None:
        return None
    return dt.isoformat()


async def async_create_event(
    subject: str,
    start: datetime | str,
    end: datetime | str,
    body: str | None = None,
    show_as: str = "busy",
    timezone: ZoneInfo | str | None = None,
    is_all_day: bool = False,
    location: str | None = None,
    attendees: list[str] | None = None,
    client=None,
) -> dict[str, Any]:
    """Create a new calendar event.

    Args:
        subject: Event subject/title.
        start: Start time as datetime or ISO string (e.g., "2026-04-21T12:00").
        end: End time as datetime or ISO string.
        body: Optional event body (plain text).
        show_as: How to show during event: busy, free, tentative, oof, workingElsewhere.
        timezone: Timezone for interpreting start/end times.
        is_all_day: Whether this is an all-day event.
        location: Optional location string.
        attendees: Optional list of email addresses to invite.
        client: Optional GraphServiceClient.

    Returns:
        Dictionary with created event details.
    """
    from msgraph.generated.models.event import Event as EventModel
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.free_busy_status import FreeBusyStatus
    from msgraph.generated.models.location import Location as LocationModel

    if client is None:
        client = get_graph_client()

    # Resolve timezone
    if timezone is None:
        local_tz = get_local_timezone()
    elif isinstance(timezone, str):
        local_tz = ZoneInfo(timezone)
    else:
        local_tz = timezone

    iana_name = str(local_tz)
    # Graph API requires Windows timezone names for DateTimeTimeZone
    windows_tz_name = iana_tz_to_windows(iana_name) or iana_name

    # Parse start/end
    start_dt = _to_datetime(start, local_tz)
    end_dt = _to_datetime(end, local_tz)

    # Build event
    event = EventModel(
        subject=subject,
        start=DateTimeTimeZone(
            date_time=start_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=windows_tz_name,
        ),
        end=DateTimeTimeZone(
            date_time=end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            time_zone=windows_tz_name,
        ),
        is_all_day=is_all_day,
    )

    # Show as
    status_map = {
        "busy": FreeBusyStatus.Busy,
        "free": FreeBusyStatus.Free,
        "tentative": FreeBusyStatus.Tentative,
        "oof": FreeBusyStatus.Oof,
        "workingelsewhere": FreeBusyStatus.WorkingElsewhere,
    }
    event.show_as = status_map.get(show_as.lower(), FreeBusyStatus.Busy)

    if body:
        event.body = ItemBody(content_type=BodyType.Text, content=body)

    if location:
        event.location = LocationModel(display_name=location)

    if attendees:
        from msgraph.generated.models.attendee import Attendee
        from msgraph.generated.models.attendee_type import AttendeeType
        from msgraph.generated.models.email_address import EmailAddress

        event.attendees = [
            Attendee(
                email_address=EmailAddress(address=email),
                type=AttendeeType.Required,
            )
            for email in attendees
        ]

    created = await client.me.events.post(event)

    # The POST response returns times in the timezone we specified (not UTC).
    # Build result directly to avoid _event_to_dict's UTC assumption.
    start_resp = created.start
    end_resp = created.end
    return {
        "id": created.id,
        "subject": created.subject,
        "start": {
            "date_time": start_resp.date_time if start_resp else None,
            "time_zone": iana_name,
            "formatted": start_dt.strftime("%I:%M %p"),
            "date": start_dt.strftime("%Y-%m-%d"),
        } if start_resp else None,
        "end": {
            "date_time": end_resp.date_time if end_resp else None,
            "time_zone": iana_name,
            "formatted": end_dt.strftime("%I:%M %p"),
            "date": end_dt.strftime("%Y-%m-%d"),
        } if end_resp else None,
        "show_as": created.show_as.value if created.show_as else None,
    }


def create_event(
    subject: str,
    start: datetime | str,
    end: datetime | str,
    **kwargs,
) -> dict[str, Any]:
    """Create a new calendar event (synchronous wrapper).

    See async_create_event() for full documentation.
    """
    return run_async(async_create_event(subject, start, end, **kwargs))


def _recurrence_pattern_to_dict(pattern) -> dict[str, Any] | None:
    """Convert RecurrencePattern to dictionary."""
    if pattern is None:
        return None
    return {
        "type": pattern.type.value if pattern.type else None,
        "interval": pattern.interval,
        "days_of_week": [d.value for d in pattern.days_of_week] if pattern.days_of_week else [],
        "day_of_month": pattern.day_of_month,
        "month": pattern.month,
        "first_day_of_week": pattern.first_day_of_week.value if pattern.first_day_of_week else None,
        "index": pattern.index.value if pattern.index else None,
    }


def _recurrence_range_to_dict(range_obj) -> dict[str, Any] | None:
    """Convert RecurrenceRange to dictionary."""
    if range_obj is None:
        return None
    return {
        "type": range_obj.type.value if range_obj.type else None,
        "start_date": str(range_obj.start_date) if range_obj.start_date else None,
        "end_date": str(range_obj.end_date) if range_obj.end_date else None,
        "recurrence_time_zone": range_obj.recurrence_time_zone,
        "number_of_occurrences": range_obj.number_of_occurrences,
    }
