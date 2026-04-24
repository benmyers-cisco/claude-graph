"""Room availability tools for Microsoft Graph API."""

import re
from datetime import datetime, date
from typing import Any
from zoneinfo import ZoneInfo

import time

import requests

from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
    GetSchedulePostRequestBody,
)
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone

from .graph_client import get_graph_client, get_token, run_async
from .calendar_tools import get_local_timezone, _to_datetime


# Beta API functions (fallback when Place.Read.All is not available)

def _requests_get_with_retry(
    url: str,
    headers: dict[str, str],
    max_retries: int = 3,
) -> requests.Response:
    """HTTP GET with retry on 429/503/504, respecting Retry-After header."""
    for attempt in range(max_retries + 1):
        r = requests.get(url, headers=headers)
        if r.status_code not in (429, 503, 504) or attempt == max_retries:
            return r
        retry_after = r.headers.get("Retry-After")
        if retry_after:
            try:
                wait = int(retry_after)
            except ValueError:
                wait = 2 ** attempt
        else:
            wait = 2 ** attempt
        time.sleep(min(wait, 30))
    return r  # unreachable, but satisfies type checkers


def _get_auth_headers(token: str | None = None) -> dict[str, str]:
    """Get authorization headers for direct API calls."""
    if token is None:
        token = get_token()
    if not token:
        raise ValueError("No access token available")
    return {"Authorization": f"Bearer {token}"}


def _list_room_lists_beta(token: str | None = None) -> list[dict[str, Any]]:
    """List all room lists using the beta findRoomLists API.

    Returns:
        List of room list dictionaries with 'name' and 'address' (email).
    """
    headers = _get_auth_headers(token)
    url = "https://graph.microsoft.com/beta/me/findRoomLists"

    all_lists = []
    while url:
        r = _requests_get_with_retry(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        all_lists.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_lists


def _list_rooms_beta(
    room_list: str | None = None,
    token: str | None = None
) -> list[dict[str, Any]]:
    """List rooms using the beta findRooms API.

    Args:
        room_list: Optional room list email address to filter by.
                   If provided, only rooms in that list are returned.
        token: Optional access token. Uses environment variable if not provided.

    Returns:
        List of room dictionaries with 'name' and 'address' (email).
    """
    headers = _get_auth_headers(token)

    if room_list:
        url = f"https://graph.microsoft.com/beta/me/findRooms(RoomList='{room_list}')"
    else:
        url = "https://graph.microsoft.com/beta/me/findRooms"

    all_rooms = []
    while url:
        r = _requests_get_with_retry(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        all_rooms.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_rooms


def _beta_room_to_dict(room: dict[str, Any]) -> dict[str, Any]:
    """Convert a beta API room response to our standard dictionary format.

    Args:
        room: Room dict from beta findRooms API with 'name' and 'address'.

    Returns:
        Dictionary with room metadata including parsed building/floor.
    """
    display_name = room.get("name", "")
    email = room.get("address", "")
    parsed = _parse_room_name(display_name)

    # Extract capacity from display name (e.g., "(4)" in "ARB02-1-Island (4) (Public)")
    capacity = None
    cap_match = re.search(r"\((\d+)\)", display_name)
    if cap_match:
        capacity = int(cap_match.group(1))

    return {
        "email": email,
        "name": display_name,
        "building": parsed["building"],
        "floor": parsed["floor"],
        "room_name": parsed["room_name"],
        "capacity": capacity,
        # Fields not available from beta API
        "phone": None,
        "address": None,
        "geo_coordinates": None,
        "audio_device_name": None,
        "video_device_name": None,
        "display_device_name": None,
        "is_wheel_chair_accessible": None,
        "tags": [],
        "booking_type": None,
    }


def _parse_room_name(display_name: str) -> dict[str, Any]:
    """Parse room display name to extract building, floor, and room name.

    Expected format: "BUILDING-floor-room name (capacity) (attributes)"
    Example: "ARB02-1-Island (4) (Public) VC"

    Args:
        display_name: The room's display name from Graph API.

    Returns:
        Dictionary with 'building', 'floor', 'room_name' keys.
        If parsing fails, 'building' and 'floor' will be None,
        and 'room_name' will be the original display name.
    """
    if not display_name:
        return {"building": None, "floor": None, "room_name": display_name}

    # Split on first two dashes
    parts = display_name.split("-", 2)

    if len(parts) >= 3:
        building = parts[0].strip()
        try:
            floor = int(parts[1].strip())
        except ValueError:
            floor = None
        # Room name is everything after second dash, before any parentheses
        room_part = parts[2]
        # Strip off parenthetical info like "(4) (Public) VC"
        room_name = re.split(r'\s*\(', room_part)[0].strip()
        return {"building": building, "floor": floor, "room_name": room_name}
    elif len(parts) == 2:
        # Only building-something format
        building = parts[0].strip()
        room_name = parts[1].strip()
        room_name = re.split(r'\s*\(', room_name)[0].strip()
        return {"building": building, "floor": None, "room_name": room_name}
    else:
        # No dashes, just return as room name
        return {"building": None, "floor": None, "room_name": display_name}


def _room_to_dict(room) -> dict[str, Any]:
    """Convert a Graph API Room object to a dictionary.

    Args:
        room: A Room object from the Graph API places endpoint.

    Returns:
        Dictionary with room metadata including parsed building/floor.
    """
    display_name = room.display_name or ""
    parsed = _parse_room_name(display_name)

    return {
        "email": room.email_address,
        "name": display_name,
        "building": parsed["building"],
        "floor": parsed["floor"],
        "room_name": parsed["room_name"],
        "capacity": room.capacity,
        "phone": room.phone,
        "address": {
            "street": room.address.street if room.address else None,
            "city": room.address.city if room.address else None,
            "state": room.address.state if room.address else None,
            "postal_code": room.address.postal_code if room.address else None,
            "country_or_region": room.address.country_or_region if room.address else None,
        } if room.address else None,
        "geo_coordinates": {
            "latitude": room.geo_coordinates.latitude if room.geo_coordinates else None,
            "longitude": room.geo_coordinates.longitude if room.geo_coordinates else None,
        } if room.geo_coordinates else None,
        "audio_device_name": room.audio_device_name,
        "video_device_name": room.video_device_name,
        "display_device_name": room.display_device_name,
        "is_wheel_chair_accessible": room.is_wheel_chair_accessible,
        "tags": room.tags or [],
        "booking_type": room.booking_type.value if room.booking_type else None,
    }


def list_room_lists() -> list[dict[str, Any]]:
    """List all room lists (buildings/floors).

    Uses the beta findRoomLists API.

    Returns:
        List of room list dictionaries with 'name' and 'email' fields.
        Room lists are typically organized by building or building-floor
        (e.g., "ARB02" or "ARB02-1").

    Example:
        >>> lists = list_room_lists()
        >>> for rl in lists:
        ...     print(f"{rl['name']}: {rl['email']}")
        ARB02: ARB02@cisco.com
        ARB02-1: ARB02-1@cisco.com
    """
    raw_lists = _list_room_lists_beta()
    return [
        {"name": rl.get("name", ""), "email": rl.get("address", "")}
        for rl in raw_lists
    ]


def list_rooms(
    building: str | None = None,
    floor: int | None = None,
    client=None
) -> list[dict[str, Any]]:
    """List all rooms, optionally filtered by building and/or floor.

    Synchronous wrapper. See async_list_rooms() for full documentation.
    """
    return run_async(async_list_rooms(building=building, floor=floor, client=client))


async def async_list_rooms(
    building: str | None = None,
    floor: int | None = None,
    client=None
) -> list[dict[str, Any]]:
    """List all rooms, optionally filtered by building and/or floor.

    Tries the v1.0 places API first (requires Place.Read.All), then falls
    back to the beta findRooms API which works with standard permissions.

    Args:
        building: Filter by building code (e.g., "ARB02"). Case-insensitive.
        floor: Filter by floor number (e.g., 1, 2, 3).
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of room dictionaries with: email, name, building, floor,
        room_name, capacity, and other metadata.
    """
    # Try v1.0 places API first
    try:
        if client is None:
            client = get_graph_client()

        rooms_response = await client.places.graph_room.get()

        rooms = []
        for room in rooms_response.value:
            room_dict = _room_to_dict(room)
            rooms.append(room_dict)

        # Handle pagination
        next_link = rooms_response.odata_next_link
        while next_link:
            next_response = await client.places.graph_room.with_url(next_link).get()
            for room in next_response.value:
                room_dict = _room_to_dict(room)
                rooms.append(room_dict)
            next_link = next_response.odata_next_link

    except ODataError:
        # Fall back to beta API (uses requests, not async SDK)
        rooms = _list_rooms_beta_filtered(building=building, floor=floor)
        rooms.sort(key=lambda r: (
            r["building"] or "",
            r["floor"] if r["floor"] is not None else 999,
            r["room_name"] or ""
        ))
        return rooms

    # Filter by building (case-insensitive)
    if building is not None:
        building_upper = building.upper()
        rooms = [r for r in rooms if r["building"] and r["building"].upper() == building_upper]

    # Filter by floor
    if floor is not None:
        rooms = [r for r in rooms if r["floor"] == floor]

    # Sort by building, floor, room_name
    rooms.sort(key=lambda r: (
        r["building"] or "",
        r["floor"] if r["floor"] is not None else 999,
        r["room_name"] or ""
    ))

    return rooms


def _list_rooms_beta_filtered(
    building: str | None = None,
    floor: int | None = None,
) -> list[dict[str, Any]]:
    """List rooms using beta API with efficient filtering via room lists.

    If building and/or floor are specified, uses findRoomLists to find
    the appropriate room list and queries only that list.

    Args:
        building: Filter by building code (e.g., "ARB02"). Case-insensitive.
        floor: Filter by floor number (e.g., 1, 2, 3).

    Returns:
        List of room dictionaries.
    """
    # If both building and floor specified, try to find exact room list
    if building and floor is not None:
        room_list_email = f"{building.upper()}-{floor}@cisco.com"
        try:
            raw_rooms = _list_rooms_beta(room_list=room_list_email)
            return [_beta_room_to_dict(r) for r in raw_rooms]
        except requests.HTTPError:
            # Room list doesn't exist, fall through to broader search
            pass

    # If only building specified, try building room list first
    if building:
        # Try to get all room lists and find matching ones
        try:
            room_lists = _list_room_lists_beta()
            building_upper = building.upper()

            # Find room lists that match the building
            matching_lists = [
                rl for rl in room_lists
                if rl.get("name", "").upper().startswith(building_upper)
            ]

            if matching_lists:
                # If floor specified, try to find exact match
                if floor is not None:
                    exact_match = f"{building_upper}-{floor}"
                    for rl in matching_lists:
                        if rl.get("name", "").upper() == exact_match:
                            raw_rooms = _list_rooms_beta(room_list=rl["address"])
                            return [_beta_room_to_dict(r) for r in raw_rooms]

                # Otherwise, fetch from all matching room lists
                all_rooms = []
                for rl in matching_lists:
                    try:
                        raw_rooms = _list_rooms_beta(room_list=rl["address"])
                        all_rooms.extend([_beta_room_to_dict(r) for r in raw_rooms])
                    except requests.HTTPError:
                        continue

                # Filter by floor if specified
                if floor is not None:
                    all_rooms = [r for r in all_rooms if r["floor"] == floor]

                # Deduplicate by email
                seen = set()
                unique_rooms = []
                for r in all_rooms:
                    if r["email"] not in seen:
                        seen.add(r["email"])
                        unique_rooms.append(r)
                return unique_rooms

        except requests.HTTPError:
            pass

    # Fallback: fetch all rooms and filter client-side
    raw_rooms = _list_rooms_beta()
    rooms = [_beta_room_to_dict(r) for r in raw_rooms]

    # Filter by building (case-insensitive)
    if building is not None:
        building_upper = building.upper()
        rooms = [r for r in rooms if r["building"] and r["building"].upper() == building_upper]

    # Filter by floor
    if floor is not None:
        rooms = [r for r in rooms if r["floor"] == floor]

    return rooms


def check_room_availability(
    rooms: list[str],
    start: datetime | date | str,
    end: datetime | date | str,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Check availability for specified rooms. Synchronous wrapper.

    See async_check_room_availability() for full documentation.
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_check_room_availability(rooms, start, end, timezone=timezone, client=client))


async def async_check_room_availability(
    rooms: list[str],
    start: datetime | date | str,
    end: datetime | date | str,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Check availability for specified rooms in a time range.

    Uses the getSchedule API to check room availability efficiently
    in a single request.

    Args:
        rooms: List of room email addresses to check.
        start: Start of time range (date, datetime, or ISO string).
        end: End of time range (date, datetime, or ISO string).
        timezone: Timezone for interpreting start/end. Auto-detects if None.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of availability dictionaries, one per room.

    Raises:
        ODataError: If the Graph API request fails.
        ValueError: If no valid token is available or if rooms list is empty.
    """
    if not rooms:
        raise ValueError("rooms list cannot be empty")

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
        schedules=rooms,
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
            "name": None,  # getSchedule doesn't return display names
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


def find_available_rooms(
    start: datetime | date | str,
    end: datetime | date | str,
    building: str | None = None,
    floor: int | None = None,
    rooms: list[str] | None = None,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Find available rooms for a time range. Synchronous wrapper.

    See async_find_available_rooms() for full documentation.
    """
    return run_async(async_find_available_rooms(
        start, end, building=building, floor=floor, rooms=rooms,
        timezone=timezone, client=client
    ))


async def async_find_available_rooms(
    start: datetime | date | str,
    end: datetime | date | str,
    building: str | None = None,
    floor: int | None = None,
    rooms: list[str] | None = None,
    timezone: ZoneInfo | str | None = None,
    client=None
) -> list[dict[str, Any]]:
    """Find available rooms for a time range.

    Convenience function that lists rooms (filtered by building/floor) and
    checks their availability, returning rooms sorted by availability.

    Args:
        start: Start of time range (date, datetime, or ISO string).
        end: End of time range (date, datetime, or ISO string).
        building: Filter by building code (e.g., "ARB02"). Ignored if rooms is provided.
        floor: Filter by floor number. Ignored if rooms is provided.
        rooms: Specific room emails to check. Overrides building/floor filters.
        timezone: Timezone for interpreting start/end. Auto-detects if None.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of room dictionaries sorted by availability (available rooms first).
    """
    if client is None:
        client = get_graph_client()

    # Get list of rooms (may use beta API which doesn't need SDK client)
    if rooms is not None:
        room_list = [{"email": email, "name": None} for email in rooms]
    else:
        room_list = await async_list_rooms(building=building, floor=floor, client=client)

    if not room_list:
        return []

    # Check availability for all rooms
    room_emails = [r["email"] for r in room_list if r["email"]]
    if not room_emails:
        return room_list

    availability = await async_check_room_availability(
        rooms=room_emails,
        start=start,
        end=end,
        timezone=timezone,
        client=client,
    )

    # Create email -> availability lookup
    availability_by_email = {a["email"]: a for a in availability}

    # Merge availability info into room data
    results = []
    for room in room_list:
        email = room.get("email")
        avail = availability_by_email.get(email, {})
        merged = {
            **room,
            "is_available": avail.get("is_available", None),
            "availability_view": avail.get("availability_view", ""),
            "busy_times": avail.get("busy_times", []),
        }
        results.append(merged)

    # Sort: available rooms first, then by building/floor/name
    results.sort(key=lambda r: (
        0 if r.get("is_available") else 1,
        r.get("building") or "",
        r.get("floor") if r.get("floor") is not None else 999,
        r.get("room_name") or r.get("name") or ""
    ))

    return results


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
