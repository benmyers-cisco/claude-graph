#!/usr/bin/env python3
"""CLI entry point for Microsoft Graph API operations.

This module provides command-line commands for calendar, room, meeting, and auth operations.
These commands are pre-approved in .claude/settings.json and can be invoked
by Claude skills without additional approval.

Usage:
    python -m src.cli auth [--email EMAIL]
    python -m src.cli calendar [days] [--date YYYY-MM-DD]
    python -m src.cli book-rooms [--building BUILDING] [--floor FLOOR]
    python -m src.cli book-room <event-id> <room-email> [--room-name NAME]
    python -m src.cli find-meeting <names...> [--duration MINUTES]
    python -m src.cli room-availability <start> <end> [--building BUILDING] [--floor FLOOR]
    python -m src.cli resolve-person <query>
    python -m src.cli email search <query> [--max N] [--folder FOLDER]
    python -m src.cli email read <message-id>
    python -m src.cli stats [days] [--date YYYY-MM-DD] [--work-start HH:MM] [--work-end HH:MM]
    python -m src.cli config [--email EMAIL] [--timezone TZ] [--building CODE] [--floors 1,2] [--favorite-rooms EMAIL,EMAIL]
"""

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from .calendar_tools import (
    get_events_in_range, get_local_timezone, windows_tz_to_iana, get_timezone_abbr,
    create_event,
)
from .config import get_config, set_config
from .email_tools import search_emails, read_email
from .stats import compute_day_stats, compute_range_stats
from .room_tools import list_rooms, find_available_rooms, check_room_availability
from .people_tools import resolve_person, get_attendee_locations, get_org_info, get_user_timezones
from .meeting_tools import find_meeting_times
from .graph_client import get_graph_client, run_async, get_token, set_token, is_token_expired, get_concurrency_semaphore


def get_current_user_email() -> str | None:
    """Get the current user's email address."""
    try:
        client = get_graph_client()
        me = run_async(client.me.get())
        return (me.mail or me.user_principal_name or "").lower()
    except Exception:
        return None


def format_duration(minutes: int) -> str:
    """Format duration in minutes to human-readable string."""
    if minutes < 60:
        return f"{minutes} min"
    hours = minutes // 60
    mins = minutes % 60
    if mins == 0:
        return f"{hours} hr" if hours == 1 else f"{hours} hrs"
    return f"{hours} hr {mins} min"


def format_attendees(attendees: list, current_user_email: str | None, max_show: int = 10) -> str:
    """Format attendee list, excluding current user."""
    if not attendees:
        return "-"

    # Filter out current user and resource attendees (rooms)
    filtered = [
        a for a in attendees
        if a.get("type") != "resource" and (
            not current_user_email or
            (a.get("address") or "").lower() != current_user_email
        )
    ]

    if not filtered:
        return "-"

    names = []
    for a in filtered[:max_show]:
        name = a.get("name") or a.get("address", "").split("@")[0]
        names.append(name)

    if len(filtered) > max_show:
        names.append(f"and {len(filtered) - max_show} others")

    return ", ".join(names)


def format_recurrence(pattern: dict | None) -> str | None:
    """Format a recurrence pattern dict into a human-readable frequency string.

    Args:
        pattern: Recurrence pattern dict with 'type', 'interval', 'days_of_week', etc.

    Returns:
        Human-readable string like "Weekly", "Biweekly", "Monthly", or None if not recurring.
    """
    if not pattern:
        return None

    pat_type = pattern.get("type", "")
    interval = pattern.get("interval", 1)

    if pat_type == "daily":
        if interval == 1:
            return "Daily"
        return f"Every {interval} days"
    elif pat_type == "weekly":
        if interval == 1:
            return "Weekly"
        elif interval == 2:
            return "Biweekly"
        return f"Every {interval} weeks"
    elif pat_type in ("absoluteMonthly", "relativeMonthly"):
        if interval == 1:
            return "Monthly"
        return f"Every {interval} months"
    elif pat_type in ("absoluteYearly", "relativeYearly"):
        if interval == 1:
            return "Yearly"
        return f"Every {interval} years"

    return None


def get_room_from_attendees(attendees: list) -> str | None:
    """Extract room name from attendees list."""
    for a in attendees:
        if a.get("type") == "resource":
            name = a.get("name") or a.get("address", "")
            # Try to extract short room name (e.g., "ARB02-1-Elm" from full name)
            if name:
                # If it looks like a room name with building-floor-room format
                match = re.match(r"^([A-Z0-9]+-\d+-[^(]+)", name)
                if match:
                    return match.group(1).strip()
                return name
    return None


def get_primary_contact(attendees: list, current_user_email: str | None) -> dict | None:
    """Determine the primary contact for a 1:1 meeting.

    Filters out the current user, resource attendees (rooms), shared mailboxes,
    and mailer lists. If exactly one person remains, that's the primary contact.

    Returns:
        The attendee dict if exactly one primary contact, else None.
    """
    if not attendees:
        return None

    filtered = []
    for a in attendees:
        # Skip resource attendees (rooms)
        if a.get("type") == "resource":
            continue
        # Skip current user
        addr = (a.get("address") or "").lower()
        if current_user_email and addr == current_user_email:
            continue
        # Skip shared mailboxes
        name = a.get("name") or ""
        if "Shared Mailbox" in name:
            continue
        # Skip mailer lists
        if "mailer list" in name.lower():
            continue
        filtered.append(a)

    if len(filtered) == 1:
        return filtered[0]
    return None


def cmd_calendar(args):
    """Show calendar events for specified days."""
    prefs = get_config()
    tz = prefs.get("timezone")

    # Resolve the effective timezone (config → system detection)
    if tz:
        local_tz = ZoneInfo(tz)
    else:
        local_tz = get_local_timezone()

    # Calculate date range
    if args.date:
        try:
            start_date = date.fromisoformat(args.date)
        except ValueError:
            print(json.dumps({"error": f"Invalid date format: {args.date}. Use YYYY-MM-DD."}))
            return 1
    else:
        start_date = date.today()
    days = args.days if args.days else 1
    end_date = start_date + timedelta(days=days)

    # Get current user email for filtering attendees
    current_user_email = get_current_user_email()

    # Fetch events
    try:
        events = get_events_in_range(start_date, end_date, timezone=tz)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    # Build events list and identify primary contacts for 1:1 meetings
    events_data = []
    primary_contacts = {}  # event index -> attendee dict
    series_master_map = {}  # event index -> series_master_id

    for idx, event in enumerate(events):
        if event.get("is_cancelled"):
            continue

        room = get_room_from_attendees(event.get("attendees", []))
        response = event.get("response_status", {})
        response_type = response.get("response") if response else None

        event_data = {
            "id": event.get("id"),
            "subject": event.get("subject"),
            "start": event.get("start"),
            "end": event.get("end"),
            "duration_minutes": event.get("duration_minutes"),
            "duration_formatted": format_duration(event.get("duration_minutes", 0)),
            "room": room,
            "response_status": response_type,
            "is_all_day": event.get("is_all_day"),
            "attendees_formatted": format_attendees(
                event.get("attendees", []),
                current_user_email
            ),
            "organizer": event.get("organizer"),
            "their_time": None,
            "frequency": None,
            "category": ", ".join(event.get("categories", [])) or None,
            "body_preview": event.get("body_preview"),
            "body": event.get("body"),
        }
        ev_idx = len(events_data)
        events_data.append(event_data)

        # Track series master for recurring events
        master_id = event.get("series_master_id")
        if master_id:
            series_master_map[ev_idx] = master_id

        # Check for primary contact (1:1 meeting)
        contact = get_primary_contact(event.get("attendees", []), current_user_email)
        if contact:
            primary_contacts[ev_idx] = contact

    # Batch-fetch timezones for all primary contacts
    unique_emails = list({
        (primary_contacts[i].get("address") or "").lower()
        for i in primary_contacts
        if primary_contacts[i].get("address")
    })

    tz_map: dict[str, str | None] = {}
    if unique_emails:
        try:
            tz_map = get_user_timezones(unique_emails)
        except Exception:
            pass  # Best-effort — degrade gracefully

    # Compute "their_time" for each event with a primary contact
    for idx, contact in primary_contacts.items():
        email = (contact.get("address") or "").lower()
        windows_tz = tz_map.get(email)
        if not windows_tz:
            continue

        iana_name = windows_tz_to_iana(windows_tz)
        if not iana_name:
            continue

        try:
            their_tz = ZoneInfo(iana_name)
        except Exception:
            continue

        # Parse the event start time and convert to their timezone
        start_info = events_data[idx].get("start")
        if not start_info or not start_info.get("date_time"):
            continue

        start_dt = datetime.fromisoformat(start_info["date_time"])
        start_dt = start_dt.replace(tzinfo=local_tz)
        their_dt = start_dt.astimezone(their_tz)

        events_data[idx]["their_time"] = {
            "formatted": their_dt.strftime("%I:%M %p"),
            "timezone_abbr": get_timezone_abbr(their_tz, their_dt),
        }

    # Batch-fetch recurrence patterns from series masters
    unique_master_ids = list(set(series_master_map.values()))
    recurrence_cache: dict[str, str | None] = {}  # master_id -> formatted frequency
    if unique_master_ids:
        try:
            import asyncio
            from msgraph.generated.users.item.events.item.event_item_request_builder import (
                EventItemRequestBuilder,
            )

            async def _fetch_recurrences():
                client = get_graph_client()
                sem = get_concurrency_semaphore()
                config = EventItemRequestBuilder.EventItemRequestBuilderGetRequestConfiguration(
                    query_parameters=EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters(
                        select=["recurrence"],
                    )
                )

                async def _get(mid):
                    async with sem:
                        return await client.me.events.by_event_id(mid).get(request_configuration=config)

                tasks = [_get(mid) for mid in unique_master_ids]
                return await asyncio.gather(*tasks, return_exceptions=True)

            results = run_async(_fetch_recurrences())
            for mid, result in zip(unique_master_ids, results):
                if isinstance(result, Exception) or not result:
                    recurrence_cache[mid] = None
                    continue
                pat = None
                if result.recurrence and result.recurrence.pattern:
                    p = result.recurrence.pattern
                    pat = {
                        "type": p.type.value if p.type else None,
                        "interval": p.interval,
                        "days_of_week": [d.value for d in p.days_of_week] if p.days_of_week else [],
                    }
                recurrence_cache[mid] = format_recurrence(pat)
        except Exception:
            pass  # Best-effort

    # Apply frequency to events
    for ev_idx, master_id in series_master_map.items():
        freq = recurrence_cache.get(master_id)
        if freq:
            events_data[ev_idx]["frequency"] = freq

    # Compute user's timezone abbreviation from first event date (or today)
    ref_dt = datetime.now(local_tz)
    if events_data and events_data[0].get("start") and events_data[0]["start"].get("date_time"):
        try:
            ref_dt = datetime.fromisoformat(events_data[0]["start"]["date_time"]).replace(tzinfo=local_tz)
        except Exception:
            pass
    user_tz_abbr = get_timezone_abbr(local_tz, ref_dt)

    # Format output
    result = {
        "date_range": {
            "start": start_date.isoformat(),
            "end": end_date.isoformat(),
            "days": days,
        },
        "timezone": str(local_tz),
        "timezone_abbr": user_tz_abbr,
        "events": events_data,
    }

    print(json.dumps(result, indent=2))
    return 0


def cmd_book_rooms(args):
    """Find meetings needing rooms and show room availability."""
    prefs = get_config()
    building = args.building or prefs.get("default_building")
    floors = [args.floor] if args.floor else prefs.get("preferred_floors", [])
    tz = prefs.get("timezone")
    skip_location = getattr(args, "skip_location", False)

    # Get today's events
    today = date.today()
    tomorrow = today + timedelta(days=1)

    try:
        events = get_events_in_range(today, tomorrow, timezone=tz)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    # Get current user email
    current_user_email = get_current_user_email()

    # Find meetings without rooms, stash raw attendees per meeting
    meetings_needing_rooms = []
    for event in events:
        if event.get("is_cancelled") or event.get("is_all_day"):
            continue

        # Check if meeting has a room
        room = get_room_from_attendees(event.get("attendees", []))
        if room:
            continue

        # Determine if user is organizer
        organizer = event.get("organizer", {})
        organizer_email = (organizer.get("address") or "").lower()
        is_organizer = current_user_email and organizer_email == current_user_email

        # Non-resource attendees
        non_resource = [
            a for a in event.get("attendees", [])
            if a.get("type") != "resource"
        ]

        meetings_needing_rooms.append({
            "id": event.get("id"),
            "subject": event.get("subject"),
            "start": event.get("start"),
            "end": event.get("end"),
            "duration_minutes": event.get("duration_minutes"),
            "is_organizer": is_organizer,
            "organizer": organizer,
            "attendees_count": len(non_resource),
            "_attendee_emails": [
                (a.get("address") or "").lower()
                for a in non_resource
                if a.get("address")
            ],
        })

    # Location lookup: collect all unique attendee emails across meetings
    location_results = {}
    location_aware = False
    if not skip_location and building and meetings_needing_rooms:
        all_emails = set()
        for m in meetings_needing_rooms:
            all_emails.update(m["_attendee_emails"])
        all_emails.discard("")

        if all_emails:
            try:
                loc_data = get_attendee_locations(list(all_emails), building)
                # Index by email for fast per-meeting lookup
                for att in loc_data.get("attendees", []):
                    location_results[att["email"].lower()] = att
                location_aware = True
            except Exception:
                pass  # Best-effort — degrade gracefully

    # For each meeting, find available rooms and attach location info
    result = {
        "building": building,
        "floors": floors,
        "meetings_needing_rooms": [],
    }

    for meeting in meetings_needing_rooms:
        start_str = meeting["start"]["date_time"]
        end_str = meeting["end"]["date_time"]

        # Parse start/end times
        start_dt = datetime.fromisoformat(start_str)
        end_dt = datetime.fromisoformat(end_str)

        # Find available rooms
        try:
            if floors:
                # Check each preferred floor
                available = []
                for floor in floors:
                    rooms = find_available_rooms(
                        start=start_dt,
                        end=end_dt,
                        building=building,
                        floor=floor,
                        timezone=tz
                    )
                    available.extend([r for r in rooms if r.get("is_available")])
            else:
                available = find_available_rooms(
                    start=start_dt,
                    end=end_dt,
                    building=building,
                    timezone=tz
                )
                available = [r for r in available if r.get("is_available")]
        except Exception:
            available = []

        # Classify this meeting's attendees using cached location results
        local_count = 0
        other_office_count = 0
        unknown_count = 0

        if location_aware:
            for email in meeting["_attendee_emails"]:
                att = location_results.get(email)
                if att:
                    cls = att["classification"]
                    if cls == "local":
                        local_count += 1
                    elif cls == "other_office":
                        other_office_count += 1
                    else:
                        unknown_count += 1
                else:
                    unknown_count += 1
            # Current user is presumably local if not already in the list
            if current_user_email and current_user_email not in meeting["_attendee_emails"]:
                local_count += 1

        # Remove internal field
        attendee_emails = meeting.pop("_attendee_emails")

        # Sort favorites first, preserving order within each group
        favorite_rooms = prefs.get("favorite_rooms", [])
        favorite_set = {e.lower() for e in favorite_rooms}
        favorites = [r for r in available if (r.get("email") or "").lower() in favorite_set]
        non_favorites = [r for r in available if (r.get("email") or "").lower() not in favorite_set]
        sorted_rooms = favorites + non_favorites

        meeting_data = {
            **meeting,
            "location_aware": location_aware,
            "local_attendees_count": local_count if location_aware else None,
            "other_office_count": other_office_count if location_aware else None,
            "unknown_location_count": unknown_count if location_aware else None,
            "in_person_estimate": local_count if location_aware else None,
            "available_rooms": [
                {
                    "email": r.get("email"),
                    "name": r.get("name"),
                    "capacity": r.get("capacity"),
                    "building": r.get("building"),
                    "floor": r.get("floor"),
                    "room_name": r.get("room_name"),
                    "is_favorite": (r.get("email") or "").lower() in favorite_set,
                }
                for r in sorted_rooms[:10]  # Limit to 10 rooms
            ],
        }
        result["meetings_needing_rooms"].append(meeting_data)

    print(json.dumps(result, indent=2))
    return 0


def cmd_book_room(args):
    """Book a conference room for an existing calendar event."""
    import asyncio

    async def _book(event_id, room_email, room_name):
        client = get_graph_client()

        # Get event and current user
        event = await client.me.events.by_event_id(event_id).get()
        me = await client.me.get()
        my_email = (me.mail or me.user_principal_name or "").lower()
        organizer_email = (event.organizer.email_address.address or "").lower()
        is_organizer = organizer_email == my_email

        if is_organizer:
            # Add room as Resource attendee
            from msgraph.generated.models.attendee import Attendee
            from msgraph.generated.models.attendee_type import AttendeeType
            from msgraph.generated.models.email_address import EmailAddress
            from msgraph.generated.models.event import Event as EventModel

            attendees = list(event.attendees) if event.attendees else []
            attendees.append(
                Attendee(
                    email_address=EmailAddress(address=room_email, name=room_name),
                    type=AttendeeType.Resource,
                )
            )
            update = EventModel(attendees=attendees)
            await client.me.events.by_event_id(event_id).patch(update)
            return {
                "status": "success",
                "method": "attendee",
                "message": f"Room {room_name or room_email} added as attendee to event.",
            }
        else:
            # Forward invite to room
            from msgraph.generated.models.email_address import EmailAddress
            from msgraph.generated.models.recipient import Recipient
            from msgraph.generated.users.item.events.item.forward.forward_post_request_body import (
                ForwardPostRequestBody,
            )

            body = ForwardPostRequestBody(
                to_recipients=[
                    Recipient(email_address=EmailAddress(address=room_email))
                ],
                comment="",
            )
            await client.me.events.by_event_id(event_id).forward.post(body)
            return {
                "status": "success",
                "method": "forward",
                "message": f"Meeting invite forwarded to {room_name or room_email}.",
            }

    try:
        result = asyncio.run(_book(args.event_id, args.room_email, args.room_name or ""))
    except Exception as e:
        result = {"status": "error", "message": str(e)}

    print(json.dumps(result, indent=2))
    return 0 if result.get("status") == "success" else 1


def cmd_find_meeting(args):
    """Find meeting times for specified people."""
    prefs = get_config()
    tz = prefs.get("timezone")
    duration = args.duration or 30
    days = args.days or 7

    # Resolve names to emails
    resolved = []
    errors = []

    for name in args.names:
        try:
            result = resolve_person(name)
            if result.get("resolved_email"):
                resolved.append({
                    "input": name,
                    "email": result["resolved_email"],
                    "name": result["best_match"]["name"] if result.get("best_match") else None,
                })
            elif result.get("is_ambiguous"):
                errors.append({
                    "input": name,
                    "error": "ambiguous",
                    "matches": [
                        {"email": m["email"], "name": m["name"]}
                        for m in result.get("matches", [])[:5]
                    ],
                })
            else:
                errors.append({
                    "input": name,
                    "error": "not_found",
                })
        except Exception as e:
            errors.append({
                "input": name,
                "error": str(e),
            })

    if errors:
        print(json.dumps({
            "resolved": resolved,
            "errors": errors,
            "message": "Could not resolve all participants",
        }, indent=2))
        return 1 if not resolved else 0

    # Get current user email and add to participants
    current_user_email = get_current_user_email()
    participant_emails = [r["email"] for r in resolved]
    if current_user_email and current_user_email not in participant_emails:
        participant_emails.insert(0, current_user_email)

    # Find meeting times
    try:
        times = find_meeting_times(
            participants=participant_emails,
            duration_minutes=duration,
            days_ahead=days,
            timezone=tz
        )
    except Exception as e:
        print(json.dumps({
            "resolved": resolved,
            "error": str(e),
        }, indent=2))
        return 1

    result = {
        "participants": resolved,
        "duration_minutes": duration,
        "duo_meeting_duration_minutes": duration - 5,
        "scheduling_note": "Duo standard: meetings start at :05/:35 with 5-min buffer",
        "available_times": [
            {
                "start": t.get("start"),
                "end": t.get("end"),
                "confidence": t.get("confidence"),
                "all_available": t.get("all_available"),
                "lunch_overlap": t.get("lunch_overlap", False),
            }
            for t in times[:10]  # Limit to 10 suggestions
        ],
    }

    print(json.dumps(result, indent=2))
    return 0


def cmd_room_availability(args):
    """Check room availability for a time range."""
    prefs = get_config()
    building = args.building or prefs.get("default_building")
    floor = args.floor
    tz = prefs.get("timezone")

    # Resolve timezone for accurate "now" handling
    local_tz = ZoneInfo(tz) if tz else get_local_timezone()

    # Parse start/end times
    try:
        if args.start.lower() == "now":
            start_dt = datetime.now(local_tz)
        else:
            start_dt = datetime.fromisoformat(args.start)

        if args.end:
            end_dt = datetime.fromisoformat(args.end)
        else:
            # Default to 1 hour from start
            end_dt = start_dt + timedelta(hours=1)
    except ValueError as e:
        print(json.dumps({"error": f"Invalid date format: {e}"}))
        return 1

    # Find available rooms
    try:
        rooms = find_available_rooms(
            start=start_dt,
            end=end_dt,
            building=building,
            floor=floor,
            timezone=tz
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    # Separate available and busy rooms
    available = [r for r in rooms if r.get("is_available")]
    busy = [r for r in rooms if not r.get("is_available")]

    result = {
        "time_range": {
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
        },
        "building": building,
        "floor": floor,
        "available_rooms": [
            {
                "email": r.get("email"),
                "name": r.get("name"),
                "capacity": r.get("capacity"),
                "building": r.get("building"),
                "floor": r.get("floor"),
                "room_name": r.get("room_name"),
            }
            for r in available
        ],
        "busy_rooms": [
            {
                "email": r.get("email"),
                "name": r.get("name"),
                "capacity": r.get("capacity"),
                "building": r.get("building"),
                "floor": r.get("floor"),
                "room_name": r.get("room_name"),
                "busy_times": r.get("busy_times", []),
            }
            for r in busy
        ],
    }

    print(json.dumps(result, indent=2))
    return 0


def cmd_resolve_person(args):
    """Resolve a name or partial email to a full email address."""
    try:
        result = resolve_person(args.query)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_org(args):
    """Look up a person's manager and/or direct reports."""
    include_manager = not getattr(args, "reports_only", False)
    include_reports = not getattr(args, "manager_only", False)

    try:
        result = get_org_info(
            args.query,
            include_manager=include_manager,
            include_reports=include_reports,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_stats(args):
    """Compute meeting time statistics."""
    prefs = get_config()
    tz = prefs.get("timezone")

    start_date = date.fromisoformat(args.date) if args.date else date.today()
    days = args.days if args.days else 1
    work_start = args.work_start or "09:00"
    work_end = args.work_end or "17:00"

    # Get current user email for solo-event filtering
    current_user_email = get_current_user_email()

    # Collect business days in range
    target_dates = []
    d = start_date
    while len(target_dates) < days:
        if d.weekday() < 5:  # Mon-Fri
            target_dates.append(d)
        d += timedelta(days=1)

    # Fetch events and compute stats per day
    daily_stats = []
    for target in target_dates:
        next_day = target + timedelta(days=1)
        try:
            events = get_events_in_range(target, next_day, timezone=tz)
        except Exception as e:
            print(json.dumps({"error": str(e)}))
            return 1
        day_stats = compute_day_stats(
            events,
            event_date=target,
            work_start=work_start,
            work_end=work_end,
            user_email=current_user_email,
        )
        daily_stats.append(day_stats)

    summary = compute_range_stats(daily_stats)

    result = {
        "date_range": {
            "start": target_dates[0].isoformat(),
            "end": target_dates[-1].isoformat(),
            "business_days": len(target_dates),
        },
        "work_hours": {
            "start": work_start,
            "end": work_end,
        },
        "days": daily_stats,
        "summary": summary,
    }

    print(json.dumps(result, indent=2))
    return 0


def cmd_email_search(args):
    """Search emails matching a query."""
    try:
        results = search_emails(
            args.query,
            max_results=args.max,
            folder=args.folder,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps({"messages": results, "count": len(results)}, indent=2))
    return 0


def cmd_email_read(args):
    """Read the full content of an email message."""
    try:
        result = read_email(args.message_id)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_delete_event(args):
    """Delete a calendar event by ID."""
    import asyncio

    async def _delete(event_id):
        client = get_graph_client()
        await client.me.events.by_event_id(event_id).delete()
        return {"status": "success", "message": f"Event {event_id} deleted."}

    try:
        result = asyncio.run(_delete(args.event_id))
    except Exception as e:
        result = {"status": "error", "message": str(e)}

    print(json.dumps(result, indent=2))
    return 0 if result.get("status") == "success" else 1


def cmd_create_event(args):
    """Create a new calendar event."""
    try:
        result = create_event(
            subject=args.subject,
            start=args.start,
            end=args.end,
            body=args.body,
            show_as=args.show_as or "busy",
            is_all_day=args.all_day,
            location=args.location,
            attendees=args.attendees.split(",") if args.attendees else None,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_config(args):
    """Get or set user configuration."""
    updates = {}
    if args.email is not None:
        updates["email"] = args.email
    if args.timezone is not None:
        updates["timezone"] = args.timezone
    if args.building is not None:
        updates["default_building"] = args.building
    if args.floors is not None:
        floors = [int(f.strip()) for f in args.floors.split(",") if f.strip().isdigit()]
        updates["preferred_floors"] = floors
    if args.favorite_rooms is not None:
        if args.favorite_rooms.strip() == "":
            updates["favorite_rooms"] = []
        else:
            updates["favorite_rooms"] = [
                r.strip() for r in args.favorite_rooms.split(",") if r.strip()
            ]

    if updates:
        config = set_config(updates)
    else:
        config = get_config()

    print(json.dumps(config, indent=2))
    return 0


def cmd_auth(args):
    """Authenticate with Graph Explorer and extract access token."""
    prefs = get_config()
    email = args.email or prefs.get("email")
    timeout = args.timeout or 120

    from .auth import authenticate

    result = authenticate(email=email, timeout=timeout)

    if result.get("status") == "success" and result.get("token"):
        set_token(result["token"])
        # Verify token works
        try:
            client = get_graph_client()
            me = run_async(client.me.get())
            result["verified_as"] = {
                "name": me.display_name,
                "email": me.mail or me.user_principal_name,
            }
        except Exception:
            pass

    print(json.dumps(result, indent=2))
    return 0 if result.get("status") == "success" else 1


def _is_token_error(output: str) -> bool:
    """Check if CLI output indicates a token/auth error."""
    lower = output.lower()
    return any(phrase in lower for phrase in [
        "invalidauthenticationtoken",
        "token is expired",
        "lifetime validation failed",
        "compacttoken",
    ])


def _auto_auth() -> bool:
    """Run automatic re-authentication. Returns True on success."""
    from .auth import authenticate

    prefs = get_config()
    email = prefs.get("email")

    print(json.dumps({
        "status": "re-authenticating",
        "message": "Token expired. Opening browser to re-authenticate...",
    }), file=sys.stderr)

    result = authenticate(email=email, timeout=120)

    if result.get("status") == "success" and result.get("token"):
        set_token(result["token"])
        print(json.dumps({
            "status": "re-authenticated",
            "message": "Token refreshed successfully.",
        }), file=sys.stderr)
        return True

    print(json.dumps({
        "status": "auth-failed",
        "message": "Re-authentication failed. Run /msgraph:auth manually.",
    }), file=sys.stderr)
    return False


def main():
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="Microsoft Graph API CLI for calendar and room operations"
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # auth command
    auth_parser = subparsers.add_parser("auth", help="Authenticate with Graph Explorer")
    auth_parser.add_argument(
        "--email", "-e",
        help="Microsoft account email address"
    )
    auth_parser.add_argument(
        "--timeout", "-t",
        type=int,
        default=120,
        help="Timeout in seconds for auth completion (default: 120)"
    )
    auth_parser.set_defaults(func=cmd_auth)

    # calendar command
    cal_parser = subparsers.add_parser("calendar", help="Show calendar events")
    cal_parser.add_argument(
        "days",
        nargs="?",
        type=int,
        default=1,
        help="Number of days to show (default: 1)"
    )
    cal_parser.add_argument(
        "--date", "-d",
        help="Start date in YYYY-MM-DD format (default: today)"
    )
    cal_parser.set_defaults(func=cmd_calendar)

    # book-rooms command
    book_parser = subparsers.add_parser("book-rooms", help="Find meetings needing rooms")
    book_parser.add_argument(
        "--building", "-b",
        help="Building to search for rooms"
    )
    book_parser.add_argument(
        "--floor", "-f",
        type=int,
        help="Floor to search for rooms"
    )
    book_parser.add_argument(
        "--skip-location",
        action="store_true",
        default=False,
        help="Skip attendee location lookup for faster results"
    )
    book_parser.set_defaults(func=cmd_book_rooms)

    # book-room command (actually book a specific room)
    book_one_parser = subparsers.add_parser("book-room", help="Book a room for an event")
    book_one_parser.add_argument(
        "event_id",
        help="Calendar event ID"
    )
    book_one_parser.add_argument(
        "room_email",
        help="Room email address"
    )
    book_one_parser.add_argument(
        "--room-name",
        help="Room display name (optional)"
    )
    book_one_parser.set_defaults(func=cmd_book_room)

    # find-meeting command
    find_parser = subparsers.add_parser("find-meeting", help="Find meeting times for people")
    find_parser.add_argument(
        "names",
        nargs="+",
        help="Names or emails of participants"
    )
    find_parser.add_argument(
        "--duration", "-d",
        type=int,
        default=30,
        help="Meeting duration in minutes (default: 30)"
    )
    find_parser.add_argument(
        "--days",
        type=int,
        default=7,
        help="Days ahead to search (default: 7)"
    )
    find_parser.set_defaults(func=cmd_find_meeting)

    # room-availability command
    room_parser = subparsers.add_parser("room-availability", help="Check room availability")
    room_parser.add_argument(
        "start",
        help="Start time (ISO format or 'now')"
    )
    room_parser.add_argument(
        "end",
        nargs="?",
        help="End time (ISO format, default: 1 hour after start)"
    )
    room_parser.add_argument(
        "--building", "-b",
        help="Building to check"
    )
    room_parser.add_argument(
        "--floor", "-f",
        type=int,
        help="Floor to check"
    )
    room_parser.set_defaults(func=cmd_room_availability)

    # resolve-person command
    resolve_parser = subparsers.add_parser("resolve-person", help="Resolve name to email")
    resolve_parser.add_argument(
        "query",
        help="Name or partial email to resolve"
    )
    resolve_parser.set_defaults(func=cmd_resolve_person)

    # org command
    org_parser = subparsers.add_parser("org", help="Look up manager and direct reports")
    org_parser.add_argument(
        "query",
        help="Name or email of the person to look up"
    )
    org_parser.add_argument(
        "--manager-only",
        action="store_true",
        default=False,
        help="Only show manager (skip direct reports)"
    )
    org_parser.add_argument(
        "--reports-only",
        action="store_true",
        default=False,
        help="Only show direct reports (skip manager)"
    )
    org_parser.set_defaults(func=cmd_org)

    # stats command
    stats_parser = subparsers.add_parser("stats", help="Compute meeting time statistics")
    stats_parser.add_argument(
        "days",
        nargs="?",
        type=int,
        default=1,
        help="Number of business days (default: 1)"
    )
    stats_parser.add_argument(
        "--date", "-d",
        help="Start date in YYYY-MM-DD format (default: today)"
    )
    stats_parser.add_argument(
        "--work-start",
        default="09:00",
        help="Work hours start in HH:MM (default: 09:00)"
    )
    stats_parser.add_argument(
        "--work-end",
        default="17:00",
        help="Work hours end in HH:MM (default: 17:00)"
    )
    stats_parser.set_defaults(func=cmd_stats)

    # email search command
    email_search_parser = subparsers.add_parser("email", help="Email operations")
    email_subparsers = email_search_parser.add_subparsers(dest="email_command", help="Email sub-commands")

    search_parser = email_subparsers.add_parser("search", help="Search emails")
    search_parser.add_argument(
        "query",
        help="KQL search query (e.g., 'from:person subject:weekly', 'webex recording')"
    )
    search_parser.add_argument(
        "--max", "-m",
        type=int,
        default=10,
        help="Maximum results to return (default: 10)"
    )
    search_parser.add_argument(
        "--folder", "-f",
        help="Mail folder (inbox, sentitems, drafts, or folder ID)"
    )
    search_parser.set_defaults(func=cmd_email_search)

    read_parser = email_subparsers.add_parser("read", help="Read full email content")
    read_parser.add_argument(
        "message_id",
        help="Message ID from a search result"
    )
    read_parser.set_defaults(func=cmd_email_read)

    # delete-event command
    delete_event_parser = subparsers.add_parser("delete-event", help="Delete a calendar event")
    delete_event_parser.add_argument(
        "event_id",
        help="Calendar event ID to delete"
    )
    delete_event_parser.set_defaults(func=cmd_delete_event)

    # create-event command
    create_event_parser = subparsers.add_parser("create-event", help="Create a calendar event")
    create_event_parser.add_argument(
        "subject",
        help="Event subject/title"
    )
    create_event_parser.add_argument(
        "start",
        help="Start time (ISO format, e.g., 2026-08-11T09:00:00)"
    )
    create_event_parser.add_argument(
        "end",
        help="End time (ISO format, e.g., 2026-08-11T09:30:00)"
    )
    create_event_parser.add_argument(
        "--body", "-b",
        help="Event body/description (plain text)"
    )
    create_event_parser.add_argument(
        "--show-as",
        choices=["busy", "free", "tentative", "oof", "workingElsewhere"],
        default="busy",
        help="Show as status (default: busy)"
    )
    create_event_parser.add_argument(
        "--all-day",
        action="store_true",
        default=False,
        help="Create as all-day event"
    )
    create_event_parser.add_argument(
        "--location", "-l",
        help="Event location"
    )
    create_event_parser.add_argument(
        "--attendees", "-a",
        help="Comma-separated email addresses to invite"
    )
    create_event_parser.set_defaults(func=cmd_create_event)

    # config command
    config_parser = subparsers.add_parser("config", help="Get or set configuration")
    config_parser.add_argument(
        "--email",
        help="Microsoft account email"
    )
    config_parser.add_argument(
        "--timezone",
        help="IANA timezone (e.g., America/Detroit)"
    )
    config_parser.add_argument(
        "--building",
        help="Default building code"
    )
    config_parser.add_argument(
        "--floors",
        help="Preferred floors, comma-separated (e.g., 1,2)"
    )
    config_parser.add_argument(
        "--favorite-rooms",
        help="Favorite room emails, comma-separated (empty string clears)"
    )
    config_parser.set_defaults(func=cmd_config)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    # Handle nested email subcommand
    if args.command == "email" and not getattr(args, "email_command", None):
        email_search_parser.print_help()
        return 1

    # Check for token (skip for auth and config commands — they don't need an existing token)
    if args.command not in ("auth", "config"):
        token = get_token()
        if not token:
            # No token at all — try auto-auth
            if not _auto_auth():
                print(json.dumps({
                    "error": "No access token available. Run /msgraph:auth to authenticate.",
                    "help": "Or get a token from https://developer.microsoft.com/en-us/graph/graph-explorer"
                }))
                return 1
        elif is_token_expired(token):
            # Token exists but is expired — auto-auth before running the command
            if not _auto_auth():
                print(json.dumps({
                    "error": "Token expired and re-authentication failed. Run /msgraph:auth manually.",
                    "help": "Or get a token from https://developer.microsoft.com/en-us/graph/graph-explorer"
                }))
                return 1

    # For auth/config commands, run directly (no retry logic needed)
    if args.command in ("auth", "config"):
        return args.func(args)

    # Run the command, with auto-retry on token errors as a fallback
    import io
    import contextlib

    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        ret = args.func(args)
    result_text = output.getvalue()

    if ret != 0 and _is_token_error(result_text):
        # Command failed with a token error — try auto-auth and retry once
        if _auto_auth():
            output2 = io.StringIO()
            with contextlib.redirect_stdout(output2):
                ret = args.func(args)
            result_text = output2.getvalue()

    print(result_text, end="")
    return ret


if __name__ == "__main__":
    sys.exit(main())
