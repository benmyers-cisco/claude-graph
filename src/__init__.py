"""Microsoft Graph SDK helper for AI assistants."""

from .graph_client import (
    get_graph_client,
    get_concurrency_semaphore,
    run_async,
    get_token,
    set_token,
    verify_token,
    print_token_instructions,
    is_token_expired,
)
from .calendar_tools import get_events_in_range, get_local_timezone, async_get_events_in_range
from .room_tools import (
    list_rooms,
    list_room_lists,
    check_room_availability,
    find_available_rooms,
    async_list_rooms,
    async_check_room_availability,
    async_find_available_rooms,
)
from .people_tools import (
    resolve_person,
    resolve_people,
    verify_emails,
    async_verify_emails,
    async_resolve_person,
    get_attendee_locations,
    async_get_attendee_locations,
    get_org_info,
    async_get_org_info,
    get_user_timezones,
    async_get_user_timezones,
)
from .meeting_tools import (
    find_meeting_times,
    check_participant_availability,
    async_find_meeting_times,
    async_check_participant_availability,
)
from .email_tools import search_emails, read_email, async_search_emails, async_read_email
from .auth import authenticate
from .config import get_config, set_config, is_configured

__all__ = [
    "get_graph_client",
    "get_concurrency_semaphore",
    "run_async",
    "get_token",
    "set_token",
    "verify_token",
    "print_token_instructions",
    "is_token_expired",
    "get_events_in_range",
    "get_local_timezone",
    "list_rooms",
    "list_room_lists",
    "check_room_availability",
    "find_available_rooms",
    "resolve_person",
    "resolve_people",
    "verify_emails",
    "async_verify_emails",
    "find_meeting_times",
    "check_participant_availability",
    "async_get_events_in_range",
    "async_list_rooms",
    "async_check_room_availability",
    "async_find_available_rooms",
    "get_attendee_locations",
    "async_get_attendee_locations",
    "get_org_info",
    "async_get_org_info",
    "get_user_timezones",
    "async_get_user_timezones",
    "async_resolve_person",
    "async_find_meeting_times",
    "async_check_participant_availability",
    "search_emails",
    "read_email",
    "async_search_emails",
    "async_read_email",
    "authenticate",
    "get_config",
    "set_config",
    "is_configured",
]
