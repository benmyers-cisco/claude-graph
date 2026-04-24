"""People resolution tools for Microsoft Graph API.

Resolves names and partial email addresses to full email addresses using
the People API and directory lookups.
"""

import re
from typing import Any

from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.people.people_request_builder import PeopleRequestBuilder
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

from .graph_client import get_graph_client, run_async, get_concurrency_semaphore

SELF_REFERENCE_WORDS = frozenset({"me", "myself", "my", "i"})


def get_org_info(
    query: str,
    *,
    include_manager: bool = True,
    include_reports: bool = True,
    domain: str = "cisco.com",
    client=None,
) -> dict[str, Any]:
    """Get manager and direct reports for a person.

    Synchronous wrapper around async_get_org_info().
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_get_org_info(
        query,
        include_manager=include_manager,
        include_reports=include_reports,
        domain=domain,
        client=client,
    ))


async def async_get_org_info(
    query: str,
    *,
    include_manager: bool = True,
    include_reports: bool = True,
    domain: str = "cisco.com",
    client=None,
) -> dict[str, Any]:
    """Get manager and direct reports for a person.

    Resolves the query to a user, then fetches their manager and/or
    direct reports from the Graph directory.

    Args:
        query: Name or partial email to look up.
        include_manager: Whether to fetch the person's manager.
        include_reports: Whether to fetch direct reports.
        domain: Default email domain for resolution.
        client: Optional GraphServiceClient.

    Returns:
        Dictionary with:
        - person: {email, name, job_title, department, office_location}
        - manager: Same shape dict, or None
        - direct_reports: List of same shape dicts
        - direct_reports_count: int
        - errors: List of non-fatal error strings
    """
    if client is None:
        client = get_graph_client()

    # Resolve query to email
    resolution = await async_resolve_person(query, domain=domain, client=client)

    if not resolution.get("best_match"):
        return {"error": "no_match", "input": query, "matches": []}

    if resolution.get("is_ambiguous"):
        return {
            "error": "ambiguous",
            "input": query,
            "matches": [
                {"email": m["email"], "name": m["name"]}
                for m in resolution.get("matches", [])[:5]
            ],
        }

    email = resolution["resolved_email"]
    errors: list[str] = []

    # Fetch the target user's profile
    try:
        user_obj = await client.users.by_user_id(email).get()
        person = _extract_person_dict(user_obj)
    except ODataError as e:
        return {"error": "user_not_found", "input": query, "email": email,
                "detail": e.error.message if e.error else str(e)}

    # Fetch manager
    manager = None
    if include_manager:
        try:
            mgr_obj = await client.users.by_user_id(email).manager.get()
            manager = _extract_person_dict(mgr_obj)
        except ODataError as e:
            code = e.error.code if e.error else ""
            if "404" in str(code) or "Request_ResourceNotFound" in str(code):
                pass  # No manager — that's fine
            else:
                errors.append(f"Could not fetch manager: {e.error.message if e.error else str(e)}")

    # Fetch direct reports
    direct_reports: list[dict[str, Any]] = []
    if include_reports:
        try:
            reports_resp = await client.users.by_user_id(email).direct_reports.graph_user.get()
            for u in reports_resp.value or []:
                direct_reports.append(_extract_person_dict(u))
            # Sort by name
            direct_reports.sort(key=lambda r: (r.get("name") or "").lower())
        except ODataError as e:
            errors.append(f"Could not fetch direct reports: {e.error.message if e.error else str(e)}")

    return {
        "person": person,
        "manager": manager,
        "direct_reports": direct_reports,
        "direct_reports_count": len(direct_reports),
        "errors": errors,
    }


def _extract_person_dict(obj) -> dict[str, Any]:
    """Extract a standard person dict from a User or DirectoryObject."""
    return {
        "email": getattr(obj, "mail", None) or getattr(obj, "user_principal_name", None),
        "name": getattr(obj, "display_name", None),
        "job_title": getattr(obj, "job_title", None),
        "department": getattr(obj, "department", None),
        "office_location": getattr(obj, "office_location", None),
    }


def get_user_timezones(
    emails: list[str],
) -> dict[str, str | None]:
    """Look up mailbox timezone settings for a list of users.

    Synchronous wrapper around async_get_user_timezones().

    Args:
        emails: List of email addresses to look up.

    Returns:
        Dictionary mapping lowercase email to Windows timezone string
        (e.g., {"user@cisco.com": "Eastern Standard Time"}). Value is None
        if the timezone could not be fetched.
    """
    client = get_graph_client()
    return run_async(async_get_user_timezones(emails, client=client))


async def async_get_user_timezones(
    emails: list[str],
    *,
    client=None,
) -> dict[str, str | None]:
    """Look up timezone settings for a list of users via getSchedule API.

    Uses the getSchedule endpoint which returns working hours including
    timezone for each user. This works with delegated Graph Explorer tokens
    (unlike /users/{email}/mailboxSettings which requires admin permissions).

    Batches up to 20 emails per request.

    Args:
        emails: List of email addresses to look up.
        client: Optional GraphServiceClient.

    Returns:
        Dictionary mapping lowercase email to Windows timezone string.
        Value is None if the timezone could not be fetched.
    """
    from datetime import date, timedelta
    from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
        GetSchedulePostRequestBody,
    )
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone

    if client is None:
        client = get_graph_client()

    # Deduplicate
    unique = list({e.strip().lower() for e in emails if e.strip()})
    if not unique:
        return {}

    result: dict[str, str | None] = {}
    batch_size = 20

    # Use a minimal time window — we only need working hours metadata
    today = date.today()
    tomorrow = today + timedelta(days=1)

    for i in range(0, len(unique), batch_size):
        batch = unique[i:i + batch_size]
        try:
            body = GetSchedulePostRequestBody()
            body.schedules = batch
            body.start_time = DateTimeTimeZone(
                date_time=f"{today.isoformat()}T00:00:00",
                time_zone="UTC",
            )
            body.end_time = DateTimeTimeZone(
                date_time=f"{tomorrow.isoformat()}T00:00:00",
                time_zone="UTC",
            )
            body.availability_view_interval = 60

            response = await client.me.calendar.get_schedule.post(body)
            for item in response.value or []:
                key = (item.schedule_id or "").lower()
                if key and item.working_hours and item.working_hours.time_zone:
                    result[key] = item.working_hours.time_zone.name
                elif key:
                    result[key] = None
        except Exception:
            # Best-effort — mark batch as unknown
            for email in batch:
                result.setdefault(email, None)

    return result


def resolve_person(
    query: str,
    *,
    domain: str = "cisco.com",
    max_results: int = 10,
    client=None
) -> dict[str, Any]:
    """Resolve a single name or partial email to full email addresses.

    Synchronous wrapper around async_resolve_person(). NOT safe to call from
    an async context (e.g., MCP server). See async_resolve_person() for full docs.
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_resolve_person(
        query, domain=domain, max_results=max_results, client=client
    ))


async def async_resolve_person(
    query: str,
    *,
    domain: str = "cisco.com",
    max_results: int = 10,
    client=None
) -> dict[str, Any]:
    """Resolve a single name or partial email to full email addresses.

    Searches the People API to find matching people, ranked by relevance.

    Args:
        query: Name or partial email to search (e.g., "john", "jdoe", "john.doe@cisco.com").
        domain: Default email domain to append (default: "cisco.com").
        max_results: Maximum number of matches to return.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        Dictionary with:
        - input: Original query string
        - matches: List of candidate matches with email, name, score, source
        - is_ambiguous: True if multiple high-confidence matches exist
        - best_match: Highest scored match (or None if no matches)
        - resolved_email: Email if unambiguous (or None if ambiguous/no matches)
    """
    if client is None:
        client = get_graph_client()

    query = query.strip()
    if not query:
        return {
            "input": query,
            "matches": [],
            "is_ambiguous": False,
            "best_match": None,
            "resolved_email": None,
        }

    # Self-references resolve to the current authenticated user
    if query.lower() in SELF_REFERENCE_WORDS:
        try:
            me = await client.me.get()
            my_email = (me.mail or me.user_principal_name or "").lower()
            my_name = me.display_name
            if my_email:
                return {
                    "input": query,
                    "matches": [{"email": my_email, "name": my_name, "score": 1.0, "source": "self"}],
                    "is_ambiguous": False,
                    "best_match": {"email": my_email, "name": my_name, "score": 1.0, "source": "self"},
                    "resolved_email": my_email,
                }
        except Exception:
            pass  # Fall through to normal resolution

    # Check if it's already a full email address
    normalized = _normalize_email(query, domain)
    if normalized and "@" in query:
        return {
            "input": query,
            "matches": [{"email": normalized, "name": None, "score": 1.0, "source": "input"}],
            "is_ambiguous": False,
            "best_match": {"email": normalized, "name": None, "score": 1.0, "source": "input"},
            "resolved_email": normalized,
        }

    # Search using People API
    all_results = []
    try:
        people_results = await _async_search_people_api(query, max_results, client)
        all_results.extend(people_results)
    except ODataError:
        pass

    # Merge, deduplicate, and rank results
    merged = _merge_and_rank_results(all_results, query, domain)
    merged = merged[:max_results]

    # Determine if ambiguous (multiple high-confidence matches)
    is_ambiguous = False
    if len(merged) > 1:
        score_diff = merged[0]["score"] - merged[1]["score"]
        if merged[0]["score"] >= 0.6:
            is_ambiguous = score_diff < 0.1
        elif merged[0]["score"] >= 0.4:
            is_ambiguous = score_diff < 0.15
        else:
            is_ambiguous = score_diff < 0.2

    best_match = merged[0] if merged else None
    resolved_email = best_match["email"] if best_match and not is_ambiguous else None

    return {
        "input": query,
        "matches": merged,
        "is_ambiguous": is_ambiguous,
        "best_match": best_match,
        "resolved_email": resolved_email,
    }


def resolve_people(
    queries: list[str],
    *,
    domain: str = "cisco.com",
    strict: bool = False,
    **kwargs
) -> list[dict[str, Any]]:
    """Batch resolve multiple names/emails.

    Note: Due to asyncio event loop limitations, this function may not work reliably
    for multiple queries. For batch resolution, use the async pattern documented in CLAUDE.md.

    Args:
        queries: List of names or partial emails to resolve.
        domain: Default email domain to append.
        strict: If True, raises ValueError on ambiguous matches.
        **kwargs: Additional arguments passed to resolve_person().

    Returns:
        List of resolution results, one per query (same format as resolve_person()).

    Raises:
        ValueError: If strict=True and any query has ambiguous matches.
        RuntimeError: May occur due to asyncio event loop issues with multiple queries.

    Example:
        >>> results = resolve_people(["john", "jane smith", "bob"])
        >>> for r in results:
        ...     print(f"{r['input']} -> {r['resolved_email'] or 'AMBIGUOUS'}")
    """
    results = []
    # Don't reuse client - each call needs a fresh one due to asyncio issues
    kwargs.pop("client", None)

    for query in queries:
        # Each call creates its own client to avoid event loop issues
        result = resolve_person(query, domain=domain, client=None, **kwargs)
        if strict and result["is_ambiguous"]:
            matches_str = ", ".join(
                f"{m['name'] or 'Unknown'} <{m['email']}>" for m in result["matches"][:3]
            )
            raise ValueError(f"Ambiguous match for '{query}': {matches_str}")
        results.append(result)

    return results


def verify_emails(
    emails: list[str],
    *,
    client=None
) -> list[dict[str, Any]]:
    """Verify and normalize emails against directory.

    Synchronous wrapper around async_verify_emails().

    Args:
        emails: List of email addresses to verify.
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        List of verification results with:
        - input: Original email
        - exists: True if found in directory
        - email: Canonical email address (if found)
        - display_name: Full name (if found)
        - job_title: Job title (if available)
        - department: Department (if available)

    Example:
        >>> results = verify_emails(["john.doe@cisco.com", "notreal@cisco.com"])
        >>> for r in results:
        ...     status = "Found" if r['exists'] else "Not found"
        ...     print(f"{r['input']}: {status} - {r['display_name']}")
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_verify_emails(emails, client=client))


async def async_verify_emails(
    emails: list[str],
    *,
    client=None
) -> list[dict[str, Any]]:
    """Verify and normalize emails against directory.

    Batch-fetches all emails concurrently with a concurrency semaphore,
    then returns results in the same order as the input list.

    Args:
        emails: List of email addresses to verify.
        client: Optional GraphServiceClient.

    Returns:
        List of verification results (same order as input).
    """
    import asyncio

    if client is None:
        client = get_graph_client()

    sem = get_concurrency_semaphore()

    async def _verify_one(email: str) -> dict[str, Any]:
        email = email.strip().lower()
        async with sem:
            try:
                config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                    query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                        filter=f"mail eq '{email}' or userPrincipalName eq '{email}'",
                        select=["displayName", "mail", "userPrincipalName", "jobTitle", "department", "officeLocation"],
                        top=1
                    )
                )
                response = await client.users.get(request_configuration=config)

                if response.value:
                    user = response.value[0]
                    return {
                        "input": email,
                        "exists": True,
                        "email": user.mail or user.user_principal_name,
                        "display_name": user.display_name,
                        "job_title": user.job_title,
                        "department": user.department,
                        "office_location": user.office_location,
                    }
                else:
                    return {
                        "input": email,
                        "exists": False,
                        "email": None,
                        "display_name": None,
                        "job_title": None,
                        "department": None,
                        "office_location": None,
                    }
            except ODataError:
                return {
                    "input": email,
                    "exists": None,  # Unknown - couldn't verify
                    "email": email,
                    "display_name": None,
                    "job_title": None,
                    "department": None,
                    "office_location": None,
                }

    return await asyncio.gather(*[_verify_one(e) for e in emails])


def get_attendee_locations(
    emails: list[str],
    building: str,
) -> dict[str, Any]:
    """Look up office locations for attendees and classify relative to a building.

    Synchronous wrapper around async_get_attendee_locations().

    Args:
        emails: List of attendee email addresses.
        building: Target building code (e.g., "ARB02") to classify against.

    Returns:
        Dictionary with:
        - attendees: List of {email, office_location, classification} dicts
        - summary: {total, local, other_office, unknown} counts
    """
    client = get_graph_client()
    return run_async(async_get_attendee_locations(emails, building, client=client))


async def async_get_attendee_locations(
    emails: list[str],
    building: str,
    *,
    client=None,
) -> dict[str, Any]:
    """Look up office locations for attendees and classify relative to a building.

    Batch-fetches officeLocation from /users using OData $filter with 'in' operator
    (up to 15 per request). Falls back to individual queries if 'in' is unsupported.

    Args:
        emails: List of attendee email addresses.
        building: Target building code (e.g., "ARB02") to classify against.
        client: Optional GraphServiceClient.

    Returns:
        Dictionary with:
        - attendees: List of {email, office_location, classification} dicts
        - summary: {total, local, other_office, unknown} counts
    """
    if client is None:
        client = get_graph_client()

    # Batch fetch office locations (up to 15 per OData 'in' query)
    location_map: dict[str, str | None] = {}
    batch_size = 15

    for i in range(0, len(emails), batch_size):
        batch = emails[i:i + batch_size]
        try:
            locations = await _async_batch_fetch_locations(batch, client)
            location_map.update(locations)
        except ODataError:
            # 'in' operator not supported — fall back to individual queries
            import asyncio
            sem = get_concurrency_semaphore()

            async def _fetch_one(em):
                async with sem:
                    return em, await _async_fetch_single_location(em, client)

            results = await asyncio.gather(
                *[_fetch_one(em) for em in batch], return_exceptions=True
            )
            for res in results:
                if isinstance(res, Exception):
                    continue
                em, loc = res
                location_map[em.lower()] = loc

    # Classify each attendee
    attendees = []
    summary = {"total": len(emails), "local": 0, "other_office": 0, "unknown": 0}

    for email in emails:
        office = location_map.get(email.lower())
        classification = _classify_location(office, building)
        summary[classification] += 1
        attendees.append({
            "email": email,
            "office_location": office,
            "classification": classification,
        })

    return {"attendees": attendees, "summary": summary}


async def _async_batch_fetch_locations(
    emails: list[str],
    client,
) -> dict[str, str | None]:
    """Batch fetch officeLocation for up to 15 emails using OData 'in' operator."""
    # Build filter: mail in ('a@x.com', 'b@x.com')
    quoted = ", ".join(f"'{e.strip().lower()}'" for e in emails)
    filter_str = f"mail in ({quoted})"

    config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
        query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=filter_str,
            select=["mail", "userPrincipalName", "officeLocation"],
            top=len(emails),
        )
    )
    response = await client.users.get(request_configuration=config)

    result: dict[str, str | None] = {}
    for user in response.value or []:
        key = (user.mail or user.user_principal_name or "").lower()
        if key:
            result[key] = user.office_location
    return result


async def _async_fetch_single_location(
    email: str,
    client,
) -> str | None:
    """Fetch officeLocation for a single email."""
    config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
        query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=f"mail eq '{email.strip().lower()}'",
            select=["mail", "officeLocation"],
            top=1,
        )
    )
    response = await client.users.get(request_configuration=config)
    if response.value:
        return response.value[0].office_location
    return None


def _matches_building(office_location: str | None, building: str) -> bool:
    """Check whether an officeLocation string matches a building code.

    Handles formats like "ARB02", "ARB02-1", "Building ARB02", etc.
    Uses word-boundary matching (case-insensitive).
    """
    if not office_location or not building:
        return False
    # Word-boundary match for the building code anywhere in the string
    pattern = rf'\b{re.escape(building)}\b'
    return bool(re.search(pattern, office_location, re.IGNORECASE))


def _classify_location(office_location: str | None, building: str) -> str:
    """Classify an officeLocation relative to a target building.

    Returns:
        "local" — officeLocation matches the target building
        "other_office" — officeLocation is set but doesn't match
        "unknown" — officeLocation is null/empty (likely remote or external)
    """
    if not office_location or not office_location.strip():
        return "unknown"
    if _matches_building(office_location, building):
        return "local"
    return "other_office"


# Private helper functions

async def _async_search_people_api(
    query: str,
    top: int,
    client=None
) -> list[dict[str, Any]]:
    """Search using People API (frequently contacted people).

    The People API ranks results by relevance based on communication patterns.
    """
    if client is None:
        client = get_graph_client()

    config = PeopleRequestBuilder.PeopleRequestBuilderGetRequestConfiguration(
        query_parameters=PeopleRequestBuilder.PeopleRequestBuilderGetQueryParameters(
            search=f'"{query}"',
            top=top,
            select=["displayName", "scoredEmailAddresses", "personType"]
        )
    )

    response = await client.me.people.get(request_configuration=config)

    results = []
    for person in response.value or []:
        if person.scored_email_addresses:
            for email_info in person.scored_email_addresses:
                if email_info.address:
                    results.append({
                        "email": email_info.address.lower(),
                        "name": person.display_name,
                        "score": 0.0,  # Will be recalculated
                        "source": "people_api",
                        "relevance_score": email_info.relevance_score or 0,
                    })

    return results


def _merge_and_rank_results(
    results: list[dict[str, Any]],
    query: str,
    domain: str
) -> list[dict[str, Any]]:
    """Merge results from multiple sources, deduplicate, and rank by relevance."""
    # Group by email
    by_email: dict[str, dict[str, Any]] = {}

    for result in results:
        email = result["email"].lower()

        if email not in by_email:
            by_email[email] = result.copy()
            by_email[email]["sources"] = [result["source"]]
        else:
            # Merge sources
            if result["source"] not in by_email[email]["sources"]:
                by_email[email]["sources"].append(result["source"])
            # Keep the best name (non-empty)
            if not by_email[email].get("name") and result.get("name"):
                by_email[email]["name"] = result["name"]
            # Keep relevance score if available (take the higher one)
            if result.get("relevance_score"):
                existing = by_email[email].get("relevance_score", 0)
                by_email[email]["relevance_score"] = max(existing, result["relevance_score"])

    # Normalize relevance scores relative to the maximum
    # This is critical because People API returns raw scores that can vary wildly (e.g., 272 vs 7)
    max_relevance = max((d.get("relevance_score", 0) for d in by_email.values()), default=0)
    if max_relevance > 0:
        for data in by_email.values():
            raw = data.get("relevance_score", 0)
            # Normalize to 0-1 range, with some dampening to not completely dominate
            data["relevance_score_normalized"] = raw / max_relevance

    # Calculate final scores
    for email, data in by_email.items():
        data["score"] = _calculate_match_score(data, query, domain)

    # Sort by score descending
    ranked = sorted(by_email.values(), key=lambda x: x["score"], reverse=True)

    # Clean up temporary fields
    for r in ranked:
        r.pop("sources", None)
        r.pop("relevance_score", None)
        r.pop("relevance_score_normalized", None)
        r.pop("job_title", None)

    return ranked


def _calculate_match_score(
    candidate: dict[str, Any],
    query: str,
    domain: str
) -> float:
    """Calculate relevance score for a candidate match.

    Scoring priorities:
    1. People API relevance (interaction frequency) - strongest signal
    2. Exact first name match - "jamie" should match "Jamie Pringle" not "James"
    3. Domain match - prefer internal domain over external
    4. Name quality - penalize weak matches
    """
    score = 0.0
    query_lower = query.lower()
    email = candidate.get("email", "").lower()
    name = (candidate.get("name") or "").lower()
    sources = candidate.get("sources", [])

    # Extract name parts for analysis
    name_parts = name.split() if name else []
    first_name = name_parts[0] if name_parts else ""
    # Handle "Display Name (username)" format - extract just the name part
    if "(" in name:
        name_only = name.split("(")[0].strip()
        name_parts = name_only.split()
        first_name = name_parts[0] if name_parts else ""

    # 1. PEOPLE API RELEVANCE - strongest signal for who you actually interact with
    # This is the key insight: if you frequently contact someone, that's likely who you mean
    # Use normalized score (0-1) which accounts for relative interaction frequency
    relevance = candidate.get("relevance_score_normalized", 0)
    if relevance:
        # Scale relevance score significantly - this is our best signal
        # A person you contact 40x more often should have a big advantage
        score += relevance * 0.4

    # Base source weight (People API results are inherently more relevant)
    if "people_api" in sources:
        score += 0.15
    if "calendar" in sources:
        score += 0.1
    if "email" in sources:
        score += 0.08

    # Multiple sources = higher confidence
    if len(sources) > 1:
        score += 0.05 * (len(sources) - 1)

    # 2. NAME MATCHING - exact first name match is critical
    if first_name:
        # Exact first name match - very strong signal
        if first_name == query_lower:
            score += 0.35
        # First name starts with query (e.g., "jam" -> "jamie")
        elif first_name.startswith(query_lower) and len(query_lower) >= 3:
            score += 0.25
        # Query starts with first name (e.g., "jamie" -> "jam") - weaker
        elif query_lower.startswith(first_name) and len(first_name) >= 3:
            score += 0.15
        # First name doesn't match at all - PENALTY
        # e.g., searching "jamie" but first name is "james" or "jennifer"
        elif not first_name.startswith(query_lower[:3]) and not query_lower.startswith(first_name[:3]):
            score -= 0.2

    # Last name match (less important than first name)
    if len(name_parts) > 1:
        last_name = name_parts[-1]
        if last_name == query_lower:
            score += 0.2
        elif last_name.startswith(query_lower):
            score += 0.1

    # Full name match
    if name == query_lower:
        score += 0.3

    # 3. EMAIL MATCHING
    if email:
        local_part = email.split("@")[0]
        email_domain = email.split("@")[1] if "@" in email else ""

        # Exact local part match
        if local_part == query_lower:
            score += 0.3
        # Local part starts with query
        elif local_part.startswith(query_lower):
            score += 0.15
        # Query in local part
        elif query_lower in local_part:
            score += 0.05

        # 4. DOMAIN MATCHING - prefer internal domain
        if domain:
            if email_domain == domain:
                score += 0.15  # Internal domain bonus
            else:
                score -= 0.1  # External domain penalty

    # Input source means exact match from user
    if "input" in sources:
        score = 1.0

    return max(0.0, min(score, 1.0))  # Clamp between 0 and 1


def _normalize_email(email: str, domain: str) -> str | None:
    """Normalize an email address, adding domain if needed."""
    email = email.strip().lower()

    if not email:
        return None

    # Already has @ - return as-is
    if "@" in email:
        return email

    # Check if it looks like a local part (no spaces, has valid chars)
    if " " not in email and re.match(r'^[\w.\-+]+$', email):
        return f"{email}@{domain}"

    return None
