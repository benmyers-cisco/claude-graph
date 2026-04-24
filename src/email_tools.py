"""Email search and retrieval via Microsoft Graph API."""

from __future__ import annotations

import asyncio
from datetime import datetime, timedelta, timezone

from .graph_client import get_graph_client, run_async, get_concurrency_semaphore


def search_emails(
    query: str,
    max_results: int = 20,
    folder: str | None = None,
) -> list[dict]:
    """Search emails using KQL via the Graph API $search parameter.

    Args:
        query: KQL search query (e.g., "from:alice subject:project update").
        max_results: Maximum number of results to return.
        folder: Mail folder to search (default: all folders).

    Returns:
        List of email dicts with id, subject, sender, date, preview, etc.
    """
    return run_async(_async_search_emails(query, max_results, folder))


async def _async_search_emails(
    query: str,
    max_results: int = 20,
    folder: str | None = None,
) -> list[dict]:
    client = get_graph_client()

    # Use the messages endpoint with $search (KQL)
    # Graph Explorer tokens have Mail.Read by default
    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
        query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            search=f'"{query}"',
            top=max_results,
            select=[
                "id", "subject", "from", "toRecipients", "ccRecipients",
                "receivedDateTime", "bodyPreview", "isRead", "importance",
                "hasAttachments", "conversationId", "flag",
            ],
        )
    )

    if folder:
        result = await client.me.mail_folders.by_mail_folder_id(folder).messages.get(
            request_configuration=config
        )
    else:
        result = await client.me.messages.get(request_configuration=config)

    return [_format_message(m) for m in (result.value or [])]


def get_recent_emails(
    days: int = 3,
    max_results: int = 30,
    important_only: bool = False,
    unread_only: bool = False,
) -> list[dict]:
    """Get recent emails with optional filters.

    Args:
        days: Number of days back to look.
        max_results: Maximum results.
        important_only: Only return high-importance emails.
        unread_only: Only return unread emails.

    Returns:
        List of email dicts.
    """
    return run_async(_async_get_recent_emails(days, max_results, important_only, unread_only))


async def _async_get_recent_emails(
    days: int = 3,
    max_results: int = 30,
    important_only: bool = False,
    unread_only: bool = False,
) -> list[dict]:
    client = get_graph_client()

    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    filters = [f"receivedDateTime ge {cutoff}"]
    if important_only:
        filters.append("importance eq 'high'")
    if unread_only:
        filters.append("isRead eq false")

    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
        query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            filter=" and ".join(filters),
            top=max_results,
            select=[
                "id", "subject", "from", "toRecipients", "ccRecipients",
                "receivedDateTime", "bodyPreview", "isRead", "importance",
                "hasAttachments", "conversationId", "flag",
            ],
            orderby=["receivedDateTime desc"],
        )
    )

    result = await client.me.messages.get(request_configuration=config)
    return [_format_message(m) for m in (result.value or [])]


def get_emails_with_person(
    email: str,
    days: int = 14,
    max_results: int = 30,
) -> list[dict]:
    """Get emails exchanged with a specific person.

    Args:
        email: The person's email address.
        days: Number of days back to look.
        max_results: Maximum results.

    Returns:
        List of email dicts, newest first.
    """
    return run_async(_async_get_emails_with_person(email, days, max_results))


async def _async_get_emails_with_person(
    email: str,
    days: int = 14,
    max_results: int = 30,
) -> list[dict]:
    client = get_graph_client()
    sem = get_concurrency_semaphore()

    # Search for emails from and to this person in parallel
    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()

    async def _search(q: str):
        async with sem:
            config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                    search=f'"{q}"',
                    top=max_results,
                    select=[
                        "id", "subject", "from", "toRecipients", "ccRecipients",
                        "receivedDateTime", "bodyPreview", "isRead", "importance",
                        "hasAttachments", "conversationId", "flag",
                    ],
                )
            )
            return await client.me.messages.get(request_configuration=config)

    from_result, to_result = await asyncio.gather(
        _search(f"from:{email}"),
        _search(f"to:{email}"),
    )

    # Merge and deduplicate by message ID
    seen_ids = set()
    messages = []
    cutoff_dt = datetime.now(timezone.utc) - timedelta(days=days)

    for msg in (from_result.value or []) + (to_result.value or []):
        if msg.id in seen_ids:
            continue
        seen_ids.add(msg.id)
        # Filter by date (search doesn't always respect date filters)
        if msg.received_date_time and msg.received_date_time < cutoff_dt:
            continue
        messages.append(msg)

    # Sort by date descending
    messages.sort(
        key=lambda m: m.received_date_time or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True,
    )

    return [_format_message(m) for m in messages[:max_results]]


def _format_message(msg) -> dict:
    """Convert a Graph SDK Message object to a clean dict."""
    sender = None
    if msg.from_:
        ea = msg.from_.email_address
        sender = {
            "name": ea.name if ea else None,
            "email": ea.address if ea else None,
        }

    to = []
    for r in (msg.to_recipients or []):
        ea = r.email_address
        if ea:
            to.append({"name": ea.name, "email": ea.address})

    cc = []
    for r in (msg.cc_recipients or []):
        ea = r.email_address
        if ea:
            cc.append({"name": ea.name, "email": ea.address})

    flag_status = None
    if msg.flag and msg.flag.flag_status:
        flag_status = msg.flag.flag_status.value

    return {
        "id": msg.id,
        "subject": msg.subject,
        "from": sender,
        "to": to,
        "cc": cc,
        "received": msg.received_date_time.isoformat() if msg.received_date_time else None,
        "preview": msg.body_preview,
        "is_read": msg.is_read,
        "importance": msg.importance.value if msg.importance else None,
        "has_attachments": msg.has_attachments,
        "conversation_id": msg.conversation_id,
        "flag": flag_status,
    }
