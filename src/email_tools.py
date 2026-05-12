"""Email tools for Microsoft Graph API.

Search and read email messages via the Messages API.
"""

import asyncio
from typing import Any

from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

from .graph_client import get_graph_client, run_async


def search_emails(
    query: str,
    *,
    max_results: int = 10,
    folder: str | None = None,
    client=None,
) -> list[dict[str, Any]]:
    """Search emails matching a query string.

    Synchronous wrapper around async_search_emails().
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_search_emails(
        query, max_results=max_results, folder=folder, client=client,
    ))


async def async_search_emails(
    query: str,
    *,
    max_results: int = 10,
    folder: str | None = None,
    client=None,
) -> list[dict[str, Any]]:
    """Search emails matching a query string.

    Uses the $search query parameter which supports KQL (Keyword Query Language):
    - from:person@example.com
    - subject:"weekly update"
    - "exact phrase"
    - hasAttachments:true
    - received>=2024-01-01

    Args:
        query: KQL search query string.
        max_results: Maximum number of results to return (default 10).
        folder: Optional folder ID or well-known name (inbox, sentitems, etc.).
        client: Optional GraphServiceClient.

    Returns:
        List of message dicts with id, subject, from, receivedDateTime, bodyPreview, hasAttachments.
    """
    if client is None:
        client = get_graph_client()

    # Graph $search requires the KQL expression wrapped in outer quotes.
    # If the query already contains quotes (e.g., subject:"weekly update"),
    # pass it through as-is since the user is providing raw KQL.
    if '"' in query:
        search_expr = query
    else:
        search_expr = f'"{query}"'

    config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
        query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            search=search_expr,
            top=max_results,
            select=["id", "subject", "from", "receivedDateTime", "bodyPreview", "hasAttachments", "webLink"],
        )
    )

    try:
        if folder:
            result = await client.me.mail_folders.by_mail_folder_id(folder).messages.get(
                request_configuration=config
            )
        else:
            result = await client.me.messages.get(request_configuration=config)
    except ODataError as e:
        raise RuntimeError(f"Graph API error: {e.error.code}: {e.error.message}") from e

    messages = []
    for msg in (result.value or []):
        from_addr = None
        if msg.from_ and msg.from_.email_address:
            from_addr = {
                "name": msg.from_.email_address.name,
                "address": msg.from_.email_address.address,
            }
        messages.append({
            "id": msg.id,
            "subject": msg.subject,
            "from": from_addr,
            "received": msg.received_date_time.isoformat() if msg.received_date_time else None,
            "body_preview": msg.body_preview,
            "has_attachments": msg.has_attachments,
            "web_link": msg.web_link,
        })

    return messages


def read_email(
    message_id: str,
    *,
    client=None,
) -> dict[str, Any]:
    """Read the full content of an email message.

    Synchronous wrapper around async_read_email().
    """
    if client is None:
        client = get_graph_client()
    return run_async(async_read_email(message_id, client=client))


async def async_read_email(
    message_id: str,
    *,
    client=None,
) -> dict[str, Any]:
    """Read the full content of an email message by ID.

    Returns the full body (HTML and/or text), all recipients, and metadata.

    Args:
        message_id: The message ID from a search result.
        client: Optional GraphServiceClient.

    Returns:
        Dict with full message content including body, recipients, attachments list.
    """
    if client is None:
        client = get_graph_client()

    from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder

    config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
        query_parameters=MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            expand=["attachments"],
        )
    )

    try:
        msg = await client.me.messages.by_message_id(message_id).get(request_configuration=config)
    except ODataError as e:
        raise RuntimeError(f"Graph API error: {e.error.code}: {e.error.message}") from e

    # From
    from_addr = None
    if msg.from_ and msg.from_.email_address:
        from_addr = {
            "name": msg.from_.email_address.name,
            "address": msg.from_.email_address.address,
        }

    # Recipients
    def _extract_recipients(recipients):
        if not recipients:
            return []
        return [
            {
                "name": r.email_address.name if r.email_address else None,
                "address": r.email_address.address if r.email_address else None,
            }
            for r in recipients
        ]

    # Body
    body_content = None
    body_type = None
    if msg.body:
        body_content = msg.body.content
        body_type = msg.body.content_type.value if msg.body.content_type else None

    # Attachments (metadata only — not downloading content)
    attachments = []
    if msg.has_attachments and msg.attachments:
        for att in msg.attachments:
            attachments.append({
                "id": att.id,
                "name": att.name,
                "content_type": att.content_type,
                "size": att.size,
            })

    return {
        "id": msg.id,
        "subject": msg.subject,
        "from": from_addr,
        "to": _extract_recipients(msg.to_recipients),
        "cc": _extract_recipients(msg.cc_recipients),
        "received": msg.received_date_time.isoformat() if msg.received_date_time else None,
        "sent": msg.sent_date_time.isoformat() if msg.sent_date_time else None,
        "has_attachments": msg.has_attachments,
        "attachments": attachments,
        "body_type": body_type,
        "body": body_content,
        "web_link": msg.web_link,
    }
