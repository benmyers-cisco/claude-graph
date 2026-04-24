"""SharePoint and OneDrive document search via Microsoft Graph API."""

from __future__ import annotations

import requests
from datetime import datetime, timedelta, timezone

from .graph_client import get_token


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    token = get_token()
    if not token:
        raise ValueError("No access token available.")
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _requests_get_with_retry(url: str, headers: dict, params: dict | None = None, max_retries: int = 2) -> requests.Response:
    """GET with retry on 429/503/504."""
    for attempt in range(max_retries + 1):
        r = requests.get(url, headers=headers, params=params, timeout=30)
        if r.status_code in (429, 503, 504) and attempt < max_retries:
            import time
            retry_after = int(r.headers.get("Retry-After", 2))
            time.sleep(min(retry_after, 5))
            continue
        return r
    return r


def search_documents(
    query: str,
    max_results: int = 20,
    file_type: str | None = None,
) -> dict:
    """Search for documents across SharePoint and OneDrive.

    Args:
        query: Search query string.
        max_results: Maximum number of results.
        file_type: Optional file extension filter (e.g., "docx", "xlsx", "pptx", "pdf").

    Returns:
        Dict with total count and list of document results.
    """
    h = _headers()

    # Build the query string with optional file type filter
    q = query
    if file_type:
        q = f"{query} filetype:{file_type}"

    payload = {
        "requests": [{
            "entityTypes": ["driveItem"],
            "query": {"queryString": q},
            "from": 0,
            "size": max_results,
        }]
    }

    r = requests.post(f"{GRAPH_BASE}/search/query", headers=h, json=payload, timeout=30)
    r.raise_for_status()
    data = r.json()

    results = []
    total = 0

    for container in data.get("value", []):
        for hc in container.get("hitsContainers", []):
            total = hc.get("total", 0)
            for hit in hc.get("hits", []):
                res = hit.get("resource", {})
                results.append(_format_search_hit(res, hit))

    return {"query": query, "total": total, "count": len(results), "documents": results}


def get_recent_documents(
    days: int = 7,
    max_results: int = 20,
) -> dict:
    """Get recently modified documents from OneDrive and shared with me.

    Args:
        days: Number of days back to look.
        max_results: Maximum results.

    Returns:
        Dict with documents from OneDrive recent and shared.
    """
    h = _headers()

    # Get recent files from OneDrive
    r = _requests_get_with_retry(
        f"{GRAPH_BASE}/me/drive/recent",
        headers=h,
        params={"$top": max_results},
    )
    r.raise_for_status()
    data = r.json()

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    results = []

    for item in data.get("value", []):
        modified = item.get("lastModifiedDateTime")
        if modified:
            mod_dt = datetime.fromisoformat(modified.replace("Z", "+00:00"))
            if mod_dt < cutoff:
                continue

        results.append(_format_drive_item(item))

    return {"days": days, "count": len(results), "documents": results}


def get_site_documents(
    site_query: str,
    path: str | None = None,
    max_results: int = 20,
) -> dict:
    """List documents in a SharePoint site's document library.

    Args:
        site_query: Site name or URL fragment to search for.
        path: Optional path within the document library.
        max_results: Maximum results.

    Returns:
        Dict with site info and documents.
    """
    h = _headers()

    # Find the site
    r = _requests_get_with_retry(
        f"{GRAPH_BASE}/sites",
        headers=h,
        params={"search": site_query, "$top": 5},
    )
    r.raise_for_status()
    sites = r.json().get("value", [])

    if not sites:
        return {"error": f"No site found matching '{site_query}'", "sites": []}

    site = sites[0]
    site_id = site["id"]

    # Get the default document library
    r2 = _requests_get_with_retry(
        f"{GRAPH_BASE}/sites/{site_id}/drive",
        headers=h,
    )
    r2.raise_for_status()
    drive = r2.json()
    drive_id = drive["id"]

    # List items
    if path:
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path}:/children"
    else:
        url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

    r3 = _requests_get_with_retry(
        url,
        headers=h,
        params={"$top": max_results, "$orderby": "lastModifiedDateTime desc"},
    )
    r3.raise_for_status()
    items = r3.json().get("value", [])

    return {
        "site": {
            "name": site.get("displayName"),
            "url": site.get("webUrl"),
            "id": site_id,
        },
        "count": len(items),
        "documents": [_format_drive_item(item) for item in items],
    }


def _format_search_hit(resource: dict, hit: dict) -> dict:
    """Format a search API hit into a clean dict."""
    # Extract site/library info from the resource path
    parent_ref = resource.get("parentReference", {})

    created_by = resource.get("createdBy", {}).get("user", {})
    modified_by = resource.get("lastModifiedBy", {}).get("user", {})

    return {
        "name": resource.get("name"),
        "url": resource.get("webUrl"),
        "modified": resource.get("lastModifiedDateTime"),
        "modified_by": modified_by.get("displayName"),
        "created_by": created_by.get("displayName"),
        "size": resource.get("size"),
        "site_name": parent_ref.get("siteId", "").split(",")[0] if parent_ref.get("siteId") else None,
        "path": parent_ref.get("path"),
        "summary": hit.get("summary"),
    }


def _format_drive_item(item: dict) -> dict:
    """Format a drive item into a clean dict."""
    modified_by = item.get("lastModifiedBy", {}).get("user", {})
    created_by = item.get("createdBy", {}).get("user", {})
    is_folder = "folder" in item

    result = {
        "name": item.get("name"),
        "url": item.get("webUrl"),
        "modified": item.get("lastModifiedDateTime"),
        "modified_by": modified_by.get("displayName"),
        "created_by": created_by.get("displayName"),
        "size": item.get("size"),
        "is_folder": is_folder,
    }

    if is_folder:
        result["child_count"] = item.get("folder", {}).get("childCount")

    return result
