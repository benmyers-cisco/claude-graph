import base64
import json
import logging
import re
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote

import requests

from .graph_client import get_auth_headers, requests_get_with_retry, requests_post_with_retry, requests_put_with_retry

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TEMP_DIR = Path.home() / ".config" / "claude-graph" / "tmp"


@dataclass
class DriveItemRef:
    """Reference to a resolved driveItem with enough info to make API calls."""
    drive_id: str
    item_id: str
    name: str
    mime_type: str
    size: int
    web_url: str
    etag: str = ""

    @property
    def base_url(self) -> str:
        """Graph API base URL for this item."""
        return f"{GRAPH_BASE}/drives/{self.drive_id}/items/{self.item_id}"

    @classmethod
    def from_api_response(cls, data: dict) -> "DriveItemRef":
        """Create from a Graph API driveItem response."""
        return cls(
            drive_id=data["parentReference"]["driveId"],
            item_id=data["id"],
            name=data.get("name", ""),
            mime_type=(data.get("file", {}) or {}).get("mimeType", ""),
            size=data.get("size", 0),
            web_url=data.get("webUrl", ""),
            etag=data.get("eTag", ""),
        )

    @classmethod
    def from_search_hit(cls, resource: dict) -> "DriveItemRef":
        """Create from a Microsoft Search API hit resource."""
        return cls(
            drive_id=resource.get("parentReference", {}).get("driveId", ""),
            item_id=resource.get("id", ""),
            name=resource.get("name", ""),
            mime_type=(resource.get("file", {}) or {}).get("mimeType", ""),
            size=resource.get("size", 0),
            web_url=resource.get("webUrl", ""),
            etag=resource.get("eTag", ""),
        )


def encode_sharing_url(url: str) -> str:
    """Encode a SharePoint sharing URL to a Graph API share token.

    Converts a URL to base64url format and prefixes with 'u!' for use
    with the /shares/{token}/driveItem endpoint.

    NOTE: This only works with sharing URLs, not arbitrary browser URLs.
    For items found via search, use DriveItemRef.from_search_hit() instead.
    """
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=")


def resolve_drive_item(url: str, token: str | None = None) -> DriveItemRef:
    """Resolve a SharePoint sharing URL to a DriveItemRef.

    Uses the /shares/{token}/driveItem endpoint. Only works with sharing URLs
    (the kind you get from SharePoint's "Copy link" or "Share" UI).

    Returns:
        DriveItemRef with drive_id, item_id, name, mime_type, size, web_url.
    """
    share_token = encode_sharing_url(url)
    headers = get_auth_headers(token)
    api_url = f"{GRAPH_BASE}/shares/{share_token}/driveItem"
    r = requests_get_with_retry(api_url, headers=headers)
    _check_read_scope_error(r)
    r.raise_for_status()
    return DriveItemRef.from_api_response(r.json())


READ_SCOPE_ERROR_MESSAGE = """\
Insufficient permissions for SharePoint file access.

To fix this:
1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Click the "Modify permissions" tab (gear icon)
3. Consent to these permissions:
   - Files.Read.All (read all files user can access)
   - Sites.Read.All (search across SharePoint sites)
4. Re-authenticate: /msgraph:auth
"""

WRITE_SCOPE_ERROR_MESSAGE = """\
Insufficient permissions for SharePoint file writes.

To fix this:
1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Click the "Modify permissions" tab (gear icon)
3. Consent to these permissions:
   - Files.ReadWrite.All (read and write all files user can access)
   - Sites.Read.All (search across SharePoint sites)
4. Re-authenticate: /msgraph:auth
"""

_SCOPE_ERROR_CODES = ("Authorization_RequestDenied", "AccessDenied", "ErrorAccessDenied")


class UploadConflictError(Exception):
    """Raised when upload fails due to eTag mismatch (HTTP 412).

    Attributes:
        current_etag: The file's current eTag (for retry).
        last_modified: ISO datetime of the last modification.
        last_modified_by: Display name of who last modified the file.
    """

    def __init__(
        self,
        message: str,
        current_etag: str = "",
        last_modified: str = "",
        last_modified_by: str = "",
    ):
        super().__init__(message)
        self.current_etag = current_etag
        self.last_modified = last_modified
        self.last_modified_by = last_modified_by


class FileLockError(Exception):
    """Raised when upload fails because the file is locked (HTTP 423) after retries.

    Attributes:
        retry_count: Number of retries attempted before giving up.
    """

    def __init__(self, message: str, retry_count: int = 0):
        super().__init__(message)
        self.retry_count = retry_count


def _check_read_scope_error(response: requests.Response) -> None:
    """Raise a clear error if the response indicates missing read permissions."""
    if response.status_code != 403:
        return
    try:
        error = response.json().get("error", {})
        code = error.get("code", "")
    except (json.JSONDecodeError, KeyError, AttributeError):
        return
    if code in _SCOPE_ERROR_CODES:
        raise ValueError(READ_SCOPE_ERROR_MESSAGE)


def _check_write_scope_error(response: requests.Response) -> None:
    """Raise a clear error if the response indicates missing write permissions."""
    if response.status_code != 403:
        return
    try:
        error = response.json().get("error", {})
        code = error.get("code", "")
    except (json.JSONDecodeError, KeyError, AttributeError):
        return
    if code in _SCOPE_ERROR_CODES:
        raise ValueError(WRITE_SCOPE_ERROR_MESSAGE)


def resolve_drive_item_by_ids(
    drive_id: str, item_id: str, token: str | None = None
) -> DriveItemRef:
    """Resolve a driveItem by drive ID and item ID (from search results).

    Makes a single GET to /drives/{driveId}/items/{itemId} to fetch full
    metadata including mime_type, which is needed for accurate file type
    detection.
    """
    headers = get_auth_headers(token)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    r = requests_get_with_retry(url, headers=headers)
    _check_read_scope_error(r)
    r.raise_for_status()
    return DriveItemRef.from_api_response(r.json())


def _purge_temp_dir(max_age_hours: int = 4) -> None:
    """Remove temp files older than max_age_hours from the shared temp directory.

    Called at the start of sharepoint-read to prevent accumulation.
    """
    if not TEMP_DIR.exists():
        return
    cutoff = time.time() - (max_age_hours * 3600)
    for f in TEMP_DIR.iterdir():
        try:
            if f.is_file() and f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def detect_file_type(ref: DriveItemRef) -> str:
    """Detect file type category from DriveItemRef metadata.

    Returns one of: 'excel', 'word', 'powerpoint', 'pdf', 'text', 'onenote', 'unknown'
    """
    mime = ref.mime_type
    name = ref.name.lower()

    if "spreadsheet" in mime or name.endswith((".xlsx", ".xls")):
        return "excel"
    if "wordprocessing" in mime or name.endswith((".docx", ".doc")):
        return "word"
    if "presentation" in mime or name.endswith((".pptx", ".ppt")):
        return "powerpoint"
    if mime == "application/pdf" or name.endswith(".pdf"):
        return "pdf"
    if mime.startswith("text/") or mime == "application/json" or name.endswith(
        (".txt", ".csv", ".json", ".md", ".xml", ".yaml", ".yml")
    ):
        return "text"
    if "onenote" in mime or name.endswith(".one"):
        return "onenote"
    return "unknown"


def _resolve_worksheet_id(
    ref: DriveItemRef, sheet_name: str, headers: dict[str, str]
) -> str:
    """Resolve a worksheet name to its ID.

    The Graph Workbook API needs worksheet IDs in URL paths to avoid escaping
    issues with names containing spaces, #, or apostrophes.

    Raises ValueError if the sheet name is not found.
    """
    r = requests_get_with_retry(
        f"{ref.base_url}/workbook/worksheets", headers=headers
    )
    r.raise_for_status()
    for ws in r.json().get("value", []):
        if ws["name"] == sheet_name:
            return ws["id"]
    available = [ws["name"] for ws in r.json().get("value", [])]
    raise ValueError(
        f"Worksheet '{sheet_name}' not found. Available: {available}"
    )


def list_worksheets(ref: DriveItemRef, token: str | None = None) -> list[dict]:
    """List worksheet names and IDs in an Excel file."""
    headers = get_auth_headers(token)
    r = requests_get_with_retry(
        f"{ref.base_url}/workbook/worksheets", headers=headers
    )
    r.raise_for_status()
    return [{"name": ws["name"], "id": ws["id"]} for ws in r.json().get("value", [])]


def read_worksheet(
    ref: DriveItemRef,
    sheet: str | None = None,
    cell_range: str | None = None,
    token: str | None = None,
) -> dict:
    """Read Excel worksheet data as JSON.

    Args:
        ref: Resolved DriveItemRef for the Excel file.
        sheet: Worksheet name. If None, reads the first sheet.
        cell_range: Optional cell range (e.g., "A1:D10"). If None, reads usedRange.
        token: Optional access token.

    Returns:
        Dict with 'sheet_name', 'address', 'values' (raw — dates are serial numbers),
        'text' (formatted as displayed in Excel — dates are human-readable),
        'row_count', and 'column_count'.
    """
    headers = get_auth_headers(token)

    if sheet:
        # Resolve name → ID to avoid URL escaping issues
        ws_id = _resolve_worksheet_id(ref, sheet, headers)
        sheet_path = f"/workbook/worksheets/{ws_id}"
        sheet_name = sheet
    else:
        # Get first worksheet by ID
        r = requests_get_with_retry(
            f"{ref.base_url}/workbook/worksheets", headers=headers
        )
        r.raise_for_status()
        worksheets = r.json().get("value", [])
        if not worksheets:
            raise ValueError("No worksheets found in workbook")
        sheet_path = f"/workbook/worksheets/{worksheets[0]['id']}"
        sheet_name = worksheets[0]["name"]

    if cell_range:
        if not re.match(r'^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$', cell_range):
            raise ValueError(f"Invalid cell range format: '{cell_range}'. Expected e.g. 'A1' or 'A1:D10'.")
        range_url = f"{ref.base_url}{sheet_path}/range(address='{cell_range}')"
    else:
        range_url = f"{ref.base_url}{sheet_path}/usedRange"

    r = requests_get_with_retry(range_url, headers=headers)
    r.raise_for_status()
    data = r.json()

    return {
        "file_name": ref.name,
        "sheet_name": sheet_name,
        "address": data.get("address", ""),
        "text": data.get("text", []),
        "values": data.get("values", []),
        "row_count": data.get("rowCount", 0),
        "column_count": data.get("columnCount", 0),
    }


def _convert_docx_to_html(content: bytes) -> str:
    """Convert a .docx file to semantic HTML using mammoth."""
    import io
    import mammoth

    result = mammoth.convert_to_html(io.BytesIO(content))
    return result.value


def _extract_docx_text(content: bytes) -> str:
    """Extract plain text from a .docx file using python-docx.

    Emergency fallback when mammoth conversion fails.
    Returns paragraphs joined by newlines.
    """
    import io
    from docx import Document

    doc = Document(io.BytesIO(content))
    return "\n".join(p.text for p in doc.paragraphs if p.text)


def read_word_content(ref: DriveItemRef, token: str | None = None) -> dict:
    """Read a Word document by downloading the raw .docx and converting locally.

    Primary: mammoth converts to semantic HTML (headings, lists, tables).
    Fallback: python-docx extracts plain text if mammoth fails.
    Legacy .doc files are not supported (mammoth/python-docx are .docx-only).

    Returns dict with 'file_name' and either 'html_content' or 'content'.
    """
    if ref.name.lower().endswith(".doc") and not ref.name.lower().endswith(".docx"):
        return {
            "file_name": ref.name,
            "readable": False,
            "message": (
                "Legacy .doc format is not supported for inline reading. "
                "Use sharepoint-download to save it locally, then open in Word or LibreOffice."
            ),
            "web_url": ref.web_url,
        }

    headers = get_auth_headers(token)
    r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
    r.raise_for_status()

    try:
        html = _convert_docx_to_html(r.content)
        return {
            "file_name": ref.name,
            "html_content": html,
            "conversion_mode": "html",
        }
    except Exception:
        logger.warning("mammoth conversion failed for %s, falling back to text extraction", ref.name, exc_info=True)
        text = _extract_docx_text(r.content)
        return {
            "file_name": ref.name,
            "content": text,
            "conversion_mode": "text",
        }


def _save_to_temp(content: bytes, stem: str, suffix: str) -> str:
    """Save binary content to the shared temp directory.

    Returns the path to the temp file.
    """
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    tmp = tempfile.NamedTemporaryFile(
        dir=TEMP_DIR, suffix=suffix, prefix=f"{stem}_", delete=False
    )
    tmp.write(content)
    tmp.close()
    return tmp.name


def _convert_pptx_to_markdown(file_path: str) -> str:
    """Convert a .pptx file to markdown using markitdown."""
    from markitdown import MarkItDown

    md = MarkItDown()
    result = md.convert(file_path)
    return result.text_content


def read_powerpoint_content(ref: DriveItemRef, token: str | None = None) -> dict:
    """Read a PowerPoint file with format-appropriate conversion.

    .pptx: markitdown text extraction (returns inline content), PDF fallback.
    .ppt: PDF conversion only (legacy binary format).
    """
    headers = get_auth_headers(token)
    name_lower = ref.name.lower()

    # Legacy .ppt — PDF conversion only
    if name_lower.endswith(".ppt") and not name_lower.endswith(".pptx"):
        r = requests_get_with_retry(
            f"{ref.base_url}/content?format=pdf", headers=headers
        )
        r.raise_for_status()
        path = _save_to_temp(r.content, Path(ref.name).stem, ".pdf")
        return {
            "file_name": ref.name,
            "temp_path": path,
            "size_bytes": len(r.content),
        }

    # Modern .pptx — try markitdown, fall back to PDF
    r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
    r.raise_for_status()

    tmp_path = _save_to_temp(r.content, Path(ref.name).stem, ".pptx")
    try:
        content = _convert_pptx_to_markdown(tmp_path)
        return {
            "file_name": ref.name,
            "content": content,
        }
    except Exception:
        logger.warning("markitdown conversion failed for %s, falling back to PDF", ref.name, exc_info=True)
    finally:
        try:
            Path(tmp_path).unlink()
        except OSError:
            pass

    # Fallback: PDF conversion
    r = requests_get_with_retry(
        f"{ref.base_url}/content?format=pdf", headers=headers
    )
    r.raise_for_status()
    path = _save_to_temp(r.content, Path(ref.name).stem, ".pdf")
    return {
        "file_name": ref.name,
        "temp_path": path,
        "size_bytes": len(r.content),
    }


def read_pdf_content(ref: DriveItemRef, token: str | None = None) -> dict:
    """Download a PDF to temp.

    Returns dict with 'file_name', 'temp_path', and 'size_bytes'.
    """
    headers = get_auth_headers(token)
    r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
    r.raise_for_status()

    path = _save_to_temp(r.content, Path(ref.name).stem, ".pdf")
    return {
        "file_name": ref.name,
        "temp_path": path,
        "size_bytes": len(r.content),
    }


def read_text_content(ref: DriveItemRef, token: str | None = None) -> dict:
    """Read a text-based file inline.

    Returns dict with 'file_name' and 'content'.
    """
    headers = get_auth_headers(token)
    r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
    r.raise_for_status()
    return {
        "file_name": ref.name,
        "content": r.text,
    }


def download_file(
    url: str | None = None,
    output_path: str | None = None,
    drive_id: str | None = None,
    item_id: str | None = None,
    token: str | None = None,
) -> dict:
    """Download a SharePoint file to a specified path on disk.

    Accepts either a sharing URL or drive_id + item_id (from search results).

    Args:
        url: SharePoint sharing URL.
        output_path: Where to save. If None, saves to CWD with original filename.
        drive_id: Drive ID (from search results). Requires item_id.
        item_id: Item ID (from search results). Requires drive_id.
        token: Optional access token.

    Returns:
        Dict with 'file_name', 'path', 'size_bytes', and 'etag'.
    """
    if drive_id and item_id:
        ref = resolve_drive_item_by_ids(drive_id, item_id, token)
    elif url:
        ref = resolve_drive_item(url, token)
    else:
        raise ValueError("Provide either a URL or both drive_id and item_id.")
    headers = get_auth_headers(token)
    r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
    r.raise_for_status()

    file_name = ref.name or "download"
    if output_path is None:
        output_path = file_name

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_bytes(r.content)

    return {
        "file_name": file_name,
        "path": str(Path(output_path).resolve()),
        "size_bytes": len(r.content),
        "etag": ref.etag,
    }


def upload_file(
    ref: DriveItemRef,
    content: bytes,
    local_filename: str,
    etag: str | None = None,
    force: bool = False,
    token: str | None = None,
    lock_retries: int = 3,
    lock_retry_wait: int = 3,
) -> dict:
    """Upload content to replace an existing SharePoint/OneDrive file.

    Args:
        ref: Resolved DriveItemRef for the target file.
        content: Raw bytes to upload.
        local_filename: Name of the local file (for extension check).
        etag: Optional eTag for concurrency check (sends If-Match header).
        force: Skip size sanity check.
        token: Optional access token.
        lock_retries: Number of retries on 423 (file locked). Default 3.
        lock_retry_wait: Seconds to wait between lock retries. Default 3.

    Returns:
        Dict with file_name, size_bytes, last_modified, web_url.

    Raises:
        ValueError: On extension mismatch, suspicious size decrease, or scope error.
        UploadConflictError: On eTag mismatch (HTTP 412) with current file metadata.
        FileLockError: On file lock (HTTP 423) after retries exhausted.
    """
    headers = get_auth_headers(token)

    # Pre-flight: fetch current item metadata
    r = requests_get_with_retry(ref.base_url, headers=headers)
    _check_write_scope_error(r)
    r.raise_for_status()
    current = r.json()

    # Extension check
    current_ext = Path(current.get("name", "")).suffix.lower()
    upload_ext = Path(local_filename).suffix.lower()
    if current_ext and upload_ext and current_ext != upload_ext:
        raise ValueError(
            f"File extension mismatch: uploading '{upload_ext}' over '{current_ext}'. "
            "Use the same file type or check your file."
        )

    # Size sanity check
    current_size = current.get("size", 0)
    if not force and current_size > 0:
        if len(content) < current_size * 0.1:
            raise ValueError(
                f"New file ({len(content)} bytes) is >90% smaller than the original "
                f"({current_size} bytes). This may indicate corruption or wrong file. "
                "Use force=True to override."
            )

    # Upload with retry on 423 (file locked)
    upload_headers = {**headers, "Content-Type": "application/octet-stream"}
    if etag:
        upload_headers["If-Match"] = etag

    for attempt in range(lock_retries + 1):
        r = requests_put_with_retry(
            f"{ref.base_url}/content", upload_headers, data=content
        )
        _check_write_scope_error(r)

        if r.status_code == 423:
            if attempt < lock_retries:
                time.sleep(lock_retry_wait)
                continue
            attempts = lock_retries + 1
            total_wait = lock_retries * lock_retry_wait
            raise FileLockError(
                f"File is locked after {attempts} attempts ({total_wait}s). "
                "Close the file in the browser and retry.",
                retry_count=lock_retries,
            )
        break

    if r.status_code == 412:
        # Re-fetch metadata to provide current state for retry
        current_etag = ""
        last_modified = ""
        last_modified_by = ""
        meta_r = requests_get_with_retry(ref.base_url, headers=headers)
        if meta_r.status_code == 200:
            meta = meta_r.json()
            current_etag = meta.get("eTag", "")
            last_modified = meta.get("lastModifiedDateTime", "")
            lmb = meta.get("lastModifiedBy", {})
            if lmb:
                last_modified_by = lmb.get("user", {}).get("displayName", "")
        raise UploadConflictError(
            "File was modified since you last read it.",
            current_etag=current_etag,
            last_modified=last_modified,
            last_modified_by=last_modified_by,
        )

    if r.status_code == 413:
        raise ValueError("File exceeds upload size limit.")
    r.raise_for_status()

    data = r.json()
    return {
        "file_name": data.get("name", ref.name),
        "size_bytes": data.get("size", len(content)),
        "last_modified": data.get("lastModifiedDateTime", ""),
        "web_url": data.get("webUrl", ref.web_url),
    }


def read_file(
    url: str,
    sheet: str | None = None,
    cell_range: str | None = None,
    list_sheets: bool = False,
    token: str | None = None,
) -> dict:
    """Read a SharePoint file, auto-detecting type and routing to the right reader.

    This is the main entry point for `sharepoint-read` when given a URL.
    Resolves the URL once, detects file type, and delegates.

    Purges stale temp files before reading to prevent accumulation.
    """
    _purge_temp_dir()
    ref = resolve_drive_item(url, token)
    return read_file_by_ref(ref, sheet=sheet, cell_range=cell_range,
                            list_sheets=list_sheets, token=token, _skip_purge=True)


def read_file_by_ref(
    ref: DriveItemRef,
    sheet: str | None = None,
    cell_range: str | None = None,
    list_sheets: bool = False,
    token: str | None = None,
    _skip_purge: bool = False,
) -> dict:
    """Read a file from a pre-resolved DriveItemRef.

    Used by both read_file (URL-based) and the search→read flow.

    Args:
        ref: Pre-resolved DriveItemRef.
        sheet: Worksheet name for Excel files.
        cell_range: Cell range for Excel files.
        list_sheets: If True, list worksheets instead of reading data.
        token: Optional access token.
        _skip_purge: Internal flag to avoid duplicate purge calls.
    """
    if not _skip_purge:
        _purge_temp_dir()
    file_type = detect_file_type(ref)

    if file_type == "excel":
        if list_sheets:
            sheets = list_worksheets(ref, token)
            result = {"file_type": "excel", "file_name": ref.name, "worksheets": sheets}
        else:
            ws_result = read_worksheet(ref, sheet=sheet, cell_range=cell_range, token=token)
            result = {"file_type": "excel", **ws_result}
    elif file_type == "word":
        word_result = read_word_content(ref, token)
        result = {"file_type": "word", **word_result}
    elif file_type == "powerpoint":
        ppt_result = read_powerpoint_content(ref, token)
        result = {"file_type": "powerpoint", **ppt_result}
    elif file_type == "pdf":
        pdf_result = read_pdf_content(ref, token)
        result = {"file_type": "pdf", **pdf_result}
    elif file_type == "text":
        text_result = read_text_content(ref, token)
        result = {"file_type": "text", **text_result}
    elif file_type == "onenote":
        result = {
            "file_type": "onenote",
            "file_name": ref.name,
            "readable": False,
            "message": "OneNote notebooks cannot be read inline. Open in the browser instead.",
            "web_url": ref.web_url,
        }
    else:
        # Unknown type — download to temp
        headers = get_auth_headers(token)
        r = requests_get_with_retry(f"{ref.base_url}/content", headers=headers)
        r.raise_for_status()
        ext = Path(ref.name).suffix or ".bin"
        path = _save_to_temp(r.content, Path(ref.name).stem, ext)
        result = {
            "file_type": "unknown",
            "file_name": ref.name,
            "temp_path": path,
            "size_bytes": len(r.content),
            "message": "Unrecognized file type, downloaded to temp.",
        }

    if ref.etag:
        result["etag"] = ref.etag
    return result


def search_files(
    query: str,
    file_type: str | None = None,
    max_results: int = 10,
    token: str | None = None,
) -> list[dict]:
    """Search SharePoint/OneDrive for files matching a keyword query.

    Uses the Microsoft Search API to find files across all accessible
    SharePoint sites and OneDrive locations.

    Args:
        query: Search keywords (e.g., "sprint schedule").
        file_type: Optional file type filter. One of: 'excel', 'word',
            'powerpoint', 'pdf', 'text'. Appends KQL filetype: filter.
        max_results: Maximum number of results to return (default: 10).
        token: Optional access token.

    Returns:
        List of dicts, each with:
            - name: File name
            - web_url: SharePoint URL
            - drive_id: Drive ID (for use with sharepoint-read --drive-id)
            - item_id: Item ID (for use with sharepoint-read --item-id)
            - site_name: SharePoint site name
            - last_modified: ISO datetime string
            - last_modified_by: Display name of last modifier
            - size_bytes: File size
            - file_type: Detected file type category
            - readable: Boolean — False for OneNote notebooks (can't be read inline)
    """
    headers = get_auth_headers(token)
    headers["Content-Type"] = "application/json"

    # Build KQL query with optional file type filter
    kql = query
    filetype_map = {
        "excel": "xlsx",
        "word": "docx",
        "powerpoint": "pptx",
        "pdf": "pdf",
        "text": "txt",
    }
    if file_type and file_type in filetype_map:
        kql += f" filetype:{filetype_map[file_type]}"

    body = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": kql},
                "from": 0,
                "size": max_results,
            }
        ]
    }

    r = requests_post_with_retry(f"{GRAPH_BASE}/search/query", headers=headers, json=body)
    _check_read_scope_error(r)
    r.raise_for_status()
    data = r.json()

    results = []
    for response in data.get("value", []):
        for hit_container in response.get("hitsContainers", []):
            for hit in hit_container.get("hits", []):
                resource = hit.get("resource", {})
                has_file = bool(resource.get("file"))
                has_package = bool(resource.get("package"))
                name = resource.get("name", "")
                has_extension = bool(Path(name).suffix)
                # The Search API may omit file/package facets.  Use the
                # filename extension as a fallback signal — folders have
                # no extension.  This misclassifies folders with dots
                # (e.g. "Q1.2024"), but they'll 404 on read, not corrupt.
                if not has_file and not has_package and not has_extension:
                    continue
                ref = DriveItemRef.from_search_hit(resource)
                # OneNote notebooks have a package facet, not file
                if has_package and not has_file:
                    file_type_detected = "onenote"
                    readable = False
                else:
                    file_type_detected = detect_file_type(ref)
                    readable = True

                parent_ref = resource.get("parentReference", {})
                site_name = parent_ref.get("name") or parent_ref.get("sharepointIds", {}).get("siteUrl", "")

                results.append({
                    "name": ref.name,
                    "web_url": ref.web_url,
                    "drive_id": ref.drive_id,
                    "item_id": ref.item_id,
                    "site_name": site_name,
                    "last_modified": resource.get("lastModifiedDateTime"),
                    "last_modified_by": (
                        resource.get("lastModifiedBy", {}).get("user", {}).get("displayName")
                        if resource.get("lastModifiedBy") else None
                    ),
                    "size_bytes": ref.size,
                    "file_type": file_type_detected,
                    "readable": readable,
                })

    return results


_MIME_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pdf": "application/pdf",
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".json": "application/json",
    ".md": "text/markdown",
    ".xml": "application/xml",
}


def get_recent_documents(
    days: int = 7,
    max_results: int = 20,
    token: str | None = None,
) -> list[dict]:
    """Get recently modified documents from OneDrive.

    Args:
        days: Number of days back to look.
        max_results: Maximum results.
        token: Optional access token.

    Returns:
        List of dicts with name, web_url, modified, modified_by, size, is_folder.
    """
    from datetime import datetime, timedelta, timezone

    headers = get_auth_headers(token)
    r = requests_get_with_retry(
        f"{GRAPH_BASE}/me/drive/recent",
        headers=headers,
        params={"$top": max_results},
    )
    r.raise_for_status()

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    results = []

    for item in r.json().get("value", []):
        modified = item.get("lastModifiedDateTime")
        if modified:
            mod_dt = datetime.fromisoformat(modified.replace("Z", "+00:00"))
            if mod_dt < cutoff:
                continue

        modified_by = item.get("lastModifiedBy", {}).get("user", {})
        is_folder = "folder" in item
        results.append({
            "name": item.get("name"),
            "web_url": item.get("webUrl"),
            "modified": modified,
            "modified_by": modified_by.get("displayName"),
            "size": item.get("size"),
            "is_folder": is_folder,
        })

    return results


def get_site_documents(
    site_query: str,
    path: str | None = None,
    max_results: int = 20,
    token: str | None = None,
) -> dict:
    """List documents in a SharePoint site's default document library.

    Args:
        site_query: Site name or URL fragment to search for.
        path: Optional path within the document library.
        max_results: Maximum results.
        token: Optional access token.

    Returns:
        Dict with site info and list of documents.
    """
    headers = get_auth_headers(token)

    r = requests_get_with_retry(
        f"{GRAPH_BASE}/sites",
        headers=headers,
        params={"search": site_query, "$top": 5},
    )
    r.raise_for_status()
    sites = r.json().get("value", [])

    if not sites:
        return {"error": f"No site found matching '{site_query}'", "sites": []}

    site = sites[0]
    site_id = site["id"]

    r2 = requests_get_with_retry(
        f"{GRAPH_BASE}/sites/{site_id}/drive",
        headers=headers,
    )
    r2.raise_for_status()
    drive_id = r2.json()["id"]

    if path:
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{quote(path, safe='/')}:/children"
    else:
        url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

    r3 = requests_get_with_retry(
        url,
        headers=headers,
        params={"$top": max_results, "$orderby": "lastModifiedDateTime desc"},
    )
    r3.raise_for_status()
    items = r3.json().get("value", [])

    documents = []
    for item in items:
        modified_by = item.get("lastModifiedBy", {}).get("user", {})
        is_folder = "folder" in item
        doc = {
            "name": item.get("name"),
            "web_url": item.get("webUrl"),
            "modified": item.get("lastModifiedDateTime"),
            "modified_by": modified_by.get("displayName"),
            "size": item.get("size"),
            "is_folder": is_folder,
        }
        if is_folder:
            doc["child_count"] = item.get("folder", {}).get("childCount")
        documents.append(doc)

    return {
        "site": {
            "name": site.get("displayName"),
            "url": site.get("webUrl"),
            "id": site_id,
        },
        "count": len(documents),
        "documents": documents,
    }


class DestinationFileExistsError(Exception):
    """Raised when upload_new_file would overwrite an existing file.

    Attributes:
        existing_name: Name of the existing file.
        existing_web_url: URL to the existing file.
        existing_last_modified: ISO datetime of last modification.
    """

    def __init__(
        self,
        message: str,
        existing_name: str = "",
        existing_web_url: str = "",
        existing_last_modified: str = "",
    ):
        super().__init__(message)
        self.existing_name = existing_name
        self.existing_web_url = existing_web_url
        self.existing_last_modified = existing_last_modified


def upload_new_file(
    folder_path: str,
    filename: str,
    content: bytes,
    drive_id: str | None = None,
    force: bool = False,
    token: str | None = None,
) -> dict:
    """Create a new file in OneDrive/SharePoint.

    Checks that no file with the same name exists at the destination.
    Raises DestinationFileExistsError if it does, unless force=True.

    Args:
        folder_path: Path relative to drive root (e.g., "Documents/Reports").
            Use "/" or "" for the drive root.
        filename: Name of the new file (e.g., "results.xlsx").
        content: Raw bytes to upload.
        drive_id: Optional drive ID. If None, uses the user's OneDrive.
        force: If True, overwrite an existing file at the destination.
        token: Optional access token.

    Returns:
        Dict with file_name, size_bytes, last_modified, web_url, etag,
        item_id, drive_id.

    Raises:
        DestinationFileExistsError: If a file already exists at the path and force is False.
    """
    headers = get_auth_headers(token)

    # Determine content type from extension
    ext = Path(filename).suffix.lower()
    content_type = _MIME_TYPES.get(ext, "application/octet-stream")

    # Build the item path — encode each segment for reserved chars (#, ?, %)
    folder = folder_path.strip("/")
    encoded_segments = [quote(s, safe="") for s in folder.split("/") if s] if folder else []
    encoded_segments.append(quote(filename, safe=""))
    item_path = "/".join(encoded_segments)

    # Build API URL for the item metadata (without /content)
    if drive_id:
        item_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{item_path}"
    else:
        item_url = f"{GRAPH_BASE}/me/drive/root:/{item_path}"

    # Check if file already exists
    if not force:
        check_r = requests_get_with_retry(item_url, headers=headers)
        if check_r.status_code == 200:
            existing = check_r.json()
            raise DestinationFileExistsError(
                f"A file named '{filename}' already exists at this location. "
                "Use --force to overwrite.",
                existing_name=existing.get("name", filename),
                existing_web_url=existing.get("webUrl", ""),
                existing_last_modified=existing.get("lastModifiedDateTime", ""),
            )
        elif check_r.status_code != 404:
            # Any error other than 404 (file not found) must be surfaced —
            # silently proceeding could overwrite a file we failed to detect.
            _check_write_scope_error(check_r)
            check_r.raise_for_status()

    upload_url = f"{item_url}:/content"
    upload_headers = {**headers, "Content-Type": content_type}
    r = requests_put_with_retry(upload_url, upload_headers, data=content)
    _check_write_scope_error(r)

    if r.status_code == 413:
        raise ValueError("File exceeds upload size limit.")
    r.raise_for_status()

    data = r.json()
    return {
        "file_name": data.get("name", filename),
        "size_bytes": data.get("size", len(content)),
        "last_modified": data.get("lastModifiedDateTime", ""),
        "web_url": data.get("webUrl", ""),
        "etag": data.get("eTag", ""),
        "item_id": data.get("id", ""),
        "drive_id": data.get("parentReference", {}).get("driveId", ""),
    }
