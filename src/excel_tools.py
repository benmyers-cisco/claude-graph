"""Excel workbook operations via Microsoft Graph API."""

from __future__ import annotations

import requests
import time
from urllib.parse import quote

from .graph_client import get_token


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    token = get_token()
    if not token:
        raise ValueError("No access token available.")
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _wb_url(item_id: str) -> str:
    """Base workbook URL for an item."""
    return f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook"


def _request_with_retry(
    method: str,
    url: str,
    headers: dict,
    json_body: dict | None = None,
    params: dict | None = None,
    max_retries: int = 2,
) -> requests.Response:
    """HTTP request with retry on 429/503/504."""
    for attempt in range(max_retries + 1):
        if method == "GET":
            r = requests.get(url, headers=headers, params=params, timeout=30)
        elif method == "PATCH":
            r = requests.patch(url, headers=headers, json=json_body, timeout=30)
        elif method == "POST":
            r = requests.post(url, headers=headers, json=json_body, timeout=30)
        else:
            raise ValueError(f"Unsupported method: {method}")

        if r.status_code in (429, 503, 504) and attempt < max_retries:
            retry_after = int(r.headers.get("Retry-After", 2))
            time.sleep(min(retry_after, 5))
            continue
        return r
    return r


def get_workbook_info(item_id: str) -> dict:
    """List worksheets and named tables in a workbook.

    Args:
        item_id: SharePoint/OneDrive item ID for the Excel file.

    Returns:
        Dict with worksheets and tables lists.
    """
    h = _headers()
    base = _wb_url(item_id)

    # Fetch worksheets
    r_ws = _request_with_retry("GET", f"{base}/worksheets", headers=h)
    r_ws.raise_for_status()
    worksheets = [
        {"name": ws.get("name"), "id": ws.get("id")}
        for ws in r_ws.json().get("value", [])
    ]

    # Fetch tables
    r_tb = _request_with_retry("GET", f"{base}/tables", headers=h)
    r_tb.raise_for_status()
    tables = []
    for tb in r_tb.json().get("value", []):
        table_entry = {
            "name": tb.get("name"),
            "id": tb.get("id"),
        }
        # Get row count
        tb_name = quote(tb.get("name", ""), safe="")
        r_rows = _request_with_retry(
            "GET", f"{base}/tables/{tb_name}/range", headers=h
        )
        if r_rows.status_code == 200:
            range_data = r_rows.json()
            row_count = range_data.get("rowCount", 0)
            # Subtract 1 for header row if present
            table_entry["rows"] = max(0, row_count - 1)
        tables.append(table_entry)

    return {
        "item_id": item_id,
        "worksheets": worksheets,
        "tables": tables,
    }


def read_range(
    item_id: str,
    sheet: str | None = None,
    range_addr: str | None = None,
) -> dict:
    """Read values from a worksheet range.

    Args:
        item_id: SharePoint/OneDrive item ID.
        sheet: Worksheet name. If None, uses first worksheet.
        range_addr: Excel range like 'A1:L10'. If None, reads used range.

    Returns:
        Dict with worksheet name, range address, dimensions, and values.
    """
    h = _headers()
    base = _wb_url(item_id)

    # Resolve sheet name if not provided
    if not sheet:
        r_ws = _request_with_retry("GET", f"{base}/worksheets", headers=h)
        r_ws.raise_for_status()
        ws_list = r_ws.json().get("value", [])
        if not ws_list:
            raise ValueError("Workbook has no worksheets.")
        sheet = ws_list[0].get("name")

    sheet_encoded = quote(sheet, safe="")

    if range_addr:
        addr_encoded = quote(range_addr, safe="")
        url = f"{base}/worksheets/{sheet_encoded}/range(address='{addr_encoded}')"
    else:
        url = f"{base}/worksheets/{sheet_encoded}/usedRange"

    r = _request_with_retry("GET", url, headers=h)
    r.raise_for_status()
    data = r.json()

    return {
        "worksheet": sheet,
        "range": data.get("address", "").split("!")[-1] if data.get("address") else None,
        "rows": data.get("rowCount", 0),
        "columns": data.get("columnCount", 0),
        "values": data.get("values", []),
    }


def write_range(
    item_id: str,
    sheet: str,
    range_addr: str,
    values: list[list],
) -> dict:
    """Update cell values in a range.

    Args:
        item_id: SharePoint/OneDrive item ID.
        sheet: Worksheet name.
        range_addr: Target range address like 'F5' or 'F5:G6'.
        values: 2D list of values matching the range dimensions.

    Returns:
        Dict with worksheet, range, status, and updated values.
    """
    h = _headers()
    base = _wb_url(item_id)

    sheet_encoded = quote(sheet, safe="")
    addr_encoded = quote(range_addr, safe="")
    url = f"{base}/worksheets/{sheet_encoded}/range(address='{addr_encoded}')"

    r = _request_with_retry("PATCH", url, headers=h, json_body={"values": values})
    r.raise_for_status()
    data = r.json()

    return {
        "worksheet": sheet,
        "range": data.get("address", "").split("!")[-1] if data.get("address") else range_addr,
        "status": "updated",
        "values": data.get("values", values),
    }


def append_rows(
    item_id: str,
    values: list[list],
    sheet: str | None = None,
    table: str | None = None,
) -> dict:
    """Append rows to end of used range or to a named table.

    Args:
        item_id: SharePoint/OneDrive item ID.
        values: 2D list of rows to append.
        sheet: Worksheet name (required if not using table mode).
        table: Named table to append to (uses table API).

    Returns:
        Dict with status, range of new rows, and values.
    """
    h = _headers()
    base = _wb_url(item_id)

    if table:
        # Table mode: POST to table rows endpoint
        table_encoded = quote(table, safe="")
        url = f"{base}/tables/{table_encoded}/rows"
        r = _request_with_retry("POST", url, headers=h, json_body={"values": values})
        r.raise_for_status()
        data = r.json()
        return {
            "table": table,
            "status": "appended",
            "rows_added": len(values),
            "values": data.get("values", values),
        }

    # Sheet mode: find last used row, then write after it
    if not sheet:
        raise ValueError("Either --sheet or --table is required for append.")

    sheet_encoded = quote(sheet, safe="")
    used_url = f"{base}/worksheets/{sheet_encoded}/usedRange"
    r_used = _request_with_retry("GET", used_url, headers=h)
    r_used.raise_for_status()
    used_data = r_used.json()

    row_count = used_data.get("rowCount", 0)

    # Parse the used range address to get the starting column letter
    address = used_data.get("address", "")
    # Address format: "Sheet1!A1:L42" or just "A1:L42"
    range_part = address.split("!")[-1] if "!" in address else address
    start_col = ""
    for ch in range_part:
        if ch.isalpha():
            start_col += ch
        else:
            break
    if not start_col:
        start_col = "A"

    # Calculate end column letter from col_count
    def col_letter(n: int) -> str:
        """Convert 1-based column number to Excel letter (1=A, 27=AA)."""
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    # Start column as number
    start_col_num = 0
    for ch in start_col.upper():
        start_col_num = start_col_num * 26 + (ord(ch) - 64)

    end_col_num = start_col_num + len(values[0]) - 1
    end_col = col_letter(end_col_num)

    next_row = row_count + 1
    last_row = next_row + len(values) - 1
    new_range = f"{start_col}{next_row}:{end_col}{last_row}"

    addr_encoded = quote(new_range, safe="")
    write_url = f"{base}/worksheets/{sheet_encoded}/range(address='{addr_encoded}')"
    r_write = _request_with_retry("PATCH", write_url, headers=h, json_body={"values": values})
    r_write.raise_for_status()
    data = r_write.json()

    return {
        "worksheet": sheet,
        "status": "appended",
        "range": data.get("address", "").split("!")[-1] if data.get("address") else new_range,
        "rows_added": len(values),
        "values": data.get("values", values),
    }
