# Excel Commands Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `msgraph excel` subcommands (info, read, write, append) so Claude can read and update SharePoint/OneDrive Excel files.

**Architecture:** New `src/excel_tools.py` module using raw `requests` calls (same pattern as `docs_tools.py`). Four CLI subcommands registered under `msgraph excel`. All Graph API calls target `/me/drive/items/{id}/workbook/...`.

**Tech Stack:** Python 3.12, `requests` library, Microsoft Graph REST API v1.0

---

### Task 1: Create `excel_tools.py` with HTTP helpers and `get_workbook_info`

**Files:**
- Create: `src/excel_tools.py`

- [ ] **Step 1: Create `excel_tools.py` with helpers and `get_workbook_info`**

```python
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
```

- [ ] **Step 2: Test manually**

Run: `cd ~/.config/claude-graph && python -c "from src.excel_tools import get_workbook_info; print('import ok')"`
Expected: `import ok`

- [ ] **Step 3: Commit**

```bash
cd ~/.config/claude-graph && git add src/excel_tools.py && git commit -m "feat: add excel_tools module with get_workbook_info"
```

---

### Task 2: Add `read_range` function

**Files:**
- Modify: `src/excel_tools.py`

- [ ] **Step 1: Add `read_range` to `excel_tools.py`**

Append after `get_workbook_info`:

```python
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
```

- [ ] **Step 2: Test import**

Run: `cd ~/.config/claude-graph && python -c "from src.excel_tools import read_range; print('ok')"`
Expected: `ok`

- [ ] **Step 3: Commit**

```bash
cd ~/.config/claude-graph && git add src/excel_tools.py && git commit -m "feat: add read_range to excel_tools"
```

---

### Task 3: Add `write_range` function

**Files:**
- Modify: `src/excel_tools.py`

- [ ] **Step 1: Add `write_range` to `excel_tools.py`**

Append after `read_range`:

```python
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
```

- [ ] **Step 2: Test import**

Run: `cd ~/.config/claude-graph && python -c "from src.excel_tools import write_range; print('ok')"`
Expected: `ok`

- [ ] **Step 3: Commit**

```bash
cd ~/.config/claude-graph && git add src/excel_tools.py && git commit -m "feat: add write_range to excel_tools"
```

---

### Task 4: Add `append_rows` function

**Files:**
- Modify: `src/excel_tools.py`

- [ ] **Step 1: Add `append_rows` to `excel_tools.py`**

Append after `write_range`:

```python
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
    col_count = used_data.get("columnCount", 1)

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
```

- [ ] **Step 2: Test import**

Run: `cd ~/.config/claude-graph && python -c "from src.excel_tools import append_rows; print('ok')"`
Expected: `ok`

- [ ] **Step 3: Commit**

```bash
cd ~/.config/claude-graph && git add src/excel_tools.py && git commit -m "feat: add append_rows to excel_tools"
```

---

### Task 5: Add CLI command handlers in `cli.py`

**Files:**
- Modify: `src/cli.py:38` (add import)
- Modify: `src/cli.py` (add cmd functions after `cmd_docs_site` at ~line 1029)

- [ ] **Step 1: Add import at line 38**

After the existing `from .email_tools import ...` line (line 38), add:

```python
from .excel_tools import get_workbook_info, read_range, write_range, append_rows
```

- [ ] **Step 2: Add command handler functions**

Add these after `cmd_docs_site` (after line 1029) and before `cmd_config`:

```python
def cmd_excel_info(args):
    """List worksheets and tables in an Excel workbook."""
    try:
        result = get_workbook_info(item_id=args.item_id)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_excel_read(args):
    """Read values from an Excel worksheet."""
    try:
        result = read_range(
            item_id=args.item_id,
            sheet=args.sheet,
            range_addr=args.range,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_excel_write(args):
    """Write values to an Excel range."""
    try:
        values = json.loads(args.values)
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"Invalid JSON for --values: {e}"}))
        return 1

    if not isinstance(values, list) or not all(isinstance(row, list) for row in values):
        print(json.dumps({"error": "--values must be a JSON 2D array, e.g. [[\"a\",\"b\"],[\"c\",\"d\"]]"}))
        return 1

    try:
        result = write_range(
            item_id=args.item_id,
            sheet=args.sheet,
            range_addr=args.range,
            values=values,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def cmd_excel_append(args):
    """Append rows to an Excel worksheet or table."""
    try:
        values = json.loads(args.values)
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"Invalid JSON for --values: {e}"}))
        return 1

    if not isinstance(values, list) or not all(isinstance(row, list) for row in values):
        print(json.dumps({"error": "--values must be a JSON 2D array, e.g. [[\"a\",\"b\"],[\"c\",\"d\"]]"}))
        return 1

    try:
        result = append_rows(
            item_id=args.item_id,
            values=values,
            sheet=args.sheet,
            table=args.table,
        )
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0
```

- [ ] **Step 3: Commit**

```bash
cd ~/.config/claude-graph && git add src/cli.py && git commit -m "feat: add excel command handler functions"
```

---

### Task 6: Register CLI subparsers for `excel` commands

**Files:**
- Modify: `src/cli.py` (add parser registration after docs parsers, before config parser ~line 1459)
- Modify: `src/cli.py:1509-1517` (add excel to subcommand dispatch)

- [ ] **Step 1: Add excel subparser group**

Insert after the docs site parser block (after line 1457) and before the config parser (line 1459):

```python
    # excel command group
    excel_parser = subparsers.add_parser("excel", help="Read and write Excel workbooks")
    excel_sub = excel_parser.add_subparsers(dest="excel_command", help="Excel subcommands")

    # excel info
    excel_info_parser = excel_sub.add_parser("info", help="List worksheets and tables")
    excel_info_parser.add_argument(
        "item_id",
        help="SharePoint/OneDrive item ID for the Excel file"
    )
    excel_info_parser.set_defaults(func=cmd_excel_info)

    # excel read
    excel_read_parser = excel_sub.add_parser("read", help="Read cell values")
    excel_read_parser.add_argument(
        "item_id",
        help="SharePoint/OneDrive item ID for the Excel file"
    )
    excel_read_parser.add_argument(
        "--sheet", "-s",
        help="Worksheet name (default: first worksheet)"
    )
    excel_read_parser.add_argument(
        "--range", "-r",
        help="Range address like A1:L10 (default: used range)"
    )
    excel_read_parser.set_defaults(func=cmd_excel_read)

    # excel write
    excel_write_parser = excel_sub.add_parser("write", help="Update cell values")
    excel_write_parser.add_argument(
        "item_id",
        help="SharePoint/OneDrive item ID for the Excel file"
    )
    excel_write_parser.add_argument(
        "--sheet", "-s",
        required=True,
        help="Worksheet name"
    )
    excel_write_parser.add_argument(
        "--range", "-r",
        required=True,
        help="Target range address like F5 or F5:G6"
    )
    excel_write_parser.add_argument(
        "--values", "-v",
        required=True,
        help='JSON 2D array of values, e.g. \'[[\"a\",\"b\"],[\"c\",\"d\"]]\''
    )
    excel_write_parser.set_defaults(func=cmd_excel_write)

    # excel append
    excel_append_parser = excel_sub.add_parser("append", help="Append rows")
    excel_append_parser.add_argument(
        "item_id",
        help="SharePoint/OneDrive item ID for the Excel file"
    )
    excel_append_parser.add_argument(
        "--sheet", "-s",
        help="Worksheet name (required unless --table is used)"
    )
    excel_append_parser.add_argument(
        "--table", "-t",
        help="Named table to append to"
    )
    excel_append_parser.add_argument(
        "--values", "-v",
        required=True,
        help='JSON 2D array of rows, e.g. \'[[\"col1\",\"col2\"]]\''
    )
    excel_append_parser.set_defaults(func=cmd_excel_append)
```

- [ ] **Step 2: Add excel to subcommand dispatch**

In the subcommand dispatch section (around line 1509-1517), after the `docs` check, add:

```python
    if args.command == "excel":
        if not args.excel_command:
            excel_parser.print_help()
            return 1
```

- [ ] **Step 3: Test CLI help renders**

Run: `cd ~/.config/claude-graph && python -m src.cli excel --help`
Expected: Shows `info`, `read`, `write`, `append` subcommands.

Run: `cd ~/.config/claude-graph && python -m src.cli excel read --help`
Expected: Shows `item_id`, `--sheet`, `--range` arguments.

- [ ] **Step 4: Commit**

```bash
cd ~/.config/claude-graph && git add src/cli.py && git commit -m "feat: register excel subcommand parsers in CLI"
```

---

### Task 7: Manual integration test against Alpha tracking spreadsheet

**Files:** None (testing only)

- [ ] **Step 1: Find the tracking spreadsheet item ID**

Run: `~/.config/claude-graph/bin/msgraph docs search "Private Preview Tracking" --type xlsx`
Note the item ID from the results. (Known ID from memory: `B0820E2C-5D97-488E-9B9A-2A7556B40848` — but confirm it works with the API.)

- [ ] **Step 2: Test `excel info`**

Run: `~/.config/claude-graph/bin/msgraph excel info <item-id>`
Expected: JSON with worksheets and tables lists.

- [ ] **Step 3: Test `excel read`**

Run: `~/.config/claude-graph/bin/msgraph excel read <item-id>`
Expected: JSON with used range values from the first worksheet.

Run with range: `~/.config/claude-graph/bin/msgraph excel read <item-id> --range A1:C3`
Expected: JSON with 3 rows, 3 columns.

- [ ] **Step 4: Test `excel write`**

Pick a safe cell to test (read first to know what's there, write the same value back):

Run: `~/.config/claude-graph/bin/msgraph excel write <item-id> --sheet Sheet1 --range <safe-cell> --values '[["<original-value>"]]'`
Expected: JSON with status "updated" and the value confirmed.

- [ ] **Step 5: Test `excel append` (if a test row is acceptable)**

Only run if the user confirms it's okay to add a test row. Otherwise skip.

Run: `~/.config/claude-graph/bin/msgraph excel append <item-id> --sheet Sheet1 --values '[["Test","","","","","",""]]'`
Expected: JSON with status "appended" and the new range address.
