# Excel Commands for msgraph CLI

**Date**: 2026-04-23
**Status**: Approved

## Goal

Add Excel read/write commands to the existing msgraph CLI so Claude can programmatically read and update SharePoint/OneDrive spreadsheets (e.g., the Alpha program tracking doc).

## Architecture

- New module: `src/excel_tools.py`
- Pattern: raw `requests` calls with bearer token (same as `docs_tools.py`)
- No msgraph Python SDK for Excel — its Excel bindings are unreliable
- CLI commands registered in `cli.py` as `msgraph excel <subcommand>`
- No workbook sessions — each command is a single atomic REST call

## File Identification

Two modes, used across all commands:

- `--item-id ID` — direct SharePoint/OneDrive item ID for known files
- Discovery via existing `msgraph docs search "query" --type xlsx` — user grabs item ID from results

Item ID is a required positional argument for all `excel` subcommands.

## Commands

### `msgraph excel info <item-id>`

Lists worksheets and named tables in a workbook.

**Graph API calls:**
- `GET /me/drive/items/{id}/workbook/worksheets`
- `GET /me/drive/items/{id}/workbook/tables` (for named tables)

**Output:**
```json
{
  "item_id": "...",
  "worksheets": [{"name": "Sheet1", "id": "..."}],
  "tables": [{"name": "Table1", "id": "...", "worksheet": "Sheet1", "rows": 42}]
}
```

### `msgraph excel read <item-id> [--sheet NAME] [--range RANGE]`

Reads cell values from a worksheet.

**Arguments:**
- `--sheet` — worksheet name (default: first worksheet)
- `--range` — Excel range address like `A1:L10` (default: used range)

**Graph API calls:**
- With range: `GET /me/drive/items/{id}/workbook/worksheets/{name}/range(address='{range}')`
- Without range: `GET /me/drive/items/{id}/workbook/worksheets/{name}/usedRange`

**Output:**
```json
{
  "worksheet": "Sheet1",
  "range": "A1:L42",
  "rows": 42,
  "columns": 12,
  "values": [["Header1", "Header2", ...], ["row1col1", "row1col2", ...]]
}
```

### `msgraph excel write <item-id> --sheet NAME --range RANGE --values JSON`

Updates specific cells.

**Arguments:**
- `--sheet` — worksheet name (required)
- `--range` — target range address like `F5` or `F5:G6` (required)
- `--values` — JSON 2D array matching range dimensions (required)

**Graph API call:**
- `PATCH /me/drive/items/{id}/workbook/worksheets/{name}/range(address='{range}')` with body `{"values": [[...]]}`

**Output:** Returns the updated range values as confirmation.

```json
{
  "worksheet": "Sheet1",
  "range": "F5",
  "status": "updated",
  "values": [["Rescheduled"]]
}
```

### `msgraph excel append <item-id> --sheet NAME --values JSON [--table NAME]`

Appends rows after the last used row, or to a named table.

**Arguments:**
- `--sheet` — worksheet name (required unless `--table` used)
- `--values` — JSON 2D array of rows to append (required)
- `--table` — named table to append to (optional; uses table API instead of range)

**Graph API calls:**
- Sheet mode: `GET usedRange` to find last row, then `PATCH` the next row range
- Table mode: `POST /me/drive/items/{id}/workbook/tables/{name}/rows` with body `{"values": [[...]]}`

**Output:**
```json
{
  "worksheet": "Sheet1",
  "status": "appended",
  "range": "A43:L43",
  "rows_added": 1,
  "values": [[...]]
}
```

## Implementation Details

### `excel_tools.py` Functions

```
resolve_item_id(item_id: str) -> str
    Validates item ID format, returns as-is. Placeholder for future search integration.

get_workbook_info(item_id: str) -> dict
    Lists worksheets and tables.

read_range(item_id: str, sheet: str | None, range_addr: str | None) -> dict
    Reads values from a range or used range.

write_range(item_id: str, sheet: str, range_addr: str, values: list[list]) -> dict
    Updates cell values in a range.

append_rows(item_id: str, sheet: str | None, table: str | None, values: list[list]) -> dict
    Appends rows to end of used range or to a named table.
```

### HTTP Helpers

Reuse `_headers()` and `_requests_get_with_retry()` patterns from `docs_tools.py`. Add `_requests_patch_with_retry()` for write operations with same retry logic (429/503/504).

### Graph API Base URL

`https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/workbook/...`

All calls go through the user's OneDrive, which includes SharePoint files the user has accessed.

## Error Handling

- **Auth errors**: Handled by existing `cli.py` auto-retry wrapper
- **Bad range address**: Graph API returns descriptive error — pass through as `{"error": "..."}`
- **`.xls` files**: Not supported by Graph Excel API (only `.xlsx`/`.xlsm`). Detect by checking the file name if possible, otherwise let the API error propagate.
- **Retries**: GET and PATCH retry on 429/503/504 (same as docs_tools.py)
- **Large used ranges**: No pagination — Graph returns full range. User can narrow with `--range`.

## CLI Registration

Add to `cli.py`:
- `excel` subparser with sub-subparsers for `info`, `read`, `write`, `append`
- Import `excel_tools` functions
- `cmd_excel_info`, `cmd_excel_read`, `cmd_excel_write`, `cmd_excel_append` handler functions
- Wire into the existing token check and auto-retry flow

## Permissions

Uses `Files.ReadWrite` delegated scope — already available through the Graph Explorer token.

## Testing

Manual testing against the Alpha tracking spreadsheet:
1. `msgraph docs search "Private Preview Tracking" --type xlsx` — find the item ID
2. `msgraph excel info <id>` — see worksheets and tables
3. `msgraph excel read <id>` — read full contents
4. `msgraph excel write <id> --sheet Sheet1 --range F2 --values '[["Rescheduled"]]'` — update a cell
5. `msgraph excel append <id> --sheet Sheet1 --values '[["Test Org","Contact",...]]'` — add a row
