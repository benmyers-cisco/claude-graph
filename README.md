# claude-graph

Microsoft Graph CLI plugin for Claude Code. Gives Claude access to your Outlook calendar, email, SharePoint docs, Excel spreadsheets, conference rooms, and org chart — all through a single `msgraph` command.

## Commands

### Calendar & Meetings

| Command | Description |
|---------|-------------|
| `msgraph auth` | Authenticate via Graph Explorer (Playwright-automated SSO) |
| `msgraph calendar [days]` | Show calendar events |
| `msgraph find-meeting <names...>` | Find available meeting times |
| `msgraph book-rooms` | Find meetings needing rooms and show availability |
| `msgraph book-room <event-id> <room-email>` | Book a room for an event |
| `msgraph room-availability <start> [end]` | Check room availability |
| `msgraph stats [days]` | Compute meeting time statistics |

### Email

| Command | Description |
|---------|-------------|
| `msgraph email search <query>` | Search emails (KQL: `from:`, `to:`, `subject:`, `isRead:false`, etc.) |
| `msgraph email read <message-id>` | Read full email content (HTML body, attachments) |

### SharePoint & OneDrive

| Command | Description |
|---------|-------------|
| `msgraph sharepoint-search <query>` | Search SharePoint/OneDrive for files |
| `msgraph sharepoint-recent [days]` | Recently modified documents |
| `msgraph sharepoint-site <site-name>` | List documents in a SharePoint site |
| `msgraph sharepoint-read <item-id>` | Read a SharePoint file (Excel, Word, PowerPoint, PDF, text) |
| `msgraph sharepoint-download <item-id>` | Download a SharePoint file to local disk |
| `msgraph sharepoint-upload <item-id>` | Upload a file to replace an existing SharePoint file |
| `msgraph sharepoint-create <path>` | Create a new file in OneDrive/SharePoint |

### Excel (via Graph API)

| Command | Description |
|---------|-------------|
| `msgraph excel info <item-id>` | List worksheets and tables in a workbook |
| `msgraph excel read <item-id>` | Read cell values from a worksheet |
| `msgraph excel write <item-id>` | Update cell values in a range |
| `msgraph excel append <item-id>` | Append rows to a worksheet or table |

### People & Config

| Command | Description |
|---------|-------------|
| `msgraph resolve-person <query>` | Resolve name to email address |
| `msgraph org <query>` | Look up manager and direct reports |
| `msgraph config` | Get or set user preferences |

## Setup

### 1. Clone and install

```bash
git clone https://github.com/benmyers-cisco/claude-graph.git ~/.config/claude-graph
cd ~/.config/claude-graph
python3.12 -m venv venv
./venv/bin/pip install -r requirements.txt
playwright install chromium
```

### 2. Configure

```bash
./bin/msgraph config --email you@company.com --timezone America/Detroit
```

Optional room booking preferences:

```bash
./bin/msgraph config --building ARB02 --floors 1,2
```

### 3. Authenticate

```bash
./bin/msgraph auth
```

This opens a browser, completes SSO via Graph Explorer, and stores the token locally. Tokens expire after ~1 hour; the CLI auto-refreshes when needed.

### 4. Add to Claude Code permissions

Add to your `.claude/settings.json`:

```json
{
  "permissions": {
    "allow": [
      "Bash(~/.config/claude-graph/bin/msgraph *)"
    ]
  }
}
```

## Excel Commands

The `excel` subcommands read and write Excel files stored in SharePoint or OneDrive using the [Microsoft Graph Excel API](https://learn.microsoft.com/en-us/graph/api/resources/excel) — no file downloads or openpyxl needed.

### File discovery

Find a spreadsheet's item ID using SharePoint search:

```bash
msgraph sharepoint-search "My Tracking Sheet"
```

### Reading data

```bash
# List worksheets and tables
msgraph excel info <item-id>

# Read entire used range (first worksheet)
msgraph excel read <item-id>

# Read specific sheet and range
msgraph excel read <item-id> --sheet "Sheet1" --range "A1:D10"
```

### Writing data

```bash
# Update a single cell
msgraph excel write <item-id> --sheet "Sheet1" --range "F5" --values '[["new value"]]'

# Update a range
msgraph excel write <item-id> --sheet "Sheet1" --range "A2:C2" --values '[["a","b","c"]]'
```

### Appending rows

```bash
# Append to end of used range
msgraph excel append <item-id> --sheet "Sheet1" --values '[["col1","col2","col3"]]'

# Append to a named table
msgraph excel append <item-id> --table "Table1" --values '[["col1","col2"]]'
```

### SharePoint site files

Files on SharePoint sites (not your personal OneDrive) require a `--drive` flag:

```bash
msgraph excel --drive <drive-id> read <item-id> --sheet "Sheet1"
```

Get the drive ID by browsing the site with `msgraph docs site <site-name>` and then resolving the file path via the Graph API.

## Authentication

The CLI uses a delegated access token from Microsoft Graph Explorer. The `auth` command automates the browser flow using Playwright:

1. Opens Graph Explorer in Chromium
2. Completes your organization's SSO (Duo MFA supported)
3. Extracts the access token
4. Stores it at `~/.config/claude-graph/token`

Tokens are checked for expiration before each command and auto-refreshed when needed.

## Output

All commands output JSON, making them easy for Claude to parse and act on. Errors are returned as `{"error": "..."}` with a non-zero exit code.

## License

Internal use.
