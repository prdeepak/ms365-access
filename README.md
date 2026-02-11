# MS365-Access

A FastAPI backend providing secure API access to Microsoft 365 services (Email, Calendar, OneDrive, SharePoint) for local use and AI agent integration.

Base URL: `http://localhost:8365`

## Features

- **Email**: Read, send, reply, forward, search, and organize messages
- **Calendar**: View, create, update events; handle meeting invites
- **OneDrive**: Browse, upload, download, and manage files
- **SharePoint**: Resolve sites, browse document libraries, search, download files
- **Security**: Token encryption at rest, audit logging, localhost-only by default

## Quick Start

1. **Configure**: `cp .env.example .env` and fill in Azure credentials
2. **Run**: `make up` (Docker) or `uvicorn app.main:app --port 8365 --reload`
3. **Authenticate**: Visit http://localhost:8365/auth/login, sign in with Microsoft
4. **Verify**: http://localhost:8365/auth/status

## Configuration

See `.env.example` for all options. Required:

| Variable | Description |
|----------|-------------|
| `AZURE_CLIENT_ID` | Azure AD App Registration ID |
| `AZURE_CLIENT_SECRET` | Azure AD App Client Secret |
| `AZURE_TENANT_ID` | Azure AD Tenant ID |
| `SECRET_KEY` | Encryption key for token storage (min 32 chars) |

Azure AD permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Files.ReadWrite.All`

---

## API Reference

All endpoints require authentication (via `/auth/login`) except the auth endpoints themselves.

Optional query params shown with `?` suffix. Defaults shown in parentheses.

### Auth (`/auth`)

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/auth/login` | Redirect to Azure AD OAuth login |
| GET | `/auth/callback?code=...` | OAuth callback (automatic) |
| GET | `/auth/status` | Auth status: authenticated, email, token expiry |
| POST | `/auth/logout` | Clear tokens and log out |

### Mail (`/mail`)

Well-known folder names: `inbox`, `drafts`, `sentitems`, `deleteditems`, `junkemail`, `archive`, `outbox`

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/mail/folders` | List all mail folders |
| GET | `/mail/folders/resolve/{name}` | Resolve well-known folder name to folder object with ID |
| GET | `/mail/messages` | List messages from a folder |
| GET | `/mail/messages/{id}` | Get full message details |
| POST | `/mail/messages` | Send a new email |
| POST | `/mail/messages/{id}/send` | Send an existing draft |
| POST | `/mail/messages/{id}/draftReply` | Create a draft reply (does not send) |
| POST | `/mail/messages/{id}/reply` | Reply to a message (sends immediately) |
| POST | `/mail/messages/{id}/forward` | Forward a message |
| PATCH | `/mail/messages/{id}` | Update message (read status, flags, categories, body) |
| POST | `/mail/messages/{id}/move` | Move message to a folder |
| DELETE | `/mail/messages/{id}` | Delete a message |
| POST | `/mail/batch/move` | Batch move messages (background job) |
| POST | `/mail/batch/delete` | Batch delete messages (background job) |
| GET | `/mail/search` | Search messages |

**GET `/mail/messages` query params:**
- `folder?` — well-known folder name (e.g. `inbox`, `drafts`)
- `folder_id?` — folder ID (takes precedence over `folder`)
- `top?` (25) — results per page, 1-100
- `skip?` (0) — pagination offset
- `search?` — search query
- `filter?` — OData filter expression
- `order_by?` (`receivedDateTime desc`) — sort order

**GET `/mail/search` query params:**
- `q` — search query (required)
- `top?` (25) — results per page, 1-100
- `skip?` (0) — pagination offset

**POST `/mail/messages/{id}/draftReply` query params:**
- `reply_all?` (false) — reply to all recipients

**POST `/mail/messages/{id}/move` query params:**
- `verify?` (true) — verify the move succeeded

**POST `/mail/messages` body:**
```json
{
  "subject": "Hello",
  "body": "<p>HTML body</p>",
  "body_type": "HTML",
  "to_recipients": ["user@example.com"],
  "cc_recipients": [],
  "bcc_recipients": [],
  "importance": "normal",
  "save_to_sent_items": true
}
```

**POST `/mail/messages/{id}/reply` body:**
```json
{ "comment": "Thanks!", "reply_all": false }
```

**POST `/mail/messages/{id}/forward` body:**
```json
{ "comment": "FYI", "to_recipients": ["user@example.com"] }
```

**PATCH `/mail/messages/{id}` body** (all fields optional):
```json
{ "is_read": true, "flag_status": "flagged", "categories": ["Blue"], "body": "...", "body_type": "HTML" }
```

**POST `/mail/messages/{id}/move` body:**
```json
{ "destination_folder_id": "AAMkAG..." }
```

**POST `/mail/batch/move` body:**
```json
{ "message_ids": ["id1", "id2"], "destination_folder_id": "AAMkAG..." }
```

**POST `/mail/batch/delete` body:**
```json
{ "message_ids": ["id1", "id2"] }
```

### Calendar (`/calendar`)

Events are returned in local timezone (configured via `LOCAL_TIMEZONE` env var).

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/calendar/calendars` | List all calendars |
| GET | `/calendar/events` | List events with pagination |
| GET | `/calendar/view` | Calendar view for date range (expands recurring events) |
| GET | `/calendar/events/{id}` | Get a specific event |
| POST | `/calendar/events` | Create an event |
| PATCH | `/calendar/events/{id}` | Update an event |
| DELETE | `/calendar/events/{id}` | Delete an event |
| POST | `/calendar/events/{id}/accept` | Accept meeting invite |
| POST | `/calendar/events/{id}/tentative` | Tentatively accept |
| POST | `/calendar/events/{id}/decline` | Decline meeting invite |

**GET `/calendar/events` query params:**
- `calendar_id?` — specific calendar (default: primary)
- `top?` (25) — results per page, 1-100
- `skip?` (0) — pagination offset
- `order_by?` (`start/dateTime`) — sort order
- `filter?` — OData filter expression

**GET `/calendar/view` query params:**
- `start_datetime` — range start, ISO 8601 (required)
- `end_datetime` — range end, ISO 8601 (required)
- `calendar_id?` — specific calendar
- `top?` (100) — max results, 1-500

**POST `/calendar/events` body:**
```json
{
  "subject": "Team Meeting",
  "start_datetime": "2024-01-15T10:00:00",
  "end_datetime": "2024-01-15T11:00:00",
  "time_zone": "UTC",
  "location": "Conference Room A",
  "attendees": ["colleague@example.com"],
  "body": "<p>Agenda...</p>",
  "body_type": "HTML",
  "is_all_day": false,
  "is_online_meeting": false,
  "reminder_minutes": 15,
  "show_as": "busy",
  "importance": "normal",
  "recurrence": null
}
```

**PATCH `/calendar/events/{id}` body:** Same fields as create, all optional.

**POST `/calendar/events/{id}/accept|tentative|decline` body:**
```json
{ "comment": "Thanks!", "send_response": true }
```

### Files (`/files`) — OneDrive

All file endpoints accept an optional `drive_id` query param. Without it, uses the user's default OneDrive.

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/files/drives` | List accessible drives |
| GET | `/files/drive/root` | Get root folder metadata |
| GET | `/files/items/{id}` | Get item metadata |
| GET | `/files/items/{id}/children` | List folder contents |
| GET | `/files/items/{id}/content` | Download file |
| PUT | `/files/items/{parent_id}:/{filename}:/content` | Upload file (multipart) |
| DELETE | `/files/items/{id}` | Delete item |
| POST | `/files/items/{parent_id}/folder` | Create folder |
| PATCH | `/files/items/{id}` | Rename/move item |
| GET | `/files/search` | Search files |

**GET `/files/items/{id}/children` query params:**
- `drive_id?` — specific drive
- `top?` (100) — results per page, 1-200
- `skip?` (0) — pagination offset
- `order_by?` (`name`) — sort order

**GET `/files/search` query params:**
- `q` — search query (required)
- `drive_id?` — specific drive
- `top?` (25) — results per page, 1-100

**POST `/files/items/{parent_id}/folder` body:**
```json
{ "name": "New Folder" }
```

**PATCH `/files/items/{id}` body:**
```json
{ "name": "new-name.txt", "parent_id": "target-folder-id" }
```

### SharePoint (`/sharepoint`)

SharePoint endpoints require a `drive_id` (document library ID) for file operations. Get it by resolving a site first, then listing its drives.

**Typical workflow:**
1. Resolve site → get `site_id`
2. List drives for that site → get `drive_id`
3. Browse/search/download using `drive_id`

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/sharepoint/sites/{host_path}` | Resolve site by hostname/path |
| GET | `/sharepoint/drives` | List document libraries for a site |
| GET | `/sharepoint/items/{id}/children` | List folder contents |
| GET | `/sharepoint/items/{id}` | Get item metadata |
| GET | `/sharepoint/items/{id}/content` | Download file (optional format conversion) |
| GET | `/sharepoint/search` | Search within a drive |
| GET | `/sharepoint/resolve` | Resolve a SharePoint sharing URL to metadata |

**GET `/sharepoint/sites/{host_path}` — site resolution:**
Path is the full hostname + site path, e.g.:
```
/sharepoint/sites/contoso.sharepoint.com/sites/Finance
```

**GET `/sharepoint/drives` query params:**
- `site_id` — site ID from resolve (required)

**GET `/sharepoint/items/{id}/children` query params:**
- `drive_id` — document library ID (required)
- `top?` (100) — results per page, 1-200
- `order_by?` (`name`) — sort order
- Use `item_id=root` to list the drive root

**GET `/sharepoint/items/{id}` query params:**
- `drive_id` — document library ID (required)

**GET `/sharepoint/items/{id}/content` query params:**
- `drive_id` — document library ID (required)
- `format?` — convert to format on download (e.g. `pdf`)

**GET `/sharepoint/search` query params:**
- `drive_id` — document library ID (required)
- `q` — search query (required)
- `top?` (25) — results per page, 1-100

**GET `/sharepoint/resolve` query params:**
- `url` — full SharePoint URL (required). Works with sharing links and direct URLs.

---

## Examples

### List inbox messages
```bash
curl 'http://localhost:8365/mail/messages?folder=inbox&top=10'
```

### Send an email
```bash
curl -X POST http://localhost:8365/mail/messages \
  -H "Content-Type: application/json" \
  -d '{"subject":"Hello","body":"<p>Test</p>","body_type":"HTML","to_recipients":["user@example.com"]}'
```

### Get calendar view for a week
```bash
curl 'http://localhost:8365/calendar/view?start_datetime=2026-02-10T00:00:00&end_datetime=2026-02-17T00:00:00'
```

### Browse OneDrive
```bash
# Get root folder ID
curl http://localhost:8365/files/drive/root

# List children
curl 'http://localhost:8365/files/items/{root_id}/children'
```

### Browse SharePoint site
```bash
# 1. Resolve site
curl http://localhost:8365/sharepoint/sites/contoso.sharepoint.com/sites/Finance

# 2. List drives (use site id from step 1)
curl 'http://localhost:8365/sharepoint/drives?site_id=contoso.sharepoint.com,guid1,guid2'

# 3. Search for a file (use drive id from step 2)
curl 'http://localhost:8365/sharepoint/search?drive_id=b!abc123&q=agenda'

# 4. Download a file (use item id from search results)
curl 'http://localhost:8365/sharepoint/items/ITEM_ID/content?drive_id=b!abc123' -o file.docx

# 4b. Download as PDF
curl 'http://localhost:8365/sharepoint/items/ITEM_ID/content?drive_id=b!abc123&format=pdf' -o file.pdf
```

## Interactive API Docs

- Swagger UI: http://localhost:8365/docs
- ReDoc: http://localhost:8365/redoc

## Docker

```bash
make up    # Build and run
make down  # Stop
```

## Security Notes

- **Single-user design**: Intended for local/single-user use, not multi-tenant
- **Token encryption**: OAuth tokens encrypted at rest using Fernet (AES-128-CBC)
- **Localhost binding**: Binds to 127.0.0.1 only by default
- **Audit logging**: Sensitive operations logged to `data/audit.log`

<!-- GEN:API_START -->

## API Reference

> Auto-generated from OpenAPI spec. Do not edit manually.
> Regenerate with: `make gen-client`

### Health

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Health |

### Mail

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/mail/batch/delete` | Batch Delete Messages |
| POST | `/mail/batch/move` | Batch Move Messages |
| GET | `/mail/folders` | List Folders |
| GET | `/mail/folders/resolve/{name}` | Resolve Folder Name |
| GET | `/mail/messages?folder=...&folder_id=...&top=...&...` | List Messages |
| POST | `/mail/messages` | Send Mail |
| DELETE | `/mail/messages/{message_id}` | Delete Message |
| GET | `/mail/messages/{message_id}` | Get Message |
| PATCH | `/mail/messages/{message_id}` | Update Message |
| POST | `/mail/messages/{message_id}/draftReply?reply_all=...` | Create Reply Draft |
| POST | `/mail/messages/{message_id}/forward` | Forward Message |
| POST | `/mail/messages/{message_id}/move?verify=...` | Move Message |
| POST | `/mail/messages/{message_id}/reply` | Reply To Message |
| POST | `/mail/messages/{message_id}/send` | Send Draft |
| GET | `/mail/search?q=...&top=...&skip=...` | Search Messages |

### Calendar

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/calendar/calendars` | List Calendars |
| GET | `/calendar/events?calendar_id=...&top=...&skip=...&...` | List Events |
| POST | `/calendar/events?calendar_id=...` | Create Event |
| DELETE | `/calendar/events/{event_id}` | Delete Event |
| GET | `/calendar/events/{event_id}` | Get Event |
| PATCH | `/calendar/events/{event_id}` | Update Event |
| POST | `/calendar/events/{event_id}/accept` | Accept Event |
| POST | `/calendar/events/{event_id}/decline` | Decline Event |
| POST | `/calendar/events/{event_id}/tentative` | Tentatively Accept Event |
| GET | `/calendar/view?start_datetime=...&end_datetime=...&calendar_id=...&...` | Get Calendar View |

### Files

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/files/drive/root?drive_id=...` | Get Drive Root |
| GET | `/files/drives` | List Drives |
| DELETE | `/files/items/{item_id}?drive_id=...` | Delete Item |
| GET | `/files/items/{item_id}?drive_id=...` | Get Item |
| PATCH | `/files/items/{item_id}?drive_id=...` | Update Item |
| GET | `/files/items/{item_id}/children?drive_id=...&top=...&skip=...&...` | List Children |
| GET | `/files/items/{item_id}/content?drive_id=...` | Download Content |
| POST | `/files/items/{parent_id}/folder?drive_id=...` | Create Folder |
| PUT | `/files/items/{parent_id}:/{filename}:/content?drive_id=...` | Upload Content |
| GET | `/files/search?q=...&drive_id=...&top=...` | Search Files |

### Sharepoint

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/sharepoint/drives?site_id=...` | List Drives |
| GET | `/sharepoint/items/{item_id}?drive_id=...` | Get Item |
| GET | `/sharepoint/items/{item_id}/children?drive_id=...&top=...&order_by=...` | List Children |
| GET | `/sharepoint/items/{item_id}/content?drive_id=...&format=...` | Download Content |
| GET | `/sharepoint/resolve?url=...` | Resolve Url |
| GET | `/sharepoint/search?q=...&drive_id=...&top=...` | Search |
| GET | `/sharepoint/sites/{host_path}` | Resolve Site |

### 

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Root |

### Auth

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/auth/logout` | Logout |
| GET | `/auth/status` | Auth Status |

<!-- GEN:API_END -->
