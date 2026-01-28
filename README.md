# MS365-Access

A FastAPI backend providing secure API access to Microsoft 365 services (Email, Calendar, OneDrive) for local use and AI agent integration.

## Features

- **Email**: Read, send, reply, forward, search, and organize messages
- **Calendar**: View, create, update events; handle meeting invites
- **OneDrive**: Browse, upload, download, and manage files
- **Security**: Token encryption at rest, audit logging, localhost-only by default

## Prerequisites

- Python 3.10+
- An Azure AD App Registration with appropriate permissions
- Microsoft 365 account

## Quick Start

1. **Clone and configure**
   ```bash
   git clone https://github.com/your-username/ms365-access.git
   cd ms365-access
   cp .env.example .env
   # Edit .env with your Azure credentials
   ```

2. **Install dependencies**
   ```bash
   cd backend
   pip install -r requirements.txt
   ```

3. **Run the server**
   ```bash
   uvicorn app.main:app --port 8365 --reload
   ```

4. **Authenticate**
   - Visit http://localhost:8365/auth/login
   - Sign in with your Microsoft account
   - Verify at http://localhost:8365/auth/status

## Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Create new registration
3. Add redirect URI: `http://localhost:8365/auth/callback`
4. Create a client secret
5. Add API permissions:
   - Microsoft Graph: User.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Files.ReadWrite.All

## Configuration

See `.env.example` for all configuration options. Required variables:

| Variable | Description |
|----------|-------------|
| `AZURE_CLIENT_ID` | Azure AD App Registration ID |
| `AZURE_CLIENT_SECRET` | Azure AD App Client Secret |
| `AZURE_TENANT_ID` | Azure AD Tenant ID |
| `SECRET_KEY` | Encryption key for token storage (min 32 chars) |

## API Endpoints

### Auth (`/auth`)
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/auth/login` | Redirect to Azure AD OAuth |
| GET | `/auth/callback` | Handle OAuth callback |
| GET | `/auth/status` | Authentication status |
| POST | `/auth/logout` | Clear tokens |

### Mail (`/mail`)
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/mail/folders` | List mail folders |
| GET | `/mail/messages` | List messages |
| GET | `/mail/messages/{id}` | Get message details |
| POST | `/mail/messages` | Send email |
| POST | `/mail/messages/{id}/reply` | Reply to message |
| POST | `/mail/messages/{id}/forward` | Forward message |
| PATCH | `/mail/messages/{id}` | Update message |
| POST | `/mail/messages/{id}/move` | Move to folder |
| DELETE | `/mail/messages/{id}` | Delete message |
| POST | `/mail/batch/move` | Batch move |
| POST | `/mail/batch/delete` | Batch delete |
| GET | `/mail/search` | Search messages |

### Calendar (`/calendar`)
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/calendar/calendars` | List calendars |
| GET | `/calendar/events` | List events |
| GET | `/calendar/view` | Calendar view (expands recurring) |
| GET | `/calendar/events/{id}` | Get event |
| POST | `/calendar/events` | Create event |
| PATCH | `/calendar/events/{id}` | Update event |
| DELETE | `/calendar/events/{id}` | Delete event |
| POST | `/calendar/events/{id}/accept` | Accept invite |
| POST | `/calendar/events/{id}/tentative` | Tentatively accept |
| POST | `/calendar/events/{id}/decline` | Decline invite |

### Files (`/files`) - OneDrive
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/files/drives` | List drives |
| GET | `/files/drive/root` | Get root folder |
| GET | `/files/items/{id}` | Get item metadata |
| GET | `/files/items/{id}/children` | List folder contents |
| GET | `/files/items/{id}/content` | Download file |
| PUT | `/files/items/{parent_id}:/{name}:/content` | Upload file |
| DELETE | `/files/items/{id}` | Delete item |
| POST | `/files/items/{parent_id}/folder` | Create folder |
| PATCH | `/files/items/{id}` | Rename/move item |
| GET | `/files/search` | Search files |

## Examples

### Send an Email
```bash
curl -X POST http://localhost:8365/mail/messages \
  -H "Content-Type: application/json" \
  -d '{
    "subject": "Hello",
    "body": "<p>Test email</p>",
    "body_type": "HTML",
    "to_recipients": ["user@example.com"]
  }'
```

### Create Calendar Event
```bash
curl -X POST http://localhost:8365/calendar/events \
  -H "Content-Type: application/json" \
  -d '{
    "subject": "Team Meeting",
    "start_datetime": "2024-01-15T10:00:00",
    "end_datetime": "2024-01-15T11:00:00",
    "time_zone": "UTC",
    "location": "Conference Room A",
    "attendees": ["colleague@example.com"]
  }'
```

### List Files in Root
```bash
# Get root folder ID
curl http://localhost:8365/files/drive/root

# List children (use the id from above)
curl http://localhost:8365/files/items/{root_id}/children
```

## API Documentation

Interactive API docs available at:
- Swagger UI: http://localhost:8365/docs
- ReDoc: http://localhost:8365/redoc

## Docker

```bash
# Build and run
make up

# Stop
make down
```

## Security Notes

- **Single-user design**: This is intended for local/single-user use, not multi-tenant deployment
- **Token encryption**: OAuth tokens are encrypted at rest using Fernet (AES-128-CBC)
- **Localhost binding**: By default, binds to 127.0.0.1 only
- **Audit logging**: Sensitive operations are logged to `data/audit.log`

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup and guidelines.

## License

[MIT](LICENSE)
