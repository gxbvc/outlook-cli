# outlook-cli

CLI tool for interacting with Microsoft Outlook email via the Microsoft Graph API. Supports listing, reading, sending, replying, drafts, and folder management.

## Prerequisites

- Node.js 18+
- An Azure AD app registration with Microsoft Graph permissions (Mail.Read, Mail.ReadWrite, Mail.Send, User.Read)

## Setup

1. Clone and install:
   ```bash
   cd ~/tools/outlook-cli
   npm install
   npm run build
   npm link
   ```

2. Configure credentials in `.env` (see `.env.example`):
   ```
   OUTLOOK_CLIENT_ID=your-app-client-id
   OUTLOOK_CLIENT_SECRET=your-client-secret
   OUTLOOK_TENANT_ID=your-tenant-id
   OUTLOOK_REDIRECT_URI=http://localhost:3000/callback
   ```

3. Authenticate:
   ```bash
   outlook-cli auth
   ```
   This opens a browser window for Microsoft OAuth2 login. Tokens are saved to `tokens/tokens.json` and auto-refreshed.

## Commands

### Authentication
- `outlook-cli auth` — Start OAuth2 flow
- `outlook-cli me` — Get current user profile

### Emails
- `outlook-cli emails list` — List inbox emails
- `outlook-cli emails get --id ID` — Get full email
- `outlook-cli emails send --to ADDR --subject SUBJ --text BODY` — Send email
- `outlook-cli emails reply --id ID --text BODY` — Reply to email
- `outlook-cli emails move --id ID --folder FOLDER_ID` — Move email
- `outlook-cli emails delete --id ID` — Delete email
- `outlook-cli emails read --id ID` — Mark as read
- `outlook-cli emails unread --id ID` — Mark as unread
- `outlook-cli emails attachments --id ID` — List attachments (id, name, type, size)
- `outlook-cli emails download --id ID [--attachment-id ATT_ID] [--out PATH]` — Download attachment(s) to disk

### Drafts
- `outlook-cli drafts list` — List drafts
- `outlook-cli drafts create --to ADDR --subject SUBJ --text BODY` — Create draft
- `outlook-cli drafts reply --id ID` — Create reply draft
- `outlook-cli drafts update --id ID --subject SUBJ` — Update draft
- `outlook-cli drafts send --id ID` — Send draft
- `outlook-cli drafts delete --id ID` — Delete draft

### Folders
- `outlook-cli folders list` — List mail folders
- `outlook-cli folders create --name NAME` — Create folder

## Tips

- **Always use `--html` instead of `--text`** when sending, replying, or creating drafts. Plain text (`--text`) renders without line breaks in Outlook — newlines are ignored. Use `<p>` tags for paragraphs and `<br>` for line breaks. Example:
  ```bash
  outlook-cli drafts create --to "someone@example.com" --subject "Hello" \
    --html '<p>Hi there,</p><p>Here is my message.</p><p>Thanks!</p>'
  ```

## Output Format

All commands output JSON to stdout:
```json
{"ok": true, "data": { ... }}
{"ok": false, "error": "message", "code": "ERROR_CODE"}
```

## How It Works

Uses the Microsoft Graph REST API v1.0 with OAuth2 delegated permissions. The app is registered in Azure AD (Microsoft Entra ID) with multi-tenant + personal account support. Tokens are stored locally and auto-refreshed using the refresh token.
