# outlook-cli

CLI tool for interacting with Microsoft Outlook email via the Microsoft Graph API. Supports listing, reading, sending, replying, drafts, and folder management.

## Prerequisites

- Node.js 18+
- An Azure AD app registration with Microsoft Graph permissions (Mail.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, User.Read, OnlineMeetings.Read, OnlineMeetingTranscript.Read.All)

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
   OUTLOOK_DEFAULT_TZ=America/Chicago
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
- `outlook-cli emails send --to ADDR --subject SUBJ --text BODY [--attach PATH ...]` — Send email
- `outlook-cli emails reply --id ID --text BODY [--reply-all] [--to ADDR] [--cc ADDR] [--bcc ADDR] [--attach PATH ...]` — Reply to email; `--to/--cc/--bcc` (comma-separated or repeated) add people and keep the reply in the thread
- `outlook-cli emails forward --id ID --to ADDR [--html NOTE] [--cc ...] [--bcc ...] [--attach PATH ...]` — Forward email; `--to/--cc/--bcc` are comma-separated or repeated (keeps the original's attachments)
- `outlook-cli emails move --id ID --folder FOLDER_ID` — Move email
- `outlook-cli emails delete --id ID` — Delete email
- `outlook-cli emails read --id ID` — Mark as read
- `outlook-cli emails unread --id ID` — Mark as unread
- `outlook-cli emails attachments --id ID` — List attachments (id, name, type, size)
- `outlook-cli emails download --id ID [--attachment-id ATT_ID] [--out PATH]` — Download attachment(s) to disk

### Drafts
- `outlook-cli drafts list` — List drafts
- `outlook-cli drafts create --to ADDR --subject SUBJ --text BODY [--cc ...] [--bcc ...] [--attach PATH ...]` — Create draft
- `outlook-cli drafts reply --id ID [--attach PATH ...]` — Create reply draft
- `outlook-cli drafts forward --id ID [--to ADDR] [--html NOTE] [--cc ...] [--bcc ...] [--attach PATH ...]` — Create forward draft
- `outlook-cli drafts update --id ID [--subject SUBJ] [--cc ...] [--bcc ...] [--attach PATH ...]` — Update draft
- `outlook-cli drafts attach --id DRAFT_ID --file PATH [--file PATH ...]` — Attach file(s) to a draft
- `outlook-cli drafts detach --id DRAFT_ID --attachment-id ATT_ID` — Remove an attachment from a draft
- `outlook-cli drafts send --id ID` — Send draft
- `outlook-cli drafts send --id ID --at ISO_TIME` — Schedule the send (Exchange deferred send; `drafts delete` cancels it before then)
- `outlook-cli drafts delete --id ID` — Delete draft

### Folders
- `outlook-cli folders list` — List mail folders
- `outlook-cli folders create --name NAME` — Create folder

### Calendar
- `outlook-cli calendar list [--today] [--start ISO] [--end ISO] [--limit N] [--search Q] [--tz TZ]` — List events
- `outlook-cli calendar get --id EVENT_ID` — Get a single event
- `outlook-cli calendar create --subject SUBJ --start ISO --end ISO [--tz TZ] [--location LOC] [--html HTML | --body TEXT] [--attendees "a@b.com,c@d.com"] [--online] [--all-day]` — Create event
- `outlook-cli calendar update --id ID [--subject ...] [--start ...] [--end ...] [--location ...] [--html ...] [--attendees ...]` — Update event
- `outlook-cli calendar delete --id ID` — Delete event (silent — attendees not notified)

Timezone defaults to `OUTLOOK_DEFAULT_TZ` (`America/Chicago` if unset). Override per-call with `--tz`. `--today` uses start/end-of-day in the chosen timezone and expands recurrences via `/me/calendarView`.

### Files (OneDrive and SharePoint)
- `outlook-cli files search QUERY [--drive me|ricky]` — Search files and folders in a drive
- `outlook-cli files ls PATH [--drive me|ricky]` — List a folder; paths are relative to the drive root (e.g. `10_Products`)
- `outlook-cli files info PATH [--drive me|ricky]` — Show one item's metadata (name, id, size, hash); never file contents
- `outlook-cli files get PATH [--drive me|ricky] [--out DIR]` — Download a file. Default cache is `~/tmp/onedrive-cache` (dir mode 700, files mode 600); saves under `<cache>/<drive>/<path>`. Streams to disk (safe for large files) and prints the local path, size and sha256.
- `outlook-cli files put LOCAL PATH [--drive me|ricky] [--force] [--mkdir]` — Upload a local file. Never deletes or moves anything. Refuses to overwrite (`ALREADY_EXISTS`) without `--force`; `--force` replaces a file only, never a folder. Missing parent folders need `--mkdir`. Files over 4 MiB upload through a Graph upload session in 5 MiB chunks (never held in memory). Prints size, sha256, remote id.
- `outlook-cli files sync LOCAL_DIR PATH [--drive me|ricky] --dry-run|--apply` — Mirror a local folder into the drive: uploads only new or changed files (size + quickXorHash compared against Graph's hashes; sha1 fallback for personal OneDrive, computed locally by Microsoft's published algorithm). A changed remote file is replaced only when the remote copy is older than the local one; a remote edit newer than local is a conflict, never replaced. The last-synced content markers (quickXorHash, cTag) per file are kept in `~/.config/outlook-cli/sync-state/` (mode 600): a remote whose content changed since the last sync is a conflict even when it looks older (both-sides-edited protection; eTag is never compared, SharePoint bumps it on metadata-only changes). Uploads of changed files send `If-Match` from the plan; a 412 is a conflict. Case-only name collisions and symlinks are reported and skipped; one bad local file name does not stop the rest. Never deletes remote files, never replaces a remote folder. `--dry-run` prints the plan and counts only. Creates missing remote folders on `--apply`. Paths accept one leading or trailing `/`; `/` is the drive root.

Drives: `me` is your OneDrive (`/me/drive`). `ricky` is the SharePoint document library "Ricky Bureau - GXB"; its drive id is resolved once via a Graph site search (exact name match only) and cached in `tokens/drive-ids.json` as `{ driveId, siteId }`; every run re-verifies that the site still owns the cached drive before using it.

Read only. There are no delete, move, or write commands, and nothing in OneDrive is ever modified.

**Setup (needs consent):** `files get/search/ls/info` require the delegated Graph scopes `Files.Read.All` and `Sites.Read.All` (read only); `files put` and `files sync` additionally need `Files.ReadWrite.All`. They must be added to the Azure app registration and to `SCOPES` in `src/config.ts`, then run `outlook-cli auth` once. Until then, `files` commands return `CONSENT_REQUIRED` (403).

## Tips

- **Always use `--html` instead of `--text`** when sending, replying, or creating drafts. Plain text (`--text`) renders without line breaks in Outlook — newlines are ignored. Use `<p>` tags for paragraphs and `<br>` for line breaks. Example:
  ```bash
  outlook-cli drafts create --to "someone@example.com" --subject "Hello" \
    --html '<p>Hi there,</p><p>Here is my message.</p><p>Thanks!</p>'
  ```

## Attachments

Repeat `--attach PATH` to attach files. Files ≤ 3 MiB are inlined into the request; larger files are transparently uploaded via Graph's `createUploadSession` flow in 5 MiB chunks. Works for `emails send`, `emails reply`, `emails forward`, `drafts create`, `drafts reply`, `drafts forward`, and `drafts update`. You can also manage attachments on an existing draft directly:

```bash
outlook-cli emails send --to a@b.com --subject Hi --html "<p>Body</p>" \
  --attach ./report.pdf --attach ./chart.png

outlook-cli emails reply --id MSG_ID --html "<p>see attached</p>" --attach ./notes.md

outlook-cli drafts create --to a@b.com --subject Hi --attach ./big.zip
outlook-cli drafts attach --id DRAFT_ID --file ./extra.pdf
outlook-cli drafts detach --id DRAFT_ID --attachment-id ATT_ID
```

## Output Format

All commands output JSON to stdout:
```json
{"ok": true, "data": { ... }}
{"ok": false, "error": "message", "code": "ERROR_CODE"}
```

## How It Works

Uses the Microsoft Graph REST API v1.0 with OAuth2 delegated permissions. The app is registered in Azure AD (Microsoft Entra ID) with multi-tenant + personal account support. Tokens are stored locally and auto-refreshed using the refresh token.
