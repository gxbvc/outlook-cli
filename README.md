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
- `outlook-cli calendar create --subject SUBJ --start ISO --end ISO [--tz TZ] [--location LOC] [--html HTML | --body TEXT] [--attendees "a@example.test,c@example.test"] [--resources "room@example.test"] [--online] [--all-day]` — Create event; `--resources` invites room mailboxes as attendee type `resource`
- `outlook-cli calendar schedule --addresses "room@example.test,b@example.test" --start ISO --end ISO [--tz TZ] [--interval 30]` — Free/busy for mailboxes or rooms (Graph `POST /me/calendar/getSchedule`): `schedules[].availabilityView` (0 free, 1 tentative, 2 busy, 3 away, 4 working elsewhere) and `scheduleItems`
- `outlook-cli calendar update --id ID [--subject ...] [--start ...] [--end ...] [--location ...] [--html ...] [--attendees ...] [--optional-attendees ...] [--replace-attendees] [--remove-attendees ...] [--yes]` — Update event. Attendee flags merge into the current guest list (nobody is removed). Removing guests needs `--replace-attendees` or `--remove-attendees` plus `--yes`; without `--yes` it prints the removal list and exits non-zero without patching. Rooms and the organizer are kept
- `outlook-cli calendar delete --id ID` — Delete event. Not silent when you organize a meeting with attendees: Graph sends them (and any room) a cancellation. Logged in the send log as `event_delete`. Use `calendar cancel` to add a comment
- `outlook-cli calendar cancel --id ID [--comment TEXT]` — Cancel an event you organize and notify the attendees and rooms (Graph `POST /me/events/{id}/cancel`)
- `outlook-cli calendar forward --id ID --to ADDRESSES [--comment TEXT]` — Forward a meeting invite (Graph `POST /me/events/{id}/forward`). It sends at once; the organizer is told. The comment ends with the agent footer

Timezone defaults to `OUTLOOK_DEFAULT_TZ` (`America/Chicago` if unset). Override per-call with `--tz`. `--today` uses start/end-of-day in the chosen timezone and expands recurrences via `/me/calendarView`.

### Files (OneDrive and SharePoint)
- `outlook-cli files search QUERY [--drive me|shared]` — Search files and folders in a drive
- `outlook-cli files ls PATH [--drive me|shared]` — List a folder; paths are relative to the drive root (e.g. `10_Products`)
- `outlook-cli files info PATH [--drive me|shared]` — Show one item's metadata (name, id, size, hash); never file contents
- `outlook-cli files get PATH [--drive me|shared] [--out DIR]` — Download a file. Default cache is `~/tmp/onedrive-cache` (dir mode 700, files mode 600); saves under `<cache>/<drive>/<path>`. Streams to disk (safe for large files) and prints the local path, size and sha256.
- `outlook-cli files put LOCAL PATH [--drive me|shared] [--force] [--mkdir]` — Upload a local file. Never deletes or moves anything. Refuses to overwrite (`ALREADY_EXISTS`) without `--force`; `--force` replaces a file only, never a folder. Missing parent folders need `--mkdir`. Files over 4 MiB upload through a Graph upload session in 5 MiB chunks (never held in memory). Prints size, sha256, remote id.
- `outlook-cli files sync LOCAL_DIR PATH [--drive me|shared] --dry-run|--apply` — Mirror a local folder into the drive: lists each remote folder once (a 429 waits Retry-After and retries, up to 5 times), then uploads only new or changed files (size + quickXorHash compared against Graph's hashes; sha1 fallback for personal OneDrive, computed locally by Microsoft's published algorithm). A changed remote file is replaced only when the remote copy is older than the local one; a remote edit newer than local is a conflict, never replaced. The last-synced content markers (quickXorHash, cTag) per file are kept in `~/.config/outlook-cli/sync-state/` (mode 600): a remote whose content changed since the last sync is a conflict even when it looks older (both-sides-edited protection; eTag is never compared, SharePoint bumps it on metadata-only changes). Uploads of changed files send `If-Match` from the plan; a 412 is a conflict. Case-only name collisions, symlinks and names OneDrive does not allow (`" * : < > ? |`, leading or trailing spaces) are reported and skipped; one bad local file name does not stop the rest. Never deletes remote files, never replaces a remote folder. `--dry-run` prints the plan and counts only. Creates missing remote folders on `--apply`. Paths accept one leading or trailing `/`; `/` is the drive root.

Drives: `me` is your OneDrive (`/me/drive`). `shared` (old alias: `ricky`) is one folder in another user's OneDrive, set in `.env` with `OUTLOOK_SHARED_DRIVE_UPN`, `OUTLOOK_SHARED_DRIVE_LIBRARY` and `OUTLOOK_SHARED_DRIVE_FOLDER` (a missing key gives `SHARED_DRIVE_NOT_CONFIGURED` naming it). It is not a SharePoint site, so paths are relative to that folder. Its drive id and folder id are found once (`/users/<owner>/drive/root:/<folder>`, exact folder only) and cached in `tokens/drive-ids.json` as `{ driveId, itemId }`; every run checks that the cached item is still that folder before using it.

Read only. There are no delete, move, or write commands, and nothing in OneDrive is ever modified.

**Setup (needs consent):** the files scopes — `Files.Read.All`, `Sites.Read.All`, `Files.ReadWrite.All` (delegated, `Files.ReadWrite.All` only for `put`/`sync`) — are in `SCOPES` in `src/config.ts` and requested by the next `outlook-cli auth`; the Azure app registration must also have them added and consented. Until both are done, `files` commands return `CONSENT_REQUIRED` (403) with the hint to run `outlook-cli auth`.

## Send log

Every command that makes Graph send mail writes to a local log first: `emails send`, `emails reply` (all forms), `emails forward`, `drafts send` (also `--at`), `calendar create` with attendees, `calendar update`, `calendar cancel`, `calendar delete` and `calendar forward`. Drafts that are only created or updated are not logged.

- Path: `~/.local/state/outlook-cli/sends.jsonl`. Set `OUTLOOK_CLI_SEND_LOG` to another absolute path to move it.
- Folder mode 0700, file mode 0600, append only. If the file is a symlink, if the file is wider than 0600 or the folder wider than 0700, or if the line cannot be written, the command sends nothing and exits non-zero with code `SEND_LOG_FAILED`.
- Two lines per send, joined by `send_id`: `phase: "attempt"` before the Graph call, then `phase: "result"` with `ok` and the HTTP `status` on failure.
- Attempt fields: `ts` (UTC), `kind` (`send`, `reply`, `reply_all`, `forward`, `draft_send`, `draft_send_scheduled`, `event_create`, `event_update`, `event_cancel`, `event_delete`, `event_forward`), `from` (the mailbox used), `to`, `cc`, `bcc` (lower-case addresses; `null` when Graph picks them and the CLI does not know them), `subject_sha8` (first 8 hex of sha256 of the exact subject; `null` with `subject_note` when the subject is not known), `message_id` or `event_id`, `note`, and `caller` (`herdr_tab_id`, `herdr_pane_id`, `herdr_workspace_id` when set, else `parent_process`; `launchd_label` from `XPC_SERVICE_NAME` when set; `cwd` basename).
- Never logged: subject text, body, comments, attachment names or contents, tokens.

Find who sent a subject: `grep "$(printf %s 'Exact subject' | shasum -a 256 | cut -c1-8)" ~/.local/state/outlook-cli/sends.jsonl`

## Tips

- **Always use `--html` instead of `--text`** when sending, replying, or creating drafts. Plain text (`--text`) renders without line breaks in Outlook — newlines are ignored. Use `<p>` tags for paragraphs and `<br>` for line breaks. Example:
  ```bash
  outlook-cli drafts create --to "someone@example.test" --subject "Hello" \
    --html '<p>Hi there,</p><p>Here is my message.</p><p>Thanks!</p>'
  ```

## Attachments

Repeat `--attach PATH` to attach files. Files ≤ 3 MiB are inlined into the request; larger files are transparently uploaded via Graph's `createUploadSession` flow in 5 MiB chunks. Works for `emails send`, `emails reply`, `emails forward`, `drafts create`, `drafts reply`, `drafts forward`, and `drafts update`. You can also manage attachments on an existing draft directly:

```bash
outlook-cli emails send --to a@example.test --subject Hi --html "<p>Body</p>" \
  --attach ./report.pdf --attach ./chart.png

outlook-cli emails reply --id MSG_ID --html "<p>see attached</p>" --attach ./notes.md

outlook-cli drafts create --to a@example.test --subject Hi --attach ./big.zip
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
