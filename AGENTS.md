# outlook-cli

CLI for Microsoft Outlook email via Microsoft Graph API. All output JSON.

This repo is public. No client names, family details, phone numbers, ~/projects/hq paths, workspace ids or server IPs. Before any push: check the repo's visibility and run 'git log origin/main..'.

## Setup

```bash
cd ~/tools/outlook-cli && npm install && npm run build && npm link
outlook-cli auth    # Opens browser for OAuth2 login
```

## Commands

```bash
outlook-cli auth                                    # Authenticate (browser OAuth2 flow)
outlook-cli me                                      # Get current user profile

# Emails. send, reply, and forward always end with "Sent from my agent 🤖" (see note below the block)
outlook-cli emails list [--folder inbox] [--limit 25] [--search "query"] [--filter "..."] [--user EMAIL]
outlook-cli emails get --id MESSAGE_ID
outlook-cli emails send --to "a@example.test" --subject "Hi" [--text "body"] [--html "<p>body</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli emails reply --id MESSAGE_ID [--text "reply"] [--html "<p>reply</p>"] [--reply-all] [--to "..."] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # --to/--cc/--bcc add people, stay in the thread
outlook-cli emails forward --id MESSAGE_ID --to "a@example.test" [--html "<p>note</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # original's attachments come along
# --to/--cc/--bcc on reply and forward: comma-separated and/or repeated (--cc a@example.test,b@example.test --cc c@example.test).
outlook-cli emails move --id MESSAGE_ID --folder FOLDER_ID
outlook-cli emails delete --id MESSAGE_ID
outlook-cli emails read --id MESSAGE_ID              # Mark as read
outlook-cli emails unread --id MESSAGE_ID            # Mark as unread
outlook-cli emails attachments --id MESSAGE_ID       # List attachments
outlook-cli emails download --id MESSAGE_ID [--attachment-id ATT_ID] [--out PATH]  # also pulls file attachments nested inside forwarded itemAttachment messages

# Drafts
outlook-cli drafts list [--limit 25]
outlook-cli drafts create --to "a@example.test" --subject "Hi" [--text "body"] [--html "..."] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli drafts reply --id MESSAGE_ID [--text "reply"] [--reply-all] [--attach PATH ...]
outlook-cli drafts forward --id MESSAGE_ID [--to "a@example.test"] [--html "<p>note</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # returns owaUrl
outlook-cli drafts update --id DRAFT_ID [--to "..."] [--subject "..."] [--text "..."] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli drafts attach --id DRAFT_ID --file PATH [--file PATH ...]
outlook-cli drafts detach --id DRAFT_ID --attachment-id ATT_ID
outlook-cli drafts send --id DRAFT_ID
outlook-cli drafts send --id DRAFT_ID --at 2026-09-28T08:30:00-05:00  # Exchange sends it then; the draft keeps its id until then, so drafts delete cancels it
outlook-cli drafts delete --id DRAFT_ID
outlook-cli drafts url --id DRAFT_ID                 # OWA browser URL that opens the draft

# Folders
outlook-cli folders list
outlook-cli folders create --name "My Folder" [--parent PARENT_ID]

# Calendar (timezone defaults to OUTLOOK_DEFAULT_TZ; falls back to America/Chicago)
outlook-cli calendar list [--today] [--start ISO] [--end ISO] [--limit 50] [--search "query"] [--tz "America/Chicago"]
outlook-cli calendar get --id EVENT_ID [--tz TZ]
outlook-cli calendar schedule --addresses "room@example.test,b@example.test" --start ISO --end ISO [--tz TZ] [--interval 30]  # free/busy (getSchedule): availabilityView 0 free, 2 busy
outlook-cli calendar create --subject "..." --start ISO --end ISO [--tz TZ] [--location "..."] [--html "<p>...</p>" | --body "..."] [--attendees "a@example.test,c@example.test"] [--optional-attendees "e@example.test"] [--resources "room@example.test"] [--online] [--all-day] [--recurrence daily|weekly|biweekly|monthly|yearly] [--recurrence-days monday,tuesday]
outlook-cli calendar update --id EVENT_ID [--subject "..."] [--start ISO] [--end ISO] [--tz TZ] [--location "..."] [--html "..." | --body "..."] [--attendees "..."] [--optional-attendees "..."] [--replace-attendees] [--remove-attendees "..."] [--yes] [--online]
# update: attendee flags merge. The CLI reads the event's current guests first and adds the addresses you give; nobody is removed, rooms and response statuses stay. An existing guest named with the other flag changes type (printed).
# Removal needs --replace-attendees or --remove-attendees plus --yes. Without --yes it prints the removal list and final list and exits non-zero without patching.
# --replace-attendees keeps the organizer and rooms; --remove-attendees of an address not on the event is an error. If the read fails or the event has no attendees field, nothing is patched.
# Every attendee change prints data.attendee_changes: {added, type_changed, removed, kept_count} (addresses only).
outlook-cli calendar delete --id EVENT_ID              # NOT silent when you organize a meeting with attendees: Graph sends them a cancellation. Logged in the send log (event_delete)
outlook-cli calendar cancel --id EVENT_ID [--comment "..."]  # Organizer only; sends cancellation to attendees. Occurrence ID cancels one instance
outlook-cli calendar forward --id EVENT_ID --to "a@example.test,c@example.test" [--comment "..."]  # Graph /forward: SENDS the invite now (organizer is told); comment gets the agent footer; occurrence ID forwards one instance. A send: needs approval like any email

# Teams meetings and transcripts (take --join-url or --event-id from calendar list)
outlook-cli meetings get --event-id EVENT_ID          # The onlineMeeting behind the event
outlook-cli meetings transcripts --event-id EVENT_ID  # List transcripts (one per time transcription started)
outlook-cli meetings transcript --event-id EVENT_ID [--since ISO] [--until ISO] [--transcript-id ID] [--format text|vtt] [--out PATH]
# transcript: text is "Speaker Name: words" lines; several transcripts are joined in order.
# --since/--until pick one occurrence of a recurring meeting (they share one join URL).
# Graph returns transcripts for meetings organized by anyone in your tenant.
# Meetings another company organized give MEETING_NOT_FOUND or 403. Meetings older than about 60 days
# give 404 "3004: Specified meeting is not found" (the online meeting expired);
# a meeting without transcription gives NO_TRANSCRIPT.
# OneDrive / SharePoint files (read + upload; scopes are in SCOPES and requested by the next auth:
# Files.Read.All + Sites.Read.All for read, Files.ReadWrite.All for put/sync; the Azure app must consent too)
outlook-cli files search QUERY [--drive me|shared]     # search a drive; slim results (name, id, size, path, hash)
outlook-cli files ls PATH [--drive me|shared]          # list a folder; paths relative to drive root (e.g. 10_Products)
outlook-cli files info PATH [--drive me|shared]        # one item's metadata, never file contents
outlook-cli files get PATH [--drive me|shared] [--out DIR]  # stream-download to cache (default ~/tmp/onedrive-cache, mode 700/600); prints localPath, size, sha256
outlook-cli files put LOCAL PATH [--drive me|shared] [--force] [--mkdir]  # upload one file; ALREADY_EXISTS without --force (--force replaces files only, never folders); files >4 MiB go through a 5 MiB-chunk upload session
outlook-cli files sync LOCAL_DIR PATH [--drive me|shared] --dry-run|--apply  # mirror up: one listing per remote folder (429 = wait and retry), then only new/changed files (size + quickXorHash vs Graph, sha1 fallback); remote newer than local OR content changed since last sync = CONFLICT, never replaced; state (qxh + cTag, never eTag) in ~/.config/outlook-cli/sync-state/ (600); If-Match sent, 412 = conflict; never deletes remote; names OneDrive refuses (" * : < > ? | or edge spaces) are skipped and reported; one leading/trailing / ok, / = drive root; --dry-run plans only
# --drive me = your OneDrive (/me/drive); --drive shared (old alias: ricky) = one folder in another user's OneDrive (not a
# SharePoint site), set in .env: OUTLOOK_SHARED_DRIVE_UPN, OUTLOOK_SHARED_DRIVE_LIBRARY, OUTLOOK_SHARED_DRIVE_FOLDER
# (a missing key gives SHARED_DRIVE_NOT_CONFIGURED naming it). Paths are relative to that folder.
# Its drive id and folder id are found once and cached in tokens/drive-ids.json. No move/delete commands.
# Local details for this machine (whose drive, cron jobs that use this CLI): AGENTS.local.md (untracked).
```

**Reply recipients:** `emails reply` without `--to/--cc/--bcc` calls Graph `/reply` (or `/replyAll`). With any of them it calls `createReply` (or `createReplyAll`), PATCHes the recipients, attaches files, and sends, so the reply stays in the thread with the quoted history. Added addresses join the existing recipients. An address already in To, Cc or Bcc is not added again (compared case-insensitively), no existing recipient is dropped, and the sending mailbox is never added to itself. Without `--reply-all` the reply goes to the sender plus any `--to`. If a step fails, the half-built draft is deleted. To let the mailbox owner review first, use `drafts reply` and send from Outlook.

**Agent footer:** every email from `emails send`, `emails reply`, and `emails forward` ends with "Sent from my agent 🤖" (gray footer in HTML; no opt-out, never doubled). This marks it as agent-written, like the 🤖 prefix in slack-cli and imessages-cli. Do not add it yourself. Drafts get no footer, because the mailbox owner reviews and sends those. Plain `--text` replies and forwards are sent as HTML with `<br>` line breaks, because Graph collapses newlines in a reply comment.

## Send log

Every command that makes Graph send mail writes to a local log first: `emails send`, `emails reply` (all forms), `emails forward`, `drafts send` (also `--at`), `calendar create` with attendees, `calendar update`, `calendar cancel`, `calendar delete` and `calendar forward`. Drafts that are only created or updated are not logged.

- Path: `~/.local/state/outlook-cli/sends.jsonl`. Set `OUTLOOK_CLI_SEND_LOG` to another absolute path to move it.
- Folder mode 0700, file mode 0600, append only. If the file is a symlink, if the file is wider than 0600 or the folder wider than 0700, or if the line cannot be written, the command sends nothing and exits non-zero with code `SEND_LOG_FAILED`.
- Two lines per send, joined by `send_id`: `phase: "attempt"` before the Graph call, then `phase: "result"` with `ok` and the HTTP `status` on failure.
- Attempt fields: `ts` (UTC), `kind` (`send`, `reply`, `reply_all`, `forward`, `draft_send`, `draft_send_scheduled`, `event_create`, `event_update`, `event_cancel`, `event_delete`, `event_forward`), `from` (the mailbox used), `to`, `cc`, `bcc` (lower-case addresses; `null` when Graph picks them and the CLI does not know them), `subject_sha8` (first 8 hex of sha256 of the exact subject; `null` with `subject_note` when the subject is not known), `message_id` or `event_id`, `note`, and `caller` (`herdr_tab_id`, `herdr_pane_id`, `herdr_workspace_id` when set, else `parent_process`; `launchd_label` from `XPC_SERVICE_NAME` when set; `cwd` basename).
- Never logged: subject text, body, comments, attachment names or contents, tokens.

Find who sent a subject: `grep "$(printf %s 'Exact subject' | shasum -a 256 | cut -c1-8)" ~/.local/state/outlook-cli/sends.jsonl`

## Examples

```bash
outlook-cli emails list --search "from:john@example.test" --limit 5
outlook-cli --user someone@other-tenant.example.test emails list --limit 5
outlook-cli emails send --to "bob@example.test" --subject "Hello" --text "Hi Bob"
outlook-cli emails reply --id AAMkAG... --text "Thanks!"
outlook-cli drafts create --to "alice@example.test" --subject "Draft" --text "WIP"

# Open a draft in the browser. `drafts create` returns the same link as `owaUrl`.
open "$(outlook-cli drafts url --id AAMkAG... | jq -r .data.owaUrl)"
outlook-cli drafts send --id AAMkAG...
```

## Attachments

Send messages with attachments by repeating `--attach` (or `--file` for `drafts attach`). Files ≤ 3 MiB are inlined into the request; larger files transparently use Graph's upload sessions via a draft → chunked PUT → send flow.

```bash
# Send with attachments (inline if small, upload session if any > 3 MiB)
outlook-cli emails send --to a@example.test --subject "Hi" --html "<p>Body</p>" \
  --attach ./report.pdf --attach ./chart.png

# Reply with an attachment
outlook-cli emails reply --id MSG_ID --html "<p>see attached</p>" --attach ./notes.md

# Create a draft with attachments (works for any size)
outlook-cli drafts create --to a@example.test --subject "Hi" --attach ./big.zip

# Reply draft with attachments (does not auto-send)
outlook-cli drafts reply --id MSG_ID --html "<p>Draft reply</p>" --attach ./a.pdf

# Update a draft (supports --cc / --bcc / --attach)
outlook-cli drafts update --id DRAFT_ID --cc "extra@example.test" --attach ./extra.pdf

# Add or remove attachments on an existing draft
outlook-cli drafts attach --id DRAFT_ID --file ./extra.pdf --file ./more.png
outlook-cli drafts detach --id DRAFT_ID --attachment-id ATT_ID
```

Inline threshold: `3 MiB` (Graph's file-attachment payload ceiling). Anything larger automatically switches to the `createUploadSession` flow in 5 MiB chunks. No user action required.

## Response Format

All commands return a top-level wrapper: `{ "ok": true, "data": ... }` (or `{ "ok": false, "error": ..., "code": ... }`).

### `emails list`
Returns an object wrapper with message list metadata:

```json
{
  "ok": true,
  "data": {
    "messages": [
      {
        "id": "AAMk...",
        "subject": "Weekly to-dos",
        "from": {"emailAddress": {"name": "Pat Sender", "address": "pat@example.test"}},
        "toRecipients": [{"emailAddress": {"name": "Sam Owner", "address": "sam@example.test"}}],
        "receivedDateTime": "2026-03-20T21:12:21Z",
        "isRead": false,
        "hasAttachments": false,
        "bodyPreview": "Sam, ..."
      }
    ],
    "nextPage": "https://graph.microsoft.com/v1.0/...",
    "count": 1
  }
}
```

### `emails get --id <id>`
Returns the full Microsoft Graph message object:

```json
{
  "ok": true,
  "data": {
    "id": "AAMk...",
    "subject": "Weekly to-dos",
    "from": {"emailAddress": {"name": "Pat Sender", "address": "pat@example.test"}},
    "toRecipients": [{"emailAddress": {"name": "Sam Owner", "address": "sam@example.test"}}],
    "ccRecipients": [],
    "receivedDateTime": "2026-03-20T21:12:21Z",
    "isRead": false,
    "body": {"contentType": "html", "content": "<html>..."},
    "bodyPreview": "Sam, ...",
    "parentFolderId": "AQMk..."
  }
}
```

### `folders list`
Returns `data` as a plain array (not `data.folders`):

```json
{
  "ok": true,
  "data": [
    {
      "id": "AQMk...",
      "displayName": "Inbox",
      "parentFolderId": "AQMk...",
      "childFolderCount": 0,
      "unreadItemCount": 16,
      "totalItemCount": 22,
      "sizeInBytes": 7636536,
      "isHidden": false
    }
  ]
}
```

Requires `.env` with `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`, `OUTLOOK_TENANT_ID`. See `.env.example`.

Tokens stored in `tokens/tokens.json` (auto-refreshed).
