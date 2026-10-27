# outlook-cli

CLI for Microsoft Outlook email via Microsoft Graph API. All output JSON.

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
outlook-cli emails send --to "a@b.com" --subject "Hi" [--text "body"] [--html "<p>body</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli emails reply --id MESSAGE_ID [--text "reply"] [--html "<p>reply</p>"] [--reply-all] [--to "..."] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # --to/--cc/--bcc add people, stay in the thread
outlook-cli emails forward --id MESSAGE_ID --to "a@b.com" [--html "<p>note</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # original's attachments come along
# --to/--cc/--bcc on reply and forward: comma-separated and/or repeated (--cc a@x.com,b@x.com --cc c@x.com).
outlook-cli emails move --id MESSAGE_ID --folder FOLDER_ID
outlook-cli emails delete --id MESSAGE_ID
outlook-cli emails read --id MESSAGE_ID              # Mark as read
outlook-cli emails unread --id MESSAGE_ID            # Mark as unread
outlook-cli emails attachments --id MESSAGE_ID       # List attachments
outlook-cli emails download --id MESSAGE_ID [--attachment-id ATT_ID] [--out PATH]  # also pulls file attachments nested inside forwarded itemAttachment messages

# Drafts
outlook-cli drafts list [--limit 25]
outlook-cli drafts create --to "a@b.com" --subject "Hi" [--text "body"] [--html "..."] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli drafts reply --id MESSAGE_ID [--text "reply"] [--reply-all] [--attach PATH ...]
outlook-cli drafts forward --id MESSAGE_ID [--to "a@b.com"] [--html "<p>note</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # returns owaUrl
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
outlook-cli calendar create --subject "..." --start ISO --end ISO [--tz TZ] [--location "..."] [--html "<p>...</p>" | --body "..."] [--attendees "a@b.com,c@d.com"] [--optional-attendees "e@f.com"] [--online] [--all-day] [--recurrence daily|weekly|biweekly|monthly|yearly] [--recurrence-days monday,tuesday]
outlook-cli calendar update --id EVENT_ID [--subject "..."] [--start ISO] [--end ISO] [--tz TZ] [--location "..."] [--html "..." | --body "..."] [--attendees "..."] [--optional-attendees "..."] [--online]
# update: passing --attendees or --optional-attendees replaces the whole attendee list. Pass both to keep both.
outlook-cli calendar delete --id EVENT_ID              # Silent — does not notify attendees
outlook-cli calendar cancel --id EVENT_ID [--comment "..."]  # Organizer only; sends cancellation to attendees. Occurrence ID cancels one instance

# Teams meetings and transcripts (take --join-url or --event-id from calendar list)
outlook-cli meetings get --event-id EVENT_ID          # The onlineMeeting behind the event
outlook-cli meetings transcripts --event-id EVENT_ID  # List transcripts (one per time transcription started)
outlook-cli meetings transcript --event-id EVENT_ID [--since ISO] [--until ISO] [--transcript-id ID] [--format text|vtt] [--out PATH]
# transcript: text is "Speaker Name: words" lines; several transcripts are joined in order.
# --since/--until pick one occurrence of a recurring meeting (they share one join URL).
# Graph returns transcripts for meetings organized by anyone in the gxb.vc tenant (Christian, Ricky, Sarah).
# Meetings another company organized give MEETING_NOT_FOUND or 403. Meetings older than about 60 days
# give 404 "3004: Specified meeting is not found" (the online meeting expired);
# a meeting without transcription gives NO_TRANSCRIPT.
# OneDrive / SharePoint files (read + upload; scopes are in SCOPES and requested by the next auth:
# Files.Read.All + Sites.Read.All for read, Files.ReadWrite.All for put/sync; the Azure app must consent too)
outlook-cli files search QUERY [--drive me|ricky]     # search a drive; slim results (name, id, size, path, hash)
outlook-cli files ls PATH [--drive me|ricky]          # list a folder; paths relative to drive root (e.g. 10_Products)
outlook-cli files info PATH [--drive me|ricky]        # one item's metadata, never file contents
outlook-cli files get PATH [--drive me|ricky] [--out DIR]  # stream-download to cache (default ~/tmp/onedrive-cache, mode 700/600); prints localPath, size, sha256
outlook-cli files put LOCAL PATH [--drive me|ricky] [--force] [--mkdir]  # upload one file; ALREADY_EXISTS without --force (--force replaces files only, never folders); files >4 MiB go through a 5 MiB-chunk upload session
outlook-cli files sync LOCAL_DIR PATH [--drive me|ricky] --dry-run|--apply  # mirror up: only new/changed files (size + quickXorHash vs Graph, sha1 fallback); remote newer than local OR content changed since last sync = CONFLICT, never replaced; state (qxh + cTag, never eTag) in ~/.config/outlook-cli/sync-state/ (600); If-Match sent, 412 = conflict; never deletes remote; names OneDrive refuses (" * : < > ? | or edge spaces) are skipped and reported; one leading/trailing / ok, / = drive root; --dry-run plans only
# --drive me = your OneDrive (/me/drive); --drive ricky = shared library "Ricky Bureau - GXB", which is
# the GXB folder in ricky@gxb.vc's OneDrive (not a SharePoint site). Paths are relative to that folder.
# Its drive id and folder id are found once and cached in tokens/drive-ids.json. No move/delete commands.

# ~/utils/sync-teams-transcripts.sh (cron) uses this to copy transcripts into ~/scribe and scribe.gxb.vc.
```

**Reply recipients:** `emails reply` without `--to/--cc/--bcc` calls Graph `/reply` (or `/replyAll`). With any of them it calls `createReply` (or `createReplyAll`), PATCHes the recipients, attaches files, and sends, so the reply stays in the thread with the quoted history. Added addresses join the existing recipients. An address already in To, Cc or Bcc is not added again (compared case-insensitively), no existing recipient is dropped, and the sending mailbox is never added to itself. Without `--reply-all` the reply goes to the sender plus any `--to`. If a step fails, the half-built draft is deleted. To let Christian review first, use `drafts reply` and send from Outlook.

**Agent footer:** every email from `emails send`, `emails reply`, and `emails forward` ends with "Sent from my agent 🤖" (gray footer in HTML; no opt-out, never doubled). This marks it as agent-written, like the 🤖 prefix in slack-cli and imessages-cli. Do not add it yourself. Drafts get no footer, because Christian reviews and sends those. Plain `--text` replies and forwards are sent as HTML with `<br>` line breaks, because Graph collapses newlines in a reply comment.

## Examples

```bash
outlook-cli emails list --search "from:john@example.com" --limit 5
outlook-cli --user charles@brenmere.com emails list --limit 5
outlook-cli emails send --to "bob@example.com" --subject "Hello" --text "Hi Bob"
outlook-cli emails reply --id AAMkAG... --text "Thanks!"
outlook-cli drafts create --to "alice@example.com" --subject "Draft" --text "WIP"

# Open a draft in the browser. `drafts create` returns the same link as `owaUrl`.
open "$(outlook-cli drafts url --id AAMkAG... | jq -r .data.owaUrl)"
outlook-cli drafts send --id AAMkAG...
```

## Attachments

Send messages with attachments by repeating `--attach` (or `--file` for `drafts attach`). Files ≤ 3 MiB are inlined into the request; larger files transparently use Graph's upload sessions via a draft → chunked PUT → send flow.

```bash
# Send with attachments (inline if small, upload session if any > 3 MiB)
outlook-cli emails send --to a@b.com --subject "Hi" --html "<p>Body</p>" \
  --attach ./report.pdf --attach ./chart.png

# Reply with an attachment
outlook-cli emails reply --id MSG_ID --html "<p>see attached</p>" --attach ./notes.md

# Create a draft with attachments (works for any size)
outlook-cli drafts create --to a@b.com --subject "Hi" --attach ./big.zip

# Reply draft with attachments (does not auto-send)
outlook-cli drafts reply --id MSG_ID --html "<p>Draft reply</p>" --attach ./a.pdf

# Update a draft (supports --cc / --bcc / --attach)
outlook-cli drafts update --id DRAFT_ID --cc "extra@b.com" --attach ./extra.pdf

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
        "subject": "To-dos for ISN/ TCY",
        "from": {"emailAddress": {"name": "Ricky Bureau", "address": "ricky@gxb.vc"}},
        "toRecipients": [{"emailAddress": {"name": "Christian Genco", "address": "christian@gxb.vc"}}],
        "receivedDateTime": "2026-03-20T21:12:21Z",
        "isRead": false,
        "hasAttachments": false,
        "bodyPreview": "Christian, ..."
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
    "subject": "To-dos for ISN/ TCY",
    "from": {"emailAddress": {"name": "Ricky Bureau", "address": "ricky@gxb.vc"}},
    "toRecipients": [{"emailAddress": {"name": "Christian Genco", "address": "christian@gxb.vc"}}],
    "ccRecipients": [],
    "receivedDateTime": "2026-03-20T21:12:21Z",
    "isRead": false,
    "body": {"contentType": "html", "content": "<html>..."},
    "bodyPreview": "Christian, ...",
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
