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

# Emails. send, reply, and forward always end with "Sent by my agent 🤖" (see note below the block)
outlook-cli emails list [--folder inbox] [--limit 25] [--search "query"] [--filter "..."] [--user EMAIL]
outlook-cli emails get --id MESSAGE_ID
outlook-cli emails send --to "a@b.com" --subject "Hi" [--text "body"] [--html "<p>body</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]
outlook-cli emails reply --id MESSAGE_ID [--text "reply"] [--html "<p>reply</p>"] [--reply-all] [--attach PATH ...]
outlook-cli emails forward --id MESSAGE_ID --to "a@b.com" [--html "<p>note</p>"] [--cc "..."] [--bcc "..."] [--attach PATH ...]  # original's attachments come along
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
# Graph only returns transcripts for meetings Christian organized. Others give MEETING_NOT_FOUND;
# a meeting without transcription gives NO_TRANSCRIPT.
# ~/utils/sync-teams-transcripts.sh (cron) uses this to copy transcripts into ~/scribe and scribe.gxb.vc.
```

**Agent footer:** every email from `emails send`, `emails reply`, and `emails forward` ends with "Sent by my agent 🤖" (gray footer in HTML; no opt-out, never doubled). This marks it as agent-written, like the 🤖 prefix in slack-cli and imessages-cli. Do not add it yourself. Drafts get no footer, because Christian reviews and sends those. Plain `--text` replies and forwards are sent as HTML with `<br>` line breaks, because Graph collapses newlines in a reply comment.

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
