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

# Emails
outlook-cli emails list [--folder inbox] [--limit 25] [--search "query"] [--filter "..."]
outlook-cli emails get --id MESSAGE_ID
outlook-cli emails send --to "a@b.com" --subject "Hi" [--text "body"] [--html "<p>body</p>"] [--cc "..."] [--bcc "..."]
outlook-cli emails reply --id MESSAGE_ID [--text "reply"] [--html "<p>reply</p>"] [--reply-all]
outlook-cli emails move --id MESSAGE_ID --folder FOLDER_ID
outlook-cli emails delete --id MESSAGE_ID
outlook-cli emails read --id MESSAGE_ID              # Mark as read
outlook-cli emails unread --id MESSAGE_ID            # Mark as unread

# Drafts
outlook-cli drafts list [--limit 25]
outlook-cli drafts create --to "a@b.com" --subject "Hi" [--text "body"] [--html "..."]
outlook-cli drafts reply --id MESSAGE_ID [--text "reply"] [--reply-all]
outlook-cli drafts update --id DRAFT_ID [--to "..."] [--subject "..."] [--text "..."]
outlook-cli drafts send --id DRAFT_ID
outlook-cli drafts delete --id DRAFT_ID

# Folders
outlook-cli folders list
outlook-cli folders create --name "My Folder" [--parent PARENT_ID]
```

## Examples

```bash
outlook-cli emails list --search "from:john@example.com" --limit 5
outlook-cli emails send --to "bob@example.com" --subject "Hello" --text "Hi Bob"
outlook-cli emails reply --id AAMkAG... --text "Thanks!"
outlook-cli drafts create --to "alice@example.com" --subject "Draft" --text "WIP"
outlook-cli drafts send --id AAMkAG...
```

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
