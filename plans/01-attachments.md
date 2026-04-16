# 01 — Attachment support for outlook-cli

## Goal

Add full attachment support to `outlook-cli`: send messages with attached files, manage attachments on existing drafts, support large files via Graph upload sessions, and tighten the existing `download` UX. End the task by sending a real test email from `christian@gxb.vc` to `christian.genco@gmail.com` with an attachment, and verifying delivery via `google-cli`.

## Context

`outlook-cli` lives at `~/tools/outlook-cli` (TypeScript + Commander, calls Microsoft Graph). Today the **read** side of attachments is built (`emails attachments`, `emails download`) but the **write** side is missing — there is no way to attach files when sending mail or creating drafts.

Source map:
- `src/cli.ts` — Commander entrypoint, registers subcommand groups
- `src/commands/emails.ts` — `list`, `get`, `send`, `reply`, `move`, `delete`, `read`, `unread`, `attachments`, `download`
- `src/commands/drafts.ts` — `list`, `create`, `reply`, `update`, `send`, `delete`
- `src/graph.ts` — `graphRequest()` wrapper, `out()`/`err()` JSON helpers
- `AGENTS.md` / `README.md` — user-facing docs (kept in sync via the `CLAUDE.md -> AGENTS.md` symlink)

Build/install: `cd ~/tools/outlook-cli && npm run build && npm link` (already linked globally).

Output convention: every command prints `{"ok":true,"data":...}` or `{"ok":false,"error":...,"code":"..."}` on a single line.

## Microsoft Graph reference

- Inline attachment (≤ ~3 MB): `POST /me/messages` (or include in `message.attachments` on `/me/sendMail`) with
  ```json
  { "@odata.type": "#microsoft.graph.fileAttachment",
    "name": "report.pdf", "contentType": "application/pdf", "contentBytes": "<base64>" }
  ```
- Add attachment to existing draft: `POST /me/messages/{id}/attachments` with the same body.
- Remove attachment: `DELETE /me/messages/{id}/attachments/{attId}`.
- Large file (> 3 MB): `POST /me/messages/{id}/attachments/createUploadSession` with
  ```json
  { "AttachmentItem": { "attachmentType": "file", "name": "big.zip", "size": 12345678 } }
  ```
  Response includes `uploadUrl`. PUT raw bytes in chunks (recommended ~5–10 MB; must be a multiple of 320 KiB except for the last chunk) with header `Content-Range: bytes {start}-{end}/{total}`. Final 201/200 response indicates success.
- Reply attachments: the `/reply` and `/replyAll` actions don't take attachments. Use `createReply` / `createReplyAll` → add attachments to the resulting draft → `send`.

## Tasks

### 1. Shared attachment helper

Create `src/attachments.ts` with:
- `loadAttachment(path: string): { name, contentType, size, bytes: Buffer }` — uses `fs.readFileSync`, `path.basename`, and a content-type lookup. Add `mime-types` to `dependencies` (`npm install mime-types @types/mime-types`); fall back to `application/octet-stream`.
- `INLINE_LIMIT_BYTES = 3 * 1024 * 1024` constant (Graph's hard limit is around 3 MB; use 3 MiB to be safe).
- `toFileAttachmentJson(att): object` — produces the `#microsoft.graph.fileAttachment` payload (base64-encoded `contentBytes`).
- `uploadLargeAttachment(messageId, att): Promise<void>` — calls `createUploadSession`, then PUTs the file in 5 MiB chunks using `fetch` directly against the returned `uploadUrl` (do NOT route through `graphRequest`, as it adds the Authorization header and JSON content-type; the upload URL is pre-authenticated and needs `Content-Length` + `Content-Range`, raw octet-stream body).
- `attachAllToDraft(messageId, paths): Promise<{name,size,id?}[]>` — for each path, branch on size: small → `POST /attachments` inline; large → `uploadLargeAttachment`. Returns a summary list.

### 2. `--attach` on `emails send`

In `src/commands/emails.ts`, change the `send` command:
- Add `.option("--attach <path>", "File to attach (repeatable)", collect, [])` where `collect` is the standard Commander pattern: `(v, acc) => { acc.push(v); return acc; }` (define once at top of file).
- If no attachments OR every attachment is small: keep the existing `/me/sendMail` flow but include them inline in `message.attachments`.
- If any attachment is large (> 3 MiB): switch to draft flow — `POST /me/messages` to create the draft (with all small attachments inline), then `attachAllToDraft` for the large ones, then `POST /me/messages/{id}/send`.
- Output should include `{ message: "Email sent", attachments: [{name,size}, ...] }` when attachments were included.

### 3. `--attach` on `drafts create`

Same pattern as `emails send` minus the final `/send`. Always uses the draft flow when attachments are present, so large-file branching is straightforward. Output the created draft object plus an `attachments` summary.

### 4. `--attach` on `drafts reply` and `emails reply`

- `drafts reply`: after `createReply` / `createReplyAll`, if `--attach` is present call `attachAllToDraft(newDraft.id, paths)` before returning. Output the draft object.
- `emails reply`: today this calls `/reply` directly. When `--attach` is present, switch to the draft flow: `createReply(All)` → attach files → `send`. When no attachments, keep the existing direct `/reply` for backward compatibility.

### 5. `--attach` on `drafts update`

Add `--attach <path>` (repeatable) to `drafts update`. After the PATCH, run `attachAllToDraft` for any provided files. Also add `--cc` and `--bcc` while you're in there (they're missing today — same parsing as `drafts create`).

### 6. New `drafts attach` and `drafts detach` subcommands

In `src/commands/drafts.ts`:

```
drafts attach --id <draftId> --file <path> [--file <path> ...]
drafts detach --id <draftId> --attachment-id <attId>
```

`attach` reuses `attachAllToDraft`. `detach` is a single `DELETE /me/messages/{id}/attachments/{attId}`. Output: `{ ok: true, data: { attached: [...] } }` or `{ ok: true, data: { message: "Attachment removed" } }`.

### 7. Tighten `emails download`

In `src/commands/emails.ts`:
- When `--attachment-id` is given, honor the same "looks like a file path" heuristic as the multi-file path. Specifically: if `outPath` exists and is a directory, write `att.name` inside it; otherwise treat `outPath` as the destination file.
- When iterating all attachments and encountering `itemAttachment` or `referenceAttachment`, include them in the response with a `skipped` array carrying `{name, type, reason}` so callers know they exist. Don't error.
- (Optional, only if trivial) For very large attachments, fetch `/attachments/{id}/$value` and stream the raw bytes to disk via `Readable.fromWeb(res.body).pipe(createWriteStream(...))` instead of base64-decoding the whole `contentBytes` string in memory. If this turns out to need significant rework of `graphRequest`, skip it and leave a TODO comment — not blocking.

### 8. Docs

Update `AGENTS.md` (the symlink covers `CLAUDE.md`) and `README.md`:
- New `## Attachments` section with examples for send / reply / drafts.
- Note the 3 MiB inline threshold and that large files use upload sessions transparently.
- Document `drafts attach` / `drafts detach` and the new `--cc`/`--bcc`/`--attach` flags on `drafts update`.

Sample examples to include:
```bash
outlook-cli emails send --to a@b.com --subject Hi --text Body \
  --attach ./report.pdf --attach ./chart.png

outlook-cli emails reply --id MSG_ID --html "<p>see attached</p>" --attach ./notes.md
outlook-cli drafts create --to a@b.com --subject Hi --attach ./big.zip
outlook-cli drafts attach --id DRAFT_ID --file ./extra.pdf
outlook-cli drafts detach --id DRAFT_ID --attachment-id ATT_ID
```

### 9. Build + smoke test

```bash
cd ~/tools/outlook-cli
npm run build
outlook-cli --help        # sanity: commands still register
outlook-cli emails send --help    # confirm --attach flag is listed
outlook-cli drafts attach --help  # confirm new subcommand
```

### 10. End-to-end test (the "did it actually work" step)

This is the verification the user explicitly asked for. Two test messages — one small inline attachment, one large via upload session — sent from `christian@gxb.vc` (Outlook) to `christian.genco@gmail.com` (Gmail), verified by reading the inbox via `google-cli`.

```bash
# Prepare test files
mkdir -p /tmp/outlook-attach-test
echo "small attachment test from outlook-cli plan 01" > /tmp/outlook-attach-test/small.txt

# 5 MB binary file to force the upload-session branch
dd if=/dev/urandom of=/tmp/outlook-attach-test/big.bin bs=1m count=5

# Tag with a unique token so the verification search is unambiguous
TOKEN="outlook-attach-test-$(date +%s)"
echo "Test token: $TOKEN"

# Send #1: small inline attachment
outlook-cli emails send \
  --to christian.genco@gmail.com \
  --subject "[$TOKEN] small inline attachment" \
  --html "<p>Plan 01 verification — small inline file.</p>" \
  --attach /tmp/outlook-attach-test/small.txt

# Send #2: large attachment (forces upload session)
outlook-cli emails send \
  --to christian.genco@gmail.com \
  --subject "[$TOKEN] large upload-session attachment" \
  --html "<p>Plan 01 verification — large upload-session file.</p>" \
  --attach /tmp/outlook-attach-test/big.bin

# Wait briefly for delivery (Outlook -> Gmail is usually a few seconds)
sleep 30

# Verify both arrived in Gmail
google-cli emails list --q "subject:$TOKEN"

# For each returned message id, confirm the attachment is present.
# (Use whatever fields google-cli returns; if it exposes payload parts
# or hasAttachment, assert on that. Otherwise emails get --id MSG_ID
# and inspect the JSON for an attachment payload part.)
google-cli emails list --q "subject:$TOKEN has:attachment"
```

Pass criteria — ALL of:
- Both `outlook-cli emails send` calls return `{"ok":true,...}`.
- `google-cli emails list --q "subject:$TOKEN"` returns 2 messages.
- `google-cli emails list --q "subject:$TOKEN has:attachment"` also returns 2 messages (i.e. Gmail saw the attachments on both).
- Reading either message via `google-cli emails get --id <id>` shows an attachment payload part with the expected filename (`small.txt` or `big.bin`).

If any of those fail, dig in:
- `outlook-cli emails list --search "subject:$TOKEN" --filter "hasAttachments eq true"` — confirm the message left Outlook with attachments attached.
- For the large-file path, the most likely failure modes are (a) wrong `Content-Range` math, (b) sending JSON content-type to the upload URL, or (c) using `graphRequest` instead of raw `fetch` for the chunk PUTs. See task 1 notes.

### 11. Clean up

- Remove `/tmp/outlook-attach-test/`.
- Don't delete the test emails — leave them in the inbox for the user to inspect.
- Print a final summary line with the test token and the two message IDs (Outlook side and/or Gmail side) so the user can find them later.

## Out of scope

- Inline `cid:` images for HTML bodies (separate plan).
- Streaming download via `/$value` if it requires rewiring `graphRequest` (mentioned as optional in task 7).
- Pagination changes, OAuth scope changes, or any auth work — `Mail.ReadWrite` and `Mail.Send` should already be granted.
