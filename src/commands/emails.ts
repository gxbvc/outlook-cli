import { Command } from "commander";
import { writeFileSync, mkdirSync, existsSync, statSync } from "fs";
import { join, resolve } from "path";
import { graphRequest, out, err } from "../graph.js";
import {
  loadAll,
  hasLarge,
  toFileAttachmentJson,
  attachAllToDraft,
  INLINE_LIMIT_BYTES,
} from "../attachments.js";

const collect = (v: string, acc: string[]) => { acc.push(v); return acc; };

export function registerEmails(program: Command) {
  const emails = program.command("emails").description("Manage emails");

  emails
    .command("list")
    .description("List emails")
    .option("--folder <folder>", "Folder name or ID", "inbox")
    .option("--limit <n>", "Max results", "25")
    .option("--search <query>", "Search query (KQL)")
    .option("--filter <filter>", "OData filter")
    .option("--page-token <token>", "Pagination token (nextLink)")
    .action(async (opts) => {
      try {
        let url: string;
        if (opts.pageToken) {
          url = opts.pageToken;
        } else {
          const params = new URLSearchParams({ "$top": opts.limit });
          params.append("$select", "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview");
          // Graph API disallows $orderby together with $search
          if (!opts.search) params.append("$orderby", "receivedDateTime desc");
          if (opts.search) params.append("$search", `"${opts.search}"`);
          if (opts.filter) params.append("$filter", opts.filter);
          url = `/me/mailFolders/${opts.folder}/messages?${params.toString()}`;
        }
        const data = await graphRequest(url);
        out({
          messages: data.value,
          nextPage: data["@odata.nextLink"] || null,
          count: data.value.length,
        });
      } catch (e: any) { err(e.message, "LIST_FAILED"); }
    });

  emails
    .command("get")
    .description("Get a single email by ID")
    .requiredOption("--id <id>", "Message ID")
    .action(async (opts) => {
      try {
        const data = await graphRequest(`/me/messages/${opts.id}`);
        out(data);
      } catch (e: any) { err(e.message, "GET_FAILED"); }
    });

  emails
    .command("send")
    .description("Send an email")
    .requiredOption("--to <addresses>", "Comma-separated recipients")
    .requiredOption("--subject <subject>", "Subject line")
    .option("--text <text>", "Plain text body")
    .option("--html <html>", "HTML body")
    .option("--cc <addresses>", "Comma-separated CC")
    .option("--bcc <addresses>", "Comma-separated BCC")
    .option("--attach <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const toRecipients = opts.to.split(",").map((e: string) => ({
          emailAddress: { address: e.trim() },
        }));
        const ccRecipients = opts.cc
          ? opts.cc.split(",").map((e: string) => ({ emailAddress: { address: e.trim() } }))
          : undefined;
        const bccRecipients = opts.bcc
          ? opts.bcc.split(",").map((e: string) => ({ emailAddress: { address: e.trim() } }))
          : undefined;

        const attachPaths: string[] = opts.attach || [];

        if (attachPaths.length === 0 || !hasLarge(attachPaths)) {
          // All small: single /me/sendMail call with inline attachments
          const message: any = {
            subject: opts.subject,
            body: {
              contentType: opts.html ? "HTML" : "Text",
              content: opts.html || opts.text || "",
            },
            toRecipients,
          };
          if (ccRecipients) message.ccRecipients = ccRecipients;
          if (bccRecipients) message.bccRecipients = bccRecipients;

          const attSummary: { name: string; size: number }[] = [];
          if (attachPaths.length > 0) {
            const loaded = loadAll(attachPaths);
            message.attachments = loaded.map(toFileAttachmentJson);
            for (const a of loaded) attSummary.push({ name: a.name, size: a.size });
          }

          await graphRequest("/me/sendMail", {
            method: "POST",
            body: JSON.stringify({ message }),
          });

          out(attachPaths.length > 0
            ? { message: "Email sent", attachments: attSummary }
            : { message: "Email sent" });
          return;
        }

        // Large attachments present: create draft, attach, send
        const draftBody: any = {
          subject: opts.subject,
          body: {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.html || opts.text || "",
          },
          toRecipients,
        };
        if (ccRecipients) draftBody.ccRecipients = ccRecipients;
        if (bccRecipients) draftBody.bccRecipients = bccRecipients;

        // Inline-attach all small ones at draft creation
        const loaded = loadAll(attachPaths);
        const smallInline = loaded.filter((a) => a.size <= INLINE_LIMIT_BYTES);
        const largeRemaining = loaded.filter((a) => a.size > INLINE_LIMIT_BYTES);
        if (smallInline.length > 0) {
          draftBody.attachments = smallInline.map(toFileAttachmentJson);
        }

        const draft = await graphRequest("/me/messages", {
          method: "POST",
          body: JSON.stringify(draftBody),
        });

        for (const a of largeRemaining) {
          await attachAllToDraft(draft.id, [a.path]);
        }

        await graphRequest(`/me/messages/${draft.id}/send`, { method: "POST" });

        out({
          message: "Email sent",
          attachments: loaded.map((a) => ({ name: a.name, size: a.size })),
        });
      } catch (e: any) { err(e.message, "SEND_FAILED"); }
    });

  emails
    .command("reply")
    .description("Reply to an email")
    .requiredOption("--id <id>", "Message ID to reply to")
    .option("--text <text>", "Plain text reply")
    .option("--html <html>", "HTML reply")
    .option("--reply-all", "Reply to all recipients")
    .option("--attach <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const attachPaths: string[] = opts.attach || [];

        if (attachPaths.length === 0) {
          const endpoint = opts.replyAll ? "replyAll" : "reply";
          await graphRequest(`/me/messages/${opts.id}/${endpoint}`, {
            method: "POST",
            body: JSON.stringify({
              comment: opts.html || opts.text || "",
            }),
          });
          out({ message: `Reply sent` });
          return;
        }

        // Attachments: createReply(All) → attach files → send
        const createEndpoint = opts.replyAll ? "createReplyAll" : "createReply";
        const draft = await graphRequest(`/me/messages/${opts.id}/${createEndpoint}`, {
          method: "POST",
          body: JSON.stringify({ comment: opts.html || opts.text || "" }),
        });

        const attached = await attachAllToDraft(draft.id, attachPaths);
        await graphRequest(`/me/messages/${draft.id}/send`, { method: "POST" });

        out({ message: "Reply sent", attachments: attached.map((a) => ({ name: a.name, size: a.size })) });
      } catch (e: any) { err(e.message, "REPLY_FAILED"); }
    });

  emails
    .command("move")
    .description("Move an email to a folder")
    .requiredOption("--id <id>", "Message ID")
    .requiredOption("--folder <folder>", "Destination folder ID")
    .action(async (opts) => {
      try {
        const data = await graphRequest(`/me/messages/${opts.id}/move`, {
          method: "POST",
          body: JSON.stringify({ destinationId: opts.folder }),
        });
        out(data);
      } catch (e: any) { err(e.message, "MOVE_FAILED"); }
    });

  emails
    .command("delete")
    .description("Delete an email")
    .requiredOption("--id <id>", "Message ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}`, { method: "DELETE" });
        out({ message: "Email deleted" });
      } catch (e: any) { err(e.message, "DELETE_FAILED"); }
    });

  emails
    .command("read")
    .description("Mark an email as read")
    .requiredOption("--id <id>", "Message ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}`, {
          method: "PATCH",
          body: JSON.stringify({ isRead: true }),
        });
        out({ message: "Marked as read" });
      } catch (e: any) { err(e.message, "UPDATE_FAILED"); }
    });

  emails
    .command("unread")
    .description("Mark an email as unread")
    .requiredOption("--id <id>", "Message ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}`, {
          method: "PATCH",
          body: JSON.stringify({ isRead: false }),
        });
        out({ message: "Marked as unread" });
      } catch (e: any) { err(e.message, "UPDATE_FAILED"); }
    });

  emails
    .command("attachments")
    .description("List attachments on an email")
    .requiredOption("--id <id>", "Message ID")
    .action(async (opts) => {
      try {
        const data = await graphRequest(`/me/messages/${opts.id}/attachments?$select=id,name,contentType,size`);
        out({
          attachments: data.value.map((a: any) => ({
            id: a.id,
            name: a.name,
            contentType: a.contentType,
            size: a.size,
          })),
          count: data.value.length,
        });
      } catch (e: any) { err(e.message, "ATTACHMENTS_FAILED"); }
    });

  emails
    .command("download")
    .description("Download attachment(s) from an email")
    .requiredOption("--id <id>", "Message ID")
    .option("--attachment-id <attachmentId>", "Specific attachment ID (downloads all if omitted)")
    .option("--out <path>", "Output file path or directory", ".")
    .action(async (opts) => {
      try {
        const outPath = resolve(opts.out);
        let rawAttachments: any[];

        if (opts.attachmentId) {
          const att = await graphRequest(`/me/messages/${opts.id}/attachments/${opts.attachmentId}`);
          rawAttachments = [att];
        } else {
          const data = await graphRequest(`/me/messages/${opts.id}/attachments`);
          rawAttachments = data.value;
        }

        const fileAttachments = rawAttachments.filter(
          (a: any) => a["@odata.type"] === "#microsoft.graph.fileAttachment"
        );
        const skipped = rawAttachments
          .filter((a: any) => a["@odata.type"] !== "#microsoft.graph.fileAttachment")
          .map((a: any) => ({
            name: a.name,
            type: a["@odata.type"],
            reason: a["@odata.type"] === "#microsoft.graph.itemAttachment"
              ? "item attachment (nested message/event) — use a different endpoint to retrieve"
              : a["@odata.type"] === "#microsoft.graph.referenceAttachment"
                ? "reference attachment (cloud link) — open the sourceUrl directly"
                : "unsupported attachment type",
          }));

        if (fileAttachments.length === 0) {
          out({ files: [], skipped, message: "No file attachments found" });
          return;
        }

        const treatAsDir =
          fileAttachments.length > 1 ||
          opts.out === "." ||
          opts.out.endsWith("/") ||
          (existsSync(outPath) && statSync(outPath).isDirectory());

        const saved: any[] = [];
        for (const att of fileAttachments) {
          const buf = Buffer.from(att.contentBytes, "base64");
          let filePath: string;

          if (treatAsDir) {
            mkdirSync(outPath, { recursive: true });
            filePath = join(outPath, att.name);
          } else {
            filePath = outPath;
          }

          writeFileSync(filePath, buf);
          saved.push({ file: filePath, name: att.name, bytes: buf.length });
        }

        out({ files: saved, count: saved.length, skipped });
      } catch (e: any) { err(e.message, "DOWNLOAD_FAILED"); }
    });
}
