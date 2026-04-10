import { Command } from "commander";
import { writeFileSync, mkdirSync } from "fs";
import { join, resolve } from "path";
import { graphRequest, out, err } from "../graph.js";

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
    .action(async (opts) => {
      try {
        const toRecipients = opts.to.split(",").map((e: string) => ({
          emailAddress: { address: e.trim() },
        }));
        const body: any = {
          message: {
            subject: opts.subject,
            body: {
              contentType: opts.html ? "HTML" : "Text",
              content: opts.html || opts.text || "",
            },
            toRecipients,
          },
        };
        if (opts.cc) {
          body.message.ccRecipients = opts.cc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        if (opts.bcc) {
          body.message.bccRecipients = opts.bcc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        await graphRequest("/me/sendMail", {
          method: "POST",
          body: JSON.stringify(body),
        });
        out({ message: "Email sent" });
      } catch (e: any) { err(e.message, "SEND_FAILED"); }
    });

  emails
    .command("reply")
    .description("Reply to an email")
    .requiredOption("--id <id>", "Message ID to reply to")
    .option("--text <text>", "Plain text reply")
    .option("--html <html>", "HTML reply")
    .option("--reply-all", "Reply to all recipients")
    .action(async (opts) => {
      try {
        const endpoint = opts.replyAll ? "replyAll" : "reply";
        await graphRequest(`/me/messages/${opts.id}/${endpoint}`, {
          method: "POST",
          body: JSON.stringify({
            comment: opts.html || opts.text || "",
          }),
        });
        out({ message: `Reply sent` });
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
        let attachments: any[];

        if (opts.attachmentId) {
          const att = await graphRequest(`/me/messages/${opts.id}/attachments/${opts.attachmentId}`);
          attachments = [att];
        } else {
          const data = await graphRequest(`/me/messages/${opts.id}/attachments`);
          attachments = data.value.filter((a: any) => a["@odata.type"] === "#microsoft.graph.fileAttachment");
        }

        if (attachments.length === 0) {
          out({ files: [], message: "No file attachments found" });
          return;
        }

        const saved: any[] = [];
        for (const att of attachments) {
          const buf = Buffer.from(att.contentBytes, "base64");
          let filePath: string;

          if (attachments.length === 1 && !outPath.endsWith("/") && opts.out !== ".") {
            // Single attachment + --out looks like a file path
            filePath = outPath;
          } else {
            mkdirSync(outPath, { recursive: true });
            filePath = join(outPath, att.name);
          }

          writeFileSync(filePath, buf);
          saved.push({ file: filePath, name: att.name, bytes: buf.length });
        }

        out({ files: saved, count: saved.length });
      } catch (e: any) { err(e.message, "DOWNLOAD_FAILED"); }
    });
}
