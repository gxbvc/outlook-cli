import { Command } from "commander";
import { graphRequest, out, err, owaUrl } from "../graph.js";
import {
  loadAll,
  hasLarge,
  toFileAttachmentJson,
  attachAllToDraft,
  INLINE_LIMIT_BYTES,
} from "../attachments.js";

const collect = (v: string, acc: string[]) => { acc.push(v); return acc; };

export function registerDrafts(program: Command) {
  const drafts = program.command("drafts").description("Manage drafts");

  drafts
    .command("list")
    .description("List drafts")
    .option("--limit <n>", "Max results", "25")
    .option("--page-token <token>", "Pagination token")
    .action(async (opts) => {
      try {
        let url: string;
        if (opts.pageToken) {
          url = opts.pageToken;
        } else {
          const params = new URLSearchParams({ "$top": opts.limit });
          params.append("$select", "id,subject,toRecipients,createdDateTime,bodyPreview");
          params.append("$orderby", "createdDateTime desc");
          url = `/me/mailFolders/drafts/messages?${params.toString()}`;
        }
        const data = await graphRequest(url);
        out({
          drafts: data.value,
          nextPage: data["@odata.nextLink"] || null,
          count: data.value.length,
        });
      } catch (e: any) { err(e.message, "LIST_FAILED"); }
    });

  drafts
    .command("create")
    .description("Create a new draft")
    .requiredOption("--to <addresses>", "Comma-separated recipients")
    .requiredOption("--subject <subject>", "Subject line")
    .option("--text <text>", "Plain text body")
    .option("--html <html>", "HTML body")
    .option("--cc <addresses>", "Comma-separated CC")
    .option("--bcc <addresses>", "Comma-separated BCC")
    .option("--attach <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const body: any = {
          subject: opts.subject,
          body: {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.html || opts.text || "",
          },
          toRecipients: opts.to.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          })),
        };
        if (opts.cc) {
          body.ccRecipients = opts.cc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        if (opts.bcc) {
          body.bccRecipients = opts.bcc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }

        const attachPaths: string[] = opts.attach || [];
        let loaded: ReturnType<typeof loadAll> = [];
        if (attachPaths.length > 0) {
          loaded = loadAll(attachPaths);
          const smallInline = loaded.filter((a) => a.size <= INLINE_LIMIT_BYTES);
          if (smallInline.length > 0) {
            body.attachments = smallInline.map(toFileAttachmentJson);
          }
        }

        const data = await graphRequest("/me/messages", {
          method: "POST",
          body: JSON.stringify(body),
        });

        // Upload any large attachments after draft creation
        if (attachPaths.length > 0 && hasLarge(attachPaths)) {
          const largePaths = loaded.filter((a) => a.size > INLINE_LIMIT_BYTES).map((a) => a.path);
          await attachAllToDraft(data.id, largePaths);
        }

        const link = await owaUrl(data.id);
        if (attachPaths.length > 0) {
          out({
            ...data,
            owaUrl: link,
            attachments: loaded.map((a) => ({ name: a.name, size: a.size })),
          });
        } else {
          out({ ...data, owaUrl: link });
        }
      } catch (e: any) { err(e.message, "CREATE_FAILED"); }
    });

  drafts
    .command("url")
    .description("Print the OWA browser URL for a draft (open it with: open \"$(...)\")")
    .requiredOption("--id <id>", "Draft ID")
    .option("--folder <folder>", "OWA folder segment", "drafts")
    .action(async (opts) => {
      try {
        const link = await owaUrl(opts.id, opts.folder);
        if (!link) err("Could not translate that ID to an OWA URL", "NO_URL");
        out({ id: opts.id, owaUrl: link });
      } catch (e: any) { err(e.message, "URL_FAILED"); }
    });

  drafts
    .command("reply")
    .description("Create a reply draft")
    .requiredOption("--id <id>", "Message ID to reply to")
    .option("--text <text>", "Plain text body")
    .option("--html <html>", "HTML body")
    .option("--reply-all", "Reply to all")
    .option("--attach <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const endpoint = opts.replyAll ? "createReplyAll" : "createReply";
        const body: any = {};
        if (opts.text || opts.html) {
          body.comment = opts.html || opts.text;
        }
        const data = await graphRequest(`/me/messages/${opts.id}/${endpoint}`, {
          method: "POST",
          body: JSON.stringify(body),
        });

        const attachPaths: string[] = opts.attach || [];
        if (attachPaths.length > 0) {
          const attached = await attachAllToDraft(data.id, attachPaths);
          out({ ...data, attachments: attached.map((a) => ({ name: a.name, size: a.size })) });
          return;
        }
        out(data);
      } catch (e: any) { err(e.message, "REPLY_DRAFT_FAILED"); }
    });

  drafts
    .command("update")
    .description("Update a draft")
    .requiredOption("--id <id>", "Draft message ID")
    .option("--to <addresses>", "Comma-separated recipients")
    .option("--subject <subject>", "Subject line")
    .option("--text <text>", "Plain text body")
    .option("--html <html>", "HTML body")
    .option("--cc <addresses>", "Comma-separated CC")
    .option("--bcc <addresses>", "Comma-separated BCC")
    .option("--attach <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const body: any = {};
        if (opts.subject) body.subject = opts.subject;
        if (opts.text || opts.html) {
          body.body = {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.html || opts.text,
          };
        }
        if (opts.to) {
          body.toRecipients = opts.to.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        if (opts.cc) {
          body.ccRecipients = opts.cc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        if (opts.bcc) {
          body.bccRecipients = opts.bcc.split(",").map((e: string) => ({
            emailAddress: { address: e.trim() },
          }));
        }
        const data = await graphRequest(`/me/messages/${opts.id}`, {
          method: "PATCH",
          body: JSON.stringify(body),
        });

        const attachPaths: string[] = opts.attach || [];
        if (attachPaths.length > 0) {
          const attached = await attachAllToDraft(opts.id, attachPaths);
          out({ ...data, attachments: attached.map((a) => ({ name: a.name, size: a.size, id: a.id })) });
          return;
        }
        out(data);
      } catch (e: any) { err(e.message, "UPDATE_FAILED"); }
    });

  drafts
    .command("send")
    .description("Send a draft")
    .requiredOption("--id <id>", "Draft message ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}/send`, { method: "POST" });
        out({ message: "Draft sent" });
      } catch (e: any) { err(e.message, "SEND_FAILED"); }
    });

  drafts
    .command("delete")
    .description("Delete a draft")
    .requiredOption("--id <id>", "Draft message ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}`, { method: "DELETE" });
        out({ message: "Draft deleted" });
      } catch (e: any) { err(e.message, "DELETE_FAILED"); }
    });

  drafts
    .command("attach")
    .description("Attach file(s) to a draft")
    .requiredOption("--id <id>", "Draft message ID")
    .requiredOption("--file <path>", "File to attach (repeatable)", collect, [])
    .action(async (opts) => {
      try {
        const paths: string[] = opts.file || [];
        if (paths.length === 0) {
          err("At least one --file is required", "MISSING_FILE");
        }
        const attached = await attachAllToDraft(opts.id, paths);
        out({ attached });
      } catch (e: any) { err(e.message, "ATTACH_FAILED"); }
    });

  drafts
    .command("detach")
    .description("Remove an attachment from a draft")
    .requiredOption("--id <id>", "Draft message ID")
    .requiredOption("--attachment-id <attId>", "Attachment ID to remove")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/messages/${opts.id}/attachments/${opts.attachmentId}`, {
          method: "DELETE",
        });
        out({ message: "Attachment removed" });
      } catch (e: any) { err(e.message, "DETACH_FAILED"); }
    });
}
