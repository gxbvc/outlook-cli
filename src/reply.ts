import { graphRequest } from "./graph.js";
import { attachAllToDraft } from "./attachments.js";
import { footedComment } from "./footer.js";

export interface Recipient {
  emailAddress: { address: string; name?: string };
}

// Option parser for repeatable, comma-separated address flags:
// `--cc a@x.com,b@x.com --cc c@x.com` gives three addresses.
export const collectAddrs = (v: string, acc: string[]) => { acc.push(v); return acc; };

export function parseAddrs(values?: string | string[]): string[] {
  const list = values === undefined ? [] : Array.isArray(values) ? values : [values];
  const out: string[] = [];
  for (const v of list) {
    for (const part of v.split(",")) {
      const a = part.trim();
      if (!a) continue;
      if (!/^[^@\s<>;]+@[^@\s<>;]+$/.test(a)) throw new Error(`Invalid email address: ${a}`);
      out.push(a);
    }
  }
  return out;
}

const key = (a: string) => a.trim().toLowerCase();

// Merge added addresses into existing recipients. An address that is already a
// recipient anywhere (to, cc or bcc) is not added again, compared case-insensitively.
// Existing recipients are never dropped. `self` (the sending mailbox) is never added.
export function mergeRecipients(
  existing: { to: Recipient[]; cc: Recipient[]; bcc: Recipient[] },
  added: { to: string[]; cc: string[]; bcc: string[] },
  self?: string | null,
): { to: Recipient[]; cc: Recipient[]; bcc: Recipient[] } {
  const seen = new Set<string>();
  for (const r of [...existing.to, ...existing.cc, ...existing.bcc]) {
    if (r?.emailAddress?.address) seen.add(key(r.emailAddress.address));
  }
  const selfKey = self ? key(self) : null;
  const add = (list: Recipient[], addrs: string[]): Recipient[] => {
    const next = [...list];
    for (const a of addrs) {
      const k = key(a);
      if (seen.has(k) || k === selfKey) continue;
      seen.add(k);
      next.push({ emailAddress: { address: a } });
    }
    return next;
  };
  return {
    to: add(existing.to, added.to),
    cc: add(existing.cc, added.cc),
    bcc: add(existing.bcc, added.bcc),
  };
}

export interface ReplyOpts {
  id: string;
  text?: string;
  html?: string;
  replyAll?: boolean;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  attach?: string[];
}

export function hasAddedRecipients(opts: { to?: string[]; cc?: string[]; bcc?: string[] }): boolean {
  return parseAddrs(opts.to).length + parseAddrs(opts.cc).length + parseAddrs(opts.bcc).length > 0;
}

async function mailboxAddress(draft: any): Promise<string | null> {
  const fromDraft = draft?.from?.emailAddress?.address;
  if (fromDraft) return fromDraft;
  if (process.env.OUTLOOK_USER) return process.env.OUTLOOK_USER;
  try {
    const me = await graphRequest("/me?$select=mail,userPrincipalName");
    return me?.mail || me?.userPrincipalName || null;
  } catch {
    return null;
  }
}

// Reply that adds recipients and stays in the thread:
// createReply(All) with the comment (Graph puts it above the quoted history),
// PATCH the recipients, attach files, send. The agent footer comes from footedComment, once.
export async function replyWithRecipients(opts: ReplyOpts) {
  const added = {
    to: parseAddrs(opts.to),
    cc: parseAddrs(opts.cc),
    bcc: parseAddrs(opts.bcc),
  };
  const endpoint = opts.replyAll ? "createReplyAll" : "createReply";
  const draft = await graphRequest(`/me/messages/${opts.id}/${endpoint}`, {
    method: "POST",
    body: JSON.stringify({ comment: footedComment(opts) }),
  });
  try {
    const self = await mailboxAddress(draft);
    const merged = mergeRecipients(
      {
        to: draft.toRecipients || [],
        cc: draft.ccRecipients || [],
        bcc: draft.bccRecipients || [],
      },
      added,
      self,
    );
    await graphRequest(`/me/messages/${draft.id}`, {
      method: "PATCH",
      body: JSON.stringify({
        toRecipients: merged.to,
        ccRecipients: merged.cc,
        bccRecipients: merged.bcc,
      }),
    });
    const attached = opts.attach?.length ? await attachAllToDraft(draft.id, opts.attach) : [];
    await graphRequest(`/me/messages/${draft.id}/send`, { method: "POST" });
    return { attached, recipients: merged };
  } catch (e) {
    // Do not leave a half-built reply in Drafts.
    try { await graphRequest(`/me/messages/${draft.id}`, { method: "DELETE" }); } catch { /* ignore */ }
    throw e;
  }
}

export interface ForwardOpts {
  id: string;
  text?: string;
  html?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  attach?: string[];
}

// createForward(comment + recipients) → attach extra files → send.
// The original's attachments come along; the footer comes from footedComment, once.
export async function forwardMessage(opts: ForwardOpts) {
  const added = { to: parseAddrs(opts.to), cc: parseAddrs(opts.cc), bcc: parseAddrs(opts.bcc) };
  if (added.to.length === 0) throw new Error("forward needs at least one --to address");
  const merged = mergeRecipients({ to: [], cc: [], bcc: [] }, added, null);
  const draft = await graphRequest(`/me/messages/${opts.id}/createForward`, {
    method: "POST",
    body: JSON.stringify({
      comment: footedComment(opts),
      message: {
        toRecipients: merged.to,
        ccRecipients: merged.cc.length ? merged.cc : undefined,
        bccRecipients: merged.bcc.length ? merged.bcc : undefined,
      },
    }),
  });
  const attached = opts.attach?.length ? await attachAllToDraft(draft.id, opts.attach) : [];
  await graphRequest(`/me/messages/${draft.id}/send`, { method: "POST" });
  return { attached, recipients: merged };
}
