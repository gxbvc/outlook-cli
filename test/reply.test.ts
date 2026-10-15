import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";

const { setTransportForTests } = await import("../src/graph.js");
const { registerEmails } = await import("../src/commands/emails.js");
const { AGENT_FOOTER } = await import("../src/footer.js");

type Call = { path: string; method: string; body: any };
let calls: Call[] = [];
let draft: any;
let lines: string[] = [];
const realLog = console.log;

function recipients(...addrs: string[]) {
  return addrs.map((a) => ({ emailAddress: { address: a } }));
}

beforeEach(() => {
  calls = [];
  lines = [];
  delete process.env.OUTLOOK_USER;
  draft = {
    id: "DRAFT1",
    from: { emailAddress: { address: "christian@gxb.vc" } },
    toRecipients: recipients("trey@example.com"),
    ccRecipients: [],
    bccRecipients: [],
  };
  setTransportForTests(async (path, options) => {
    calls.push({
      path,
      method: String(options.method || "GET"),
      body: options.body ? JSON.parse(String(options.body)) : undefined,
    });
    if (/\/(createReply|createReplyAll|createForward)$/.test(path)) return draft;
    return null;
  });
  console.log = (s: string) => { lines.push(String(s)); };
});

after(() => {
  console.log = realLog;
  setTransportForTests(null);
});

async function run(...args: string[]) {
  const program = new Command();
  program.exitOverride();
  registerEmails(program);
  await program.parseAsync(["node", "outlook-cli", "emails", ...args]);
}

const addrsOf = (list: any[]) => list.map((r) => r.emailAddress.address);
const patch = () => calls.find((c) => c.method === "PATCH")!;
const create = () => calls.find((c) => /\/createReply(All)?$/.test(c.path))!;
const count = (s: string, sub: string) => s.split(sub).length - 1;

test("reply without recipient flags uses plain /reply once, with footer", async () => {
  await run("reply", "--id", "MSG1", "--text", "Thanks");
  assert.equal(calls.length, 1);
  assert.equal(calls[0].path, "/me/messages/MSG1/reply");
  assert.equal(count(calls[0].body.comment, AGENT_FOOTER), 1);
});

test("reply --reply-all without flags uses /replyAll", async () => {
  await run("reply", "--id", "MSG1", "--text", "Thanks", "--reply-all");
  assert.equal(calls.length, 1);
  assert.equal(calls[0].path, "/me/messages/MSG1/replyAll");
});

test("reply --cc: createReply, PATCH recipients, send; sender kept, cc added", async () => {
  await run("reply", "--id", "MSG1", "--text", "Thanks", "--cc", "dan@example.com");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), [
    "POST /me/messages/MSG1/createReply",
    "PATCH /me/messages/DRAFT1",
    "POST /me/messages/DRAFT1/send",
  ]);
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["trey@example.com"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["dan@example.com"]);
  assert.deepEqual(patch().body.bccRecipients, []);
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
  assert.deepEqual(JSON.parse(lines[0]), { ok: true, data: { message: "Reply sent" } });
});

test("reply --cc repeated and comma-separated", async () => {
  await run("reply", "--id", "MSG1", "--html", "<p>Hi</p>", "--cc", "a@x.com,b@x.com", "--cc", "c@x.com");
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["a@x.com", "b@x.com", "c@x.com"]);
  assert.equal(create().body.comment.startsWith("<p>Hi</p>"), true);
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
});

test("reply --reply-all --cc dedupes against existing recipients, case-insensitively", async () => {
  draft.toRecipients = recipients("Trey@Example.com");
  draft.ccRecipients = recipients("ricky@gxb.vc");
  await run(
    "reply", "--id", "MSG1", "--text", "x", "--reply-all",
    "--cc", "TREY@example.com,dan@example.com,RICKY@GXB.VC,dan@EXAMPLE.com",
  );
  assert.equal(create().path, "/me/messages/MSG1/createReplyAll");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["Trey@Example.com"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["ricky@gxb.vc", "dan@example.com"]);
});

test("reply never adds the sending mailbox to itself", async () => {
  await run("reply", "--id", "MSG1", "--text", "x", "--reply-all", "--cc", "Christian@GXB.vc,dan@example.com");
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["dan@example.com"]);
});

test("reply never drops an existing recipient, even the sending mailbox", async () => {
  draft.toRecipients = recipients("christian@gxb.vc");
  await run("reply", "--id", "MSG1", "--text", "x", "--cc", "christian.genco@gmail.com");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["christian@gxb.vc"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["christian.genco@gmail.com"]);
});

test("reply --to and --bcc are added without --reply-all", async () => {
  await run("reply", "--id", "MSG1", "--text", "x", "--to", "extra@x.com", "--bcc", "hidden@x.com");
  assert.equal(create().path, "/me/messages/MSG1/createReply");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["trey@example.com", "extra@x.com"]);
  assert.deepEqual(patch().body.ccRecipients, []);
  assert.deepEqual(addrsOf(patch().body.bccRecipients), ["hidden@x.com"]);
});

test("reply footer is not doubled when the text already has it", async () => {
  await run("reply", "--id", "MSG1", "--text", `Hi\n\n${AGENT_FOOTER}`, "--cc", "dan@example.com");
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
});

test("reply with --user maps every call to the shared mailbox", async () => {
  process.env.OUTLOOK_USER = "shared@gxb.vc";
  await run("reply", "--id", "MSG1", "--text", "x", "--cc", "dan@example.com");
  for (const c of calls) assert.match(c.path, /^\/users\/shared%40gxb\.vc\/messages\//);
});

test("reply with an invalid address sends nothing", async () => {
  const exit = process.exit;
  (process as any).exit = () => { throw new Error("exit"); };
  try {
    await assert.rejects(run("reply", "--id", "MSG1", "--text", "x", "--cc", "not an address"));
  } finally {
    process.exit = exit;
  }
  assert.equal(calls.length, 0);
  assert.match(lines[0], /Invalid email address/);
});

test("reply deletes the draft if PATCH fails", async () => {
  setTransportForTests(async (path, options) => {
    calls.push({ path, method: String(options.method || "GET"), body: undefined });
    if (path.endsWith("/createReply")) return draft;
    if (options.method === "PATCH") throw new Error("boom");
    return null;
  });
  const exit = process.exit;
  (process as any).exit = () => { throw new Error("exit"); };
  try {
    await assert.rejects(run("reply", "--id", "MSG1", "--text", "x", "--cc", "dan@example.com"));
  } finally {
    process.exit = exit;
  }
  assert.ok(calls.some((c) => c.method === "DELETE" && c.path === "/me/messages/DRAFT1"));
  assert.ok(!calls.some((c) => c.path.endsWith("/send")));
});

test("forward with --cc and --bcc, repeated and comma-separated", async () => {
  draft = { id: "FWD1" };
  await run("forward", "--id", "MSG1", "--to", "a@x.com,b@x.com", "--cc", "c@x.com", "--cc", "d@x.com", "--bcc", "e@x.com", "--html", "<p>FYI</p>");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), [
    "POST /me/messages/MSG1/createForward",
    "POST /me/messages/FWD1/send",
  ]);
  const msg = calls[0].body.message;
  assert.deepEqual(addrsOf(msg.toRecipients), ["a@x.com", "b@x.com"]);
  assert.deepEqual(addrsOf(msg.ccRecipients), ["c@x.com", "d@x.com"]);
  assert.deepEqual(addrsOf(msg.bccRecipients), ["e@x.com"]);
  assert.equal(count(calls[0].body.comment, AGENT_FOOTER), 1);
});

test("forward drops a duplicate address in different case", async () => {
  draft = { id: "FWD1" };
  await run("forward", "--id", "MSG1", "--to", "A@x.com", "--cc", "a@X.com,c@x.com");
  const msg = calls[0].body.message;
  assert.deepEqual(addrsOf(msg.toRecipients), ["A@x.com"]);
  assert.deepEqual(addrsOf(msg.ccRecipients), ["c@x.com"]);
  assert.equal(msg.bccRecipients, undefined);
});

test("forward without --to fails and sends nothing", async () => {
  const exit = process.exit;
  (process as any).exit = () => { throw new Error("exit"); };
  try {
    await assert.rejects(run("forward", "--id", "MSG1", "--cc", "c@x.com"));
  } finally {
    process.exit = exit;
  }
  assert.equal(calls.length, 0);
});
