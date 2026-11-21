import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";
// Never write the real send log from tests.
process.env.OUTLOOK_CLI_SEND_LOG = join(mkdtempSync(join(tmpdir(), "olsl-")), "sends.jsonl");

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
    from: { emailAddress: { address: "owner@corp.example.test" } },
    toRecipients: recipients("sender@example.test"),
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
  await run("reply", "--id", "MSG1", "--text", "Thanks", "--cc", "added@example.test");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), [
    "POST /me/messages/MSG1/createReply",
    "PATCH /me/messages/DRAFT1",
    "POST /me/messages/DRAFT1/send",
  ]);
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["sender@example.test"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["added@example.test"]);
  assert.deepEqual(patch().body.bccRecipients, []);
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
  assert.deepEqual(JSON.parse(lines[0]), { ok: true, data: { message: "Reply sent" } });
});

test("reply --cc repeated and comma-separated", async () => {
  await run("reply", "--id", "MSG1", "--html", "<p>Hi</p>", "--cc", "a@x.example.test,b@x.example.test", "--cc", "c@x.example.test");
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["a@x.example.test", "b@x.example.test", "c@x.example.test"]);
  assert.equal(create().body.comment.startsWith("<p>Hi</p>"), true);
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
});

test("reply --reply-all --cc dedupes against existing recipients, case-insensitively", async () => {
  draft.toRecipients = recipients("Sender@Example.test");
  draft.ccRecipients = recipients("partner@corp.example.test");
  await run(
    "reply", "--id", "MSG1", "--text", "x", "--reply-all",
    "--cc", "SENDER@example.test,added@example.test,PARTNER@CORP.EXAMPLE.TEST,added@EXAMPLE.test",
  );
  assert.equal(create().path, "/me/messages/MSG1/createReplyAll");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["Sender@Example.test"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["partner@corp.example.test", "added@example.test"]);
});

test("reply never adds the sending mailbox to itself", async () => {
  await run("reply", "--id", "MSG1", "--text", "x", "--reply-all", "--cc", "Owner@CORP.example.test,added@example.test");
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["added@example.test"]);
});

test("reply never drops an existing recipient, even the sending mailbox", async () => {
  draft.toRecipients = recipients("owner@corp.example.test");
  await run("reply", "--id", "MSG1", "--text", "x", "--cc", "owner.personal@mail.example.test");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["owner@corp.example.test"]);
  assert.deepEqual(addrsOf(patch().body.ccRecipients), ["owner.personal@mail.example.test"]);
});

test("reply --to and --bcc are added without --reply-all", async () => {
  await run("reply", "--id", "MSG1", "--text", "x", "--to", "extra@x.example.test", "--bcc", "hidden@x.example.test");
  assert.equal(create().path, "/me/messages/MSG1/createReply");
  assert.deepEqual(addrsOf(patch().body.toRecipients), ["sender@example.test", "extra@x.example.test"]);
  assert.deepEqual(patch().body.ccRecipients, []);
  assert.deepEqual(addrsOf(patch().body.bccRecipients), ["hidden@x.example.test"]);
});

test("reply footer is not doubled when the text already has it", async () => {
  await run("reply", "--id", "MSG1", "--text", `Hi\n\n${AGENT_FOOTER}`, "--cc", "added@example.test");
  assert.equal(count(create().body.comment, AGENT_FOOTER), 1);
});

test("reply with --user maps every call to the shared mailbox", async () => {
  process.env.OUTLOOK_USER = "shared@corp.example.test";
  await run("reply", "--id", "MSG1", "--text", "x", "--cc", "added@example.test");
  for (const c of calls) assert.match(c.path, /^\/users\/shared%40corp\.example\.test\/messages\//);
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
    await assert.rejects(run("reply", "--id", "MSG1", "--text", "x", "--cc", "added@example.test"));
  } finally {
    process.exit = exit;
  }
  assert.ok(calls.some((c) => c.method === "DELETE" && c.path === "/me/messages/DRAFT1"));
  assert.ok(!calls.some((c) => c.path.endsWith("/send")));
});

test("forward with --cc and --bcc, repeated and comma-separated", async () => {
  draft = { id: "FWD1" };
  await run("forward", "--id", "MSG1", "--to", "a@x.example.test,b@x.example.test", "--cc", "c@x.example.test", "--cc", "d@x.example.test", "--bcc", "e@x.example.test", "--html", "<p>FYI</p>");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), [
    "POST /me/messages/MSG1/createForward",
    "POST /me/messages/FWD1/send",
  ]);
  const msg = calls[0].body.message;
  assert.deepEqual(addrsOf(msg.toRecipients), ["a@x.example.test", "b@x.example.test"]);
  assert.deepEqual(addrsOf(msg.ccRecipients), ["c@x.example.test", "d@x.example.test"]);
  assert.deepEqual(addrsOf(msg.bccRecipients), ["e@x.example.test"]);
  assert.equal(count(calls[0].body.comment, AGENT_FOOTER), 1);
});

test("forward drops a duplicate address in different case", async () => {
  draft = { id: "FWD1" };
  await run("forward", "--id", "MSG1", "--to", "A@x.example.test", "--cc", "a@X.example.test,c@x.example.test");
  const msg = calls[0].body.message;
  assert.deepEqual(addrsOf(msg.toRecipients), ["A@x.example.test"]);
  assert.deepEqual(addrsOf(msg.ccRecipients), ["c@x.example.test"]);
  assert.equal(msg.bccRecipients, undefined);
});

test("forward without --to fails and sends nothing", async () => {
  const exit = process.exit;
  (process as any).exit = () => { throw new Error("exit"); };
  try {
    await assert.rejects(run("forward", "--id", "MSG1", "--cc", "c@x.example.test"));
  } finally {
    process.exit = exit;
  }
  assert.equal(calls.length, 0);
});
