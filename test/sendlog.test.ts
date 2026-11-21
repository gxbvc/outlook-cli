import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";
import {
  chmodSync, mkdirSync, mkdtempSync, readFileSync, statSync, symlinkSync, writeFileSync, existsSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";

const { setTransportForTests, GraphError } = await import("../src/graph.js");
const { registerEmails } = await import("../src/commands/emails.js");
const { registerDrafts } = await import("../src/commands/drafts.js");
const { registerCalendar } = await import("../src/commands/calendar.js");

// Strings that must never reach the log file.
const SUBJECT = "Quokka-Lantern subject 7731";
const BODY = "Marmalade-Harbor body 4419";
const ATT_NAME = "Pelican-Ledger-attach.txt";

type Call = { path: string; method: string };
let calls: Call[] = [];
let lines: string[] = [];
let logPath = "";
let failSend: number | null = null;
const realLog = console.log;
const realExit = process.exit;
const realFetch = globalThis.fetch;

const sha8 = (s: string) => createHash("sha256").update(s, "utf8").digest("hex").slice(0, 8);
const rcpt = (...a: string[]) => a.map((address) => ({ emailAddress: { address } }));
const isSend = (c: Call) => {
  const p = c.path.replace(/^\/users\/[^/]+/, "/me");
  return /\/(sendMail|send|reply|replyAll|forward|cancel)$/.test(p) ||
    (c.method === "POST" && p === "/me/events") ||
    (p.startsWith("/me/events/") && (c.method === "PATCH" || c.method === "DELETE"));
};

function freshLog(): string {
  const dir = join(mkdtempSync(join(tmpdir(), "olsl-test-")), "state");
  return join(dir, "sends.jsonl");
}

beforeEach(() => {
  calls = [];
  lines = [];
  failSend = null;
  logPath = freshLog();
  process.env.OUTLOOK_CLI_SEND_LOG = logPath;
  process.env.OUTLOOK_USER = "Agent@Example.test";
  for (const k of ["HERDR_TAB_ID", "HERDR_PANE_ID", "HERDR_WORKSPACE_ID"]) delete process.env[k];
  setTransportForTests(async (path, options) => {
    const method = String(options.method || "GET");
    calls.push({ path, method });
    if (failSend && /\/send$/.test(path)) throw new GraphError(failSend, `Graph API ${failSend}: no`);
    if (/\/createReply(All)?$/.test(path)) {
      return { id: "RDRAFT", subject: `RE: ${SUBJECT}`, toRecipients: rcpt("Sender@Example.test"), ccRecipients: rcpt("cc1@example.test"), bccRecipients: [] };
    }
    if (/\/createForward$/.test(path)) {
      return { id: "FDRAFT", subject: `FW: ${SUBJECT}`, toRecipients: rcpt("fwd@example.test"), ccRecipients: [], bccRecipients: [] };
    }
    if (method === "GET" && path.startsWith("/users/") && path.includes("/messages/D1?")) {
      return { subject: SUBJECT, toRecipients: rcpt("To1@Example.test"), ccRecipients: [], bccRecipients: rcpt("hidden@example.test") };
    }
    if (method === "GET" && path.includes("$select=organizer,attendees")) {
      return { organizer: { emailAddress: { address: "agent@example.test" } }, attendees: [{ emailAddress: { address: "old@example.test" }, type: "required" }] };
    }
    if (method === "POST" && /\/messages$/.test(path)) return { id: "BIGDRAFT" };
    if (/createUploadSession$/.test(path)) return { uploadUrl: "https://upload.example.test/session" };
    if (/\/attachments$/.test(path)) return { id: "ATT1" };
    return { id: "X1" };
  });
  console.log = (s: string) => { lines.push(String(s)); };
  (process as any).exit = (code?: number) => { throw new Error(`exit ${code}`); };
});

after(() => {
  console.log = realLog;
  process.exit = realExit;
  globalThis.fetch = realFetch;
  setTransportForTests(null);
  delete process.env.OUTLOOK_USER;
});

async function run(...args: string[]) {
  const program = new Command();
  program.exitOverride();
  registerEmails(program);
  registerDrafts(program);
  registerCalendar(program);
  await program.parseAsync(["node", "outlook-cli", ...args]);
}

function logLines(path = logPath): any[] {
  return readFileSync(path, "utf8").trim().split("\n").map((l) => JSON.parse(l));
}

function assertPair(kind: string, extra: (attempt: any) => void = () => {}) {
  const recs = logLines();
  assert.equal(recs.length, 2, "one attempt line and one result line");
  const [a, r] = recs;
  assert.equal(a.phase, "attempt");
  assert.equal(r.phase, "result");
  assert.equal(a.kind, kind);
  assert.equal(r.kind, kind);
  assert.equal(a.send_id, r.send_id);
  assert.equal(r.ok, true);
  assert.equal(a.from, "agent@example.test");
  assert.match(a.ts, /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/);
  assert.equal(typeof a.caller.cwd, "string");
  if (a.subject_sha8 === null) assert.equal(typeof a.subject_note, "string");
  extra(a);
  assertNoSecrets();
}

function assertNoSecrets(path = logPath) {
  const raw = readFileSync(path, "utf8");
  for (const s of [SUBJECT, BODY, ATT_NAME, "Quokka", "Marmalade", "Pelican"]) {
    assert.ok(!raw.includes(s), `log must not contain ${s}`);
  }
}

// ---- every send path writes attempt + result ----

test("emails send (sendMail): attempt + result, lower-case addresses, subject hash", async () => {
  await run("emails", "send", "--to", "Bob@Example.test, amy@example.test", "--cc", "Cc@Example.test", "--bcc", "B@example.test", "--subject", SUBJECT, "--text", BODY);
  assert.ok(calls.some((c) => c.path.endsWith("/sendMail")));
  assertPair("send", (a) => {
    assert.deepEqual(a.to, ["bob@example.test", "amy@example.test"]);
    assert.deepEqual(a.cc, ["cc@example.test"]);
    assert.deepEqual(a.bcc, ["b@example.test"]);
    assert.equal(a.subject_sha8, sha8(SUBJECT));
  });
  assert.equal(statSync(logPath).mode & 0o777, 0o600);
  assert.equal(statSync(join(logPath, "..")).mode & 0o777, 0o700);
});

test("emails send with a large attachment (draft + /send): attempt + result, no attachment name", async () => {
  const dir = mkdtempSync(join(tmpdir(), "olsl-att-"));
  const file = join(dir, ATT_NAME);
  writeFileSync(file, Buffer.alloc(3 * 1024 * 1024 + 10, 65));
  globalThis.fetch = (async () => new Response(null, { status: 201 })) as any;
  try {
    await run("emails", "send", "--to", "bob@example.test", "--subject", SUBJECT, "--text", BODY, "--attach", file);
  } finally {
    globalThis.fetch = realFetch;
  }
  assert.ok(calls.some((c) => c.path.endsWith("/BIGDRAFT/send")));
  assertPair("send", (a) => {
    assert.equal(a.message_id, "BIGDRAFT");
    assert.equal(a.subject_sha8, sha8(SUBJECT));
  });
});

test("emails reply (plain /reply): subject unknown, null with a note", async () => {
  await run("emails", "reply", "--id", "MSG1", "--text", BODY);
  assertPair("reply", (a) => {
    assert.equal(a.message_id, "MSG1");
    assert.equal(a.subject_sha8, null);
    assert.match(a.subject_note, /Graph sets the reply subject/);
    assert.equal(a.to, null);
  });
});

test("emails reply --reply-all (plain /replyAll)", async () => {
  await run("emails", "reply", "--id", "MSG1", "--text", BODY, "--reply-all");
  assertPair("reply_all");
});

test("emails reply with attachment (createReply + /send): subject from the draft", async () => {
  const dir = mkdtempSync(join(tmpdir(), "olsl-att-"));
  const file = join(dir, ATT_NAME);
  writeFileSync(file, "small");
  await run("emails", "reply", "--id", "MSG1", "--text", BODY, "--attach", file);
  assertPair("reply", (a) => {
    assert.equal(a.subject_sha8, sha8(`RE: ${SUBJECT}`));
    assert.deepEqual(a.to, ["sender@example.test"]);
    assert.deepEqual(a.cc, ["cc1@example.test"]);
  });
});

test("emails reply --cc (createReply + PATCH + /send)", async () => {
  await run("emails", "reply", "--id", "MSG1", "--text", BODY, "--cc", "Added@Example.test");
  assertPair("reply", (a) => {
    assert.deepEqual(a.cc, ["cc1@example.test", "added@example.test"]);
    assert.equal(a.subject_sha8, sha8(`RE: ${SUBJECT}`));
  });
});

test("emails forward (createForward + /send)", async () => {
  await run("emails", "forward", "--id", "MSG1", "--to", "fwd@example.test", "--text", BODY);
  assertPair("forward", (a) => {
    assert.deepEqual(a.to, ["fwd@example.test"]);
    assert.equal(a.subject_sha8, sha8(`FW: ${SUBJECT}`));
  });
});

test("drafts send: recipients and subject hash from the draft", async () => {
  await run("drafts", "send", "--id", "D1");
  assertPair("draft_send", (a) => {
    assert.deepEqual(a.to, ["to1@example.test"]);
    assert.deepEqual(a.bcc, ["hidden@example.test"]);
    assert.equal(a.subject_sha8, sha8(SUBJECT));
    assert.equal(a.message_id, "D1");
  });
});

test("drafts send --at (scheduled)", async () => {
  await run("drafts", "send", "--id", "D1", "--at", new Date(Date.now() + 86_400_000).toISOString());
  assertPair("draft_send_scheduled");
});

test("calendar create with attendees: invites logged; without attendees: no log", async () => {
  await run("calendar", "create", "--subject", SUBJECT, "--start", "2030-01-01T09:00:00", "--end", "2030-01-01T10:00:00",
    "--body", BODY, "--attendees", "Req@Example.test", "--optional-attendees", "opt@example.test", "--resources", "room@example.test");
  assertPair("event_create", (a) => {
    assert.deepEqual(a.to, ["req@example.test", "room@example.test"]);
    assert.deepEqual(a.cc, ["opt@example.test"]);
    assert.equal(a.subject_sha8, sha8(SUBJECT));
  });
  logPath = freshLog();
  process.env.OUTLOOK_CLI_SEND_LOG = logPath;
  await run("calendar", "create", "--subject", SUBJECT, "--start", "2030-01-01T09:00:00", "--end", "2030-01-01T10:00:00");
  assert.equal(existsSync(logPath), false);
});

test("calendar update: logged with the attendee list after the update", async () => {
  await run("calendar", "update", "--id", "EVT1", "--subject", SUBJECT, "--attendees", "new@example.test");
  assertPair("event_update", (a) => {
    assert.deepEqual(a.to, ["old@example.test", "new@example.test"]);
    assert.equal(a.event_id, "EVT1");
    assert.equal(a.subject_sha8, sha8(SUBJECT));
  });
});

test("calendar cancel, delete and forward: each logged", async () => {
  await run("calendar", "cancel", "--id", "EVT1", "--comment", BODY);
  assertPair("event_cancel", (a) => assert.equal(a.event_id, "EVT1"));
  for (const [args, kind] of [
    [["calendar", "delete", "--id", "EVT1"], "event_delete"],
    [["calendar", "forward", "--id", "EVT1", "--to", "Fw@Example.test", "--comment", BODY], "event_forward"],
  ] as [string[], string][]) {
    logPath = freshLog();
    process.env.OUTLOOK_CLI_SEND_LOG = logPath;
    await run(...args);
    assertPair(kind);
  }
  assert.deepEqual(logLines()[0].to, ["fw@example.test"]);
});

// ---- result line on failure ----

test("a failed Graph send writes a result line with the HTTP status", async () => {
  failSend = 503;
  await assert.rejects(run("drafts", "send", "--id", "D1"), /exit 1/);
  const [a, r] = logLines();
  assert.equal(a.phase, "attempt");
  assert.deepEqual([r.phase, r.ok, r.status], ["result", false, 503]);
});

// ---- caller ----

test("caller: herdr ids when set, else the parent process name", async () => {
  process.env.HERDR_TAB_ID = "tab-test";
  process.env.HERDR_PANE_ID = "pane-test";
  await run("emails", "send", "--to", "bob@example.test", "--subject", SUBJECT, "--text", BODY);
  const c1 = logLines()[0].caller;
  assert.equal(c1.herdr_tab_id, "tab-test");
  assert.equal(c1.herdr_pane_id, "pane-test");
  assert.equal(c1.parent_process, undefined);
  delete process.env.HERDR_TAB_ID;
  delete process.env.HERDR_PANE_ID;
  logPath = freshLog();
  process.env.OUTLOOK_CLI_SEND_LOG = logPath;
  await run("emails", "send", "--to", "bob@example.test", "--subject", SUBJECT, "--text", BODY);
  const c2 = logLines()[0].caller;
  // ps may time out on a loaded machine; then the field is null, but it is always present.
  assert.ok("parent_process" in c2);
  assert.ok(c2.parent_process === null || typeof c2.parent_process === "string");
});

// ---- fail closed ----

async function assertRefused(args: string[], why: RegExp) {
  await assert.rejects(run(...args), /exit 1/, "command must exit non-zero");
  assert.equal(calls.filter(isSend).length, 0, "the Graph send must not be called");
  const env = JSON.parse(lines[lines.length - 1]);
  assert.equal(env.ok, false);
  assert.equal(env.code, "SEND_LOG_FAILED");
  assert.match(env.error, /Send refused/);
  assert.match(env.error, why);
}

const SEND = ["emails", "send", "--to", "bob@example.test", "--subject", SUBJECT, "--text", BODY];

test("refuses a symlinked log file: no send, target untouched", async () => {
  mkdirSync(join(logPath, ".."), { recursive: true, mode: 0o700 });
  const target = join(logPath, "..", "target.jsonl");
  writeFileSync(target, "", { mode: 0o600 });
  symlinkSync(target, logPath);
  await assertRefused(SEND, /symlink/);
  assert.equal(readFileSync(target, "utf8"), "");
});

test("refuses a 0644 log file: no send, file unchanged", async () => {
  mkdirSync(join(logPath, ".."), { recursive: true, mode: 0o700 });
  writeFileSync(logPath, "");
  chmodSync(logPath, 0o644);
  await assertRefused(SEND, /wider than 600/);
  assert.equal(readFileSync(logPath, "utf8"), "");
});

test("unwritable log (read-only folder): the Graph send is called 0 times, exit non-zero", async () => {
  const dir = join(logPath, "..");
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  chmodSync(dir, 0o500);
  try {
    await assertRefused(SEND, /cannot open/);
    // The same holds for every other send path.
    for (const args of [
      ["emails", "reply", "--id", "MSG1", "--text", BODY],
      ["emails", "forward", "--id", "MSG1", "--to", "fwd@example.test"],
      ["drafts", "send", "--id", "D1"],
      ["calendar", "cancel", "--id", "EVT1"],
      ["calendar", "forward", "--id", "EVT1", "--to", "x@example.test"],
    ]) {
      calls = [];
      await assertRefused(args, /cannot open/);
    }
  } finally {
    chmodSync(dir, 0o700);
  }
});

test("refuses a log folder wider than 0700 and a relative OUTLOOK_CLI_SEND_LOG", async () => {
  const dir = join(logPath, "..");
  mkdirSync(dir, { recursive: true });
  chmodSync(dir, 0o755);
  await assertRefused(SEND, /folder mode 755/);
  calls = [];
  process.env.OUTLOOK_CLI_SEND_LOG = "relative/sends.jsonl";
  await assertRefused(SEND, /absolute path/);
});

test("refused forward deletes its half-built draft and never calls /send", async () => {
  process.env.OUTLOOK_CLI_SEND_LOG = "relative.jsonl";
  await assertRefused(["emails", "forward", "--id", "MSG1", "--to", "fwd@example.test"], /absolute/);
  assert.ok(calls.some((c) => c.method === "DELETE" && c.path.endsWith("/messages/FDRAFT")));
});

// ---- append-only ----

test("append-only: a second send leaves the first lines unchanged", async () => {
  await run(...SEND);
  const first = readFileSync(logPath, "utf8");
  await run("emails", "reply", "--id", "MSG2", "--text", BODY);
  const both = readFileSync(logPath, "utf8");
  assert.ok(both.startsWith(first), "earlier bytes are never rewritten");
  const recs = logLines();
  assert.equal(recs.length, 4);
  assert.equal(JSON.stringify(recs[0]), first.split("\n")[0]);
  assert.equal(recs[2].kind, "reply");
  assertNoSecrets();
  if (process.env.SENDLOG_SAMPLE) realLog(first.split("\n")[0]);
});
