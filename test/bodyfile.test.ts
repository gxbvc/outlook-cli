import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, symlinkSync, writeFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";

const { setTransportForTests } = await import("../src/graph.js");
const { registerEmails } = await import("../src/commands/emails.js");
const { registerDrafts } = await import("../src/commands/drafts.js");
const { readBodyFile, applyBodyFiles, BodyFileError, MAX_BODY_FILE_BYTES } = await import("../src/bodyfile.js");
const { AGENT_FOOTER } = await import("../src/footer.js");

// Made-up text that must reach Graph's request body and never the send log or stdout.
const BODY = "Tangerine-Lighthouse body 5512\nsecond line\n";
const HTML = "<p>Tangerine-Lighthouse html 5512</p>";

type Call = { path: string; method: string; body: any };
let calls: Call[] = [];
let lines: string[] = [];
let logPath = "";
let dir = "";
const realLog = console.log;
const realExit = process.exit;

function bodyFile(name: string, content: string | Buffer, mode = 0o600): string {
  const p = join(dir, name);
  writeFileSync(p, content);
  chmodSync(p, mode);
  return p;
}

beforeEach(() => {
  calls = [];
  lines = [];
  dir = mkdtempSync(join(tmpdir(), "olbf-test-"));
  logPath = join(dir, "state", "sends.jsonl");
  process.env.OUTLOOK_CLI_SEND_LOG = logPath;
  process.env.OUTLOOK_USER = "agent@example.test";
  setTransportForTests(async (path, options) => {
    const method = String(options.method || "GET");
    calls.push({ path, method, body: options.body ? JSON.parse(String(options.body)) : null });
    if (/\/create(Reply|ReplyAll|Forward)$/.test(path)) return { id: "D1", subject: "RE: x", toRecipients: [], ccRecipients: [], bccRecipients: [] };
    return { id: "X1" };
  });
  console.log = (s: string) => { lines.push(String(s)); };
  (process as any).exit = (code?: number) => { throw new Error(`exit ${code}`); };
});

after(() => {
  console.log = realLog;
  process.exit = realExit;
  setTransportForTests(null);
  delete process.env.OUTLOOK_USER;
});

async function run(...args: string[]) {
  const program = new Command();
  program.exitOverride();
  registerEmails(program);
  registerDrafts(program);
  await program.parseAsync(["node", "outlook-cli", ...args]);
}

async function refused(...args: string[]): Promise<string> {
  await assert.rejects(run(...args), /exit 1/);
  const last = JSON.parse(lines[lines.length - 1]);
  assert.equal(last.ok, false);
  assert.equal(last.code, "BODY_FILE_REFUSED");
  assert.equal(calls.length, 0, "no Graph call");
  assert.ok(!existsSync(logPath), "no send log line");
  return last.error;
}

function assertNotLogged() {
  const raw = existsSync(logPath) ? readFileSync(logPath, "utf8") : "";
  assert.ok(!raw.includes("Tangerine"), "the body is never in the send log");
  assert.ok(!lines.join("\n").includes("Tangerine"), "the body is never printed");
}

// ---- readBodyFile rules ----

test("readBodyFile reads a 0600 or 0400 regular file", () => {
  assert.equal(readBodyFile(bodyFile("a.txt", BODY)), BODY);
  assert.equal(readBodyFile(bodyFile("b.txt", BODY, 0o400)), BODY);
});

test("readBodyFile refuses a wider mode, a symlink, a FIFO, a folder, a missing file", () => {
  assert.throws(() => readBodyFile(bodyFile("w.txt", BODY, 0o644)), /mode is 0644/);
  assert.throws(() => readBodyFile(bodyFile("g.txt", BODY, 0o640)), /mode is 0640/);
  assert.throws(() => readBodyFile(bodyFile("x.txt", BODY, 0o700)), /mode is 0700/);
  const target = bodyFile("t.txt", BODY);
  const link = join(dir, "link.txt");
  symlinkSync(target, link);
  assert.throws(() => readBodyFile(link), /symlink/);
  const fifo = join(dir, "fifo");
  execFileSync("mkfifo", ["-m", "600", fifo]);
  assert.throws(() => readBodyFile(fifo), /not a regular file/);
  const sub = join(dir, "sub");
  mkdirSync(sub, { mode: 0o700 });
  assert.throws(() => readBodyFile(sub), BodyFileError);
  assert.throws(() => readBodyFile(join(dir, "missing.txt")), /cannot be opened: ENOENT/);
});

test("readBodyFile refuses a file over 1 MB and invalid UTF-8, and never quotes the body", () => {
  assert.equal(MAX_BODY_FILE_BYTES, 1024 * 1024);
  assert.equal(readBodyFile(bodyFile("max.txt", Buffer.alloc(MAX_BODY_FILE_BYTES, 65))).length, MAX_BODY_FILE_BYTES);
  assert.throws(() => readBodyFile(bodyFile("big.txt", Buffer.alloc(MAX_BODY_FILE_BYTES + 1, 65))), /limit is 1048576/);
  try {
    readBodyFile(bodyFile("bad.txt", Buffer.from([0x54, 0x61, 0x6e, 0xff, 0xfe])));
    assert.fail("expected a refusal");
  } catch (e: any) {
    assert.match(e.message, /not valid UTF-8/);
    assert.ok(!e.message.includes("Tan"));
  }
});

test("applyBodyFiles: file options exclude --text, --html and each other", () => {
  const f = bodyFile("a.txt", BODY);
  assert.throws(() => applyBodyFiles({ textFile: f, text: "x" }), /cannot be combined/);
  assert.throws(() => applyBodyFiles({ textFile: f, html: "x" }), /cannot be combined/);
  assert.throws(() => applyBodyFiles({ htmlFile: f, text: "" }), /cannot be combined/);
  assert.throws(() => applyBodyFiles({ textFile: f, htmlFile: f }), /only one/);
  const o: any = { textFile: f };
  applyBodyFiles(o);
  assert.deepEqual(o, { text: BODY });
  const h: any = { htmlFile: bodyFile("h.html", HTML) };
  applyBodyFiles(h);
  assert.deepEqual(h, { html: HTML });
  const none: any = { text: "plain" };
  applyBodyFiles(none);
  assert.deepEqual(none, { text: "plain" });
});

// ---- commands ----

test("emails send --text-file: the file is the body (plus the footer), logged without the body", async () => {
  await run("emails", "send", "--to", "bob@example.test", "--subject", "s1", "--text-file", bodyFile("a.txt", BODY));
  const send = calls.find((c) => c.path.endsWith("/sendMail"))!;
  assert.equal(send.body.message.body.contentType, "Text");
  assert.equal(send.body.message.body.content, `${BODY.trimEnd()}\n\n${AGENT_FOOTER}`);
  const recs = readFileSync(logPath, "utf8").trim().split("\n").map((l) => JSON.parse(l));
  assert.deepEqual(recs.map((r) => r.phase), ["attempt", "result"]);
  assert.deepEqual(recs[0].to, ["bob@example.test"]);
  assertNotLogged();
});

test("emails send --html-file: HTML body", async () => {
  await run("emails", "send", "--to", "bob@example.test", "--subject", "s2", "--html-file", bodyFile("h.html", HTML));
  const send = calls.find((c) => c.path.endsWith("/sendMail"))!;
  assert.equal(send.body.message.body.contentType, "HTML");
  assert.ok(send.body.message.body.content.startsWith(HTML));
  assertNotLogged();
});

test("emails send refuses a bad body file before any Graph call or log line", async () => {
  assert.match(await refused("emails", "send", "--to", "bob@example.test", "--subject", "s", "--text-file", bodyFile("w.txt", BODY, 0o644)), /mode/);
  assert.match(await refused("emails", "send", "--to", "bob@example.test", "--subject", "s", "--text-file", bodyFile("a.txt", BODY), "--text", "x"), /cannot be combined/);
  assert.ok(!lines.join("\n").includes("Tangerine"));
});

test("emails reply and forward take --text-file as the comment", async () => {
  await run("emails", "reply", "--id", "M1", "--text-file", bodyFile("r.txt", BODY));
  const reply = calls.find((c) => c.path.endsWith("/M1/reply"))!;
  assert.ok(reply.body.comment.includes("Tangerine-Lighthouse body 5512<br>second line"));
  calls = [];
  await run("emails", "forward", "--id", "M2", "--to", "fwd@example.test", "--html-file", bodyFile("f.html", HTML));
  const create = calls.find((c) => c.path.endsWith("/M2/createForward"))!;
  assert.ok(JSON.stringify(create.body).includes("Tangerine-Lighthouse html 5512"));
  assertNotLogged();
});

test("drafts create, update, reply and forward take the file options (no footer on drafts)", async () => {
  await run("drafts", "create", "--to", "bob@example.test", "--subject", "d", "--text-file", bodyFile("d.txt", BODY));
  const create = calls.find((c) => c.method === "POST" && c.path.endsWith("/messages"))!;
  assert.deepEqual(create.body.body, { contentType: "Text", content: BODY });
  calls = [];
  await run("drafts", "update", "--id", "D9", "--html-file", bodyFile("u.html", HTML));
  const upd = calls.find((c) => c.method === "PATCH")!;
  assert.deepEqual(upd.body.body, { contentType: "HTML", content: HTML });
  calls = [];
  await run("drafts", "reply", "--id", "M3", "--text-file", bodyFile("dr.txt", BODY));
  assert.ok(JSON.stringify(calls[0].body).includes("Tangerine"));
  calls = [];
  await run("drafts", "forward", "--id", "M4", "--to", "fwd@example.test", "--text-file", bodyFile("df.txt", BODY));
  assert.ok(JSON.stringify(calls[0].body).includes("Tangerine"));
  await assert.rejects(run("drafts", "create", "--to", "bob@example.test", "--subject", "d", "--html-file", bodyFile("bad.html", HTML, 0o644)), /exit 1/);
  assert.equal(JSON.parse(lines[lines.length - 1]).code, "BODY_FILE_REFUSED");
});
