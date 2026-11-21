// Local send log. Every Graph call that makes Exchange send mail goes through
// loggedSend(). It writes an "attempt" line BEFORE the call (if that write fails,
// the send is refused) and a "result" line after. One JSON object per line.
//
// Never logged: subject text, body, comments, attachment names or contents, tokens.
// The subject is only kept as the first 8 hex characters of its sha256.
import { createHash, randomBytes } from "crypto";
import {
  closeSync, constants, fstatSync, fsyncSync, lstatSync, mkdirSync, openSync, writeSync,
} from "fs";
import { homedir } from "os";
import { basename, dirname, isAbsolute, join } from "path";
import { execFileSync } from "child_process";
import { loadTokens } from "./auth.js";

export type SendKind =
  | "send" | "reply" | "reply_all" | "forward" | "draft_send" | "draft_send_scheduled"
  | "event_create" | "event_update" | "event_cancel" | "event_delete" | "event_forward";

export interface SendInfo {
  kind: SendKind;
  // Addresses or Graph recipient objects; null = not known before the call.
  to?: any[] | null;
  cc?: any[] | null;
  bcc?: any[] | null;
  // The exact subject Graph will use, or null when it is not known (say why in subjectNote).
  subject?: string | null;
  subjectNote?: string;
  message_id?: string;           // the message replied to / forwarded / the draft sent
  event_id?: string;
  note?: string;                 // fixed text from the code (never user input)
}

export class SendLogError extends Error {
  code = "SEND_LOG_FAILED";
  constructor(message: string) {
    super(`Send refused: the send log could not be written (${message}). Nothing was sent.`);
  }
}

// Error code for a command's JSON envelope: SEND_LOG_FAILED when the log refused the send.
export function failCode(e: unknown, fallback: string): string {
  return e instanceof SendLogError ? e.code : fallback;
}

export function sendLogPath(): string {
  const env = process.env.OUTLOOK_CLI_SEND_LOG;
  if (env !== undefined && env !== "") {
    if (!isAbsolute(env)) throw new SendLogError("OUTLOOK_CLI_SEND_LOG must be an absolute path");
    return env;
  }
  return join(homedir(), ".local", "state", "outlook-cli", "sends.jsonl");
}

export function subjectSha8(subject: string | null | undefined): string | null {
  if (subject === null || subject === undefined) return null;
  return createHash("sha256").update(subject, "utf8").digest("hex").slice(0, 8);
}

// Addresses only, lower-case. Accepts strings or Graph recipient objects.
export function addrs(list: any[] | null | undefined): string[] | null {
  if (list === null || list === undefined) return null;
  const out: string[] = [];
  for (const r of list) {
    const a = typeof r === "string" ? r : r?.emailAddress?.address;
    if (typeof a === "string" && a.trim()) out.push(a.trim().toLowerCase());
  }
  return out;
}

// The mailbox used: --user (OUTLOOK_USER), else the account in the cached access token
// (only the upn claim is read; the token itself is never logged), else "me".
function fromMailbox(): string {
  if (process.env.OUTLOOK_USER) return process.env.OUTLOOK_USER.toLowerCase();
  try {
    const tok = loadTokens()?.access_token;
    const part = tok?.split(".")[1];
    if (part) {
      const claims = JSON.parse(Buffer.from(part, "base64url").toString("utf8"));
      const who = claims.upn || claims.preferred_username || claims.unique_name;
      if (typeof who === "string" && who) return who.toLowerCase();
    }
  } catch { /* fall through */ }
  return "me";
}

function parentProcessName(): string | null {
  try {
    const name = execFileSync("ps", ["-o", "comm=", "-p", String(process.ppid)], {
      encoding: "utf8", timeout: 2000, stdio: ["ignore", "pipe", "ignore"],
    }).trim();
    return name ? basename(name) : null;
  } catch {
    return null;
  }
}

export function callerInfo(): Record<string, string | null> {
  const c: Record<string, string | null> = {};
  const herdr: [string, string][] = [
    ["herdr_tab_id", "HERDR_TAB_ID"], ["herdr_pane_id", "HERDR_PANE_ID"], ["herdr_workspace_id", "HERDR_WORKSPACE_ID"],
  ];
  let anyHerdr = false;
  for (const [field, env] of herdr) {
    if (process.env[env]) { c[field] = process.env[env]!; anyHerdr = true; }
  }
  if (!anyHerdr) c.parent_process = parentProcessName();
  const xpc = process.env.XPC_SERVICE_NAME;
  if (xpc && xpc !== "0") c.launchd_label = xpc;
  c.cwd = basename(process.cwd());
  return c;
}

function checkDir(dir: string): void {
  let st;
  try {
    st = lstatSync(dir);
  } catch (e: any) {
    if (e.code !== "ENOENT") throw new SendLogError(`cannot stat folder: ${e.code || e.message}`);
    mkdirSync(dir, { recursive: true, mode: 0o700 });
    st = lstatSync(dir);
  }
  if (st.isSymbolicLink()) throw new SendLogError("log folder is a symlink");
  if (!st.isDirectory()) throw new SendLogError("log folder is not a directory");
  if ((st.mode & 0o077) !== 0) {
    throw new SendLogError(`log folder mode ${(st.mode & 0o777).toString(8)} is wider than 700`);
  }
}

function checkFileStat(st: { mode: number; isFile(): boolean; isSymbolicLink(): boolean; uid: number }): void {
  if (st.isSymbolicLink()) throw new SendLogError("log file is a symlink");
  if (!st.isFile()) throw new SendLogError("log file is not a regular file");
  if ((st.mode & 0o177) !== 0) {
    throw new SendLogError(`log file mode ${(st.mode & 0o777).toString(8)} is wider than 600`);
  }
  if (typeof process.getuid === "function" && st.uid !== process.getuid()) {
    throw new SendLogError("log file is owned by another user");
  }
}

// Append one line. Never truncates or rewrites. Throws SendLogError on any problem.
export function appendLine(record: object): void {
  const path = sendLogPath();
  checkDir(dirname(path));
  try {
    checkFileStat(lstatSync(path));
  } catch (e: any) {
    if (e instanceof SendLogError) throw e;
    if (e.code !== "ENOENT") throw new SendLogError(`cannot stat file: ${e.code || e.message}`);
  }
  let fd: number;
  try {
    fd = openSync(path, constants.O_WRONLY | constants.O_APPEND | constants.O_CREAT | constants.O_NOFOLLOW, 0o600);
  } catch (e: any) {
    throw new SendLogError(e.code === "ELOOP" ? "log file is a symlink" : `cannot open: ${e.code || e.message}`);
  }
  try {
    checkFileStat(fstatSync(fd) as any);
    const buf = Buffer.from(JSON.stringify(record) + "\n", "utf8");
    let off = 0;
    while (off < buf.length) off += writeSync(fd, buf, off, buf.length - off);
    fsyncSync(fd);
  } catch (e: any) {
    if (e instanceof SendLogError) throw e;
    throw new SendLogError(`cannot write: ${e.code || e.message}`);
  } finally {
    closeSync(fd);
  }
}

// The one gate for every Graph call that sends mail.
export async function loggedSend<T>(info: SendInfo, call: () => Promise<T>): Promise<T> {
  const send_id = randomBytes(6).toString("hex");
  const base: Record<string, unknown> = {
    send_id,
    kind: info.kind,
    from: fromMailbox(),
  };
  const attempt: Record<string, unknown> = {
    ts: new Date().toISOString(),
    phase: "attempt",
    ...base,
    to: addrs(info.to ?? null),
    cc: addrs(info.cc ?? null),
    bcc: addrs(info.bcc ?? null),
    subject_sha8: subjectSha8(info.subject),
  };
  if (attempt.subject_sha8 === null) attempt.subject_note = info.subjectNote || "subject not known before the send";
  if (info.message_id) attempt.message_id = info.message_id;
  if (info.event_id) attempt.event_id = info.event_id;
  if (info.note) attempt.note = info.note;
  attempt.caller = callerInfo();

  appendLine(attempt); // throws SendLogError: the Graph call never happens

  try {
    const res = await call();
    writeResult({ ts: new Date().toISOString(), phase: "result", ...base, ok: true });
    return res;
  } catch (e: any) {
    const status = typeof e?.status === "number" ? e.status : null;
    writeResult({ ts: new Date().toISOString(), phase: "result", ...base, ok: false, status });
    throw e;
  }
}

// The mail may already be out, so a failed result write warns instead of failing the command.
function writeResult(record: object): void {
  try {
    appendLine(record);
  } catch (e: any) {
    process.stderr.write(JSON.stringify({ warning: "send log result line not written", detail: e.message }) + "\n");
  }
}
