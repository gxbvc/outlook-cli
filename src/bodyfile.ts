// --text-file / --html-file: read a message body from a file instead of argv, so the body never shows in
// the process list or in shell history. The file must be a regular file (no symlink, FIFO or device),
// owner-only (mode 0600 or narrower), at most MAX_BODY_FILE_BYTES, and valid UTF-8. Errors name the
// rule that failed, never the body.
import { closeSync, constants, fstatSync, openSync, readSync } from "node:fs";
import { err } from "./graph.js";

export const MAX_BODY_FILE_BYTES = 1024 * 1024;

export class BodyFileError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BodyFileError";
  }
}

export interface BodyOpts {
  text?: string;
  html?: string;
  textFile?: string;
  htmlFile?: string;
}

export function readBodyFile(path: string): string {
  let fd: number;
  try {
    // O_NOFOLLOW: a symlink is refused. O_NONBLOCK: a FIFO cannot hang the open.
    fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW | (constants.O_NONBLOCK ?? 0));
  } catch (e: any) {
    if (e.code === "ELOOP" || e.code === "EMLINK") throw new BodyFileError("body file is a symlink (not allowed)");
    throw new BodyFileError(`body file cannot be opened: ${e.code || "error"}`);
  }
  try {
    const st = fstatSync(fd); // the SAME fd that is read: no race between check and read
    if (!st.isFile()) throw new BodyFileError("body file is not a regular file");
    const mode = st.mode & 0o777;
    if (mode & ~0o600) {
      throw new BodyFileError(`body file mode is ${mode.toString(8).padStart(4, "0")}; must be 0600 or narrower`);
    }
    if (st.size > MAX_BODY_FILE_BYTES) {
      throw new BodyFileError(`body file is ${st.size} bytes; the limit is ${MAX_BODY_FILE_BYTES}`);
    }
    const buf = Buffer.alloc(st.size + 1);
    let n = 0;
    while (n < buf.length) {
      const got = readSync(fd, buf, n, buf.length - n, null);
      if (got === 0) break;
      n += got;
    }
    if (n !== st.size) throw new BodyFileError("body file changed size while it was read");
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(buf.subarray(0, n));
    } catch {
      throw new BodyFileError("body file is not valid UTF-8");
    }
  } finally {
    closeSync(fd);
  }
}

// Moves a --text-file / --html-file body into opts.text / opts.html. A file option cannot be combined
// with --text, --html or the other file option.
export function applyBodyFiles(opts: BodyOpts): void {
  const files = [opts.textFile, opts.htmlFile].filter((v) => v !== undefined);
  if (files.length === 0) return;
  if (files.length > 1) throw new BodyFileError("use only one of --text-file and --html-file");
  if (opts.text !== undefined || opts.html !== undefined) {
    throw new BodyFileError("--text-file and --html-file cannot be combined with --text or --html");
  }
  if (opts.textFile !== undefined) opts.text = readBodyFile(opts.textFile);
  else opts.html = readBodyFile(opts.htmlFile!);
  delete opts.textFile;
  delete opts.htmlFile;
}

// Command entry point: a refused body file ends the command (code BODY_FILE_REFUSED) before any Graph call
// or send-log line.
export function applyBodyFilesOrExit(opts: BodyOpts): void {
  try {
    applyBodyFiles(opts);
  } catch (e: any) {
    if (e instanceof BodyFileError) err(e.message, "BODY_FILE_REFUSED");
    throw e;
  }
}
