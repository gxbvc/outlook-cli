import { Command } from "commander";
import { chmodSync, closeSync, createWriteStream, existsSync, lstatSync, mkdirSync, openSync, readFileSync, readSync, readdirSync, realpathSync, renameSync, statSync, unlinkSync, writeFileSync } from "fs";
import { Readable } from "stream";
import { pipeline } from "stream/promises";
import { createHash } from "crypto";
import { basename, dirname, isAbsolute, join, relative, resolve as resolvePath, sep } from "path";
import { homedir } from "os";
import { TOKENS_DIR } from "../config.js";
import { GraphError, graphRawRequest, graphRequest, out, err } from "../graph.js";
import { QuickXorHash } from "../quickxor.js";

// OneDrive and SharePoint files via Graph. Read: search, ls, info, get.
// Upload: put, sync (needs Files.ReadWrite.All on top of Files.Read.All and
// Sites.Read.All; re-run `outlook-cli auth` once those scopes are granted).
// Never deletes or moves anything.

const DEFAULT_CACHE = join(homedir(), "tmp", "onedrive-cache");
const SHARED_LIBRARY_NAME = "Ricky Bureau - GXB";

// Files over 4 MiB upload through a Graph upload session in chunks; the chunk
// size must be a multiple of 320 KiB, so 5 MiB it is.
const UPLOAD_SESSION_THRESHOLD = 4 * 1024 * 1024;
const UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024;

// --drive ricky resolves the shared library's drive id once through a site
// search and caches only the id here (never tokens or file data).
function driveCacheFile(): string {
  return process.env.OUTLOOK_DRIVE_CACHE || join(TOKENS_DIR, "drive-ids.json");
}

function readDriveCache(): Record<string, string> {
  try {
    return JSON.parse(readFileSync(driveCacheFile(), "utf8"));
  } catch {
    return {};
  }
}

function writeDriveCache(cache: Record<string, string>): void {
  mkdirSync(dirname(driveCacheFile()), { recursive: true });
  writeFileSync(driveCacheFile(), JSON.stringify(cache, null, 2) + "\n", { mode: 0o600 });
}

async function rickyDriveId(): Promise<string> {
  const cache = readDriveCache();
  const cached = cache[SHARED_LIBRARY_NAME];
  if (cached) {
    // Re-verify the cached id once per run: a stale id (for example one cached
    // by the old first-hit fallback) must never be written through.
    try {
      const drive = await graphRequest(`/drives/${cached}`);
      if (drive?.name === SHARED_LIBRARY_NAME) return cached;
    } catch (e: any) {
      if (!(e instanceof GraphError)) throw e;
      // 404 or access error: treat as stale and re-resolve below.
    }
    delete cache[SHARED_LIBRARY_NAME];
    writeDriveCache(cache);
  }
  const sites = await graphRequest(`/sites?search=${encodeURIComponent(SHARED_LIBRARY_NAME)}`);
  // Exact library name only. A fuzzy or first-hit fallback would cache the
  // wrong drive id and later puts/syncs would write into an unrelated library.
  const site = sites?.value?.find((s: any) => s.displayName === SHARED_LIBRARY_NAME);
  if (!site) err(`SharePoint site "${SHARED_LIBRARY_NAME}" not found (exact match). Check Sites.Read.All consent.`, "SITE_NOT_FOUND");
  const drive = await graphRequest(`/sites/${site.id}/drive`);
  cache[SHARED_LIBRARY_NAME] = drive.id;
  writeDriveCache(cache);
  return drive.id;
}

async function driveBase(drive: string): Promise<string> {
  // The global --user remaps /me to another mailbox; files commands do not support it.
  if (process.env.OUTLOOK_USER) err("files does not support --user", "USER_UNSUPPORTED");
  if (drive === "me") return "/me/drive";
  if (drive === "ricky") return `/drives/${await rickyDriveId()}`;
  err(`Unknown drive "${drive}" (use me or ricky)`, "BAD_DRIVE");
}

// Encode each path segment so spaces, #, % and unicode survive the URL, but
// first reject anything that could escape the drive or the local cache: empty
// segments, ".", "..", backslashes and control characters. WHATWG URL parsing
// would otherwise resolve dot segments out of the encoded path. ONE leading
// and ONE trailing "/" are allowed and stripped; "" or "/" means the root.
function badSegmentReason(rel: string): string | null {
  for (const seg of rel.split("/")) {
    if (!seg || seg === "." || seg === "..") return `bad path segment "${seg}"`;
    if (seg.includes("\\") || /[\u0000-\u001f\u007f]/.test(seg)) return `bad character in path segment "${seg}"`;
  }
  return null;
}

function normalizeRel(rel: string): string {
  let p = rel;
  if (p.startsWith("/")) p = p.slice(1);
  if (p.endsWith("/")) p = p.slice(0, -1);
  return p;
}

function encPath(rel: string): string {
  const p = normalizeRel(rel);
  if (p === "") return "";
  const reason = badSegmentReason(p);
  if (reason) err(`Bad path: ${reason} in "${rel}"`, "BAD_PATH");
  return p.split("/").map(encodeURIComponent).join("/");
}

// Graph item path relative to the drive root ("root:/segment/segment").
function itemPath(base: string, rel: string): string {
  const enc = encPath(rel);
  return enc ? `${base}/root:/${enc}` : `${base}/root`;
}

function relPathOf(item: any): string {
  const parent = item.parentReference?.path || "";
  const idx = parent.indexOf("root:");
  const base = idx >= 0 ? parent.slice(idx + 5).replace(/^\//, "") : "";
  return base ? `${base}/${item.name}` : item.name;
}

// Show names, ids, sizes and hashes only. Never file contents.
function slim(item: any) {
  return {
    name: item.name,
    id: item.id,
    type: item.folder ? "folder" : "file",
    size: item.size,
    path: relPathOf(item),
    lastModified: item.lastModifiedDateTime,
    hash: item.file?.hashes?.quickXorHash || item.file?.hashes?.sha1Hash || null,
    webUrl: item.webUrl,
  };
}

function graphFail(e: any): never {
  const status = e instanceof GraphError ? e.status : 0;
  if (status === 404) err("Not found in this drive (404). Check the path and --drive.", "NOT_FOUND");
  if (status === 403)
    err(
      "Access denied (403). The app likely lacks Files.Read.All / Sites.Read.All consent. Grant the scopes, then run `outlook-cli auth` again.",
      "CONSENT_REQUIRED"
    );
  if (status === 429)
    err(
      /after \d+ retries/.test(e.message)
        ? e.message
        : `Throttled (429). Retry after ${e.retryAfter ?? "unknown"} seconds.`,
      "THROTTLED"
    );
  throw e;
}

function expandTilde(p: string): string {
  return p.replace(/^~(?=\/|$)/, homedir());
}

// Create the folder chain from root down to dir and chmod every level 700 —
// a cache root or drive folder that already existed keeps its old mode otherwise.
// Never follows a symlink: refuses instead.
function mkdirChmodChain(root: string, dir: string): void {
  mkdirSync(root, { recursive: true, mode: 0o700 });
  chmodSync(root, 0o700);
  const rel = relative(root, dir);
  if (!rel || rel.startsWith("..") || isAbsolute(rel)) return;
  let cur = root;
  for (const seg of rel.split(sep).filter(Boolean)) {
    cur = join(cur, seg);
    if (existsSync(cur) && lstatSync(cur).isSymbolicLink()) err(`Refusing to follow the symlink at ${cur}`, "PATH_ESCAPE");
    mkdirSync(cur, { recursive: true, mode: 0o700 });
    chmodSync(cur, 0o700);
  }
}

// Turn a raw Response (streaming download) into a GraphError when not ok.
async function checkRaw(res: Response, path: string): Promise<Response> {
  if (res.ok) return res;
  const body = await res.text();
  const retryAfter = res.headers.get("Retry-After");
  throw new GraphError(res.status, `Graph API ${res.status} ${path}: ${body}`, retryAfter ? parseInt(retryAfter, 10) : null);
}

function isNotFound(e: any): boolean {
  return e instanceof GraphError && e.status === 404;
}

// Create missing folders one segment at a time (conflictBehavior fail: never
// replaces an existing folder). Only called when the user asked for it.
async function ensureFolder(base: string, segments: string[]): Promise<void> {
  let cur = "";
  for (const seg of segments) {
    const item = cur ? `${base}/root:/${cur}/${encodeURIComponent(seg)}` : `${base}/root:/${encodeURIComponent(seg)}`;
    const parent = cur ? `${base}/root:/${cur}` : `${base}/root`;
    try {
      await graphRequest(item);
    } catch (e: any) {
      if (!isNotFound(e)) throw e;
      await graphRequest(`${parent}/children`, {
        method: "POST",
        body: JSON.stringify({ name: seg, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
    }
    cur = cur ? `${cur}/${encodeURIComponent(seg)}` : encodeURIComponent(seg);
  }
}

// Stream the sha1 of a local file (fallback for personal OneDrive items).
function sha1Of(localPath: string): string {
  const hash = createHash("sha1");
  const fd = openSync(localPath, "r");
  try {
    const buf = Buffer.alloc(65536);
    for (;;) {
      const n = readSync(fd, buf, 0, buf.length, null);
      if (n === 0) break;
      hash.update(buf.subarray(0, n));
    }
  } finally {
    closeSync(fd);
  }
  return hash.digest("hex").toUpperCase();
}

// Honor Graph's Retry-After (seconds or HTTP-date), capped at 60 s; fall back
// to exponential backoff when the header is absent.
export function retryDelayMs(header: string | null, attempt: number): number {
  const CAP = 60_000;
  if (header) {
    const secs = parseInt(header, 10);
    if (!isNaN(secs)) return Math.min(secs * 1000, CAP);
    const at = Date.parse(header);
    if (!isNaN(at)) return Math.min(Math.max(at - Date.now(), 0), CAP);
  }
  return Math.min(200 * 2 ** attempt, CAP);
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

// Upload one local file. Small: single PUT with conflictBehavior on the URL.
// Large: an upload session, streamed one chunk at a time (never the whole file
// in memory), chunks retried on 429/5xx/network errors with Retry-After (max
// 3), the session cancelled on final failure. Hashes while sending; refuses
// when the local file changes size during the upload; verifies the stored
// quickXorHash against the local one when Graph reports it.
async function uploadFile(base: string, rel: string, localPath: string, replace: boolean, ifMatch?: string) {
  const startSize = statSync(localPath).size;
  const hash = createHash("sha256");
  const enc = encPath(rel);
  const headers: Record<string, string> = { "Content-Type": "application/octet-stream" };
  if (ifMatch) headers["If-Match"] = ifMatch;

  const checkStoredHash = (item: any) => {
    const remoteQxh = item?.file?.hashes?.quickXorHash;
    if (remoteQxh && remoteQxh !== quickXorHashOf(localPath)) {
      err(`Uploaded quickXorHash mismatch for "${rel}"; the upload did not round-trip`, "UPLOAD_HASH_MISMATCH");
    }
  };

  if (startSize <= UPLOAD_SESSION_THRESHOLD) {
    const data = readFileSync(localPath);
    if (data.length !== startSize) err(`Local file "${localPath}" changed during upload`, "LOCAL_FILE_CHANGED");
    hash.update(data);
    const res = await checkRaw(
      await graphRawRequest(`${base}/root:/${enc}:/content?@microsoft.graph.conflictBehavior=${replace ? "replace" : "fail"}`, {
        method: "PUT",
        headers,
        body: data,
      }),
      rel
    );
    const item = await res.json();
    checkStoredHash(item);
    return { id: item.id, size: startSize, sha256: hash.digest("hex"), mode: "simple", eTag: item.eTag };
  }

  const sessionHeaders: Record<string, string> = {};
  if (ifMatch) sessionHeaders["If-Match"] = ifMatch;
  const session = await graphRequest(`${base}/root:/${enc}:/createUploadSession`, {
    method: "POST",
    headers: sessionHeaders,
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": replace ? "replace" : "fail" } }),
  });
  const uploadUrl = session?.uploadUrl;
  if (!uploadUrl) err("Graph returned no upload URL for the session", "UPLOAD_SESSION_FAILED");

  const cancelSession = async () => {
    try {
      await graphRawRequest(uploadUrl, { method: "DELETE" });
    } catch {
      // cancelling is best effort
    }
  };

  let item: any = null;
  let sent = 0;
  for (let start = 0; start < startSize; start += UPLOAD_CHUNK_SIZE) {
    const end = Math.min(start + UPLOAD_CHUNK_SIZE, startSize) - 1;
    const len = end - start + 1;
    const buf = Buffer.alloc(len);
    const fd = openSync(localPath, "r");
    let nRead = 0;
    try {
      nRead = readSync(fd, buf, 0, len, start);
    } finally {
      closeSync(fd);
    }
    if (nRead !== len || statSync(localPath).size !== startSize) {
      await cancelSession();
      err(`Local file "${localPath}" changed during upload`, "LOCAL_FILE_CHANGED");
    }
    const chunk = nRead === buf.length ? buf : buf.subarray(0, nRead);
    hash.update(chunk);

    let attempt = 0;
    for (;;) {
      let chunkRes: Response;
      try {
        chunkRes = await graphRawRequest(uploadUrl, {
          method: "PUT",
          headers: {
            "Content-Length": String(chunk.length),
            "Content-Range": `bytes ${start}-${start + chunk.length - 1}/${startSize}`,
          },
          body: chunk,
        });
      } catch (e: any) {
        // A network error (fetch rejects) retries like a 5xx.
        if (attempt < 3) {
          attempt++;
          await sleep(retryDelayMs(null, attempt));
          continue;
        }
        await cancelSession();
        throw new GraphError(0, `Network error on chunk ${start}-${start + chunk.length - 1} after ${attempt} retries: ${e?.message || e}`);
      }
      if (chunkRes.ok) {
        sent += chunk.length;
        const text = await chunkRes.text();
        if (text) {
          try {
            const parsed = JSON.parse(text);
            if (parsed.id) item = parsed; // the final chunk returns the drive item
          } catch {
            // 202 progress responses are JSON, but never fail the upload over parsing
          }
        }
        break;
      }
      const retryAfter = chunkRes.headers.get("Retry-After");
      const body = await chunkRes.text();
      const retryable = chunkRes.status === 429 || chunkRes.status >= 500;
      if (retryable && attempt < 3) {
        attempt++;
        await sleep(retryDelayMs(retryAfter, attempt));
        continue;
      }
      await cancelSession();
      throw new GraphError(
        chunkRes.status,
        `Graph API ${chunkRes.status} chunk ${start}-${start + chunk.length - 1} after ${attempt} retries: ${body}`,
        retryAfter ? parseInt(retryAfter, 10) : null
      );
    }
  }
  if (statSync(localPath).size !== startSize || sent !== startSize) {
    await cancelSession();
    err(`Local file "${localPath}" changed during upload`, "LOCAL_FILE_CHANGED");
  }
  if (!item) err("Upload finished but Graph returned no drive item", "UPLOAD_INCOMPLETE");
  checkStoredHash(item);
  return { id: item.id, size: startSize, sha256: hash.digest("hex"), mode: "session", eTag: item.eTag };
}

// OneDrive for Business and SharePoint return only quickXorHash in
// file.hashes (sha1Hash is personal OneDrive only), so this is the primary
// comparison for sync. Streamed off disk; never the whole file in memory.
function quickXorHashOf(localPath: string): string {
  const q = new QuickXorHash();
  const fd = openSync(localPath, "r");
  try {
    const buf = Buffer.alloc(65536);
    for (;;) {
      const n = readSync(fd, buf, 0, buf.length, null);
      if (n === 0) break;
      q.update(buf.subarray(0, n));
    }
  } finally {
    closeSync(fd);
  }
  return q.digestBase64();
}

// OneDrive is case-insensitive: a.txt and A.txt map to one remote item, so
// uploading both would silently replace one with the other. Refuse both.
export function detectCaseCollisions(rels: string[]): Set<string> {
  const seen = new Map<string, string>();
  const collisions = new Set<string>();
  for (const rel of rels) {
    const key = rel.toLowerCase();
    if (seen.has(key)) {
      collisions.add(seen.get(key)!);
      collisions.add(rel);
    } else {
      seen.set(key, rel);
    }
  }
  return collisions;
}

// Per-target sync state: the eTag each remote file had at its last successful
// sync. A remote whose eTag differs from the stored one changed since then —
// even when it looks older — so both-sides-edited becomes a conflict.
function syncStateDir(): string {
  return process.env.OUTLOOK_SYNC_STATE_DIR || join(homedir(), ".config", "outlook-cli", "sync-state");
}

function syncStateFile(drive: string, target: string): string {
  const h = createHash("sha256").update(`${drive}:${target}`).digest("hex").slice(0, 16);
  return join(syncStateDir(), `${drive}-${h}.json`);
}

function readSyncState(file: string): Record<string, { eTag?: string; syncedAt?: string }> {
  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch {
    return {};
  }
}

function writeSyncState(file: string, state: Record<string, { eTag?: string; syncedAt?: string }>): void {
  mkdirSync(dirname(file), { recursive: true, mode: 0o700 });
  writeFileSync(file, JSON.stringify(state, null, 2) + "\n", { mode: 0o600 });
}

export function registerFiles(program: Command) {
  const files = program.command("files").description("OneDrive and SharePoint files (read and upload)");

  files
    .command("search")
    .description("Search files and folders in a drive")
    .argument("<query>", "Search text")
    .option("--drive <drive>", "me (your OneDrive) or ricky (Ricky Bureau - GXB shared library)", "me")
    .action(async (query, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const sq = encodeURIComponent(query.replace(/'/g, "''"));
        // Follow @odata.nextLink so large result sets are not cut off.
        const results: any[] = [];
        let url: string | null = `${base}/search(q='${sq}')?$top=50`;
        for (let hops = 0; url && hops < 20; hops++) {
          const data = await graphRequest(url);
          results.push(...(data?.value || []));
          url = data?.["@odata.nextLink"] || null;
        }
        out({ query, drive: opts.drive, count: results.length, results: results.map(slim) });
      } catch (e: any) {
        graphFail(e);
      }
    });

  files
    .command("ls")
    .description("List a folder (paths are relative to the drive root)")
    .argument("<path>", "Folder path, e.g. 10_Products or /")
    .option("--drive <drive>", "me or ricky", "me")
    .action(async (p, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const isRoot = p === "/" || p.trim() === "";
        const enc = isRoot ? "" : encPath(p);
        const select = "$select=name,id,size,folder,file,lastModifiedDateTime,parentReference,webUrl";
        // Follow @odata.nextLink so folders over the page size are complete.
        const items: any[] = [];
        let url: string | null = enc
          ? `${base}/root:/${enc}:/children?${select}`
          : `${base}/root/children?${select}`;
        for (let hops = 0; url && hops < 50; hops++) {
          const data = await graphRequest(url);
          items.push(...(data?.value || []));
          url = data?.["@odata.nextLink"] || null;
        }
        out({ drive: opts.drive, path: p, count: items.length, items: items.map(slim) });
      } catch (e: any) {
        graphFail(e);
      }
    });

  files
    .command("info")
    .description("Show one item's metadata (no file contents)")
    .argument("<path>", "File or folder path")
    .option("--drive <drive>", "me or ricky", "me")
    .action(async (p, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const item = await graphRequest(itemPath(base, p));
        out({ drive: opts.drive, ...slim(item) });
      } catch (e: any) {
        graphFail(e);
      }
    });

  files
    .command("get")
    .description("Download a file to the local cache (streams to disk, never loads it whole in memory)")
    .argument("<path>", "File path relative to the drive root")
    .option("--drive <drive>", "me or ricky", "me")
    .option("--out <dir>", "Cache root directory", DEFAULT_CACHE)
    .action(async (p, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const enc = encPath(p);
        if (!enc) err("Give a file path, e.g. 10_Products/notes.txt", "BAD_PATH");
        const meta = await graphRequest(`${base}/root:/${enc}`);
        if (meta.folder) err("That path is a folder; `files get` works on files only", "IS_FOLDER");

        const cacheRoot = resolvePath(expandTilde(opts.out));
        mkdirSync(cacheRoot, { recursive: true, mode: 0o700 });
        // The cache root itself is the user's choice: resolve it once through
        // any symlinks. Every level below it must be symlink-free.
        const realRoot = realpathSync(cacheRoot);
        const driveRoot = join(realRoot, opts.drive);
        const target = join(driveRoot, ...p.split("/"));
        let cur = driveRoot;
        for (const seg of relative(driveRoot, dirname(target)).split(sep).filter(Boolean)) {
          cur = join(cur, seg);
          if (existsSync(cur)) {
            if (lstatSync(cur).isSymbolicLink()) err(`Refusing to follow the symlink at ${cur}`, "PATH_ESCAPE");
            if (!statSync(cur).isDirectory()) err(`Not a folder: ${cur}`, "PATH_ESCAPE");
          }
        }
        if (existsSync(target) && lstatSync(target).isSymbolicLink()) err(`Refusing to follow the symlink at ${target}`, "PATH_ESCAPE");

        const res = await checkRaw(await graphRawRequest(`${base}/root:/${enc}:/content`), p);
        mkdirChmodChain(realRoot, dirname(target));
        if (!realpathSync(dirname(target)).startsWith(realRoot + sep)) err("Resolved path escapes the cache directory", "PATH_ESCAPE");

        // Write to a temp file first, verify, then rename: a failed download
        // never leaves a partial file behind that looks complete.
        const tmp = join(dirname(target), `.${basename(target)}.part-${process.pid}-${Date.now()}`);
        const hash = createHash("sha256");
        let size = 0;
        try {
          await pipeline(
            Readable.fromWeb(res.body as any),
            async function* (source) {
              for await (const chunk of source) {
                hash.update(chunk);
                size += chunk.length;
                yield chunk;
              }
            },
            createWriteStream(tmp, { flags: "wx", mode: 0o600 })
          );
          chmodSync(tmp, 0o600);
          if (meta.size != null && size !== meta.size) {
            throw new GraphError(0, `Downloaded ${size} bytes but Graph says the file is ${meta.size}`);
          }
          const remoteQxh = meta.file?.hashes?.quickXorHash;
          if (remoteQxh && quickXorHashOf(tmp) !== remoteQxh) {
            throw new GraphError(0, "Downloaded quickXorHash does not match the remote item");
          }
          renameSync(tmp, target);
        } catch (e) {
          try {
            unlinkSync(tmp);
          } catch {
            // nothing to clean up
          }
          throw e;
        }

        out({ drive: opts.drive, path: p, localPath: target, size, sha256: hash.digest("hex") });
      } catch (e: any) {
        graphFail(e);
      }
    });

  files
    .command("put")
    .description(
      "Upload a local file. Never deletes or moves anything; refuses to overwrite without --force (files only, never a folder)."
    )
    .argument("<local>", "Local file path")
    .argument("<path>", "Remote path relative to the drive root")
    .option("--drive <drive>", "me or ricky", "me")
    .option("--force", "Replace the remote file if it exists (never replaces a folder)")
    .option("--mkdir", "Create missing parent folders")
    .action(async (local, p, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const enc = encPath(p);
        if (!enc) err("Give a remote path, e.g. 10_Products/notes.txt", "BAD_PATH");
        const localPath = resolvePath(expandTilde(local));
        if (!existsSync(localPath) || !statSync(localPath).isFile()) err(`No local file: ${localPath}`, "LOCAL_MISSING");

        const segments = p.split("/").filter(Boolean);
        if (segments.length > 1) {
          const dir = segments.slice(0, -1);
          try {
            await graphRequest(itemPath(base, dir.join("/")));
          } catch (e: any) {
            if (!isNotFound(e)) throw e;
            if (!opts.mkdir) err(`Remote folder "${dir.join("/")}" does not exist; pass --mkdir to create it`, "PARENT_MISSING");
            await ensureFolder(base, dir);
          }
        }

        // Exists? Without --force: refuse. With --force: replace files only,
        // never a folder (conflictBehavior=replace on a file below).
        let replace = false;
        try {
          const existing = await graphRequest(itemPath(base, p));
          if (existing.folder) err(`Remote path "${p}" is a folder; put only replaces files`, "IS_FOLDER");
          if (!opts.force) err(`"${p}" already exists; pass --force to replace the file`, "ALREADY_EXISTS");
          replace = true;
        } catch (e: any) {
          if (!isNotFound(e)) throw e;
        }

        const r = await uploadFile(base, p, localPath, replace);
        out({ drive: opts.drive, path: p, localPath, size: r.size, sha256: r.sha256, id: r.id, uploadedVia: r.mode, replaced: replace });
      } catch (e: any) {
        graphFail(e);
      }
    });

  files
    .command("sync")
    .description(
      "Mirror a local folder into the drive: uploads only new or changed files (size + quickXorHash compared against Graph, sha1 fallback). Never deletes remote files."
    )
    .argument("<dir>", "Local folder to upload")
    .argument("<path>", "Remote target folder relative to the drive root")
    .option("--drive <drive>", "me or ricky", "me")
    .option("--dry-run", "Show the upload plan without uploading")
    .option("--apply", "Upload the new or changed files")
    .action(async (dir, p, opts) => {
      try {
        if (!!opts.dryRun === !!opts.apply) err("Give exactly one of --dry-run or --apply", "BAD_MODE");
        const base = await driveBase(opts.drive);
        const localDir = resolvePath(expandTilde(dir));
        if (!existsSync(localDir) || !statSync(localDir).isDirectory()) err(`No local folder: ${localDir}`, "LOCAL_MISSING");
        // Validate the remote target the same way as every other subcommand.
        // One leading and one trailing "/" are allowed and stripped; "/" is the drive root.
        const rootTarget = normalizeRel(p) === "";
        const normalizedTarget = rootTarget ? "/" : normalizeRel(p);
        if (!rootTarget) encPath(normalizedTarget); // throws BAD_PATH on dot/backslash/control segments

        const stateFile = syncStateFile(opts.drive, normalizedTarget);
        const state = readSyncState(stateFile);

        const localFiles: string[] = [];
        const skippedEntries: any[] = [];
        (function walk(d: string, rel: string) {
          for (const e of readdirSync(join(d, rel), { withFileTypes: true })) {
            if (e.name === ".DS_Store") continue;
            const r = rel ? `${rel}/${e.name}` : e.name;
            if (e.isDirectory()) walk(d, r);
            else if (e.isFile()) localFiles.push(r);
            else skippedEntries.push({ file: r, skipped: e.isSymbolicLink() ? "symlink" : "special file" });
          }
        })(localDir, "");
        localFiles.sort();

        // Case-only name collisions (OneDrive is case-insensitive; APFS cannot
        // hold both spellings, so this catches them before anything uploads).
        const caseCollisions = detectCaseCollisions(localFiles);

        const plan: any[] = [];
        let unchanged = 0;
        for (const rel of localFiles) {
          const remoteRel = rootTarget ? rel : `${normalizedTarget}/${rel}`;
          // One bad local file name (a backslash or control character is legal
          // on macOS) is skipped and reported, not an abort of the whole sync.
          const badName = badSegmentReason(remoteRel);
          if (badName) {
            plan.push({ file: rel, remote: remoteRel, skipped: `bad file name: ${badName}` });
            continue;
          }
          const st = statSync(join(localDir, rel));
          if (caseCollisions.has(rel)) {
            plan.push({ file: rel, remote: remoteRel, conflict: "case-only name collision with another local file; upload refused" });
            continue;
          }
          let remote: any = null;
          try {
            remote = await graphRequest(itemPath(base, remoteRel));
          } catch (e: any) {
            if (!isNotFound(e)) throw e;
          }
          if (remote?.folder) {
            plan.push({ file: rel, remote: remoteRel, skipped: "remote path is a folder; never replaced" });
            continue;
          }
          // OneDrive for Business/SharePoint give quickXorHash only; personal
          // OneDrive also has sha1Hash. Prefer quickXorHash, fall back to sha1;
          // a remote item with neither hash counts as changed (re-upload).
          const hashes = remote?.file?.hashes || {};
          let hashMatches = false;
          if (hashes.quickXorHash) hashMatches = quickXorHashOf(join(localDir, rel)) === hashes.quickXorHash;
          else if (hashes.sha1Hash) hashMatches = sha1Of(join(localDir, rel)) === String(hashes.sha1Hash).toUpperCase();
          const same = remote && remote.size === st.size && hashMatches;
          if (same) {
            unchanged++;
            // Identical content: record the remote's current eTag as the
            // last-synced one (apply only; dry-run never writes state).
            if (opts.apply && remote.eTag) state[remoteRel] = { eTag: remote.eTag, syncedAt: new Date().toISOString() };
            continue;
          }
          const times = {
            remoteModified: remote?.lastModifiedDateTime || null,
            localModified: st.mtime.toISOString(),
          };
          if (!remote) {
            plan.push({ file: rel, remote: remoteRel, size: st.size, changed: false, ...times });
            continue;
          }
          // Both-sides-edited guard: a remote whose eTag differs from the one
          // recorded at the last sync changed since then, even when it looks
          // older. First sync with no stored state falls back to the mtime rule.
          const stored = state[remoteRel];
          if (stored?.eTag && remote.eTag && stored.eTag !== remote.eTag) {
            plan.push({
              file: rel,
              remote: remoteRel,
              size: st.size,
              conflict: "remote changed since the last sync (eTag differs from the stored one)",
              ...times,
            });
            continue;
          }
          // A changed remote file is only replaced when the remote copy is
          // older than the local one; a remote edit newer than local is a
          // conflict that sync never overwrites.
          const remoteNewer = remote.lastModifiedDateTime && new Date(remote.lastModifiedDateTime) > st.mtime;
          if (remoteNewer) {
            plan.push({ file: rel, remote: remoteRel, size: st.size, conflict: "remote is newer than the local file", ...times });
          } else {
            plan.push({ file: rel, remote: remoteRel, size: st.size, changed: true, eTag: remote.eTag, ...times });
          }
        }

        for (const s of skippedEntries) plan.push(s);
        const toUpload = plan.filter((u) => !u.skipped && !u.conflict && (u.changed || u.eTag === undefined));
        if (opts.apply) {
          try {
            await graphRequest(itemPath(base, normalizedTarget));
          } catch (e: any) {
            if (!isNotFound(e)) throw e;
            await ensureFolder(base, normalizedTarget.split("/").filter(Boolean));
          }
          for (const u of toUpload) {
            const relSegments = u.remote.split("/").filter(Boolean);
            const dir = relSegments.slice(0, -1);
            if (dir.length) {
              try {
                await graphRequest(itemPath(base, dir.join("/")));
              } catch (e: any) {
                if (!isNotFound(e)) throw e;
                await ensureFolder(base, dir);
              }
            }
            try {
              // Replace only files we saw older remotes for, and send If-Match
              // with the eTag from the plan so an item that changed after the
              // plan fails with 412 instead of overwriting.
              const r = await uploadFile(base, u.remote, join(localDir, u.file), !!u.changed, u.eTag);
              u.id = r.id;
              u.sha256 = r.sha256;
              u.uploadedVia = r.mode;
              if (r.eTag) state[u.remote] = { eTag: r.eTag, syncedAt: new Date().toISOString() };
            } catch (e: any) {
              if (e instanceof GraphError && e.status === 412) {
                u.conflict = "remote changed during upload (412); skipped";
                continue;
              }
              throw e;
            }
          }
        }

        out({
          drive: opts.drive,
          localDir,
          remote: p,
          mode: opts.apply ? "apply" : "dry-run",
          scanned: localFiles.length,
          uploads: toUpload.filter((u) => !u.conflict).length,
          unchanged,
          conflicts: plan.filter((u) => u.conflict).length,
          files: plan,
        });
        if (opts.apply) writeSyncState(stateFile, state);
      } catch (e: any) {
        graphFail(e);
      }
    });
}
