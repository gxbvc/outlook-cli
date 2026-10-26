import { Command } from "commander";
import { chmodSync, closeSync, createWriteStream, existsSync, mkdirSync, openSync, readFileSync, readSync, readdirSync, statSync, writeFileSync } from "fs";
import { Readable } from "stream";
import { pipeline } from "stream/promises";
import { createHash } from "crypto";
import { dirname, join, resolve as resolvePath } from "path";
import { homedir } from "os";
import { TOKENS_DIR } from "../config.js";
import { GraphError, graphRawRequest, graphRequest, out, err } from "../graph.js";
import { QuickXorHash } from "../quickxor.js";

// OneDrive and SharePoint files via Graph. Read only: search, ls, info, get.
// Needs Files.Read.All and Sites.Read.All (delegated); re-run `outlook-cli auth`
// once those scopes are granted. Never deletes or moves anything.

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
  if (cache[SHARED_LIBRARY_NAME]) return cache[SHARED_LIBRARY_NAME];
  const sites = await graphRequest(`/sites?search=${encodeURIComponent(SHARED_LIBRARY_NAME)}`);
  const site = sites?.value?.find((s: any) => s.displayName === SHARED_LIBRARY_NAME) || sites?.value?.[0];
  if (!site) err(`SharePoint site "${SHARED_LIBRARY_NAME}" not found. Check Sites.Read.All consent.`, "SITE_NOT_FOUND");
  const drive = await graphRequest(`/sites/${site.id}/drive`);
  cache[SHARED_LIBRARY_NAME] = drive.id;
  writeDriveCache(cache);
  return drive.id;
}

async function driveBase(drive: string): Promise<string> {
  if (drive === "me") return "/me/drive";
  if (drive === "ricky") return `/drives/${await rickyDriveId()}`;
  err(`Unknown drive "${drive}" (use me or ricky)`, "BAD_DRIVE");
}

// Encode each path segment so spaces, #, % and unicode survive the URL.
function encPath(rel: string): string {
  return rel.split("/").filter(Boolean).map(encodeURIComponent).join("/");
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
  if (status === 429) err(`Throttled (429). Retry after ${e.retryAfter ?? "unknown"} seconds.`, "THROTTLED");
  throw e;
}

function expandTilde(p: string): string {
  return p.replace(/^~(?=\/|$)/, homedir());
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

// Upload one local file. Small: single PUT. Large: an upload session, streamed
// one chunk at a time (never the whole file in memory). Hashes while sending.
async function uploadFile(base: string, rel: string, localPath: string, replace: boolean) {
  const total = statSync(localPath).size;
  const hash = createHash("sha256");
  const enc = encPath(rel);

  if (total <= UPLOAD_SESSION_THRESHOLD) {
    const data = readFileSync(localPath);
    hash.update(data);
    const res = await checkRaw(
      await graphRawRequest(`${base}/root:/${enc}:/content`, {
        method: "PUT",
        headers: { "Content-Type": "application/octet-stream" },
        body: data,
      }),
      rel
    );
    const item = await res.json();
    return { id: item.id, size: total, sha256: hash.digest("hex"), mode: "simple" };
  }

  const session = await graphRequest(`${base}/root:/${enc}:/createUploadSession`, {
    method: "POST",
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": replace ? "replace" : "fail" } }),
  });
  const uploadUrl = session?.uploadUrl;
  if (!uploadUrl) err("Graph returned no upload URL for the session", "UPLOAD_SESSION_FAILED");

  let item: any = null;
  for (let start = 0; start < total; start += UPLOAD_CHUNK_SIZE) {
    const end = Math.min(start + UPLOAD_CHUNK_SIZE, total) - 1;
    const len = end - start + 1;
    const buf = Buffer.alloc(len);
    const fd = openSync(localPath, "r");
    try {
      readSync(fd, buf, 0, len, start);
    } finally {
      closeSync(fd);
    }
    hash.update(buf);
    const chunkRes = await graphRawRequest(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(len),
        "Content-Range": `bytes ${start}-${end}/${total}`,
      },
      body: buf,
    });
    if (!chunkRes.ok) {
      const body = await chunkRes.text();
      throw new GraphError(chunkRes.status, `Graph API ${chunkRes.status} chunk ${start}-${end}: ${body}`);
    }
    const text = await chunkRes.text();
    if (text) {
      try {
        const parsed = JSON.parse(text);
        if (parsed.id) item = parsed; // the final chunk returns the drive item
      } catch {
        // 202 progress responses are JSON, but never fail the upload over parsing
      }
    }
  }
  if (!item) err("Upload finished but Graph returned no drive item", "UPLOAD_INCOMPLETE");
  return { id: item.id, size: total, sha256: hash.digest("hex"), mode: "session" };
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

export function registerFiles(program: Command) {
  const files = program.command("files").description("OneDrive and SharePoint files (read only)");

  files
    .command("search")
    .description("Search files and folders in a drive")
    .argument("<query>", "Search text")
    .option("--drive <drive>", "me (your OneDrive) or ricky (Ricky Bureau - GXB shared library)", "me")
    .action(async (query, opts) => {
      try {
        const base = await driveBase(opts.drive);
        const sq = encodeURIComponent(query.replace(/'/g, "''"));
        const data = await graphRequest(`${base}/search(q='${sq}')?$top=50`);
        const results = (data?.value || []).map(slim);
        out({ query, drive: opts.drive, count: results.length, results });
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
        const enc = encPath(p);
        const url = enc
          ? `${base}/root:/${enc}:/children?$select=name,id,size,folder,file,lastModifiedDateTime,parentReference,webUrl`
          : `${base}/root/children?$select=name,id,size,folder,file,lastModifiedDateTime,parentReference,webUrl`;
        const data = await graphRequest(url);
        const items = (data?.value || []).map(slim);
        out({ drive: opts.drive, path: p, count: items.length, items });
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

        const res = await checkRaw(await graphRawRequest(`${base}/root:/${enc}:/content`), p);
        const cacheRoot = resolvePath(expandTilde(opts.out));
        const segments = p.split("/").filter(Boolean);
        const target = join(cacheRoot, opts.drive, ...segments);
        mkdirSync(dirname(target), { recursive: true, mode: 0o700 });
        chmodSync(dirname(target), 0o700);

        const hash = createHash("sha256");
        let size = 0;
        await pipeline(
          Readable.fromWeb(res.body as any),
          async function* (source) {
            for await (const chunk of source) {
              hash.update(chunk);
              size += chunk.length;
              yield chunk;
            }
          },
          createWriteStream(target, { mode: 0o600 })
        );
        chmodSync(target, 0o600);

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

        const localFiles: string[] = [];
        (function walk(d: string, rel: string) {
          for (const e of readdirSync(join(d, rel), { withFileTypes: true })) {
            if (e.name === ".DS_Store") continue;
            const r = rel ? `${rel}/${e.name}` : e.name;
            if (e.isDirectory()) walk(d, r);
            else if (e.isFile()) localFiles.push(r);
          }
        })(localDir, "");
        localFiles.sort();

        const plan: any[] = [];
        let unchanged = 0;
        for (const rel of localFiles) {
          const remoteRel = p ? `${p.replace(/\/+$/, "")}/${rel}` : rel;
          const st = statSync(join(localDir, rel));
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
          else if (hashes.sha1Hash) {
            const localSha1 = createHash("sha1").update(readFileSync(join(localDir, rel))).digest("hex").toUpperCase();
            hashMatches = localSha1 === String(hashes.sha1Hash).toUpperCase();
          }
          const same = remote && remote.size === st.size && hashMatches;
          if (same) {
            unchanged++;
            continue;
          }
          plan.push({ file: rel, remote: remoteRel, size: st.size, changed: !!remote });
        }

        const toUpload = plan.filter((u) => !u.skipped);
        if (opts.apply) {
          try {
            await graphRequest(itemPath(base, p));
          } catch (e: any) {
            if (!isNotFound(e)) throw e;
            await ensureFolder(base, p.split("/").filter(Boolean));
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
            const r = await uploadFile(base, u.remote, join(localDir, u.file), !!u.changed);
            u.id = r.id;
            u.sha256 = r.sha256;
            u.uploadedVia = r.mode;
          }
        }

        out({
          drive: opts.drive,
          localDir,
          remote: p,
          mode: opts.apply ? "apply" : "dry-run",
          scanned: localFiles.length,
          uploads: toUpload.length,
          unchanged,
          files: plan,
        });
      } catch (e: any) {
        graphFail(e);
      }
    });
}
