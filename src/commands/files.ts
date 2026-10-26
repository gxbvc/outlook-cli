import { Command } from "commander";
import { chmodSync, createWriteStream, mkdirSync, readFileSync, writeFileSync } from "fs";
import { Readable } from "stream";
import { pipeline } from "stream/promises";
import { createHash } from "crypto";
import { dirname, join, resolve as resolvePath } from "path";
import { homedir } from "os";
import { TOKENS_DIR } from "../config.js";
import { GraphError, graphRawRequest, graphRequest, out, err } from "../graph.js";

// OneDrive and SharePoint files via Graph. Read only: search, ls, info, get.
// Needs Files.Read.All and Sites.Read.All (delegated); re-run `outlook-cli auth`
// once those scopes are granted. Never deletes or moves anything.

const DEFAULT_CACHE = join(homedir(), "tmp", "onedrive-cache");
const SHARED_LIBRARY_NAME = "Ricky Bureau - GXB";

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
}
