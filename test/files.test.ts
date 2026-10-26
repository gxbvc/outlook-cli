import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";
import { appendFileSync, existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, realpathSync, rmSync, statSync, symlinkSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { dirname, join } from "path";
import { createHash } from "crypto";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";

const { setTransportForTests, GraphError } = await import("../src/graph.js");
const { registerFiles } = await import("../src/commands/files.js");

type Call = { path: string; method: string; body?: any; contentRange?: string; bodyLen?: number };
let calls: Call[] = [];
let lines: string[] = [];
// Extra GET routes per test: remote drive path (no "/me/drive/root:/" prefix) → drive item.
let remoteItems: Record<string, any> = {};
// Per-test override: content endpoint path → HTTP status to force (412 races).
let putStatus: Record<string, number> = {};
const realLog = console.log;
const realExit = process.exit;

// Offline: every Graph call goes through the stub transport below. No network.
const driveItem = {
  name: "report.pdf",
  id: "ITEM1",
  size: 18,
  file: { hashes: { quickXorHash: "p0V3g8nQAjG8gQx5esEY3uAGOpQ=" } },
  lastModifiedDateTime: "2026-10-26T12:00:00Z",
  parentReference: { path: "/drive/root:/10_Products" },
  webUrl: "https://example.sharepoint.com/report.pdf",
};
const childItem = { ...driveItem, name: "child.txt", parentReference: { path: "/drive/root:/10_Products/My Folder" } };

function searchItem() {
  return { ...driveItem, parentReference: { path: "/drives/DRIVE1/root:/10_Products" } };
}

beforeEach(() => {
  calls = [];
  lines = [];
  delete process.env.OUTLOOK_USER;
  const cacheFile = join(mkdtempSync(join(tmpdir(), "odcache-")), "drive-ids.json");
  process.env.OUTLOOK_DRIVE_CACHE = cacheFile;
  process.env.OUTLOOK_SYNC_STATE_DIR = join(mkdtempSync(join(tmpdir(), "odstate-")), "state");
  remoteItems = {};
  putStatus = {};
  setTransportForTests(async (path, options) => {
    const method = String(options.method || "GET");
    const q = path.split("?")[0];
    const body = options.body;
    const record: Call = {
      path,
      method,
      contentRange: (options.headers as any)?.["Content-Range"],
      ifMatch: (options.headers as any)?.["If-Match"],
      bodyLen: body ? (body as any).length : undefined,
    };
    if (typeof body === "string") {
      try { record.body = JSON.parse(body); } catch { record.body = body; }
    }
    calls.push(record);
    if (q.startsWith("https://upload.example/")) {
      // Chunk PUT. The final chunk (end == total-1) returns the drive item.
      const range = record.contentRange || ""; // "bytes start-end/total"
      const end = parseInt(range.split(" ")[1]?.split("-")[1] || "0", 10);
      const total = parseInt(range.split("/")[1] || "0", 10);
      if (end === total - 1)
        return new Response(JSON.stringify({ name: "big.bin", id: "UPLOADED1", size: total }), {
          headers: { "Content-Type": "application/json" },
        });
      return new Response(JSON.stringify({ nextExpectedRanges: [`${end + 1}-`] }), {
        status: 202,
        headers: { "Content-Type": "application/json" },
      });
    }
    if (q.endsWith(":/createUploadSession") && method === "POST")
      return { uploadUrl: "https://upload.example/session1", nextExpectedRanges: [] };
    if (q.endsWith(":/content") && method === "PUT") {
      if (putStatus[q]) return new Response(JSON.stringify({ error: "forced" }), { status: putStatus[q] });
      return new Response(JSON.stringify({ name: "file", id: "UPLOADED-SIMPLE", size: record.bodyLen, eTag: '"ETAG-UPLOADED-1"' }), {
        headers: { "Content-Type": "application/json" },
      });
    }
    if (q.endsWith("/children") && method === "POST")
      return { name: record.body?.name, id: `FOLDER-CREATED-${record.body?.name}`, folder: {} };
    if (q === "/sites/SITE1/drive") return { id: "DRIVE1", name: "Ricky Bureau - GXB" };
    if (q === "/drives/DRIVE1" || q === "/drives/CACHED1") return { id: q.split("/")[2], name: "Ricky Bureau - GXB" };
    if (q === "/me/drive/root") return { name: "root", id: "ROOT1", folder: {} };
    if (q.startsWith("/sites")) return { value: [{ id: "SITE1", displayName: "Ricky Bureau - GXB" }] };
    if (q.includes("/search(q=")) return { value: [searchItem()] };
    if (q === "/me/drive/root/children") return { value: [driveItem] };
    if (q === "/me/drive/root:/10_Products/My%20Folder:/children") return { value: [childItem] };
    if (q === "/drives/DRIVE1/root:/docs:/children") return { value: [searchItem()] };
    if (q === "/drives/DRIVE1/root:/docs/report.pdf")
      return { ...driveItem, size: 24, file: { hashes: { quickXorHash: "002gYFWLb0aeFt0NPkMLxPAGMuQ=" } }, parentReference: { path: "/drives/DRIVE1/root:/docs" } };
    if (q === "/me/drive/root:/10_Products/report.pdf") return driveItem;
    if (q === "/me/drive/root:/docs" || q === "/me/drive/root:/somedir" || q === "/me/drive/root:/syncdst" || q === "/me/drive/root:/uploads")
      return { name: q.split("/").pop(), id: "FOLDER1", folder: {} };
    if (q === "/me/drive/root:/docs/report.pdf") return driveItem;
    if (q === "/me/drive/root:/10_Products/report.pdf:/content") return new Response("file-body-contents");
    if (q === "/drives/DRIVE1/root:/docs/report.pdf:/content") return new Response("ricky-file-body-contents");
    if (q.startsWith("/me/drive/root:/")) {
      const rel = q.slice("/me/drive/root:/".length);
      if (rel in remoteItems) return remoteItems[rel];
    }
    throw new GraphError(404, "Graph API 404: item not found");
  });
  console.log = (s: string) => { lines.push(String(s)); };
});

after(() => {
  console.log = realLog;
  process.exit = realExit;
  setTransportForTests(null);
});

// `err()` calls process.exit(1); make that observable instead of fatal.
async function run(...args: string[]): Promise<any> {
  const program = new Command();
  program.exitOverride();
  registerFiles(program);
  process.exit = ((code?: number) => {
    throw new Error(`exit ${code}`);
  }) as any;
  try {
    await program.parseAsync(["node", "outlook-cli", "files", ...args]);
  } catch (e: any) {
    if (!/^exit \d+$/.test(e.message)) throw e;
  }
  return JSON.parse(lines[lines.length - 1]);
}

function outObj(): any {
  return JSON.parse(lines[lines.length - 1]);
}

test("ls at root hits /root/children", async () => {
  const o = await run("ls", "/");
  assert.equal(calls[0].path.split("?")[0], "/me/drive/root/children");
  assert.equal(o.ok, true);
  assert.equal(o.data.count, 1);
  assert.equal(o.data.items[0].name, "report.pdf");
  assert.equal(o.data.items[0].type, "file");
});

test("ls encodes spaces in the path", async () => {
  await run("ls", "10_Products/My Folder");
  assert.equal(calls[0].path.split("?")[0], "/me/drive/root:/10_Products/My%20Folder:/children");
});

test("ls encodes unicode, # and %", async () => {
  await run("ls", "A#B 100%/café ünicode");
  assert.equal(calls[0].path.split("?")[0], "/me/drive/root:/A%23B%20100%25/caf%C3%A9%20%C3%BCnicode:/children");
});

test("info shows metadata only (name, id, size, hash), no contents", async () => {
  const o = await run("info", "10_Products/report.pdf");
  assert.equal(calls[0].path, "/me/drive/root:/10_Products/report.pdf");
  assert.deepEqual(o.data, {
    drive: "me",
    name: "report.pdf",
    id: "ITEM1",
    type: "file",
    size: 18,
    path: "10_Products/report.pdf",
    lastModified: "2026-10-26T12:00:00Z",
    hash: "p0V3g8nQAjG8gQx5esEY3uAGOpQ=",
    webUrl: "https://example.sharepoint.com/report.pdf",
  });
});

test("search escapes single quotes and returns slim results", async () => {
  const o = await run("search", "it's a plan");
  assert.equal(calls[0].path, "/me/drive/search(q='it''s%20a%20plan')?$top=50");
  assert.equal(o.data.results[0].path, "10_Products/report.pdf");
  assert.equal(o.data.results[0].size, 18);
});

test("get streams to a temp cache dir with mode 600 and correct sha256", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odget-"));
  try {
    const o = await run("get", "10_Products/report.pdf", "--out", dir);
    const expected = join(dir, "me", "10_Products", "report.pdf");
    assert.equal(o.ok, true);
    assert.equal(o.data.localPath, realpathSync(expected)); // cache root resolved through realpath
    assert.equal(o.data.size, Buffer.byteLength("file-body-contents"));
    assert.equal(o.data.sha256, createHash("sha256").update("file-body-contents").digest("hex"));
    assert.equal(readFileSync(expected, "utf8"), "file-body-contents");
    assert.equal(statSync(expected).mode & 0o777, 0o600);
    assert.equal(statSync(join(dir, "me", "10_Products")).mode & 0o777, 0o700);
    assert.equal(statSync(join(dir, "me")).mode & 0o777, 0o700, "the whole folder chain is 700, not just the leaf");
    assert.equal(statSync(dir).mode & 0o777, 0o700);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("get on a --drive ricky file resolves the shared library drive id once and caches it", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odricky-"));
  try {
    await run("get", "docs/report.pdf", "--drive", "ricky", "--out", dir);
    await run("ls", "docs", "--drive", "ricky");
    const siteCalls = calls.filter((c) => c.path.startsWith("/sites")).length;
    assert.equal(siteCalls, 2, "site search + site drive lookup happen only on first resolve");
    assert.deepEqual(
      calls.filter((c) => c.path.startsWith("/drives/") && c.path.split("?")[0] !== "/drives/DRIVE1").map((c) => c.path.split("?")[0]),
      [
        "/drives/DRIVE1/root:/docs/report.pdf",
        "/drives/DRIVE1/root:/docs/report.pdf:/content",
        "/drives/DRIVE1/root:/docs:/children",
      ]
    );
    assert.equal(calls.filter((c) => c.path.split("?")[0] === "/drives/DRIVE1").length, 1, "verified once per run (only the cached run)");
    assert.equal(readFileSync(join(dir, "ricky", "docs", "report.pdf"), "utf8"), "ricky-file-body-contents");
    // The cache file holds only the drive id.
    const cached = JSON.parse(readFileSync(process.env.OUTLOOK_DRIVE_CACHE, "utf8"));
    assert.deepEqual(cached, { "Ricky Bureau - GXB": "DRIVE1" });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("ricky drive id is reused from cache (verified once) without a site search", async () => {
  writeFileSync(process.env.OUTLOOK_DRIVE_CACHE, JSON.stringify({ "Ricky Bureau - GXB": "CACHED1" }));
  await run("info", "docs/report.pdf", "--drive", "ricky");
  assert.equal(calls.filter((c) => c.path.startsWith("/sites")).length, 0);
  assert.equal(calls[0].path, "/drives/CACHED1", "the cached id is re-verified once per run");
  assert.equal(calls[1].path, "/drives/CACHED1/root:/docs/report.pdf");
});

test("ricky: a cached id whose drive no longer matches is dropped and re-resolved", async () => {
  writeFileSync(process.env.OUTLOOK_DRIVE_CACHE, JSON.stringify({ "Ricky Bureau - GXB": "STALE1" }));
  const o = await run("ls", "docs", "--drive", "ricky");
  assert.equal(o.ok, true);
  assert.equal(calls.some((c) => c.path.startsWith("/sites")), true, "re-resolves after dropping the stale id");
  const cached = JSON.parse(readFileSync(process.env.OUTLOOK_DRIVE_CACHE, "utf8"));
  assert.deepEqual(cached, { "Ricky Bureau - GXB": "DRIVE1" });
});

test("404 becomes a friendly NOT_FOUND error", async () => {
  const o = await run("info", "nope/missing.txt");
  assert.equal(o.ok, false);
  assert.equal(o.code, "NOT_FOUND");
});

test("403 becomes CONSENT_REQUIRED with the scope hint", async () => {
  setTransportForTests(async () => {
    throw new GraphError(403, "Graph API 403: Access denied. Missing scope Files.Read.All");
  });
  const o = await run("ls", "/");
  assert.equal(o.ok, false);
  assert.equal(o.code, "CONSENT_REQUIRED");
  assert.match(o.error, /Files\.Read\.All/);
});

test("429 with Retry-After becomes THROTTLED and names the wait", async () => {
  setTransportForTests(async () => {
    throw new GraphError(429, "Graph API 429: too many requests", 30);
  });
  const o = await run("ls", "/");
  assert.equal(o.ok, false);
  assert.equal(o.code, "THROTTLED");
  assert.match(o.error, /30 seconds/);
});

test("put uploads a small new file with a simple PUT and prints sha256", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odput-"));
  try {
    const local = join(dir, "hello.txt");
    writeFileSync(local, "hello put body");
    const o = await run("put", local, "uploads/hello.txt");
    const contentPut = calls.find((c) => c.method === "PUT" && c.path.split("?")[0].endsWith("/uploads/hello.txt:/content"));
    assert.ok(contentPut, "expected a PUT to the content endpoint");
    assert.equal(o.ok, true);
    assert.equal(o.data.replaced, false);
    assert.equal(o.data.size, 14);
    assert.equal(o.data.sha256, createHash("sha256").update("hello put body").digest("hex"));
    assert.equal(o.data.uploadedVia, "simple");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("put refuses to overwrite an existing file without --force", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odput-"));
  try {
    const local = join(dir, "report.pdf");
    writeFileSync(local, "new body");
    const o = await run("put", local, "docs/report.pdf");
    assert.equal(o.ok, false);
    assert.equal(o.code, "ALREADY_EXISTS");
    assert.equal(calls.some((c) => c.method === "PUT"), false, "no upload may happen");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("put --force replaces an existing file, never a folder", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odput-"));
  try {
    const local = join(dir, "report.pdf");
    writeFileSync(local, "new body");
    const o = await run("put", local, "docs/report.pdf", "--force");
    assert.equal(o.ok, true);
    assert.equal(o.data.replaced, true);
    assert.equal(calls.some((c) => c.method === "PUT"), true);

    writeFileSync(join(dir, "x.txt"), "x");
    const o2 = await run("put", join(dir, "x.txt"), "somedir", "--force");
    assert.equal(o2.ok, false);
    assert.equal(o2.code, "IS_FOLDER");
    assert.equal(calls.some((c) => c.method === "PUT" && c.path.endsWith("/somedir:/content")), false);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("put needs --mkdir for a missing parent and then creates the folders", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odput-"));
  try {
    const local = join(dir, "file.txt");
    writeFileSync(local, "deep");
    const o = await run("put", local, "newdir/deep/file.txt");
    assert.equal(o.code, "PARENT_MISSING");
    assert.equal(calls.some((c) => c.method === "POST"), false);

    const o2 = await run("put", local, "newdir/deep/file.txt", "--mkdir");
    assert.equal(o2.ok, true);
    const created = calls.filter((c) => c.method === "POST" && c.path.endsWith("/children")).map((c) => c.body.name);
    assert.deepEqual(created, ["newdir", "deep"]);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("put uses a chunked upload session over 4 MiB with Content-Range", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odput-"));
  try {
    const local = join(dir, "big.bin");
    // 5.5 MiB: two chunks (5 MiB + 0.5 MiB), both multiples of 320 KiB.
    writeFileSync(local, Buffer.alloc(5 * 1024 * 1024 + 512 * 1024, 7));
    const o = await run("put", local, "uploads/big.bin");
    assert.equal(o.ok, true);
    assert.equal(o.data.uploadedVia, "session");
    const session = calls.find((c) => c.method === "POST" && c.path.endsWith("/uploads/big.bin:/createUploadSession"));
    assert.ok(session, "expected createUploadSession");
    assert.equal(session.body.item["@microsoft.graph.conflictBehavior"], "fail");
    const chunks = calls.filter((c) => c.path.startsWith("https://upload.example/"));
    assert.equal(chunks.length, 2);
    assert.equal(chunks[0].contentRange, "bytes 0-5242879/5767168");
    assert.equal(chunks[1].contentRange, "bytes 5242880-5767167/5767168");
    assert.equal(o.data.sha256, createHash("sha256").update(readFileSync(local)).digest("hex"));
    // No bearer token may be sent to the preauthenticated upload URL.
    assert.equal(o.data.id, "UPLOADED1");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("sync dry-run plans uploads (new, changed, unchanged, folder) without uploading", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odsync-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "unchanged.txt"), "same");
    writeFileSync(join(src, "changed.txt"), "new content");
    writeFileSync(join(src, "new.txt"), "brand new");
    writeFileSync(join(src, "folder.txt"), "collides with a remote folder");
    const sha1 = (s: string) => createHash("sha1").update(s).digest("hex").toUpperCase();
    remoteItems = {
      "syncdst/unchanged.txt": { name: "unchanged.txt", id: "U1", size: 4, file: { hashes: { sha1Hash: sha1("same") } } },
      "syncdst/changed.txt": { name: "changed.txt", id: "C1", size: 11, file: { hashes: { sha1Hash: sha1("old content") } } },
      "syncdst/folder.txt": { name: "folder.txt", id: "F1", folder: {} },
    };
    const o = await run("sync", src, "syncdst", "--dry-run");
    assert.equal(o.ok, true);
    assert.equal(o.data.mode, "dry-run");
    assert.equal(o.data.scanned, 4);
    assert.equal(o.data.uploads, 2);
    assert.equal(o.data.unchanged, 1);
    assert.deepEqual(o.data.files.map((f: any) => f.file), ["changed.txt", "folder.txt", "new.txt"]);
    assert.equal(o.data.files[0].changed, true);
    assert.equal(o.data.files[1].skipped, "remote path is a folder; never replaced");
    assert.equal(o.data.files[2].changed, false);
    assert.equal(calls.some((c) => c.method === "PUT" || c.method === "POST"), false, "dry-run uploads nothing");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("sync apply uploads only new and changed files, never deletes remote", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odsync-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "unchanged.txt"), "same");
    writeFileSync(join(src, "changed.txt"), "new content");
    writeFileSync(join(src, "new.txt"), "brand new");
    const sha1 = (s: string) => createHash("sha1").update(s).digest("hex").toUpperCase();
    remoteItems = {
      "syncdst/unchanged.txt": { name: "unchanged.txt", id: "U1", size: 4, file: { hashes: { sha1Hash: sha1("same") } } },
      "syncdst/changed.txt": { name: "changed.txt", id: "C1", size: 11, file: { hashes: { sha1Hash: sha1("old content") } } },
    };
    const o = await run("sync", src, "syncdst", "--apply");
    assert.equal(o.ok, true);
    assert.equal(o.data.uploads, 2);
    assert.equal(o.data.unchanged, 1);
    const puts = calls
      .filter((c) => c.method === "PUT" && c.path.split("?")[0].endsWith(":/content"))
      .map((c) => c.path.split("?")[0]);
    assert.deepEqual(puts, [
      "/me/drive/root:/syncdst/changed.txt:/content",
      "/me/drive/root:/syncdst/new.txt:/content",
    ]);
    const o2 = await run("sync", src, "syncdst", "--dry-run");
    assert.deepEqual(
      o2.data.files.filter((f: any) => !f.skipped).map((f: any) => [f.file, f.changed]),
      [["changed.txt", true], ["new.txt", false]]
    );
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("sync requires exactly one of --dry-run or --apply", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odsync-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    const o = await run("sync", src, "syncdst");
    assert.equal(o.code, "BAD_MODE");
    const o2 = await run("sync", src, "syncdst", "--dry-run", "--apply");
    assert.equal(o2.code, "BAD_MODE");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

// Pinned QuickXorHash vectors, computed with the reference implementation
// (@cloudraker/quickxorhash 1.1.0, a port of Microsoft's published
// quickxorhash algorithm) on 2026-10-26. The port in src/quickxor.ts must
// reproduce these exactly, or sync's change detection breaks.
test("QuickXorHash matches Microsoft's published algorithm (pinned vectors)", async () => {
  const { QuickXorHash } = await import("../src/quickxor.js");
  const hash = (s: string) => { const q = new QuickXorHash(); q.update(Buffer.from(s, "utf8")); return q.digestBase64(); };
  assert.equal(hash(""), "AAAAAAAAAAAAAAAAAAAAAAAAAAA=");
  assert.equal(hash("hello world"), "aCgDG9jwBhDc4Q1yawMZAAAAAAA=");
  assert.equal(hash("The quick brown fox jumps over the lazy dog"), "bMSlbysmxJL6S75XwfMcQZOpcr4=");
  // Streaming (64 KiB chunks) must agree with one-shot on a 5 MiB + 777 byte input.
  const big = Buffer.alloc(5 * 1024 * 1024 + 777);
  for (let i = 0; i < big.length; i++) big[i] = (i * 31 + 17) % 251;
  const one = new QuickXorHash(); one.update(big);
  const streamed = new QuickXorHash();
  for (let off = 0; off < big.length; off += 65536) streamed.update(big.subarray(off, Math.min(off + 65536, big.length)));
  assert.equal(one.digestBase64(), streamed.digestBase64());
});

test("sync compares quickXorHash (business drives) and falls back to sha1", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odqxh-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "same.txt"), "same");
    writeFileSync(join(src, "differs.txt"), "same");
    writeFileSync(join(src, "sha1only.txt"), "same");
    const { QuickXorHash } = await import("../src/quickxor.js");
    const q = new QuickXorHash(); q.update(Buffer.from("same"));
    const qxh = q.digestBase64();
    const sha1 = (s: string) => createHash("sha1").update(s).digest("hex").toUpperCase();
    remoteItems = {
      "syncdst/same.txt": { name: "same.txt", id: "S1", size: 4, file: { hashes: { quickXorHash: qxh } } },
      "syncdst/differs.txt": { name: "differs.txt", id: "D1", size: 4, file: { hashes: { quickXorHash: "WRONGQXH=" } } },
      // personal OneDrive shape: sha1Hash only
      "syncdst/sha1only.txt": { name: "sha1only.txt", id: "H1", size: 4, file: { hashes: { sha1Hash: sha1("same") } } },
      // no hashes at all: counts as changed (re-upload)
      "syncdst/nohash.txt": { name: "nohash.txt", id: "N1", size: 4, file: { hashes: {} } },
    };
    writeFileSync(join(src, "nohash.txt"), "same");
    const o = await run("sync", src, "syncdst", "--dry-run");
    assert.equal(o.data.unchanged, 2, "same.txt (quickXorHash) and sha1only.txt (sha1 fallback) are unchanged");
    assert.deepEqual(
      o.data.files.filter((f: any) => !f.skipped).map((f: any) => f.file).sort(),
      ["differs.txt", "nohash.txt"]
    );
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

// --- Review round 1 fixes ---

test("M1: dot-dot segments are refused before any Graph call", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odm1-"));
  try {
    const o = await run("get", "../../../drives/OTHER/root:/f", "--out", dir);
    assert.equal(o.code, "BAD_PATH");
    assert.equal(calls.length, 0, "no Graph request may leave the machine");
    const o2 = await run("info", "a/../b");
    assert.equal(o2.code, "BAD_PATH");
    const o3 = await run("ls", "a/../b");
    assert.equal(o3.code, "BAD_PATH");
    writeFileSync(join(dir, "local.txt"), "x");
    const o4 = await run("put", join(dir, "local.txt"), "a/../../escape.txt");
    assert.equal(o4.code, "BAD_PATH");
    const o5 = await run("sync", dir, "dst/../escape", "--dry-run");
    assert.equal(o5.code, "BAD_PATH");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("M1: backslashes and control characters are refused", async () => {
  const o = await run("info", "a\\b.txt");
  assert.equal(o.code, "BAD_PATH");
  const o2 = await run("ls", "a\u0001b");
  assert.equal(o2.code, "BAD_PATH");
});

test("M1: get refuses a symlink at the target instead of writing through it", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odsym-"));
  try {
    const target = join(dir, "me", "10_Products", "report.pdf");
    mkdirSync(dirname(target), { recursive: true });
    symlinkSync("/etc/passwd", target);
    const o = await run("get", "10_Products/report.pdf", "--out", dir);
    assert.equal(o.code, "PATH_ESCAPE");
    assert.equal(readFileSync("/etc/passwd", "utf8").includes("file-body-contents"), false);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("M2: the simple PUT always carries @microsoft.graph.conflictBehavior", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odm2-"));
  try {
    const local = join(dir, "hello.txt");
    writeFileSync(local, "hello put body");
    const o = await run("put", local, "uploads/hello.txt");
    assert.equal(o.ok, true);
    const put1 = calls.find((c) => c.method === "PUT" && c.path.split("?")[0].endsWith("/uploads/hello.txt:/content"));
    assert.match(put1!.path, /\?@microsoft\.graph\.conflictBehavior=fail$/);
    const o2 = await run("put", local, "docs/report.pdf", "--force");
    assert.equal(o2.ok, true);
    const put2 = calls.filter((c) => c.method === "PUT" && c.path.split("?")[0].endsWith("/docs/report.pdf:/content")).pop();
    assert.match(put2!.path, /\?@microsoft\.graph\.conflictBehavior=replace$/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("M3: sync --apply replaces only files whose remote copy is older; If-Match sent; 412 is a conflict", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odm3-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "older.txt"), "older remote");
    writeFileSync(join(src, "newer.txt"), "newer remote");
    writeFileSync(join(src, "race.txt"), "raced file");
    writeFileSync(join(src, "fresh.txt"), "brand new");
    remoteItems = {
      // remote older than local: replaced, with If-Match from the plan
      "syncdst/older.txt": {
        name: "older.txt", id: "OLD1", size: 99, eTag: "\"ETAG-OLD1\"",
        lastModifiedDateTime: "2000-01-01T00:00:00Z", file: { hashes: { quickXorHash: "WRONG=" } },
      },
      // remote newer than local: CONFLICT, never replaced
      "syncdst/newer.txt": {
        name: "newer.txt", id: "NEW1", size: 99, eTag: "\"ETAG-NEW1\"",
        lastModifiedDateTime: "2030-01-01T00:00:00Z", file: { hashes: { quickXorHash: "WRONG=" } },
      },
      // remote older but the PUT races with a 412: skipped as conflict
      "syncdst/race.txt": {
        name: "race.txt", id: "RACE1", size: 99, eTag: "\"ETAG-RACE1\"",
        lastModifiedDateTime: "2000-01-01T00:00:00Z", file: { hashes: { quickXorHash: "WRONG=" } },
      },
    };
    putStatus["/me/drive/root:/syncdst/race.txt:/content"] = 412;
    const o = await run("sync", src, "syncdst", "--apply");
    const entry = (n: string) => o.data.files.find((f: any) => f.file === n);
    assert.equal(entry("newer.txt").conflict, "remote is newer than the local file");
    assert.equal(o.data.conflicts, 2); // newer.txt + the 412 on race.txt
    assert.equal(o.data.uploads, 2); // older.txt + fresh.txt
    const olderPut = calls.find((c) => c.method === "PUT" && c.path.split("?")[0].endsWith("/syncdst/older.txt:/content"));
    assert.equal(olderPut!.ifMatch, "\"ETAG-OLD1\"");
    assert.equal(calls.some((c) => c.method === "PUT" && c.path.split("?")[0].endsWith("/syncdst/newer.txt:/content")), false);
    assert.equal(entry("race.txt").conflict, "remote changed during upload (412); skipped");
    assert.equal(entry("fresh.txt").uploadedVia, "simple");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("M4: --drive ricky needs an exact library name and caches nothing on miss", async () => {
  setTransportForTests(async (path) => {
    const q = path.split("?")[0];
    if (q.startsWith("/sites")) return { value: [{ id: "SITE9", displayName: "Wrong Library" }] };
    throw new GraphError(404, "Graph API 404: should never get past site search");
  });
  const o = await run("info", "docs/report.pdf", "--drive", "ricky");
  assert.equal(o.code, "SITE_NOT_FOUND");
  assert.equal(calls.some((c) => c.path.startsWith("/drives/")), false, "no drive may be picked as a fallback");
  assert.equal(existsSync(process.env.OUTLOOK_DRIVE_CACHE), false, "nothing cached on a miss");
});

test("SHOULD: a throttled chunk is retried with Retry-After, then the session is cancelled", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odretry-"));
  try {
    const local = join(dir, "big.bin");
    writeFileSync(local, Buffer.alloc(5 * 1024 * 1024 + 512 * 1024, 7));
    let chunkAttempts = 0;
    setTransportForTests(async (path, options) => {
      const method = String(options.method || "GET");
      calls.push({ path, method });
      const q = path.split("?")[0];
      if (q.startsWith("https://upload.example/")) {
        if (method === "DELETE") return new Response(null, { status: 204 });
        chunkAttempts++;
        return new Response(JSON.stringify({ error: "throttled" }), { status: 429, headers: { "Retry-After": "0" } });
      }
      if (q === "/me/drive/root:/uploads") return { name: "uploads", id: "FOLDER1", folder: {} };
      if (q.endsWith(":/createUploadSession") && method === "POST")
        return { uploadUrl: "https://upload.example/session1" };
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    const o = await run("put", local, "uploads/big.bin");
    assert.equal(o.code, "THROTTLED");
    assert.equal(chunkAttempts, 4, "1 attempt + 3 retries");
    assert.match(o.error, /after 3 retries/);
    assert.equal(calls.some((c) => c.method === "DELETE" && c.path.startsWith("https://upload.example/")), true, "session cancelled");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: a throttled chunk succeeds on retry", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odretry2-"));
  try {
    const local = join(dir, "big.bin");
    writeFileSync(local, Buffer.alloc(5 * 1024 * 1024 + 512 * 1024, 7));
    let chunkAttempts = 0;
    setTransportForTests(async (path, options) => {
      const method = String(options.method || "GET");
      const q = path.split("?")[0];
      if (q.startsWith("https://upload.example/")) {
        chunkAttempts++;
        if (chunkAttempts === 1) return new Response(JSON.stringify({ e: 1 }), { status: 429, headers: { "Retry-After": "0" } });
        if (chunkAttempts < 3) return new Response(JSON.stringify({ nextExpectedRanges: ["1-"] }), { status: 202 });
        return new Response(JSON.stringify({ name: "big.bin", id: "UPLOADED1", size: 5 * 1024 * 1024 + 512 * 1024 }));
      }
      if (q === "/me/drive/root:/uploads") return { name: "uploads", id: "FOLDER1", folder: {} };
      if (q.endsWith(":/createUploadSession") && method === "POST") return { uploadUrl: "https://upload.example/session1" };
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    const o = await run("put", local, "uploads/big.bin");
    assert.equal(o.ok, true);
    assert.equal(chunkAttempts, 3);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: a failed download leaves no partial file or .part leftover", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odpart-"));
  try {
    setTransportForTests(async (path) => {
      const q = path.split("?")[0];
      if (q === "/me/drive/root:/10_Products/report.pdf") return driveItem;
      if (q.endsWith(":/content")) {
        return new Response(
          new ReadableStream({
            start(controller) {
              controller.enqueue(Buffer.from("partial"));
              controller.error(new Error("connection dropped"));
            },
          })
        );
      }
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    let o: any;
    try {
      o = await run("get", "10_Products/report.pdf", "--out", dir);
    } catch (e: any) {
      o = { ok: false, error: e.message }; // the network error surfaces through parseAsync
    }
    assert.match(String(o.error), /connection dropped/);
    const target = join(dir, "me", "10_Products", "report.pdf");
    assert.equal(existsSync(target), false);
    const leftovers = (function walk(d: string): string[] {
      return readdirSync(d, { withFileTypes: true }).flatMap((e: any) =>
        e.isDirectory() ? walk(join(d, e.name)) : e.name.includes(".part") ? [join(d, e.name)] : []
      );
    })(dir);
    assert.deepEqual(leftovers, []);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: ls follows @odata.nextLink", async () => {
  setTransportForTests(async (path) => {
    const q = path.split("?")[0];
    if (q === "/me/drive/root/children") {
      return { value: [{ name: "p1.txt", id: "1", size: 1, file: {} }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/drive/root/children?page=2" };
    }
    if (path.includes("page=2")) return { value: [{ name: "p2.txt", id: "2", size: 2, file: {} }] };
    throw new GraphError(404, "Graph API 404: unexpected " + q);
  });
  const o = await run("ls", "/");
  assert.equal(o.data.count, 2);
  assert.deepEqual(o.data.items.map((i: any) => i.name), ["p1.txt", "p2.txt"]);
});

test("SHOULD: sync reports case-only collisions and symlinks, never uploads them", async () => {
  // APFS here is case-insensitive, so the two spellings cannot both exist on
  // disk; the collision rule itself is a pure function, tested directly.
  const { detectCaseCollisions } = await import("../src/commands/files.js");
  assert.deepEqual(
    [...detectCaseCollisions(["A.txt", "a.txt", "plain.txt"])].sort(),
    ["A.txt", "a.txt"]
  );
  assert.equal(detectCaseCollisions(["a.txt", "b.txt"]).size, 0);

  const dir = mkdtempSync(join(tmpdir(), "odcase-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "a.txt"), "plain");
    symlinkSync(join(src, "a.txt"), join(src, "link.txt"));
    writeFileSync(join(src, "plain.txt"), "plain");
    const o = await run("sync", src, "syncdst", "--dry-run");
    assert.equal(o.data.files.find((f: any) => f.file === "link.txt").skipped, "symlink");
    assert.equal(o.data.uploads, 2);
    assert.equal(o.data.conflicts, 0);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: the bearer token never goes on an absolute (upload session) URL", async () => {
  const { authHeadersFor } = await import("../src/graph.js");
  assert.deepEqual(authHeadersFor("https://upload.example/session1", "SECRET"), {});
  assert.deepEqual(authHeadersFor("/me/drive/root:/f", "SECRET"), { Authorization: "Bearer SECRET" });
});

test("SHOULD: a local file that changes size during upload cancels the session", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odgrow-"));
  try {
    const local = join(dir, "big.bin");
    writeFileSync(local, Buffer.alloc(5 * 1024 * 1024 + 512 * 1024, 7));
    setTransportForTests(async (path, options) => {
      const method = String(options.method || "GET");
      calls.push({ path, method });
      const q = path.split("?")[0];
      if (q.startsWith("https://upload.example/")) {
        if (method === "DELETE") return new Response(null, { status: 204 });
        // Simulate another process appending to the file mid-upload.
        appendFileSync(local, "grown");
        return new Response(JSON.stringify({ nextExpectedRanges: ["1-"] }), { status: 202 });
      }
      if (q === "/me/drive/root:/uploads") return { name: "uploads", id: "FOLDER1", folder: {} };
      if (q.endsWith(":/createUploadSession") && method === "POST") return { uploadUrl: "https://upload.example/session1" };
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    const o = await run("put", local, "uploads/big.bin");
    assert.equal(o.code, "LOCAL_FILE_CHANGED");
    assert.equal(calls.some((c) => c.method === "DELETE" && c.path.startsWith("https://upload.example/")), true, "session cancelled");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("NIT: files refuses the global --user remap", async () => {
  process.env.OUTLOOK_USER = "someone@example.com";
  try {
    const o = await run("ls", "/");
    assert.equal(o.code, "USER_UNSUPPORTED");
  } finally {
    delete process.env.OUTLOOK_USER;
  }
});

// --- Review round 2 fixes ---

test("R2-M1: get refuses a symlinked drive folder and never chmods through it", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odr2m1-"));
  try {
    const outside = join(dir, "outside");
    mkdirSync(outside, { mode: 0o755 });
    symlinkSync(outside, join(dir, "me")); // cache/me -> outside
    const o = await run("get", "10_Products/report.pdf", "--out", dir);
    assert.equal(o.code, "PATH_ESCAPE");
    assert.equal(existsSync(join(outside, "10_Products")), false, "nothing created through the link");
    assert.equal(statSync(outside).mode & 0o777, 0o755, "the outside folder is not chmodded");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("R2-M1: get refuses a symlinked middle folder in the cache", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odr2m1b-"));
  try {
    const outside = join(dir, "outside");
    mkdirSync(outside, { mode: 0o755 });
    mkdirSync(join(dir, "me"), { recursive: true });
    symlinkSync(outside, join(dir, "me", "10_Products"));
    const o = await run("get", "10_Products/report.pdf", "--out", dir);
    assert.equal(o.code, "PATH_ESCAPE");
    assert.equal(existsSync(join(outside, "10_Products")), false);
    assert.equal(statSync(outside).mode & 0o777, 0o755);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("R2-M2: one leading slash is allowed on every subcommand", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odr2m2-"));
  try {
    const local = join(dir, "hello.txt");
    writeFileSync(local, "lead");
    const o1 = await run("info", "/10_Products/report.pdf");
    assert.equal(o1.ok, true);
    const o2 = await run("get", "/10_Products/report.pdf", "--out", dir);
    assert.equal(o2.ok, true);
    const o3 = await run("put", local, "/uploads/hello.txt");
    assert.equal(o3.ok, true);
    const o4 = await run("ls", "/10_Products/My Folder");
    assert.equal(o4.ok, true);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("R2-M2: a trailing slash is stripped; dry-run and apply agree", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odr2m2b-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "plain.txt"), "plain");
    const o = await run("sync", src, "syncdst/", "--dry-run");
    assert.equal(o.ok, true);
    assert.equal(o.data.uploads, 1);
    const o2 = await run("sync", src, "syncdst/", "--apply");
    assert.equal(o2.ok, true);
    assert.deepEqual(o2.data.files.map((f: any) => f.remote), ["syncdst/plain.txt"]);
    // dot-dot through a trailing slash is still refused
    const o3 = await run("sync", src, "syncdst/..", "--dry-run");
    assert.equal(o3.code, "BAD_PATH");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("R2-M2: sync to the drive root (/) plans and applies the same way", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odr2root-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "plain.txt"), "plain");
    const dry = await run("sync", src, "/", "--dry-run");
    assert.equal(dry.ok, true);
    assert.equal(dry.data.uploads, 1);
    assert.equal(dry.data.files[0].remote, "plain.txt");
    const apply = await run("sync", src, "/", "--apply");
    assert.equal(apply.ok, true);
    assert.equal(apply.data.uploads, 1);
    assert.deepEqual(
      calls.filter((c) => c.method === "PUT" && c.path.split("?")[0].endsWith(":/content")).map((c) => c.path.split("?")[0]),
      ["/me/drive/root:/plain.txt:/content"]
    );
    // dry-run and apply agree on the plan
    assert.deepEqual(
      apply.data.files.map((f: any) => [f.file, f.remote, f.changed]),
      dry.data.files.map((f: any) => [f.file, f.remote, f.changed])
    );
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: one bad local file name is skipped and reported, the rest sync", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odbad-"));
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "good.txt"), "good");
    writeFileSync(join(src, "bad\\name.txt"), "bad"); // backslash is legal on macOS
    const o = await run("sync", src, "syncdst", "--dry-run");
    assert.equal(o.ok, true);
    const bad = o.data.files.find((f: any) => f.file === "bad\\name.txt");
    assert.match(bad.skipped, /bad file name/);
    assert.equal(o.data.files.find((f: any) => f.file === "good.txt").changed, false);
    assert.equal(o.data.uploads, 1);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: the sync state file stores eTags; a remote changed since the last sync is a conflict", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odstate-"));
  const stateDir = join(dir, "state");
  process.env.OUTLOOK_SYNC_STATE_DIR = stateDir;
  try {
    const src = join(dir, "src");
    mkdirSync(src);
    writeFileSync(join(src, "doc.txt"), "local edit");
    remoteItems = {
      // remote looks OLD (mtime rule would replace) and has a matching-content hash? no:
      // content differs (changed), remote mtime old, eTag known
      "syncdst/doc.txt": {
        name: "doc.txt", id: "D1", size: 99, eTag: "\"ETAG-D1\"",
        lastModifiedDateTime: "2000-01-01T00:00:00Z", file: { hashes: { quickXorHash: "WRONG=" } },
      },
    };
    // First apply: no stored state, mtime rule applies -> replaced, state written.
    const o1 = await run("sync", src, "syncdst", "--apply");
    assert.equal(o1.data.conflicts, 0);
    const stateFile = readdirSync(stateDir)[0];
    assert.match(stateFile, /^me-/.test(stateFile) ? /^me-/ : /./); // drive-prefixed name
    const state = JSON.parse(readFileSync(join(stateDir, stateFile), "utf8"));
    assert.equal(state["syncdst/doc.txt"].eTag, "\"ETAG-UPLOADED-1\"");
    assert.equal(statSync(join(stateDir, stateFile)).mode & 0o777, 0o600);

    // Second apply: remote STILL has the old mtime but a DIFFERENT eTag than
    // the stored one -> both sides edited -> CONFLICT, never replaced.
    remoteItems["syncdst/doc.txt"] = {
      ...remoteItems["syncdst/doc.txt"],
      eTag: "\"ETAG-D2\"", // Ricky edited it after our sync
    };
    const putsBefore = calls.filter((c) => c.method === "PUT").length;
    const o2 = await run("sync", src, "syncdst", "--apply");
    const entry = o2.data.files.find((f: any) => f.file === "doc.txt");
    assert.match(entry.conflict, /changed since the last sync/);
    assert.equal(o2.data.conflicts, 1);
    assert.equal(calls.filter((c) => c.method === "PUT").length, putsBefore, "the conflicted file is never uploaded");

    // dry-run never writes state
    rmSync(stateDir, { recursive: true, force: true });
    await run("sync", src, "syncdst", "--dry-run");
    assert.equal(existsSync(stateDir), false);
  } finally {
    delete process.env.OUTLOOK_SYNC_STATE_DIR;
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: a network error on a chunk is retried, then the session cancelled", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odnet-"));
  try {
    const local = join(dir, "big.bin");
    writeFileSync(local, Buffer.alloc(5 * 1024 * 1024 + 512 * 1024, 7));
    let chunkAttempts = 0;
    setTransportForTests(async (path, options) => {
      const method = String(options.method || "GET");
      calls.push({ path, method });
      const q = path.split("?")[0];
      if (q.startsWith("https://upload.example/")) {
        if (method === "DELETE") return new Response(null, { status: 204 });
        chunkAttempts++;
        if (chunkAttempts <= 2) throw new TypeError("fetch failed"); // network errors
        if (chunkAttempts < 4) return new Response(JSON.stringify({ nextExpectedRanges: ["1-"] }), { status: 202 });
        return new Response(JSON.stringify({ name: "big.bin", id: "UP1", size: 5 * 1024 * 1024 + 512 * 1024 }));
      }
      if (q === "/me/drive/root:/uploads") return { id: "F1", folder: {} };
      if (q.endsWith(":/createUploadSession") && method === "POST") return { uploadUrl: "https://upload.example/session1" };
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    const o = await run("put", local, "uploads/big.bin");
    assert.equal(o.ok, true);
    assert.equal(chunkAttempts, 4);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("SHOULD: Retry-After is honored in full up to a 60 s cap, HTTP-dates parsed", async () => {
  const { retryDelayMs } = await import("../src/commands/files.js");
  assert.equal(retryDelayMs("120", 1), 60000, "capped at 60 s");
  assert.equal(retryDelayMs("2", 1), 2000);
  assert.equal(retryDelayMs(null, 1), 400, "backoff fallback");
  assert.equal(retryDelayMs(null, 3), 1600);
  const inAnHour = new Date(Date.now() + 3600_000).toUTCString();
  assert.equal(retryDelayMs(inAnHour, 1), 60000, "HTTP-date capped at 60 s");
  const in5s = new Date(Date.now() + 5000).toUTCString();
  const d = retryDelayMs(in5s, 1);
  assert.ok(d > 3000 && d <= 5000, `HTTP-date parsed to a wait near 5 s, got ${d}`);
});

test("SHOULD: the simple PUT is hash-checked against the returned item", async () => {
  const dir = mkdtempSync(join(tmpdir(), "odhash-"));
  try {
    const local = join(dir, "hello.txt");
    writeFileSync(local, "hello put body");
    setTransportForTests(async (path, options) => {
      const method = String(options.method || "GET");
      calls.push({ path, method });
      const q = path.split("?")[0];
      if (q === "/me/drive/root:/uploads") return { id: "F1", folder: {} };
      if (q.endsWith(":/uploads/hello.txt:/content") && method === "PUT")
        return new Response(JSON.stringify({ id: "X1", file: { hashes: { quickXorHash: "CORRUPTED=" } } }));
      throw new GraphError(404, "Graph API 404: unexpected " + q);
    });
    const o = await run("put", local, "uploads/hello.txt");
    assert.equal(o.code, "UPLOAD_HASH_MISMATCH");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
