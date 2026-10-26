import { test, beforeEach, after } from "node:test";
import assert from "node:assert/strict";
import { Command } from "commander";
import { mkdtempSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { createHash } from "crypto";

process.env.OUTLOOK_CLIENT_ID ||= "test";
process.env.OUTLOOK_CLIENT_SECRET ||= "test";
process.env.OUTLOOK_TENANT_ID ||= "test";

const { setTransportForTests, GraphError } = await import("../src/graph.js");
const { registerFiles } = await import("../src/commands/files.js");

type Call = { path: string; method: string };
let calls: Call[] = [];
let lines: string[] = [];
const realLog = console.log;
const realExit = process.exit;

// Offline: every Graph call goes through the stub transport below. No network.
const driveItem = {
  name: "report.pdf",
  id: "ITEM1",
  size: 18,
  file: { hashes: { quickXorHash: "QUICKXORHASH1" } },
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
  setTransportForTests(async (path, options) => {
    calls.push({ path, method: String(options.method || "GET") });
    const q = path.split("?")[0];
    if (q === "/sites/SITE1/drive") return { id: "DRIVE1", name: "Ricky Bureau - GXB" };
    if (q.startsWith("/sites")) return { value: [{ id: "SITE1", displayName: "Ricky Bureau - GXB" }] };
    if (q.includes("/search(q=")) return { value: [searchItem()] };
    if (q === "/me/drive/root/children") return { value: [driveItem] };
    if (q === "/me/drive/root:/10_Products/My%20Folder:/children") return { value: [childItem] };
    if (q === "/drives/DRIVE1/root:/docs:/children") return { value: [searchItem()] };
    if (q === "/drives/DRIVE1/root:/docs/report.pdf") return { ...driveItem, parentReference: { path: "/drives/DRIVE1/root:/docs" } };
    if (q === "/me/drive/root:/10_Products/report.pdf") return driveItem;
    if (q === "/me/drive/root:/10_Products/report.pdf:/content") return new Response("file-body-contents");
    if (q === "/drives/DRIVE1/root:/docs/report.pdf:/content") return new Response("ricky-file-body-contents");
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
  return JSON.parse(lines[0]);
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
    hash: "QUICKXORHASH1",
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
    assert.equal(o.data.localPath, expected);
    assert.equal(o.data.size, Buffer.byteLength("file-body-contents"));
    assert.equal(o.data.sha256, createHash("sha256").update("file-body-contents").digest("hex"));
    assert.equal(readFileSync(expected, "utf8"), "file-body-contents");
    assert.equal(statSync(expected).mode & 0o777, 0o600);
    assert.equal(statSync(join(dir, "me", "10_Products")).mode & 0o777, 0o700);
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
      calls.filter((c) => c.path.startsWith("/drives/")).map((c) => c.path.split("?")[0]),
      [
        "/drives/DRIVE1/root:/docs/report.pdf",
        "/drives/DRIVE1/root:/docs/report.pdf:/content",
        "/drives/DRIVE1/root:/docs:/children",
      ]
    );
    assert.equal(readFileSync(join(dir, "ricky", "docs", "report.pdf"), "utf8"), "ricky-file-body-contents");
    // The cache file holds only the drive id.
    const cached = JSON.parse(readFileSync(process.env.OUTLOOK_DRIVE_CACHE, "utf8"));
    assert.deepEqual(cached, { "Ricky Bureau - GXB": "DRIVE1" });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("ricky drive id is reused from cache without a site search", async () => {
  writeFileSync(process.env.OUTLOOK_DRIVE_CACHE, JSON.stringify({ "Ricky Bureau - GXB": "CACHED1" }));
  await run("info", "docs/report.pdf", "--drive", "ricky");
  assert.equal(calls.filter((c) => c.path.startsWith("/sites")).length, 0);
  assert.equal(calls[0].path, "/drives/CACHED1/root:/docs/report.pdf");
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
