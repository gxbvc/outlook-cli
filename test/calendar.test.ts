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
const { registerCalendar } = await import("../src/commands/calendar.js");

type Call = { path: string; method: string; body: any; headers: any };
let calls: Call[] = [];
let lines: string[] = [];
const realLog = console.log;

beforeEach(() => {
  calls = [];
  lines = [];
  delete process.env.OUTLOOK_USER;
  setTransportForTests(async (path, options) => {
    calls.push({
      path,
      method: String(options.method || "GET"),
      body: options.body ? JSON.parse(String(options.body)) : undefined,
      headers: options.headers,
    });
    if (path === "/me/calendar/getSchedule") {
      return { value: [{ scheduleId: "room@venue.example.test", availabilityView: "0020", scheduleItems: [] }] };
    }
    return { id: "EVT1" };
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
  registerCalendar(program);
  await program.parseAsync(["node", "outlook-cli", "calendar", ...args]);
}

test("schedule posts getSchedule with the addresses, window, timezone and interval", async () => {
  await run("schedule", "--addresses", "room@venue.example.test, other@venue.example.test", "--start", "2026-11-05T09:00:00",
    "--end", "2026-11-05T11:00:00", "--tz", "America/Chicago");
  assert.equal(calls.length, 1);
  assert.equal(calls[0].method, "POST");
  assert.equal(calls[0].path, "/me/calendar/getSchedule");
  assert.deepEqual(calls[0].body, {
    schedules: ["room@venue.example.test", "other@venue.example.test"],
    startTime: { dateTime: "2026-11-05T09:00:00", timeZone: "America/Chicago" },
    endTime: { dateTime: "2026-11-05T11:00:00", timeZone: "America/Chicago" },
    availabilityViewInterval: 30,
  });
  const result = JSON.parse(lines[0]);
  assert.equal(result.ok, true);
  assert.equal(result.data.schedules[0].availabilityView, "0020");
});

test("create --resources adds the room as a resource attendee; no body when none is given", async () => {
  await run("create", "--subject", "Owner - Team meeting", "--start", "2026-11-05T09:00:00",
    "--end", "2026-11-05T10:00:00", "--tz", "America/Chicago", "--location", "Room A",
    "--resources", "room@venue.example.test");
  assert.equal(calls[0].path, "/me/events");
  assert.deepEqual(calls[0].body.attendees, [{ emailAddress: { address: "room@venue.example.test" }, type: "resource" }]);
  assert.equal(calls[0].body.body, undefined);
  assert.equal(calls[0].body.isOnlineMeeting, undefined);
});

test("forward posts /events/{id}/forward with each address once and the agent footer on the comment", async () => {
  await run("forward", "--id", "EVT9", "--to", "guest@corp.example.test, a@x.example.test", "--to", "A@X.example.test",
    "--comment", "A teammate will join to listen.");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), ["POST /me/events/EVT9/forward"]);
  assert.deepEqual(calls[0].body, {
    toRecipients: [{ emailAddress: { address: "guest@corp.example.test" } }, { emailAddress: { address: "a@x.example.test" } }],
    comment: "A teammate will join to listen.\n\nSent from my agent \u{1F916}",
  });
  const result = JSON.parse(lines[0]);
  assert.equal(result.ok, true);
  assert.deepEqual(result.data.to, ["guest@corp.example.test", "a@x.example.test"]);
});

test("forward without --comment still sends the agent footer", async () => {
  await run("forward", "--id", "EVT9", "--to", "guest@corp.example.test");
  assert.equal(calls[0].body.comment, "Sent from my agent \u{1F916}");
});

test("cancel notifies attendees through POST /events/{id}/cancel (never a silent delete)", async () => {
  await run("cancel", "--id", "EVT1");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), ["POST /me/events/EVT1/cancel"]);
  assert.deepEqual(calls[0].body, { comment: "" });
});

// ---- update: attendee changes merge, removal needs --yes ----

const guest = (address: string, type = "required", response = "accepted") =>
  ({ emailAddress: { address, name: address }, type, status: { response, time: "2026-11-09T10:00:00Z" } });

function serveEvent(event: any | Error) {
  setTransportForTests(async (path, options) => {
    const method = String(options.method || "GET");
    calls.push({ path, method, body: options.body ? JSON.parse(String(options.body)) : undefined, headers: options.headers });
    if (method === "GET" && path.startsWith("/me/events/EVT1")) {
      if (event instanceof Error) throw event;
      return event;
    }
    return { id: "EVT1" };
  });
}

const threeGuests = () => ({
  organizer: { emailAddress: { address: "owner@corp.example.test" } },
  attendees: [guest("a@x.example.test"), guest("b@x.example.test", "required", "tentativelyAccepted"), guest("c@x.example.test", "required", "none")],
});
const patches = () => calls.filter((c) => c.method === "PATCH");
const addrs = (list: any[]) => list.map((a) => a.emailAddress.address);

class ExitSignal extends Error {}
async function runExpectingExit(...args: string[]) {
  const realExit = process.exit;
  (process as any).exit = (code?: number) => { throw new ExitSignal(String(code)); };
  try {
    await assert.rejects(run(...args), ExitSignal);
  } finally {
    process.exit = realExit;
  }
}
async function runExpectingNonZero(...args: string[]) {
  const before = process.exitCode;
  await run(...args);
  const code = process.exitCode;
  process.exitCode = before;
  return code;
}

test("update --optional-attendees keeps all 3 required guests and adds the new one (PATCH holds 4)", async () => {
  serveEvent(threeGuests());
  await run("update", "--id", "EVT1", "--optional-attendees", "new@x.example.test");
  assert.equal(patches().length, 1);
  const body = patches()[0].body;
  assert.deepEqual(addrs(body.attendees), ["a@x.example.test", "b@x.example.test", "c@x.example.test", "new@x.example.test"]);
  assert.deepEqual(body.attendees.map((a: any) => a.type), ["required", "required", "required", "optional"]);
  const result = JSON.parse(lines[0]);
  assert.deepEqual(result.data.attendee_changes, { added: ["new@x.example.test"], type_changed: [], removed: [], kept_count: 3 });
});

test("update prints a required -> optional type change and keeps the entry's response status", async () => {
  serveEvent(threeGuests());
  await run("update", "--id", "EVT1", "--optional-attendees", " B@X.example.test ,b@x.example.test");
  const body = patches()[0].body;
  assert.equal(body.attendees.length, 3);
  const b = body.attendees[1];
  assert.equal(b.type, "optional");
  assert.equal(b.status.response, "tentativelyAccepted");
  const summary = JSON.parse(lines[0]).data.attendee_changes;
  assert.deepEqual(summary.type_changed, [{ address: "b@x.example.test", from: "required", to: "optional" }]);
  assert.deepEqual(summary.added, []);
  assert.equal(summary.kept_count, 3);
});

test("a merge keeps rooms and response statuses exactly as they were", async () => {
  const room = guest("room@x.example.test", "resource", "accepted");
  serveEvent({ ...threeGuests(), attendees: [...threeGuests().attendees, room] });
  await run("update", "--id", "EVT1", "--attendees", "d@x.example.test");
  const body = patches()[0].body;
  assert.deepEqual(body.attendees.slice(0, 4), [...threeGuests().attendees, room]);
  assert.deepEqual(addrs(body.attendees), ["a@x.example.test", "b@x.example.test", "c@x.example.test", "room@x.example.test", "d@x.example.test"]);
});

test("the organizer is never added as an attendee", async () => {
  serveEvent(threeGuests());
  await run("update", "--id", "EVT1", "--attendees", "Owner@corp.example.test,d@x.example.test");
  assert.deepEqual(addrs(patches()[0].body.attendees), ["a@x.example.test", "b@x.example.test", "c@x.example.test", "d@x.example.test"]);
});

test("--replace-attendees without --yes: no PATCH, removal list printed, non-zero exit", async () => {
  serveEvent(threeGuests());
  const code = await runExpectingNonZero("update", "--id", "EVT1", "--replace-attendees", "--attendees", "a@x.example.test,n@x.example.test");
  assert.equal(code, 1);
  assert.equal(patches().length, 0);
  const result = JSON.parse(lines[0]);
  assert.equal(result.ok, false);
  assert.equal(result.code, "CONFIRMATION_REQUIRED");
  assert.deepEqual(result.data.removed, ["b@x.example.test", "c@x.example.test"]);
  assert.deepEqual(result.data.final_attendees, [{ address: "a@x.example.test", type: "required" }, { address: "n@x.example.test", type: "required" }]);
});

test("--replace-attendees --yes: PATCH has exactly the new list and the removed list is printed", async () => {
  serveEvent(threeGuests());
  await run("update", "--id", "EVT1", "--replace-attendees", "--yes", "--attendees", "a@x.example.test,n@x.example.test");
  assert.equal(patches().length, 1);
  assert.deepEqual(addrs(patches()[0].body.attendees), ["a@x.example.test", "n@x.example.test"]);
  const summary = JSON.parse(lines[0]).data.attendee_changes;
  assert.deepEqual(summary.removed, ["b@x.example.test", "c@x.example.test"]);
  assert.deepEqual(summary.added, ["n@x.example.test"]);
  assert.equal(summary.kept_count, 1);
});

test("--replace-attendees leaves rooms alone", async () => {
  const room = guest("room@x.example.test", "resource");
  serveEvent({ ...threeGuests(), attendees: [guest("a@x.example.test"), room] });
  await run("update", "--id", "EVT1", "--replace-attendees", "--yes", "--attendees", "n@x.example.test");
  assert.deepEqual(addrs(patches()[0].body.attendees), ["room@x.example.test", "n@x.example.test"]);
  assert.deepEqual(JSON.parse(lines[0]).data.attendee_changes.removed, ["a@x.example.test"]);
});

test("--remove-attendees without --yes: no PATCH, non-zero exit; with --yes removes just that guest", async () => {
  serveEvent(threeGuests());
  assert.equal(await runExpectingNonZero("update", "--id", "EVT1", "--remove-attendees", "B@x.example.test"), 1);
  assert.equal(patches().length, 0);
  assert.deepEqual(JSON.parse(lines[0]).data.removed, ["b@x.example.test"]);

  lines = [];
  await run("update", "--id", "EVT1", "--remove-attendees", "B@x.example.test", "--yes");
  assert.deepEqual(addrs(patches()[0].body.attendees), ["a@x.example.test", "c@x.example.test"]);
  assert.deepEqual(JSON.parse(lines[0]).data.attendee_changes.removed, ["b@x.example.test"]);
});

test("--remove-attendees of an unknown address: error, no PATCH", async () => {
  serveEvent(threeGuests());
  await runExpectingExit("update", "--id", "EVT1", "--remove-attendees", "nobody@x.example.test", "--yes");
  assert.equal(patches().length, 0);
  assert.equal(JSON.parse(lines[0]).code, "NOT_AN_ATTENDEE");
});

test("GET failure: error and no PATCH (fail closed)", async () => {
  serveEvent(new Error("Graph API 503"));
  await runExpectingExit("update", "--id", "EVT1", "--optional-attendees", "new@x.example.test");
  assert.equal(patches().length, 0);
  assert.equal(JSON.parse(lines[0]).code, "GET_FAILED");
});

test("event without an attendees field: refuse, no PATCH", async () => {
  serveEvent({ organizer: { emailAddress: { address: "owner@corp.example.test" } } });
  await runExpectingExit("update", "--id", "EVT1", "--optional-attendees", "new@x.example.test");
  assert.equal(patches().length, 0);
  assert.equal(JSON.parse(lines[0]).code, "NO_ATTENDEES_FIELD");
});

test("update without attendee flags does not read the event and patches only the given fields", async () => {
  await run("update", "--id", "EVT1", "--subject", "New title");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), ["PATCH /me/events/EVT1"]);
  assert.deepEqual(calls[0].body, { subject: "New title" });
});

test("create still sends the attendee list as given (unchanged)", async () => {
  await run("create", "--subject", "S", "--start", "2026-11-05T09:00:00", "--end", "2026-11-05T10:00:00",
    "--attendees", "a@x.example.test", "--optional-attendees", "o@x.example.test");
  assert.deepEqual(calls.map((c) => `${c.method} ${c.path}`), ["POST /me/events"]);
  assert.deepEqual(calls[0].body.attendees, [
    { emailAddress: { address: "a@x.example.test" }, type: "required" },
    { emailAddress: { address: "o@x.example.test" }, type: "optional" },
  ]);
});
