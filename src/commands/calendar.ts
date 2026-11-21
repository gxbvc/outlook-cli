import { Command } from "commander";
import { graphRequest, out, err } from "../graph.js";
import { DEFAULT_TZ } from "../config.js";
import { withFooterText } from "../footer.js";
import { loggedSend, failCode } from "../sendlog.js";

const MAYBE_MAILS = "Exchange mails attendees only when the event has attendees and this mailbox organizes it";

function collectList(value: string, previous: string[] = []) {
  return [...previous, ...value.split(",").map((a) => a.trim()).filter(Boolean)];
}

// Body of POST /me/events/{id}/forward. Graph sends the invite to each address; the organizer is told. The
// comment always ends with the agent footer (plain text), like emails forward.
export function forwardEventRequest(addresses: string[], comment?: string) {
  const seen = new Set<string>();
  const toRecipients = addresses
    .filter((a) => { const k = a.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; })
    .map((address) => ({ emailAddress: { address } }));
  return { toRecipients, comment: withFooterText(comment) };
}

function parseAttendees(csv?: string, type: "required" | "optional" | "resource" = "required") {
  if (!csv) return undefined;
  return csv
    .split(",")
    .map((a) => a.trim())
    .filter(Boolean)
    .map((address) => ({
      emailAddress: { address },
      type,
    }));
}

function combineAttendees(required?: string, optional?: string, resources?: string) {
  const r = parseAttendees(required, "required") || [];
  const o = parseAttendees(optional, "optional") || [];
  const rooms = parseAttendees(resources, "resource") || [];
  if (r.length === 0 && o.length === 0 && rooms.length === 0) return undefined;
  return [...r, ...o, ...rooms];
}

// ---- Safe attendee updates ----
// Graph PATCH of `attendees` REPLACES the whole list, and dropped guests get cancellation emails. So the update
// command reads the current list first and builds the new one here. Default is MERGE: nobody is removed unless
// --replace-attendees or --remove-attendees is given AND confirmed with --yes.

export class AttendeeError extends Error {
  code: string;
  constructor(message: string, code: string) { super(message); this.code = code; }
}

const norm = (a: string) => a.trim().toLowerCase();

export function parseAddressList(csv: string | undefined, flag: string): string[] {
  if (csv === undefined) return [];
  const seen = new Set<string>();
  const list: string[] = [];
  for (const raw of String(csv).split(",")) {
    const a = raw.trim();
    if (!a) continue;
    if (!a.includes("@")) throw new AttendeeError(`${flag}: "${a}" is not an email address`, "INVALID_ADDRESS");
    if (seen.has(norm(a))) continue;
    seen.add(norm(a));
    list.push(a);
  }
  if (list.length === 0) throw new AttendeeError(`${flag} has no addresses`, "EMPTY_ADDRESS_LIST");
  return list;
}

export interface AttendeePlanInput {
  current: any;            // the event as returned by GET (needs `attendees`; `organizer` is used)
  required?: string[];     // --attendees
  optional?: string[];     // --optional-attendees
  replace?: boolean;       // --replace-attendees
  remove?: string[];       // --remove-attendees
}

export interface AttendeePlan {
  attendees: any[];        // full list to PATCH
  added: string[];
  type_changed: { address: string; from: string; to: string }[];
  removed: string[];
  kept_count: number;
}

export function planAttendees({ current, required = [], optional = [], replace = false, remove = [] }: AttendeePlanInput): AttendeePlan {
  if (!current || !Array.isArray(current.attendees)) {
    throw new AttendeeError("Event has no attendees field; refusing to change attendees", "NO_ATTENDEES_FIELD");
  }
  if (replace && remove.length) {
    throw new AttendeeError("Use --replace-attendees or --remove-attendees, not both", "CONFLICTING_FLAGS");
  }
  if (replace && required.length + optional.length === 0) {
    throw new AttendeeError("--replace-attendees needs --attendees and/or --optional-attendees", "EMPTY_REPLACE");
  }
  const organizer = norm(current.organizer?.emailAddress?.address || "");
  const isOrganizer = (a: string) => organizer !== "" && norm(a) === organizer;

  const wanted = new Map<string, { address: string; type: "required" | "optional" }>();
  for (const [list, type] of [[required, "required"], [optional, "optional"]] as const) {
    for (const address of list) {
      if (isOrganizer(address)) continue; // never add the organizer as an attendee
      const prev = wanted.get(norm(address));
      if (prev && prev.type !== type) {
        throw new AttendeeError(`${address} is in both --attendees and --optional-attendees`, "CONFLICTING_TYPES");
      }
      wanted.set(norm(address), { address, type });
    }
  }

  const existing: any[] = current.attendees;
  const existingByAddr = new Map<string, any>();
  for (const a of existing) existingByAddr.set(norm(a?.emailAddress?.address || ""), a);

  const removeKeys = new Set<string>();
  for (const address of remove) {
    if (isOrganizer(address)) throw new AttendeeError(`${address} is the organizer and cannot be removed`, "ORGANIZER");
    if (!existingByAddr.has(norm(address))) {
      throw new AttendeeError(`${address} is not an attendee of this event; nothing changed`, "NOT_AN_ATTENDEE");
    }
    if (wanted.has(norm(address))) {
      throw new AttendeeError(`${address} is both added and removed`, "CONFLICTING_FLAGS");
    }
    removeKeys.add(norm(address));
  }

  if (replace) {
    // Everyone not named is removed, except the organizer and resources (rooms). Rooms go only through --remove-attendees.
    for (const a of existing) {
      const k = norm(a?.emailAddress?.address || "");
      if (!wanted.has(k) && !isOrganizer(k) && a.type !== "resource") removeKeys.add(k);
    }
  }

  const added: string[] = [];
  const type_changed: AttendeePlan["type_changed"] = [];
  const attendees: any[] = [];
  for (const a of existing) {
    const k = norm(a?.emailAddress?.address || "");
    if (removeKeys.has(k)) continue;
    const want = wanted.get(k);
    if (want && a.type !== want.type) {
      type_changed.push({ address: a.emailAddress.address, from: a.type, to: want.type });
      attendees.push({ ...a, type: want.type });
    } else {
      attendees.push(a); // untouched: keeps response status and everything else
    }
  }
  for (const [k, want] of wanted) {
    if (existingByAddr.has(k)) continue;
    added.push(want.address);
    attendees.push({ emailAddress: { address: want.address }, type: want.type });
  }
  const removed = existing
    .filter((a) => removeKeys.has(norm(a?.emailAddress?.address || "")))
    .map((a) => a.emailAddress.address as string);
  const kept_count = attendees.length - added.length;
  return { attendees, added, type_changed, removed, kept_count };
}

function summaryOf(plan: AttendeePlan) {
  return { added: plan.added, type_changed: plan.type_changed, removed: plan.removed, kept_count: plan.kept_count };
}

function finalList(plan: AttendeePlan) {
  return plan.attendees.map((a) => ({ address: a.emailAddress.address, type: a.type }));
}

// Body of POST /me/calendar/getSchedule: free/busy of mailboxes (people or rooms) in a window.
export function scheduleRequest(addresses: string[], start: string, end: string, tz: string, interval: number) {
  return {
    schedules: addresses,
    startTime: { dateTime: start, timeZone: tz },
    endTime: { dateTime: end, timeZone: tz },
    availabilityViewInterval: interval,
  };
}

function tzHeaders(tz: string) {
  return { Prefer: `outlook.timezone="${tz}"` };
}

function startOfDayISO(tz: string): string {
  const now = new Date();
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone: tz,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = fmt.formatToParts(now);
  const y = parts.find((p) => p.type === "year")!.value;
  const m = parts.find((p) => p.type === "month")!.value;
  const d = parts.find((p) => p.type === "day")!.value;
  return `${y}-${m}-${d}T00:00:00`;
}

function endOfDayISO(tz: string): string {
  const now = new Date();
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone: tz,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = fmt.formatToParts(now);
  const y = parts.find((p) => p.type === "year")!.value;
  const m = parts.find((p) => p.type === "month")!.value;
  const d = parts.find((p) => p.type === "day")!.value;
  return `${y}-${m}-${d}T23:59:59`;
}

export function registerCalendar(program: Command) {
  const calendar = program.command("calendar").description("Manage calendar events");

  calendar
    .command("list")
    .description("List calendar events (expands recurrences when --start/--end or --today is used)")
    .option("--today", "Show events for today (in the default timezone)")
    .option("--start <iso>", "Start datetime (ISO, e.g. 2026-04-24T00:00:00)")
    .option("--end <iso>", "End datetime (ISO)")
    .option("--limit <n>", "Max results", "50")
    .option("--search <query>", "Search query (KQL)")
    .option("--tz <tz>", "IANA timezone for returned start/end (defaults to OUTLOOK_DEFAULT_TZ)")
    .option("--page-token <token>", "Pagination token (nextLink)")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        let url: string;

        if (opts.pageToken) {
          url = opts.pageToken;
        } else if (opts.today || opts.start || opts.end) {
          const start = opts.start || (opts.today ? startOfDayISO(tz) : undefined);
          const end = opts.end || (opts.today ? endOfDayISO(tz) : undefined);
          if (!start || !end) {
            err("--start and --end are both required (or use --today)", "MISSING_RANGE");
          }
          const params = new URLSearchParams({
            startDateTime: start!,
            endDateTime: end!,
            $top: opts.limit,
          });
          params.append("$select", "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,onlineMeeting,webLink,bodyPreview");
          params.append("$orderby", "start/dateTime");
          url = `/me/calendarView?${params.toString()}`;
        } else {
          const params = new URLSearchParams({ $top: opts.limit });
          params.append("$select", "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,onlineMeeting,webLink,bodyPreview");
          if (!opts.search) params.append("$orderby", "start/dateTime desc");
          if (opts.search) params.append("$search", `"${opts.search}"`);
          url = `/me/events?${params.toString()}`;
        }

        const data = await graphRequest(url, { headers: tzHeaders(tz) });
        out({
          events: data.value,
          nextPage: data["@odata.nextLink"] || null,
          count: data.value.length,
          timezone: tz,
        });
      } catch (e: any) { err(e.message, "LIST_FAILED"); }
    });

  calendar
    .command("schedule")
    .description("Free/busy for mailboxes or rooms in a window (Graph getSchedule; no event details needed)")
    .requiredOption("--addresses <addresses>", "Comma-separated mailbox or room addresses")
    .requiredOption("--start <iso>", "Start datetime (ISO, e.g. 2026-11-05T09:00:00)")
    .requiredOption("--end <iso>", "End datetime (ISO)")
    .option("--tz <tz>", "IANA timezone for --start/--end (defaults to OUTLOOK_DEFAULT_TZ)")
    .option("--interval <minutes>", "availabilityView slot length in minutes", "30")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        const addresses = String(opts.addresses).split(",").map((a) => a.trim()).filter(Boolean);
        if (addresses.length === 0) err("--addresses is empty", "MISSING_ADDRESSES");
        const data = await graphRequest("/me/calendar/getSchedule", {
          method: "POST",
          headers: tzHeaders(tz),
          body: JSON.stringify(scheduleRequest(addresses, opts.start, opts.end, tz, Number(opts.interval) || 30)),
        });
        out({ schedules: data?.value || [], timezone: tz });
      } catch (e: any) { err(e.message, "SCHEDULE_FAILED"); }
    });

  calendar
    .command("get")
    .description("Get a single event by ID")
    .requiredOption("--id <id>", "Event ID")
    .option("--tz <tz>", "IANA timezone for returned start/end")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        const data = await graphRequest(`/me/events/${opts.id}`, { headers: tzHeaders(tz) });
        out(data);
      } catch (e: any) { err(e.message, "GET_FAILED"); }
    });

  calendar
    .command("create")
    .description("Create a calendar event")
    .requiredOption("--subject <subject>", "Event subject/title")
    .requiredOption("--start <iso>", "Start datetime (ISO, e.g. 2026-04-24T09:00:00)")
    .requiredOption("--end <iso>", "End datetime (ISO)")
    .option("--tz <tz>", "IANA timezone for --start/--end (defaults to OUTLOOK_DEFAULT_TZ)")
    .option("--location <location>", "Location display name")
    .option("--body <text>", "Plain text body")
    .option("--html <html>", "HTML body (overrides --body)")
    .option("--attendees <addresses>", "Comma-separated required attendee email addresses")
    .option("--optional-attendees <addresses>", "Comma-separated optional attendee email addresses")
    .option("--resources <addresses>", "Comma-separated room/resource mailboxes (attendee type resource)")
    .option("--online", "Create as an online (Teams) meeting")
    .option("--all-day", "All-day event")
    .option("--recurrence <type>", "Recurrence type: daily, weekly, biweekly, monthly, yearly")
    .option("--recurrence-days <days>", "Comma-separated days for weekly recurrence (e.g. monday,tuesday)")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        const event: any = {
          subject: opts.subject,
          start: { dateTime: opts.start, timeZone: tz },
          end: { dateTime: opts.end, timeZone: tz },
        };
        if (opts.html || opts.body) {
          event.body = {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.html || opts.body,
          };
        }
        if (opts.location) event.location = { displayName: opts.location };
        const attendees = combineAttendees(opts.attendees, opts.optionalAttendees, opts.resources);
        if (attendees) event.attendees = attendees;
        if (opts.online) {
          event.isOnlineMeeting = true;
          event.onlineMeetingProvider = "teamsForBusiness";
        }
        if (opts.allDay) event.isAllDay = true;
        if (opts.recurrence) {
          const typeMap: Record<string, { type: string; interval: number }> = {
            daily:    { type: "daily",   interval: 1 },
            weekly:   { type: "weekly",  interval: 1 },
            biweekly: { type: "weekly",  interval: 2 },
            monthly:  { type: "absoluteMonthly", interval: 1 },
            yearly:   { type: "absoluteYearly",  interval: 1 },
          };
          const r = typeMap[opts.recurrence.toLowerCase()];
          if (!r) err(`Unknown recurrence type: ${opts.recurrence}`, "INVALID_RECURRENCE");
          const startDate = opts.start.split("T")[0];
          const daysOfWeek = opts.recurrenceDays
            ? opts.recurrenceDays.split(",").map((d: string) => d.trim().toLowerCase())
            : undefined;
          event.recurrence = {
            pattern: {
              type: r.type,
              interval: r.interval,
              ...(daysOfWeek ? { daysOfWeek } : {}),
            },
            range: {
              type: "noEnd",
              startDate,
            },
          };
        }

        const create = () => graphRequest("/me/events", {
          method: "POST",
          headers: tzHeaders(tz),
          body: JSON.stringify(event),
        });
        // An event with attendees (rooms too) sends invitations; one without sends nothing.
        const data = attendees
          ? await loggedSend({
              kind: "event_create",
              to: attendees.filter((a) => a.type !== "optional"),
              cc: attendees.filter((a) => a.type === "optional"),
              bcc: [],
              subject: opts.subject,
              note: "to = required attendees and rooms, cc = optional attendees",
            }, create)
          : await create();
        out(data);
      } catch (e: any) { err(e.message, failCode(e, "CREATE_FAILED")); }
    });

  calendar
    .command("update")
    .description("Update a calendar event. Attendee flags MERGE into the current guest list (nobody is removed). " +
      "Removing guests needs --replace-attendees or --remove-attendees plus --yes; without --yes it prints the " +
      "removal list and exits non-zero without changing anything. The organizer and rooms are never touched by --replace-attendees.")
    .requiredOption("--id <id>", "Event ID")
    .option("--subject <subject>", "New subject")
    .option("--start <iso>", "New start datetime")
    .option("--end <iso>", "New end datetime")
    .option("--tz <tz>", "IANA timezone for --start/--end (defaults to OUTLOOK_DEFAULT_TZ)")
    .option("--location <location>", "New location")
    .option("--body <text>", "New plain text body")
    .option("--html <html>", "New HTML body")
    .option("--attendees <addresses>", "Add these required attendees (comma-separated); existing guests stay. An existing optional guest named here becomes required")
    .option("--optional-attendees <addresses>", "Add these optional attendees (comma-separated); existing guests stay. An existing required guest named here becomes optional")
    .option("--replace-attendees", "Make --attendees/--optional-attendees the whole guest list (removes everyone else except the organizer and rooms). Needs --yes")
    .option("--remove-attendees <addresses>", "Remove these guests (comma-separated; each must be on the event). Needs --yes")
    .option("--yes", "Confirm removals from --replace-attendees or --remove-attendees")
    .option("--online", "Convert to an online (Teams) meeting")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        const patch: any = {};
        let plan: AttendeePlan | null = null;
        const touchesAttendees = opts.attendees !== undefined || opts.optionalAttendees !== undefined ||
          opts.removeAttendees !== undefined || opts.replaceAttendees;
        if (touchesAttendees) {
          const required = parseAddressList(opts.attendees, "--attendees");
          const optional = parseAddressList(opts.optionalAttendees, "--optional-attendees");
          const remove = parseAddressList(opts.removeAttendees, "--remove-attendees");
          let current: any;
          try {
            current = await graphRequest(`/me/events/${opts.id}?$select=organizer,attendees`, { headers: tzHeaders(tz) });
          } catch (e: any) {
            throw new AttendeeError(`Could not read the event's current attendees (${e.message}); nothing changed`, "GET_FAILED");
          }
          plan = planAttendees({ current, required, optional, replace: !!opts.replaceAttendees, remove });
          if (plan.removed.length > 0 && !opts.yes) {
            console.log(JSON.stringify({
              ok: false,
              error: "Removing attendees needs --yes. Nothing was changed.",
              code: "CONFIRMATION_REQUIRED",
              data: { ...summaryOf(plan), final_attendees: finalList(plan) },
            }));
            process.exitCode = 1;
            return;
          }
          patch.attendees = plan.attendees;
        }
        if (opts.subject) patch.subject = opts.subject;
        if (opts.start) patch.start = { dateTime: opts.start, timeZone: tz };
        if (opts.end) patch.end = { dateTime: opts.end, timeZone: tz };
        if (opts.location) patch.location = { displayName: opts.location };
        if (opts.html || opts.body) {
          patch.body = {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.html || opts.body,
          };
        }
        if (opts.online) {
          patch.isOnlineMeeting = true;
          patch.onlineMeetingProvider = "teamsForBusiness";
        }

        if (Object.keys(patch).length === 0) {
          err("Nothing to update. Provide at least one field.", "NO_CHANGES");
        }

        const final = plan ? plan.attendees : null;
        const data = await loggedSend({
          kind: "event_update",
          to: final ? final.filter((a) => a.type !== "optional") : null,
          cc: final ? final.filter((a) => a.type === "optional") : null,
          bcc: final ? [] : null,
          subject: opts.subject ?? null,
          subjectNote: "subject not changed by this update; the event's subject was not read",
          event_id: opts.id,
          note: final ? `${MAYBE_MAILS}; to/cc = attendee list after the update` : `${MAYBE_MAILS}; attendees not read`,
        }, () => graphRequest(`/me/events/${opts.id}`, {
          method: "PATCH",
          headers: tzHeaders(tz),
          body: JSON.stringify(patch),
        }));
        out(plan ? { ...(data || {}), attendee_changes: summaryOf(plan) } : data);
      } catch (e: any) {
        if (e instanceof AttendeeError) err(e.message, e.code);
        err(e.message, failCode(e, "UPDATE_FAILED"));
      }
    });

  calendar
    .command("delete")
    .description("Delete a calendar event (if you organize a meeting with attendees, Graph sends them a cancellation; logged in the send log)")
    .requiredOption("--id <id>", "Event ID")
    .action(async (opts) => {
      try {
        // Graph: deleting a meeting from the organizer's calendar sends a cancellation to attendees.
        await loggedSend({
          kind: "event_delete", to: null, cc: null, bcc: null, subject: null,
          subjectNote: "event not read before delete", event_id: opts.id, note: MAYBE_MAILS,
        }, () => graphRequest(`/me/events/${opts.id}`, { method: "DELETE" }));
        out({ message: "Event deleted" });
      } catch (e: any) { err(e.message, failCode(e, "DELETE_FAILED")); }
    });

  calendar
    .command("cancel")
    .description("Cancel an event you organize and notify attendees (use an occurrence ID to cancel one instance of a series)")
    .requiredOption("--id <id>", "Event ID")
    .option("--comment <text>", "Message included in the cancellation notice")
    .action(async (opts) => {
      try {
        await loggedSend({
          kind: "event_cancel", to: null, cc: null, bcc: null, subject: null,
          subjectNote: "Graph sets the cancellation subject; event not read", event_id: opts.id,
          note: "cancellation goes to every attendee of the event",
        }, () => graphRequest(`/me/events/${opts.id}/cancel`, {
          method: "POST",
          body: JSON.stringify({ comment: opts.comment || "" }),
        }));
        out({ message: "Event cancelled" });
      } catch (e: any) { err(e.message, failCode(e, "CANCEL_FAILED")); }
    });

  calendar
    .command("forward")
    .description("Forward a meeting invite to other people (Graph /forward; an occurrence ID forwards one instance)")
    .requiredOption("--id <id>", "Event ID")
    .requiredOption("--to <addresses>", "Comma-separated and/or repeated addresses", collectList, [])
    .option("--comment <text>", "Note sent with the invite (the agent footer is added)")
    .action(async (opts) => {
      try {
        if (!opts.to.length) return err("--to needs at least one address", "NO_RECIPIENTS");
        const body = forwardEventRequest(opts.to, opts.comment);
        await loggedSend({
          kind: "event_forward", to: body.toRecipients, cc: [], bcc: [], subject: null,
          subjectNote: "Graph uses the event's subject; event not read", event_id: opts.id,
        }, () => graphRequest(`/me/events/${opts.id}/forward`, { method: "POST", body: JSON.stringify(body) }));
        out({ message: "Event forwarded", id: opts.id, to: body.toRecipients.map((r) => r.emailAddress.address) });
      } catch (e: any) { err(e.message, failCode(e, "FORWARD_FAILED")); }
    });
}
