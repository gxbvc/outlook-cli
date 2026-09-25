import { Command } from "commander";
import { graphRequest, out, err } from "../graph.js";
import { DEFAULT_TZ } from "../config.js";

function parseAttendees(csv?: string, type: "required" | "optional" = "required") {
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

function combineAttendees(required?: string, optional?: string) {
  const r = parseAttendees(required, "required") || [];
  const o = parseAttendees(optional, "optional") || [];
  if (r.length === 0 && o.length === 0) return undefined;
  return [...r, ...o];
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
        const attendees = combineAttendees(opts.attendees, opts.optionalAttendees);
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

        const data = await graphRequest("/me/events", {
          method: "POST",
          headers: tzHeaders(tz),
          body: JSON.stringify(event),
        });
        out(data);
      } catch (e: any) { err(e.message, "CREATE_FAILED"); }
    });

  calendar
    .command("update")
    .description("Update a calendar event")
    .requiredOption("--id <id>", "Event ID")
    .option("--subject <subject>", "New subject")
    .option("--start <iso>", "New start datetime")
    .option("--end <iso>", "New end datetime")
    .option("--tz <tz>", "IANA timezone for --start/--end (defaults to OUTLOOK_DEFAULT_TZ)")
    .option("--location <location>", "New location")
    .option("--body <text>", "New plain text body")
    .option("--html <html>", "New HTML body")
    .option("--attendees <addresses>", "Replace required attendees (comma-separated)")
    .option("--optional-attendees <addresses>", "Replace optional attendees (comma-separated)")
    .option("--online", "Convert to an online (Teams) meeting")
    .action(async (opts) => {
      try {
        const tz = opts.tz || DEFAULT_TZ;
        const patch: any = {};
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
        if (opts.attendees !== undefined || opts.optionalAttendees !== undefined) {
          patch.attendees = combineAttendees(opts.attendees, opts.optionalAttendees) || [];
        }
        if (opts.online) {
          patch.isOnlineMeeting = true;
          patch.onlineMeetingProvider = "teamsForBusiness";
        }

        if (Object.keys(patch).length === 0) {
          err("Nothing to update. Provide at least one field.", "NO_CHANGES");
        }

        const data = await graphRequest(`/me/events/${opts.id}`, {
          method: "PATCH",
          headers: tzHeaders(tz),
          body: JSON.stringify(patch),
        });
        out(data);
      } catch (e: any) { err(e.message, "UPDATE_FAILED"); }
    });

  calendar
    .command("delete")
    .description("Delete a calendar event (silent — does not notify attendees)")
    .requiredOption("--id <id>", "Event ID")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/events/${opts.id}`, { method: "DELETE" });
        out({ message: "Event deleted" });
      } catch (e: any) { err(e.message, "DELETE_FAILED"); }
    });

  calendar
    .command("cancel")
    .description("Cancel an event you organize and notify attendees (use an occurrence ID to cancel one instance of a series)")
    .requiredOption("--id <id>", "Event ID")
    .option("--comment <text>", "Message included in the cancellation notice")
    .action(async (opts) => {
      try {
        await graphRequest(`/me/events/${opts.id}/cancel`, {
          method: "POST",
          body: JSON.stringify({ comment: opts.comment || "" }),
        });
        out({ message: "Event cancelled" });
      } catch (e: any) { err(e.message, "CANCEL_FAILED"); }
    });
}
