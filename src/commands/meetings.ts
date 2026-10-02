import { Command } from "commander";
import { writeFileSync } from "fs";
import { getAccessToken } from "../auth.js";
import { GRAPH_BASE } from "../config.js";
import { graphRequest, out, err } from "../graph.js";

// Teams meetings and their transcripts. Graph only finds a meeting by its join
// URL, so every command takes --join-url or --event-id (a calendar event whose
// onlineMeeting.joinUrl we read). Transcripts need OnlineMeetings.Read and
// OnlineMeetingTranscript.Read.All; run `outlook-cli auth` once after adding them.

async function joinUrlFor(opts: { joinUrl?: string; eventId?: string }): Promise<string> {
  if (opts.joinUrl) return opts.joinUrl;
  if (!opts.eventId) err("Give --join-url or --event-id", "MISSING_MEETING");
  const event = await graphRequest(`/me/events/${encodeURIComponent(opts.eventId)}?$select=subject,onlineMeeting`);
  const url = event?.onlineMeeting?.joinUrl;
  if (!url) err(`Event "${event?.subject}" has no Teams join URL`, "NOT_ONLINE_MEETING");
  return url;
}

async function findMeeting(joinUrl: string): Promise<any> {
  const filter = encodeURIComponent(`JoinWebUrl eq '${joinUrl.replace(/'/g, "''")}'`);
  const res = await graphRequest(`/me/onlineMeetings?$filter=${filter}`);
  const meeting = res?.value?.[0];
  if (!meeting) err("No Teams meeting found for that join URL", "MEETING_NOT_FOUND");
  return meeting;
}

async function listTranscripts(meetingId: string): Promise<any[]> {
  const res = await graphRequest(`/me/onlineMeetings/${encodeURIComponent(meetingId)}/transcripts`);
  return res?.value || [];
}

// The content endpoint returns WebVTT, not JSON, so graphRequest cannot read it.
async function fetchVtt(meetingId: string, transcriptId: string): Promise<string> {
  const token = await getAccessToken();
  const url = `${GRAPH_BASE}/me/onlineMeetings/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}/content?$format=text/vtt`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${await res.text()}`);
  return res.text();
}

function decodeEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");
}

// Teams cues look like `<v Christian Genco>Hello there.</v>`. Join a speaker's
// back-to-back cues into one "Name: text" line, like the ~/scribe .txt files.
export function vttToText(vtt: string): string {
  const lines: string[] = [];
  let speaker = "";
  for (const raw of vtt.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line === "WEBVTT" || line.includes("-->") || /^[0-9a-f-]+\/\d+-\d+$/i.test(line)) continue;
    const m = line.match(/^<v ([^>]+)>(.*?)(<\/v>)?$/);
    const who = m ? m[1].trim() : speaker || "Speaker";
    const text = decodeEntities((m ? m[2] : line).replace(/<[^>]+>/g, "")).trim();
    if (!text) continue;
    if (who === speaker && lines.length) lines[lines.length - 1] += ` ${text}`;
    else lines.push(`${who}: ${text}`);
    speaker = who;
  }
  return lines.join("\n") + "\n";
}

export function registerMeetings(program: Command) {
  const meetings = program.command("meetings").description("Teams meetings and transcripts");

  meetings
    .command("get")
    .description("Show the Teams meeting behind a join URL or calendar event")
    .option("--join-url <url>", "Teams join URL")
    .option("--event-id <id>", "Calendar event ID")
    .action(async (opts) => {
      try {
        out(await findMeeting(await joinUrlFor(opts)));
      } catch (e: any) {
        err(e.message);
      }
    });

  meetings
    .command("transcripts")
    .description("List a meeting's transcripts (one per time transcription was started)")
    .option("--join-url <url>", "Teams join URL")
    .option("--event-id <id>", "Calendar event ID")
    .action(async (opts) => {
      try {
        const meeting = await findMeeting(await joinUrlFor(opts));
        const transcripts = await listTranscripts(meeting.id);
        out({ meetingId: meeting.id, subject: meeting.subject, transcripts });
      } catch (e: any) {
        err(e.message);
      }
    });

  meetings
    .command("transcript")
    .description("Download a transcript as speaker-labeled text (default) or raw WebVTT")
    .option("--join-url <url>", "Teams join URL")
    .option("--event-id <id>", "Calendar event ID")
    .option("--transcript-id <id>", "Which transcript; default joins all of them in order")
    .option("--since <iso>", "Only transcripts started at or after this time")
    .option("--until <iso>", "Only transcripts started before this time")
    .option("--format <format>", "text or vtt", "text")
    .option("--out <path>", "Write to this file instead of stdout JSON")
    .action(async (opts) => {
      try {
        const meeting = await findMeeting(await joinUrlFor(opts));
        let transcripts = await listTranscripts(meeting.id);
        if (opts.transcriptId) transcripts = transcripts.filter((t) => t.id === opts.transcriptId);
        // A recurring meeting has one join URL, so its transcripts pile up across
        // occurrences. --since/--until pick out one occurrence.
        const at = (t: any) => Date.parse(t.createdDateTime);
        if (opts.since) transcripts = transcripts.filter((t) => at(t) >= Date.parse(opts.since));
        if (opts.until) transcripts = transcripts.filter((t) => at(t) < Date.parse(opts.until));
        if (!transcripts.length) err("This meeting has no transcript. Transcription was not turned on.", "NO_TRANSCRIPT");
        transcripts.sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));

        const parts: string[] = [];
        for (const t of transcripts) {
          const vtt = await fetchVtt(meeting.id, t.id);
          parts.push(opts.format === "vtt" ? vtt : vttToText(vtt));
        }
        const content = parts.join(opts.format === "vtt" ? "\n" : "");
        const meta = {
          meetingId: meeting.id,
          subject: meeting.subject,
          startDateTime: meeting.startDateTime,
          endDateTime: meeting.endDateTime,
          transcriptIds: transcripts.map((t) => t.id),
          createdDateTime: transcripts[0].createdDateTime,
          endedDateTime: transcripts[transcripts.length - 1].endDateTime,
        };
        if (opts.out) {
          writeFileSync(opts.out, content);
          out({ ...meta, path: opts.out, bytes: Buffer.byteLength(content) });
        } else {
          out({ ...meta, content });
        }
      } catch (e: any) {
        err(e.message);
      }
    });
}
