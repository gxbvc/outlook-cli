import { getAccessToken } from "./auth.js";
import { GRAPH_BASE } from "./config.js";

function mailboxPath(path: string): string {
  const user = process.env.OUTLOOK_USER;
  if (!user || !path.startsWith("/me")) return path;
  return `/users/${encodeURIComponent(user)}` + path.slice(3);
}

// Graph error with the status code and Retry-After (when throttled) attached,
// so commands can turn 404/403/429 into friendly messages.
export class GraphError extends Error {
  status: number;
  retryAfter: number | null;
  constructor(status: number, message: string, retryAfter: number | null = null) {
    super(message);
    this.status = status;
    this.retryAfter = retryAfter;
  }
}

// Tests replace the HTTP layer with this. Receives the resolved path (after --user mapping).
export type Transport = (path: string, options: RequestInit) => Promise<any>;
let testTransport: Transport | null = null;
export function setTransportForTests(t: Transport | null): void { testTransport = t; }

export async function graphRequest(path: string, options: RequestInit = {}): Promise<any> {
  if (testTransport) return testTransport(mailboxPath(path), options);
  const token = await getAccessToken();
  const resolved = mailboxPath(path);
  const url = resolved.startsWith("http") ? resolved : `${GRAPH_BASE}${resolved}`;


  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });

  if (res.status === 204 || res.status === 202) return null;

  if (!res.ok) {
    const err = await res.text();
    const retryAfter = res.headers.get("Retry-After");
    throw new GraphError(res.status, `Graph API ${res.status}: ${err}`, retryAfter ? parseInt(retryAfter, 10) : null);
  }

  return res.json();
}

// Like graphRequest but returns the raw Response, for binary/streaming endpoints
// (file downloads). Tests may return a Response object from the stub transport.
export async function graphRawRequest(path: string, options: RequestInit = {}): Promise<Response> {
  if (testTransport) return testTransport(mailboxPath(path), options);
  const token = await getAccessToken();
  const resolved = mailboxPath(path);
  const url = resolved.startsWith("http") ? resolved : `${GRAPH_BASE}${resolved}`;
  return fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...options.headers,
    },
  });
}

// The `webLink` Graph hands back is a legacy /owa/?ItemID= link that answers 401
// in a browser, and OWA's own URLs use a different id format than the REST id.
// translateExchangeIds converts one to the other; OWA then wants standard base64,
// not the base64url Graph returns. Result: a link that opens the message in OWA.
export async function owaUrl(messageId: string, folder = "drafts"): Promise<string | null> {
  try {
    const res = await graphRequest("/me/translateExchangeIds", {
      method: "POST",
      body: JSON.stringify({
        inputIds: [messageId],
        sourceIdType: "restId",
        targetIdType: "restImmutableEntryId",
      }),
    });
    const target = res?.value?.[0]?.targetId;
    if (!target) return null;
    const std = target.replace(/-/g, "+").replace(/_/g, "/");
    return `https://outlook.office.com/mail/${folder}/id/${encodeURIComponent(std)}`;
  } catch {
    return null; // a link lookup must never fail the command that created the draft
  }
}

export function out(data: any): void {
  console.log(JSON.stringify({ ok: true, data }));
}

export function err(message: string, code?: string): never {
  console.log(JSON.stringify({ ok: false, error: message, code }));
  process.exit(1);
}
