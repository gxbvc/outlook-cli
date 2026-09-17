import { getAccessToken } from "./auth.js";
import { GRAPH_BASE } from "./config.js";

function mailboxPath(path: string): string {
  const user = process.env.OUTLOOK_USER;
  if (!user || !path.startsWith("/me")) return path;
  return `/users/${encodeURIComponent(user)}` + path.slice(3);
}

export async function graphRequest(path: string, options: RequestInit = {}): Promise<any> {
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
    throw new Error(`Graph API ${res.status}: ${err}`);
  }

  return res.json();
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
