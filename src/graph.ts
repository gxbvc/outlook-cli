import { getAccessToken } from "./auth.js";
import { GRAPH_BASE } from "./config.js";

export async function graphRequest(path: string, options: RequestInit = {}): Promise<any> {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;

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

export function out(data: any): void {
  console.log(JSON.stringify({ ok: true, data }));
}

export function err(message: string, code?: string): never {
  console.log(JSON.stringify({ ok: false, error: message, code }));
  process.exit(1);
}
