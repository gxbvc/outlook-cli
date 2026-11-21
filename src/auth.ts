import { createServer } from "http";
import { URL } from "url";
import { mkdirSync, writeFileSync, readFileSync, existsSync } from "fs";
import { CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, AUTHORITY, SCOPES, TOKENS_DIR, TOKENS_FILE } from "./config.js";

interface TokenData {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  expires_at: number;
  token_type: string;
  scope: string;
}

function buildAuthUrl(): string {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: SCOPES.join(" "),
    prompt: "consent",
  });
  return `${AUTHORITY}/oauth2/v2.0/authorize?${params.toString()}`;
}

async function exchangeCode(code: string): Promise<TokenData> {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    code,
    redirect_uri: REDIRECT_URI,
    grant_type: "authorization_code",
    scope: SCOPES.join(" "),
  });

  const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Token exchange failed: ${res.status} ${err}`);
  }

  const data = await res.json() as any;
  return {
    ...data,
    expires_at: Date.now() + data.expires_in * 1000,
  };
}

// Refresh with the scopes this token was granted, not SCOPES. When SCOPES gains
// a new permission, asking for it on refresh fails until `outlook-cli auth` runs
// again, and that would break email too.
export async function refreshTokens(refreshToken: string, grantedScope?: string): Promise<TokenData> {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: grantedScope ? `${grantedScope} offline_access` : SCOPES.join(" "),
  });

  const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Token refresh failed: ${res.status} ${err}`);
  }

  const data = await res.json() as any;
  return {
    ...data,
    expires_at: Date.now() + data.expires_in * 1000,
  };
}

function saveTokens(tokens: TokenData): void {
  mkdirSync(TOKENS_DIR, { recursive: true });
  writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
}

export function loadTokens(): TokenData | null {
  if (!existsSync(TOKENS_FILE)) return null;
  try {
    return JSON.parse(readFileSync(TOKENS_FILE, "utf-8"));
  } catch {
    return null;
  }
}

export async function getAccessToken(): Promise<string> {
  let tokens = loadTokens();
  if (!tokens) {
    console.error(JSON.stringify({ ok: false, error: "Not authenticated. Run: outlook-cli auth", code: "NOT_AUTHENTICATED" }));
    process.exit(1);
  }

  // Refresh if expired or about to expire (5 min buffer)
  if (Date.now() >= tokens.expires_at - 300_000) {
    if (!tokens.refresh_token) {
      console.error(JSON.stringify({ ok: false, error: "No refresh token. Run: outlook-cli auth", code: "NO_REFRESH_TOKEN" }));
      process.exit(1);
    }
    tokens = await refreshTokens(tokens.refresh_token, tokens.scope);
    saveTokens(tokens);
  }

  return tokens.access_token;
}

export async function runAuthFlow(): Promise<void> {
  const url = new URL(REDIRECT_URI);
  const port = parseInt(url.port || "3000");

  return new Promise((resolve, reject) => {
    const server = createServer(async (req, res) => {
      try {
        const reqUrl = new URL(req.url || "/", REDIRECT_URI);
        const code = reqUrl.searchParams.get("code");
        const error = reqUrl.searchParams.get("error");

        if (error) {
          res.writeHead(400, { "Content-Type": "text/html" });
          res.end(`<h1>Error: ${error}</h1><p>${reqUrl.searchParams.get("error_description")}</p>`);
          server.close();
          reject(new Error(error));
          return;
        }

        if (!code) {
          res.writeHead(400, { "Content-Type": "text/html" });
          res.end("<h1>No code received</h1>");
          return;
        }

        const tokens = await exchangeCode(code);
        saveTokens(tokens);

        res.writeHead(200, { "Content-Type": "text/html" });
        res.end("<h1>✅ Authenticated!</h1><p>You can close this window.</p>");

        console.log(JSON.stringify({ ok: true, data: { message: "Authentication successful", expires_at: new Date(tokens.expires_at).toISOString() } }));

        server.close();
        resolve();
      } catch (err: any) {
        res.writeHead(500, { "Content-Type": "text/html" });
        res.end(`<h1>Error</h1><p>${err.message}</p>`);
        server.close();
        reject(err);
      }
    });

    server.listen(port, () => {
      const authUrl = buildAuthUrl();
      console.error(`\nOpen this URL in your browser to authenticate:\n\n${authUrl}\n`);
      console.error(`Waiting for callback on port ${port}...`);
    });
  });
}

// If run directly
if (process.argv[1] && (process.argv[1].endsWith("auth.ts") || process.argv[1].endsWith("auth.js"))) {
  runAuthFlow().catch((err) => {
    console.error(JSON.stringify({ ok: false, error: err.message }));
    process.exit(1);
  });
}
