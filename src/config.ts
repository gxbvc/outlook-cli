import { config } from "dotenv";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

config({ path: join(__dirname, "..", ".env") });

function requireEnv(name: string): string {
  const val = process.env[name];
  if (!val) {
    console.error(JSON.stringify({ ok: false, error: `Missing env var: ${name}`, code: "MISSING_ENV" }));
    process.exit(1);
  }
  return val;
}

export const CLIENT_ID = requireEnv("OUTLOOK_CLIENT_ID");
export const CLIENT_SECRET = requireEnv("OUTLOOK_CLIENT_SECRET");
export const TENANT_ID = requireEnv("OUTLOOK_TENANT_ID");
export const REDIRECT_URI = process.env["OUTLOOK_REDIRECT_URI"] || "http://localhost:3000/callback";

// Use "common" for multi-tenant + personal accounts
export const AUTHORITY = "https://login.microsoftonline.com/common";
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const SCOPES = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadWrite", "https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read", "offline_access"];

export const TOKENS_DIR = join(__dirname, "..", "tokens");
export const TOKENS_FILE = join(TOKENS_DIR, "tokens.json");
