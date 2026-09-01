import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const env = loadEnv(path.join(__dirname, "..", ".env"));

export const PORT = Number(env.PORT || process.env.PORT || 7002);
export const APP_SECRET = env.APP_SECRET || process.env.APP_SECRET || "dev-auto-bid-secret";
export const DATABASE_URL = env.DATABASE_URL || process.env.DATABASE_URL || "postgres://postgres:postgres@localhost:5432/autobid";
export const DATABASE_POOL_MAX = normalizePositiveInt(env.DATABASE_POOL_MAX || process.env.DATABASE_POOL_MAX, 50);
export const GEMINI_API_KEY = env.GEMINI_API_KEY || process.env.GEMINI_API_KEY || "";
export const GEMINI_MODEL = env.GEMINI_MODEL || process.env.GEMINI_MODEL || "gemini-2.5-flash-lite";
export const GEMINI_MODELS = normalizeList(env.GEMINI_MODELS || process.env.GEMINI_MODELS || GEMINI_MODEL);
export const OPENAI_API_KEY = env.OPENAI_API_KEY || env.GPT_API_KEY || process.env.OPENAI_API_KEY || process.env.GPT_API_KEY || "";
export const OPENAI_MODEL = env.OPENAI_MODEL || process.env.OPENAI_MODEL || "gpt-5-mini";
export const OPENAI_ROUTE_ENABLED = normalizeBoolean(env.OPENAI_ROUTE_ENABLED || process.env.OPENAI_ROUTE_ENABLED, false);
export const GOOGLE_APPS_SCRIPT_WEB_APP_URL = env.GOOGLE_APPS_SCRIPT_WEB_APP_URL || env.APPS_SCRIPT_WEB_APP_URL || process.env.GOOGLE_APPS_SCRIPT_WEB_APP_URL || process.env.APPS_SCRIPT_WEB_APP_URL || "";
export const GOOGLE_APPS_SCRIPT_SECRET = env.GOOGLE_APPS_SCRIPT_SECRET || env.CHATGPT_EXTENSION_SECRET || env.EXTENSION_SECRET || process.env.GOOGLE_APPS_SCRIPT_SECRET || process.env.CHATGPT_EXTENSION_SECRET || process.env.EXTENSION_SECRET || "";
export const GOOGLE_APPS_SCRIPT_CONNECT_HOST = env.GOOGLE_APPS_SCRIPT_CONNECT_HOST || process.env.GOOGLE_APPS_SCRIPT_CONNECT_HOST || "www.google.com";
export const GOOGLE_APPS_SCRIPT_TIMEOUT_MS = normalizePositiveInt(env.GOOGLE_APPS_SCRIPT_TIMEOUT_MS || process.env.GOOGLE_APPS_SCRIPT_TIMEOUT_MS, 30_000);
export const GOOGLE_SHEETS_SPREADSHEET_ID = env.GOOGLE_SHEETS_SPREADSHEET_ID || process.env.GOOGLE_SHEETS_SPREADSHEET_ID || "";
export const GOOGLE_SERVICE_ACCOUNT_EMAIL = env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "";
export const GOOGLE_PRIVATE_KEY = normalizePrivateKey(env.GOOGLE_PRIVATE_KEY || process.env.GOOGLE_PRIVATE_KEY || "");
export const MICROSOFT_CLIENT_ID = env.MICROSOFT_CLIENT_ID || process.env.MICROSOFT_CLIENT_ID || "";
export const MICROSOFT_CLIENT_SECRET = env.MICROSOFT_CLIENT_SECRET || process.env.MICROSOFT_CLIENT_SECRET || "";
export const MICROSOFT_TENANT_ID = env.MICROSOFT_TENANT_ID || process.env.MICROSOFT_TENANT_ID || "common";
export const MICROSOFT_OUTLOOK_SCOPES = normalizeList(
  env.MICROSOFT_OUTLOOK_SCOPES || process.env.MICROSOFT_OUTLOOK_SCOPES || "openid,profile,email,offline_access,User.Read,Mail.ReadWrite"
);
export const DEV_AUTH_BYPASS = (env.DEV_AUTH_BYPASS || process.env.DEV_AUTH_BYPASS || "true") !== "false";
export const DEV_USER_EMAIL = normalizeEmail(env.DEV_USER_EMAIL || process.env.DEV_USER_EMAIL || "dev@autobid.local");
export const TOKEN_VERSION = "ab1";
export const TOKEN_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
export const CACHE_SCOPES = ["global", "profile", "profile_job"];

export function loadEnv(filePath) {
  try {
    const content = readFileSync(filePath, "utf8");
    return Object.fromEntries(
      content
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line && !line.startsWith("#") && line.includes("="))
        .map((line) => {
          const index = line.indexOf("=");
          return [line.slice(0, index).trim(), line.slice(index + 1).trim().replace(/^["']|["']$/g, "")];
        })
    );
  } catch {
    return {};
  }
}

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function normalizePrivateKey(value) {
  return String(value || "").replace(/\\n/g, "\n");
}

function normalizeList(value) {
  return String(value || "")
    .split(/[\s,]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizePositiveInt(value, fallback) {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function normalizeBoolean(value, fallback = false) {
  if (value == null || value === "") return fallback;
  return /^(1|true|yes|on)$/i.test(String(value).trim());
}
