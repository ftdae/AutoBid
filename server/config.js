import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const env = loadEnv(path.join(__dirname, "..", ".env"));

export const PORT = Number(env.PORT || process.env.PORT || 7002);
export const APP_SECRET = env.APP_SECRET || process.env.APP_SECRET || "dev-auto-bid-secret";
export const DATABASE_URL = env.DATABASE_URL || process.env.DATABASE_URL || "postgres://postgres:postgres@localhost:5432/autobid";
export const OPENAI_API_KEY = env.OPENAI_API_KEY || env.GPT_API_KEY || process.env.OPENAI_API_KEY || process.env.GPT_API_KEY || "";
export const OPENAI_MODEL = env.OPENAI_MODEL || process.env.OPENAI_MODEL || "gpt-5-mini";
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
