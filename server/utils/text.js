import { createHash } from "node:crypto";

export function extractOutputText(data) {
  if (typeof data?.output_text === "string") return data.output_text;
  if (!Array.isArray(data?.output)) return "";
  return data.output
    .flatMap((item) => Array.isArray(item?.content) ? item.content : [])
    .map((content) => content?.text || "")
    .filter(Boolean)
    .join("");
}

export function normalizeFieldType(type) {
  const normalized = String(type || "text").toLowerCase();
  if (normalized === "select-one" || normalized === "select-multiple") return "select";
  return normalized;
}

export function normalizeText(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

export function normalizeDomain(domain) {
  return String(domain || "").toLowerCase().replace(/^www\./, "");
}

export function normalizeUrl(url) {
  try {
    const parsed = new URL(url);
    const cleanPath = parsed.pathname.replace(/\/+$/, "") || "/";
    return `${parsed.origin}${cleanPath}`.toLowerCase();
  } catch {
    return String(url || "").split("?")[0].toLowerCase();
  }
}

export function safeDomain(url) {
  try {
    return normalizeDomain(new URL(url).hostname);
  } catch {
    return "";
  }
}

export function hashText(text) {
  return createHash("sha256").update(String(text)).digest("hex");
}

export function trimForPrompt(text, maxLength) {
  const value = String(text || "").trim();
  return value.length > maxLength ? value.slice(0, maxLength) : value;
}
