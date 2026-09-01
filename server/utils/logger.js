import { randomUUID } from "node:crypto";

const HTTP_LOG_META = Symbol("autoBidHttpLogMeta");
const SENSITIVE_KEY = /(password|secret|token|authorization|cookie|api[_-]?key|private[_-]?key|client[_-]?secret|verification[_-]?code|passcode|^otp$|^codes?$)/i;
const MAX_STRING_LENGTH = 600;
const MAX_ARRAY_ITEMS = 30;
const MAX_OBJECT_KEYS = 60;
const MAX_DEPTH = 6;

export function beginHttpRequestLog(req, res) {
  const requestId = createLogId("req");
  const startedAt = Date.now();
  const meta = {
    request_id: requestId,
    started_at: startedAt,
    response_payload: undefined
  };
  res[HTTP_LOG_META] = meta;

  logBackendEvent("HTTP_REQUEST", {
    method: req.method || "GET",
    url: req.url || "/"
  }, { requestId });

  res.once("finish", () => {
    logBackendEvent("HTTP_RESPONSE", {
      method: req.method || "GET",
      url: req.url || "/",
      status: res.statusCode,
      duration_ms: Date.now() - startedAt,
      payload: meta.response_payload
    }, {
      requestId,
      level: res.statusCode >= 500 ? "error" : res.statusCode >= 400 ? "warn" : "info"
    });
  });

  return requestId;
}

export function logHttpRequestBody(res, body) {
  const requestId = getHttpRequestId(res);
  if (!body || typeof body !== "object" || Object.keys(body).length === 0) return;
  logBackendEvent("HTTP_REQUEST_BODY", { body }, { requestId });
}

export function recordHttpResponsePayload(res, payload) {
  if (res?.[HTTP_LOG_META]) res[HTTP_LOG_META].response_payload = payload;
}

export function getHttpRequestId(res) {
  return res?.[HTTP_LOG_META]?.request_id || "request-unknown";
}

export function logBackendEvent(stage, data = {}, options = {}) {
  const timestamp = new Date().toISOString();
  const requestId = String(options.requestId || options.request_id || "system");
  const level = ["error", "warn", "info"].includes(options.level) ? options.level : "info";
  const safeData = sanitizeForLog(data);
  const message = `${timestamp} [AutoBid] [${requestId}] ${String(stage || "EVENT")} ${JSON.stringify(safeData)}`;
  console[level](message);
}

export function sanitizeForLog(value, key = "", depth = 0) {
  if (SENSITIVE_KEY.test(String(key || ""))) return "[REDACTED]";
  if (value === null || value === undefined) return value;
  if (value instanceof Error) {
    return {
      name: value.name,
      message: truncateString(value.message),
      code: value.code || undefined,
      cause: value.cause ? sanitizeForLog(value.cause, "error_cause", depth + 1) : undefined
    };
  }
  if (typeof value === "string") return truncateString(value);
  if (["number", "boolean"].includes(typeof value)) return value;
  if (typeof value === "bigint") return String(value);
  if (depth >= MAX_DEPTH) return "[MAX_DEPTH]";

  if (Array.isArray(value)) {
    const items = value.slice(0, MAX_ARRAY_ITEMS)
      .map((item) => sanitizeForLog(item, key, depth + 1));
    if (value.length > MAX_ARRAY_ITEMS) items.push(`[${value.length - MAX_ARRAY_ITEMS} more items]`);
    return items;
  }

  if (typeof value === "object") {
    const keys = Object.keys(value);
    const entries = Object.entries(value).slice(0, MAX_OBJECT_KEYS);
    const result = Object.fromEntries(entries.map(([childKey, childValue]) => [
      childKey,
      sanitizeForLog(childValue, childKey, depth + 1)
    ]));
    if (keys.length > MAX_OBJECT_KEYS) result.__truncated_keys = keys.length - MAX_OBJECT_KEYS;
    return result;
  }

  return truncateString(String(value));
}

function createLogId(prefix) {
  if (typeof randomUUID === "function") return `${prefix}_${randomUUID()}`;
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function truncateString(value) {
  const text = String(value || "");
  if (text.length <= MAX_STRING_LENGTH) return text;
  return `${text.slice(0, MAX_STRING_LENGTH)}…[${text.length - MAX_STRING_LENGTH} more chars]`;
}
