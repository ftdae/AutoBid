import assert from "node:assert/strict";
import test from "node:test";
import { sanitizeForLog } from "../server/utils/logger.js";

test("backend logs redact secrets and truncate large values", () => {
  const safe = sanitizeForLog({
    password: "do-not-print",
    api_key: "do-not-print",
    nested: {
      authorization: "Bearer private",
      page_text: "x".repeat(900)
    },
    answers: Array.from({ length: 35 }, (_, index) => ({ field_id: `field_${index}` }))
  });

  assert.equal(safe.password, "[REDACTED]");
  assert.equal(safe.api_key, "[REDACTED]");
  assert.equal(safe.nested.authorization, "[REDACTED]");
  assert.match(safe.nested.page_text, /more chars/);
  assert.equal(safe.answers.length, 31);
  assert.equal(safe.answers.at(-1), "[5 more items]");
});

test("backend logs preserve nested network error causes", () => {
  const cause = Object.assign(new Error("connection timed out"), { code: "ETIMEDOUT" });
  const error = new TypeError("fetch failed", { cause });
  const safe = sanitizeForLog({ error });

  assert.equal(safe.error.message, "fetch failed");
  assert.equal(safe.error.cause.message, "connection timed out");
  assert.equal(safe.error.cause.code, "ETIMEDOUT");
});
