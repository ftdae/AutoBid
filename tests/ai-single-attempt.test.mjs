import assert from "node:assert/strict";
import test from "node:test";

test("the disabled OpenAI route makes no API request", async () => {
  process.env.OPENAI_API_KEY = process.env.OPENAI_API_KEY || "test-openai-key";
  process.env.OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5-mini";
  process.env.OPENAI_ROUTE_ENABLED = "false";

  const calls = [];
  globalThis.fetch = async (url) => {
    calls.push(String(url));
    throw new Error("OpenAI fetch must remain disabled");
  };

  const { generateAiAnswers } = await import(`../server/assist/ai.js?single-attempt=${Date.now()}`);
  const fields = [
    { id: "required_a", label: "Why this role?", type: "textarea", required: true, options: [], cache_scope: "profile_job" },
    { id: "required_b", label: "Describe your experience", type: "textarea", required: true, options: [], cache_scope: "profile_job" }
  ];
  const warnings = [];
  const answers = await generateAiAnswers(
    fields,
    { name: "Test", static_fields: {}, resume_text: "", preferences: {}, profile_version: 1 },
    { url: "https://jobs.example.test/1", title: "Job", job_title: "Engineer", text: "" },
    "job-hash",
    warnings
  );

  assert.deepEqual(calls, []);
  assert.deepEqual(answers, []);
  assert.match(warnings.join(" "), /OpenAI autofill routing is disabled/i);
});
