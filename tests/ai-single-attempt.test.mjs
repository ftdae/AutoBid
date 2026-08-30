import assert from "node:assert/strict";
import test from "node:test";

test("a partial OpenAI second-router result stops after one request", async () => {
  process.env.OPENAI_API_KEY = process.env.OPENAI_API_KEY || "test-openai-key";
  process.env.OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5-mini";

  const calls = [];
  globalThis.fetch = async (url, options) => {
    const parsedUrl = new URL(url);
    const body = JSON.parse(options.body);
    const payload = JSON.parse(body.input);
    calls.push({
      host: parsedUrl.hostname,
      fieldIds: payload.fields.map((field) => field.field_id),
      store: body.store,
      format: body.text?.format?.type
    });

    const result = JSON.stringify({
      answers: [{
        field_id: "required_a",
        value: "OpenAI answer",
        cache_scope: "profile_job",
        confidence: 0.9,
        warning: null
      }],
      warnings: []
    });
    return new Response(JSON.stringify({
      output: [{ type: "message", content: [{ type: "output_text", text: result }] }],
      usage: { input_tokens: 100, output_tokens: 20, total_tokens: 120 }
    }), { status: 200, headers: { "content-type": "application/json" } });
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

  assert.deepEqual(calls, [{
    host: "api.openai.com",
    fieldIds: ["required_a", "required_b"],
    store: false,
    format: "json_schema"
  }]);
  assert.deepEqual(answers.map(({ field_id, provider }) => ({ field_id, provider })), [
    { field_id: "required_a", provider: "openai" }
  ]);
});
