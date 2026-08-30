import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { normalizeFields, shouldAnswerWithAi } from "../server/assist/field-policy.js";
import { matchStaticFieldKey } from "../server/profiles/static-fields.js";

test("autofill executes ChatGPT before the OpenAI second router", async () => {
  const source = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
  const workerSource = await readFile(new URL("../extension/gpt-answer-worker.js", import.meta.url), "utf8");
  const chatGptIndex = source.indexOf('"chatgpt-first-provider"');
  const openAiIndex = source.indexOf('"openai-second-provider"');

  assert.ok(chatGptIndex >= 0);
  assert.ok(openAiIndex > chatGptIndex);
  assert.match(source, /max_attempts:\s*1/);
  assert.match(workerSource, /requireComplete:\s*!isRuntimeRequest\(row\)/);
});

test("complex required BambooHR textareas reach the API router but basic and optional fields do not", () => {
  const rawFields = [
    {
      id: "architecture",
      label: "Describe a full-stack feature you designed. How did you structure the front-end, API layer, and database?",
      type: "textarea",
      required: true
    },
    {
      id: "ai_workflow",
      label: "How have you used AI-generated code in your development workflow?",
      type: "textarea",
      required: true
    },
    {
      id: "production_process",
      label: "Walk me through how you take a feature from concept to production.",
      type: "textarea",
      required: true
    },
    { id: "address", label: "Current address", type: "textarea", required: true },
    { id: "optional", label: "Describe your interests", type: "textarea", required: false }
  ];
  const fields = normalizeFields(rawFields, "jobs.example.test", "https://jobs.example.test/1");
  const decisions = Object.fromEntries(fields.map((field) => [field.id, shouldAnswerWithAi(field)]));

  assert.deepEqual(decisions, {
    architecture: true,
    ai_workflow: true,
    production_process: true,
    address: false,
    optional: false
  });
});

test("AI-generated does not match expected rate", () => {
  assert.equal(matchStaticFieldKey({
    label: "How have you used AI-generated code in your development workflow?",
    type: "textarea"
  }), null);
  assert.equal(matchStaticFieldKey({ label: "Expected rate", type: "text" }), "expected_rate");
});
