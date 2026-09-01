import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");
const workerSource = await readFile(new URL("../extension/gpt-answer-worker.js", import.meta.url), "utf8");

test("three persistent GPT workers process one application per prompt", () => {
  assert.match(backgroundSource, /const RUNTIME_GPT_PROMPT_BATCH_SIZE = 1;/);
  assert.match(backgroundSource, /const RUNTIME_GPT_MAX_WORKERS = 3;/);
  assert.match(workerSource, /const MAX_REQUESTS_PER_PROMPT = 1;/);
  assert.match(workerSource, /one application per prompt/i);
});

test("runtime GPT keeps valid partial answers, rejects placeholders, and can recover a changed field id by question", () => {
  assert.match(workerSource, /function resolveAnswerField/);
  assert.match(workerSource, /const field = resolveAnswerField\(answer, fieldList, fieldsById\)/);
  assert.match(workerSource, /return matches\.length === 1 \? matches\[0\] : null/);
  assert.match(workerSource, /\{ requireComplete: false \}/);
  assert.match(workerSource, /function isRejectedPlaceholderAnswer/);
  assert.match(workerSource, /ChatGPT returned a partial answer set/);
  assert.doesNotMatch(workerSource, /throw new Error\(`ChatGPT response omitted \$\{missingFieldIds\.length\}/);
});
