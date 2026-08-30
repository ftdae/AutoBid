import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const contentSource = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");

test("generated answers stop retrying a field after three failed fills", () => {
  assert.match(contentSource, /const MAX_FIELD_FILL_ATTEMPTS = 3;/);
  assert.match(contentSource, /answer:stopped-after-max-attempts/);
  assert.match(contentSource, /answer:invalid-choice/);
  assert.match(contentSource, /attempt >= MAX_FIELD_FILL_ATTEMPTS/);

  const pushedAnswerHandler = contentSource.slice(
    contentSource.indexOf("async function applyPushedRuntimeGptAnswers"),
    contentSource.indexOf("function applyRuntimeGptAnswersOnce")
  );
  assert.match(pushedAnswerHandler, /getRuntimeGptAnswerSettlement/);
  assert.doesNotMatch(pushedAnswerHandler, /applyProfileStaticFallbacks|reapplyRuntimeGptAnswers/);
});

test("dynamic dropdown options are collected before an AI request", () => {
  assert.match(contentSource, /await hydrateGeneratedChoiceOptions\(candidateFields\)/);
  assert.match(contentSource, /ai:choice-options-hydrated/);
  assert.match(contentSource, /field\.options = uniqueNonEmptyValues/);
});

test("location autocomplete primes with country for two seconds before selecting", () => {
  assert.match(contentSource, /const LOCATION_AUTOCOMPLETE_WAIT_MS = 2000;/);
  assert.match(contentSource, /location-autocomplete:country-prime/);
  assert.match(contentSource, /filterWaitMs: LOCATION_AUTOCOMPLETE_WAIT_MS/);
  assert.match(contentSource, /text\.includes\(answer\) && text\.includes\(country\)/);
});

test("native dropdown input reconnects after Chrome detaches its debugger session", () => {
  assert.match(backgroundSource, /chrome\.debugger\?\.onDetach\?\.addListener/);
  assert.match(backgroundSource, /isDetachedNativeDebuggerError/);
  assert.match(backgroundSource, /attempt <= 2/);
});

test("inactive GPT and application tabs receive lifecycle and focus emulation", () => {
  assert.match(backgroundSource, /Page\.setWebLifecycleState/);
  assert.match(backgroundSource, /Emulation\.setFocusEmulationEnabled/);
  assert.match(backgroundSource, /autoDiscardable: false/);
  assert.match(backgroundSource, /RUNTIME_GPT_DELIVERY_ATTEMPTS = 3/);
  assert.match(contentSource, /return true;\s*\n\s*}\s*\n\s*return false;/);
});
