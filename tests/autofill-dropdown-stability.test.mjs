import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const contentSource = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");
const atsAdapterSource = await readFile(new URL("../extension/content-modules/ats-adapters.js", import.meta.url), "utf8");

test("generated answers stop retrying a field after three failed fills", () => {
  assert.match(contentSource, /const MAX_FIELD_FILL_ATTEMPTS = 3;/);
  assert.match(contentSource, /answer:stopped-after-max-attempts/);
  assert.match(contentSource, /answer:invalid-choice/);
  assert.match(contentSource, /answer:verification-failed/);
  assert.match(contentSource, /attempt >= MAX_FIELD_FILL_ATTEMPTS/);

  const pushedAnswerHandler = contentSource.slice(
    contentSource.indexOf("async function applyPushedRuntimeGptAnswers"),
    contentSource.indexOf("function applyRuntimeGptAnswersOnce")
  );
  assert.match(pushedAnswerHandler, /getRuntimeGptAnswerSettlement/);
  assert.doesNotMatch(pushedAnswerHandler, /applyProfileStaticFallbacks|reapplyRuntimeGptAnswers/);
});

test("Ashby-style checkbox questions use explicit verified Yes or No choices", () => {
  assert.match(contentSource, /const CHOICE_FIELD_TYPES = \["select", "radio", "checkbox", "combobox", "button-group"\]/);
  assert.match(contentSource, /function getCheckboxBooleanChoiceOptions/);
  assert.match(contentSource, /function getSelectedCheckboxBooleanChoiceLabel/);
  assert.match(contentSource, /const shouldCheck = parseBooleanAnswer\(value\)/);
  assert.match(contentSource, /if \(!option && shouldCheck === false\) return false/);
  assert.match(contentSource, /await waitForGeneratedAnswerMatch/);
  assert.match(contentSource, /doesGeneratedAnswerMatchField\(answer, binding\.field, binding\.controls\)/);
});

test("custom choices use verified native, DOM, and keyboard activation fallbacks", () => {
  assert.match(contentSource, /async function activateChoiceTarget/);
  assert.match(contentSource, /for \(const key of \[" ", "Enter"\]\)/);
  assert.match(contentSource, /await activateChoiceTarget\(target, \(\) => isRadioChecked\(radio\)\)/);
});

test("generated placeholder answers are rejected but explicitly permitted N-A fields are filled locally", () => {
  assert.match(contentSource, /function isRejectedGeneratedPlaceholder/);
  assert.match(contentSource, /not specified\|unspecified\|unknown\|not provided/);
  assert.match(contentSource, /applyExplicitNotApplicableAnswers/);
  assert.match(contentSource, /explicitlyAllowsNotApplicable/);
});

test("optional fields cannot inherit a required marker from sibling controls", () => {
  const requiredSource = contentSource.slice(
    contentSource.indexOf("function isRequired"),
    contentSource.indexOf("function hasRequiredSemanticMarker")
  );
  assert.doesNotMatch(requiredSource, /getNearbyText\(control\)/);
  assert.match(contentSource, /if \(nestedControls > 1/);
  assert.match(atsAdapterSource, /const scopedContainer = nestedControls <= 1/);
});

test("combined identity and location prompts map to complete saved profile values", () => {
  assert.match(contentSource, /function isCombinedProfileLocationField/);
  assert.match(contentSource, /return "location"/);
  assert.match(contentSource, /first\(\?:\\s\+and\|\\s\*\\\/\)\\s\*last/);
});

test("language checkbox scales select one supported level and localized None for unsupported languages", () => {
  assert.match(contentSource, /\["italian", \["italian", "italiano"\]\]/);
  assert.match(contentSource, /field\?\.type === "checkbox" && field\.option/);
  assert.match(contentSource, /nessuno/);
  assert.match(contentSource, /getPreferredLanguageProficiencyOption/);
});

test("semantic availability questions are not claimed by saved start availability", () => {
  assert.match(contentSource, /function isSemanticAvailabilityQuestion/);
  assert.doesNotMatch(contentSource, /\["availability", \["availability", "available", "start date"\]\]/);
  assert.match(contentSource, /type === "checkbox" && parseBooleanAnswer\(rawValue\) === null/);
});

test("required checkbox groups inherit the star from their question label", () => {
  assert.match(contentSource, /required: isRequired\(control\) \|\| Boolean\(choiceQuestionLabel && \/\\\*\/\.test\(choiceQuestionLabel\)\)/);
  assert.match(contentSource, /runtime-gpt:candidates-selected/);
});

test("dynamic dropdown options are collected before an AI request", () => {
  assert.match(contentSource, /await hydrateGeneratedChoiceOptions\(candidateFields\)/);
  assert.match(contentSource, /ai:choice-options-hydrated/);
  assert.match(contentSource, /field\.options = uniqueNonEmptyValues/);
});

test("a native click is verified and falls back to the page bridge before keyboard opening", () => {
  const openComboboxSource = contentSource.slice(
    contentSource.indexOf("async function openCombobox"),
    contentSource.indexOf("async function applyComboboxFilter")
  );
  assert.match(openComboboxSource, /waitForComboboxOptions\(control, 500\)/);
  assert.match(openComboboxSource, /runPageCommand\("combobox-open", trigger\)/);
  assert.match(openComboboxSource, /dropdown:open-after-page-bridge/);
  assert.match(openComboboxSource, /for \(const key of \["ArrowDown", " "\]\)/);
});

test("plain-text based-in questions receive a local Yes or No answer", () => {
  const basedInSource = contentSource.slice(
    contentSource.indexOf("async function applyBasedInLocationAnswers"),
    contentSource.indexOf("async function applySensitiveDemographicDeclineAnswers")
  );
  const classifierSource = contentSource.slice(
    contentSource.indexOf("function isBasedInLocationField"),
    contentSource.indexOf("function isLanguageChoiceField")
  );
  assert.match(basedInSource, /isSemanticBooleanFieldType\(field\.type\)/);
  assert.match(basedInSource, /locationAnswerMatchesQuestion\(locationAnswer, field\.question \|\| field\.label\) \? "Yes" : "No"/);
  assert.doesNotMatch(classifierSource, /hasYesNoOptions/);
});

test("required demographic choices use only an explicit non-disclosure option", () => {
  assert.match(contentSource, /sensitive-demographic-decline/);
  assert.match(contentSource, /function findSensitiveDemographicDeclineOption/);
  assert.match(contentSource, /prefer not to say\|prefer not to answer\|decline to self identify/);
  assert.match(contentSource, /if \(!answer\)[\s\S]*sensitive-demographic:no-decline-option/);
});

test("complex required qualification choices are deferred to ChatGPT", () => {
  const deferSource = contentSource.slice(
    contentSource.indexOf("function shouldDeferChoiceFieldToRuntimeGpt"),
    contentSource.indexOf("function shouldUsePositiveCheckboxFallback")
  );
  assert.match(deferSource, /experience\|years\|skill\|proficien/);
  assert.match(deferSource, /sponsor\|sponsorship\|visa\|work permit\|right to work/);
  assert.match(contentSource, /items\.filter\(\(field\) => !shouldDeferChoiceFieldToRuntimeGpt\(field\)\)/);
});

test("deterministic experience matching cannot inherit nearby unrelated questions", () => {
  assert.match(contentSource, /function getDirectFieldLabel/);
  assert.match(contentSource, /const label = getDirectFieldLabel\(field, control\);/);
  assert.match(atsAdapterSource, /adapter\("greenhouse", "Greenhouse"/);
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

test("current BambooHR Fabric selects expose state and country choices", () => {
  assert.match(atsAdapterSource, /\.fab-SelectToggle/);
  assert.match(atsAdapterSource, /\.fab-MenuOption/);
  assert.match(contentSource, /getProfileAddressFillPriority/);
  assert.match(contentSource, /if \(key === "country"\) return 0/);
  assert.match(contentSource, /if \(key === "state_region"\) return 1/);
  assert.match(contentSource, /masovian voivodeship/);
});

test("Workable phone entry receives a national number after the dial code selector", () => {
  assert.match(contentSource, /function getPhoneEntryCandidateValues/);
  assert.match(contentSource, /hasSeparateDialCodeControl/);
  assert.match(contentSource, /uniqueNonEmptyValues\(\[nationalDigits, nationalFormatted, raw\]\)/);
  assert.match(contentSource, /function areTextControlValuesEquivalent/);
  assert.match(contentSource, /currentDigits === expectedDigits/);
});

test("NewRocket dropdowns are hydrated and deferred to GPT instead of guessed locally", () => {
  assert.match(atsAdapterSource, /adapter\("newrocket", "NewRocket Greenhouse"/);
  assert.match(atsAdapterSource, /\.select2-results__option/);
  assert.match(contentSource, /function shouldDeferChoiceFieldToRuntimeGpt/);
  assert.match(contentSource, /if \(shouldDeferChoiceFieldToRuntimeGpt\(field\)\) return false/);
  assert.match(contentSource, /await hydrateGeneratedChoiceOptions\(candidateFields\)/);
});
