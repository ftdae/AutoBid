import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const contentSource = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");
const gptWorkerSource = await readFile(new URL("../extension/gpt-answer-worker.js", import.meta.url), "utf8");
const manifestSource = await readFile(new URL("../extension/manifest.json", import.meta.url), "utf8");
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

test("a GPT delivery is finalized after one application pass and duplicate or stale pushes are ignored", () => {
  assert.match(contentSource, /const finalizedRuntimeGptRequestIds = new Set\(\)/);
  assert.match(contentSource, /runtime-gpt:request-finalized/);
  assert.match(contentSource, /runtime-gpt:push-ignored-finalized/);
  assert.match(contentSource, /runtime-gpt:push-ignored-stale-run/);
  assert.match(contentSource, /client_run_id: autoBidRunId/);
  assert.match(backgroundSource, /client_run_id: request\.client_run_id \|\| ""/);
  assert.match(contentSource, /if \(!isActiveContentInstance\(\)\) return false;[\s\S]*const first = controls\[0\]/);
  assert.match(contentSource, /reason: "superseded-run"/);
  assert.match(backgroundSource, /const holdReason = `autofill-run:\$\{frameId\}:\$\{runId\}`/);
});

test("each application page shows a persistent spinning autofill stage indicator", () => {
  const statusSource = contentSource.slice(
    contentSource.indexOf("function showStatus"),
    contentSource.indexOf("function hashSmall")
  );
  assert.match(contentSource, /showStatus\("Questions sent to GPT", "waiting"/);
  assert.match(contentSource, /showStatus\("Autofilling GPT answers", "autofilling"/);
  assert.match(contentSource, /"Autofill done"/);
  assert.match(statusSource, /dataset\.autoBidStatusKind = kind/);
  assert.match(statusSource, /icon\.animate\?\./);
  assert.match(statusSource, /Dismiss Auto Bid status/);
  assert.doesNotMatch(statusSource, /setTimeout\(\(\) => status\.remove/);
});

test("generated answers preserve every current manual value including choice controls", () => {
  const applySource = contentSource.slice(
    contentSource.indexOf("async function applyAnswers"),
    contentSource.indexOf("function getGeneratedAnswerAttemptKey")
  );
  assert.match(applySource, /if \(alreadyFilled\)/);
  assert.match(applySource, /reason: "preserve-current-value"/);
  assert.doesNotMatch(applySource, /replacing-mismatched-choice/);
});

test("later hotkey runs treat a selected required checkbox group as complete", () => {
  const filledSource = contentSource.slice(
    contentSource.indexOf("function getCurrentlyFilledFieldIds"),
    contentSource.indexOf("async function closeOpenChoiceMenus")
  );
  assert.match(filledSource, /const checkboxGroups = new Map\(\)/);
  assert.match(filledSource, /group\.some\(\(field\) => hasFieldCurrentValue\(field\)\)/);
  assert.match(filledSource, /group\.forEach\(\(field\) => filledIds\.add\(field\.id\)\)/);
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
  assert.match(contentSource, /function getCheckboxGroupQuestionLabel/);
  assert.match(contentSource, /getCheckboxGroupQuestionLabel\(control, checkboxGroups\.get\(checkboxName\) \|\| \[control\]\)/);
  assert.match(contentSource, /runtime-gpt:candidates-selected/);
});

test("Recruitee textareas keep their own visual question instead of a broad Questions section label", () => {
  const fieldLabelSource = contentSource.slice(
    contentSource.indexOf("function getFieldLabel"),
    contentSource.indexOf("function getNearbyText")
  );
  const recruiteeAdapterSource = atsAdapterSource.slice(
    atsAdapterSource.indexOf('adapter("recruitee"'),
    atsAdapterSource.indexOf('adapter("smartrecruiters"')
  );
  assert.match(contentSource, /function getVisualFieldLabel/);
  assert.match(fieldLabelSource, /visualLabel,[\s\S]*getLabelCandidates/);
  assert.match(contentSource, /const question = getPlainQuestionText\(visibleQuestion \|\| field\.question \|\| field\.label\)/);
  assert.doesNotMatch(recruiteeAdapterSource, /"\.question"/);
});

test("bare boolean answers cannot be written into non-boolean narrative fields", () => {
  assert.match(contentSource, /function isSuspiciousNarrativeBooleanAnswer/);
  assert.match(contentSource, /function isSemanticBooleanQuestion/);
  assert.match(contentSource, /answer:rejected-narrative-boolean/);
  assert.match(contentSource, /getVisualFieldLabel\(first\)/);
});

test("dynamic dropdown options are collected before an AI request", () => {
  assert.match(contentSource, /await hydrateGeneratedChoiceOptions\(candidateFields\)/);
  assert.match(contentSource, /ai:choice-options-hydrated/);
  assert.match(contentSource, /collectAllComboboxOptionTexts\(control, options\)/);
  assert.match(contentSource, /VIRTUALIZED_CHOICE_SCROLL_STEPS = 24/);
  assert.match(contentSource, /container\.dispatchEvent\(new Event\("scroll", \{ bubbles: true \}\)\)/);
  assert.match(contentSource, /MAX_COLLECTED_CHOICE_OPTIONS = 200/);
  assert.match(contentSource, /field\.options = uniqueNonEmptyValues/);
  assert.match(contentSource, /function hasCompleteGeneratedChoiceOptions/);
  assert.match(contentSource, /runtime-gpt:choice-options-unavailable/);
  assert.match(contentSource, /const requestableFields = unresolvedCandidateFields\.filter\(\(field\) => hasCompleteGeneratedChoiceOptions\(field\)\)/);
});

test("composite controls are canonicalized before questions are sent to GPT", () => {
  const collectSource = contentSource.slice(
    contentSource.indexOf("function collectFields"),
    contentSource.indexOf("function getFormControls")
  );
  assert.match(collectSource, /const canonicalControls = new Set\(\)/);
  assert.match(collectSource, /getCanonicalFieldControl\(control\)/);
  assert.match(collectSource, /fields:duplicate-control-skipped/);
  assert.match(contentSource, /return getComboboxShell\(control\) \|\| getComboboxInput\(control\) \|\| control/);
});

test("required static choice failures are promoted to GPT with their live options", () => {
  assert.match(contentSource, /const failedStaticChoiceFieldKeys = new Set\(\)/);
  assert.match(contentSource, /failedStaticChoiceFieldKeys\.clear\(\)/);
  assert.match(contentSource, /failedStaticChoiceFieldKeys\.add\(getStableFieldBindingKey\(field\)\)/);
  assert.match(contentSource, /isProfileStaticQuestionField\(field\) && !failedStaticChoiceFieldKeys\.has\(getStableFieldBindingKey\(field\)\)/);
  assert.match(contentSource, /await hydrateGeneratedChoiceOptions\(candidateFields\)/);
});

test("framework choice state and multi-select values are verified from rendered UI", () => {
  const stateSource = contentSource.slice(
    contentSource.indexOf("function getExplicitChoiceSelectionState"),
    contentSource.indexOf("function cleanLabel")
  );
  assert.match(stateSource, /data-automation-selected/);
  assert.match(stateSource, /aria-pressed/);
  assert.match(stateSource, /hasSelectedStateClass/);
  assert.doesNotMatch(stateSource, /\[class\*='active'/);

  const selectedTextSource = contentSource.slice(
    contentSource.indexOf("function getComboboxSelectedText"),
    contentSource.indexOf("function getComboboxValueText")
  );
  assert.match(selectedTextSource, /react-select__multi-value__label/);
  assert.match(selectedTextSource, /uniqueNonEmptyValues\(multiValues\)\.join\(", "\)/);
  assert.match(selectedTextSource, /select__value-container--has-value/);
});

test("the final required-field audit is read-only and reports newly rendered gaps", () => {
  const auditSource = contentSource.slice(
    contentSource.indexOf("function auditRequiredFieldStates"),
    contentSource.indexOf("function hasMissingRequiredResumeUpload")
  );
  assert.match(contentSource, /auditRequiredFieldStates\("initial", fields\)/);
  assert.match(contentSource, /const finalRequiredAudit = auditRequiredFieldStates\("final", fields\)/);
  assert.match(auditSource, /fields:required-audit/);
  assert.match(auditSource, /missing_fields: missing/);
  assert.doesNotMatch(auditSource, /setControlsValue|applyAnswers|click\(/);
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

test("dropdown fallbacks use page-world actions without Chrome debugger attachment", () => {
  assert.match(backgroundSource, /func: clickPagePoint/);
  assert.match(backgroundSource, /func: typeIntoFocusedPageControl/);
  assert.doesNotMatch(backgroundSource, /chrome\.debugger/);
  assert.doesNotMatch(manifestSource, /"debugger"/);
});

test("inactive GPT and application tabs stay non-discardable without a debugger banner", () => {
  assert.match(backgroundSource, /autoDiscardable: false/);
  assert.match(backgroundSource, /debugger-free tab protection enabled/);
  assert.match(backgroundSource, /RUNTIME_GPT_DELIVERY_ATTEMPTS = 3/);
  assert.match(contentSource, /return true;\s*\n\s*}\s*\n\s*return false;/);
});

test("inactive ChatGPT tabs submit immediately in the existing conversation", () => {
  const askSource = gptWorkerSource.slice(
    gptWorkerSource.indexOf("async function askChatGpt"),
    gptWorkerSource.indexOf("async function saveAnswersWithRetry")
  );
  assert.match(askSource, /await setComposerText\(composer, prompt\);[\s\S]*await clickSend\(composer\);/);
  assert.doesNotMatch(askSource, /sleepInterruptible\(250\)/);
  assert.doesNotMatch(gptWorkerSource, /startFreshChat|findNewChatButton|AUTOBID_GPT_NATIVE_SUBMIT/);
  assert.match(gptWorkerSource, /button\.click\(\)/);
  assert.match(gptWorkerSource, /form\.requestSubmit\(\)/);
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
