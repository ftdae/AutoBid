(() => {
const AUTO_BID_INSTANCE_ID = `ab_${Date.now()}_${Math.random().toString(36).slice(2)}`;
window.__autoBidActiveContentInstance = AUTO_BID_INSTANCE_ID;
window.__autoBidContentLoaded = true;

const STATUS_ID = "auto-bid-assistant-status";
const DEBUG_ATTR = "data-auto-bid-debug";
const EU_COUNTRY_NAMES = new Set([
  "austria",
  "belgium",
  "bulgaria",
  "croatia",
  "cyprus",
  "czech republic",
  "czechia",
  "denmark",
  "estonia",
  "finland",
  "france",
  "germany",
  "greece",
  "hungary",
  "ireland",
  "italy",
  "latvia",
  "lithuania",
  "luxembourg",
  "malta",
  "netherlands",
  "poland",
  "portugal",
  "romania",
  "slovakia",
  "slovenia",
  "spain",
  "sweden"
]);
const DROPDOWN_OPEN_TIMEOUT_MS = 2200;
const DROPDOWN_SELECT_TIMEOUT_MS = 2200;
const DROPDOWN_SETTLE_MS = 450;
const LOCATION_AUTOCOMPLETE_WAIT_MS = 2000;
const MAX_FIELD_FILL_ATTEMPTS = 3;
const SHEET_ANSWER_FIRST_WAIT_MS = 10000;
const SHEET_ANSWER_RETRY_ATTEMPTS = 3;
const SHEET_ANSWER_RETRY_MS = 10000;
const RUNTIME_GPT_ANSWER_TIMEOUT_MS = 90000;
const RUNTIME_GPT_ANSWER_POLL_MS = 1000;
const OPENAI_AUTOFILL_ROUTE_ENABLED = false;
const SHEET_CONTEXT_TIMEOUT_MS = 5000;
const RESUME_ATTACH_TIMEOUT_MS = 30000;
const RESUME_FILE_INPUT_WAIT_MS = 2500;
const RESUME_FILE_INPUT_RETRY_MS = 250;
const RESUME_SERVER_CONFIRM_TIMEOUT_MS = 6000;
const RESUME_MANAGED_UPLOAD_TIMEOUT_MS = 30000;
const OUTLOOK_VERIFICATION_FIELD_WAIT_MS = 12000;
const OUTLOOK_VERIFICATION_BUTTON_WAIT_MS = 8000;
const OUTLOOK_VERIFICATION_LOOKBACK_MS = 4 * 60 * 1000;
const CHOICE_FIELD_TYPES = ["select", "radio", "checkbox", "combobox", "button-group"];
const PHONE_DIAL_CODES_BY_COUNTRY = {
  albania: "+355",
  austria: "+43",
  belgium: "+32",
  bulgaria: "+359",
  canada: "+1",
  croatia: "+385",
  cyprus: "+357",
  "czech republic": "+420",
  czechia: "+420",
  denmark: "+45",
  estonia: "+372",
  finland: "+358",
  france: "+33",
  germany: "+49",
  greece: "+30",
  hungary: "+36",
  ireland: "+353",
  italy: "+39",
  latvia: "+371",
  lithuania: "+370",
  luxembourg: "+352",
  malta: "+356",
  netherlands: "+31",
  norway: "+47",
  poland: "+48",
  portugal: "+351",
  romania: "+40",
  slovakia: "+421",
  slovenia: "+386",
  spain: "+34",
  sweden: "+46",
  switzerland: "+41",
  ukraine: "+380",
  "united kingdom": "+44",
  uk: "+44",
  "united states": "+1",
  usa: "+1"
};
const LANGUAGE_ALIASES = [
  ["english", ["english", "inglese"]],
  ["ukrainian", ["ukrainian"]],
  ["polish", ["polish"]],
  ["russian", ["russian"]],
  ["spanish", ["spanish"]],
  ["portuguese", ["portuguese"]],
  ["german", ["german"]],
  ["french", ["french"]],
  ["italian", ["italian", "italiano"]],
  ["dutch", ["dutch"]],
  ["croatian", ["croatian"]],
  ["czech", ["czech"]],
  ["slovak", ["slovak"]],
  ["romanian", ["romanian"]],
  ["bulgarian", ["bulgarian"]],
  ["greek", ["greek"]],
  ["turkish", ["turkish"]],
  ["arabic", ["arabic"]],
  ["hindi", ["hindi"]],
  ["chinese", ["chinese", "mandarin"]],
  ["japanese", ["japanese"]],
  ["korean", ["korean"]],
  ["hebrew", ["hebrew"]],
  ["swedish", ["swedish"]],
  ["danish", ["danish"]],
  ["norwegian", ["norwegian"]],
  ["finnish", ["finnish"]]
];
const FIELD_CONTAINER_SELECTOR = [
  "fieldset",
  "[data-ui*='form-field' i]",
  "[data-testid*='form-field' i]",
  "[class*='form-field' i]",
  "[class*='question' i]",
  ".form-group",
  ".field",
  ".input",
  "li"
].join(",");
const FIELD_LABEL_SELECTOR = [
  "label",
  "legend",
  "[data-ui*='label' i]",
  "[data-testid*='label' i]",
  "[class*='label' i]"
].join(",");
let autoBidRunning = false;
let lastMousePoint = null;
let lastNativeClickError = null;
let autoBidTrace = [];
let autoBidRunId = "";
let autoBidRunStartedAt = "";
let activeAutoBidProfileId = "";
let activeAutoBidProfileEmail = "";
let tracePublishTimer = null;
const runtimeGptSourceFields = new Map();
const runtimeGptAnswersByRequest = new Map();
const runtimeGptApplyPromises = new Map();
const finalizedRuntimeGptRequestIds = new Set();
const generatedAnswerFillAttempts = new Map();
const completedOutlookVerificationMessages = new Set();
const atsAdapters = window.AutoBidAtsAdapters?.create({
  queryAll,
  isVisible,
  cleanText: cleanLabel
}) || null;

function isActiveContentInstance() {
  return window.__autoBidActiveContentInstance === AUTO_BID_INSTANCE_ID;
}

window.addEventListener("mousemove", (event) => {
  if (!isActiveContentInstance()) return;
  lastMousePoint = { x: event.clientX, y: event.clientY };
}, true);

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!isActiveContentInstance()) return false;
  if (message?.type === "AUTO_BID_TRIGGER") {
    runAutoBid();
    sendResponse?.({ ok: true });
    return false;
  }
  if (message?.type === "AUTO_BID_GPT_ANSWERS_READY") {
    const payload = message.payload || {};
    applyPushedRuntimeGptAnswers(payload)
      .then((result) => sendResponse?.({ ok: true, ...result }))
      .catch((error) => {
        traceAutoBid("runtime-gpt:push-apply-error", {
          request_id: payload.request_id || payload.requestId || "",
          message: error.message || String(error)
        });
        sendResponse?.({ ok: false, settled: false, error: error.message || String(error) });
      });
    return true;
  }
  if (message?.type === "AUTO_BID_OUTLOOK_VERIFICATION_READY") {
    const payload = message.payload || {};
    completeOutlookVerification(payload)
      .then((result) => sendResponse?.({ ok: true, ...result }))
      .catch((error) => {
        traceAutoBid("outlook-verification:completion-error", {
          monitor_id: payload.monitor_id || "",
          message_id: payload.message?.id || "",
          message: error.message || String(error)
        });
        sendResponse?.({ ok: false, applied: false, error: error.message || String(error) });
      });
    return true;
  }
  return false;
});

runAutoBid();

async function runAutoBid() {
  if (!isActiveContentInstance() || autoBidRunning) return;
  autoBidRunning = true;
  lastNativeClickError = null;
  resetTrace();
  send("AUTOBID_AUTOFILL_STATE", { running: true, run_id: autoBidRunId, url: location.href }).catch(() => {});
  traceAutoBid("run:start", {
    url: location.href,
    title: document.title,
    frame: getFrameScope(),
    ats: atsAdapters?.describe?.() || { id: "common", name: "Common form" }
  });
  showStatus("Scanning required fields", "working", { detail: "Preparing this application for autofill…" });
  await closeOpenChoiceMenus();

  let fields = collectFields();
  traceAutoBid("fields:collected", {
    count: fields.length,
    fields: fields.map((field) => ({
      id: field.id,
      question: field.question || "",
      option: field.option || "",
      label: field.label,
      type: field.type,
      required: field.required,
      value: field.value,
      options: field.options
    }))
  });

  if (fields.length === 0) {
    const embeddedFrames = isTopFrame() ? queryAll("iframe").length : 0;
    traceAutoBid("run:no-fields", { frame: getFrameScope(), embedded_frames: embeddedFrames });
    flushTrace();
    showStatus(embeddedFrames ? "Checking embedded application frames..." : "No fillable fields found.", embeddedFrames ? "success" : "error");
    autoBidRunning = false;
    await send("AUTOBID_AUTOFILL_STATE", { running: false, run_id: autoBidRunId, reason: "no-fields", url: location.href }).catch(() => {});
    return;
  }

  try {
    const initiallyFilledIds = getCurrentlyFilledFieldIds(fields);
    traceAutoBid("fields:already-filled", { count: initiallyFilledIds.size, field_ids: Array.from(initiallyFilledIds) });

    showStatus("Autofilling saved information", "autofilling", {
      detail: "Filling known profile values before requesting generated answers…"
    });
    const staticFallbackResult = await runStep("static-profile", () => applyProfileStaticFallbacks(fields, initiallyFilledIds), emptyFillResult());
    const locallyFilledIds = new Set([...initiallyFilledIds, ...staticFallbackResult.filledIds]);
    const localFallbackResult = await runStep("local-fallbacks", () => applyLocalGeneratedFallbacks(fields, locallyFilledIds), emptyFillResult());
    fields = collectFields();
    const localAnswerIds = new Set([
      ...getCurrentlyFilledFieldIds(fields),
      ...localFallbackResult.filledIds
    ]);
    traceAutoBid("fields:refreshed-after-local", {
      count: fields.length,
      filled: localAnswerIds.size
    });
    const resumeResultPromise = runStepWithTimeout("resume-upload", attachResumeFromSheet, emptyResumeResult(), RESUME_ATTACH_TIMEOUT_MS);

    traceAutoBid("ai:router-order", {
      first: "chatgpt-extension",
      second: null,
      openai_enabled: OPENAI_AUTOFILL_ROUTE_ENABLED,
      candidates: getGeneratedAnswerCandidateFields(fields, localAnswerIds).length
    });
    const runtimeGptResult = await runStep(
      "chatgpt-first-provider",
      () => applyRuntimeGptAnswerExchange(fields, localAnswerIds),
      emptySheetResult()
    );
    fields = collectFields();
    const afterRuntimeGptIds = getCurrentlyFilledFieldIds(fields);
    traceAutoBid("fields:refreshed-after-chatgpt", {
      count: fields.length,
      filled: afterRuntimeGptIds.size,
      chatgpt_filled: runtimeGptResult.filled,
      chatgpt_pending: runtimeGptResult.pending
    });
    const openAiResult = OPENAI_AUTOFILL_ROUTE_ENABLED
      ? await runStep(
        "openai-second-provider",
        () => applyDirectAiAnswers(fields, afterRuntimeGptIds),
        emptyFillResult()
      )
      : emptyFillResult();
    if (!OPENAI_AUTOFILL_ROUTE_ENABLED) {
      traceAutoBid("ai:openai-route-disabled", {
        unresolved_required: getMissingRequiredFields(fields).length
      });
    }
    fields = collectFields();
    traceAutoBid("fields:refreshed-after-openai", {
      count: fields.length,
      filled: getCurrentlyFilledFieldIds(fields).size,
      openai_filled: openAiResult.filled,
      unresolved_required: openAiResult.missed
    });
    const resumeResult = await resumeResultPromise;
    fields = collectFields();
    const postGeneratedFilledIds = getCurrentlyFilledFieldIds(fields);
    const profileReconcileResult = await runStep(
      "profile-reconcile",
      () => applyProfileStaticFallbacks(fields, postGeneratedFilledIds),
      emptyFillResult()
    );
    const runtimeGptReconcileResult = emptyFillResult();
    fields = collectFields();
    const finalDefaultIds = getCurrentlyFilledFieldIds(fields);
    const finalDefaultResult = await runStep(
      "final-deterministic-defaults",
      () => applyDeterministicDefaults(
        fields.filter((field) => !shouldDeferChoiceFieldToRuntimeGpt(field)),
        finalDefaultIds
      ),
      emptyFillResult()
    );
    const submitResult = await maybeAutoSubmitApplication(fields, {
      filled: runtimeGptResult.filled + openAiResult.filled + finalDefaultResult.filled,
      pending: runtimeGptResult.pending
    }, resumeResult);
    const filled = staticFallbackResult.filled + localFallbackResult.filled + runtimeGptResult.filled + openAiResult.filled + profileReconcileResult.filled + runtimeGptReconcileResult.filled + finalDefaultResult.filled + resumeResult.filled;
    const missed = getMissingRequiredFields(collectFields()).length;
    traceAutoBid("run:complete", { filled, missed, runtime_gpt: runtimeGptResult, openai: openAiResult, profile_reconcile: profileReconcileResult, runtime_gpt_reconcile: runtimeGptReconcileResult, final_defaults: finalDefaultResult, resume: resumeResult, submit: submitResult });
    showStatus(
      submitResult.clicked ? "Application submitted" : missed > 0 ? "Autofill finished" : "Autofill done",
      submitResult.clicked || missed === 0 ? "success" : "warning",
      {
        detail: submitResult.clicked
          ? "Required fields were completed and the application was submitted."
          : missed > 0
            ? `${missed} required field${missed === 1 ? " is" : "s are"} still empty. Press Ctrl+Q to retry only those fields.`
            : "All detected required fields are complete."
      }
    );
  } catch (error) {
    traceAutoBid("run:error", { message: error.message || String(error) });
    showStatus("Autofill stopped", "error", { detail: error.message || String(error) });
  } finally {
    flushTrace();
    autoBidRunning = false;
    await send("AUTOBID_AUTOFILL_STATE", { running: false, run_id: autoBidRunId, reason: "run-finished", url: location.href }).catch(() => {});
  }
}

async function runStep(name, runner, fallback) {
  try {
    return await runner();
  } catch (error) {
    traceAutoBid(`${name}:error`, { message: error.message || String(error) });
    return fallback;
  }
}

async function runStepWithTimeout(name, runner, fallback, timeoutMs) {
  try {
    return await promiseWithTimeout(runner(), timeoutMs, `${name} timed out after ${timeoutMs}ms`);
  } catch (error) {
    traceAutoBid(`${name}:error`, { message: error.message || String(error), timeout_ms: timeoutMs });
    return fallback;
  }
}

function promiseWithTimeout(promise, timeoutMs, message) {
  let timeoutId = null;
  return Promise.race([
    Promise.resolve(promise),
    new Promise((_resolve, reject) => {
      timeoutId = setTimeout(() => reject(new Error(message)), timeoutMs);
    })
  ]).finally(() => {
    if (timeoutId) clearTimeout(timeoutId);
  });
}

function emptyFillResult() {
  return { filled: 0, missed: 0, filledIds: new Set(), missedIds: new Set() };
}

function emptySheetResult() {
  return { ...emptyFillResult(), pending: 0 };
}

function emptyResumeResult() {
  return { filled: 0, missed: 0, reason: "skipped-after-error" };
}

async function applyLocalGeneratedFallbacks(fields, filledIds) {
  const activeFilledIds = new Set(filledIds);
  const localFilledIds = new Set();
  let filled = 0;
  let missed = 0;

  const runLocalPass = async (name, runner) => {
    let result;
    try {
      result = await runner(fields, activeFilledIds);
    } catch (error) {
      traceAutoBid(`local:${name}:error`, { message: error.message || String(error) });
      return;
    }
    result.filledIds?.forEach((fieldId) => {
      activeFilledIds.add(fieldId);
      localFilledIds.add(fieldId);
    });
    filled += result.filled || 0;
    missed += result.missed || 0;
    traceAutoBid(`local:${name}`, {
      filled: result.filled || 0,
      missed: result.missed || 0,
      field_ids: Array.from(result.filledIds || [])
    });
  };

  await runLocalPass("language-choice", applyLanguageChoiceAnswers);
  await runLocalPass("outlook-verification", applyOutlookVerificationAnswers);
  await runLocalPass("based-in-location", (items, ids) => applyBasedInLocationAnswers(items, [], ids));
  await runLocalPass("sensitive-demographic-decline", applySensitiveDemographicDeclineAnswers);
  await runLocalPass("referral-source", applyReferralSourceAnswers);
  await runLocalPass("consent-choice", applyConsentChoiceAnswers);
  await runLocalPass("explicit-not-applicable", applyExplicitNotApplicableAnswers);
  await runLocalPass("deterministic-defaults", (items, ids) => applyDeterministicDefaults(
    items.filter((field) => !shouldDeferChoiceFieldToRuntimeGpt(field)),
    ids
  ));
  await runLocalPass("positive-dropdowns", applyPositiveDropdownFallbacks);
  await runLocalPass("positive-checkboxes", applyPositiveCheckboxFallbacks);

  return { filled, missed, filledIds: localFilledIds };
}

async function applyExplicitNotApplicableAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !field.required || !["text", "search", "textarea"].includes(field.type)) continue;
    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasFieldCurrentValue(field, controls)) continue;

    const instructions = normalize([
      field.question,
      field.label,
      field.placeholder,
      getDescribedByText(controls[0]),
      getNearbyText(controls[0])
    ].filter(Boolean).join(" "));
    if (!explicitlyAllowsNotApplicable(instructions)) continue;

    const answer = /\banswer n a\b|\brespond n a\b|\benter n a\b|\bwrite n a\b/.test(instructions)
      ? "N/A"
      : "Not applicable";
    const selected = await setControlsValue(controls, answer, field);
    traceAutoBid("explicit-not-applicable:result", {
      field_id: field.id,
      label: field.question || field.label,
      answer,
      selected,
      current: getCurrentChoiceSummary(controls)
    });
    if (selected) {
      filled += 1;
      filledLocalIds.add(field.id);
    } else {
      missed += 1;
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

function explicitlyAllowsNotApplicable(text) {
  const value = normalize(text);
  if (!value) return false;
  const sentinel = /\b(n a|not applicable)\b/;
  const conditional = /\b(if not|if no|if none|if you (?:were|are|have|do|did) not|if you (?:haven t|don t|didn t)|if you answered no|otherwise)\b/;
  const instruction = /\b(answer|respond|enter|write|put|use|please answer|please enter)\b/;
  return sentinel.test(value) && conditional.test(value) && instruction.test(value);
}

async function applyOutlookVerificationAnswers(fields, filledIds) {
  const candidates = fields.filter((field) =>
    field.required &&
    !filledIds.has(field.id) &&
    !hasFieldCurrentValue(field) &&
    isOutlookVerificationCodeField(field)
  );
  const result = { filled: 0, missed: 0, filledIds: new Set(), missedIds: new Set() };
  if (candidates.length === 0) return result;

  let verification;
  try {
    if (!activeAutoBidProfileId) {
      const profileStatic = await send("GET_PROFILE_STATIC_FIELDS");
      captureAutoBidProfileContext(profileStatic);
    }
    verification = await send("OUTLOOK_FIND_VERIFICATION", {
      domain: location.hostname.replace(/^www\./, ""),
      page_url: location.href,
      title: document.title || "",
      since: new Date(Date.now() - OUTLOOK_VERIFICATION_LOOKBACK_MS).toISOString(),
      top: 10,
      profile_id: activeAutoBidProfileId,
      mailbox_email: activeAutoBidProfileEmail
    });
  } catch (error) {
    traceAutoBid("outlook-verification:unavailable", { message: error.message || String(error) });
    return result;
  }
  if (!verification?.code) {
    traceAutoBid("outlook-verification:no-code", {
      reason: verification?.reason || "not-found",
      candidates: candidates.map((field) => field.id)
    });
    return result;
  }

  for (const field of candidates) {
    const controls = getControlsByFieldId(field.id);
    const applied = controls.length > 0 && await setControlsValue(controls, verification.code, field);
    if (applied) {
      result.filled += 1;
      result.filledIds.add(field.id);
      traceAutoBid("outlook-verification:applied", {
        field_id: field.id,
        message_id: verification.message?.id || ""
      });
    } else {
      result.missed += 1;
      result.missedIds.add(field.id);
    }
  }
  if (result.filled > 0 && verification.message?.id) {
    send("OUTLOOK_MARK_READ", {
      messageId: verification.message.id,
      connection_id: verification.message.connection_id || ""
    }).catch(() => {});
  }
  return result;
}

function isOutlookVerificationCodeField(field) {
  const text = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (/(postal|zip|country|dial|phone|promo|coupon|referral)/.test(text)) return false;
  return /(verification|security|confirmation|one time|otp|passcode|authentication)\s*(code|pin|number)?/.test(text) ||
    /\b(code|pin)\b/.test(text) && /(email|account|application|verify|confirm)/.test(text);
}

async function completeOutlookVerification(payload = {}) {
  const code = String(payload.code || "").replace(/[^a-z0-9]/gi, "");
  const messageId = String(payload.message?.id || payload.message_id || "");
  const monitorId = String(payload.monitor_id || "");
  const completionKeys = [messageId, monitorId].filter(Boolean);
  if (!/^\w{4,10}$/i.test(code) || !/\d/.test(code)) {
    return { applied: false, clicked: false, reason: "invalid-verification-code" };
  }
  if (completionKeys.some((key) => completedOutlookVerificationMessages.has(key))) {
    return { applied: true, clicked: false, settled: true, reason: "already-completed" };
  }

  showStatus("Verification code received", "autofilling", {
    detail: "Filling the email verification step and confirming it once…"
  });
  const candidates = await waitForOutlookVerificationCodeFields();
  if (candidates.length === 0) {
    traceAutoBid("outlook-verification:field-not-ready", {
      monitor_id: payload.monitor_id || "",
      message_id: messageId
    });
    return { applied: false, clicked: false, reason: "verification-field-not-ready" };
  }

  const controls = uniqueElements(candidates.flatMap((field) => getControlsByFieldId(field.id)));
  const alreadyApplied = controls.some((control) => normalizeComparableValue(control.value) === normalizeComparableValue(code));
  const hasDifferentManualValue = !alreadyApplied && controls.some((control) => String(control.value || "").trim());
  if (hasDifferentManualValue) {
    traceAutoBid("outlook-verification:manual-value-preserved", {
      monitor_id: payload.monitor_id || "",
      message_id: messageId
    });
    return { applied: false, clicked: false, settled: true, reason: "manual-value-preserved" };
  }
  const applied = alreadyApplied || await applyOutlookVerificationCode(candidates, controls, code);
  if (!applied) {
    traceAutoBid("outlook-verification:code-fill-failed", {
      monitor_id: payload.monitor_id || "",
      message_id: messageId,
      fields: candidates.map((field) => field.id)
    });
    return { applied: false, clicked: false, reason: "verification-code-fill-failed" };
  }

  completionKeys.forEach((key) => completedOutlookVerificationMessages.add(key));
  const button = await waitForOutlookVerificationSubmitButton(controls);
  let clicked = false;
  if (button && !isDisabledSubmitButton(button)) {
    await scrollElementIntoView(button, "center");
    await sleep(150);
    clicked = await nativeClickElement(button);
    if (!clicked) {
      dispatchRealisticMouseClick(button);
      clicked = true;
    }
  }

  traceAutoBid("outlook-verification:completed", {
    monitor_id: payload.monitor_id || "",
    message_id: messageId,
    applied,
    clicked,
    button: button ? getSubmitButtonText(button) : ""
  });
  showStatus(clicked ? "Email verification submitted" : "Verification code filled", clicked ? "success" : "warning", {
    detail: clicked
      ? "The code was filled and the verification button was clicked once."
      : "The code was filled once, but no safe verification button was found."
  });
  return { applied: true, clicked, settled: true, reason: clicked ? "verification-submitted" : "button-not-found" };
}

async function waitForOutlookVerificationCodeFields(timeoutMs = OUTLOOK_VERIFICATION_FIELD_WAIT_MS) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const candidates = collectFields().filter((field) =>
      isOutlookVerificationCodeField(field) &&
      getControlsByFieldId(field.id).some((control) => isVisible(control))
    );
    if (candidates.length > 0) return candidates;
    await sleep(300);
  }
  return [];
}

async function applyOutlookVerificationCode(fields, controls, code) {
  const textControls = controls.filter((control) => {
    const type = getControlType(control);
    return ["text", "number", "tel"].includes(type) || control.getAttribute?.("inputmode") === "numeric";
  });
  const segmented = textControls.length >= code.length &&
    textControls.slice(0, code.length).every((control) => Number(control.maxLength || control.getAttribute?.("maxlength") || 0) === 1);
  if (segmented) {
    for (let index = 0; index < code.length; index += 1) {
      const control = textControls[index];
      control.focus?.({ preventScroll: true });
      setNativeValue(control, code[index]);
      dispatchInput(control);
    }
    textControls[Math.min(code.length, textControls.length) - 1]?.blur?.();
    await sleep(200);
    return textControls.slice(0, code.length).map((control) => String(control.value || "")).join("") === code;
  }

  for (const field of fields) {
    const fieldControls = getControlsByFieldId(field.id);
    if (fieldControls.length === 0) continue;
    if (await setControlsValue(fieldControls, code, field)) return true;
  }
  return false;
}

async function waitForOutlookVerificationSubmitButton(controls, timeoutMs = OUTLOOK_VERIFICATION_BUTTON_WAIT_MS) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const button = findOutlookVerificationSubmitButton(controls);
    if (button && !isDisabledSubmitButton(button)) return button;
    await sleep(300);
  }
  return findOutlookVerificationSubmitButton(controls);
}

function findOutlookVerificationSubmitButton(controls) {
  const roots = uniqueElements(controls.flatMap((control) => [
    control.closest?.("form"),
    control.closest?.("[role='dialog'], main, section, article, [class*='verification' i], [class*='confirm' i]"),
    document.body
  ].filter(Boolean)));
  const candidates = uniqueElements(roots.flatMap((root) =>
    queryAll("button, input[type='submit'], input[type='button'], [role='button']", root)
  )).filter((element) => isVisible(element) && !isDisabledSubmitButton(element))
    .map((element) => ({ element, text: normalize(getSubmitButtonText(element)) }))
    .filter(({ text }) => text && !/(resend|send again|new code|change email|cancel|back|previous)/.test(text))
    .map((candidate) => ({ ...candidate, score: scoreOutlookVerificationButton(candidate.text) }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score);
  return candidates[0]?.element || null;
}

function scoreOutlookVerificationButton(text) {
  if (/^(verify|verify code|confirm|confirm code|submit code)$/.test(text)) return 140;
  if (/(verify|verification|confirm).*(email|code|application)|(email|code|application).*(verify|confirm)/.test(text)) return 130;
  if (/(complete|finish).*(application|verification)/.test(text)) return 120;
  if (/^(submit|continue|send)$/.test(text)) return 90;
  if (/(submit|continue|send).*(application|verification|code)/.test(text)) return 80;
  return 0;
}

function uniqueElements(values) {
  return Array.from(new Set((values || []).filter(Boolean)));
}

async function applyDirectAiAnswers(fields, filledIds) {
  const localFilledIds = new Set();
  const candidateFields = getGeneratedAnswerCandidateFields(fields, filledIds);
  if (candidateFields.length === 0) {
    traceAutoBid("ai:no-candidates", {
      fields: fields.map((field) => ({
        field_id: field.id,
        label: field.label,
        type: field.type,
        required: field.required,
        reason: getGeneratedAnswerCandidateSkipReason(field, filledIds)
      }))
    });
    return { filled: 0, missed: 0, filledIds: localFilledIds };
  }

  await hydrateGeneratedChoiceOptions(candidateFields);

  traceAutoBid("ai:candidates", {
    count: candidateFields.length,
    fields: candidateFields.map((field) => ({
      field_id: field.id,
      label: field.label,
      type: field.type,
      options: field.options || []
    }))
  });

  let data;
  try {
    data = await send("ASSIST", {
      page: collectPageContext(),
      fields: candidateFields
    });
  } catch (error) {
    traceAutoBid("ai:error", { message: error.message || String(error) });
    return {
      filled: 0,
      missed: candidateFields.length,
      filledIds: localFilledIds,
      missedIds: new Set(candidateFields.map((field) => field.id))
    };
  }

  const candidateIds = new Set(candidateFields.map((field) => field.id));
  const answers = normalizeDirectAiAnswers(data)
    .filter((answer) => candidateIds.has(answer.field_id))
    .filter((answer) => !hasFieldCurrentValue(fields.find((field) => field.id === answer.field_id)));

  traceAutoBid("ai:answers-received", {
    count: answers.length,
    warnings: Array.isArray(data?.warnings) ? data.warnings : [],
    answers: answers.map((answer) => ({
      field_id: answer.field_id,
      value: shortText(answer.value),
      provider: answer.provider || "",
      model: answer.model || "",
      estimated_request_cost_usd: answer.estimated_request_cost_usd ?? null
    }))
  });

  if (answers.length === 0) {
    return {
      filled: 0,
      missed: candidateFields.length,
      filledIds: localFilledIds,
      missedIds: new Set(candidateFields.map((field) => field.id))
    };
  }

  const result = await applyAnswers(answers, filledIds, fields);
  result.filledIds.forEach((fieldId) => localFilledIds.add(fieldId));
  const unresolvedFields = candidateFields.filter((field) => !hasFieldCurrentValue(field));
  return {
    filled: result.filled,
    missed: unresolvedFields.length,
    filledIds: localFilledIds,
    missedIds: new Set(unresolvedFields.map((field) => field.id))
  };
}

function collectPageContext() {
  return {
    url: location.href,
    domain: location.hostname.replace(/^www\./, ""),
    title: document.title || "",
    job_title: findJobTitle(),
    ats: atsAdapters?.describe?.() || { id: "common", name: "Common form" },
    text: (document.body?.innerText || "").replace(/\s+/g, " ").trim().slice(0, 18000)
  };
}

function findJobTitle() {
  const heading = document.querySelector("h1") || document.querySelector("[data-testid*='title' i]");
  return heading?.textContent?.trim() || document.title || "";
}

function collectFields() {
  const controls = getFormControls().filter(isFillableControl);
  const fields = [];
  const radioGroups = new Map();
  const checkboxGroups = new Map();

  controls.forEach((control) => {
    if (getControlType(control) !== "checkbox") return;
    const name = cleanLabel(control.name || control.getAttribute?.("name") || "");
    if (!name) return;
    if (!checkboxGroups.has(name)) checkboxGroups.set(name, []);
    checkboxGroups.get(name).push(control);
  });

  controls.forEach((control, index) => {
    const type = getControlType(control);
    if (type === "radio") {
      const key = getRadioGroupKey(control, index);
      if (!radioGroups.has(key)) radioGroups.set(key, []);
      radioGroups.get(key).push(control);
      return;
    }

    const label = getFieldLabel(control);
    const id = `ab_${index}_${hashSmall([
      label,
      control.name,
      control.id,
      control.getAttribute("placeholder"),
      control.getAttribute("autocomplete")
    ].filter(Boolean).join(" "))}`;
    const checkboxName = cleanLabel(control.name || control.getAttribute?.("name") || "");
    const choiceQuestionLabel = type === "checkbox"
      ? getCheckboxGroupQuestionLabel(control, checkboxGroups.get(checkboxName) || [control])
      : "";
    const question = getPlainQuestionText(choiceQuestionLabel || label);
    const option = type === "checkbox" ? getCheckboxOptionLabel(control, question) : "";
    control.dataset.autoBidFieldId = id;
    fields.push({
      id,
      label,
      question,
      option,
      name: control.name || "",
      placeholder: control.getAttribute("placeholder") || "",
      autocomplete: control.getAttribute("autocomplete") || "",
      type,
      required: isRequired(control) || Boolean(choiceQuestionLabel && /\*/.test(choiceQuestionLabel)),
      options: getControlOptions(control),
      value: getControlValue(control)
    });
  });

  Array.from(radioGroups.values()).forEach((group, groupIndex) => {
    const first = group[0];
    const label = getRadioGroupLabel(group);
    const id = `ab_radio_${groupIndex}_${hashSmall([label, first.name || "", group.map(getRadioOptionLabel).join(" ")].join(" "))}`;
    group.forEach((control) => {
      control.dataset.autoBidFieldId = id;
    });
    fields.push({
      id,
      label,
      question: getPlainQuestionText(label),
      option: "",
      name: first.name || "",
      placeholder: "",
      autocomplete: "",
      type: "radio",
      required: group.some(isRequired),
      options: group.map(getRadioOptionLabel).filter(Boolean),
      value: getRadioGroupValue(group)
    });
  });

  fields.push(...collectButtonChoiceFields(fields.length));
  fields.push(...collectConsentCheckboxFields(fields.length, fields));

  return fields;
}

function getFormControls() {
  return queryAll(Array.from(new Set([
    "input",
    "textarea",
    "select",
    "[role='checkbox']",
    "[role='radio']",
    "button[role='combobox']",
    "[role='combobox']",
    "[aria-haspopup='listbox']",
    "[aria-haspopup='menu'][aria-expanded]",
    "[data-radix-select-trigger]",
    "[data-slot='select-trigger']",
    ".select__control",
    ".fab-SelectToggle",
    "[data-fabric-component='SelectToggle']",
    ".select2-selection",
    ".select2-choice",
    ".chosen-single",
    "[class*='select'][aria-expanded]",
    ...(atsAdapters?.getControlSelectors?.() || [])
  ])).join(","));
}

function getControlsByFieldId(fieldId) {
  if (!fieldId) return [];
  return queryAll(`[data-auto-bid-field-id="${cssEscape(fieldId)}"]`);
}

function queryOne(selector, root = document) {
  return queryAll(selector, root)[0] || null;
}

function queryAll(selector, root = document) {
  const roots = getQueryableRoots(root);
  const results = [];

  for (const queryRoot of roots) {
    if (queryRoot.matches?.(selector)) results.push(queryRoot);
    if (queryRoot.querySelectorAll) {
      results.push(...Array.from(queryRoot.querySelectorAll(selector)));
    }
  }

  return Array.from(new Set(results));
}

function getQueryableRoots(root = document) {
  const roots = [];
  const stack = [root || document];

  for (let index = 0; index < stack.length; index += 1) {
    const current = stack[index];
    if (!current || roots.includes(current)) continue;
    roots.push(current);

    if (!current.querySelectorAll) continue;
    for (const element of Array.from(current.querySelectorAll("*"))) {
      if (element.shadowRoot) stack.push(element.shadowRoot);
    }
  }

  return roots;
}

function collectButtonChoiceFields(startIndex) {
  const groups = [];
  const groupRoots = Array.from(new Set(
    queryAll("button, [role='button'], [role='radio']")
      .filter(isChoiceButton)
      .map((button) => button.parentElement)
      .filter(Boolean)
  ));

  groupRoots.forEach((root, groupIndex) => {
    const buttons = Array.from(root.children).filter(isChoiceButton);
    if (!isLikelyButtonChoiceGroup(buttons)) return;
    if (buttons.some((button) => button.dataset.autoBidFieldId)) return;

    const first = buttons[0];
    const options = buttons.map(getChoiceButtonLabel).filter(Boolean);
    const label = getChoiceGroupLabel(buttons);
    if (!label || normalize(label) === normalize(options.join(" "))) return;

    const id = `ab_button_${startIndex + groupIndex}_${hashSmall([label, options.join(" ")].join(" "))}`;
    buttons.forEach((button) => {
      button.dataset.autoBidFieldId = id;
      button.dataset.autoBidControlType = "button-group";
    });

    groups.push({
      id,
      label,
      question: getPlainQuestionText(label),
      option: "",
      name: root.getAttribute("name") || first.name || root.id || first.id || "",
      placeholder: "",
      autocomplete: "",
      type: "button-group",
      required: buttons.some(isRequired),
      options,
      value: getSelectedChoiceButtonLabel(buttons)
    });
  });

  return groups;
}

function collectConsentCheckboxFields(startIndex, existingFields = []) {
  const fields = [];
  const seenControls = new Set();
  const roots = queryAll([
    "input[type='checkbox']",
    "label",
    "[role='checkbox']",
    "[class*='checkbox' i]",
    "[class*='terms' i]",
    "[class*='privacy' i]",
    "[class*='consent' i]"
  ].join(","));

  roots.forEach((root, index) => {
    if (!isVisible(root)) return;
    const control = root.matches?.("input[type='checkbox'], [role='checkbox']")
      ? root
      : root.querySelector?.("input[type='checkbox'], [role='checkbox']");
    if (!control) return;

    const label = getConsentCheckboxLabel(root, control);
    if (!label || !isConsentCheckboxText(label)) return;

    if (seenControls.has(control)) return;
    if (control.matches?.("input") && control.disabled) return;

    const existingId = control.dataset?.autoBidFieldId || "";
    const existingField = existingFields.find((field) => field.id === existingId);
    if (existingField) {
      existingField.label = label;
      existingField.question = getPlainQuestionText(label);
      existingField.option = "";
      existingField.required = existingField.required || isRequired(control) || /\*/.test(label);
      existingField.options = ["Yes", "No"];
      existingField.value = isCheckboxChecked(control) ? "Yes" : "";
      seenControls.add(control);
      return;
    }

    const id = `ab_checkbox_${startIndex + index}_${hashSmall(label)}`;
    control.dataset.autoBidFieldId = id;
    if (!control.matches?.("input[type='checkbox']")) control.dataset.autoBidControlType = "checkbox";
    if (root !== control) root.dataset.autoBidFieldId = id;
    seenControls.add(control);

    fields.push({
      id,
      label,
      question: getPlainQuestionText(label),
      option: "",
      name: control.name || root.getAttribute("name") || root.id || control.id || "",
      placeholder: "",
      autocomplete: "",
      type: "checkbox",
      required: isRequired(control) || /\*/.test(label),
      options: ["Yes", "No"],
      value: isCheckboxChecked(control) ? "Yes" : ""
    });
  });

  return fields;
}

function isConsentCheckboxText(text) {
  const label = normalize(text);
  if (isFileUploadChoiceText(label)) return false;
  if (isSensitiveOrPersonalChoiceCheckbox(label)) return false;
  return /(terms|privacy|policy|consent|agree|accept|acknowledge|confirm|certify|accurate|process.*data|data.*process)/.test(label);
}

function isFileUploadChoiceText(label) {
  const text = normalize(label);
  return /(accepted file types|file types|resume|cv|cover letter|attach|upload|dropbox|google drive|enter manually|pdf|docx|rtf)/.test(text);
}

function getConsentCheckboxLabel(root, control) {
  const input = getCheckboxInput(control) || control;
  const rootNode = input.getRootNode?.() || document;
  const forLabel = input.id ? queryOne(`label[for="${cssEscape(input.id)}"]`, rootNode) : null;
  const row = getConsentCheckboxRow(input, root);
  return cleanLabel([
    root.textContent,
    input.getAttribute?.("aria-label"),
    input.getAttribute?.("title"),
    forLabel?.textContent,
    input.closest?.("label")?.textContent,
    row?.textContent,
    ...getNearbyConsentTextCandidates(input)
  ].filter(Boolean).join(" "));
}

function getConsentCheckboxRow(input, root = null) {
  const candidates = [
    root,
    input.closest?.("label"),
    input.closest?.("[class*='checkbox' i], [class*='consent' i], [class*='privacy' i], [class*='terms' i], li, p"),
    input.parentElement,
    input.parentElement?.parentElement,
    getFieldContainer(input)
  ].filter(Boolean);

  return candidates
    .filter((candidate) => isVisible(candidate))
    .map((candidate) => ({
      element: candidate,
      text: cleanLabel(candidate.textContent || ""),
      rect: candidate.getBoundingClientRect()
    }))
    .filter((candidate) => candidate.text && candidate.text.length <= 900)
    .sort((left, right) => {
      const leftConsent = isConsentCheckboxText(left.text) ? 0 : 1;
      const rightConsent = isConsentCheckboxText(right.text) ? 0 : 1;
      return leftConsent - rightConsent || left.rect.height - right.rect.height;
    })[0]?.element || input.parentElement;
}

function getNearbyConsentTextCandidates(input) {
  const candidates = [];
  const baseRect = input.getBoundingClientRect();
  const scope = input.closest?.("form, main, [role='main'], section") || document.body;
  const selector = "label, span, div, p";

  queryAll(selector, scope).forEach((element) => {
    if (!isVisible(element) || element === input || element.contains(input)) return;
    if (element.querySelector?.("input, textarea, select, button, [role='button'], [role='combobox']")) return;

    const text = cleanLabel(element.textContent || "");
    if (!text || text.length > 900 || !isConsentCheckboxText(text)) return;

    const rect = element.getBoundingClientRect();
    const verticalOverlap = Math.max(0, Math.min(rect.bottom, baseRect.bottom + 28) - Math.max(rect.top, baseRect.top - 28));
    const horizontalDistance = Math.max(0, rect.left - baseRect.right, baseRect.left - rect.right);
    if (verticalOverlap <= 0 && Math.abs(rect.top - baseRect.top) > 60) return;
    if (horizontalDistance > 120) return;
    candidates.push(text);
  });

  return candidates.slice(0, 4);
}

function isChoiceButton(button) {
  if (!button || !isVisible(button) || button.disabled) return false;
  if (isLikelyDropdownTrigger(button)) return false;
  if (button.closest("[role='listbox'], [role='menu'], .select__menu")) return false;
  const type = String(button.getAttribute("type") || "button").toLowerCase();
  if (["submit", "reset"].includes(type)) return false;
  const label = normalize(getChoiceButtonLabel(button));
  if (!label || /^(clear|remove|delete|attach|upload|browse|submit|apply|next|back|continue|cancel)$/.test(label)) return false;
  return true;
}

function isLikelyButtonChoiceGroup(buttons) {
  if (buttons.length < 2 || buttons.length > 4) return false;
  const labels = buttons.map((button) => normalize(getChoiceButtonLabel(button)));
  return labels.includes("yes") && labels.includes("no");
}

function isFillableControl(control) {
  const type = getControlType(control);
  if (control.getAttribute("role") === "radio" && control.querySelector?.("input[type='radio']")) return false;
  const visibleChoiceControl = isVisibleChoiceControl(control, type);
  if ((control.getAttribute("aria-hidden") === "true" && !visibleChoiceControl) || isCompositeComboboxShell(control)) return false;
  if (!isVisible(control) && !visibleChoiceControl) return false;
  if (control.disabled || (control.readOnly && type !== "combobox")) return false;
  if (isPhoneDialCodeSelector(control, type)) return false;
  return !["hidden", "submit", "button", "reset", "image", "file", "password"].includes(type);
}

function isPhoneDialCodeSelector(control, type = getControlType(control)) {
  if (!["combobox", "select"].includes(type)) return false;

  const ownText = cleanLabel([
    control.value,
    control.textContent,
    getComboboxSelectedText(control),
    control.getAttribute?.("aria-label"),
    control.getAttribute?.("title"),
    control.name,
    control.id
  ].filter(Boolean).join(" "));
  const context = normalize([
    ownText,
    control.getAttribute?.("aria-label"),
    control.getAttribute?.("title"),
    control.name,
    control.id,
    getNearbyText(control)
  ].filter(Boolean).join(" "));

  const phoneContext = /\b(phone|mobile|telephone|cell)\b/.test(context);
  const dialCodeText = /\+\s*\d{1,4}/.test(ownText) ||
    /(country code|calling code|dial code|phone code|phone prefix|country prefix)/.test(context);

  return phoneContext && dialCodeText && hasNearbyPhoneNumberInput(control);
}

function hasNearbyPhoneNumberInput(control) {
  const roots = [
    control.parentElement,
    control.closest("label"),
    control.closest("[class*='phone' i], [class*='field' i], [class*='input' i], .form-group"),
    getFieldContainer(control)
  ].filter(Boolean);

  return roots.some((root) => Array.from(root.querySelectorAll("input"))
    .some((input) => input !== control && input !== getComboboxInput(control) && isPhoneNumberEntryInput(input)));
}

function isPhoneNumberEntryInput(input) {
  if (!input || input.disabled || input.readOnly || !isVisible(input)) return false;
  const type = String(input.getAttribute("type") || "text").toLowerCase();
  if (!["tel", "text", "search", "number"].includes(type)) return false;
  const text = normalize([
    input.getAttribute("autocomplete"),
    input.getAttribute("placeholder"),
    input.getAttribute("aria-label"),
    input.name,
    input.id
  ].join(" "));
  return type === "tel" || /\b(phone|mobile|telephone|cell|number)\b/.test(text);
}

function isVisible(element) {
  const rect = element.getBoundingClientRect();
  const style = window.getComputedStyle(element);
  return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
}

function isVisibleChoiceControl(control, type = getControlType(control)) {
  if (!["checkbox", "radio"].includes(type)) return false;
  const label = control.closest("label");
  const wrapper = control.closest(".application-question, [class*='question' i], [class*='field' i], [class*='checkbox' i], [class*='option' i], [role='group'], [role='radiogroup'], li");
  return Boolean((label && isVisible(label)) || (wrapper && isVisible(wrapper)));
}

function getControlType(control) {
  const adapterType = atsAdapters?.getControlType?.(control);
  if (adapterType) return adapterType;
  if (control.dataset?.autoBidControlType === "button-group") return "button-group";
  if (control.getAttribute("role") === "checkbox") return "checkbox";
  if (control.getAttribute("role") === "radio") return "radio";
  if (
    control.getAttribute("role") === "combobox" ||
    ["listbox", "menu"].includes(control.getAttribute("aria-haspopup") || "") ||
    control.hasAttribute("data-radix-select-trigger") ||
    control.getAttribute("data-slot") === "select-trigger" ||
    control.classList?.contains("select__control") ||
    control.matches?.(".fab-Select__control, [class*='Select__control'], .fab-SelectToggle, [data-fabric-component='SelectToggle'], .select2-selection, .select2-choice, .chosen-single")
  ) return "combobox";
  if (control.tagName === "TEXTAREA") return "textarea";
  if (control.tagName === "SELECT") return "select";
  return (control.getAttribute("type") || "text").toLowerCase();
}

function isCompositeComboboxShell(control) {
  return control.classList?.contains("select__control") &&
    Boolean(control.querySelector("input[role='combobox'], input[aria-autocomplete='list']"));
}

function getControlOptions(control) {
  if (control.tagName === "SELECT") {
    return Array.from(control.options)
      .filter((option) => !option.disabled)
      .map((option) => ({ text: cleanLabel(option.textContent || ""), value: cleanLabel(option.value || "") }))
      .filter((option) => option.text && !isGeneratedChoicePlaceholder(option.text, option.value))
      .map((option) => option.text);
  }
  if (getControlType(control) === "combobox") {
    const options = [
      ...getVisibleChoiceElements(control),
      ...(atsAdapters?.getVisibleOptions?.(control) || [])
    ];
    return Array.from(new Set(options.map((option) => cleanLabel(option.textContent || option.getAttribute?.("aria-label") || "")).filter(Boolean)));
  }
  if (getControlType(control) === "checkbox") return ["Yes", "No"];
  return [];
}

function getControlValue(control) {
  if (getControlType(control) === "checkbox") return isCheckboxChecked(control) ? "Yes" : "";
  if (getControlType(control) === "button-group") {
    const controls = getControlsByFieldId(control.dataset.autoBidFieldId || "");
    return getSelectedChoiceButtonLabel(controls);
  }
  if (getControlType(control) === "combobox") {
    const value = cleanLabel(control.value || getComboboxSelectedText(control) || control.textContent || "");
    return isPlaceholderChoice(value, value) ? "" : value;
  }
  if (getControlType(control) === "contenteditable") return cleanLabel(control.textContent || "");
  return control.value || "";
}

function isRequired(control) {
  const rootNode = control.getRootNode?.() || document;
  const explicitLabel = control.id ? queryOne(`label[for="${cssEscape(control.id)}"]`, rootNode) : null;
  const fieldContainer = getFieldContainer(control);
  const requiredMarkerElements = [
    control,
    explicitLabel,
    control.closest?.("label"),
    control.closest?.("fieldset"),
    fieldContainer,
    fieldContainer?.querySelector?.("label, legend, [class*='label' i], [class*='required' i]")
  ].filter(Boolean);
  const requiredText = [
    getFieldLabel(control),
    ["checkbox", "radio"].includes(getControlType(control)) ? getChoiceQuestionLabel(control) : "",
    getDescribedByText(control)
  ].join(" ");
  return Boolean(
    atsAdapters?.isRequired?.(control) ||
    control.required ||
    control.getAttribute("aria-required") === "true" ||
    control.closest("[aria-required='true'], [data-required='true']") ||
    requiredMarkerElements.some(hasRequiredSemanticMarker) ||
    /\*/.test(requiredText) ||
    /\brequired\b/i.test(requiredText)
  );
}

function hasRequiredSemanticMarker(element) {
  if (!element) return false;
  if (element.matches?.("[required], [aria-required='true'], [data-required='true']")) return true;

  const markerText = [
    element.getAttribute?.("class"),
    element.getAttribute?.("data-testid"),
    element.getAttribute?.("data-ui"),
    element.getAttribute?.("data-qa"),
    element.getAttribute?.("aria-label")
  ].filter(Boolean).join(" ").toLowerCase();

  if (/(^|[^a-z])required([^a-z]|$)/.test(markerText) && !/(not-required|optional)/.test(markerText)) return true;
  const nestedControls = element.querySelectorAll?.("input, textarea, select, [role='checkbox'], [role='radio'], [role='combobox']")?.length || 0;
  if (nestedControls > 1 && !element.matches?.("label, legend, [class*='label' i]")) return false;
  return /\*/.test(cleanLabel(element.textContent || ""));
}

function getFieldLabel(control) {
  const id = control.id;
  const fromFor = id ? queryOne(`label[for="${cssEscape(id)}"]`, control.getRootNode?.() || document) : null;
  const closestLabel = control.closest("label");
  const ariaLabelledBy = control.getAttribute("aria-labelledby");
  const fromAria = ariaLabelledBy
    ? ariaLabelledBy.split(/\s+/).map((part) => document.getElementById(part)?.textContent || "").join(" ")
    : "";
  const container = getFieldContainer(control);
  const containerLabels = container
    ? Array.from(container.querySelectorAll(FIELD_LABEL_SELECTOR))
      .filter((label) => label !== control && !label.contains(control))
      .map((label) => label.textContent)
    : [];
  const siblingLabels = getSiblingLabelCandidates(control);
  const isChoiceControl = ["checkbox", "radio"].includes(getControlType(control));
  const choiceLabels = isChoiceControl ? getChoiceInlineLabelCandidates(control) : [];
  const visualLabel = isChoiceControl ? "" : getVisualFieldLabel(control);
  const candidates = [
    fromFor?.textContent,
    closestLabel?.textContent,
    fromAria,
    control.getAttribute("aria-label"),
    visualLabel,
    ...siblingLabels,
    ...(atsAdapters?.getLabelCandidates?.(control) || []),
    ...choiceLabels,
    ...containerLabels,
    control.getAttribute("placeholder"),
    control.name,
    control.id
  ];
  return chooseBestFieldLabel(candidates);
}

function getNearbyText(control) {
  const container = getFieldContainer(control) || control.closest("fieldset, .form-group, .field, .input, div, li") || control.parentElement;
  return cleanLabel(container?.textContent || "");
}

function getFieldText(control) {
  return [getFieldLabel(control), control.name, control.id, control.getAttribute("placeholder"), control.getAttribute("autocomplete")].filter(Boolean).join(" ");
}

function getFieldContainer(control) {
  return atsAdapters?.getFieldContainer?.(control) || control.closest(FIELD_CONTAINER_SELECTOR) || control.parentElement;
}

function getDescribedByText(control) {
  return (control.getAttribute("aria-describedby") || "")
    .split(/\s+/)
    .map((part) => document.getElementById(part)?.textContent || "")
    .join(" ");
}

function getSiblingLabelCandidates(control) {
  const candidates = [];
  let current = control.previousElementSibling;
  for (let index = 0; current && index < 3; index += 1, current = current.previousElementSibling) {
    if (isLabelLikeElement(current)) candidates.push(current.textContent);
  }

  current = control.parentElement?.previousElementSibling || null;
  for (let index = 0; current && index < 2; index += 1, current = current.previousElementSibling) {
    if (isLabelLikeElement(current)) candidates.push(current.textContent);
  }

  return candidates;
}

function getChoiceInlineLabelCandidates(control) {
  const candidates = [];
  const root = control.closest("label, [class*='checkbox' i], [class*='option' i], li");
  if (root && isVisible(root)) candidates.push(root.textContent);

  let current = control.nextElementSibling;
  for (let index = 0; current && index < 4; index += 1, current = current.nextElementSibling) {
    if (isLabelLikeElement(current)) candidates.push(current.textContent);
  }

  current = control.parentElement?.nextElementSibling || null;
  for (let index = 0; current && index < 3; index += 1, current = current.nextElementSibling) {
    if (isLabelLikeElement(current)) candidates.push(current.textContent);
  }

  const parentText = getDirectQuestionText(control.parentElement, control);
  if (parentText) candidates.push(parentText);
  return candidates;
}

function isLabelLikeElement(element) {
  if (!element || !isVisible(element)) return false;
  if (element.matches("input, textarea, select, button, [role='combobox']")) return false;
  if (element.querySelector("input, textarea, select, button, [role='combobox']")) return false;
  return Boolean(cleanLabel(element.textContent || ""));
}

function chooseBestFieldLabel(candidates) {
  const labels = candidates
    .map(cleanLabel)
    .filter(Boolean)
    .filter((label, index, list) => list.indexOf(label) === index);
  const readable = labels.filter((label) => /[a-z]/i.test(label) && !isPlaceholderChoice(label, label));
  return (readable.find((label) => /[?*]|\brequired\b/i.test(label)) || readable[0] || labels[0] || "").slice(0, 300);
}

function getChoiceButtonLabel(button) {
  return cleanLabel(button?.textContent || button?.getAttribute?.("aria-label") || button?.getAttribute?.("value") || button?.value || "");
}

function getRadioGroupLabel(group) {
  const first = group[0];
  const questionLabel = getChoiceQuestionLabel(first);
  if (questionLabel) return questionLabel;

  const options = group.map(getRadioOptionLabel).filter(Boolean);
  const label = getFieldLabel(first);
  if (options.some((option) => normalize(option) === normalize(label))) {
    return cleanLabel(first.name || label || options.join(" "));
  }
  return label;
}

function getRadioGroupKey(control, index) {
  const name = cleanLabel(control.name || control.getAttribute?.("name") || "");
  if (name) return `name:${name}`;

  const explicitGroup = control.closest?.("[role='radiogroup'], fieldset");
  if (explicitGroup) return explicitGroup;

  if (control.getAttribute?.("role") === "radio") {
    let current = control.parentElement;
    for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
      const radios = Array.from(current.querySelectorAll?.("[role='radio']") || [])
        .filter((candidate) => isVisible(candidate) && !candidate.querySelector?.("input[type='radio']"));
      if (radios.length >= 2 && radios.length <= 20) return current;
    }
  }

  return control.id ? `id:${control.id}` : `radio:${index}`;
}

function getRadioGroupValue(group) {
  const checked = group.find(isRadioChecked);
  return checked ? getRadioOptionLabel(checked) || checked.value || "" : "";
}

function getChoiceGroupLabel(group) {
  const first = group[0];
  const questionLabel = getChoiceQuestionLabel(first);
  if (questionLabel) return questionLabel;

  const options = group.map(getChoiceButtonLabel).filter(Boolean);
  const label = getFieldLabel(first);
  if (options.some((option) => normalize(option) === normalize(label))) {
    return cleanLabel(first.name || first.id || label || options.join(" "));
  }
  return label;
}

function getChoiceQuestionLabel(control) {
  const question = control.closest(".application-question, [class*='question' i], fieldset, [role='radiogroup'], [role='group']");
  const questionText = findQuestionLabelText(question, control);
  if (questionText) return questionText;

  const list = control.closest("ul[data-qa='multiple-choice'], [data-qa*='multiple-choice' i], [role='radiogroup']");
  let current = list?.parentElement || control.parentElement;
  for (let depth = 0; current && depth < 5; depth += 1, current = current.parentElement) {
    const label = findQuestionLabelText(current, control);
    if (label) return label;
    const previous = current.previousElementSibling;
    const previousLabel = findQuestionLabelText(previous, control);
    if (previousLabel) return previousLabel;
    const nearbyLabel = getNearbyPreviousQuestionText(current, control);
    if (nearbyLabel) return nearbyLabel;
  }

  const visualLabel = getVisualQuestionLabel(control);
  if (visualLabel) return visualLabel;

  return "";
}

function getCheckboxGroupQuestionLabel(control, checkboxGroup = []) {
  const group = checkboxGroup.length > 0 ? checkboxGroup : [control];
  const anchor = group[0] || control;
  const anchorQuestion = getChoiceQuestionLabel(anchor);
  if (anchorQuestion) return anchorQuestion;
  if (anchor !== control) return getChoiceQuestionLabel(control);
  return "";
}

function findQuestionLabelText(root, control) {
  if (!root || root === control || root.contains?.(control) && root.matches?.("label")) return "";
  const selectors = [
    ".application-label .text",
    ".application-label",
    "label",
    FIELD_LABEL_SELECTOR,
    "[data-qa*='label' i]",
    "[data-ui*='label' i]",
    "[data-testid*='label' i]",
    "[class*='question-text' i]",
    "[class*='question-label' i]",
    "[class*='field-label' i]",
    "legend"
  ];
  for (const selector of selectors) {
    const element = root.matches?.(selector) ? root : root.querySelector?.(selector);
    if (element && !element.contains(control)) {
      const text = cleanLabel(element.textContent || "");
      if (text && !isRadioOptionText(control, text)) return text;
    }
  }
  const directText = getDirectQuestionText(root, control);
  if (directText && !isRadioOptionText(control, directText)) return directText;
  return "";
}

function isRadioOptionText(control, text) {
  const optionList = control.closest("ul[data-qa='multiple-choice'], [data-qa*='multiple-choice' i], [role='radiogroup']") || control.parentElement;
  const optionTexts = [
    ...Array.from(optionList?.querySelectorAll("label, button, [role='button'], [role='radio']") || []),
    control
  ].map((item) => normalize(item.textContent || item.getAttribute?.("aria-label") || item.value || ""));
  return optionTexts.includes(normalize(text));
}

function getDirectQuestionText(root, control) {
  if (!root || root === control || root.contains?.(control) && root.matches?.("label")) return "";
  const pieces = [];
  for (const child of Array.from(root.childNodes || [])) {
    if (child.nodeType === Node.TEXT_NODE) {
      pieces.push(child.textContent);
      continue;
    }
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    if (child === control || child.contains(control)) continue;
    if (child.matches?.("input, textarea, select, button, [role='button'], [role='radio'], [role='combobox']")) continue;
    if (child.querySelector?.("input, textarea, select, button, [role='button'], [role='radio'], [role='combobox']")) continue;
    pieces.push(child.textContent);
  }
  return cleanLabel(pieces.join(" "));
}

function getNearbyPreviousQuestionText(element, control) {
  let current = element?.previousElementSibling || null;
  for (let index = 0; current && index < 4; index += 1, current = current.previousElementSibling) {
    if (!isLabelLikeElement(current)) continue;
    const text = cleanLabel(current.textContent || "");
    if (text && !isRadioOptionText(control, text)) return text;
  }
  return "";
}

function getVisualQuestionLabel(control) {
  const groupRect = getChoiceGroupRect(control);
  if (!groupRect) return "";

  const scope = control.closest("form, main, [role='main']") || document.body;
  const selector = [
    "label",
    "legend",
    "div",
    "p",
    "div",
    "span",
    "strong",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6"
  ].join(",");
  const candidates = Array.from(scope.querySelectorAll(selector))
    .map((element) => {
      if (!isVisible(element) || element === control || element.contains(control)) return null;
      if (element.querySelector("input, textarea, select, button, [role='button'], [role='radio'], [role='combobox']")) return null;
      const text = cleanLabel(element.textContent || "");
      if (!text || text.length > 240 || isRadioOptionText(control, text)) return null;

      const rect = element.getBoundingClientRect();
      const verticalDistance = groupRect.top - rect.bottom;
      if (verticalDistance < -8 || verticalDistance > 220) return null;
      if (rect.right < groupRect.left - 80 || rect.left > groupRect.right + 80) return null;

      return {
        text,
        verticalDistance,
        score: verticalDistance - (isQuestionLikeText(text) ? 80 : 0)
      };
    })
    .filter(Boolean)
    .sort((left, right) => left.score - right.score);

  return candidates[0]?.text || "";
}

function getVisualFieldLabel(control) {
  const controlRect = control?.getBoundingClientRect?.();
  if (!controlRect || controlRect.width <= 0 || controlRect.height <= 0) return "";

  const scope = control.closest?.("form, main, [role='main']") || document.body;
  const candidates = Array.from(scope?.querySelectorAll?.([
    "label",
    "legend",
    "p",
    "span",
    "strong",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "[class*='label' i]",
    "[data-testid*='label' i]"
  ].join(",")) || [])
    .map((element) => {
      if (!isVisible(element) || element === control || element.contains(control)) return null;
      if (element.querySelector?.("input, textarea, select, button, [role='button'], [role='radio'], [role='combobox']")) return null;
      const text = cleanLabel(element.textContent || element.getAttribute?.("aria-label") || "");
      if (!text || text.length > 300) return null;

      const rect = element.getBoundingClientRect?.();
      if (!rect) return null;
      const verticalDistance = controlRect.top - rect.bottom;
      if (verticalDistance < -8 || verticalDistance > 180) return null;
      if (rect.right < controlRect.left - 80 || rect.left > controlRect.right + 80) return null;

      const labelLike = element.matches?.("label, legend, strong, [class*='label' i], [data-testid*='label' i]");
      if (!labelLike && !isQuestionLikeText(text)) return null;
      const horizontalDistance = Math.abs(rect.left - controlRect.left);
      return {
        text,
        score: verticalDistance + Math.min(horizontalDistance, 120) * 0.15 - (isQuestionLikeText(text) ? 80 : 0)
      };
    })
    .filter(Boolean)
    .sort((left, right) => left.score - right.score);

  return candidates[0]?.text || "";
}

function getChoiceGroupRect(control) {
  const group = control.closest("[role='radiogroup'], ul[data-qa='multiple-choice'], [data-qa*='multiple-choice' i], fieldset") ||
    control.closest("label")?.parentElement ||
    control.parentElement;
  return (group || control).getBoundingClientRect?.() || null;
}

function isQuestionLikeText(text) {
  return /[?*]$|\b(do you|are you|can you|have you|how many|what|which|where|when|please|speak|language|experience|salary|notice|authorization|eligible|consent)\b/i.test(text);
}

function getChoiceButtonValue(button) {
  return cleanLabel(button?.getAttribute?.("data-value") || button?.getAttribute?.("value") || button?.value || getChoiceButtonLabel(button));
}

function getSelectedChoiceButtonLabel(buttons) {
  const selected = buttons.find(isChoiceButtonSelected);
  return selected ? getChoiceButtonLabel(selected) || getChoiceButtonValue(selected) : "";
}

function isChoiceButtonSelected(button) {
  if (!button) return false;
  if (button.getAttribute("aria-pressed") === "true" || button.getAttribute("aria-checked") === "true") return true;
  if (button.getAttribute("data-selected") === "true" || button.getAttribute("data-state") === "checked") return true;
  if (button.matches("[aria-selected='true'], .selected, .active, [class*='selected' i], [class*='active' i]")) return true;
  return false;
}

function getRadioOptionLabel(control) {
  const forLabel = control.id ? queryOne(`label[for="${cssEscape(control.id)}"]`, control.getRootNode?.() || document) : null;
  const siblingLabel = control.closest("[class*='option' i], li, div")?.querySelector?.("label");
  const value = cleanLabel(control.value || "");
  return cleanLabel(
    forLabel?.textContent ||
    control.closest("label")?.textContent ||
    (siblingLabel && !siblingLabel.contains(control) ? siblingLabel.textContent : "") ||
    control.getAttribute("aria-label") ||
    control.textContent ||
    (normalize(value) === "on" ? "" : value)
  );
}

function isRadioChecked(control) {
  if (!control) return false;
  if (control.matches?.("input[type='radio']")) return Boolean(control.checked);
  if (control.getAttribute?.("aria-checked") === "true") return true;
  if (["checked", "selected", "on"].includes(control.getAttribute?.("data-state"))) return true;
  return Boolean(control.matches?.("[aria-selected='true'], .checked, .selected, .active, [class*='checked' i], [class*='selected' i], [class*='active' i]"));
}

function cleanLabel(text) {
  return String(text || "").replace(/\s+/g, " ").replace(/\s+\*/g, " *").trim().slice(0, 300);
}

function getPlainQuestionText(text) {
  return cleanLabel(text)
    .replace(/\s*\*\s*/g, " ")
    .replace(/\s+\(required\)\s*$/i, "")
    .replace(/\brequired\b\s*:?\s*/gi, "")
    .replace(/\s+([?.!,])/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function getCheckboxOptionLabel(control, questionText = "") {
  const root = control.closest("label, [class*='checkbox' i], [class*='option' i], li, [role='checkbox']");
  const rootNode = control.getRootNode?.() || document;
  const forLabel = control.id ? queryOne(`label[for="${cssEscape(control.id)}"]`, rootNode) : null;
  const candidates = [
    forLabel?.textContent,
    control.closest("label")?.textContent,
    root && root !== control ? root.textContent : "",
    ...getChoiceInlineLabelCandidates(control),
    control.getAttribute("aria-label"),
    normalize(control.value) === "on" ? "" : control.value
  ];
  const question = getPlainQuestionText(questionText);

  return candidates
    .map((candidate) => stripQuestionFromOptionText(candidate, question))
    .filter(Boolean)
    .find((candidate) => normalize(candidate) !== normalize(question) && !/^(yes|no|true|false|on|off)$/.test(normalize(candidate))) || "";
}

function stripQuestionFromOptionText(text, questionText) {
  let value = getPlainQuestionText(text);
  const question = getPlainQuestionText(questionText);
  if (!value) return "";
  if (!question) return value;

  const normalizedValue = normalize(value);
  const normalizedQuestion = normalize(question);
  if (normalizedValue === normalizedQuestion) return "";

  if (normalizedValue.includes(normalizedQuestion)) {
    value = value.replace(new RegExp(escapeRegExp(question), "i"), "").trim();
  }

  return getPlainQuestionText(value.replace(/^[-:\s]+/, ""));
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getCurrentlyFilledFieldIds(fields) {
  const filledIds = new Set();
  for (const field of fields) {
    if (hasFieldCurrentValue(field)) filledIds.add(field.id);
  }

  const checkboxGroups = new Map();
  for (const field of fields) {
    if (field.type !== "checkbox" || !field.option || !field.name) continue;
    const key = normalize(field.name);
    if (!key) continue;
    if (!checkboxGroups.has(key)) checkboxGroups.set(key, []);
    checkboxGroups.get(key).push(field);
  }
  for (const group of checkboxGroups.values()) {
    if (group.length < 2 || !group.some((field) => hasFieldCurrentValue(field))) continue;
    group.forEach((field) => filledIds.add(field.id));
  }

  return filledIds;
}

async function closeOpenChoiceMenus() {
  const targets = [
    document.activeElement,
    ...queryAll("[aria-expanded='true']")
  ]
    .filter(Boolean)
    .filter((target) => target?.getAttribute?.("aria-expanded") === "true")
    .filter((target) => {
      const role = normalize(target.getAttribute?.("role") || "");
      const hasPopup = normalize(target.getAttribute?.("aria-haspopup") || "");
      return /combobox|button|listbox|menu/.test(role) || /listbox|menu|true/.test(hasPopup);
    });

  targets.forEach((target) => {
    target.dispatchEvent?.(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
    target.dispatchEvent?.(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
  });

  if (targets.length > 0) await sleep(120);
}

async function applyAnswers(answers, skipFilledIds = new Set(), fields = []) {
  let filled = 0;
  let missed = 0;
  const filledIds = new Set();
  const missedIds = new Set();
  const fieldsById = new Map((fields || []).map((field) => [field.id, field]));
  const claimedLiveFields = new Set();

  for (const answer of answers) {
    const requestedFieldId = answer.field_id;
    const sourceField = fieldsById.get(requestedFieldId) || null;
    const binding = resolveLiveFieldForAnswer(answer, sourceField, claimedLiveFields);
    if (!binding) {
      const attemptKey = getGeneratedAnswerAttemptKey(answer, sourceField);
      const attempt = recordGeneratedAnswerFillFailure(attemptKey);
      missed += 1;
      missedIds.add(requestedFieldId);
      traceAutoBid("answer:unbound", {
        field_id: requestedFieldId,
        question: answer.question || sourceField?.question || sourceField?.label || "",
        source: answer.source || "",
        attempt,
        max_attempts: MAX_FIELD_FILL_ATTEMPTS
      });
      continue;
    }

    const { field, controls, rebound, score } = binding;
    const attemptKey = getGeneratedAnswerAttemptKey(answer, field);
    claimedLiveFields.add(getStableFieldBindingKey(field));
    if (rebound) {
      traceAutoBid("answer:rebound", {
        requested_field_id: requestedFieldId,
        live_field_id: field.id,
        question: field.question || field.label || "",
        score
      });
    }

    const answerAlreadyMatches = doesGeneratedAnswerMatchField(answer, field, controls);
    if (answerAlreadyMatches) {
      traceAutoBid("answer:skipped-filled", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        source: answer.source || "",
        current: getCurrentChoiceSummary(controls)
      });
      filledIds.add(requestedFieldId);
      filledIds.add(field.id);
      continue;
    }

    const alreadyFilled = hasFieldCurrentValue(field, controls);
    if (alreadyFilled) {
      traceAutoBid("answer:skipped-filled", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        source: answer.source || "",
        current: getCurrentChoiceSummary(controls),
        reason: "preserve-current-value"
      });
      filledIds.add(requestedFieldId);
      filledIds.add(field.id);
      continue;
    }

    const previousAttempts = getGeneratedAnswerFillAttempts(attemptKey);
    if (previousAttempts >= MAX_FIELD_FILL_ATTEMPTS) {
      traceAutoBid("answer:stopped-after-max-attempts", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        source: answer.source || "",
        attempts: previousAttempts,
        max_attempts: MAX_FIELD_FILL_ATTEMPTS
      });
      missed += 1;
      missedIds.add(requestedFieldId);
      continue;
    }

    if (isSuspiciousNarrativeBooleanAnswer(answer, field, controls)) {
      const attempt = recordGeneratedAnswerFillFailure(attemptKey);
      traceAutoBid("answer:rejected-narrative-boolean", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        question: field.question || field.label || "",
        answer: answer.value || "",
        source: answer.source || "",
        attempt,
        max_attempts: MAX_FIELD_FILL_ATTEMPTS,
        stopped: attempt >= MAX_FIELD_FILL_ATTEMPTS
      });
      missed += 1;
      missedIds.add(requestedFieldId);
      continue;
    }

    const declaredOptions = sourceField?.options?.length > 0 ? sourceField.options : field.options || [];
    const declaredChoice = isChoiceFieldType(field.type) && declaredOptions.length > 0
      ? findBestChoice(declaredOptions.map((option) => ({ text: option, value: option })), answer.value || "")
      : null;
    if (isChoiceFieldType(field.type) && declaredOptions.length > 0 && !declaredChoice) {
      const attempt = recordGeneratedAnswerFillFailure(attemptKey);
      traceAutoBid("answer:invalid-choice", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        source: answer.source || "",
        options: declaredOptions,
        attempt,
        max_attempts: MAX_FIELD_FILL_ATTEMPTS,
        stopped: attempt >= MAX_FIELD_FILL_ATTEMPTS
      });
      missed += 1;
      missedIds.add(requestedFieldId);
      continue;
    }

    const valueToApply = declaredChoice?.text || declaredChoice?.value || answer.value || "";
    const selected = await setControlsValue(controls, valueToApply, field);
    const verified = selected && await waitForGeneratedAnswerMatch({
      ...answer,
      field_id: field.id,
      value: valueToApply
    }, field, controls);
    if (verified) {
      generatedAnswerFillAttempts.delete(attemptKey);
      traceAutoBid("answer:applied", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        applied_answer: valueToApply,
        source: answer.source || "",
        current: getCurrentChoiceSummary(controls)
      });
      filled += 1;
      filledIds.add(requestedFieldId);
      filledIds.add(field.id);
    } else {
      const attempt = recordGeneratedAnswerFillFailure(attemptKey);
      traceAutoBid(selected ? "answer:verification-failed" : "answer:missed", {
        field_id: field.id,
        requested_field_id: requestedFieldId,
        answer: answer.value || "",
        source: answer.source || "",
        selected,
        current: getCurrentChoiceSummary(controls),
        attempt,
        max_attempts: MAX_FIELD_FILL_ATTEMPTS,
        stopped: attempt >= MAX_FIELD_FILL_ATTEMPTS
      });
      missed += 1;
      missedIds.add(requestedFieldId);
    }
  }

  return { filled, missed, filledIds, missedIds };
}

function getGeneratedAnswerAttemptKey(answer, field) {
  const stableField = field ? getStableFieldBindingKey(field) : "";
  return stableField || [
    normalize(answer?.question || ""),
    normalize(answer?.option || ""),
    String(answer?.field_id || "")
  ].join("|");
}

function getGeneratedAnswerFillAttempts(attemptKey) {
  return Number(generatedAnswerFillAttempts.get(attemptKey) || 0);
}

function recordGeneratedAnswerFillFailure(attemptKey) {
  const attempts = Math.min(MAX_FIELD_FILL_ATTEMPTS, getGeneratedAnswerFillAttempts(attemptKey) + 1);
  generatedAnswerFillAttempts.set(attemptKey, attempts);
  return attempts;
}

function resolveLiveFieldForAnswer(answer, sourceField, claimedLiveFields = new Set()) {
  const liveFields = collectFields();
  const directField = liveFields.find((field) => field.id === answer.field_id);
  if (directField) {
    const controls = getControlsByFieldId(directField.id);
    if (controls.length > 0) {
      return { field: directField, controls, rebound: false, score: 10000 };
    }
  }

  const scored = liveFields
    .filter((field) => !claimedLiveFields.has(getStableFieldBindingKey(field)))
    .map((field) => ({
      field,
      score: scoreLiveFieldForAnswer(field, sourceField, answer)
    }))
    .filter((candidate) => candidate.score >= 260)
    .sort((left, right) => right.score - left.score);

  const best = scored[0];
  if (!best) return null;
  const controls = getControlsByFieldId(best.field.id);
  if (controls.length === 0) return null;
  return { field: best.field, controls, rebound: best.field.id !== answer.field_id, score: best.score };
}

function scoreLiveFieldForAnswer(liveField, sourceField, answer) {
  if (sourceField && !areFieldTypesCompatible(sourceField.type, liveField.type)) return -1;

  let score = sourceField?.type === liveField.type ? 120 : 45;
  const liveQuestion = normalize(liveField.question || liveField.label || "");
  const sourceQuestion = normalize(sourceField?.question || sourceField?.label || "");
  const answerQuestion = normalize(answer.question || "");
  const liveLabel = normalize(liveField.label || "");
  const sourceLabel = normalize(sourceField?.label || "");

  score += scoreFieldIdentityText(liveQuestion, sourceQuestion, 720, 330);
  score += scoreFieldIdentityText(liveQuestion, answerQuestion, 640, 300);
  score += scoreFieldIdentityText(liveLabel, sourceLabel, 420, 180);
  score += scoreFieldIdentityText(normalize(liveField.name), normalize(sourceField?.name), 320, 100);
  score += scoreFieldIdentityText(normalize(liveField.placeholder), normalize(sourceField?.placeholder), 180, 70);

  const liveOption = normalize(liveField.option || "");
  const sourceOption = normalize(sourceField?.option || answer.option || "");
  score += scoreFieldIdentityText(liveOption, sourceOption, 520, 220);

  const liveOptions = new Set((liveField.options || []).map(normalize).filter(Boolean));
  const sourceOptions = (sourceField?.options || []).map(normalize).filter(Boolean);
  const sharedOptions = sourceOptions.filter((option) => liveOptions.has(option)).length;
  if (sharedOptions > 0) score += Math.min(240, sharedOptions * 80);
  if (sourceField && Boolean(sourceField.required) === Boolean(liveField.required)) score += 20;

  return score;
}

function scoreFieldIdentityText(left, right, exactScore, partialScore) {
  if (!left || !right) return 0;
  if (left === right) return exactScore;
  if (Math.min(left.length, right.length) >= 8 && (left.includes(right) || right.includes(left))) return partialScore;
  return 0;
}

function areFieldTypesCompatible(left, right) {
  if (!left || !right || left === right) return true;
  if (isChoiceFieldType(left) && isChoiceFieldType(right)) return true;
  const textTypes = new Set(["text", "email", "tel", "url", "number", "date", "textarea"]);
  return textTypes.has(left) && textTypes.has(right);
}

function getStableFieldBindingKey(field) {
  return [
    normalize(field.question || field.label || ""),
    normalize(field.option || ""),
    normalize(field.name || ""),
    field.type || ""
  ].join("|");
}

async function applyResidenceAnswers(fields, answers, filledIds) {
  const filledLocalIds = new Set();
  const countryAnswer = findProfileCountryAnswer(fields, answers);
  if (!countryAnswer) return { filled: 0, missed: 0, filledIds: filledLocalIds };

  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !isResidenceField(field) || !["select", "radio", "combobox"].includes(field.type)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasCurrentChoiceValue(controls)) continue;

    if (await setControlsValue(controls, countryAnswer)) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("residence:applied", {
        field_id: field.id,
        answer: countryAnswer,
        current: getCurrentChoiceSummary(controls)
      });
    } else {
      missed += 1;
      traceAutoBid("residence:missed", {
        field_id: field.id,
        answer: countryAnswer
      });
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

async function applyBasedInLocationAnswers(fields, answers, filledIds) {
  const filledLocalIds = new Set();
  const locationAnswer = findProfileCountryAnswer(fields, answers) || await getProfileStaticLocationAnswer();
  if (!locationAnswer) return { filled: 0, missed: 0, filledIds: filledLocalIds };

  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !isBasedInLocationField(field) || !isSemanticBooleanFieldType(field.type)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasCurrentChoiceValue(controls)) continue;

    const answer = locationAnswerMatchesQuestion(locationAnswer, field.question || field.label) ? "Yes" : "No";
    if (await setControlsValue(controls, answer)) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("location-choice:applied", {
        field_id: field.id,
        label: field.label,
        answer,
        location: locationAnswer,
        current: getCurrentChoiceSummary(controls)
      });
    } else {
      missed += 1;
      traceAutoBid("location-choice:missed", {
        field_id: field.id,
        label: field.label,
        answer,
        location: locationAnswer
      });
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

async function applySensitiveDemographicDeclineAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !field.required || !isSensitiveDemographicChoiceField(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasCurrentChoiceValue(controls)) continue;

    let options = field.options || [];
    if (field.type === "combobox" && options.length === 0) {
      options = (await getComboboxChoices(controls[0])).map((option) => cleanLabel(
        option.textContent || option.getAttribute("data-value") || option.getAttribute("value") || ""
      ));
      field.options = uniqueNonEmptyValues(options);
      await closeCombobox(controls[0]);
    }

    const answer = findSensitiveDemographicDeclineOption(options);
    if (!answer) {
      traceAutoBid("sensitive-demographic:no-decline-option", {
        field_id: field.id,
        label: field.label,
        options
      });
      continue;
    }

    const selected = await setControlsValue(controls, answer, field);
    traceAutoBid("sensitive-demographic:decline-result", {
      field_id: field.id,
      label: field.label,
      answer,
      selected,
      current: getCurrentChoiceSummary(controls)
    });
    if (selected) {
      filled += 1;
      filledLocalIds.add(field.id);
    } else {
      missed += 1;
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

function isSensitiveDemographicChoiceField(field) {
  if (!field || !isChoiceFieldType(field.type)) return false;
  const text = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(gender|race|ethnicity|ethnic|disability|veteran|protected veteran|sexual orientation|pronoun)/.test(text);
}

function findSensitiveDemographicDeclineOption(options) {
  const choices = (options || [])
    .map((option) => String(option || "").trim())
    .filter(Boolean)
    .map((option, index) => ({ option, score: scoreSensitiveDemographicDeclineOption(option, index) }))
    .filter((item) => item.score > 0)
    .sort((left, right) => right.score - left.score);
  return choices[0]?.option || "";
}

function scoreSensitiveDemographicDeclineOption(option, index) {
  const text = normalize(option);
  if (!text || isPlaceholderChoice(text, text)) return -100;
  let score = 0;
  if (/(prefer|choose|wish|want).{0,30}not.{0,30}(say|answer|disclose|identify|respond)/.test(text)) score += 500;
  if (/(decline|do not wish|don t wish|not disclosed|rather not|no answer)/.test(text)) score += 420;
  if (/(prefer not to say|prefer not to answer|decline to self identify)/.test(text)) score += 300;
  return score > 0 ? score + Math.max(0, 20 - index) : 0;
}

function isSemanticBooleanFieldType(type) {
  return isChoiceFieldType(type) || ["text", "search", "textarea"].includes(type);
}

async function applyLanguageChoiceAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;
  let knownLanguages = [];

  try {
    const profile = await send("GET_PROFILE_STATIC_FIELDS");
    knownLanguages = getProfileLanguages(profile?.static_fields || {});
  } catch (error) {
    traceAutoBid("language-choice:profile-unavailable", { message: error.message || String(error) });
  }

  for (const field of fields) {
    if (!isLanguageChoiceField(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    const questionLanguages = getQuestionLanguageAliases(field.question || field.label);
    const answer = getLanguageChoiceAnswer(field, questionLanguages, knownLanguages);
    if (!answer) continue;

    const current = getCurrentChoiceSummary(controls);
    if (isLanguageChoiceAlreadyAcceptable(field, current, answer)) {
      filledLocalIds.add(field.id);
      continue;
    }

    if (await setControlsValue(controls, answer)) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("language-choice:applied", {
        field_id: field.id,
        label: field.label,
        asked_languages: questionLanguages,
        answer,
        current: getCurrentChoiceSummary(controls)
      });
    } else {
      missed += 1;
      traceAutoBid("language-choice:missed", {
        field_id: field.id,
        label: field.label,
        asked_languages: questionLanguages,
        answer
      });
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

async function applyReferralSourceAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !isReferralSourceField(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasFieldCurrentValue(field, controls)) continue;

    if (await setReferralSourceValue(controls, field)) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("referral-source:applied", {
        field_id: field.id,
        label: field.label,
        answer: "LinkedIn",
        current: getCurrentChoiceSummary(controls)
      });
    } else {
      missed += 1;
      traceAutoBid("referral-source:missed", {
        field_id: field.id,
        label: field.label,
        answer: "LinkedIn"
      });
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

async function setReferralSourceValue(controls, field) {
  if (await setControlsValue(controls, "LinkedIn", field)) return true;

  const first = controls[0];
  const type = getControlType(first);
  if (type !== "combobox") return false;

  const options = await getComboboxChoices(first);
  const choices = options.map((option, index) => ({
    control: option,
    text: option.textContent,
    value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent,
    index
  }));
  const match = findBestChoice(choices, "LinkedIn") || choices[0] || null;
  if (!match) {
    await closeCombobox(first);
    return false;
  }

  traceAutoBid("referral-source:default-choice", {
    field_id: field.id,
    label: field.label,
    choice: summarizeChoice(match)
  });
  const selected = await selectComboboxChoice(first, match);
  if (!selected) await closeCombobox(first);
  return selected;
}

async function applyConsentChoiceAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !isConsentChoiceField(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    const current = getCurrentChoiceSummary(controls);
    const answer = getConsentChoiceAnswer(field);
    if (!answer) continue;

    if (isConsentChoiceAlreadyAccepted(current)) {
      filledLocalIds.add(field.id);
      continue;
    }

    const selected = await setControlsValue(controls, answer, field);
    traceAutoBid("consent-choice:result", {
      field_id: field.id,
      label: field.label,
      answer,
      selected,
      previous: current,
      current: getCurrentChoiceSummary(controls)
    });

    if (selected) {
      filled += 1;
      filledLocalIds.add(field.id);
    } else {
      missed += 1;
    }
  }

  return { filled, missed, filledIds: filledLocalIds };
}

async function applyProfileStaticFallbacks(fields, filledIds) {
  const filledLocalIds = new Set();
  let profileStatic;

  try {
    profileStatic = await send("GET_PROFILE_STATIC_FIELDS");
    captureAutoBidProfileContext(profileStatic);
  } catch (error) {
    traceAutoBid("profile-static:error", { message: error.message || String(error) });
    return { filled: 0, missed: 0, filledIds: filledLocalIds };
  }

  const staticFields = profileStatic?.static_fields || {};
  const availableKeys = Object.keys(staticFields).filter((key) => String(staticFields[key] || "").trim());
  const phoneDialResult = await applyPhoneDialCodeSelectors(staticFields);
  let filled = phoneDialResult.filled;
  let missed = phoneDialResult.missed;
  const matched = [];
  const missing = [];

  const orderedFields = [...fields].sort((left, right) =>
    getProfileAddressFillPriority(left) - getProfileAddressFillPriority(right)
  );

  for (const field of orderedFields) {
    if (filledIds.has(field.id)) continue;
    const key = matchProfileStaticFieldKey(field);
    if (!key) continue;

    matched.push({ field_id: field.id, key, label: field.label });
    const rawValue = getProfileStaticValue(staticFields, key);
    let controls = getControlsByFieldId(field.id);
    const candidateValues = getProfileStaticCandidateValues(staticFields, key, field, controls);
    if (candidateValues.length === 0) {
      missing.push({ field_id: field.id, key, label: field.label });
      continue;
    }

    if (controls.length === 0 || !shouldApplyProfileStaticValue(field, controls, rawValue, key)) continue;

    let appliedValue = "";
    for (let candidateIndex = 0; candidateIndex < candidateValues.length; candidateIndex += 1) {
      const value = formatProfileStaticValueForField(key, candidateValues[candidateIndex], field);
      if (candidateIndex > 0) {
        await clearControlsForRetry(controls);
        const rebound = resolveLiveFieldForAnswer({
          field_id: field.id,
          question: field.question || field.label || ""
        }, field);
        if (rebound) controls = rebound.controls;
      }

      const selected = await setControlsValue(controls, value, field, {
        country: getProfileStaticValue(staticFields, "country"),
        locationCountryFirst: ["country", "location", "city", "state_region"].includes(key),
        profileKey: key
      });
      traceAutoBid("profile-static:attempt", {
        field_id: field.id,
        key,
        label: field.label,
        candidate_index: candidateIndex,
        value: shortText(value),
        selected
      });
      if (selected) {
        appliedValue = value;
        break;
      }
      await clearControlsForRetry(controls);
    }

    if (appliedValue) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("profile-static:applied", {
        field_id: field.id,
        key,
        label: field.label,
        value: shortText(appliedValue),
        current: getCurrentChoiceSummary(controls)
      });
      if (key === "country") await sleep(350);
    } else {
      missed += 1;
      traceAutoBid("profile-static:missed", {
        field_id: field.id,
        key,
        label: field.label,
        values: candidateValues.map((value) => shortText(value))
      });
    }
  }

  traceAutoBid("profile-static:summary", {
    profile_id: profileStatic?.profile_id || "",
    available_keys: availableKeys,
    matched,
    missing,
    phone_dial_code: phoneDialResult,
    filled
  });

  return { filled, missed, filledIds: filledLocalIds };
}

function captureAutoBidProfileContext(profileStatic) {
  if (!activeAutoBidProfileId) activeAutoBidProfileId = String(profileStatic?.profile_id || "");
  if (!activeAutoBidProfileEmail) {
    activeAutoBidProfileEmail = String(profileStatic?.static_fields?.email || "").trim().toLowerCase();
  }
}

function getProfileAddressFillPriority(field) {
  const key = matchProfileStaticFieldKey(field);
  if (key === "country") return 0;
  if (key === "state_region") return 1;
  return 2;
}

async function applyPhoneDialCodeSelectors(staticFields) {
  const country = getPhoneDialCountryValue(staticFields);
  if (!country) return { filled: 0, missed: 0, country: "" };

  const controls = getPhoneDialCodeControls();
  let filled = 0;
  let missed = 0;

  for (const control of controls) {
    if (isDialCodeSelectorAlreadyCountry(control, country)) {
      await closeCombobox(control);
      continue;
    }

    const selected = await setControlsValue([control], country);
    traceAutoBid("phone-dial-code:applied", {
      country,
      selected,
      current: getCurrentChoiceSummary([control])
    });

    if (selected || isDialCodeSelectorAlreadyCountry(control, country)) {
      filled += 1;
      await closeCombobox(control);
    } else {
      missed += 1;
      await closeCombobox(control);
    }
  }

  return { filled, missed, country, controls: controls.length };
}

function getPhoneDialCodeControls() {
  const selector = [
    "select",
    "button[role='combobox']",
    "[role='combobox']",
    "[aria-haspopup='listbox']",
    "[aria-haspopup='menu'][aria-expanded]",
    "[data-radix-select-trigger]",
    "[data-slot='select-trigger']",
    ".select__control",
    "[class*='select'][aria-expanded]"
  ].join(",");

  return queryAll(selector).filter((control) => isVisible(control) && isPhoneDialCodeSelector(control));
}

function getPhoneDialCountryValue(staticFields) {
  const country = String(getProfileStaticValue(staticFields, "country") || "").trim();
  if (country) return country;

  const location = String(getProfileStaticValue(staticFields, "location") || "").trim();
  if (!location) return "";
  const pieces = location.split(",").map((part) => part.trim()).filter(Boolean);
  return pieces.length > 1 ? pieces[pieces.length - 1] : "";
}

function isDialCodeSelectorAlreadyCountry(control, country) {
  const rawCurrent = [
    getCurrentChoiceSummary([control]),
    control.value,
    control.textContent,
    getComboboxSelectedText(control)
  ].filter(Boolean).join(" ");
  const current = normalize(rawCurrent);
  const expectedCountry = normalize(country);
  const expectedCode = getDialCodeForCountry(country);

  if (expectedCountry && (current.includes(expectedCountry) || scoreChoice(current, expectedCountry) >= 70)) return true;
  return Boolean(expectedCode && rawCurrent.includes(expectedCode));
}

function getDialCodeForCountry(country) {
  const normalized = normalize(country);
  if (PHONE_DIAL_CODES_BY_COUNTRY[normalized]) return PHONE_DIAL_CODES_BY_COUNTRY[normalized];

  const match = Object.entries(PHONE_DIAL_CODES_BY_COUNTRY)
    .find(([name]) => normalized.includes(name) || name.includes(normalized));
  return match?.[1] || "";
}

async function attachResumeFromSheet() {
  let inputs = getResumeFileInputs();
  if (inputs.length > 0 && inputs.every((input) => isResumeInputAttached(input))) {
    traceAutoBid("resume:already-attached", {
      files: inputs.flatMap(getResumeInputAttachedFilenames),
      server_confirmed: inputs.some(hasServerConfirmedResumeUpload),
      reason: "upload-confirmed"
    });
    return { filled: 0, missed: 0, reason: "already-attached" };
  }

  const existingManagedAttempt = inputs
    .map((input) => ({
      input,
      filename: input.files?.[0]?.name || "",
      context: getManagedResumeUploadContext(input)
    }))
    .map((attempt) => ({
      ...attempt,
      status: getManagedResumeUploadStatus(attempt.context, attempt.filename)
    }))
    .find((attempt) => attempt.filename && attempt.context.managed && attempt.status.status !== "complete");

  if (existingManagedAttempt) {
    let attached = false;
    if (["pending", "uploading"].includes(existingManagedAttempt.status.status)) {
      attached = await waitForResumeInputConfirmation(
        existingManagedAttempt.input,
        existingManagedAttempt.filename,
        RESUME_MANAGED_UPLOAD_TIMEOUT_MS,
        getResumeServerConfirmationSnapshot(existingManagedAttempt.input),
        existingManagedAttempt.context
      );
    }
    const finalStatus = getManagedResumeUploadStatus(
      existingManagedAttempt.context,
      existingManagedAttempt.filename
    );
    traceAutoBid("resume:managed-upload-resumed", {
      attached,
      filename: existingManagedAttempt.filename,
      previous_status: existingManagedAttempt.status.status,
      final_status: finalStatus.status,
      message: finalStatus.message || existingManagedAttempt.status.message || ""
    });
    return attached
      ? { filled: 0, missed: 0, filename: existingManagedAttempt.filename, reason: "existing-upload-completed" }
      : { filled: 0, missed: 1, filename: existingManagedAttempt.filename, reason: "existing-managed-upload-failed" };
  }

  let payload;
  try {
    payload = await fetchResumePayload(inputs);
  } catch (error) {
    traceAutoBid("resume:fetch-error", { message: error.message || String(error) });
    return { filled: 0, missed: 1, reason: "fetch-error" };
  }

  if (!payload?.base64) {
    traceAutoBid("resume:no-file", { payload: summarizeResumePayload(payload) });
    return { filled: 0, missed: 1, reason: "no-file" };
  }

  let file;
  try {
    file = fileFromBase64(payload);
  } catch (error) {
    traceAutoBid("resume:file-error", { message: error.message || String(error), payload: summarizeResumePayload(payload) });
    return { filled: 0, missed: 1, reason: "file-error" };
  }

  let chooserResult = null;
  if (inputs.length === 0) {
    traceAutoBid("resume:no-input-before-chooser", {
      candidates: getFileInputCandidatesDebug(),
      zones: getResumeUploadZoneDebug()
    });
    chooserResult = await attachResumeViaNativeFileChooser(file, payload);
    traceAutoBid("resume:file-chooser-upload", chooserResult);
    if (chooserResult.attached) {
      return { filled: 1, missed: 0, filename: file.name, method: "file-chooser" };
    }
    inputs = await waitForResumeFileInputs();
  }

  let filled = 0;
  let missed = 0;
  for (const input of inputs) {
    if (input.files?.length) continue;
    const attached = await setFileInputValue(input, file, payload);
    traceAutoBid("resume:attached", {
      selected: attached,
      server_confirmed: attached && isServerTrackedResumeInput(input),
      filename: file.name,
      mime_type: file.type,
      context: shortText(getFileInputContextText(input))
    });
    if (attached) filled += 1;
    else missed += 1;
  }

  if (inputs.length === 0) {
    traceAutoBid("resume:no-input", {
      candidates: getFileInputCandidatesDebug(),
      zones: getResumeUploadZoneDebug()
    });
    return { filled: 0, missed: 1, filename: file.name, reason: "no-input" };
  }

  return { filled, missed, filename: file.name };
}

function fetchResumePayload(inputs) {
  return send("SHEET_FETCH_RESUME_FILE", {
    page: collectPageContext(),
    accept: Array.from(new Set(inputs.map((input) => input.accept).filter(Boolean)))
  });
}

async function attachResumeViaNativeFileChooser(file, payload = {}) {
  const base64 = String(payload.base64 || "").replace(/^data:[^,]+,/, "");
  if (!base64) return { attempted: false, attached: false, reason: "no-file-bytes" };

  const button = findResumeAttachButton();
  if (!button) return { attempted: false, attached: false, reason: "attach-button-not-found" };
  const confirmationBefore = getResumeServerConfirmationSnapshot();

  await scrollElementIntoView(button, "center");
  const target = getHitTarget(button);
  const rect = target.getBoundingClientRect();
  const x = rect.left + rect.width / 2;
  const y = rect.top + rect.height / 2;
  if (x < 0 || y < 0 || x > window.innerWidth || y > window.innerHeight) {
    return { attempted: true, attached: false, reason: "attach-button-not-clickable" };
  }

  try {
    const result = await send("NATIVE_FILE_CHOOSER_UPLOAD", {
      x,
      y,
      filename: file.name,
      mime_type: file.type || payload.mime_type || payload.mimeType || "application/pdf",
      base64
    });
    await sleep(900);
    const attached = await waitForAnyResumeUploadConfirmation(
      file.name,
      Boolean(result?.uploaded),
      RESUME_SERVER_CONFIRM_TIMEOUT_MS,
      confirmationBefore
    );
    return {
      attempted: true,
      attached,
      server_confirmed: hasNewResumeServerConfirmation(file.name, confirmationBefore),
      uploaded: Boolean(result?.uploaded),
      files: Number(result?.files || 0),
      button: shortText(button.textContent || button.getAttribute?.("aria-label") || ""),
      filename: file.name,
      reason: result?.reason || ""
    };
  } catch (error) {
    return {
      attempted: true,
      attached: false,
      uploaded: false,
      button: shortText(button.textContent || button.getAttribute?.("aria-label") || ""),
      filename: file.name,
      reason: error.message || String(error)
    };
  }
}

async function waitForResumeFileInputs(timeoutMs = RESUME_FILE_INPUT_WAIT_MS) {
  const started = Date.now();
  let inputs = getResumeFileInputs();
  let revealAttempted = false;

  while (inputs.length === 0 && Date.now() - started < timeoutMs) {
    if (!revealAttempted) {
      revealAttempted = true;
      const revealResult = await revealResumeFileInput();
      if (revealResult.attempted) {
        traceAutoBid("resume:reveal", revealResult);
      }
    }
    await sleep(RESUME_FILE_INPUT_RETRY_MS);
    inputs = getResumeFileInputs();
  }

  return inputs;
}

async function revealResumeFileInput() {
  const button = findResumeAttachButton();
  if (!button) return { attempted: false, clicked: false, reason: "button-not-found" };

  await scrollElementIntoView(button, "center");
  dispatchRealisticMouseClick(button);
  await sleep(600);
  return {
    attempted: true,
    clicked: true,
    text: shortText(button.textContent || button.getAttribute?.("aria-label") || ""),
    inputs_after: getResumeFileInputs().length
  };
}

function findResumeAttachButton() {
  const candidates = queryAll("button, [role='button'], label, a")
    .filter((element) => isVisible(element) && !isDisabledSubmitButton(element))
    .map((element) => ({
      element,
      text: normalize([element.textContent, element.getAttribute?.("aria-label"), element.getAttribute?.("title")].filter(Boolean).join(" ")),
      context: normalize(getResumeAttachButtonContext(element))
    }))
    .filter((candidate) => isResumeAttachButtonText(candidate.text) && isResumeAttachButtonContext(candidate.context))
    .map((candidate) => ({
      ...candidate,
      score: scoreResumeAttachButton(candidate.text, candidate.context)
    }))
    .sort((left, right) => right.score - left.score);

  return candidates[0]?.element || null;
}

function getResumeAttachButtonContext(element) {
  const roots = [
    element.closest?.("section, fieldset, label, [class*='resume' i], [class*='cv' i], [class*='upload' i], [class*='attachment' i], [class*='field' i], .form-group"),
    element.parentElement,
    element.closest?.("form")
  ].filter(Boolean);
  return roots.map((root) => cleanLabel(root.textContent || "")).join(" ");
}

function isResumeAttachButtonText(text) {
  if (!text) return false;
  if (/(dropbox|google drive|drive|manual|manually|paste|cover letter|portfolio|photo|avatar|image)/.test(text)) return false;
  return /^(attach|upload|browse|choose file|choose a file|select file|select a file|add file)$/.test(text) ||
    /\b(attach|upload|browse|choose (?:a )?file|select (?:a )?file|add file)\b/.test(text);
}

function isResumeAttachButtonContext(context) {
  if (!context) return false;
  return /\b(resume|cv|curriculum vitae)\b/.test(context) ||
    /\baccepted file types\b.*\b(pdf|doc|docx|rtf|txt|jpg|jpeg|png)\b/.test(context) ||
    isGenericApplicationFileUploadText(context);
}

function scoreResumeAttachButton(text, context) {
  let score = 0;
  if (/\b(resume|cv|curriculum vitae)\b/.test(context)) score += 120;
  if (/^attach$/.test(text)) score += 80;
  if (/^upload$|^upload file$/.test(text)) score += 70;
  if (/\b(pdf|doc|docx|rtf|txt)\b/.test(context)) score += 30;
  if (isGenericApplicationFileUploadText(context)) score += 35;
  return score;
}

function getResumeFileInputs() {
  const inputs = queryAll("input[type='file']")
    .filter((input) => !input.disabled);
  const teamtailorInputs = getTeamtailorResumeFileInputs(inputs);
  if (teamtailorInputs.length > 0) return teamtailorInputs;

  const scored = inputs
    .map((input) => ({
      input,
      score: scoreResumeFileInput(input),
      context: getFileInputContextText(input)
    }))
    .filter((candidate) => candidate.score > -100)
    .sort((left, right) => right.score - left.score);

  const requiredInputs = scored
    .filter((candidate) => candidate.input.required || candidate.input.getAttribute("aria-required") === "true")
    .filter((candidate) => candidate.score >= 90)
    .map((candidate) => candidate.input);
  if (requiredInputs.length === 1) return requiredInputs;

  const resumeInputs = scored
    .filter((candidate) => candidate.score >= 90)
    .map((candidate) => candidate.input);
  if (resumeInputs.length > 0) return resumeInputs;

  const nonCoverLetterInputs = scored
    .filter((candidate) => !isCoverLetterFileText(candidate.context))
    .map((candidate) => candidate.input);
  if (nonCoverLetterInputs.length === 1) return nonCoverLetterInputs;

  return getBambooResumeFileInputFallback(inputs);
}

function getTeamtailorResumeFileInputs(inputs) {
  const uploadRoots = queryAll("[data-controller~='forms--inputs--upload']");
  const resumeRoot = uploadRoots.find((element) => {
    const text = normalize([element.id, element.textContent].filter(Boolean).join(" "));
    const required = element.getAttribute("data-forms--inputs--upload-required-value") === "true";
    return required && /\b(resume|cv|curriculum vitae)\b/.test(text) && !isCoverLetterFileText(text);
  });
  if (!resumeRoot) return [];

  const directInput = queryAll("input[type='file']", resumeRoot).find((input) => !input.disabled);
  if (directInput) return [directInput];

  const dropzoneInputs = inputs.filter((input) => input.matches?.(".dz-hidden-input"));
  if (dropzoneInputs.length === 0) return [];

  const expectedAccept = normalizeFileAccept(
    resumeRoot.getAttribute("data-forms--inputs--upload-accepted-files-value") || ""
  );
  if (expectedAccept) {
    const acceptMatches = dropzoneInputs.filter((input) => normalizeFileAccept(input.accept || "") === expectedAccept);
    if (acceptMatches.length === 1) return acceptMatches;
  }

  const documentInputs = dropzoneInputs.filter(acceptsResumeDocument);
  if (documentInputs.length === 1) return documentInputs;

  const rootIndex = uploadRoots.indexOf(resumeRoot);
  if (rootIndex >= 0 && uploadRoots.length === dropzoneInputs.length && dropzoneInputs[rootIndex]) {
    return [dropzoneInputs[rootIndex]];
  }
  return [];
}

function normalizeFileAccept(value) {
  return String(value || "")
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean)
    .sort()
    .join(",");
}

function getBambooResumeFileInputFallback(inputs) {
  if (atsAdapters?.describe?.().id !== "bamboohr" || inputs.length === 0) return [];

  const uploadLabels = queryAll([
    ".fab-FormRow__label",
    ".fab-Label",
    "label",
    "legend",
    "[class*='label' i]"
  ].join(","))
    .filter(isVisible)
    .map((element) => ({ element, text: normalize(element.textContent || "") }))
    .filter(({ text }) => /^(cover letter|resume|cv|curriculum vitae)( required| optional)?$/.test(text))
    .filter((candidate, index, items) => items.findIndex((item) => item.element === candidate.element) === index);
  const resumeIndex = uploadLabels.findIndex(({ text }) => /^(resume|cv|curriculum vitae)/.test(text));
  if (resumeIndex < 0 || resumeIndex >= inputs.length || uploadLabels.length > inputs.length) return [];

  const input = inputs[resumeIndex];
  traceAutoBid("resume:bamboo-input-mapped", {
    input_index: resumeIndex,
    input_count: inputs.length,
    upload_labels: uploadLabels.map(({ text }) => text)
  });
  return [input];
}

function isResumeFileInput(input) {
  return scoreResumeFileInput(input) >= 90;
}

function scoreResumeFileInput(input) {
  const context = getFileInputContextText(input);
  const text = normalize(context);
  const uploadZone = getAssociatedResumeUploadZone(input);
  const uploadZoneText = normalize(uploadZone?.textContent || "");
  const attributes = normalize([
    input.name,
    input.id,
    input.getAttribute("aria-label"),
    input.getAttribute("title"),
    input.getAttribute("data-testid"),
    input.getAttribute("data-test"),
    input.getAttribute("class")
  ].join(" "));

  if (isCoverLetterFileText(text) || isCoverLetterFileText(attributes)) return -1000;

  let score = 0;
  if (atsAdapters?.describe?.().id === "bamboohr" &&
      (input.required || input.getAttribute("aria-required") === "true")) score += 125;
  if (uploadZone && !isCoverLetterFileText(uploadZoneText)) score += 110;
  if (/\b(resume|cv|curriculum vitae)\b/.test(text)) score += 120;
  if (/\b(resume|cv|curriculum vitae)\b/.test(attributes)) score += 100;
  if (/\b(upload|attach|attachment|file|drop|drag|browse|choose)\b/.test(text)) score += 30;
  if (/\b(pdf|doc|docx|rtf|jpg|jpeg|png)\b/.test(text) || acceptsResumeDocument(input)) score += 15;
  if (/\b(upload|attach|file)\b/.test(attributes)) score += 10;
  if (isGenericApplicationFileUploadText(text)) score += 75;
  if (isGenericApplicationFileUploadText(attributes)) score += 45;

  return score;
}

function isCoverLetterFileText(value) {
  return /(cover letter|motivation letter|supporting statement|portfolio|photo|avatar|image|transcript|certificate)/.test(normalize(value));
}

function acceptsResumeDocument(input) {
  const accept = normalize(input.accept || input.getAttribute("accept") || "");
  if (!accept) return false;
  return /(pdf|doc|docx|rtf|txt|text|msword|officedocument wordprocessingml document|jpeg|jpg|png|image)/.test(accept);
}

function getFileInputCandidatesDebug() {
  return queryAll("input[type='file']")
    .map((input) => ({
      score: scoreResumeFileInput(input),
      disabled: Boolean(input.disabled),
      has_file: Boolean(input.files?.length),
      attached: isResumeInputAttached(input),
      server_confirmed: hasServerConfirmedResumeUpload(input),
      component_files: getResumeUploadComponentFileMetadata(input),
      accept: input.accept || input.getAttribute("accept") || "",
      name: input.name || "",
      id: input.id || "",
      zone: shortText(getAssociatedResumeUploadZone(input)?.textContent || ""),
      context: shortText(getFileInputContextText(input))
    }));
}

function getFileInputContextText(input) {
  const rootNode = input.getRootNode?.() || document;
  const forLabel = input.id ? queryOne(`label[for="${cssEscape(input.id)}"]`, rootNode) : null;
  const container = getFileInputContextContainer(input);
  return cleanLabel([
    forLabel?.textContent,
    input.closest("label")?.textContent,
    getReferencedText(input, "aria-labelledby", rootNode),
    getReferencedText(input, "aria-describedby", rootNode),
    input.getAttribute("aria-label"),
    input.getAttribute("title"),
    input.getAttribute("accept"),
    input.getAttribute("data-testid"),
    input.getAttribute("data-test"),
    input.getAttribute("class"),
    input.name,
    input.id,
    container?.textContent,
    getFieldContainer(input)?.textContent,
    getAssociatedResumeUploadZone(input)?.textContent,
    getFileInputAncestorLabelCandidates(input).join(" "),
    getSiblingLabelCandidates(input).join(" ")
  ].filter(Boolean).join(" "));
}

function getResumeUploadZoneDebug() {
  return getResumeUploadZones().map((zone) => ({
    tag: zone.tagName?.toLowerCase?.() || "",
    text: shortText(zone.textContent || ""),
    inputs_inside: queryAll("input[type='file']", zone).length
  }));
}

function getResumeUploadZones() {
  const selector = [
    "label",
    "fieldset",
    "[role='button']",
    "[class*='upload' i]",
    "[class*='dropzone' i]",
    "[class*='drop-zone' i]",
    "[class*='attachment' i]",
    "[class*='resume' i]",
    "[class*='cv' i]",
    "[data-testid*='upload' i]",
    "[data-testid*='drop' i]",
    "[data-testid*='file' i]",
    "[data-test*='upload' i]",
    "[data-test*='drop' i]",
    "[data-test*='file' i]"
  ].join(",");

  return queryAll(selector)
    .filter((element) => isVisible(element))
    .filter((element) => isResumeUploadZoneText(element.textContent || element.getAttribute?.("aria-label") || ""));
}

function isResumeUploadZoneText(value) {
  const text = normalize(value);
  if (!text || isCoverLetterFileText(text)) return false;
  return (/\b(resume|cv|curriculum vitae)\b/.test(text) &&
    /\b(upload|attach|attachment|file|drop|drag|browse|choose)\b/.test(text)) ||
    isGenericApplicationFileUploadText(text);
}

function isGenericApplicationFileUploadText(value) {
  const text = normalize(value);
  if (!text || isCoverLetterFileText(text)) return false;
  if (/(dropbox|google drive|drive|manual|manually|paste|photo|avatar|image|portfolio|certificate|transcript)/.test(text)) return false;
  const hasGenericUpload = /\b(choose|select|upload|attach|browse|drop|drag)\b.*\bfile\b|\bfile\b.*\b(drop|upload|attach|browse|choose|select)\b/.test(text);
  if (!hasGenericUpload) return false;
  return /\b(easy apply|autocomplete your application|application|personal information|apply|mb size limit|size limit|pdf|doc|docx|rtf|txt)\b/.test(text);
}

function getAssociatedResumeUploadZone(input) {
  const direct = input.closest?.([
    "label",
    "fieldset",
    "[class*='upload' i]",
    "[class*='dropzone' i]",
    "[class*='drop-zone' i]",
    "[class*='attachment' i]",
    "[class*='resume' i]",
    "[class*='cv' i]",
    "[data-testid*='upload' i]",
    "[data-testid*='drop' i]",
    "[data-testid*='file' i]"
  ].join(","));
  if (direct && isResumeUploadZoneText(direct.textContent || direct.getAttribute?.("aria-label") || "")) return direct;

  const zones = getResumeUploadZones();
  if (zones.length === 0) return null;

  const inputForm = input.closest?.("form");
  const inputSection = input.closest?.("section, fieldset, main, [role='main'], article, [class*='application' i], [class*='form' i]");
  const candidates = zones
    .map((zone) => ({
      zone,
      score: scoreResumeUploadZoneAssociation(input, zone, inputForm, inputSection)
    }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score);

  return candidates[0]?.zone || null;
}

function scoreResumeUploadZoneAssociation(input, zone, inputForm, inputSection) {
  if (zone.contains(input)) return 1000;
  let score = 0;
  const zoneForm = zone.closest?.("form");
  const zoneSection = zone.closest?.("section, fieldset, main, [role='main'], article, [class*='application' i], [class*='form' i]");
  if (inputForm && zoneForm && inputForm === zoneForm) score += 180;
  if (inputSection && zoneSection && inputSection === zoneSection) score += 80;
  if (input.parentElement && zone.parentElement && input.parentElement === zone.parentElement) score += 140;

  const zoneInputs = queryAll("input[type='file']", zone);
  if (zoneInputs.includes(input)) score += 400;

  const sameFormInputs = inputForm ? queryAll("input[type='file']", inputForm).filter((item) => !item.disabled) : [];
  const sameFormZones = inputForm ? getResumeUploadZones().filter((item) => item.closest?.("form") === inputForm) : [];
  if (sameFormInputs.length === 1 && sameFormZones.length === 1) score += 120;

  return score;
}

function getFileInputContextContainer(input) {
  const direct = input.closest([
    "label",
    "fieldset",
    "[class*='upload' i]",
    "[class*='dropzone' i]",
    "[class*='drop-zone' i]",
    "[class*='file' i]",
    "[class*='resume' i]",
    "[class*='cv' i]",
    "[class*='field' i]",
    "[data-testid*='upload' i]",
    "[data-testid*='drop' i]",
    "[data-testid*='file' i]",
    ".form-group"
  ].join(","));
  if (direct) return direct;

  let current = input.parentElement;
  for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
    const text = cleanLabel(current.textContent || "");
    if (text && text.length < 1800 && /(resume|cv|curriculum vitae|upload|attach|drop|drag|browse|choose file|file)/.test(normalize(text))) {
      return current;
    }
  }

  return input.parentElement;
}

function getReferencedText(element, attribute, rootNode = document) {
  return String(element.getAttribute(attribute) || "")
    .split(/\s+/)
    .map((id) => id.trim())
    .filter(Boolean)
    .map((id) => rootNode.getElementById?.(id)?.textContent || document.getElementById(id)?.textContent || "")
    .join(" ");
}

function getFileInputAncestorLabelCandidates(input) {
  const candidates = [];
  let current = input;

  for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
    const label = current.querySelector?.(FIELD_LABEL_SELECTOR);
    if (label && label !== input && !label.contains(input)) candidates.push(label.textContent);

    let sibling = current.previousElementSibling;
    for (let index = 0; sibling && index < 4; index += 1, sibling = sibling.previousElementSibling) {
      const text = cleanLabel(sibling.textContent || "");
      if (text && text.length < 700 && (isLabelLikeElement(sibling) || /(resume|cv|curriculum vitae|upload|attach|file)/.test(normalize(text)))) {
        candidates.push(text);
      }
    }
  }

  return candidates;
}

function fileFromBase64(payload) {
  const filename = sanitizeFilename(payload.filename || payload.name || "resume.pdf");
  const mimeType = String(payload.mime_type || payload.mimeType || "application/pdf").trim() || "application/pdf";
  const binary = atob(String(payload.base64 || "").replace(/^data:[^,]+,/, ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new File([bytes], filename, { type: mimeType });
}

function sanitizeFilename(value) {
  return String(value || "resume.pdf").replace(/[\\/:*?"<>|]+/g, "_").trim() || "resume.pdf";
}

async function setFileInputValue(input, file, payload = {}) {
  try {
    const pageResult = await setFileInputValueInPage(input, file, payload);
    if (pageResult.attempted) {
      traceAutoBid("resume:page-upload", {
        attached: pageResult.attached,
        server_confirmed: pageResult.server_confirmed,
        managed_status: pageResult.managed_status || "",
        filename: file.name
      });
    }
    if (pageResult.attached) return true;
    // Once a page accepted a file-selection event, retrying through CDP, drop,
    // or a file chooser can enqueue the same File again. A rejected/timed-out
    // managed upload is a real failure, not permission to submit a duplicate.
    if (pageResult.attempted) return false;

    const nativeResult = await setFileInputValueNatively(input, file, payload);
    if (nativeResult.attempted) {
      traceAutoBid("resume:native-upload", {
        uploaded: nativeResult.uploaded,
        attached: nativeResult.attached,
        server_confirmed: nativeResult.server_confirmed,
        managed_status: nativeResult.managed_status || "",
        files: nativeResult.files,
        reason: nativeResult.reason || "",
        filename: file.name,
        transport: "bytes"
      });
    }
    if (nativeResult.attached) return true;
    if (nativeResult.attempted) return false;

    const confirmationBefore = getResumeUploadComponentFileMetadata(input);
    const uploadContext = getManagedResumeUploadContext(input);
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    dispatchFileInputEvents(input, transfer);
    return waitForResumeInputConfirmation(
      input,
      file.name,
      RESUME_SERVER_CONFIRM_TIMEOUT_MS,
      confirmationBefore,
      uploadContext
    );
  } catch (error) {
    traceAutoBid("resume:set-file-error", { message: error.message || String(error) });
    return false;
  }
}

async function setFileInputValueNatively(input, file, payload) {
  const base64 = String(payload.base64 || "").replace(/^data:[^,]+,/, "");
  if (!base64) return { attempted: false, attached: false, reason: "no-file-bytes" };

  const token = `ab_file_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const confirmationBefore = getResumeUploadComponentFileMetadata(input);
  const uploadContext = getManagedResumeUploadContext(input);
  input.setAttribute("data-auto-bid-native-file-token", token);
  try {
    const result = await send("NATIVE_FILE_UPLOAD", {
      token,
      filename: file.name,
      mime_type: file.type || payload.mime_type || payload.mimeType || "application/pdf",
      base64
    });
    const attached = await waitForResumeInputConfirmation(
      input,
      file.name,
      RESUME_SERVER_CONFIRM_TIMEOUT_MS,
      confirmationBefore,
      uploadContext
    );
    const managedStatus = getManagedResumeUploadStatus(uploadContext, file.name);
    return {
      attempted: true,
      uploaded: Boolean(result?.uploaded),
      files: Number(result?.files || input.files?.length || 0),
      attached,
      server_confirmed: attached && isServerTrackedResumeInput(input),
      managed_status: managedStatus.status
    };
  } catch (error) {
    return {
      // With debugger-free operation, allow the existing in-page DataTransfer
      // fallback to run instead of treating an unavailable native transport as
      // a completed upload attempt.
      attempted: false,
      attached: false,
      uploaded: false,
      reason: error.message || String(error)
    };
  } finally {
    window.setTimeout(() => {
      if (input.getAttribute("data-auto-bid-native-file-token") === token) {
        input.removeAttribute("data-auto-bid-native-file-token");
      }
    }, 1200);
  }
}

async function setFileInputValueInPage(input, file, payload) {
  const filePayload = {
    name: file.name,
    mime_type: file.type || payload.mime_type || payload.mimeType || "application/pdf",
    base64: String(payload.base64 || "").replace(/^data:[^,]+,/, "")
  };
  if (!filePayload.base64) return { attempted: false, attached: false };

  const confirmationBefore = getResumeUploadComponentFileMetadata(input);
  const uploadContext = getManagedResumeUploadContext(input);
  const acknowledged = runPageCommand("file-upload", input, { file: filePayload });
  if (!acknowledged) return { attempted: false, attached: false };

  const attached = await waitForResumeInputConfirmation(
    input,
    file.name,
    RESUME_SERVER_CONFIRM_TIMEOUT_MS,
    confirmationBefore,
    uploadContext
  );
  const managedStatus = getManagedResumeUploadStatus(uploadContext, file.name);
  return {
    attempted: true,
    attached,
    server_confirmed: attached && isServerTrackedResumeInput(input),
    managed_status: managedStatus.status
  };
}

function dispatchFileInputEvents(input, transfer) {
  input.focus?.();
  const eventTypes = getManagedResumeUploadContext(input).managed ? ["change"] : ["input", "change"];
  eventTypes.forEach((type) => {
    const event = new Event(type, { bubbles: true, cancelable: true, composed: true });
    defineEventDataTransfer(event, transfer);
    input.dispatchEvent(event);
  });
  input.blur?.();
}

function defineEventDataTransfer(event, transfer) {
  try {
    Object.defineProperty(event, "dataTransfer", { value: transfer });
  } catch {
    // Some browser events already expose dataTransfer.
  }
}

function hasUploadedFileUi(input, filename) {
  if (hasServerConfirmedResumeUpload(input, filename)) return true;
  const text = normalize([
    getFileInputContextText(input),
    getAssociatedResumeUploadZone(input)?.textContent
  ].filter(Boolean).join(" "));
  const name = normalize(filename);
  return Boolean(name && text.includes(name));
}

function hasResumeUploadedUi(filename) {
  const name = normalize(filename);
  if (!name) return false;

  if (getResumeFileInputs().some((input) => hasServerConfirmedResumeUpload(input, filename))) return true;

  const zones = getResumeUploadZones();
  if (zones.some((zone) => normalize(zone.textContent || "").includes(name))) return true;

  const form = queryOne("form") || document.body;
  const text = normalize(form?.textContent || "");
  return text.includes(name);
}

function isResumeInputAttached(input) {
  if (!input) return false;
  if (hasServerConfirmedResumeUpload(input)) return true;
  const managedStatus = getManagedResumeUploadStatus(getManagedResumeUploadContext(input), input.files?.[0]?.name || "");
  if (managedStatus.managed) return managedStatus.status === "complete";
  if (isServerTrackedResumeInput(input)) return false;
  return Boolean(input.files?.length) || hasUploadedFileUi(input, input.files?.[0]?.name || "");
}

async function waitForResumeInputConfirmation(input, filename, timeoutMs, confirmationBefore = [], uploadContext = null) {
  const requiresServerConfirmation = isServerTrackedResumeInput(input);
  const managedContext = uploadContext || getManagedResumeUploadContext(input);
  const waitTimeoutMs = managedContext.managed
    ? Math.max(timeoutMs, RESUME_MANAGED_UPLOAD_TIMEOUT_MS)
    : timeoutMs;
  const started = Date.now();

  while (Date.now() - started < waitTimeoutMs) {
    if (hasNewServerConfirmedResumeUpload(input, filename, confirmationBefore)) return true;
    const managedStatus = getManagedResumeUploadStatus(managedContext, filename);
    if (managedStatus.status === "complete") return true;
    if (managedStatus.status === "failed") {
      traceAutoBid("resume:managed-upload-failed", {
        filename,
        message: managedStatus.message
      });
      return false;
    }
    if (!requiresServerConfirmation &&
        !managedContext.managed &&
        (Boolean(input.files?.length) || hasUploadedFileUi(input, filename))) {
      return true;
    }
    await sleep(200);
  }

  return hasNewServerConfirmedResumeUpload(input, filename, confirmationBefore) ||
    !requiresServerConfirmation &&
      !managedContext.managed &&
      (Boolean(input.files?.length) || hasUploadedFileUi(input, filename));
}

async function waitForAnyResumeUploadConfirmation(filename, nativeUploaded, timeoutMs, confirmationBefore = []) {
  const started = Date.now();
  let inputs = getResumeFileInputs();
  const serverTracked = inputs.some(isServerTrackedResumeInput) || hasServerTrackedResumeUploadZone();
  const managedContext = getManagedResumeUploadContext(inputs[0] || null);
  const waitTimeoutMs = managedContext.managed
    ? Math.max(timeoutMs, RESUME_MANAGED_UPLOAD_TIMEOUT_MS)
    : timeoutMs;

  while (Date.now() - started < waitTimeoutMs) {
    inputs = getResumeFileInputs();
    if (hasNewResumeServerConfirmation(filename, confirmationBefore)) {
      return true;
    }
    const managedStatus = getManagedResumeUploadStatus(managedContext, filename);
    if (managedStatus.status === "complete") return true;
    if (managedStatus.status === "failed") return false;
    if (!managedContext.managed && !serverTracked && inputs.some((input) => isResumeInputAttached(input))) return true;
    if (!managedContext.managed && !serverTracked && (nativeUploaded || hasResumeUploadedUi(filename))) return true;
    await sleep(200);
  }

  return hasNewResumeServerConfirmation(filename, confirmationBefore) ||
    !managedContext.managed && !serverTracked && inputs.some(isResumeInputAttached) ||
    !managedContext.managed && !serverTracked && (nativeUploaded || hasResumeUploadedUi(filename));
}

function getManagedResumeUploadContext(input) {
  const root = findTeamtailorResumeUploadRoot(input) ||
    input?.closest?.(".dropzone, [data-controller~='forms--inputs--upload']") ||
    null;
  const managed = Boolean(
    root ||
    input?.matches?.(".dz-hidden-input") ||
    atsAdapters?.describe?.().id === "teamtailor"
  );
  return { managed, root };
}

function findTeamtailorResumeUploadRoot(input) {
  const direct = input?.closest?.("[data-controller~='forms--inputs--upload']");
  if (direct && isResumeUploadZoneText(direct.textContent || direct.id || "")) return direct;

  return queryAll("[data-controller~='forms--inputs--upload']")
    .find((element) => {
      const text = normalize([element.id, element.textContent].filter(Boolean).join(" "));
      return /\b(resume|cv|curriculum vitae)\b/.test(text) && !isCoverLetterFileText(text);
    }) || null;
}

function getManagedResumeUploadStatus(context, filename = "") {
  if (!context?.managed) return { managed: false, status: "unmanaged", message: "" };
  const root = context.root?.isConnected ? context.root : findTeamtailorResumeUploadRoot(null);
  if (!root) return { managed: true, status: "pending", message: "" };

  const visibleErrors = queryAll(".dz-error, [data-dz-errormessage], [role='alert']", root)
    .filter(isVisible)
    .map((element) => cleanLabel(element.textContent || element.getAttribute?.("data-dz-errormessage") || ""))
    .filter(Boolean);
  const rootText = cleanLabel(root.textContent || "");
  const duplicateQueueMessage = /already been processed or was rejected/i.test(rootText);
  if (visibleErrors.length > 0 || duplicateQueueMessage) {
    return {
      managed: true,
      status: "failed",
      message: shortText(visibleErrors.join(" ") || rootText)
    };
  }

  const urlInputs = queryAll("[data-forms--inputs--upload-preview-target='urlInput']", root);
  const hasRemoteUrl = urlInputs.some((element) => !element.disabled && Boolean(String(element.value || "").trim()));
  const visibleName = queryAll("[data-forms--inputs--upload-preview-target='name'], [data-dz-name]", root)
    .filter(isVisible)
    .some((element) => {
      const text = normalize(element.textContent || "");
      return Boolean(text) && (!filename || text.includes(normalize(filename)));
    });
  const dropzoneSucceeded = Boolean(root.querySelector?.(".dz-success:not(.dz-error)"));
  if (hasRemoteUrl || visibleName || dropzoneSucceeded) {
    return { managed: true, status: "complete", message: "" };
  }

  const visibleProgress = queryAll("[data-progress], .dz-progress, [data-dz-uploadprogress]", root)
    .some(isVisible);
  const dropzoneProcessing = Boolean(root.querySelector?.(".dz-processing, .dz-uploading, .dz-queued"));
  if (visibleProgress || dropzoneProcessing || /\buploading\b/i.test(rootText)) {
    return { managed: true, status: "uploading", message: "" };
  }

  return { managed: true, status: "pending", message: "" };
}

function hasServerTrackedResumeUploadZone() {
  return queryAll("spl-dropzone, [data-test*='apply-with-resume' i], [data-testid*='apply-with-resume' i]")
    .some((element) => !isCoverLetterFileText(element.textContent || ""));
}

function isServerTrackedResumeInput(input) {
  return getResumeUploadComponentElements(input).some((element) =>
    normalize(element.tagName || "") === "spl dropzone" ||
    /apply-with-resume/.test(normalize([
      element.getAttribute?.("data-test"),
      element.getAttribute?.("data-testid"),
      element.tagName
    ].filter(Boolean).join(" ")))
  );
}

function hasServerConfirmedResumeUpload(input, filename = "") {
  const expectedName = normalize(filename);
  return getResumeUploadComponentFileMetadata(input).some((metadata) => {
    const text = normalize(metadata);
    if (expectedName && !text.includes(expectedName)) return false;
    return /fileid|file id|attachmenttype|attachment type|uploaded/.test(text);
  });
}

function hasNewServerConfirmedResumeUpload(input, filename, confirmationBefore = []) {
  const previous = new Set((confirmationBefore || []).map(String));
  const expectedName = normalize(filename);
  return getResumeUploadComponentFileMetadata(input).some((metadata) => {
    if (previous.has(String(metadata))) return false;
    const text = normalize(metadata);
    if (expectedName && !text.includes(expectedName)) return false;
    return /fileid|file id|attachmenttype|attachment type|uploaded/.test(text);
  });
}

function hasNewResumeServerConfirmation(filename, confirmationBefore = []) {
  const previous = new Set((confirmationBefore || []).map(String));
  const expectedName = normalize(filename);
  return getResumeServerConfirmationSnapshot().some((metadata) => {
    if (previous.has(String(metadata))) return false;
    const text = normalize(metadata);
    if (expectedName && !text.includes(expectedName)) return false;
    return /fileid|file id|attachmenttype|attachment type|uploaded/.test(text);
  });
}

function getResumeServerConfirmationSnapshot() {
  const metadata = queryAll("input[type='file']")
    .flatMap(getResumeUploadComponentFileMetadata);
  queryAll("spl-dropzone, [data-test*='apply-with-resume' i], [data-testid*='apply-with-resume' i]")
    .forEach((element) => {
      ["files", "data-files", "attachments", "data-attachments"].forEach((attribute) => {
        const value = element.getAttribute?.(attribute);
        if (value) metadata.push(String(value));
      });
    });
  return uniqueNonEmptyValues(metadata);
}

function getResumeInputAttachedFilenames(input) {
  const names = [];
  if (input?.files?.length) {
    names.push(...Array.from(input.files).map((file) => file.name).filter(Boolean));
  }

  for (const metadata of getResumeUploadComponentFileMetadata(input)) {
    try {
      const parsed = JSON.parse(metadata);
      const items = Array.isArray(parsed) ? parsed : [parsed];
      items.forEach((item) => {
        const name = item?.fileName || item?.filename || item?.name || "";
        if (name) names.push(String(name));
      });
    } catch (_error) {
      const match = metadata.match(/(?:fileName|filename|name)["']?\s*[:=]\s*["']([^"']+)/i);
      if (match?.[1]) names.push(match[1]);
    }
  }

  return uniqueNonEmptyValues(names);
}

function getResumeUploadComponentFileMetadata(input) {
  const values = [];
  for (const element of getResumeUploadComponentElements(input)) {
    ["files", "data-files", "attachments", "data-attachments"].forEach((attribute) => {
      const value = element.getAttribute?.(attribute);
      if (value) values.push(String(value));
    });

    const propertyValue = element !== input ? element.files || element.attachments : null;
    if (propertyValue && !(propertyValue instanceof FileList)) {
      try {
        values.push(typeof propertyValue === "string" ? propertyValue : JSON.stringify(propertyValue));
      } catch (_error) {
        // Attribute metadata is sufficient when a component property is not serializable.
      }
    }
  }
  return uniqueNonEmptyValues(values);
}

function getResumeUploadComponentElements(input) {
  const elements = [];
  let current = input;
  for (let depth = 0; current && depth < 14; depth += 1) {
    if (!elements.includes(current)) elements.push(current);
    if (current.parentElement) {
      current = current.parentElement;
      continue;
    }
    const root = current.getRootNode?.();
    current = root?.host || null;
  }
  return elements;
}

function summarizeResumePayload(payload) {
  if (!payload) return null;
  return {
    row_number: payload.row_number || payload.rowNumber || null,
    filename: payload.filename || payload.name || "",
    mime_type: payload.mime_type || payload.mimeType || "",
    size: payload.size || payload.bytes || 0,
    source_url: payload.source_url || payload.url || "",
    local_path: payload.local_path || payload.localPath || "",
    refreshed: Boolean(payload.refreshed),
    reason: payload.reason || ""
  };
}

async function applySheetAnswerExchange(fields, filledIds) {
  const localFilledIds = new Set();
  const page = collectPageContext();
  let context = null;

  try {
    context = await send("SHEET_CONTEXT", { page, url: page.url });
  } catch (error) {
    traceAutoBid("sheet:context-error", { message: error.message || String(error) });
    return { filled: 0, missed: 0, pending: 0, filledIds: localFilledIds };
  }

  if (!context?.rowNumber) {
    traceAutoBid("sheet:no-context", { url: page.url });
    showStatus("No matching Google Sheet row found for this page.", "error");
    return { filled: 0, missed: 0, pending: 0, filledIds: localFilledIds };
  }

  traceAutoBid("sheet:context-found", {
    row_number: context.rowNumber,
    sheet_name: context.sheetName || "",
    url: context.url || page.url,
    match_source: context.matchSource || ""
  });

  const candidateFields = getSheetQuestionCandidateFields(fields, filledIds);
  if (candidateFields.length === 0) {
    traceAutoBid("sheet:no-candidates", {
      row_number: context.rowNumber,
      fields: fields.map((field) => ({
        field_id: field.id,
        label: field.label,
        type: field.type,
        current: shortText(getCurrentChoiceSummary(getControlsByFieldId(field.id)) || field.value || ""),
        reason: getSheetCandidateSkipReason(field, filledIds)
      }))
    });
    showStatus("No unanswered generated-answer fields found for this page.", "success");
    return { filled: 0, missed: 0, pending: 0, filledIds: localFilledIds };
  }

  traceAutoBid("sheet:candidates", {
    row_number: context.rowNumber,
    count: candidateFields.length,
    fields: candidateFields.map((field) => ({
      field_id: field.id,
      label: field.label,
      type: field.type
    }))
  });

  const candidateIds = new Set(candidateFields.map((field) => field.id));
  let totalFilled = 0;
  let totalMissed = 0;
  const activeFilledIds = new Set(filledIds);

  const applySheetAnswers = async (data, stage) => {
    const answers = normalizeSheetAnswers(data)
      .filter((answer) => candidateIds.has(answer.field_id))
      .filter((answer) => !hasFieldCurrentValue(fields.find((field) => field.id === answer.field_id)));
    if (answers.length === 0) return { filled: 0, missed: 0, filledIds: new Set() };

    traceAutoBid("sheet:answers-received", {
      stage,
      row_number: context.rowNumber,
      answers: answers.map((answer) => ({
        field_id: answer.field_id,
        value: shortText(answer.value)
      }))
    });

    const result = await applyAnswers(answers, activeFilledIds, fields);
    result.filledIds.forEach((fieldId) => {
      localFilledIds.add(fieldId);
      activeFilledIds.add(fieldId);
    });
    totalFilled += result.filled;
    totalMissed += result.missed;
    return result;
  };

  let remaining = candidateFields.filter((field) => !activeFilledIds.has(field.id) && !hasFieldCurrentValue(field));
  if (remaining.length === 0) {
    return { filled: totalFilled, missed: totalMissed, pending: 0, filledIds: localFilledIds };
  }

  const profileContext = await getGeneratedAnswerProfileContext();
  const payload = buildSheetQuestionPayload(page, context, remaining, profileContext);
  try {
    await send("SHEET_SUBMIT_QUESTIONS", { context, page, payload });
    traceAutoBid("sheet:questions-sent", {
      row_number: context.rowNumber,
      fields: remaining.map((field) => ({
        field_id: field.id,
        question: field.question || field.label,
        option: field.option || "",
        type: field.type,
        options: field.options || []
      }))
    });
    showStatus(`Sent ${remaining.length} question${remaining.length === 1 ? "" : "s"} to Google Sheet`, "success");
    triggerSheetGptAnswerWorker(context).catch((error) => {
      traceAutoBid("sheet:gpt-worker-trigger-error", { message: error.message || String(error) });
    });
  } catch (error) {
    traceAutoBid("sheet:submit-error", { message: error.message || String(error) });
    return { filled: totalFilled, missed: totalMissed, pending: remaining.length, filledIds: localFilledIds };
  }

  await sleep(SHEET_ANSWER_FIRST_WAIT_MS);

  for (let attempt = 1; attempt <= SHEET_ANSWER_RETRY_ATTEMPTS; attempt += 1) {

    try {
      const data = await send("SHEET_FETCH_ANSWERS", { context, page });
      await applySheetAnswers(data, attempt === 1 ? "after-wait" : `retry-${attempt - 1}`);
      remaining = candidateFields.filter((field) => !activeFilledIds.has(field.id) && !hasFieldCurrentValue(field));
      traceAutoBid("sheet:read", {
        row_number: context.rowNumber,
        attempt,
        remaining: remaining.length
      });
      if (remaining.length === 0) break;
    } catch (error) {
      traceAutoBid("sheet:fetch-error", { stage: `read-${attempt}`, message: error.message || String(error) });
    }

    if (attempt < SHEET_ANSWER_RETRY_ATTEMPTS) await sleep(SHEET_ANSWER_RETRY_MS);
  }

  if (remaining.length > 0) {
    traceAutoBid("sheet:answers-pending", {
      row_number: context.rowNumber,
      remaining: remaining.map((field) => ({
        field_id: field.id,
        label: field.label,
        type: field.type
      }))
    });
    showStatus(`${remaining.length} Sheet answer${remaining.length === 1 ? "" : "s"} still pending.`, "success");
  }

  return { filled: totalFilled, missed: totalMissed, pending: remaining.length, filledIds: localFilledIds };
}

async function applyRuntimeGptAnswerExchange(fields, filledIds) {
  const localFilledIds = new Set();
  const page = collectPageContext();
  const context = await getSheetContextForRuntime(page);

  const candidateFields = getSheetQuestionCandidateFields(fields, filledIds);
  traceAutoBid("runtime-gpt:candidates-selected", {
    candidates: candidateFields.map((field) => ({
      field_id: field.id,
      question: field.question || field.label,
      option: field.option || "",
      type: field.type
    })),
    skipped: fields
      .map((field) => ({
        field_id: field.id,
        question: field.question || field.label,
        option: field.option || "",
        type: field.type,
        reason: getSheetCandidateSkipReason(field, filledIds)
      }))
      .filter((field) => field.reason)
  });
  if (candidateFields.length === 0) {
    traceAutoBid("runtime-gpt:no-candidates", {
      fields: fields.map((field) => ({
        field_id: field.id,
        label: field.label,
        type: field.type,
        current: shortText(getCurrentChoiceSummary(getControlsByFieldId(field.id)) || field.value || ""),
        reason: getSheetCandidateSkipReason(field, filledIds)
      }))
    });
    return { filled: 0, missed: 0, pending: 0, filledIds: localFilledIds };
  }

  showStatus("Preparing GPT questions", "working", {
    detail: "Reading candidate values from required selects, radios, and dropdowns…"
  });
  await hydrateGeneratedChoiceOptions(candidateFields);
  const unresolvedCandidateFields = candidateFields.filter((field) => !hasFieldCurrentValue(field));
  const incompleteChoiceFields = unresolvedCandidateFields.filter((field) => !hasCompleteGeneratedChoiceOptions(field));
  const requestableFields = unresolvedCandidateFields.filter((field) => hasCompleteGeneratedChoiceOptions(field));
  if (incompleteChoiceFields.length > 0) {
    traceAutoBid("runtime-gpt:choice-options-unavailable", {
      fields: incompleteChoiceFields.map((field) => ({
        field_id: field.id,
        question: field.question || field.label,
        type: field.type
      }))
    });
  }
  if (requestableFields.length === 0) {
    return {
      filled: 0,
      missed: incompleteChoiceFields.length,
      pending: 0,
      filledIds: localFilledIds,
      missedIds: new Set(incompleteChoiceFields.map((field) => field.id))
    };
  }

  const candidateIds = new Set(requestableFields.map((field) => field.id));
  const activeFilledIds = new Set(filledIds);
  let remaining = requestableFields.filter((field) => !activeFilledIds.has(field.id) && !hasFieldCurrentValue(field));
  if (remaining.length === 0) {
    return { filled: 0, missed: incompleteChoiceFields.length, pending: 0, filledIds: localFilledIds };
  }

  const profileContext = await getGeneratedAnswerProfileContext();
  const payload = buildSheetQuestionPayload(page, context, remaining, profileContext);
  traceAutoBid("runtime-gpt:context-quality", {
    row_number: context.rowNumber || null,
    match_source: context.matchSource || "",
    tailored_resume_chars: String(payload.gpt_context?.tailored_resume_content || "").length,
    job_description_chars: String(payload.gpt_context?.job_description || "").length,
    profile_static_keys: Object.keys(payload.profile?.static_fields || {}).filter((key) => String(payload.profile.static_fields[key] || "").trim())
  });
  traceAutoBid("runtime-gpt:request", {
    row_number: context.rowNumber || null,
    sheet_name: context.sheetName || "",
    fields: remaining.map((field) => ({
      field_id: field.id,
      question: field.question || field.label,
      option: field.option || "",
      type: field.type,
      options: field.options || []
    }))
  });

  let request;
  try {
    request = await send("GPT_ANSWER_REQUEST", {
      context,
      page,
      payload,
      client_run_id: autoBidRunId,
      timeout_ms: RUNTIME_GPT_ANSWER_TIMEOUT_MS,
      max_attempts: 1
    });
  } catch (error) {
    traceAutoBid("runtime-gpt:start-error", { message: error.message || String(error) });
    return { filled: 0, missed: incompleteChoiceFields.length, pending: remaining.length, filledIds: localFilledIds };
  }

  const requestId = request?.request_id || request?.requestId || "";
  if (requestId) runtimeGptSourceFields.set(requestId, remaining.map((field) => ({ ...field })));
  showStatus("Questions sent to GPT", "waiting", {
    detail: `Waiting for ${remaining.length} required answer${remaining.length === 1 ? "" : "s"}…`
  });

  const status = await waitForRuntimeGptAnswer(requestId);
  if (!isActiveContentInstance()) {
    return { filled: 0, missed: 0, pending: 0, request_id: requestId, filledIds: localFilledIds, ignored: true, reason: "superseded-run" };
  }
  if (status?.status === "error") {
    traceAutoBid("runtime-gpt:error", {
      request_id: status.request_id || status.requestId || "",
      message: status.error || ""
    });
    return { filled: 0, missed: remaining.length + incompleteChoiceFields.length, pending: 0, request_id: requestId, filledIds: localFilledIds };
  }

  if (status?.status !== "complete") {
    traceAutoBid("runtime-gpt:pending", {
      request_id: requestId,
      status: status?.status || "timeout",
      remaining: remaining.length
    });
    return { filled: 0, missed: incompleteChoiceFields.length, pending: remaining.length, request_id: requestId, filledIds: localFilledIds };
  }

  const answers = normalizeRuntimeGptAnswers(status)
    .filter((answer) => candidateIds.has(answer.field_id));

  if (answers.length === 0) {
    traceAutoBid("runtime-gpt:no-usable-answers", {
      request_id: status.request_id || status.requestId || "",
      answers: Array.isArray(status.answers) ? status.answers.length : 0
    });
    return { filled: 0, missed: remaining.length + incompleteChoiceFields.length, pending: 0, filledIds: localFilledIds };
  }

  traceAutoBid("runtime-gpt:answers-received", {
    request_id: status.request_id || status.requestId || "",
    answers: answers.map((answer) => ({
      field_id: answer.field_id,
      value: shortText(answer.value)
    }))
  });

  showStatus("Autofilling GPT answers", "autofilling", {
    detail: `${answers.length} answer${answers.length === 1 ? " was" : "s were"} received and will be applied once…`
  });
  const result = await applyRuntimeGptAnswersOnce(requestId, answers, remaining, activeFilledIds);
  result.filledIds.forEach((fieldId) => localFilledIds.add(fieldId));
  remaining = requestableFields.filter((field) => !result.filledIds.has(field.id) && !hasFieldCurrentValue(field));

  return {
    filled: result.filled,
    missed: result.missed + incompleteChoiceFields.length,
    pending: 0,
    remaining: remaining.length,
    request_id: requestId,
    filledIds: localFilledIds
  };
}

async function applyPushedRuntimeGptAnswers(payload) {
  const requestId = String(payload.request_id || payload.requestId || "");
  if (!requestId) return emptyFillResult();
  const clientRunId = String(payload.client_run_id || payload.clientRunId || "");
  const knownToActiveRun = runtimeGptSourceFields.has(requestId);
  if ((clientRunId && clientRunId !== autoBidRunId) || (!clientRunId && !knownToActiveRun)) {
    traceAutoBid("runtime-gpt:push-ignored-stale-run", {
      request_id: requestId,
      request_run_id: clientRunId,
      active_run_id: autoBidRunId,
      legacy_untagged: !clientRunId
    });
    flushTrace();
    return { ...emptyFillResult(), applied: false, settled: true, ignored: true, reason: "stale-run" };
  }

  if (finalizedRuntimeGptRequestIds.has(requestId)) {
    traceAutoBid("runtime-gpt:push-ignored-finalized", { request_id: requestId });
    flushTrace();
    return { ...emptyFillResult(), applied: false, settled: true, ignored: true, reason: "already-finalized" };
  }

  if (!runtimeGptSourceFields.has(requestId) && Array.isArray(payload.fields)) {
    runtimeGptSourceFields.set(requestId, payload.fields.map((field) => ({
      ...field,
      id: String(field.id || field.field_id || "")
    })).filter((field) => field.id));
  }

  const answers = normalizeRuntimeGptAnswers({ answers: payload.answers || [] });
  traceAutoBid("runtime-gpt:push-received", {
    request_id: requestId,
    answers: answers.map((answer) => ({
      field_id: answer.field_id,
      value: shortText(answer.value)
    }))
  });
  if (answers.length === 0) return emptyFillResult();

  showStatus("Autofilling GPT answers", "autofilling", {
    detail: `${answers.length} answer${answers.length === 1 ? " was" : "s were"} received and will be applied once…`
  });
  const result = await applyRuntimeGptAnswersOnce(requestId, answers);
  const settlement = getRuntimeGptAnswerSettlement(requestId);
  traceAutoBid("runtime-gpt:push-applied", {
    request_id: requestId,
    filled: result.filled,
    missed: result.missed,
    settlement
  });
  if (settlement.settled) {
    await acknowledgeRuntimeGptAnswersApplied(requestId, {
      ...result,
      exhausted: settlement.exhausted
    }).catch(() => {});
  }
  flushTrace();
  if (!autoBidRunning) {
    showStatus(
      result.missed > 0 ? "Autofill finished" : "Autofill done",
      result.missed > 0 ? "warning" : "success",
      {
        detail: result.missed > 0
          ? `${result.missed} required field${result.missed === 1 ? " is" : "s are"} still empty. Press Ctrl+Q to retry.`
          : "The GPT answers were applied once. Manual changes will be preserved."
      }
    );
  }
  return {
    filled: Number(result.filled || 0),
    missed: Number(result.missed || 0),
    applied: result.filled > 0,
    settled: settlement.settled,
    exhausted: settlement.exhausted,
    pending: settlement.pending
  };
}

function getRuntimeGptAnswerSettlement(requestId) {
  const answers = runtimeGptAnswersByRequest.get(requestId) || [];
  const sourceFields = runtimeGptSourceFields.get(requestId) || [];
  if (finalizedRuntimeGptRequestIds.has(requestId)) {
    return {
      settled: true,
      filled: 0,
      exhausted: 0,
      pending: 0,
      max_attempts: 1
    };
  }
  const sourceById = new Map(sourceFields.map((field) => [field.id, field]));
  let filled = 0;
  let exhausted = 0;
  let pending = 0;

  for (const answer of answers) {
    const sourceField = sourceById.get(answer.field_id) || null;
    const binding = resolveLiveFieldForAnswer(answer, sourceField);
    const field = binding?.field || sourceField;
    if (binding && doesGeneratedAnswerMatchField(answer, binding.field, binding.controls)) {
      filled += 1;
      continue;
    }

    const attempts = getGeneratedAnswerFillAttempts(getGeneratedAnswerAttemptKey(answer, field));
    if (attempts >= MAX_FIELD_FILL_ATTEMPTS) exhausted += 1;
    else pending += 1;
  }

  return {
    settled: answers.length > 0 && pending === 0,
    filled,
    exhausted,
    pending,
    max_attempts: MAX_FIELD_FILL_ATTEMPTS
  };
}

function applyRuntimeGptAnswersOnce(requestId, rawAnswers, fallbackFields = [], skipFilledIds = new Set()) {
  if (finalizedRuntimeGptRequestIds.has(requestId)) {
    return Promise.resolve({ ...emptyFillResult(), ignored: true, reason: "already-finalized" });
  }
  if (runtimeGptApplyPromises.has(requestId)) return runtimeGptApplyPromises.get(requestId);

  const promise = (async () => {
    const sourceFields = runtimeGptSourceFields.get(requestId) || fallbackFields || [];
    const sourceIds = new Set(sourceFields.map((field) => field.id));
    const answers = normalizeRuntimeGptAnswers({ answers: rawAnswers || [] })
      .filter((answer) => sourceIds.size === 0 || sourceIds.has(answer.field_id));
    if (answers.length === 0) return emptyFillResult();

    runtimeGptAnswersByRequest.set(requestId, answers);
    try {
      return await applyAnswers(answers, skipFilledIds, sourceFields);
    } finally {
      finalizedRuntimeGptRequestIds.add(requestId);
      traceAutoBid("runtime-gpt:request-finalized", {
        request_id: requestId,
        answers: answers.length
      });
    }
  })().finally(() => {
    runtimeGptApplyPromises.delete(requestId);
  });

  runtimeGptApplyPromises.set(requestId, promise);
  return promise;
}

function acknowledgeRuntimeGptAnswersApplied(requestId, result) {
  return send("GPT_ANSWER_APPLIED", {
    request_id: requestId,
    result: {
      filled: Number(result?.filled || 0),
      missed: Number(result?.missed || 0),
      applied_at: new Date().toISOString()
    }
  });
}

async function getSheetContextForRuntime(page) {
  try {
    const context = await promiseWithTimeout(
      send("SHEET_CONTEXT", { page, url: page.url }),
      SHEET_CONTEXT_TIMEOUT_MS,
      "Sheet context lookup timed out"
    );
    if (context?.rowNumber) {
      traceAutoBid("runtime-gpt:context-found", {
        row_number: context.rowNumber,
        sheet_name: context.sheetName || "",
        url: context.url || page.url,
        match_source: context.matchSource || ""
      });
      return context;
    }
  } catch (error) {
    traceAutoBid("runtime-gpt:context-error", { message: error.message || String(error) });
  }

  traceAutoBid("runtime-gpt:no-sheet-context", { url: page.url });
  return {
    rowNumber: null,
    sheetName: "",
    spreadsheetId: "",
    url: page.url,
    values: {},
    raw: []
  };
}

async function waitForRuntimeGptAnswer(requestId) {
  if (!requestId) return { status: "error", error: "Missing runtime GPT request id" };

  const started = Date.now();
  while (Date.now() - started < RUNTIME_GPT_ANSWER_TIMEOUT_MS) {
    let status;
    try {
      status = await send("GPT_ANSWER_STATUS", { request_id: requestId });
    } catch (error) {
      traceAutoBid("runtime-gpt:status-error", { request_id: requestId, message: error.message || String(error) });
      return { status: "error", error: error.message || String(error) };
    }

    if (status?.status === "complete" || status?.status === "error") return status;
    await sleep(RUNTIME_GPT_ANSWER_POLL_MS);
  }

  try {
    const finalStatus = await send("GPT_ANSWER_STATUS", { request_id: requestId });
    if (finalStatus?.status === "complete" || finalStatus?.status === "error") return finalStatus;
  } catch (error) {
    traceAutoBid("runtime-gpt:final-status-error", { request_id: requestId, message: error.message || String(error) });
  }

  traceAutoBid("runtime-gpt:wait-timeout-queue-preserved", { request_id: requestId });
  return { status: "timeout", request_id: requestId, queued: true };
}

async function maybeAutoSubmitApplication(fields, sheetResult, resumeResult) {
  if (!sheetResult || sheetResult.filled <= 0) {
    traceAutoBid("submit:checking-without-generated-fill", { reason: "no-generated-answers-filled" });
  }

  if (sheetResult.pending > 0) {
    traceAutoBid("submit:sheet-pending-after-fallback", { pending: sheetResult.pending });
  }

  if (resumeResult?.missed > 0 && hasMissingRequiredResumeUpload(fields)) {
    traceAutoBid("submit:skipped", { reason: "resume-attach-missed", resume: resumeResult });
    return { clicked: false, reason: "resume-attach-missed" };
  }

  await sleep(900);

  const missingRequired = getMissingRequiredFields(fields);
  if (missingRequired.length > 0) {
    traceAutoBid("submit:skipped", {
      reason: "required-fields-missing",
      missing: missingRequired.map((field) => ({
        field_id: field.id,
        label: field.label,
        type: field.type
      }))
    });
    return { clicked: false, reason: "required-fields-missing", missing: missingRequired.length };
  }

  const button = await waitForApplicationSubmitButton();
  if (!button) {
    traceAutoBid("submit:skipped", { reason: "button-not-found" });
    return { clicked: false, reason: "button-not-found" };
  }

  if (isDisabledSubmitButton(button)) {
    traceAutoBid("submit:skipped", {
      reason: "button-disabled",
      text: getSubmitButtonText(button)
    });
    return { clicked: false, reason: "button-disabled", text: getSubmitButtonText(button) };
  }

  const submissionStartedAt = new Date().toISOString();
  let verificationMonitor = null;
  try {
    verificationMonitor = await promiseWithTimeout(send("OUTLOOK_MONITOR_START", {
      since: submissionStartedAt,
      page_url: location.href,
      domain: location.hostname.replace(/^www\./, ""),
      title: document.title || "",
      run_id: autoBidRunId,
      profile_id: activeAutoBidProfileId,
      mailbox_email: activeAutoBidProfileEmail,
      submit_text: getSubmitButtonText(button)
    }), 2500, "Outlook verification monitor setup timed out");
    traceAutoBid("outlook-verification:monitor-start", {
      started: Boolean(verificationMonitor?.started),
      reason: verificationMonitor?.reason || "",
      monitor_id: verificationMonitor?.monitor_id || ""
    });
  } catch (error) {
    traceAutoBid("outlook-verification:monitor-unavailable", { message: error.message || String(error) });
  }

  await scrollElementIntoView(button, "center");
  await sleep(250);
  let clicked = await nativeClickElement(button);
  if (!clicked) {
    dispatchRealisticMouseClick(button);
    clicked = true;
  }
  await sleep(600);

  if (clicked && verificationMonitor?.started && verificationMonitor.monitor_id) {
    send("OUTLOOK_MONITOR_ARM", { monitor_id: verificationMonitor.monitor_id }).catch((error) => {
      traceAutoBid("outlook-verification:monitor-arm-error", {
        monitor_id: verificationMonitor.monitor_id,
        message: error.message || String(error)
      });
    });
    announceOutlookVerificationWait(verificationMonitor.monitor_id).catch(() => {});
  }

  traceAutoBid("submit:clicked", {
    clicked: Boolean(clicked),
    text: getSubmitButtonText(button)
  });

  return {
    clicked: Boolean(clicked),
    reason: clicked ? "submitted" : "click-failed",
    text: getSubmitButtonText(button),
    verification_monitor: verificationMonitor?.started ? verificationMonitor.monitor_id : ""
  };
}

async function announceOutlookVerificationWait(monitorId) {
  const candidates = await waitForOutlookVerificationCodeFields(8000);
  if (candidates.length === 0 || completedOutlookVerificationMessages.has(monitorId)) return;
  showStatus("Waiting for Outlook verification code", "waiting", {
    detail: "The application requested email verification. Auto Bid is watching the connected mailbox in the background."
  });
}

async function triggerSheetGptAnswerWorker(context) {
  const result = await send("RUN_GPT_ANSWER", {
    mode: "once",
    rowNumber: context?.rowNumber || null,
    sheetName: context?.sheetName || ""
  });
  traceAutoBid("sheet:gpt-worker-triggered", {
    row_number: context?.rowNumber || null,
    result
  });
  return result;
}

function getMissingRequiredFields(fields) {
  return (fields || []).filter((field) => {
    if (!field.required) return false;
    if (field.type === "file") return !hasRequiredFileFieldValue(field);
    if (isGeneratedAnswerIgnoredFieldType(field.type)) return false;
    const controls = getControlsByFieldId(field.id);
    return controls.length > 0 && !hasFieldCurrentValue(field, controls);
  });
}

function hasMissingRequiredResumeUpload(fields) {
  const collectedFileFieldMissing = (fields || []).some((field) =>
    field.required && field.type === "file" && isResumeUploadField(field) && !hasRequiredFileFieldValue(field)
  );
  if (collectedFileFieldMissing) return true;

  const requiredResumeInputs = getResumeFileInputs().filter((input) =>
    input.required || input.getAttribute("aria-required") === "true"
  );
  return requiredResumeInputs.some((input) => !isResumeInputAttached(input)) && !hasResumeUploadedUi("");
}

function hasRequiredFileFieldValue(field) {
  const controls = getControlsByFieldId(field.id);
  if (controls.some((control) => control.files?.length)) return true;
  return hasResumeUploadedUi("");
}

function isResumeUploadField(field) {
  const text = normalize([field.question, field.label, field.option, field.name, field.placeholder].filter(Boolean).join(" "));
  return /\b(resume|cv|curriculum vitae)\b/.test(text);
}

function findApplicationSubmitButton() {
  const candidates = queryAll("button, input[type='submit'], input[type='button'], [role='button'], a")
    .filter((element) => isVisible(element))
    .map((element) => ({
      element,
      text: normalize(getSubmitButtonText(element)),
      disabled: isDisabledSubmitButton(element)
    }))
    .filter((candidate) => candidate.text && !isNonSubmitActionText(candidate.text))
    .filter((candidate) => isLikelyFinalApplicationSubmit(candidate.element, candidate.text))
    .map((candidate) => ({
      ...candidate,
      score: scoreSubmitButton(candidate.text, candidate.disabled)
    }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score);

  return candidates[0]?.element || null;
}

async function waitForApplicationSubmitButton(timeoutMs = 8000) {
  const started = Date.now();
  let lastButton = null;

  while (Date.now() - started < timeoutMs) {
    const button = findApplicationSubmitButton();
    if (button) {
      lastButton = button;
      if (!isDisabledSubmitButton(button)) return button;
    }
    await sleep(400);
  }

  return lastButton;
}

function getSubmitButtonText(element) {
  return cleanLabel([
    element?.textContent,
    element?.value,
    element?.getAttribute?.("aria-label"),
    element?.getAttribute?.("title")
  ].filter(Boolean).join(" "));
}

function isDisabledSubmitButton(element) {
  return Boolean(
    element?.disabled ||
    element?.getAttribute?.("aria-disabled") === "true" ||
    element?.matches?.("[disabled], .disabled, [class*='disabled' i]")
  );
}

function isNonSubmitActionText(text) {
  return /^apply!?$/.test(text) ||
    /(attach|upload|browse|choose file|clear|remove|delete|cancel|back|previous|save draft|google drive|enter manually|view website|view all jobs|help|login|sign up|log out)/.test(text);
}

function isLikelyFinalApplicationSubmit(element, text) {
  if (/(submit application|submit my application|send application|send my application)/.test(text)) return true;
  if (/(submit|send)/.test(text) && isElementNearApplicationForm(element)) return true;
  if (/(apply now|apply for this job)/.test(text) && isElementNearApplicationForm(element)) return true;
  if (/(continue|next)/.test(text) && /(application|submit|review)/.test(text) && isElementNearApplicationForm(element)) return true;
  return false;
}

function isElementNearApplicationForm(element) {
  const form = element.closest?.("form");
  if (form && queryAll("input, textarea, select, [role='checkbox'], [role='radio'], [role='combobox']", form).length > 0) return true;

  const container = element.closest?.("main, section, article, [class*='application' i], [class*='apply' i], [class*='form' i]") || document.body;
  const text = normalize(container?.textContent || "");
  const controls = queryAll("input, textarea, select, [role='checkbox'], [role='radio'], [role='combobox']", container)
    .filter((control) => isVisible(control));
  return controls.length > 0 &&
    /(required|resume|cv|cover letter|privacy|consent|first name|last name|email|phone|submit application)/.test(text);
}

function scoreSubmitButton(text, disabled) {
  let score = disabled ? -20 : 0;
  if (/(submit application|submit my application|send application|send my application)/.test(text)) score += 120;
  else if (/(apply now|apply for this job|submit|send application)/.test(text)) score += 100;
  else if (/^apply$|^send$/.test(text)) score += 80;
  else if (/(continue|next)/.test(text) && /(application|submit|review)/.test(text)) score += 50;
  return score;
}

function getSheetQuestionCandidateFields(fields, filledIds) {
  return getGeneratedAnswerCandidateFields(fields, filledIds);
}

function getSheetCandidateSkipReason(field, filledIds) {
  return getGeneratedAnswerCandidateSkipReason(field, filledIds);
}

function getGeneratedAnswerCandidateFields(fields, filledIds) {
  return fields.filter((field) => !getGeneratedAnswerCandidateSkipReason(field, filledIds));
}

async function hydrateGeneratedChoiceOptions(fields) {
  for (const field of fields) {
    if (!field || !CHOICE_FIELD_TYPES.includes(field.type)) continue;
    const controls = getControlsByFieldId(field.id);
    const control = controls[0];
    if (!control || hasFieldCurrentValue(field, controls)) continue;

    let discoveredOptions = [];
    if (field.type === "combobox") {
      await clearComboboxSearchValue(control);
      const options = await getComboboxChoices(control);
      discoveredOptions = options.map((option) => cleanLabel(
        option.textContent || option.getAttribute("data-value") || option.getAttribute("value") || ""
      ));
      await closeCombobox(control);
    } else if (field.type === "radio") {
      discoveredOptions = controls.map(getRadioOptionLabel);
    } else if (field.type === "button-group") {
      discoveredOptions = controls.map(getChoiceButtonLabel);
    } else if (field.type === "checkbox") {
      discoveredOptions = ["Yes", "No"];
    } else {
      discoveredOptions = getControlOptions(control);
    }

    field.options = uniqueNonEmptyValues([...(field.options || []), ...discoveredOptions])
      .filter((option) => !isGeneratedChoicePlaceholder(option, option));
    traceAutoBid("ai:choice-options-hydrated", {
      field_id: field.id,
      label: field.label,
      type: field.type,
      options: field.options,
      complete: field.options.length > 0
    });
  }
}

function hasCompleteGeneratedChoiceOptions(field) {
  if (!CHOICE_FIELD_TYPES.includes(field?.type)) return true;
  return Array.isArray(field.options) && field.options.some((option) => !isGeneratedChoicePlaceholder(option, option));
}

function getGeneratedAnswerCandidateSkipReason(field, filledIds) {
  if (!field?.id) return "missing-field-id";
  if (isGeneratedAnswerIgnoredFieldType(field.type)) return "ignored-field-type";
  if (!field.required) return "not-required";
  if (isSensitiveGeneratedAnswerField(field)) return "sensitive-field";
  if (isProfileStaticQuestionField(field)) return "profile-static-field";
  if (filledIds?.has?.(field.id)) return "already-resolved-locally";

  const controls = getControlsByFieldId(field.id);
  if (controls.length === 0) return "control-not-found";
  if (hasFieldCurrentValue(field, controls)) return "already-has-value";

  return "";
}

function isGeneratedAnswerIgnoredFieldType(type) {
  return ["file", "hidden", "password", "submit", "button", "reset", "image"].includes(type);
}

function isSensitiveGeneratedAnswerField(field) {
  const text = normalize([field.question, field.label, field.option, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(captcha|recaptcha|one time|otp|verification code|coupon|promo|password|gender|race|ethnicity|ethnic|disability|veteran|protected veteran|sexual orientation|date of birth|birth date|national id|social security|ssn|passport)/.test(text);
}

function isProfileStaticQuestionField(field) {
  const prompt = normalize(field.question || field.label || field.name || field.placeholder || "");
  const metadata = normalize([field.name, field.placeholder, field.autocomplete].filter(Boolean).join(" "));
  if (!prompt && !metadata) return false;

  if (/(experience|years|skill|proficiency|technology|framework|develop|authorization|authorisation|authorized|authorised|visa|sponsor|sponsorship|work permit|right to work|relocat|willing|comfortable|why|describe|explain)/.test(prompt)) {
    return false;
  }

  if (field.type === "email" || /^(email|e mail)$/.test(normalize(field.autocomplete || ""))) return true;
  if (field.type === "tel" || /^(tel|phone|mobile)$/.test(normalize(field.autocomplete || ""))) return true;

  const contactPrompt = /^(?:(?:what is|please enter|please provide|enter|provide|type|select)\s+)?(?:your\s+)?(?:first name|given name|last name|family name|surname|full name|name|email(?: address)?|e mail(?: address)?|phone(?: number)?|mobile(?: number)?|telephone(?: number)?|linkedin(?: profile)?(?: url| link)?|github(?: profile)?(?: url| link)?|portfolio(?: url| link)?|personal (?:site|website)(?: url| link)?|website url|profile url)\b/;
  if (contactPrompt.test(prompt)) return true;

  const locationPrompt = /^(?:(?:what is|please enter|please provide|enter|provide|type|select)\s+)?(?:your\s+)?(?:current\s+)?(?:city|country|country of residence|location|address|residence|postal code|post code|postcode|zip code|state|province|region|state province region)\b|^where (?:are|do) you (?:currently )?(?:live|reside|located|based)\b|^where is your current residence\b/;
  if (locationPrompt.test(prompt)) return true;

  return /^(first name|last name|full name|email|phone|mobile|city|country|location|address|postal code|postcode|zip code|state|province|region|state region|linkedin|github|portfolio|website)(?: id| field| input)?$/.test(metadata);
}

function isLocallyAnswerableGeneratedField(field) {
  const text = normalize([field.question, field.label, field.option, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!text) return false;
  return /(terms|privacy|policy|consent|agree|accept|acknowledge|confirm|certify|accurate|contact me|future job opportunit|future opportunit|talent community|job alert|recruiting communication|recruitment communication|how did you hear|how.*hear.*role|where did you hear|how did you become aware|become aware.*(?:vacancy|role|job|position|opportunity)|aware.*(?:vacancy|role|job|position|opportunity)|source|referral|referred|job board|date available|earliest available|start date|notice period)/.test(text);
}

function hasFieldCurrentValue(field, controls = getControlsByFieldId(field?.id)) {
  if (!controls || controls.length === 0) return false;
  if (getControlType(controls[0]) === "combobox" && !hasCurrentChoiceValue(controls)) return false;
  const value = cleanLabel(getCurrentChoiceSummary(controls) || field?.value || "");
  if (!value || isPlaceholderChoice(value, value)) return false;
  if (isRejectedGeneratedPlaceholder(value)) return false;
  if (isExperienceValueField(field, controls) && isZeroLikeExperienceValue(value)) return false;

  const placeholder = cleanLabel(field?.placeholder || controls[0]?.getAttribute?.("placeholder") || "");
  if (placeholder && normalize(value) === normalize(placeholder)) return false;
  return true;
}

function isGeneratedAnswerChoiceField(field) {
  return field?.type === "checkbox" || isChoiceFieldType(field?.type);
}

function parseBooleanAnswer(value) {
  const text = normalize(value);
  if (/^(yes|true|agree|agreed|accept|accepted|checked|on|1)$/.test(text)) return true;
  if (/^(no|false|decline|declined|reject|rejected|unchecked|off|0)$/.test(text)) return false;
  return null;
}

function isSuspiciousNarrativeBooleanAnswer(answer, field, controls = []) {
  if (parseBooleanAnswer(answer?.value ?? answer?.answer ?? "") === null) return false;
  const first = controls?.[0];
  const type = first ? getControlType(first) : field?.type;
  if (!["text", "search", "textarea", "contenteditable"].includes(type)) return false;

  const visibleQuestion = first ? getVisualFieldLabel(first) : "";
  const question = cleanLabel(visibleQuestion || field?.question || field?.label || answer?.question || "");
  return !isSemanticBooleanQuestion(question);
}

function isSemanticBooleanQuestion(question) {
  const text = normalize(question);
  if (!text) return false;
  return /^(?:do|does|did|are|is|was|were|have|has|had|can|could|will|would|should|may|must)\b/.test(text) ||
    /\b(?:yes\s*(?:or|\/)\s*no|confirm whether|please confirm (?:that|whether|if)|eligible to|authorized to|authorised to|willing to|open to|able to)\b/.test(text);
}

function generatedAnswerValuesMatch(currentValue, expectedValue) {
  const current = normalize(currentValue);
  const expected = normalize(expectedValue);
  if (!current || !expected) return false;
  if (current === expected) return true;

  const currentBoolean = parseBooleanAnswer(current);
  const expectedBoolean = parseBooleanAnswer(expected);
  if (currentBoolean !== null && expectedBoolean !== null) return currentBoolean === expectedBoolean;

  return scoreChoice(current, expected) >= 80;
}

function doesGeneratedAnswerMatchField(answer, field, controls = getControlsByFieldId(field?.id)) {
  const first = controls?.[0];
  if (!first) return false;

  const expectedValue = String(answer?.value ?? answer?.answer ?? "").trim();
  if (!expectedValue) return false;
  const type = getControlType(first);

  if (type === "checkbox") {
    const expected = parseBooleanAnswer(expectedValue);
    if (expected === null) return false;

    const explicitChoice = getSelectedCheckboxBooleanChoiceLabel(first);
    if (explicitChoice) return parseBooleanAnswer(explicitChoice) === expected;

    const option = cleanLabel(field?.option || answer?.option || getCheckboxOptionLabel(first, field?.question || field?.label || ""));
    if (option) return isCheckboxChecked(first) === expected;

    return expected === true && isCheckboxChecked(first);
  }

  const currentValue = getCurrentChoiceSummary(controls);
  return generatedAnswerValuesMatch(currentValue, expectedValue);
}

async function waitForGeneratedAnswerMatch(answer, sourceField, fallbackControls = [], timeoutMs = 900) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const binding = resolveLiveFieldForAnswer(answer, sourceField);
    const field = binding?.field || sourceField;
    const controls = binding?.controls?.length ? binding.controls : fallbackControls;
    if (field && doesGeneratedAnswerMatchField(answer, field, controls)) return true;
    await sleep(90);
  }

  const binding = resolveLiveFieldForAnswer(answer, sourceField);
  return Boolean(binding && doesGeneratedAnswerMatchField(answer, binding.field, binding.controls));
}

function isZeroLikeExperienceValue(value) {
  return /^(0|0\.0+|0\s*(?:years?|yrs?)?)$/i.test(String(value || "").trim());
}

function normalizeSheetAnswers(data) {
  const answers = Array.isArray(data?.answers) ? data.answers : [];
  return answers
    .map((answer) => ({
      field_id: String(answer.field_id || answer.id || ""),
      value: String(answer.value ?? answer.answer ?? "").trim(),
      source: "sheet"
    }))
    .filter((answer) => answer.field_id && answer.value && !isRejectedGeneratedPlaceholder(answer.value));
}

function normalizeRuntimeGptAnswers(data) {
  const answers = Array.isArray(data?.answers) ? data.answers : [];
  return answers
    .map((answer) => ({
      field_id: String(answer.field_id || answer.id || ""),
      question: String(answer.question || ""),
      option: String(answer.option || ""),
      value: String(answer.value ?? answer.answer ?? "").trim(),
      source: "runtime-gpt"
    }))
    .filter((answer) => answer.field_id && answer.value && !isRejectedGeneratedPlaceholder(answer.value));
}

function normalizeDirectAiAnswers(data) {
  const answers = Array.isArray(data?.answers) ? data.answers : [];
  return answers
    .map((answer) => ({
      field_id: String(answer.field_id || answer.id || ""),
      value: String(answer.value ?? answer.answer ?? "").trim(),
      source: answer.source || "ai",
      provider: answer.provider || "",
      model: answer.model || "",
      estimated_request_cost_usd: answer.estimated_request_cost_usd ?? null
    }))
    .filter((answer) => answer.field_id && answer.value && !isRejectedGeneratedPlaceholder(answer.value));
}

function isRejectedGeneratedPlaceholder(value) {
  return /^(?:not specified|unspecified|unknown|not provided|not available|no information(?: provided| available)?|information unavailable|to be determined|tbd)$/i
    .test(String(value || "").trim());
}

function buildSheetQuestionPayload(page, context, fields, profileContext = {}) {
  const rowValues = compactSheetRowValues(context.values || {});
  const gptContext = getSheetGptContext(context, rowValues);

  return {
    version: 1,
    source: "autobid-extension",
    requested_at: new Date().toISOString(),
    gpt_context: gptContext,
    profile: {
      profile_id: profileContext.profile_id || "",
      static_fields: compactSheetRowValues(profileContext.static_fields || {})
    },
    page: {
      url: page.url,
      domain: page.domain,
      title: page.title,
      job_title: page.job_title,
      text: limitSheetText(page.text, 6000)
    },
    row: {
      row_number: context.rowNumber,
      url: context.url || page.url,
      values: rowValues
    },
    response_format: {
      column: "autobid_answers",
      json: { answers: [{ field_id: "field id from fields", question: "plain question text", option: "checkbox option when present", value: "answer to fill" }] }
    },
    fields: fields.map((field) => serializeSheetQuestionField(field))
  };
}

async function getGeneratedAnswerProfileContext() {
  try {
    return await send("GET_PROFILE_STATIC_FIELDS");
  } catch (error) {
    traceAutoBid("runtime-gpt:profile-context-error", { message: error.message || String(error) });
    return {};
  }
}

function getSheetGptContext(context, rowValues) {
  const tailoredResumeContent = getSheetContextValue(context, rowValues, 7, [
    "column_g",
    "tailored_resume_content",
    "tailored resume content",
    "tailor_resume_content",
    "tailor resume content",
    "tailored_resume",
    "tailored resume",
    "generated_resume_content",
    "generated resume content",
    "resume_content",
    "resume content",
    "candidate_profile",
    "candidate profile",
    "profile"
  ]);
  const jobDescription = getSheetContextValue(context, rowValues, 13, [
    "column_m",
    "job_description",
    "job description",
    "jd",
    "job_desc",
    "job desc",
    "description",
    "job_posting",
    "job posting",
    "job_details",
    "job details"
  ]);

  return {
    source_columns: {
      tailored_resume_content: "G",
      job_description: "M"
    },
    tailored_resume_content: limitSheetText(tailoredResumeContent, 12000),
    job_description: limitSheetText(jobDescription, 12000)
  };
}

function getSheetContextValue(context, rowValues, oneBasedColumn, aliases) {
  const fromRawValues = findSheetValueByAlias(context.values || {}, aliases);
  if (fromRawValues) return fromRawValues;

  if (Array.isArray(context.raw)) {
    const fromRawColumn = String(context.raw[oneBasedColumn - 1] || "").trim();
    if (fromRawColumn) return fromRawColumn;
  }

  const fromCompactedValues = findSheetValueByAlias(rowValues, aliases);
  if (fromCompactedValues) return fromCompactedValues;

  return "";
}

function findSheetValueByAlias(values, aliases) {
  const normalizedAliases = new Set(aliases.map(normalize));
  for (const [key, value] of Object.entries(values || {})) {
    const text = String(value || "").trim();
    if (text && normalizedAliases.has(normalize(key))) return text;
  }
  return "";
}

function serializeSheetQuestionField(field) {
  const controls = getControlsByFieldId(field.id);
  const first = controls[0];
  const type = first ? getControlType(first) : field.type;
  const visibleQuestion = !["checkbox", "radio", "button-group"].includes(type) && first
    ? getVisualFieldLabel(first)
    : "";
  const question = getPlainQuestionText(visibleQuestion || field.question || field.label);
  const option = getPlainQuestionText(field.option || getChoiceOptionTextFromControls(field, controls));
  return {
    field_id: field.id,
    question,
    option,
    label: question,
    raw_label: field.label,
    help_text: getGeneratedFieldHelpText(controls[0], question),
    name: field.name,
    placeholder: field.placeholder,
    autocomplete: field.autocomplete,
    type: field.type,
    required: Boolean(field.required),
    options: field.options || [],
    current_value: hasFieldCurrentValue(field, controls) ? getCurrentChoiceSummary(controls) || field.value || "" : ""
  };
}

function getGeneratedFieldHelpText(control, question) {
  if (!control) return "";
  const described = cleanLabel(getDescribedByText(control));
  const nearby = cleanLabel(getNearbyText(control));
  const combined = cleanLabel([described, nearby].filter(Boolean).join(" "));
  if (!combined) return "";
  const withoutQuestion = question
    ? combined.replace(new RegExp(escapeRegExp(question), "ig"), " ")
    : combined;
  return cleanLabel(withoutQuestion).slice(0, 800);
}

function getChoiceOptionTextFromControls(field, controls) {
  const control = controls?.[0];
  if (!control) return "";
  if (field.type === "checkbox") return getCheckboxOptionLabel(control, field.question || field.label);
  return "";
}

function compactSheetRowValues(values) {
  return Object.fromEntries(
    Object.entries(values || {})
      .filter(([_key, value]) => String(value || "").trim())
      .slice(0, 80)
      .map(([key, value]) => [key, limitSheetText(value, 3000)])
  );
}

function limitSheetText(value, maxLength) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text.length > maxLength ? `${text.slice(0, maxLength - 3)}...` : text;
}

function shouldApplyProfileStaticValue(field, controls, rawValue, key = "") {
  const first = controls[0];
  const type = getControlType(first);
  if (key === "phone" && ["checkbox", "radio", "select", "combobox", "button-group"].includes(type)) {
    return false;
  }
  if (type === "checkbox" && parseBooleanAnswer(rawValue) === null) return false;

  const current = String(getCurrentChoiceSummary(controls) || field.value || "").trim();
  if (!current) return true;
  if (isRejectedGeneratedPlaceholder(current)) return true;
  if (normalize(current) === normalize(rawValue)) return false;

  if (["checkbox", "radio", "select", "combobox", "button-group"].includes(type)) {
    return !hasCurrentChoiceValue(controls);
  }

  const placeholder = String(field.placeholder || first.getAttribute?.("placeholder") || "").trim();
  if (placeholder && normalize(current) === normalize(placeholder)) return true;

  return isLikelyExampleStaticValue(current);
}

function isLikelyExampleStaticValue(value) {
  const text = normalize(value);
  return /^(your|enter|type|sample|example|test)\b/.test(text) ||
    /\b(example|sample|placeholder)\b/.test(text) ||
    /\bolivia manatal com\b/.test(text);
}

function findProfileCountryAnswer(fields, answers) {
  const fieldsById = new Map(fields.map((field) => [field.id, field]));
  const candidates = answers
    .map((answer) => ({
      value: String(answer.value || "").trim(),
      field: fieldsById.get(answer.field_id)
    }))
    .filter((candidate) => candidate.value && candidate.field)
    .map((candidate) => ({
      ...candidate,
      score: scoreCountryAnswerField(candidate.field)
    }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score);

  return candidates[0]?.value || "";
}

async function getProfileStaticLocationAnswer() {
  try {
    const profileStatic = await send("GET_PROFILE_STATIC_FIELDS");
    const staticFields = profileStatic?.static_fields || {};
    return String(staticFields.country || staticFields.location || staticFields.city || "").trim();
  } catch {
    return "";
  }
}

function scoreCountryAnswerField(field) {
  const label = normalize([field.label, field.name, field.placeholder, field.autocomplete].filter(Boolean).join(" "));
  if (isResidenceField(field)) return 100;
  if (/\bcountry\b/.test(label)) return 80;
  if (/(current location|location|where are you based|where.*based)/.test(label)) return 60;
  if (/\b(address|city)\b/.test(label)) return 50;
  return 0;
}

function isResidenceField(field) {
  const label = normalize([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(current residence|residence|where is your current residence|where are you based|where.*based)/.test(label);
}

function matchProfileStaticFieldKey(field) {
  const text = normalize([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  const optionText = normalize((field.options || []).join(" "));
  if (isLanguageChoiceField(field)) return "";
  if (isCombinedProfileLocationField(field)) return "location";
  if (isBasedInLocationField(field)) return "";
  if (isAuthorizationSupportRequiredText(text)) return "";
  if (isSemanticAvailabilityQuestion(field)) return "";
  if (isPlainFullNameField(field)) return "full_name";
  if (isProfilePhoneField(field)) return "phone";
  const addressComponentKey = matchProfileAddressComponentKey(field);
  if (addressComponentKey) return addressComponentKey;
  if (/(work status|right to work|employment status)/.test(text) &&
      /(national|citizen|work permit|residence permit|third country)/.test(optionText)) {
    return "work_authorization";
  }
  if (/\bcity\b/.test(text) || /(location city|city location)/.test(text)) return "city";
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name", "candidate name"]],
    ["email", ["email", "e mail", "mail"]],
    ["notice_period", ["notice period", "current notice", "notice"]],
    ["expected_rate", ["hourly rate", "rate", "expected rate", "expected salary", "salary expectation", "salary expectations", "expected compensation", "desired salary", "desired compensation", "desired pay", "desired annual salary", "annual salary", "day rate", "pay expectation", "pay expectations", "gross monthly", "monthly salary", "salary", "compensation"]],
    ["work_authorization", ["authorized", "authorization", "legally work", "eligible to work", "right to work", "work status", "employment status"]],
    ["sponsorship", ["sponsor", "sponsorship", "visa"]],
    ["availability", ["availability", "date available", "available date", "available start date", "earliest start date", "start date", "when can you start"]],
    ["linkedin", ["linkedin"]],
    ["github", ["github"]],
    ["portfolio", ["portfolio"]],
    ["website", ["website", "personal site", "web site"]],
    ["languages", ["languages", "spoken languages", "language proficiency", "fluent languages", "languages spoken"]],
    ["nationality", ["nationality", "citizenship", "country of citizenship", "citizen of"]],
    ["country", ["country", "residence", "current residence", "where is your current residence", "where are you based"]],
    ["location", ["location", "address", "current location", "currently located", "where are you currently located", "where are you located"]],
    ["city", ["city"]]
  ];

  for (const [key, needles] of patterns) {
    if (needles.some((needle) => includesNormalizedProfilePhrase(text, needle))) return key;
  }

  return "";
}

function isProfilePhoneField(field) {
  const type = normalize(field?.type || "");
  const autocomplete = normalize(field?.autocomplete || "");
  const text = normalize([field?.name, field?.label, field?.placeholder].filter(Boolean).join(" "));
  if (type === "tel" || /^(tel|phone|mobile)$/.test(autocomplete)) return true;
  if (/\b(phone|telephone|cell)(?:\s+number)?\b/.test(text)) return true;
  return /\bmobile\b/.test(text) && (
    /\bmobile\s+(?:phone|number|contact)\b/.test(text) ||
    /^(?:your\s+)?mobile(?:\s+number)?$/.test(text)
  );
}

function isSemanticAvailabilityQuestion(field) {
  const text = normalize([field?.question, field?.label, field?.name, field?.placeholder].filter(Boolean).join(" "));
  if (!text) return false;
  const asksForAvailability = /(are|would|will|can|could|do) you.{0,80}(available|open|willing|able)|\bavailable for\b|\bopen (?:to|for)\b/.test(text);
  const semanticWorkCondition = /(part time|full time|remote|onsite|on site|hybrid|contract|contractor|b2b|shift|hours|weekend|travel|relocat|collaboration)/.test(text);
  return asksForAvailability && semanticWorkCondition;
}

function includesNormalizedProfilePhrase(text, phrase) {
  const normalizedPhrase = normalize(phrase);
  return Boolean(normalizedPhrase) && ` ${text} `.includes(` ${normalizedPhrase} `);
}

function matchProfileAddressComponentKey(field) {
  const autocomplete = normalize(field.autocomplete || "");
  const name = normalize(field.name || "");
  const prompt = simplifyProfileAddressPrompt(field.label || field.placeholder || "");

  if (autocomplete === "postal code") return "postal_code";
  if (autocomplete === "address level1") return "state_region";

  const postalAliases = new Set([
    "postal code",
    "postalcode",
    "post code",
    "postcode",
    "zip code",
    "zipcode",
    "zip",
    "pin code",
    "postal code zip code",
    "zip code postal code",
    "postal zip code",
    "zip postal code"
  ]);
  if (postalAliases.has(prompt) || postalAliases.has(name)) return "postal_code";

  const regionAliases = new Set([
    "state",
    "province",
    "region",
    "county",
    "prefecture",
    "administrative area",
    "state province",
    "state region",
    "province region",
    "state province region"
  ]);
  if (regionAliases.has(prompt) || regionAliases.has(name)) return "state_region";

  return "";
}

function simplifyProfileAddressPrompt(value) {
  return normalize(value)
    .replace(/^(?:what is|please enter|please select|enter|select|choose|provide)\s+(?:your\s+)?(?:current\s+)?/, "")
    .replace(/\s+(?:select|choose)(?:\s+(?:a|an)\s+option)?$/, "")
    .replace(/\s+(?:required|optional)$/, "")
    .trim();
}

function isPlainFullNameField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const candidates = [field.label, field.name, field.autocomplete]
    .map(normalize)
    .filter(Boolean);
  return candidates.some((candidate) =>
    ["name", "your name", "applicant name", "candidate name", "full name", "preferred name"].includes(candidate) ||
    /\bfirst(?:\s+and|\s*\/)\s*last\s+name\b|\blast(?:\s+and|\s*\/)\s*first\s+name\b/.test(candidate)
  );
}

function isCombinedProfileLocationField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const text = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!/(location|where.*based|based.*in|residence)/.test(text)) return false;
  const parts = ["city", "state", "country"].filter((part) => new RegExp(`\\b${part}\\b`).test(text));
  return parts.length >= 2 || /what location.*based|where.*(?:located|based|reside)/.test(text);
}

function isTextLikeStaticField(field) {
  return !["checkbox", "radio", "select", "combobox", "button-group", "file", "hidden", "password", "submit", "button", "reset"].includes(field.type);
}

function getProfileStaticValue(staticFields, key) {
  const aliases = {
    postal_code: ["postal_code", "postalcode", "post_code", "postcode", "zip_code", "zipcode", "zip"],
    state_region: ["state_region", "state_province_region", "state_province", "state", "province", "region", "administrative_area"],
    expected_rate: ["expected_rate", "expected_salary", "salary_expectation", "salary_expectations", "monthly_salary", "monthly_salary_expectation", "desired_salary", "desired_compensation", "desired_pay", "pay_expectation", "pay_expectations", "annual_salary", "day_rate"],
    notice_period: ["notice_period", "current_notice_period", "availability_notice"],
    languages: ["languages", "language", "spoken_languages", "language_proficiency", "fluent_languages", "languages_spoken"],
    work_authorization: ["work_authorization", "right_to_work", "work_status", "employment_status"],
    nationality: ["nationality", "citizenship", "citizen_of", "country_of_citizenship"]
  };

  if (key === "full_name") {
    const fullName = staticFields?.full_name;
    if (fullName !== undefined && fullName !== null && String(fullName).trim()) return fullName;
    const composed = [staticFields?.first_name, staticFields?.last_name].filter(Boolean).join(" ").trim();
    if (composed) return composed;
  }

  if (key === "location") {
    const location = String(staticFields?.location || "").trim();
    if (location) return location;
    return [staticFields?.city, staticFields?.country].filter(Boolean).join(", ").trim();
  }

  if (key === "city") {
    const city = String(staticFields?.city || "").trim();
    if (city) return city;
    const location = String(staticFields?.location || "").trim();
    return location.split(",").map((part) => part.trim()).filter(Boolean)[0] || "";
  }

  if (key === "country") {
    const country = String(staticFields?.country || "").trim();
    if (country) return country;
    const parts = String(staticFields?.location || "").split(",").map((part) => part.trim()).filter(Boolean);
    return parts.length > 1 ? parts[parts.length - 1] : "";
  }

  for (const candidate of [key, ...(aliases[key] || [])]) {
    const value = staticFields?.[candidate];
    if (value !== undefined && value !== null && String(value).trim()) return value;
  }

  return staticFields?.[key];
}

function getProfileStaticCandidateValues(staticFields, key, field, controls = []) {
  const primary = getProfileStaticValue(staticFields, key);
  if (key === "phone") {
    return getPhoneEntryCandidateValues(primary, staticFields, field, controls);
  }
  if (!["location", "city"].includes(key)) {
    return uniqueNonEmptyValues([primary]);
  }

  const city = getProfileStaticValue(staticFields, "city");
  const country = getProfileStaticValue(staticFields, "country");
  const location = getProfileStaticValue(staticFields, "location");
  const searchableLocation = field.type === "combobox" || controls.some((control) => shouldSelectTextAutocompleteOption(control, field));

  if (!searchableLocation) return uniqueNonEmptyValues([primary]);
  if (key === "city") return uniqueNonEmptyValues([city, location, primary]);
  return uniqueNonEmptyValues([location, city, primary]);
}

function getPhoneEntryCandidateValues(primary, staticFields, field, controls = []) {
  const raw = String(primary || "").trim();
  if (!raw) return [];

  const country = getPhoneDialCountryValue(staticFields);
  const dialCode = getDialCodeForCountry(country);
  const hasSeparateDialCodeControl = location.hostname.toLowerCase().includes("workable.com") ||
    controls.some((control) => hasNearbyPhoneDialCodeSelector(control)) ||
    getPhoneDialCodeControls().length > 0;
  if (!hasSeparateDialCodeControl || !dialCode) return uniqueNonEmptyValues([raw]);

  const rawDigits = raw.replace(/\D/g, "");
  const dialDigits = dialCode.replace(/\D/g, "");
  const nationalDigits = rawDigits.startsWith(dialDigits)
    ? rawDigits.slice(dialDigits.length).replace(/^0+/, "")
    : rawDigits.replace(/^0+/, "");
  const nationalFormatted = raw
    .replace(new RegExp(`^(?:\\+|00)?${dialDigits}\\s*`), "")
    .trim();

  traceAutoBid("phone:national-candidates", {
    field_id: field?.id || "",
    country,
    dial_code: dialCode,
    values: uniqueNonEmptyValues([nationalDigits, nationalFormatted, raw]).map((value) => shortText(value))
  });
  return uniqueNonEmptyValues([nationalDigits, nationalFormatted, raw]);
}

function hasNearbyPhoneDialCodeSelector(phoneInput) {
  const roots = [
    phoneInput?.parentElement,
    phoneInput?.closest?.("[class*='phone' i], [class*='field' i], [class*='input' i], .form-group"),
    getFieldContainer(phoneInput)
  ].filter(Boolean);
  return roots.some((root) => queryAll([
    "select",
    "[role='combobox']",
    "[aria-haspopup='listbox']",
    "[aria-haspopup='menu']",
    ".select__control"
  ].join(","), root).some((control) => control !== phoneInput && isPhoneDialCodeSelector(control)));
}

function uniqueNonEmptyValues(values) {
  const seen = new Set();
  return values
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .filter((value) => {
      const key = normalize(value);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function formatProfileStaticValueForField(key, value, field) {
  const textValue = String(value || "").trim();
  if (key !== "expected_rate") return textValue;

  const label = normalize([field.label, field.name, field.placeholder].join(" "));
  if (!/(gross|monthly|eur amount|amount)/.test(label)) return textValue;

  const amount = textValue.match(/\b\d{1,3}(?:[,\s]\d{3})+(?:[.,]\d+)?\b|\b\d+(?:[.,]\d+)?\b/);
  return amount ? amount[0].replace(/\s+/g, "").replace(/,/g, "") : textValue;
}

function isBasedInLocationField(field) {
  const label = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(currently.*based.*in|based.*in|currently.*located.*in|located.*in|currently.*living.*in|living.*in|fully.*living.*resident.*in|living.*resident.*in|currently.*residing.*in|residing.*in|resident.*in|residency.*work permit.*in|work permit.*in|current.*residence.*in)/.test(label);
}

function isLanguageChoiceField(field) {
  if (!field || !isChoiceFieldType(field.type)) return false;
  const label = normalize([field.question, field.label, field.option, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!/(speak|language|fluent|fluency|proficien|native speaker|bilingual|multilingual)/.test(label)) return false;
  if (getQuestionLanguageAliases(label).length === 0) return false;
  return hasYesNoOptions(field.options) || isLanguageProficiencyScaleField(field);
}

function isReferralSourceField(field) {
  if (!field || !isChoiceFieldType(field.type)) return false;
  const label = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!/(how did you hear|how.*hear.*role|how.*hear.*job|how.*hear.*position|where did you hear|where.*hear.*role|how did you become aware|become aware.*(?:vacancy|role|job|position|opportunity)|aware.*(?:vacancy|role|job|position|opportunity)|source|referral|referred|job board|found.*role|found.*job|learn.*role|learn.*job|vacancy source)/.test(label)) {
    return false;
  }
  const options = (field.options || []).map(normalize);
  return options.length === 0 || options.some((option) => /(linkedin|indeed|job board|website|referral|google|other)/.test(option));
}

function isConsentChoiceField(field) {
  if (!field || !isChoiceFieldType(field.type)) return false;
  const label = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!label) return false;
  if (/(gender|ethnicity|ethnic|race|disability|veteran|sexual orientation|date of birth|birth date|pronoun)/.test(label) &&
    !/(consent|acknowledge|privacy|policy|process|processing|data protection)/.test(label)) {
    return false;
  }
  if (isConsentChoiceLabelText(label)) return true;
  if (hasYesNoOptions(field.options)) return false;
  return (field.options || []).some((option, index) => scoreConsentChoiceOption(option, index) > 0);
}

function getConsentChoiceAnswer(field) {
  if (hasYesNoOptions(field.options)) return "Yes";
  const label = normalize([field?.question, field?.label, field?.name, field?.placeholder].filter(Boolean).join(" "));

  const choices = (field.options || [])
    .map((option) => String(option || "").trim())
    .filter(Boolean)
    .map((option, index) => ({ option, score: scoreConsentChoiceOption(option, index) }))
    .filter((item) => item.score > 0)
    .sort((left, right) => right.score - left.score);

  return choices[0]?.option || (isConsentChoiceLabelText(label) ? "consent" : "");
}

function isConsentChoiceLabelText(label) {
  return /(privacy|privacy notice|privacy policy|policy|terms|data protection|process.*(?:information|data)|processing.*(?:information|data)|consent.*(?:process|processing|data|information)|provide.*consent|require.*consent|giving.*consent|grant.*consent|acknowledg.*consent|equal opportunity|monitoring|analysis|strictest confidence|voluntary.*form|form.*voluntary|please select yes no.*consent)/.test(label);
}

function scoreConsentChoiceOption(option, index) {
  const text = normalize(option);
  if (!text || isPlaceholderChoice(text, text)) return -100;
  if (/\b(no|decline|do not|dont|not agree|disagree|withdraw|reject)\b/.test(text)) return -200;

  let score = 0;
  if (/\b(yes|agree|accept|accepted|acknowledge|confirm)\b/.test(text)) score += 300;
  if (/(grant|give|provide).{0,30}consent|consent.{0,30}(processing|process|recruitment|data|personal data)/.test(text)) score += 420;
  if (/(privacy|policy|notice|terms|data protection|processing.*data|personal data)/.test(text)) score += 220;
  if (score > 0) score += Math.max(0, 20 - index);
  return score;
}

function isConsentChoiceAlreadyAccepted(current) {
  const text = normalize(current);
  if (!text) return false;
  if (text === "yes") return true;
  if (/\b(no|decline|do not|dont|not agree|disagree|withdraw|reject)\b/.test(text)) return false;
  return /(agree|accept|accepted|acknowledge|confirm|consent|privacy|policy|data protection|processing.*data|personal data)/.test(text);
}

function isChoiceFieldType(type) {
  return CHOICE_FIELD_TYPES.includes(type);
}

function hasYesNoOptions(options) {
  const normalized = (options || []).map(normalize);
  return normalized.includes("yes") && normalized.includes("no");
}

function locationAnswerMatchesQuestion(locationAnswer, label) {
  const questionLocations = getQuestionLocationAliases(label);
  if (questionLocations.length === 0) return false;
  const answer = normalize(locationAnswer);
  const answerAliases = new Set([answer, ...buildLocationAliases(answer, { includeCountryRegions: true })]);

  return questionLocations.some((location) => {
    const normalizedLocation = normalize(location);
    return answerAliases.has(normalizedLocation) ||
      (normalizedLocation.length > 1 && answer.includes(normalizedLocation)) ||
      (answer.length > 1 && normalizedLocation.includes(answer)) ||
      scoreChoice(answer, normalizedLocation) >= 70 ||
      scoreChoice(normalizedLocation, answer) >= 70;
  });
}

function getQuestionLocationAliases(label) {
  const text = normalize(label);
  const aliases = [];

  EU_COUNTRY_NAMES.forEach((country) => {
    if (text.includes(country)) aliases.push(country);
  });

  if (text.includes("united states") || /\busa\b|\bus\b/.test(text)) aliases.push("united states", "usa", "us");
  if (text.includes("united kingdom") || /\buk\b/.test(text)) aliases.push("united kingdom", "uk");
  if (/\beu\b|europe|european union/.test(text)) aliases.push("eu", "europe", "european union");

  return Array.from(new Set(aliases));
}

function getProfileLanguages(staticFields) {
  const raw = getProfileStaticValue(staticFields, "languages");
  return getLanguageAliasesFromText(raw);
}

function getQuestionLanguageAliases(label) {
  return getLanguageAliasesFromText(label);
}

function getLanguageChoiceAnswer(field, questionLanguages, knownLanguages = []) {
  const isEnglish = (questionLanguages || []).some((language) => normalize(language) === "english");
  const isKnownLanguage = isEnglish || (questionLanguages || []).some((language) => knownLanguages.includes(language));

  if (field?.type === "checkbox" && field.option) {
    const option = normalize(field.option);
    const isNoProficiency = /\b(none|nessuno|no proficiency|not applicable|n a)\b/.test(option);
    const isProfessional = /\b(c1|c2|professional|professionale|advanced|avanzato|proficient|fluent)\b/.test(option);
    const isNative = /\b(native|mother tongue|madrelingua)\b/.test(option);
    const isScaleOption = isNoProficiency || isProfessional || isNative ||
      /\b(a1|a2|b1|b2|basic|basico|intermediate|intermedio|elementary|beginner)\b/.test(option);

    if (isScaleOption) {
      if (!isKnownLanguage) return isNoProficiency ? "Yes" : "No";
      return isProfessional && !isNative ? "Yes" : "No";
    }
  }

  if (isLanguageProficiencyScaleField(field)) {
    return isKnownLanguage ? getPreferredLanguageProficiencyOption(field) : getLowestLanguageProficiencyOption(field);
  }
  return isKnownLanguage ? "Yes" : "No";
}

function getPreferredLanguageProficiencyOption(field) {
  const options = field.options || [];
  const preferred = options.find((option) => /\b(c1|professional|professionale|advanced|avanzato|proficient|fluent)\b/.test(normalize(option)));
  return preferred || options.find((option) => /\bc2\b/.test(normalize(option))) || "C1 Advanced";
}

function isLanguageProficiencyScaleField(field) {
  if (!field || !isChoiceFieldType(field.type)) return false;
  const options = (field.options || []).map(normalize).filter(Boolean);
  if (options.length === 0 || hasYesNoOptions(field.options)) return false;
  return options.some((option) => /\b(a1|a2|b1|b2|c1|c2)\b|elementary|beginner|intermediate|advanced|proficient|fluent|native/.test(option));
}

function isLanguageChoiceAlreadyAcceptable(field, current, answer) {
  const normalizedCurrent = normalize(current);
  const normalizedAnswer = normalize(answer);
  if (!normalizedCurrent) return false;
  if (normalizeComparableValue(normalizedCurrent) === normalizeComparableValue(normalizedAnswer)) return true;

  if (isLanguageProficiencyScaleField(field) && /english/.test(normalize(field.label))) {
    return /\b(c1|c2)\b|advanced|proficient|fluent|native/.test(normalizedCurrent);
  }

  return false;
}

function getLowestLanguageProficiencyOption(field) {
  const options = field.options || [];
  const scored = options
    .map((option) => ({ option, score: scoreLanguageProficiencyLevel(option) }))
    .filter((item) => Number.isFinite(item.score))
    .sort((a, b) => a.score - b.score);
  return scored[0]?.option || "";
}

function scoreLanguageProficiencyLevel(option) {
  const text = normalize(option);
  if (/\b(no|none|not applicable|n a)\b/.test(text)) return 0;
  if (/\ba1\b|beginner|basic/.test(text)) return 1;
  if (/\ba2\b|elementary/.test(text)) return 2;
  if (/\bb2\b|upper intermediate/.test(text)) return 4;
  if (/\bb1\b|intermediate/.test(text)) return 3;
  if (/\bc1\b|advanced/.test(text)) return 5;
  if (/\bc2\b|proficient|fluent|native/.test(text)) return 6;
  return NaN;
}

function getLanguageAliasesFromText(text) {
  const normalized = normalize(text);
  if (!normalized) return [];

  const found = [];
  for (const [language, aliases] of LANGUAGE_ALIASES) {
    if (aliases.some((alias) => hasPhrase(normalized, alias))) found.push(language);
  }

  return Array.from(new Set(found));
}

function hasPhrase(normalizedText, phrase) {
  const text = ` ${normalize(normalizedText)} `;
  const needle = ` ${normalize(phrase)} `;
  return needle.trim() ? text.includes(needle) : false;
}

async function applyHoveredDropdownSelection(fields, skipFilledIds = new Set()) {
  const filledIds = new Set();
  const fieldId = getHoveredAutoBidFieldId();
  if (!fieldId) return { filled: 0, missed: 0, filledIds };
  if (skipFilledIds.has(fieldId)) return { filled: 0, missed: 0, filledIds };

  const field = fields.find((item) => item.id === fieldId);
  if (!field || !CHOICE_FIELD_TYPES.includes(field.type) || !shouldUsePositiveDropdownFallback(field)) {
    return { filled: 0, missed: 0, filledIds };
  }

  const controls = getControlsByFieldId(field.id);
  if (controls.length === 0) return { filled: 0, missed: 0, filledIds };
  if (hasFieldCurrentValue(field, controls)) return { filled: 0, missed: 0, filledIds };

  showStatus(`Selecting hovered dropdown: ${shortText(field.label)}`, "success");
  const positiveChoice = await findMostPositiveChoice(controls, field);
  if (!positiveChoice) {
    showStatus(`Could not open hovered dropdown options: ${shortText(field.label)}`, "error");
    return { filled: 0, missed: 1, filledIds };
  }

  const selected = await setChoiceValue(controls, positiveChoice);
  traceAutoBid("dropdown:hover-result", {
    field_id: field.id,
    label: field.label,
    selected,
    choice: summarizeChoice(positiveChoice),
    current: getCurrentChoiceSummary(controls)
  });
  if (!selected) {
    showStatus(`Could not select option: ${shortText(positiveChoice.text || positiveChoice.value)}`, "error");
    await waitForDropdownSettled(controls[0]);
    return { filled: 0, missed: 1, filledIds };
  }

  await waitForDropdownSettled(controls[0]);
  filledIds.add(field.id);
  showStatus(`Selected ${shortText(positiveChoice.text || positiveChoice.value)}`, "success");
  return { filled: 1, missed: 0, filledIds };
}

async function applyPositiveDropdownFallbacks(fields, filledIds) {
  let filled = 0;
  let missed = 0;
  const localFilledIds = new Set();
  const orderedFields = prioritizeHoveredField(fields);
  const dropdownCount = orderedFields.filter((field) => CHOICE_FIELD_TYPES.includes(field.type) && !filledIds.has(field.id) && shouldUsePositiveDropdownFallback(field)).length;
  let dropdownIndex = 0;

  for (const field of orderedFields) {
    if (filledIds.has(field.id) || !CHOICE_FIELD_TYPES.includes(field.type)) continue;
    if (!shouldUsePositiveDropdownFallback(field)) continue;
    dropdownIndex += 1;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    if (field.type === "combobox") {
      showStatus(`Opening dropdown ${dropdownIndex}/${dropdownCount}: ${shortText(field.label)}`, "success");
    }

    const positiveChoice = await findMostPositiveChoice(controls, field);
    if (!positiveChoice) {
      if (field.type === "combobox") {
        showStatus(`No selectable option found: ${shortText(field.label)}`, "error");
      }
      missed += 1;
      continue;
    }

    if (await setChoiceValue(controls, positiveChoice)) {
      traceAutoBid("dropdown:positive-result", {
        field_id: field.id,
        label: field.label,
        selected: true,
        choice: summarizeChoice(positiveChoice),
        current: getCurrentChoiceSummary(controls)
      });
      filled += 1;
      localFilledIds.add(field.id);
    } else {
      traceAutoBid("dropdown:positive-result", {
        field_id: field.id,
        label: field.label,
        selected: false,
        choice: summarizeChoice(positiveChoice),
        current: getCurrentChoiceSummary(controls)
      });
      missed += 1;
    }

    if (field.type === "combobox") {
      await waitForDropdownSettled(controls[0]);
      await sleep(DROPDOWN_SETTLE_MS);
    }
  }

  return { filled, missed, filledIds: localFilledIds };
}

async function applyPositiveCheckboxFallbacks(fields, filledIds) {
  let filled = 0;
  let missed = 0;
  const localFilledIds = new Set();

  for (const field of fields) {
    if (filledIds.has(field.id) || field.type !== "checkbox") continue;
    if (!shouldUsePositiveCheckboxFallback(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    if (hasCurrentChoiceValue(controls)) {
      localFilledIds.add(field.id);
      continue;
    }

    const selected = await setControlsValue(controls, "Yes");
    traceAutoBid("checkbox:positive-result", {
      field_id: field.id,
      label: field.label,
      selected,
      current: getCurrentChoiceSummary(controls)
    });

    if (selected) {
      filled += 1;
      localFilledIds.add(field.id);
    } else {
      missed += 1;
    }
  }

  return { filled, missed, filledIds: localFilledIds };
}

let deterministicDefaultsRuntime = null;

function getDeterministicDefaultsRuntime() {
  if (deterministicDefaultsRuntime) return deterministicDefaultsRuntime;
  const factory = window.AutoBidDeterministicDefaults;
  if (!factory?.create) return null;

  deterministicDefaultsRuntime = factory.create({
    cleanLabel,
    dispatchInput,
    dispatchRealisticMouseClick,
    getControlType,
    getControlsByFieldId,
    getCurrentChoiceSummary,
    getFieldContainer,
    getNearbyText,
    getVisualQuestionLabel,
    isChoiceFieldType,
    isVisible,
    nativeClickElement,
    normalize,
    scrollElementIntoView,
    send,
    setControlsValue,
    setNativeValue,
    sleep,
    traceAutoBid
  });

  return deterministicDefaultsRuntime;
}

async function applyDeterministicDefaults(fields, filledIds) {
  const runtime = getDeterministicDefaultsRuntime();
  if (runtime) return runtime.apply(fields, filledIds);

  let filled = 0;
  let missed = 0;
  const localFilledIds = new Set();

  for (const field of fields) {
    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    const fallback = getDeterministicDefault(field, controls);
    if (!fallback) continue;
    if (filledIds.has(field.id) && !shouldForceDeterministicDefault(field, controls, fallback)) continue;

    const current = getCurrentChoiceSummary(controls);
    const alreadyApplied = getControlType(controls[0]) === "range"
      ? isRangeValueApplied(controls[0], fallback.value)
      : normalizeComparableValue(current) === normalizeComparableValue(fallback.value);
    if (alreadyApplied) {
      localFilledIds.add(field.id);
      continue;
    }

    const candidateValues = uniqueValues([fallback.value, ...(fallback.values || [])]);
    let selected = false;
    for (const value of candidateValues) {
      selected = await setDeterministicDefaultValue(controls, value, fallback);
      if (selected) break;
    }
    traceAutoBid("default:applied", {
      field_id: field.id,
      label: field.label,
      reason: fallback.reason,
      value: fallback.value,
      candidates: candidateValues,
      selected,
      current: getCurrentChoiceSummary(controls)
    });

    if (selected) {
      filled += 1;
      localFilledIds.add(field.id);
    } else {
      missed += 1;
    }
  }

  return { filled, missed, filledIds: localFilledIds };
}

function getDeterministicDefault(field, controls) {
  if (isDatabaseChoiceField(field, controls)) {
    return { value: "PostgreSQL", reason: "database-default" };
  }

  if (isAvailabilityDateField(field, controls)) {
    return { value: formatDateForControl(getNextMondayDate(), controls[0], field), reason: "next-monday" };
  }

  const experienceDefault = getExperienceYearsDefault(field, controls);
  if (experienceDefault) {
    return experienceDefault;
  }

  return null;
}

function isDatabaseChoiceField(field, controls) {
  if (!isChoiceFieldType(field.type)) return false;
  const label = getFieldContextLabel(field, controls?.[0]);
  const options = (field.options || []).map(normalize);
  return /(database|databases|\bdb\b|data store|datastore)/.test(label) &&
    options.some((option) => /(postgresql|postgres|postgre sql)/.test(option));
}

function isAuthorizationSupportRequiredChoiceField(field, controls) {
  if (!isChoiceFieldType(field.type)) return false;
  if (!hasYesNoOptions(field.options)) return false;
  const label = getFieldContextLabel(field, controls?.[0]);
  return isAuthorizationSupportRequiredText(label);
}

function isAuthorizationSupportRequiredText(value) {
  const text = normalize(value);
  if (!text) return false;
  if (/(authorized|authorised|eligible|legally).*(work|employ)|(work|employ).*(authorized|authorised|eligible|legally)/.test(text)) return false;
  if (/without.{0,80}(sponsor|sponsorship|visa|work permit|support)/.test(text)) return false;
  const supportNoun = /(authorization|authorisation|sponsor|sponsorship|visa|work permit|work authorization support|work authorisation support)/;
  const requireVerb = /(require|need|needs|needed|seek|seeking|request|support|depend|dependent)/;
  return (requireVerb.test(text) && supportNoun.test(text)) ||
    /(now|future).{0,80}(authorization|authorisation|sponsor|sponsorship|visa|work permit)/.test(text) ||
    /(authorization|authorisation|sponsor|sponsorship|visa|work permit).{0,80}(now|future|support|required|needed)/.test(text);
}

function isAvailabilityDateField(field, controls) {
  const control = controls?.[0];
  const type = getControlType(control);
  if (!["date", "text", "search"].includes(type)) return false;
  const label = getDirectFieldLabel(field, control);
  return /(date available|available date|available start date|earliest.*start|start date|when.*start|availability date)/.test(label);
}

function getExperienceYearsDefault(field, controls) {
  const control = controls?.[0];
  const type = getControlType(control);
  const label = getDirectFieldLabel(field, control);
  if (!isExperienceYearsLabel(label)) return null;

  const kind = getExperienceYearsKind(label);
  const targetYears = getExperienceDefaultTargetYears(kind);
  const reason = `experience-years-${kind}`;

  if (isChoiceFieldType(type)) {
    const bestOption = findBestExperienceYearsOption(field.options || [], targetYears);
    const values = buildExperienceChoiceValueCandidates(targetYears, bestOption);
    return {
      value: values[0],
      values,
      reason,
      target_years: targetYears
    };
  }

  if (["range", "number", "text"].includes(type)) {
    const textEntryTargetYears = getTextEntryExperienceDefaultTargetYears(label, kind);
    return { value: getDefaultExperienceYears(control, textEntryTargetYears), reason, target_years: textEntryTargetYears };
  }

  return null;
}

function isExperienceValueField(field, controls) {
  const control = controls?.[0];
  const type = getControlType(control);
  if (!["range", "number", "text"].includes(type)) return false;
  const label = getDirectFieldLabel(field, control);
  return isExperienceYearsLabel(label);
}

function getDirectFieldLabel(field, control) {
  return normalize([
    field?.question,
    field?.label,
    field?.name,
    field?.placeholder,
    field?.autocomplete,
    control?.getAttribute?.("aria-label"),
    control?.getAttribute?.("placeholder")
  ].filter(Boolean).join(" "));
}

function isExperienceYearsLabel(label) {
  return /(how many.*years.*experience|years.*professional.*experience|years.*experience|experience.*years)/.test(label);
}

function getExperienceYearsKind(label) {
  const text = normalize(label);
  if (isDomainExperienceText(text)) return "domain";
  if (isTechSkillExperienceText(text)) return "tech";
  return "general";
}

function getExperienceDefaultTargetYears(kind) {
  if (kind === "domain") return 7;
  if (kind === "tech") return 9;
  return 10;
}

function getTextEntryExperienceDefaultTargetYears(_label, _kind) {
  return 7;
}

function isDomainExperienceText(text) {
  return /(backend development|frontend development|front end development|full stack development|software development|web development|mobile development|domain|industry|sector|e commerce|ecommerce|commerce|d2c|direct to consumer|b2b|b2c|fintech|financial tech|banking|finance|payments|igaming|gaming|retail|marketplace|healthcare|health care|medtech|edtech|insurtech|proptech|martech|adtech|travel|hospitality|logistics|media|saas)/.test(text);
}

function isTechSkillExperienceText(text) {
  return /(python|react|node|node js|nestjs|nest js|next js|javascript|typescript|java\b|kotlin|spring|c sharp|c#|\.net|dotnet|php|ruby|rails|go\b|golang|rust|scala|aws|azure|gcp|cloud|api|graphql|rest|sql|postgres|postgresql|mysql|mongodb|redis|docker|kubernetes|k8s|terraform|angular|vue|svelte|frontend|front end|mobile|android|ios|react native)/.test(text);
}

function findBestExperienceYearsOption(options, targetYears) {
  const scored = (options || [])
    .map((option) => ({ option: String(option || "").trim(), score: scoreExperienceYearsOption(option, targetYears) }))
    .filter((item) => item.option && item.score > 0)
    .sort((left, right) => right.score - left.score);
  return scored[0]?.option || "";
}

function buildExperienceChoiceValueCandidates(targetYears, bestOption = "") {
  const target = String(targetYears);
  const candidates = [bestOption, target, `${target} years`];

  if (targetYears === 7) {
    candidates.push("7-9", "6-7", "5-7", "7+", "4-6");
  } else if (targetYears === 9) {
    candidates.push("7-9", "8-9", "8+", "9+", "6-9");
  } else {
    candidates.push("10+", "10 years", "10 or more", "8-10", "7-10");
  }

  return uniqueValues(candidates);
}

function scoreExperienceYearsOption(option, targetYears) {
  const range = parseExperienceYearsOption(option);
  if (!range) return 0;

  const span = Number.isFinite(range.max) ? Math.max(0, range.max - range.min) : 8;
  let score = 0;

  if (range.min <= targetYears && targetYears <= range.max) {
    score = 1000 - span * 8;
  } else if (Number.isFinite(range.max) && range.max <= targetYears) {
    score = 820 + range.max * 12 - (targetYears - range.max) * 24;
  } else if (range.min > targetYears) {
    score = 420 - (range.min - targetYears) * 45;
  }

  if (targetYears < 10 && range.min >= 10) score -= 260;
  return score;
}

function parseExperienceYearsOption(option) {
  const text = normalize(option);
  if (!text || isPlaceholderChoice(text, text)) return null;
  const numbers = text.match(/\d+(?:\.\d+)?/g)?.map(Number).filter(Number.isFinite) || [];
  if (numbers.length === 0) return null;

  if (/less than|under|below|fewer than|up to|at most/.test(text)) {
    return { min: 0, max: numbers[0] };
  }

  const hasPlus = /\+|plus|or more|more than|over|above|greater than|at least/.test(String(option || "").toLowerCase()) ||
    /or more|more than|over|above|greater than|at least/.test(text);
  if (numbers.length >= 2) {
    return { min: Math.min(numbers[0], numbers[1]), max: Math.max(numbers[0], numbers[1]) };
  }

  if (hasPlus) return { min: numbers[0], max: Infinity };
  return { min: numbers[0], max: numbers[0] };
}

function getFieldContextLabel(field, control) {
  return normalize([
    field.label,
    field.name,
    field.placeholder,
    control?.getAttribute?.("placeholder"),
    control ? getNearbyText(control) : "",
    control ? getVisualQuestionLabel(control) : ""
  ].filter(Boolean).join(" "));
}

async function setDeterministicDefaultValue(controls, value, fallback = null) {
  const first = controls[0];
  const type = getControlType(first);

  if (type === "range") {
    return setRangeValueWithVisibleEditor(first, value);
  }

  if (type === "number" || (type === "text" && String(fallback?.reason || "").startsWith("experience-years-"))) {
    return setNumberInputDefaultValue(first, value, controls);
  }

  if (type === "date") {
    await scrollElementIntoView(first, "center");
    setNativeValue(first, value);
    dispatchInput(first);
    await sleep(80);
    return normalizeComparableValue(getCurrentChoiceSummary(controls)) === normalizeComparableValue(value);
  }

  return setControlsValue(controls, value);
}

function shouldForceDeterministicDefault(field, controls, fallback) {
  if (!fallback || !isExperienceValueField(field, controls)) return false;
  const current = getCurrentChoiceSummary(controls) || field?.value || "";
  return isZeroLikeExperienceValue(current);
}

async function setNumberInputDefaultValue(input, value, controls) {
  await scrollElementIntoView(input, "center");

  if (!await nativeClickElement(input)) {
    dispatchRealisticMouseClick(input);
  }
  input.focus?.();
  await sleep(140);

  clearTextEntryValue(input);
  await sleep(80);

  if (await nativeTypeText(value, false)) {
    dispatchInput(input);
    input.blur?.();
    await sleep(500);
    if (normalizeComparableValue(getCurrentChoiceSummary(controls)) === normalizeComparableValue(value)) return true;
  }

  if (replaceTextEntryValue(input, value)) {
    await sleep(500);
    if (normalizeComparableValue(getCurrentChoiceSummary(controls)) === normalizeComparableValue(value)) return true;
  }

  setNativeValue(input, value);
  dispatchInput(input);
  input.blur?.();
  await sleep(500);
  return normalizeComparableValue(getCurrentChoiceSummary(controls)) === normalizeComparableValue(value);
}

async function setRangeValueWithVisibleEditor(control, value) {
  const editor = findRangeValueEditor(control);
  if (editor) {
    await scrollElementIntoView(editor, "center");
    if (!await nativeClickElement(editor)) {
      dispatchRealisticMouseClick(editor);
    }
    await sleep(160);

    const activeEditor = findActiveRangeTextEntry(control, editor);
    if (activeEditor && replaceTextEntryValue(activeEditor, value)) {
      await sleep(160);
      if (isRangeValueApplied(control, value)) return true;
    }

    if (await nativeTypeText(value)) {
      await sleep(220);
      if (isRangeValueApplied(control, value)) return true;
    }
  }

  await scrollElementIntoView(control, "center");
  setNativeValue(control, value);
  dispatchInput(control);
  await sleep(140);
  return isRangeValueApplied(control, value);
}

function findRangeValueEditor(control) {
  return getRangeValueCandidates(control)[0]?.element || null;
}

function findActiveRangeTextEntry(control, clickedElement) {
  const active = document.activeElement;
  if (isRangeTextEntry(active) && active !== control) return active;
  if (isRangeTextEntry(clickedElement) && clickedElement !== control) return clickedElement;

  const roots = getRangeSearchRoots(control);
  for (const root of roots) {
    const entry = Array.from(root.querySelectorAll("input:not([type='range']):not([type='hidden']), textarea, [contenteditable='true'], [role='spinbutton'], [role='textbox']"))
      .find((element) => isVisible(element) && isRangeTextEntry(element) && element !== control);
    if (entry) return entry;
  }

  return null;
}

function replaceTextEntryValue(element, value) {
  if (!element) return false;
  const tag = element.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA") {
    element.focus?.();
    safelySelectTextEntry(element);
    setNativeValue(element, value);
    dispatchInput(element);
    element.blur?.();
    return normalizeComparableValue(element.value) === normalizeComparableValue(value);
  }

  if (element.isContentEditable) {
    element.focus?.();
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.execCommand("insertText", false, value);
    dispatchInput(element);
    element.blur?.();
    return normalizeComparableValue(element.textContent) === normalizeComparableValue(value);
  }

  return false;
}

function clearTextEntryValue(element) {
  if (!element) return false;
  element.focus?.();
  safelySelectTextEntry(element);
  setNativeValue(element, "");
  dispatchInput(element);
  return String(element.value || "") === "";
}

function safelySelectTextEntry(element) {
  try {
    element.select?.();
    return true;
  } catch (_error) {
    // Number inputs reject select() in Chromium; native Ctrl+A still selects them.
  }

  try {
    const length = String(element.value || "").length;
    element.setSelectionRange?.(0, length);
    return true;
  } catch (_error) {
    return false;
  }
}

async function nativeTypeText(value, commit = true) {
  try {
    const result = await send("NATIVE_TYPE", { text: String(value), commit });
    return result?.typed === true;
  } catch (error) {
    traceAutoBid("native-type:failed", { message: error.message || String(error) });
    return false;
  }
}

function isRangeValueApplied(control, value) {
  const expected = normalizeComparableValue(value);
  const displayValue = getRangeDisplayValue(control);
  if (displayValue) return normalizeComparableValue(displayValue) === expected;
  return normalizeComparableValue(control.value) === expected;
}

function getRangeDisplayValue(control) {
  const candidate = getRangeValueCandidates(control)[0];
  return candidate?.value || "";
}

function getRangeValueCandidates(control) {
  const rangeRect = control.getBoundingClientRect();
  const selector = [
    "input:not([type='range']):not([type='hidden'])",
    "textarea",
    "[contenteditable='true']",
    "[role='spinbutton']",
    "[role='textbox']",
    "output",
    "span",
    "div",
    "p",
    "strong",
    "b",
    "button"
  ].join(",");

  const candidates = getRangeSearchRoots(control)
    .flatMap((root) => Array.from(root.querySelectorAll(selector)))
    .filter((element, index, list) => list.indexOf(element) === index)
    .map((element) => scoreRangeValueCandidate(element, control, rangeRect))
    .filter(Boolean)
    .sort((left, right) => right.score - left.score);

  return candidates;
}

function getRangeSearchRoots(control) {
  const roots = [];
  let current = control.parentElement;
  for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
    roots.push(current);
  }
  const container = getFieldContainer(control);
  if (container) roots.push(container);
  return Array.from(new Set(roots)).filter(Boolean);
}

function scoreRangeValueCandidate(element, control, rangeRect) {
  if (!element || element === control || element.contains(control) || !isVisible(element)) return null;
  if (!isRangeTextEntry(element) && element.querySelector("input, textarea, select, button, [role='slider']")) return null;

  const value = getRangeCandidateValue(element);
  if (!/^\d+(?:\.\d+)?$/.test(value)) return null;

  const rect = element.getBoundingClientRect();
  const verticalDistance = Math.abs((rect.top + rect.bottom) / 2 - rangeRect.top);
  const horizontallyNear = rect.right >= rangeRect.left - 80 && rect.left <= rangeRect.right + 80;
  if (!horizontallyNear || rect.bottom < rangeRect.top - 260 || rect.top > rangeRect.bottom + 180) return null;

  let score = 1000 - verticalDistance;
  if (isRangeTextEntry(element)) score += 500;
  if (rect.bottom <= rangeRect.top + 40) score += 120;
  if (normalizeComparableValue(value) === normalizeComparableValue(control.value)) score += 80;
  return { element, value, score };
}

function isRangeTextEntry(element) {
  if (!element || !element.matches) return false;
  if (element.matches("textarea, [contenteditable='true'], [role='spinbutton'], [role='textbox']")) return true;
  if (element.tagName !== "INPUT") return false;
  const type = String(element.getAttribute("type") || "text").toLowerCase();
  return !["range", "hidden", "checkbox", "radio", "submit", "button", "reset", "file"].includes(type);
}

function getRangeCandidateValue(element) {
  if (element.matches("input, textarea")) return cleanLabel(element.value || "");
  const ownText = Array.from(element.childNodes || [])
    .filter((node) => node.nodeType === Node.TEXT_NODE)
    .map((node) => node.textContent)
    .join(" ");
  return cleanLabel(ownText || (element.children.length === 0 ? element.textContent : ""));
}

function getDefaultExperienceYears(control, preferredYears = 10) {
  const min = parseFiniteNumber(control?.min, 0);
  const max = parseFiniteNumber(control?.getAttribute?.("max") || control?.max, NaN);
  const step = parseFiniteNumber(control?.step, 1);
  let value = preferredYears;

  if (value < min) value = min;
  if (Number.isFinite(max) && value > max) value = max;
  if (Number.isFinite(step) && step > 0) {
    value = min + Math.round((value - min) / step) * step;
  }

  return formatNumberValue(value);
}

function parseFiniteNumber(value, fallback) {
  if (value == null || String(value).trim() === "") return fallback;
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function formatNumberValue(value) {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(2)));
}

function uniqueValues(values) {
  return Array.from(new Set((values || []).map((value) => String(value || "").trim()).filter(Boolean)));
}

function getNextMondayDate(now = new Date()) {
  const daysUntilMonday = ((8 - now.getDay()) % 7) || 7;
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysUntilMonday);
}

function formatDateForControl(date, control, field) {
  const type = getControlType(control);
  if (type === "date") return formatIsoDate(date);

  const hint = normalize([field.placeholder, control?.getAttribute?.("placeholder"), field.label].filter(Boolean).join(" "));
  if (/dd mm yyyy|dd\/mm\/yyyy|dd-mm-yyyy/.test(hint)) return formatSlashDate(date, "DMY");
  if (/yyyy mm dd|yyyy-mm-dd/.test(hint)) return formatIsoDate(date);
  return formatSlashDate(date, "MDY");
}

function formatIsoDate(date) {
  return [
    date.getFullYear(),
    pad2(date.getMonth() + 1),
    pad2(date.getDate())
  ].join("-");
}

function formatSlashDate(date, order) {
  const month = pad2(date.getMonth() + 1);
  const day = pad2(date.getDate());
  const year = String(date.getFullYear());
  return order === "DMY" ? `${day}/${month}/${year}` : `${month}/${day}/${year}`;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function normalizeComparableValue(value) {
  return normalize(String(value || "").replace(/\.0+$/, ""));
}

function hasCurrentChoiceValue(controls) {
  const first = controls[0];
  const type = getControlType(first);

  if (type === "radio") return controls.some(isRadioChecked);
  if (type === "checkbox") return isCheckboxChecked(first);
  if (type === "button-group") return Boolean(getSelectedChoiceButtonLabel(controls));
  if (first.tagName === "SELECT") {
    const option = first.selectedOptions?.[0];
    if (!option) return false;
    return Boolean(first.value && !isPlaceholderChoice(option.textContent, option.value));
  }
  if (type === "combobox") {
    const selectedText = getComboboxSelectedText(first);
    if (selectedText && !isPlaceholderChoice(selectedText, selectedText)) return true;

    const shell = getComboboxShell(first);
    const hiddenValue = Array.from(shell?.querySelectorAll?.("input[type='hidden']") || [])
      .map((input) => cleanLabel(input.value || ""))
      .find((value) => value && !isPlaceholderChoice(value, value));
    if (hiddenValue) return true;

    const input = getComboboxInput(first);
    const isSearchableChoice = Boolean(
      input.getAttribute?.("aria-autocomplete") ||
      input.getAttribute?.("aria-controls") ||
      getReactSelectInstanceId(first)
    );
    if (isSearchableChoice) return false;

    const value = cleanLabel(input.value || first.value || first.textContent || "");
    return Boolean(value && !isPlaceholderChoice(value, value));
  }

  return Boolean(String(first.value || "").trim());
}

function getCurrentChoiceSummary(controls) {
  const first = controls?.[0];
  if (!first) return "";
  const type = getControlType(first);

  if (type === "radio") {
    const checked = controls.find(isRadioChecked);
    return checked ? getRadioOptionLabel(checked) || checked.value || "" : "";
  }

  if (type === "checkbox") {
    return getSelectedCheckboxBooleanChoiceLabel(first) || (isCheckboxChecked(first) ? "Yes" : "");
  }

  if (type === "button-group") {
    return getSelectedChoiceButtonLabel(controls);
  }

  if (first.tagName === "SELECT") {
    const option = first.selectedOptions?.[0];
    return cleanLabel(option ? `${option.textContent || ""} ${option.value || ""}` : "");
  }

  if (type === "combobox") {
    return cleanLabel(getComboboxSelectedText(first) || first.value || first.textContent || "");
  }

  return cleanLabel(first.value || "");
}

function shouldUsePositiveDropdownFallback(field) {
  if (shouldDeferChoiceFieldToRuntimeGpt(field)) return false;
  if (isBasedInLocationField(field)) return false;
  if (isLanguageChoiceField(field)) return false;
  if (isAuthorizationSupportRequiredText(field.label)) return true;
  const label = normalize(field.label);
  if (/have you (?:ever )?used .+ before|used .+ previously/.test(label)) return false;
  return /(experience|years|level|proficiency|skill|expertise|knowledge|familiar|comfortable|rating|seniority|sponsor|sponsorship|visa|authorization|authorisation|authorized|eligible|willing|available|open to|agree|accept|consent|confirm|legally|relocat|remote|can you|able to|do you have|have you|do you use|have you used|use.*ai|ai.*assist|artificial intelligence|development workflow|meet.*requirement|salary|compensation)/.test(label);
}

function shouldDeferChoiceFieldToRuntimeGpt(field) {
  if (!field?.required || !CHOICE_FIELD_TYPES.includes(field.type)) return false;
  if (isSensitiveGeneratedAnswerField(field) || isLanguageChoiceField(field) || isConsentChoiceField(field) || isBasedInLocationField(field)) {
    return false;
  }
  const atsId = atsAdapters?.describe?.().id || "";
  const host = location.hostname.toLowerCase();
  if (atsId === "newrocket" || host === "newrocket.com" || host.endsWith(".newrocket.com")) return true;

  const text = normalize([field.question, field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(experience|years|skill|proficien|expertise|knowledge|familiar|technology|framework|programming|software engineering|frontend|front end|backend|back end|full stack|mobile development|react|typescript|javascript|python|java|golang|cloud|aws|azure|gcp|docker|kubernetes|terraform|payment|fiscalization|database|sponsor|sponsorship|visa|work permit|right to work|authorized|authorised|eligible|legally work)/.test(text);
}

function shouldUsePositiveCheckboxFallback(field) {
  const label = normalize([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!label) return false;
  if (isFileUploadChoiceText(label)) return false;
  if (isSensitiveOrPersonalChoiceCheckbox(label)) return false;
  if (/(terms|privacy|policy|consent|agree|accept|acknowledge|confirm|certify|accurate|contact me|future job opportunit|future opportunit|talent community|job alert|keep my data|retain my data|store my data|process my data|data processing|email me|reach out|recruiting communication|recruitment communication|consider me for future)/.test(label)) return true;
  return false;
}

function isSensitiveOrPersonalChoiceCheckbox(label) {
  return /(pronoun|he him|she her|they them|xe xem|ze hir|ey em|hir hir|fae faer|hu hu|use name only|custom|gender|race|ethnicity|ethnic|disability|veteran|protected veteran|sexual orientation|date of birth|birth date|newsletter|marketing email|product update|event update)/.test(label);
}

async function setControlsValue(controls, value, field = null, context = {}) {
  if (!isActiveContentInstance()) return false;
  const first = controls[0];
  const type = getControlType(first);
  const textValue = String(value || "").trim();
  if (!textValue) return false;

  if (type === "radio") {
    return setRadioValue(controls, textValue);
  }

  if (type === "checkbox") {
    return setCheckboxValue(first, textValue, field);
  }

  if (type === "button-group") {
    return setButtonGroupValue(controls, textValue, field);
  }

  if (type === "contenteditable") {
    first.focus?.();
    first.textContent = textValue;
    first.dispatchEvent(new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      composed: true,
      inputType: "insertText",
      data: textValue
    }));
    first.dispatchEvent(new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
      data: textValue
    }));
    first.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
    return normalize(first.textContent || "") === normalize(textValue);
  }

  if (first.tagName === "SELECT") {
    const option = findBestChoice(
      Array.from(first.options)
        .filter((item) => !item.disabled)
        .map((item) => ({
          control: item,
          text: item.textContent,
          value: item.value
        })),
      textValue
    );
    if (option) {
      first.value = option.control.value;
      dispatchInput(first);
      return true;
    }
    return false;
  }

  if (type === "combobox") {
    return setComboboxValue(first, textValue, field, context);
  }

  const requiresAutocompleteSelection = shouldSelectTextAutocompleteOption(first, field);
  const country = String(context.country || "").trim();
  if (requiresAutocompleteSelection &&
      context.locationCountryFirst &&
      country &&
      normalize(country) !== normalize(textValue)) {
    first.focus?.();
    setNativeValue(first, country);
    dispatchInput(first);
    traceAutoBid("location-autocomplete:country-prime", {
      field_id: field?.id || first.dataset?.autoBidFieldId || "",
      label: field?.question || field?.label || "",
      country
    });
    await sleep(LOCATION_AUTOCOMPLETE_WAIT_MS);
  }
  first.focus?.();
  setNativeValue(first, textValue);
  dispatchInput(first);
  if (requiresAutocompleteSelection && !await maybeSelectTextAutocompleteOption(first, textValue, field, context)) {
    await clearControlsForRetry(controls);
    return false;
  }
  if (requiresAutocompleteSelection) return true;

  if (await waitForTextControlValue(first, textValue, 900)) return true;
  first.focus?.();
  setNativeValue(first, textValue);
  first.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText", data: textValue }));
  first.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
  const retained = await waitForTextControlValue(first, textValue, 1200);
  if (!retained) {
    traceAutoBid("text-value:not-retained", {
      field_id: field?.id || first.dataset?.autoBidFieldId || "",
      label: field?.question || field?.label || "",
      expected: shortText(textValue),
      current: shortText(first.value || ""),
      connected: first.isConnected
    });
  }
  return retained;
}

async function waitForTextControlValue(control, expectedValue, timeoutMs) {
  const expected = String(expectedValue || "").trim();
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (areTextControlValuesEquivalent(control, getControlValue(control), expected)) return true;
    await sleep(80);
  }
  return areTextControlValuesEquivalent(control, getControlValue(control), expected);
}

function areTextControlValuesEquivalent(control, currentValue, expectedValue) {
  const current = String(currentValue || "").trim();
  const expected = String(expectedValue || "").trim();
  if (!expected) return false;
  const phoneLike = String(control?.type || "").toLowerCase() === "tel" ||
    /\btel\b|phone|mobile/.test(normalize([
      control?.getAttribute?.("autocomplete"),
      control?.getAttribute?.("aria-label"),
      control?.name,
      control?.id
    ].filter(Boolean).join(" ")));
  if (phoneLike) {
    const currentDigits = current.replace(/\D/g, "");
    const expectedDigits = expected.replace(/\D/g, "");
    return Boolean(expectedDigits) && currentDigits === expectedDigits;
  }
  return normalize(current) === normalize(expected);
}

async function clearControlsForRetry(controls) {
  const first = controls?.[0];
  if (!first) return;
  const type = getControlType(first);
  if (type === "combobox") await closeCombobox(first);

  const target = type === "combobox" ? getComboboxInput(first) : first;
  if (target?.matches?.("input, textarea") && !target.readOnly && !target.disabled) {
    target.focus?.();
    setNativeValue(target, "");
    dispatchInput(target);
    target.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
    target.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
    target.blur?.();
  }
  await sleep(180);
}

async function maybeSelectTextAutocompleteOption(control, value, field, context = {}) {
  if (!shouldSelectTextAutocompleteOption(control, field)) return false;
  const waitMs = context.locationCountryFirst &&
    normalize(context.country || "") === normalize(value)
    ? LOCATION_AUTOCOMPLETE_WAIT_MS
    : 350;
  await sleep(waitMs);

  const options = getTextAutocompleteOptions(control);
  if (options.length === 0) {
    pressAutocompleteKey(control, "ArrowDown");
    await sleep(300);
  }

  const choices = getTextAutocompleteOptions(control)
    .map((option, index) => ({
      control: option,
      text: option.textContent || option.getAttribute("aria-label") || option.getAttribute("data-value") || "",
      value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent || "",
      index
    }));
  const match = findBestAutocompleteChoice(choices, value, field, context);
  if (!match?.control) {
    return false;
  }

  await clickChoice(match.control);
  await sleep(300);
  return true;
}

function shouldSelectTextAutocompleteOption(control, field) {
  if (!control?.matches?.("input, textarea")) return false;
  const text = normalize([
    field?.question,
    field?.label,
    field?.name,
    field?.placeholder,
    control.getAttribute("aria-label"),
    control.getAttribute("placeholder"),
    control.getAttribute("autocomplete"),
    control.name,
    control.id
  ].filter(Boolean).join(" "));
  if (!/(city|country|location|located|residence|address|state|province|region)/.test(text)) return false;
  return Boolean(
    control.getAttribute("aria-autocomplete") ||
    control.getAttribute("aria-controls") ||
    control.getAttribute("aria-owns") ||
    control.getAttribute("role") === "combobox" ||
    /combobox|autocomplete|autosuggest|typeahead|places/.test(control.className || "")
  );
}

function getTextAutocompleteOptions(control) {
  const roots = getComboboxRoots(control);
  const selector = [
    "[role='option']",
    "[role='menuitem']",
    "[data-value]",
    "[cmdk-item]",
    "[class*='option' i]",
    "[class*='suggestion' i]",
    "[class*='autocomplete' i] li",
    "[class*='typeahead' i] li",
    "li"
  ].join(",");
  return collectChoiceElements(roots, selector, control)
    .filter((option) => !isPlaceholderChoice(option.textContent || "", option.getAttribute("data-value") || option.getAttribute("value") || ""));
}

function pressAutocompleteKey(control, key) {
  const normalizedKey = key === " " ? " " : key;
  const code = key === " " ? "Space" : key;
  const keyCode = key === "ArrowDown" ? 40 : key === "Escape" ? 27 : 13;
  control.focus?.();
  control.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: normalizedKey, code, keyCode, which: keyCode }));
  control.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: normalizedKey, code, keyCode, which: keyCode }));
}

async function setRadioValue(controls, value) {
  const match = findBestChoice(
    controls.map((control) => ({
      control,
      text: getRadioOptionLabel(control),
      value: control.value
    })),
    value
  );
  if (!match) return false;

  const radio = match.control;
  const targets = Array.from(new Set([
    getRadioClickTarget(radio),
    radio.closest?.("label, [class*='option' i], li, [role='radio'], [role='option']"),
    radio
  ].filter(Boolean)));

  for (const target of targets) {
    if (await activateChoiceTarget(target, () => isRadioChecked(radio))) return true;
  }

  if (!isRadioChecked(radio) && radio.matches?.("input[type='radio']")) {
    setNativeChecked(radio, true);
    dispatchInput(radio);
    await sleep(120);
  }

  return isRadioChecked(radio);
}

async function activateChoiceTarget(target, isApplied) {
  if (!target || !isVisible(target)) return false;
  await scrollElementIntoView(target, "center");

  await nativeClickElement(target);
  await sleep(140);
  if (isApplied()) return true;

  dispatchRealisticMouseClick(target);
  await sleep(140);
  if (isApplied()) return true;

  for (const key of [" ", "Enter"]) {
    target.focus?.();
    runPageCommand("key", target, { key });
    const code = key === " " ? "Space" : "Enter";
    const keyCode = key === " " ? 32 : 13;
    target.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, composed: true, key, code, keyCode, which: keyCode }));
    target.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, composed: true, key, code, keyCode, which: keyCode }));
    await sleep(140);
    if (isApplied()) return true;
  }

  return false;
}

function getRadioClickTarget(radio) {
  const forLabel = radio.id ? queryOne(`label[for="${cssEscape(radio.id)}"]`, radio.getRootNode?.() || document) : null;
  if (forLabel && isVisible(forLabel)) return forLabel;

  const closestLabel = radio.closest("label");
  if (closestLabel && isVisible(closestLabel)) return closestLabel;

  const option = radio.closest("[class*='option' i], li, [role='radio'], [role='option']");
  if (option && isVisible(option)) return option;

  return radio;
}

async function setCheckboxValue(control, value, field = null) {
  const shouldCheck = parseBooleanAnswer(value);
  if (shouldCheck === null) return false;

  const booleanChoices = getCheckboxBooleanChoiceOptions(control);
  const desiredChoice = booleanChoices.find((choice) => parseBooleanAnswer(choice.text) === shouldCheck);
  if (desiredChoice) {
    if (isCheckboxBooleanChoiceSelected(desiredChoice)) {
      control.dataset.autoBidExplicitBooleanValue = shouldCheck ? "Yes" : "No";
      return true;
    }

    let controlStateEventObserved = false;
    const observeControlState = () => { controlStateEventObserved = true; };
    control.addEventListener("input", observeControlState);
    control.addEventListener("change", observeControlState);
    await activateChoiceTarget(desiredChoice.control, () => {
      const refreshed = getCheckboxBooleanChoiceOptions(control)
        .find((choice) => parseBooleanAnswer(choice.text) === shouldCheck);
      return Boolean(refreshed && isCheckboxBooleanChoiceSelected(refreshed)) ||
        (controlStateEventObserved && isCheckboxChecked(control) === shouldCheck);
    });
    control.removeEventListener("input", observeControlState);
    control.removeEventListener("change", observeControlState);

    const refreshedChoice = getCheckboxBooleanChoiceOptions(control)
      .find((choice) => parseBooleanAnswer(choice.text) === shouldCheck);
    const verifiedByChoiceState = Boolean(refreshedChoice && isCheckboxBooleanChoiceSelected(refreshedChoice));
    const verifiedByControlEvent = controlStateEventObserved && isCheckboxChecked(control) === shouldCheck;
    if (verifiedByChoiceState || verifiedByControlEvent) {
      control.dataset.autoBidExplicitBooleanValue = shouldCheck ? "Yes" : "No";
      return true;
    }
    return false;
  }

  const option = cleanLabel(field?.option || getCheckboxOptionLabel(control, field?.question || field?.label || ""));
  if (!option && shouldCheck === false) return false;
  if (isCheckboxChecked(control) === shouldCheck) {
    control.dataset.autoBidExplicitBooleanValue = shouldCheck ? "Yes" : "No";
    return true;
  }

  for (const target of getCheckboxClickTargets(control)) {
    if (await activateChoiceTarget(target, () => isCheckboxChecked(control) === shouldCheck)) {
      control.dataset.autoBidExplicitBooleanValue = shouldCheck ? "Yes" : "No";
      return true;
    }
  }

  const input = getCheckboxInput(control);
  if (input && isCheckboxChecked(input) !== shouldCheck) {
    setNativeChecked(input, shouldCheck);
    dispatchInput(input);
  } else if (!input && control.getAttribute("role") === "checkbox") {
    control.setAttribute("aria-checked", shouldCheck ? "true" : "false");
    dispatchInput(control);
  }
  await sleep(120);

  const applied = isCheckboxChecked(control) === shouldCheck;
  if (applied) control.dataset.autoBidExplicitBooleanValue = shouldCheck ? "Yes" : "No";
  return applied;
}

function getCheckboxBooleanChoiceOptions(control) {
  const fieldContainer = getFieldContainer(control);
  const roots = [];
  let current = control.parentElement;
  for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
    if (current.matches?.("form, main, [role='main']")) break;
    roots.push(current);
    if (current === fieldContainer) break;
  }
  if (fieldContainer && !roots.includes(fieldContainer) && !fieldContainer.matches?.("form, main, [role='main']")) {
    roots.push(fieldContainer);
  }

  const selector = "button, label, [role='button'], [role='radio'], [data-value], span, div";

  for (const root of Array.from(new Set(roots))) {
    const choices = [];
    for (const element of queryAll(selector, root)) {
      if (!isVisible(element) || element === control) continue;
      const text = cleanLabel(element.getAttribute?.("data-value") || element.getAttribute?.("aria-label") || element.textContent || "");
      if (!/^(yes|no)$/i.test(text)) continue;

      const target = element.closest?.("button, label, [role='button'], [role='radio'], [data-value]") || element;
      if (!root.contains?.(target) || !isVisible(target)) continue;
      choices.push({ control: target, text });
    }

    const uniqueChoices = choices.filter((choice, index, list) => list.findIndex((item) => item.control === choice.control) === index);
    const values = new Set(uniqueChoices.map((choice) => normalize(choice.text)));
    if (values.has("yes") && values.has("no")) return uniqueChoices;
  }

  return [];
}

function isCheckboxBooleanChoiceSelected(choice) {
  const target = choice?.control;
  if (!target) return false;
  const expectedText = normalize(choice.text);
  const stateNodes = [target];
  let parent = target.parentElement;
  for (let depth = 0; parent && depth < 2; depth += 1, parent = parent.parentElement) {
    if (normalize(parent.textContent || "") !== expectedText) break;
    stateNodes.push(parent);
  }

  for (const node of stateNodes) {
    if (isChoiceButtonSelected(node)) return true;
    const selectedDescendant = node.querySelector?.("[aria-pressed='true'], [aria-checked='true'], [aria-selected='true'], [data-selected='true'], [data-state='checked'], [data-state='selected'], .selected, .active, [class*='selected' i], [class*='active' i]");
    if (selectedDescendant) return true;
    const radio = node.matches?.("input[type='radio']") ? node : node.querySelector?.("input[type='radio']");
    if (radio?.checked) return true;
  }

  return false;
}

function getSelectedCheckboxBooleanChoiceLabel(control) {
  const selected = getCheckboxBooleanChoiceOptions(control).find(isCheckboxBooleanChoiceSelected);
  if (selected?.text) return selected.text;

  const stored = cleanLabel(control?.dataset?.autoBidExplicitBooleanValue || "");
  const storedBoolean = parseBooleanAnswer(stored);
  if (storedBoolean === null) return "";
  if (isCheckboxChecked(control) !== storedBoolean) {
    delete control.dataset.autoBidExplicitBooleanValue;
    return "";
  }
  return stored;
}

function getCheckboxClickTargets(control) {
  const targets = [];
  const input = getCheckboxInput(control);
  const root = input || control;
  const rootNode = root.getRootNode?.() || document;
  const forLabel = root.id ? queryOne(`label[for="${cssEscape(root.id)}"]`, rootNode) : null;
  const consentRow = isConsentCheckboxText(getConsentCheckboxLabel(root, root)) ? getConsentCheckboxRow(root) : null;

  [
    forLabel,
    root.closest?.("label"),
    consentRow,
    root.closest?.("[class*='checkbox' i]"),
    root.closest?.("[class*='consent' i], [class*='privacy' i], [class*='terms' i]"),
    root.closest?.("[class*='option' i]"),
    root.closest?.("li"),
    root.parentElement?.parentElement,
    root.parentElement,
    control
  ].forEach((target) => {
    if (target && isVisible(target)) targets.push(target);
  });

  return Array.from(new Set(targets));
}

function getCheckboxInput(control) {
  if (control?.matches?.("input[type='checkbox']")) return control;
  return control?.querySelector?.("input[type='checkbox']") || null;
}

function isCheckboxChecked(control) {
  const input = getCheckboxInput(control);
  if (input) return Boolean(input.checked);
  if (control?.matches?.("input[type='checkbox']")) return Boolean(control.checked);
  if (control?.getAttribute?.("aria-checked") === "true") return true;
  if (control?.getAttribute?.("data-state") === "checked") return true;
  return Boolean(control?.matches?.(".checked, .selected, .active, [class*='checked' i], [class*='selected' i], [class*='active' i]"));
}

async function setButtonGroupValue(controls, value, field = null) {
  const match = findBestChoice(
    controls.map((control) => ({
      control,
      text: getChoiceButtonLabel(control),
      value: getChoiceButtonValue(control)
    })),
    value
  );
  if (!match) return false;

  const applied = await activateChoiceTarget(match.control, () => {
    const liveControls = field?.id ? getControlsByFieldId(field.id) : controls;
    return generatedAnswerValuesMatch(getSelectedChoiceButtonLabel(liveControls), value);
  });
  if (applied) return true;

  dispatchInput(match.control);

  const started = Date.now();
  while (Date.now() - started < 900) {
    const liveControls = field?.id ? getControlsByFieldId(field.id) : controls;
    if (generatedAnswerValuesMatch(getSelectedChoiceButtonLabel(liveControls), value)) return true;
    await sleep(90);
  }
  return false;
}

async function setChoiceValue(controls, choice) {
  const first = controls[0];
  const type = getControlType(first);

  if (type === "combobox" && choice.control) {
    return selectComboboxChoice(first, choice);
  }

  return setControlsValue(controls, choice.value || choice.text);
}

function dispatchInput(control) {
  control.dispatchEvent(new Event("input", { bubbles: true }));
  control.dispatchEvent(new Event("change", { bubbles: true }));
}

function normalize(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function findBestChoice(choices, answer) {
  const normalizedAnswer = normalize(answer);
  if (!normalizedAnswer) return null;
  const answerAliases = buildAnswerAliases(normalizedAnswer);

  const scored = choices
    .map((choice) => {
      const normalizedText = normalize(choice.text);
      const normalizedValue = normalize(choice.value);
      const aliases = buildChoiceAliases(normalizedText, normalizedValue);
      const directScore = Math.max(scoreChoice(normalizedText, normalizedAnswer), scoreChoice(normalizedValue, normalizedAnswer));
      const aliasScore = Math.max(...aliases.flatMap((alias) => answerAliases.map((answerAlias) => scoreChoice(alias, answerAlias))));
      const score = Math.max(directScore >= 80 ? directScore + 20 : directScore, Math.min(aliasScore, 90));
      return { ...choice, score };
    })
    .filter((choice) => choice.score > 0)
    .sort((a, b) => b.score - a.score);

  return scored[0] || null;
}

async function findMostPositiveChoice(controls, field) {
  const first = controls[0];
  const choices = first.tagName === "SELECT"
    ? Array.from(first.options)
      .filter((option) => !option.disabled && !isPlaceholderChoice(option.textContent, option.value))
      .map((option) => ({
        text: option.textContent,
        value: option.value,
        index: option.index
      }))
    : getControlType(first) === "combobox"
      ? (await getComboboxChoices(first)).map((option, index) => ({
        control: option,
        text: option.textContent,
        value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent,
        index
      }))
    : getControlType(first) === "button-group"
      ? controls.map((control, index) => ({
        control,
        text: getChoiceButtonLabel(control),
        value: getChoiceButtonValue(control),
        index
      }))
    : controls.map((control, index) => ({
      text: getRadioOptionLabel(control),
      value: control.value,
      index
    }));

  const scored = choices
    .map((choice) => ({
      ...choice,
      score: scorePositiveChoice(field.label, `${choice.text} ${choice.value}`, choice.index)
    }))
    .sort((a, b) => b.score - a.score || b.index - a.index);

  traceAutoBid("dropdown:scored", {
    field_id: field.id,
    label: field.label,
    type: getControlType(first),
    current: getCurrentChoiceSummary(controls),
    choices: scored.map((choice) => ({
      text: cleanLabel(choice.text),
      value: cleanLabel(choice.value),
      index: choice.index,
      score: choice.score
    }))
  });

  return scored.find((choice) => choice.score > 0) || null;
}

async function setComboboxValue(control, value, field = null, context = {}) {
  const country = String(context.country || "").trim();
  if (context.locationCountryFirst &&
      country &&
      normalize(country) !== normalize(value)) {
    await clearComboboxSearchValue(control);
    await getComboboxChoices(control, country, {
      filterWaitMs: LOCATION_AUTOCOMPLETE_WAIT_MS
    });
    traceAutoBid("location-autocomplete:country-prime", {
      field_id: field?.id || control.dataset?.autoBidFieldId || "",
      label: field?.question || field?.label || "",
      country
    });
  }

  const options = await getComboboxChoices(control, value, {
    filterWaitMs: context.locationCountryFirst &&
      normalize(country) === normalize(value)
      ? LOCATION_AUTOCOMPLETE_WAIT_MS
      : 220
  });
  const choices = options.map((option, index) => ({
      control: option,
      text: option.textContent,
      value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent,
      index
    }));
  const match = findBestAutocompleteChoice(choices, value, field, context);

  if (!match) {
    await clearComboboxSearchValue(control);
    await closeCombobox(control);
    return false;
  }

  const selected = await selectComboboxChoice(control, match);
  if (!selected) await closeCombobox(control);
  return selected;
}

async function getComboboxChoices(control, filterValue, options = {}) {
  const choices = await openCombobox(control, filterValue, options);
  return choices.filter((option) => !isGeneratedChoicePlaceholder(
    option.textContent,
    option.getAttribute("data-value") || option.getAttribute("value") || ""
  ));
}

async function openCombobox(control, filterValue, config = {}) {
  await scrollElementIntoView(control, "center");
  const input = getComboboxInput(control);
  const canFilter = Boolean(
    filterValue &&
    input?.matches?.("input, textarea") &&
    !input.readOnly &&
    !input.disabled
  );
  if (document.activeElement === input && !isComboboxOpen(control)) {
    input.blur();
    await sleep(80);
  }

  let options = getVisibleChoiceElements(control);
  if (isComboboxOpen(control) && options.length > 0) {
    if (canFilter) {
      await applyComboboxFilter(input, filterValue, config.filterWaitMs);
      options = await waitForComboboxFilteredOptions(control, filterValue, DROPDOWN_OPEN_TIMEOUT_MS);
    }
    return options;
  }

  const trigger = getComboboxTrigger(control);
  const nativeClicked = await nativeClickElement(trigger);
  traceAutoBid("dropdown:open-click", describeComboboxState(control, { nativeClicked }));
  if (canFilter) await applyComboboxFilter(input, filterValue, config.filterWaitMs);
  options = canFilter
    ? await waitForComboboxFilteredOptions(control, filterValue, 500)
    : await waitForComboboxOptions(control, 500);
  traceAutoBid("dropdown:open-after-click", describeComboboxState(control, { optionCount: options.length }));
  if (options.length > 0) return options;

  const pageTriggerOpened = runPageCommand("combobox-open", trigger);
  const pageInputOpened = input !== trigger ? runPageCommand("combobox-open", input) : false;
  if (canFilter) await applyComboboxFilter(input, filterValue, config.filterWaitMs);
  options = canFilter
    ? await waitForComboboxFilteredOptions(control, filterValue, DROPDOWN_OPEN_TIMEOUT_MS)
    : await waitForComboboxOptions(control, DROPDOWN_OPEN_TIMEOUT_MS);
  traceAutoBid("dropdown:open-after-page-bridge", describeComboboxState(control, {
    optionCount: options.length,
    pageTriggerOpened,
    pageInputOpened
  }));
  if (options.length > 0) return options;

  for (const key of ["ArrowDown", " "]) {
    pressComboboxKey(input || control, key);
    options = canFilter
      ? await waitForComboboxFilteredOptions(control, filterValue, 900)
      : await waitForComboboxOptions(control, 900);
    traceAutoBid("dropdown:open-after-key", describeComboboxState(control, { key, optionCount: options.length }));
    if (options.length > 0) break;
  }

  if (options.length === 0) {
    reportAndClearNativeClickError(control);
    return [];
  }

  if (!canFilter && "value" in input && filterValue && !input.readOnly) {
    await applyComboboxFilter(input, filterValue, config.filterWaitMs);
    options = await waitForComboboxOptions(control, DROPDOWN_OPEN_TIMEOUT_MS);
  }

  return options;
}

async function applyComboboxFilter(input, filterValue, waitMs = 220) {
  if (!input?.matches?.("input, textarea") || input.readOnly || input.disabled) return false;
  input.focus?.();
  setNativeValue(input, filterValue);
  input.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: filterValue }));
  await sleep(Math.max(0, Number(waitMs) || 220));
  return true;
}

async function clearComboboxSearchValue(control) {
  const input = getComboboxInput(control);
  if (!input?.matches?.("input, textarea") || input.readOnly || input.disabled || !input.value) return;
  input.focus?.();
  setNativeValue(input, "");
  dispatchInput(input);
  await sleep(80);
}

function findBestAutocompleteChoice(choices, value, field, context = {}) {
  if (!context.locationCountryFirst) return findBestChoice(choices, value);

  const answer = normalize(value);
  const country = normalize(context.country || "");
  let candidates = choices;

  if (context.profileKey === "country" && country && answer === country) {
    const exactCountryChoices = choices.filter((choice) => {
      const rawText = cleanLabel(choice.text || choice.value || "");
      const text = normalize(rawText);
      return text === country || text.startsWith(`${country} `) && !rawText.includes(",");
    });
    if (exactCountryChoices.length === 0) return null;
    candidates = exactCountryChoices;
  } else if (country) {
    const answerAndCountryChoices = choices.filter((choice) => {
      const text = normalize(choice.text || choice.value || "");
      return text.includes(answer) && text.includes(country);
    });
    if (answerAndCountryChoices.length > 0) {
      candidates = answerAndCountryChoices;
    } else {
      const exactAnswerChoices = choices.filter((choice) => normalize(choice.text || choice.value || "") === answer);
      if (exactAnswerChoices.length === 0) return null;
      candidates = exactAnswerChoices;
    }
  }

  return findBestChoice(candidates, value);
}

async function selectComboboxChoice(control, choice) {
  const option = findCurrentChoiceElement(control, choice);
  if (!option) return false;

  await clickChoice(option);
  if (await waitForChoiceApplied(control, choice, DROPDOWN_SELECT_TIMEOUT_MS)) return true;

  const options = await openCombobox(control);
  const retryOption = findCurrentChoiceElement(control, choice, options);
  if (retryOption) {
    await clickChoice(retryOption);
    if (await waitForChoiceApplied(control, choice, DROPDOWN_SELECT_TIMEOUT_MS)) return true;
  }

  const input = getComboboxInput(control);
  const expectedOptionId = retryOption?.id || option.id;
  const steps = Math.max(options.length + 1, Number.isInteger(choice.index) ? choice.index + 2 : 2);
  for (let index = 0; index < steps; index += 1) {
    if (expectedOptionId && input.getAttribute("aria-activedescendant") === expectedOptionId) break;
    pressComboboxKey(input, "ArrowDown");
    await sleep(80);
  }
  pressComboboxKey(control, "Enter");
  const selected = await waitForChoiceApplied(control, choice, DROPDOWN_SELECT_TIMEOUT_MS);
  if (!selected) reportAndClearNativeClickError(control);
  return selected;
}

function reportAndClearNativeClickError(control) {
  if (!lastNativeClickError) return;
  const error = consumeNativeClickError();
  traceAutoBid("dropdown:native-click-unavailable", {
    field_id: control?.dataset?.autoBidFieldId || "",
    message: error.message || String(error)
  });
}

async function closeCombobox(control) {
  if (!control) return;
  const input = getComboboxInput(control);
  [input, control, document.activeElement, document.body].filter(Boolean).forEach((target) => {
    target.dispatchEvent?.(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
    target.dispatchEvent?.(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Escape", code: "Escape", keyCode: 27, which: 27 }));
  });
  input.blur?.();
  control.blur?.();
  await sleep(160);
}

function findCurrentChoiceElement(control, choice, options = getVisibleChoiceElements(control)) {
  if (choice.control?.id) {
    const current = document.getElementById(choice.control.id);
    if (current && isVisible(current)) return current;
  }

  const expected = normalize(choice.text || choice.value || "");
  return options.find((option) => {
    const text = normalize(option.textContent || option.getAttribute("data-value") || option.getAttribute("value") || "");
    return text === expected || text.includes(expected) || expected.includes(text);
  }) || null;
}

function getVisibleChoiceElements(control) {
  const roots = getComboboxRoots(control);
  const portalSelector = Array.from(new Set([
    "[role='option']",
    "[role='menuitemradio']",
    "[data-radix-collection-item]",
    "[data-value]",
    "[cmdk-item]",
    ".select__option",
    ".fab-MenuOption",
    "[data-fabric-component='MenuOption']",
    ".select2-results__option",
    ".select2-result-label",
    ".chosen-results li.active-result",
    "[data-testid*='option' i]",
    "[id*='option' i]",
    ...(atsAdapters?.getOptionSelectors?.() || [])
  ])).join(",");
  const localSelector = `${portalSelector}, li, button`;
  const scopedRoots = roots.filter((root) => root !== document);
  const scopedOptions = collectChoiceElements(scopedRoots, localSelector, control);
  if (scopedOptions.length > 0) return scopedOptions;

  const scopedGenericOptions = getGenericVisibleChoiceElements(scopedRoots, control);
  if (scopedGenericOptions.length > 0) return scopedGenericOptions;

  const reactSelectId = getReactSelectInstanceId(control);
  if (reactSelectId) {
    const reactOptions = collectChoiceElements([
      document.getElementById(`${reactSelectId}-listbox`),
      document
    ].filter(Boolean), `[id^="${cssEscape(reactSelectId)}-option-"], [aria-labelledby^="${cssEscape(reactSelectId)}-option-"]`, control);
    if (reactOptions.length > 0) return reactOptions;
  }

  const portalOptions = collectChoiceElements([document], portalSelector, control);
  if (portalOptions.length > 0) return portalOptions;

  return getGenericVisibleChoiceElements(getOpenChoiceContainers(control), control);
}

function getGenericVisibleChoiceElements(roots, control) {
  return collectChoiceElements(roots.filter(Boolean), [
    "[role='option']",
    "[role='menuitem']",
    "[data-value]",
    "[class*='option' i]",
    "[class*='item' i]",
    "li",
    "button",
    "div"
  ].join(","), control)
    .filter((option) => isLikelyChoiceRow(option, control));
}

function getOpenChoiceContainers(control) {
  const selector = [
    "[role='listbox']",
    "[role='menu']",
    "[aria-expanded='true']",
    "[class*='listbox' i]",
    "[class*='dropdown' i]",
    "[class*='drop-down' i]",
    "[class*='menu' i]",
    "[class*='options' i]",
    "[class*='option-list' i]",
    "[class*='select-menu' i]",
    "[class*='select__menu' i]",
    ".fab-MenuVessel",
    ".fab-MenuList",
    ".select2-container--open",
    ".select2-drop-active",
    ".chosen-container-active",
    ".chosen-drop",
    "[data-radix-popper-content-wrapper]"
  ].join(",");

  return queryAll(selector)
    .filter((container) => container !== control && isVisible(container))
    .filter((container) => isLikelyChoiceContainerForControl(container, control));
}

function isLikelyChoiceContainerForControl(container, control) {
  if (container.contains(control)) return true;

  const trigger = getComboboxTrigger(control);
  const triggerRect = trigger.getBoundingClientRect();
  const rect = container.getBoundingClientRect();
  if (rect.width < 80 || rect.height < 20 || rect.height > window.innerHeight * 0.85) return false;
  if (rect.bottom < triggerRect.top - 24 || rect.top > triggerRect.bottom + 650) return false;

  const horizontalOverlap = Math.max(0, Math.min(rect.right, triggerRect.right) - Math.max(rect.left, triggerRect.left));
  if (horizontalOverlap < Math.min(60, triggerRect.width * 0.25)) return false;

  const text = normalize(container.textContent || "");
  return /\b(linkedin|indeed|glassdoor|job board|website|referral|facebook|other)\b/.test(text) ||
    /\b(select|option|dropdown|menu|listbox)\b/.test(normalize([container.className, container.id, container.getAttribute?.("role")].join(" ")));
}

function isLikelyChoiceRow(option, control) {
  if (!option || option === control || option.contains(control) || isLikelyDropdownTrigger(option)) return false;
  if (option.querySelector?.("input, textarea, select")) return false;

  const text = cleanLabel(option.textContent || option.getAttribute?.("aria-label") || option.getAttribute?.("data-value") || "");
  const normalized = normalize(text);
  if (!normalized || text.length > 140) return false;
  if (/^(search|select|select an option|choose|choose an option|type to search)$/.test(normalized)) return false;
  if (isPlaceholderChoice(text, option.getAttribute?.("data-value") || option.getAttribute?.("value") || "")) return false;

  const rect = option.getBoundingClientRect();
  if (rect.width < 40 || rect.height < 8 || rect.height > 96) return false;

  const trigger = getComboboxTrigger(control);
  const triggerRect = trigger.getBoundingClientRect();
  if (rect.bottom < triggerRect.top - 24 || rect.top > triggerRect.bottom + 650) return false;

  const horizontalOverlap = Math.max(0, Math.min(rect.right, triggerRect.right) - Math.max(rect.left, triggerRect.left));
  if (horizontalOverlap < Math.min(40, triggerRect.width * 0.18)) return false;

  const meaningfulChildren = Array.from(option.children || [])
    .filter((child) => isVisible(child) && cleanLabel(child.textContent || "") && !child.querySelector?.("input, textarea, select"))
    .filter((child) => {
      const childRect = child.getBoundingClientRect();
      return childRect.height >= 8 && childRect.height <= 96 && normalize(child.textContent || "") !== normalized;
    });
  return meaningfulChildren.length === 0;
}

function getComboboxRoots(control) {
  const roots = [];
  const controls = (control.getAttribute("aria-controls") || "").split(/\s+/).filter(Boolean);
  const owns = (control.getAttribute("aria-owns") || "").split(/\s+/).filter(Boolean);
  const active = control.getAttribute("aria-activedescendant");

  [...controls, ...owns, active].filter(Boolean).forEach((id) => {
    const element = document.getElementById(id);
    if (element) roots.push(element.closest("[role='listbox'], [role='menu'], .select__menu") || element);
  });

  const reactSelectId = getReactSelectInstanceId(control);
  if (reactSelectId) {
    const listbox = document.getElementById(`${reactSelectId}-listbox`);
    if (listbox) roots.push(listbox);
  }

  const shell = getComboboxShell(control);
  if (shell) roots.push(shell);
  const menu = shell?.parentElement?.querySelector(".select__menu, [role='listbox'], [role='menu']");
  if (menu) roots.push(menu);

  const expandedRoot = control.closest("[role='combobox'], [data-radix-select-trigger], [data-state]");
  if (expandedRoot?.parentElement) roots.push(expandedRoot.parentElement);
  roots.push(document);

  return Array.from(new Set(roots));
}

async function clickChoice(option) {
  await scrollElementIntoView(option, "nearest");
  const target = option.closest("[role='option'], [role='menuitemradio'], [data-value], .select__option, li, button") || option;
  if (!await nativeClickElement(target)) {
    runPageCommand("combobox-choose", target);
  }
}

async function nativeClickElement(element) {
  if (!element || !isVisible(element)) return false;
  if (!isTopFrame()) return false;
  const target = getHitTarget(element);
  const rect = target.getBoundingClientRect();
  const x = rect.left + rect.width / 2;
  const y = rect.top + rect.height / 2;
  if (x < 0 || y < 0 || x > window.innerWidth || y > window.innerHeight) return false;

  try {
    const result = await send("NATIVE_CLICK", { x, y });
    return result?.clicked === true;
  } catch (error) {
    lastNativeClickError = error;
    console.warn("Auto Bid native click failed", error);
    return false;
  }
}

function isTopFrame() {
  return window.top === window;
}

function getFrameScope() {
  return isTopFrame() ? "top" : "iframe";
}

function consumeNativeClickError() {
  const error = lastNativeClickError || new Error("Native dropdown click failed");
  lastNativeClickError = null;
  return error;
}

async function scrollElementIntoView(element, block = "center") {
  const scrollAncestors = getScrollAncestors(element);
  const previousBehavior = scrollAncestors.map((ancestor) => ({
    ancestor,
    value: ancestor.style.getPropertyValue("scroll-behavior"),
    priority: ancestor.style.getPropertyPriority("scroll-behavior")
  }));

  scrollAncestors.forEach((ancestor) => ancestor.style.setProperty("scroll-behavior", "auto", "important"));
  element.scrollIntoView({ behavior: "auto", block, inline: "nearest" });
  await sleep(120);

  const rect = element.getBoundingClientRect();
  if (rect.top < 0 || rect.bottom > window.innerHeight || rect.left < 0 || rect.right > window.innerWidth) {
    element.scrollIntoView({ behavior: "auto", block, inline: "nearest" });
    await sleep(100);
  }

  previousBehavior.forEach(({ ancestor, value, priority }) => {
    if (value) ancestor.style.setProperty("scroll-behavior", value, priority);
    else ancestor.style.removeProperty("scroll-behavior");
  });
}

function getScrollAncestors(element) {
  const ancestors = [document.documentElement, document.body].filter(Boolean);
  let current = element.parentElement;
  while (current && current !== document.body && current !== document.documentElement) {
    const style = window.getComputedStyle(current);
    if (/(auto|scroll|overlay)/.test(`${style.overflow} ${style.overflowY} ${style.overflowX}`)) {
      ancestors.push(current);
    }
    current = current.parentElement;
  }
  return Array.from(new Set(ancestors));
}

function isLikelyDropdownTrigger(element) {
  return element.getAttribute("role") === "combobox" ||
    element.getAttribute("aria-haspopup") === "listbox" ||
    element.hasAttribute("data-radix-select-trigger") ||
    element.getAttribute("data-slot") === "select-trigger" ||
    element.classList?.contains("select__control");
}

function collectChoiceElements(roots, selector, control) {
  const options = roots.flatMap((root) => [
    root.matches?.(selector) ? root : null,
    ...Array.from(root.querySelectorAll(selector))
  ].filter(Boolean));
  return Array.from(new Set(options))
    .filter((option) => option !== control && isVisible(option) && cleanLabel(option.textContent || ""))
    .filter((option) => !option.contains(control) && !isLikelyDropdownTrigger(option));
}

function describeComboboxState(control, extra = {}) {
  const input = getComboboxInput(control);
  const ownedListbox = getOwnedListbox(control);
  const reactSelectId = getReactSelectInstanceId(control);
  const visibleOptions = getVisibleChoiceElements(control);
  return {
    ...extra,
    controlId: control.id || "",
    expanded: input.getAttribute("aria-expanded") || control.getAttribute("aria-expanded") || "",
    activeDescendant: input.getAttribute("aria-activedescendant") || "",
    ariaControls: input.getAttribute("aria-controls") || control.getAttribute("aria-controls") || "",
    reactSelectId,
    listboxId: ownedListbox?.id || "",
    listboxVisible: ownedListbox ? isVisible(ownedListbox) : false,
    visibleOptionCount: visibleOptions.length,
    visibleOptions: visibleOptions.slice(0, 8).map((option, index) => ({
      index,
      id: option.id || "",
      role: option.getAttribute("role") || "",
      text: cleanLabel(option.textContent || "")
    }))
  };
}

async function waitForComboboxOptions(control, timeoutMs) {
  const started = Date.now();
  let options = [];
  while (Date.now() - started < timeoutMs) {
    options = getVisibleChoiceElements(control);
    if (isComboboxOpen(control) && options.length > 0) return options;
    await sleep(80);
  }
  return options;
}

async function waitForComboboxFilteredOptions(control, filterValue, timeoutMs) {
  const started = Date.now();
  let options = [];
  while (Date.now() - started < timeoutMs) {
    options = getVisibleChoiceElements(control)
      .filter((option) => !isPlaceholderChoice(
        option.textContent,
        option.getAttribute("data-value") || option.getAttribute("value") || ""
      ));
    const choices = options.map((option) => ({
      text: option.textContent,
      value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent
    }));
    if (isComboboxOpen(control) && findBestChoice(choices, filterValue)) return options;
    await sleep(80);
  }
  return options;
}

async function waitForChoiceApplied(control, choice, timeoutMs) {
  const started = Date.now();
  const expected = normalize(choice.text || choice.value || "");

  while (Date.now() - started < timeoutMs) {
    const selectedText = normalize(getComboboxValueText(control));
    if (choiceMatchesAppliedValue(selectedText, expected)) {
      return true;
    }
    await sleep(100);
  }

  return choiceMatchesAppliedValue(normalize(getComboboxValueText(control)), expected);
}

function choiceMatchesAppliedValue(selectedText, expected) {
  return Boolean(selectedText && (!expected || selectedText === expected || selectedText.includes(expected) || expected.includes(selectedText)));
}

async function waitForDropdownSettled(control) {
  const started = Date.now();
  let emptySince = 0;

  while (Date.now() - started < DROPDOWN_SELECT_TIMEOUT_MS) {
    if (!isComboboxOpen(control)) {
      if (!emptySince) emptySince = Date.now();
      if (Date.now() - emptySince >= 250) return true;
    } else {
      emptySince = 0;
    }
    await sleep(100);
  }

  pressComboboxKey(control, "Escape");
  await sleep(200);
  return !isComboboxOpen(control);
}

function isComboboxOpen(control) {
  const input = getComboboxInput(control);
  if (input.getAttribute("aria-expanded") === "true") return true;

  const listbox = getOwnedListbox(control);
  return Boolean(listbox && isVisible(listbox));
}

function getOwnedListbox(control) {
  const input = getComboboxInput(control);
  const ids = [
    ...(input.getAttribute("aria-controls") || "").split(/\s+/),
    ...(input.getAttribute("aria-owns") || "").split(/\s+/),
    `${getReactSelectInstanceId(control)}-listbox`
  ].filter(Boolean);

  for (const id of ids) {
    const element = document.getElementById(id);
    const listbox = element?.closest?.("[role='listbox'], [role='menu'], .select__menu") || element;
    if (listbox) return listbox;
  }
  return null;
}

function pressComboboxKey(control, key) {
  const normalizedKey = key === " " ? " " : key;
  const code = key === " " ? "Space" : key;
  const keyCode = key === "ArrowDown" ? 40 : key === " " ? 32 : key === "Escape" ? 27 : 13;
  const target = getComboboxInput(control);
  target.focus();
  if (runPageCommand("key", target, { key: normalizedKey })) return;
  target.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: normalizedKey, code, keyCode, which: keyCode }));
  target.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: normalizedKey, code, keyCode, which: keyCode }));
}

function getComboboxShell(control) {
  return control.closest(".select-shell") ||
    control.closest(".select__control") ||
    control.closest(".fab-SelectToggle") ||
    control.closest("[data-fabric-component='SelectToggle']") ||
    control.closest(".select2-container") ||
    control.closest(".chosen-container") ||
    control.closest("[data-radix-select-trigger], [data-slot='select-trigger']") ||
    (control.matches("[role='combobox']") ? control : control.closest("[role='combobox']"));
}

function getComboboxTrigger(control) {
  return control.closest(".select__control") ||
    control.closest(".fab-SelectToggle, [data-fabric-component='SelectToggle']") ||
    control.closest(".select2-selection, .select2-choice, .chosen-single") ||
    control.closest("[data-radix-select-trigger], [data-slot='select-trigger'], [role='combobox']") ||
    control;
}

function getComboboxToggle(control) {
  const shell = getComboboxShell(control);
  return shell?.querySelector("button[aria-label*='toggle' i], button[aria-haspopup], .select__indicators button") || null;
}

function getComboboxSelectedText(control) {
  const shell = getComboboxShell(control);
  const selected = shell?.querySelector(".select__single-value, [class*='single-value'], .fab-SelectToggle__content, .select2-selection__rendered, .select2-chosen, .chosen-single span, [aria-selected='true']");
  return cleanLabel(selected?.textContent || "");
}

function getComboboxValueText(control) {
  const input = getComboboxInput(control);
  return cleanLabel(input.value || control.value || getComboboxSelectedText(control) || control.textContent || "");
}

function getComboboxInput(control) {
  if (control.matches("input, textarea")) return control;
  const shell = getComboboxShell(control);
  return control.querySelector("input, textarea") ||
    shell?.querySelector("input.select2-search__field, .select2-search input, .chosen-search input, input[role='combobox'], input[aria-autocomplete]") ||
    control;
}

function setNativeValue(control, value) {
  if (runPageCommand("input", control, { value })) return;
  const setter = Object.getOwnPropertyDescriptor(control.constructor.prototype, "value")?.set;
  if (setter) setter.call(control, value);
  else control.value = value;
}

function setNativeChecked(control, checked) {
  if (runPageCommand("checked", control, { checked })) {
    dispatchInput(control);
    return;
  }
  const setter = Object.getOwnPropertyDescriptor(control.constructor.prototype, "checked")?.set;
  if (setter) setter.call(control, checked);
  else control.checked = checked;
}

function dispatchRealisticMouseClick(element) {
  runPageCommand("click", element);
  const target = getHitTarget(element);
  const rect = target.getBoundingClientRect();
  const clientX = Math.max(rect.left + Math.min(rect.width / 2, rect.width - 2), rect.left + 1);
  const clientY = Math.max(rect.top + Math.min(rect.height / 2, rect.height - 2), rect.top + 1);
  const eventInit = { bubbles: true, cancelable: true, composed: true, view: window, clientX, clientY, screenX: clientX, screenY: clientY, button: 0, buttons: 1 };

  target.dispatchEvent(new PointerEvent("pointerover", { ...eventInit, pointerType: "mouse" }));
  target.dispatchEvent(new PointerEvent("pointerenter", { ...eventInit, bubbles: false, pointerType: "mouse" }));
  target.dispatchEvent(new MouseEvent("mouseover", eventInit));
  target.dispatchEvent(new MouseEvent("mouseenter", { ...eventInit, bubbles: false }));
  target.dispatchEvent(new PointerEvent("pointermove", { ...eventInit, pointerType: "mouse" }));
  target.dispatchEvent(new MouseEvent("mousemove", eventInit));
  target.dispatchEvent(new PointerEvent("pointerdown", { ...eventInit, pointerType: "mouse" }));
  target.dispatchEvent(new MouseEvent("mousedown", eventInit));
  target.dispatchEvent(new PointerEvent("pointerup", { ...eventInit, buttons: 0, pointerType: "mouse" }));
  target.dispatchEvent(new MouseEvent("mouseup", { ...eventInit, buttons: 0 }));
  target.dispatchEvent(new MouseEvent("click", { ...eventInit, buttons: 0, detail: 1 }));
}

function getHitTarget(element) {
  const rect = element.getBoundingClientRect();
  const x = rect.left + rect.width / 2;
  const y = rect.top + rect.height / 2;
  const hit = document.elementFromPoint(x, y);
  if (hit && (element === hit || element.contains(hit))) return hit;
  return element;
}

function prioritizeHoveredField(fields) {
  const hoveredId = getHoveredAutoBidFieldId();
  if (!hoveredId) return fields;
  return [...fields].sort((left, right) => {
    if (left.id === hoveredId) return -1;
    if (right.id === hoveredId) return 1;
    return 0;
  });
}

function getHoveredAutoBidFieldId() {
  if (!lastMousePoint) return "";
  const element = document.elementFromPoint(lastMousePoint.x, lastMousePoint.y);
  const control = findAutoBidControl(element);
  return control?.dataset?.autoBidFieldId || "";
}

function findAutoBidControl(element) {
  return element?.closest?.("[data-auto-bid-field-id]") ||
    element?.closest?.(".field-wrapper, .select-shell, .select__control, .select")?.querySelector?.("[data-auto-bid-field-id]") ||
    null;
}

function shortText(value) {
  const text = cleanLabel(value);
  return text.length > 72 ? `${text.slice(0, 69)}...` : text;
}

function summarizeChoice(choice) {
  if (!choice) return null;
  return {
    text: cleanLabel(choice.text || ""),
    value: cleanLabel(choice.value || ""),
    index: choice.index,
    score: choice.score
  };
}

function resetTrace() {
  autoBidTrace = [];
  generatedAnswerFillAttempts.clear();
  autoBidRunId = `abr_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  autoBidRunStartedAt = new Date().toISOString();
  activeAutoBidProfileId = "";
  activeAutoBidProfileEmail = "";
  document.documentElement.removeAttribute(DEBUG_ATTR);
  publishTraceSnapshot("running");
}

function traceAutoBid(event, data = {}) {
  const entry = {
    at: new Date().toISOString(),
    event,
    data
  };
  autoBidTrace.push(entry);
  if (autoBidTrace.length > 240) autoBidTrace.shift();
  console.info("[AutoBid]", event, data);
  scheduleTracePublish();
}

function flushTrace() {
  try {
    const payload = {
      url: location.href,
      title: document.title || "",
      generated_at: new Date().toISOString(),
      entries: autoBidTrace
    };
    document.documentElement.setAttribute(DEBUG_ATTR, JSON.stringify(payload).slice(0, 120000));
    const lastEvent = autoBidTrace.at(-1)?.event || "";
    publishTraceSnapshot(lastEvent === "run:error" ? "failed" : lastEvent === "run:complete" ? "completed" : "running");
  } catch (error) {
    console.warn("Auto Bid could not export debug trace", error);
  }
}

function scheduleTracePublish() {
  if (tracePublishTimer) return;
  tracePublishTimer = window.setTimeout(() => {
    tracePublishTimer = null;
    publishTraceSnapshot("running");
  }, 500);
}

function publishTraceSnapshot(status) {
  if (!autoBidRunId || !isActiveContentInstance()) return;
  const completeEntry = [...autoBidTrace].reverse().find((entry) => entry.event === "run:complete" || entry.event === "run:error");
  const completeData = completeEntry?.data || {};
  const payload = {
    run_id: autoBidRunId,
    url: location.href,
    title: document.title || "",
    ats: atsAdapters?.describe?.() || { id: "common", name: "Common form" },
    status,
    started_at: autoBidRunStartedAt,
    updated_at: new Date().toISOString(),
    completed_at: ["completed", "failed", "cancelled"].includes(status) ? new Date().toISOString() : null,
    summary: {
      filled: Number(completeData.filled || 0),
      missed: Number(completeData.missed || 0),
      submitted: Boolean(completeData.submit?.clicked),
      message: completeEntry?.event === "run:error" ? String(completeData.message || "Autofill failed") : ""
    },
    entries: autoBidTrace
  };
  try {
    chrome.runtime.sendMessage({ type: "AUTOBID_LOG_UPSERT", payload }, () => void chrome.runtime.lastError);
  } catch (_error) {
    // Extension reloads can invalidate a content-script context mid-run.
  }
}

function runPageCommand(type, element, extra = {}) {
  if (document.documentElement.getAttribute("data-auto-bid-page-helper") !== "ready") return false;
  const token = `ab_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  element.setAttribute("data-auto-bid-bridge-token", token);
  document.documentElement.removeAttribute("data-auto-bid-command-result");
  document.documentElement.setAttribute("data-auto-bid-command", JSON.stringify({ type, token, ...extra }));
  document.dispatchEvent(new CustomEvent("autoBid:pageCommand"));
  const acknowledged = document.documentElement.getAttribute("data-auto-bid-command-result") === token;
  document.documentElement.removeAttribute("data-auto-bid-command");
  window.setTimeout(() => {
    if (element.getAttribute("data-auto-bid-bridge-token") === token) {
      element.removeAttribute("data-auto-bid-bridge-token");
    }
  }, 1000);
  return acknowledged;
}

function getReactSelectInstanceId(control) {
  const text = [
    control.getAttribute("aria-describedby"),
    control.getAttribute("aria-controls"),
    control.getAttribute("aria-activedescendant"),
    control.id ? `react-select-${control.id}` : "",
    getComboboxShell(control)?.querySelector("[id^='react-select-']")?.id
  ].filter(Boolean).join(" ");
  const match = text.match(/\b(react-select-[a-z0-9_-]+?)(?:-(?:placeholder|live-region|listbox|option-\d+))?\b/i);
  return match?.[1] || "";
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function scorePositiveChoice(label, choiceText, index) {
  const rawChoice = String(choiceText || "").toLowerCase();
  const normalizedLabel = normalize(label);
  const normalizedChoice = normalize(choiceText);
  const combined = `${normalizedLabel} ${normalizedChoice}`.trim();
  const numbers = normalizedChoice.match(/\d+/g)?.map(Number) || [];
  const bestNumber = numbers.length > 0 ? Math.max(...numbers) : 0;
  const experienceContext = /(experience|years|level|proficiency|skill|expertise|knowledge|familiar|rating|seniority)/.test(normalizedLabel);
  const scaleContext = /(rating|scale|level|proficiency|seniority)/.test(normalizedLabel);
  const sponsorshipContext = /(sponsor|sponsorship|visa|work permit|authorization|authorisation|work authorization support|work authorisation support)/.test(normalizedLabel);
  const sponsorshipNeedContext = /(require|need|seek|request|support).*(sponsor|sponsorship|visa|work permit|authorization|authorisation)|(sponsor|sponsorship|visa|work permit|authorization|authorisation).*(require|need|support|request)/.test(normalizedLabel);
  const workAuthorizationContext = /(authorized|eligible|legally).*(work|employ)|(work|employ).*(authorized|eligible|legally)|without.*(sponsor|sponsorship|visa)/.test(normalizedLabel);
  const salaryComfortContext = /(comfortable|accept|agree|ok|okay).*(salary|compensation|pay|range)|(salary|compensation|pay|range).*(comfortable|accept|agree|ok|okay)/.test(normalizedLabel);
  const positiveYesContext = /(comfortable|authorized|eligible|willing|available|open to|agree|accept|consent|confirm|legally|relocat|remote|can you|able to|do you have|have you|do you use|have you used|use.*ai|ai.*assist|artificial intelligence|development workflow|meet.*requirement|salary|compensation|pay range)/.test(normalizedLabel);
  const yesChoice = /\b(yes|y|true|agree|accepted|authorized|eligible|willing|available|comfortable|ok|okay)\b/.test(normalizedChoice);
  const noChoice = /\b(no|n|false)\b|not required|do not|dont|does not|will not|won t|cannot|can not/.test(normalizedChoice);
  let score = 0;

  if (/select|choose|please|placeholder|n\/a|not applicable/.test(normalizedChoice)) return -200;

  if (isExperienceYearsLabel(normalizedLabel)) {
    const targetYears = getExperienceDefaultTargetYears(getExperienceYearsKind(normalizedLabel));
    const yearScore = scoreExperienceYearsOption(choiceText, targetYears);
    if (yearScore > 0) return yearScore + index;
  }

  if (sponsorshipContext && sponsorshipNeedContext && !workAuthorizationContext) {
    if (noChoice || /not require|not need|no sponsorship|without sponsorship/.test(normalizedChoice)) return 900 + index;
    if (yesChoice || /require|need|sponsorship/.test(normalizedChoice)) return -900;
  }

  if (salaryComfortContext) {
    if (yesChoice) return 850 + index;
    if (noChoice) return -850;
  }

  if ((positiveYesContext || experienceContext) && yesChoice) score += 650;
  if ((positiveYesContext || experienceContext) && noChoice) score -= 650;

  if (experienceContext || scaleContext || numbers.length > 0) {
    score += index * 3;
    score += bestNumber * 25;
    if (/[0-9]\s*\+/.test(rawChoice) || /plus|more than|over|above|greater than|or more|maximum|max|highest|most/.test(normalizedChoice)) score += 260;
    if (/less than|under|below|fewer than|up to|at most|maximum of 0/.test(normalizedChoice)) score -= 160;
  }

  if (/distinguished|principal|staff|architect|expert|advanced|excellent|extensive|very strong|strong|highly|high|highest|senior|lead|master|proficient|fluent|native/.test(normalizedChoice)) score += 260;
  if (/intermediate|moderate|medium|working|practical|hands on|some|familiar/.test(normalizedChoice)) score += 110;
  if (/beginner|basic|entry|junior|limited|little|none|no experience|not at all|false/.test(normalizedChoice)) score -= 260;

  if (/\b(yes|true|agree)\b/.test(combined) && score === 0) score += 80;
  if (/\b(no|false)\b/.test(normalizedChoice) && score === 0) score -= 80;

  return score;
}

function isGeneratedChoicePlaceholder(text, value) {
  const normalizedText = normalize(text);
  if (/^(?:n a|not applicable)$/.test(normalizedText)) return false;
  return isPlaceholderChoice(text, value);
}

function isPlaceholderChoice(text, value) {
  const normalizedText = normalize(text);
  const normalizedValue = normalize(value);
  if (!normalizedText && !normalizedValue) return true;
  return /^(select|choose|please select|placeholder|n a|not applicable|search|type to search|start typing|start searching)$/.test(normalizedText) ||
    /^(?:start typing|start searching|type to search)(?: no results?)?$/.test(normalizedText) ||
    /^(no results?|no options?|nothing found|no matches?|no locations? found|no suggestions?)$/.test(normalizedText) ||
    /^(select|choose)(?: an| a)? option\b/.test(normalizedText) ||
    (/this field is required/.test(normalizedText) && /\b(select|choose).*\boption\b/.test(normalizedText));
}

function buildChoiceAliases(text, value) {
  const aliases = [text, value].filter(Boolean);
  if (["yes", "y", "true", "1"].includes(text) || ["yes", "y", "true", "1"].includes(value)) aliases.push("yes", "true", "agree", "authorized");
  if (["no", "n", "false", "0"].includes(text) || ["no", "n", "false", "0"].includes(value)) aliases.push("no", "false", "not");
  aliases.push(...buildLocationAliases(`${text} ${value}`, { includeCountryRegions: false }));
  return Array.from(new Set(aliases));
}

function buildAnswerAliases(answer) {
  return Array.from(new Set([answer, ...buildLocationAliases(answer, { includeCountryRegions: true })].filter(Boolean)));
}

function buildLocationAliases(value, options = {}) {
  const text = normalize(value);
  const aliases = [];
  const includeCountryRegions = options.includeCountryRegions !== false;

  if (text.includes("united states") || /\busa\b|\bus\b|\bamerica\b/.test(text)) {
    aliases.push("usa", "us", "united states", "america", "united states eastern time zone", "usa eastern time zone");
  }

  if (text === "uk" || text.includes("united kingdom") || text.includes("great britain") || text.includes("england") || text.includes("scotland") || text.includes("wales") || text.includes("northern ireland")) {
    aliases.push("uk", "united kingdom", "great britain");
  }

  if (text.includes("european union") || /\beu\b/.test(text) || text.includes("europe")) {
    aliases.push("eu", "europe", "european union");
  }

  if (text.includes("masovian") || text.includes("masovia") || text.includes("mazowieckie")) {
    aliases.push("masovian voivodeship", "masovian", "masovia", "mazowieckie");
  }

  if (includeCountryRegions && (EU_COUNTRY_NAMES.has(text) || Array.from(EU_COUNTRY_NAMES).some((country) => text.includes(country)))) {
    aliases.push("eu", "europe", "european union");
  }

  return aliases;
}

function scoreChoice(choice, answer) {
  if (!choice) return 0;
  if (choice === answer) return 100;
  if (choice.length > 1 && answer.length > 1 && choice.includes(answer)) return 80;
  if (choice.length > 1 && answer.length > 1 && answer.includes(choice)) return 70;

  const choiceWords = new Set(choice.split(" ").filter(Boolean));
  const answerWords = new Set(answer.split(" ").filter(Boolean));
  const intersection = Array.from(choiceWords).filter((word) => answerWords.has(word));
  if (intersection.length === 0) return 0;
  return Math.round((intersection.length / Math.max(choiceWords.size, answerWords.size)) * 60);
}

function showStatus(message, kind = "working", options = {}) {
  if (!isActiveContentInstance()) return;
  let status = document.getElementById(STATUS_ID);
  if (!status) {
    status = document.createElement("div");
    status.id = STATUS_ID;
    document.documentElement.append(status);
  }

  const detail = typeof options === "string" ? options : String(options.detail || "");
  const active = ["working", "waiting", "autofilling", "progress"].includes(kind);
  const palette = kind === "error"
    ? { background: "#fff7ed", border: "#fdba74", text: "#9a3412", icon: "#c2410c" }
    : kind === "warning"
      ? { background: "#fffbeb", border: "#fcd34d", text: "#92400e", icon: "#d97706" }
      : kind === "waiting"
        ? { background: "#eff6ff", border: "#93c5fd", text: "#1e3a8a", icon: "#2563eb" }
        : { background: "#ecfdf5", border: "#86efac", text: "#065f46", icon: "#059669" };

  window.clearTimeout(status._autoBidTimer);
  status.replaceChildren();
  status.dataset.autoBidStatusKind = kind;
  status.setAttribute("role", kind === "error" ? "alert" : "status");
  status.setAttribute("aria-live", kind === "error" ? "assertive" : "polite");
  status.style.cssText = [
    "position:fixed",
    "right:16px",
    "top:16px",
    "z-index:2147483647",
    "display:flex",
    "align-items:flex-start",
    "gap:10px",
    "width:min(320px, calc(100vw - 32px))",
    "box-sizing:border-box",
    "padding:11px 12px",
    "border-radius:12px",
    "font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif",
    `background:${palette.background}`,
    `color:${palette.text}`,
    `border:1px solid ${palette.border}`,
    "box-shadow:0 14px 40px rgba(19,35,29,.20)",
    "pointer-events:auto"
  ].join(";");

  const icon = document.createElement("span");
  icon.setAttribute("aria-hidden", "true");
  if (active) {
    icon.style.cssText = [
      "flex:0 0 auto",
      "width:16px",
      "height:16px",
      "margin-top:2px",
      "box-sizing:border-box",
      "border-radius:999px",
      `border:2px solid ${palette.border}`,
      `border-top-color:${palette.icon}`
    ].join(";");
    icon.animate?.(
      [{ transform: "rotate(0deg)" }, { transform: "rotate(360deg)" }],
      { duration: 750, iterations: Infinity, easing: "linear" }
    );
  } else {
    icon.textContent = kind === "error" ? "!" : kind === "warning" ? "!" : "✓";
    icon.style.cssText = [
      "display:inline-flex",
      "align-items:center",
      "justify-content:center",
      "flex:0 0 auto",
      "width:18px",
      "height:18px",
      "margin-top:1px",
      "border-radius:999px",
      `background:${palette.icon}`,
      "color:#fff",
      "font:700 12px/1 system-ui,sans-serif"
    ].join(";");
  }

  const content = document.createElement("span");
  content.style.cssText = "display:block;min-width:0;flex:1";

  const brand = document.createElement("span");
  brand.textContent = "AUTO BID";
  brand.style.cssText = "display:block;margin-bottom:1px;font:700 9px/1.3 Inter,system-ui,sans-serif;letter-spacing:.11em;opacity:.68";

  const title = document.createElement("span");
  title.textContent = String(message || "Autofill status");
  title.style.cssText = "display:block;font:700 13px/1.35 Inter,system-ui,sans-serif";

  content.append(brand, title);
  if (detail) {
    const description = document.createElement("span");
    description.textContent = detail;
    description.style.cssText = "display:block;margin-top:2px;font:500 11.5px/1.4 Inter,system-ui,sans-serif;opacity:.82";
    content.append(description);
  }

  const close = document.createElement("button");
  close.type = "button";
  close.textContent = "×";
  close.setAttribute("aria-label", "Dismiss Auto Bid status");
  close.style.cssText = [
    "flex:0 0 auto",
    "width:20px",
    "height:20px",
    "margin:-4px -5px 0 0",
    "padding:0",
    "border:0",
    "border-radius:6px",
    "background:transparent",
    "color:inherit",
    "font:600 18px/18px system-ui,sans-serif",
    "cursor:pointer",
    "opacity:.58"
  ].join(";");
  close.addEventListener("click", () => status.remove(), { once: true });

  status.append(icon, content, close);
}

function hashSmall(value) {
  let hash = 0;
  for (let i = 0; i < value.length; i++) {
    hash = ((hash << 5) - hash + value.charCodeAt(i)) | 0;
  }
  return Math.abs(hash).toString(36);
}

function cssEscape(value) {
  if (window.CSS?.escape) return CSS.escape(value);
  return String(value).replace(/["\\]/g, "\\$&");
}

function send(type, payload, extra = {}) {
  if (typeof chrome === "undefined" || !chrome.runtime?.id) {
    throw new Error("Auto Bid extension context changed. Reload this page once, then press the hotkey again.");
  }

  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type, payload, ...extra }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      if (!response?.ok) {
        reject(new Error(response?.error || "Request failed"));
        return;
      }
      resolve(response.data);
    });
  });
}
})();
