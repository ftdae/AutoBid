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
const CHOICE_FIELD_TYPES = ["select", "radio", "combobox", "button-group"];
const LANGUAGE_ALIASES = [
  ["english", ["english"]],
  ["ukrainian", ["ukrainian"]],
  ["polish", ["polish"]],
  ["russian", ["russian"]],
  ["spanish", ["spanish"]],
  ["portuguese", ["portuguese"]],
  ["german", ["german"]],
  ["french", ["french"]],
  ["italian", ["italian"]],
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
  }
});

runAutoBid();

async function runAutoBid() {
  if (!isActiveContentInstance() || autoBidRunning) return;
  autoBidRunning = true;
  lastNativeClickError = null;
  resetTrace();
  traceAutoBid("run:start", { url: location.href, title: document.title, frame: getFrameScope() });
  showStatus("Auto Bid is filling this page...", "success");

  const fields = collectFields();
  traceAutoBid("fields:collected", {
    count: fields.length,
    fields: fields.map((field) => ({
      id: field.id,
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
    return;
  }

  try {
    const hoverResult = await applyHoveredDropdownSelection(fields);
    const prefillResult = await applyPositiveDropdownFallbacks(fields, hoverResult.filledIds);
    const checkboxResult = await applyPositiveCheckboxFallbacks(fields, new Set([...hoverResult.filledIds, ...prefillResult.filledIds]));
    const defaultResult = await applyDeterministicDefaults(fields, new Set([...hoverResult.filledIds, ...prefillResult.filledIds, ...checkboxResult.filledIds]));
    const staticFallbackResult = await applyProfileStaticFallbacks(fields, new Set([...hoverResult.filledIds, ...prefillResult.filledIds, ...checkboxResult.filledIds, ...defaultResult.filledIds]));
    const languageChoiceResult = await applyLanguageChoiceAnswers(fields, new Set([...hoverResult.filledIds, ...prefillResult.filledIds, ...checkboxResult.filledIds, ...defaultResult.filledIds, ...staticFallbackResult.filledIds]));
    let data;
    try {
      data = await send("ASSIST", {
        page: collectPageContext(),
        fields
      });
    } catch (error) {
      traceAutoBid("assist:error", { message: error.message || String(error) });
      if (!staticFallbackResult.filled && !checkboxResult.filled && !defaultResult.filled && !languageChoiceResult.filled) throw error;
      data = {
        answers: [],
        warnings: [error.message || String(error)],
        cache: null
      };
    }
    const answers = data.answers || [];
    traceAutoBid("assist:received", {
      answers: answers.length,
      answer_fields: answers.map((answer) => ({
        field_id: answer.field_id,
        source: answer.source || "",
        cache_scope: answer.cache_scope || "",
        value: shortText(answer.value || "")
      })),
      warnings: data.warnings || [],
      cache: data.cache || null,
      static_merge: data.static_merge || null
    });
    (data.warnings || []).forEach((message) => {
      traceAutoBid("assist:warning", { message });
    });
    const dropdownFilledIds = new Set([...hoverResult.filledIds, ...prefillResult.filledIds, ...checkboxResult.filledIds, ...defaultResult.filledIds, ...staticFallbackResult.filledIds, ...languageChoiceResult.filledIds]);
    const result = await applyAnswers(answers, dropdownFilledIds);
    const residenceResult = await applyResidenceAnswers(fields, answers, new Set([...dropdownFilledIds, ...result.filledIds]));
    const locationChoiceResult = await applyBasedInLocationAnswers(fields, answers, new Set([...dropdownFilledIds, ...result.filledIds, ...residenceResult.filledIds]));
    const alreadyFilled = new Set([...dropdownFilledIds, ...result.filledIds, ...residenceResult.filledIds, ...locationChoiceResult.filledIds, ...checkboxResult.filledIds, ...defaultResult.filledIds, ...staticFallbackResult.filledIds, ...languageChoiceResult.filledIds]);
    const fallbackResult = await applyPositiveDropdownFallbacks(fields, alreadyFilled);
    if (data.draft_id) {
      send("DRAFT_STATUS", null, { draftId: data.draft_id, status: "filled" }).catch(() => {});
    }
    const filled = hoverResult.filled + prefillResult.filled + checkboxResult.filled + defaultResult.filled + languageChoiceResult.filled + result.filled + residenceResult.filled + locationChoiceResult.filled + staticFallbackResult.filled + fallbackResult.filled;
    const missed = hoverResult.missed + prefillResult.missed + checkboxResult.missed + defaultResult.missed + languageChoiceResult.missed + result.missed + residenceResult.missed + locationChoiceResult.missed + staticFallbackResult.missed + fallbackResult.missed;
    traceAutoBid("run:complete", { filled, missed });
    showStatus(`Filled ${filled} fields. ${missed} skipped.`, filled ? "success" : "error");
  } catch (error) {
    traceAutoBid("run:error", { message: error.message || String(error) });
    showStatus(error.message || String(error), "error");
  } finally {
    flushTrace();
    autoBidRunning = false;
  }
}

function collectPageContext() {
  return {
    url: location.href,
    domain: location.hostname.replace(/^www\./, ""),
    title: document.title || "",
    job_title: findJobTitle(),
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

  controls.forEach((control, index) => {
    const type = getControlType(control);
    if (type === "radio") {
      const key = control.name || control.id || `radio_${index}`;
      if (!radioGroups.has(key)) radioGroups.set(key, []);
      radioGroups.get(key).push(control);
      return;
    }

    const id = `ab_${index}_${hashSmall(getFieldText(control))}`;
    control.dataset.autoBidFieldId = id;
    fields.push({
      id,
      label: getFieldLabel(control),
      name: control.name || "",
      placeholder: control.getAttribute("placeholder") || "",
      autocomplete: control.getAttribute("autocomplete") || "",
      type,
      required: isRequired(control),
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
      name: first.name || "",
      placeholder: "",
      autocomplete: "",
      type: "radio",
      required: group.some(isRequired),
      options: group.map(getRadioOptionLabel).filter(Boolean),
      value: group.find((control) => control.checked)?.value || ""
    });
  });

  fields.push(...collectButtonChoiceFields(fields.length));

  return fields;
}

function getFormControls() {
  return queryAll([
    "input",
    "textarea",
    "select",
    "[role='checkbox']",
    "button[role='combobox']",
    "[role='combobox']",
    "[aria-haspopup='listbox']",
    "[aria-haspopup='menu'][aria-expanded]",
    "[data-radix-select-trigger]",
    "[data-slot='select-trigger']",
    ".select__control",
    "[class*='select'][aria-expanded]"
  ].join(","));
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
  const visibleChoiceControl = isVisibleChoiceControl(control, type);
  if ((control.getAttribute("aria-hidden") === "true" && !visibleChoiceControl) || isCompositeComboboxShell(control)) return false;
  if (!isVisible(control) && !visibleChoiceControl) return false;
  if (control.disabled || (control.readOnly && type !== "combobox")) return false;
  return !["hidden", "submit", "button", "reset", "image", "file", "password"].includes(type);
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
  if (control.dataset?.autoBidControlType === "button-group") return "button-group";
  if (control.getAttribute("role") === "checkbox") return "checkbox";
  if (
    control.getAttribute("role") === "combobox" ||
    ["listbox", "menu"].includes(control.getAttribute("aria-haspopup") || "") ||
    control.hasAttribute("data-radix-select-trigger") ||
    control.getAttribute("data-slot") === "select-trigger" ||
    control.classList?.contains("select__control")
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
    return Array.from(control.options).map((option) => option.textContent.trim()).filter(Boolean);
  }
  if (getControlType(control) === "combobox") return getVisibleChoiceElements(control).map((option) => cleanLabel(option.textContent || ""));
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
  return control.value || "";
}

function isRequired(control) {
  const requiredText = [
    getFieldLabel(control),
    getNearbyText(control),
    getDescribedByText(control)
  ].join(" ");
  return Boolean(
    control.required ||
    control.getAttribute("aria-required") === "true" ||
    control.closest("[aria-required='true'], [data-required='true']") ||
    /\*/.test(requiredText) ||
    /\brequired\b/i.test(requiredText)
  );
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
  const choiceLabels = ["checkbox", "radio"].includes(getControlType(control)) ? getChoiceInlineLabelCandidates(control) : [];
  const candidates = [
    fromFor?.textContent,
    closestLabel?.textContent,
    fromAria,
    control.getAttribute("aria-label"),
    ...choiceLabels,
    ...containerLabels,
    ...siblingLabels,
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
  return control.closest(FIELD_CONTAINER_SELECTOR) || control.parentElement;
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
    (normalize(value) === "on" ? "" : value)
  );
}

function cleanLabel(text) {
  return String(text || "").replace(/\s+/g, " ").replace(/\s+\*/g, " *").trim().slice(0, 300);
}

async function applyAnswers(answers, skipFilledIds = new Set()) {
  let filled = 0;
  let missed = 0;
  const filledIds = new Set();
  const missedIds = new Set();

  for (const answer of answers) {
    const controls = getControlsByFieldId(answer.field_id);
    if (controls.length === 0) {
      missed += 1;
      missedIds.add(answer.field_id);
      continue;
    }
    if (skipFilledIds.has(answer.field_id) && hasCurrentChoiceValue(controls)) {
      traceAutoBid("answer:skipped-positive-dropdown", {
        field_id: answer.field_id,
        answer: answer.value || "",
        source: answer.source || "",
        current: getCurrentChoiceSummary(controls)
      });
      filledIds.add(answer.field_id);
      continue;
    }
    if (await setControlsValue(controls, answer.value || "")) {
      traceAutoBid("answer:applied", {
        field_id: answer.field_id,
        answer: answer.value || "",
        source: answer.source || "",
        current: getCurrentChoiceSummary(controls)
      });
      filled += 1;
      filledIds.add(answer.field_id);
    } else {
      traceAutoBid("answer:missed", {
        field_id: answer.field_id,
        answer: answer.value || "",
        source: answer.source || ""
      });
      missed += 1;
      missedIds.add(answer.field_id);
    }
  }

  return { filled, missed, filledIds, missedIds };
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
    if (filledIds.has(field.id) || !isBasedInLocationField(field) || !isChoiceFieldType(field.type)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || hasCurrentChoiceValue(controls)) continue;

    const answer = locationAnswerMatchesQuestion(locationAnswer, field.label) ? "Yes" : "No";
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

async function applyLanguageChoiceAnswers(fields, filledIds) {
  const filledLocalIds = new Set();
  let filled = 0;
  let missed = 0;

  for (const field of fields) {
    if (filledIds.has(field.id) || !isLanguageChoiceField(field)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    const questionLanguages = getQuestionLanguageAliases(field.label);
    const answer = getLanguageChoiceAnswer(questionLanguages);
    const current = getCurrentChoiceSummary(controls);
    if (normalize(current) === normalize(answer)) {
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

async function applyProfileStaticFallbacks(fields, filledIds) {
  const filledLocalIds = new Set();
  let profileStatic;

  try {
    profileStatic = await send("GET_PROFILE_STATIC_FIELDS");
  } catch (error) {
    traceAutoBid("profile-static:error", { message: error.message || String(error) });
    return { filled: 0, missed: 0, filledIds: filledLocalIds };
  }

  const staticFields = profileStatic?.static_fields || {};
  const availableKeys = Object.keys(staticFields).filter((key) => String(staticFields[key] || "").trim());
  let filled = 0;
  let missed = 0;
  const matched = [];
  const missing = [];

  for (const field of fields) {
    if (filledIds.has(field.id)) continue;
    const key = matchProfileStaticFieldKey(field);
    if (!key) continue;

    matched.push({ field_id: field.id, key, label: field.label });
    const rawValue = getProfileStaticValue(staticFields, key);
    if (rawValue === undefined || rawValue === null || !String(rawValue).trim()) {
      missing.push({ field_id: field.id, key, label: field.label });
      continue;
    }

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0 || !shouldApplyProfileStaticValue(field, controls, rawValue)) continue;

    const value = formatProfileStaticValueForField(key, rawValue, field);
    if (await setControlsValue(controls, value)) {
      filled += 1;
      filledLocalIds.add(field.id);
      traceAutoBid("profile-static:applied", {
        field_id: field.id,
        key,
        label: field.label,
        value: shortText(value),
        current: getCurrentChoiceSummary(controls)
      });
    } else {
      missed += 1;
      traceAutoBid("profile-static:missed", {
        field_id: field.id,
        key,
        label: field.label,
        value: shortText(value)
      });
    }
  }

  traceAutoBid("profile-static:summary", {
    profile_id: profileStatic?.profile_id || "",
    available_keys: availableKeys,
    matched,
    missing,
    filled
  });

  return { filled, missed, filledIds: filledLocalIds };
}

function shouldApplyProfileStaticValue(field, controls, rawValue) {
  const current = String(getCurrentChoiceSummary(controls) || field.value || "").trim();
  if (!current) return true;
  if (normalize(current) === normalize(rawValue)) return false;

  const first = controls[0];
  if (["checkbox", "radio", "select", "combobox", "button-group"].includes(getControlType(first))) {
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
  if (isLanguageChoiceField(field)) return "";
  if (isPlainFullNameField(field)) return "full_name";
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name", "candidate name"]],
    ["notice_period", ["notice period", "current notice", "notice"]],
    ["expected_rate", ["hourly rate", "rate", "expected rate", "expected salary", "salary expectation", "salary expectations", "expected compensation", "desired salary", "desired compensation", "gross monthly", "monthly salary", "salary", "compensation"]],
    ["work_authorization", ["authorized", "authorization", "legally work", "eligible to work"]],
    ["sponsorship", ["sponsor", "sponsorship", "visa"]],
    ["availability", ["availability", "available", "start date"]],
    ["linkedin", ["linkedin"]],
    ["github", ["github"]],
    ["portfolio", ["portfolio"]],
    ["website", ["website", "personal site", "web site"]],
    ["languages", ["languages", "spoken languages", "language proficiency", "fluent languages", "languages spoken"]],
    ["country", ["country", "residence", "current residence", "where is your current residence", "where are you based"]],
    ["location", ["location", "address", "current city", "current location"]],
    ["city", ["city"]],
    ["phone", ["phone", "mobile", "telephone", "cell"]],
    ["email", ["email", "e mail", "mail"]]
  ];

  for (const [key, needles] of patterns) {
    if (needles.some((needle) => text.includes(needle))) return key;
  }

  return "";
}

function isPlainFullNameField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const candidates = [field.label, field.name, field.autocomplete]
    .map(normalize)
    .filter(Boolean);
  return candidates.some((candidate) => ["name", "your name", "applicant name", "candidate name", "full name"].includes(candidate));
}

function isTextLikeStaticField(field) {
  return !["checkbox", "radio", "select", "combobox", "button-group", "file", "hidden", "password", "submit", "button", "reset"].includes(field.type);
}

function getProfileStaticValue(staticFields, key) {
  const aliases = {
    expected_rate: ["expected_rate", "expected_salary", "salary_expectation", "salary_expectations", "monthly_salary", "monthly_salary_expectation", "desired_salary", "desired_compensation"],
    notice_period: ["notice_period", "current_notice_period", "availability_notice"],
    languages: ["languages", "language", "spoken_languages", "language_proficiency", "fluent_languages", "languages_spoken"]
  };

  if (key === "full_name") {
    const fullName = staticFields?.full_name;
    if (fullName !== undefined && fullName !== null && String(fullName).trim()) return fullName;
    const composed = [staticFields?.first_name, staticFields?.last_name].filter(Boolean).join(" ").trim();
    if (composed) return composed;
  }

  for (const candidate of [key, ...(aliases[key] || [])]) {
    const value = staticFields?.[candidate];
    if (value !== undefined && value !== null && String(value).trim()) return value;
  }

  return staticFields?.[key];
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
  const label = normalize([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  return /(currently.*based.*in|based.*in|currently.*located.*in|located.*in|currently.*living.*in|living.*in|fully.*living.*resident.*in|living.*resident.*in|currently.*residing.*in|residing.*in|resident.*in|residency.*work permit.*in|work permit.*in|current.*residence.*in)/.test(label) &&
    hasYesNoOptions(field.options);
}

function isLanguageChoiceField(field) {
  if (!field || !isChoiceFieldType(field.type) || !hasYesNoOptions(field.options)) return false;
  const label = normalize([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!/(speak|language|fluent|fluency|proficien|native speaker|bilingual|multilingual)/.test(label)) return false;
  return getQuestionLanguageAliases(label).length > 0;
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

function getLanguageChoiceAnswer(questionLanguages) {
  return (questionLanguages || []).some((language) => normalize(language) === "english") ? "Yes" : "No";
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

async function applyHoveredDropdownSelection(fields) {
  const filledIds = new Set();
  const fieldId = getHoveredAutoBidFieldId();
  if (!fieldId) return { filled: 0, missed: 0, filledIds };

  const field = fields.find((item) => item.id === fieldId);
  if (!field || !CHOICE_FIELD_TYPES.includes(field.type) || !shouldUsePositiveDropdownFallback(field)) {
    return { filled: 0, missed: 0, filledIds };
  }

  const controls = getControlsByFieldId(field.id);
  if (controls.length === 0) return { filled: 0, missed: 0, filledIds };

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
    if (filledIds.has(field.id)) continue;

    const controls = getControlsByFieldId(field.id);
    if (controls.length === 0) continue;

    const fallback = getDeterministicDefault(field, controls);
    if (!fallback) continue;

    const current = getCurrentChoiceSummary(controls);
    const alreadyApplied = getControlType(controls[0]) === "range"
      ? isRangeValueApplied(controls[0], fallback.value)
      : normalizeComparableValue(current) === normalizeComparableValue(fallback.value);
    if (alreadyApplied) {
      localFilledIds.add(field.id);
      continue;
    }

    const selected = await setDeterministicDefaultValue(controls, fallback.value);
    traceAutoBid("default:applied", {
      field_id: field.id,
      label: field.label,
      reason: fallback.reason,
      value: fallback.value,
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

  if (isExperienceValueField(field, controls)) {
    return { value: getDefaultExperienceYears(controls[0]), reason: "experience-years" };
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

function isAvailabilityDateField(field, controls) {
  const control = controls?.[0];
  const type = getControlType(control);
  if (!["date", "text", "search"].includes(type)) return false;
  const label = getFieldContextLabel(field, control);
  return /(date available|available date|available start date|earliest.*start|start date|when.*start|availability date)/.test(label);
}

function isExperienceValueField(field, controls) {
  const control = controls?.[0];
  const type = getControlType(control);
  if (!["range", "number", "text"].includes(type)) return false;
  const label = getFieldContextLabel(field, control);
  return /(how many.*years.*experience|years.*professional.*experience|years.*experience|experience.*years)/.test(label);
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

async function setDeterministicDefaultValue(controls, value) {
  const first = controls[0];
  const type = getControlType(first);

  if (type === "range") {
    return setRangeValueWithVisibleEditor(first, value);
  }

  if (["number", "date"].includes(type)) {
    await scrollElementIntoView(first, "center");
    setNativeValue(first, value);
    dispatchInput(first);
    await sleep(80);
    return normalizeComparableValue(getCurrentChoiceSummary(controls)) === normalizeComparableValue(value);
  }

  return setControlsValue(controls, value);
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
    element.select?.();
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

async function nativeTypeText(value) {
  try {
    const result = await send("NATIVE_TYPE", { text: String(value), commit: true });
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

function getDefaultExperienceYears(control) {
  const min = parseFiniteNumber(control?.min, 0);
  const max = parseFiniteNumber(control?.getAttribute?.("max") || control?.max, NaN);
  const step = parseFiniteNumber(control?.step, 1);
  let value = Number.isFinite(max) && max > min ? max : 10;

  if (value < min) value = min;
  if (Number.isFinite(max) && value > max) value = max;
  if (Number.isFinite(step) && step > 0) {
    value = min + Math.round((value - min) / step) * step;
  }

  return formatNumberValue(value);
}

function parseFiniteNumber(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function formatNumberValue(value) {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(2)));
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

  if (type === "radio") return controls.some((control) => control.checked);
  if (type === "checkbox") return isCheckboxChecked(first);
  if (type === "button-group") return Boolean(getSelectedChoiceButtonLabel(controls));
  if (first.tagName === "SELECT") {
    const option = first.selectedOptions?.[0];
    if (!option) return false;
    return Boolean(first.value && !isPlaceholderChoice(option.textContent, option.value));
  }
  if (type === "combobox") {
    const value = cleanLabel(first.value || getComboboxSelectedText(first) || first.textContent || "");
    return Boolean(value && !isPlaceholderChoice(value, value));
  }

  return Boolean(String(first.value || "").trim());
}

function getCurrentChoiceSummary(controls) {
  const first = controls?.[0];
  if (!first) return "";
  const type = getControlType(first);

  if (type === "radio") {
    const checked = controls.find((control) => control.checked);
    return checked ? getRadioOptionLabel(checked) || checked.value || "" : "";
  }

  if (type === "checkbox") {
    return isCheckboxChecked(first) ? "Yes" : "";
  }

  if (type === "button-group") {
    return getSelectedChoiceButtonLabel(controls);
  }

  if (first.tagName === "SELECT") {
    const option = first.selectedOptions?.[0];
    return cleanLabel(option ? `${option.textContent || ""} ${option.value || ""}` : "");
  }

  if (type === "combobox") {
    return cleanLabel(first.value || getComboboxSelectedText(first) || first.textContent || "");
  }

  return cleanLabel(first.value || "");
}

function shouldUsePositiveDropdownFallback(field) {
  if (isBasedInLocationField(field)) return false;
  if (isLanguageChoiceField(field)) return false;
  const label = normalize(field.label);
  return /(experience|years|level|proficiency|skill|expertise|knowledge|familiar|comfortable|rating|seniority|sponsor|sponsorship|visa|authorized|eligible|willing|available|open to|agree|accept|consent|confirm|legally|relocat|remote|can you|able to|do you have|have you|do you use|have you used|use.*ai|ai.*assist|artificial intelligence|development workflow|meet.*requirement|salary|compensation)/.test(label);
}

function shouldUsePositiveCheckboxFallback(field) {
  const label = normalize([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!label) return false;
  if (isSensitiveOrPersonalChoiceCheckbox(label)) return false;
  if (/(terms|privacy|policy|consent|agree|accept|acknowledge|confirm|certify|accurate|contact me|future job opportunit|future opportunit|talent community|job alert|keep my data|retain my data|store my data|process my data|data processing|email me|reach out|recruiting communication|recruitment communication|consider me for future)/.test(label)) return true;
  return false;
}

function isSensitiveOrPersonalChoiceCheckbox(label) {
  return /(pronoun|he him|she her|they them|xe xem|ze hir|ey em|hir hir|fae faer|hu hu|use name only|custom|gender|race|ethnicity|ethnic|disability|veteran|protected veteran|sexual orientation|date of birth|birth date|newsletter|marketing email|product update|event update)/.test(label);
}

async function setControlsValue(controls, value) {
  const first = controls[0];
  const type = getControlType(first);
  const textValue = String(value || "").trim();
  if (!textValue) return false;

  if (type === "radio") {
    return setRadioValue(controls, textValue);
  }

  if (type === "checkbox") {
    return setCheckboxValue(first, textValue);
  }

  if (type === "button-group") {
    return setButtonGroupValue(controls, textValue);
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
    return setComboboxValue(first, textValue);
  }

  const setter = Object.getOwnPropertyDescriptor(first.constructor.prototype, "value")?.set;
  if (setter) setter.call(first, textValue);
  else first.value = textValue;
  dispatchInput(first);
  return true;
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
  const target = getRadioClickTarget(radio);
  await scrollElementIntoView(target, "center");
  if (!await nativeClickElement(target)) {
    dispatchRealisticMouseClick(target);
  }
  await sleep(120);

  if (!radio.checked) {
    setNativeChecked(radio, true);
    dispatchInput(radio);
  }

  return radio.checked;
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

async function setCheckboxValue(control, value) {
  const shouldCheck = /^(yes|true|agree|accepted|checked|1)$/i.test(String(value || "").trim());
  if (isCheckboxChecked(control) === shouldCheck) return true;

  for (const target of getCheckboxClickTargets(control)) {
    await scrollElementIntoView(target, "center");
    if (!await nativeClickElement(target)) {
      dispatchRealisticMouseClick(target);
    }
    await sleep(140);
    if (isCheckboxChecked(control) === shouldCheck) return true;
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

  return isCheckboxChecked(control) === shouldCheck;
}

function getCheckboxClickTargets(control) {
  const targets = [];
  const input = getCheckboxInput(control);
  const root = input || control;
  const rootNode = root.getRootNode?.() || document;
  const forLabel = root.id ? queryOne(`label[for="${cssEscape(root.id)}"]`, rootNode) : null;

  [
    forLabel,
    root.closest?.("label"),
    root.closest?.("[class*='checkbox' i]"),
    root.closest?.("[class*='option' i]"),
    root.closest?.("li"),
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

async function setButtonGroupValue(controls, value) {
  const match = findBestChoice(
    controls.map((control) => ({
      control,
      text: getChoiceButtonLabel(control),
      value: getChoiceButtonValue(control)
    })),
    value
  );
  if (!match) return false;

  await scrollElementIntoView(match.control, "center");
  if (!await nativeClickElement(match.control)) {
    dispatchRealisticMouseClick(match.control);
  }
  await sleep(160);
  dispatchInput(match.control);
  return true;
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

async function setComboboxValue(control, value) {
  const options = await getComboboxChoices(control, value);
  const match = findBestChoice(
    options.map((option, index) => ({
      control: option,
      text: option.textContent,
      value: option.getAttribute("data-value") || option.getAttribute("value") || option.textContent,
      index
    })),
    value
  );

  if (!match) return false;

  return selectComboboxChoice(control, match);
}

async function getComboboxChoices(control, filterValue) {
  const options = await openCombobox(control, filterValue);
  return options.filter((option) => !isPlaceholderChoice(option.textContent, option.getAttribute("data-value") || option.getAttribute("value") || ""));
}

async function openCombobox(control, filterValue) {
  await scrollElementIntoView(control, "center");
  const input = getComboboxInput(control);
  if (document.activeElement === input && !isComboboxOpen(control)) {
    input.blur();
    await sleep(80);
  }

  let options = getVisibleChoiceElements(control);
  if (isComboboxOpen(control) && options.length > 0) return options;

  const trigger = getComboboxTrigger(control);
  const nativeClicked = await nativeClickElement(trigger);
  traceAutoBid("dropdown:open-click", describeComboboxState(control, { nativeClicked }));
  if (!nativeClicked) {
    runPageCommand("combobox-open", input || trigger);
  }
  options = await waitForComboboxOptions(control, DROPDOWN_OPEN_TIMEOUT_MS);
  traceAutoBid("dropdown:open-after-click", describeComboboxState(control, { optionCount: options.length }));
  if (options.length > 0) return options;

  pressComboboxKey(input || control, "ArrowDown");
  options = await waitForComboboxOptions(control, DROPDOWN_OPEN_TIMEOUT_MS);
  traceAutoBid("dropdown:open-after-key", describeComboboxState(control, { optionCount: options.length }));
  if (options.length === 0) {
    if (lastNativeClickError) throw consumeNativeClickError();
    return [];
  }

  if ("value" in input && filterValue && !input.readOnly) {
    setNativeValue(input, filterValue);
    input.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: filterValue }));
    options = await waitForComboboxOptions(control, DROPDOWN_OPEN_TIMEOUT_MS);
  }

  return options;
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
  if (!selected && lastNativeClickError) throw consumeNativeClickError();
  return selected;
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
  const portalSelector = [
    "[role='option']",
    "[role='menuitemradio']",
    "[data-radix-collection-item]",
    "[data-value]",
    "[cmdk-item]",
    ".select__option",
    "[data-testid*='option' i]",
    "[id*='option' i]"
  ].join(",");
  const localSelector = `${portalSelector}, li, button`;
  const scopedRoots = roots.filter((root) => root !== document);
  const scopedOptions = collectChoiceElements(scopedRoots, localSelector, control);
  if (scopedOptions.length > 0) return scopedOptions;

  const reactSelectId = getReactSelectInstanceId(control);
  if (reactSelectId) {
    const reactOptions = collectChoiceElements([
      document.getElementById(`${reactSelectId}-listbox`),
      document
    ].filter(Boolean), `[id^="${cssEscape(reactSelectId)}-option-"], [aria-labelledby^="${cssEscape(reactSelectId)}-option-"]`, control);
    if (reactOptions.length > 0) return reactOptions;
  }

  return collectChoiceElements([document], portalSelector, control);
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
    control.closest("[data-radix-select-trigger], [data-slot='select-trigger']") ||
    (control.matches("[role='combobox']") ? control : control.closest("[role='combobox']"));
}

function getComboboxTrigger(control) {
  return control.closest(".select__control") ||
    control.closest("[data-radix-select-trigger], [data-slot='select-trigger'], [role='combobox']") ||
    control;
}

function getComboboxToggle(control) {
  const shell = getComboboxShell(control);
  return shell?.querySelector("button[aria-label*='toggle' i], button[aria-haspopup], .select__indicators button") || null;
}

function getComboboxSelectedText(control) {
  const shell = getComboboxShell(control);
  const selected = shell?.querySelector(".select__single-value, [class*='single-value'], [aria-selected='true']");
  return cleanLabel(selected?.textContent || "");
}

function getComboboxValueText(control) {
  const input = getComboboxInput(control);
  return cleanLabel(input.value || control.value || getComboboxSelectedText(control) || control.textContent || "");
}

function getComboboxInput(control) {
  return control.matches("input, textarea") ? control : control.querySelector("input, textarea") || control;
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
  document.documentElement.removeAttribute(DEBUG_ATTR);
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
  } catch (error) {
    console.warn("Auto Bid could not export debug trace", error);
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
  const sponsorshipContext = /(sponsor|sponsorship|visa|work permit|work authorization support)/.test(normalizedLabel);
  const sponsorshipNeedContext = /(require|need|seek|request|support).*(sponsor|sponsorship|visa|work permit)|(sponsor|sponsorship|visa|work permit).*(require|need|support|request)/.test(normalizedLabel);
  const workAuthorizationContext = /(authorized|eligible|legally).*(work|employ)|(work|employ).*(authorized|eligible|legally)|without.*(sponsor|sponsorship|visa)/.test(normalizedLabel);
  const salaryComfortContext = /(comfortable|accept|agree|ok|okay).*(salary|compensation|pay|range)|(salary|compensation|pay|range).*(comfortable|accept|agree|ok|okay)/.test(normalizedLabel);
  const positiveYesContext = /(comfortable|authorized|eligible|willing|available|open to|agree|accept|consent|confirm|legally|relocat|remote|can you|able to|do you have|have you|do you use|have you used|use.*ai|ai.*assist|artificial intelligence|development workflow|meet.*requirement|salary|compensation|pay range)/.test(normalizedLabel);
  const yesChoice = /\b(yes|y|true|agree|accepted|authorized|eligible|willing|available|comfortable|ok|okay)\b/.test(normalizedChoice);
  const noChoice = /\b(no|n|false)\b|not required|do not|dont|does not|will not|won t|cannot|can not/.test(normalizedChoice);
  let score = 0;

  if (/select|choose|please|placeholder|n\/a|not applicable/.test(normalizedChoice)) return -200;

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

function isPlaceholderChoice(text, value) {
  const normalizedText = normalize(text);
  const normalizedValue = normalize(value);
  if (!normalizedText && !normalizedValue) return true;
  return /^(select|choose|please select|placeholder|n a|not applicable)$/.test(normalizedText);
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

function showStatus(message, kind) {
  let status = document.getElementById(STATUS_ID);
  if (!status) {
    status = document.createElement("div");
    status.id = STATUS_ID;
    document.documentElement.append(status);
  }

  status.textContent = message;
  status.style.cssText = [
    "position:fixed",
    "right:18px",
    "top:18px",
    "z-index:2147483647",
    "max-width:min(360px, calc(100vw - 36px))",
    "padding:10px 12px",
    "border-radius:8px",
    "font:13px/1.4 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif",
    "box-shadow:0 14px 40px rgba(19,35,29,.22)",
    kind === "error" ? "background:#fff7ed;color:#9a3412;border:1px solid #fdba74" : "background:#ecfdf5;color:#065f46;border:1px solid #86efac"
  ].join(";");

  window.clearTimeout(status._autoBidTimer);
  status._autoBidTimer = window.setTimeout(() => status.remove(), kind === "error" ? 6000 : 2600);
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
