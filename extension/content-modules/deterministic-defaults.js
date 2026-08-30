(() => {
  if (window.AutoBidDeterministicDefaults?.version) return;

  window.AutoBidDeterministicDefaults = {
    version: "0.1.0",
    create
  };

  function create(helpers) {
    const {
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
    } = helpers;

    return { apply };

    async function apply(fields, filledIds) {
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

      if (isAuthorizationSupportRequiredChoiceField(field, controls)) {
        return { value: "No", reason: "authorization-support-not-required" };
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

    function hasYesNoOptions(options) {
      const normalized = (options || []).map(normalize);
      return normalized.includes("yes") && normalized.includes("no");
    }

    function isAvailabilityDateField(field, controls) {
      const control = controls?.[0];
      const type = getControlType(control);
      if (!["date", "text", "search"].includes(type)) return false;
      const label = getFieldContextLabel(field, control);
      return /(date available|available date|available start date|earliest.*start|start date|when.*start|availability date)/.test(label);
    }

    function getExperienceYearsDefault(field, controls) {
      const control = controls?.[0];
      const type = getControlType(control);
      const label = getFieldContextLabel(field, control);
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
      const label = getFieldContextLabel(field, control);
      return isExperienceYearsLabel(label);
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
      if (!text || /^(select|choose|please select|placeholder|n a|not applicable)$/.test(text)) return null;
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
      return /^(0|0\.0+|0\s*(?:years?|yrs?)?)$/i.test(String(current || "").trim());
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
        const result = await send("NATIVE_TYPE", { text: String(value), commit: commit });
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

      return getRangeSearchRoots(control)
        .flatMap((root) => Array.from(root.querySelectorAll(selector)))
        .filter((element, index, list) => list.indexOf(element) === index)
        .map((element) => scoreRangeValueCandidate(element, control, rangeRect))
        .filter(Boolean)
        .sort((left, right) => right.score - left.score);
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
  }
})();
