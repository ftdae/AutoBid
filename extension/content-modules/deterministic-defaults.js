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
  }
})();
