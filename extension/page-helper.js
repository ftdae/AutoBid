(() => {
  if (window.__autoBidPageHelperLoaded) {
    markReady();
    return;
  }
  window.__autoBidPageHelperLoaded = true;
  markReady();
  document.addEventListener("DOMContentLoaded", markReady, { once: true });

  document.addEventListener("autoBid:pageCommand", () => {
    const raw = getRoot()?.getAttribute("data-auto-bid-command");
    if (!raw) return;

    let command;
    try {
      command = JSON.parse(raw);
    } catch (_error) {
      return;
    }

    const element = queryDeep(`[data-auto-bid-bridge-token="${cssEscape(command.token)}"]`);
    if (!element) return;

    if (command.type === "click") dispatchRealisticMouseClick(element);
    if (command.type === "input") setNativeValue(element, command.value || "");
    if (command.type === "checked") setNativeChecked(element, Boolean(command.checked));
    if (command.type === "key" && !callReactHandlers(element, ["onKeyDown"], "keydown", { key: command.key || "Enter" })) {
      dispatchKey(element, command.key || "Enter");
    }
    if (command.type === "combobox-open") {
      if (callReactHandlers(element, ["onMouseDown"], "mousedown")) {
        callReactHandlers(element, ["onFocus"], "focus");
      } else {
        dispatchMouseDown(element);
      }
    }
    if (command.type === "combobox-toggle") element.click();
    if (command.type === "combobox-choose" && !callReactHandlers(element, ["onClick"], "click")) {
      dispatchClick(element);
    }
    if (command.type === "file-upload") {
      uploadFile(element, command.file || {});
    }
    getRoot()?.setAttribute("data-auto-bid-command-result", command.token);
  });

  function setNativeValue(element, value) {
    const setter = Object.getOwnPropertyDescriptor(element.constructor.prototype, "value")?.set;
    if (setter) setter.call(element, value);
    else element.value = value;

    const event = typeof InputEvent === "function"
      ? new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText", data: value })
      : new Event("input", { bubbles: true, composed: true });
    element.dispatchEvent(event);
    element.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
    callReactHandlers(element, ["onInput"], "input");
    callReactHandlers(element, ["onChange"], "change");
  }

  function setNativeChecked(element, checked) {
    const setter = Object.getOwnPropertyDescriptor(element.constructor.prototype, "checked")?.set;
    if (setter) setter.call(element, checked);
    else element.checked = checked;

    element.dispatchEvent(new Event("input", { bubbles: true, composed: true }));
    element.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
    callReactHandlers(element, ["onChange"], "change");
  }

  function uploadFile(input, payload) {
    if (!input || !input.matches?.("input[type='file']")) return false;

    const file = fileFromBase64(payload);
    if (!file) return false;

    const transfer = new DataTransfer();
    transfer.items.add(file);
    setInputFiles(input, transfer.files);
    input.focus?.();

    dispatchFileInput(input, transfer, "input");
    dispatchFileInput(input, transfer, "change");
    const inputHandled = callReactHandlers(input, ["onInput"], "input", { dataTransfer: transfer }) ||
      callReactHandlers(input, ["onChange"], "change", { dataTransfer: transfer });

    if (!inputHandled) {
      const target = getFileDropTargets(input)[0];
      if (target) {
        dispatchFileDrop(target, transfer, "dragenter");
        dispatchFileDrop(target, transfer, "dragover");
        callReactHandlers(target, ["onDragEnter"], "dragenter", { dataTransfer: transfer });
        callReactHandlers(target, ["onDragOver"], "dragover", { dataTransfer: transfer });
        dispatchFileDrop(target, transfer, "drop");
        callReactHandlers(target, ["onDrop"], "drop", { dataTransfer: transfer });
      }
    }

    input.blur?.();
    return true;
  }

  function fileFromBase64(payload) {
    const base64 = String(payload.base64 || "").replace(/^data:[^,]+,/, "");
    if (!base64) return null;

    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }

    const name = sanitizeFilename(payload.name || payload.filename || "resume.pdf");
    const type = String(payload.mime_type || payload.mimeType || "application/pdf").trim() || "application/pdf";
    return new File([bytes], name, { type });
  }

  function sanitizeFilename(value) {
    return String(value || "resume.pdf").replace(/[\\/:*?"<>|]+/g, "_").trim() || "resume.pdf";
  }

  function setInputFiles(input, files) {
    try {
      input.files = files;
      return;
    } catch (_error) {
      // Some frameworks wrap file inputs; defineProperty is a fallback for their event handlers.
    }

    try {
      Object.defineProperty(input, "files", {
        configurable: true,
        get() {
          return files;
        }
      });
    } catch (_error) {
      // If this also fails, the dispatched drop events still carry the file list.
    }
  }

  function dispatchFileInput(input, transfer, type) {
    const event = new Event(type, { bubbles: true, cancelable: true, composed: true });
    defineFileEventProperties(event, input, transfer);
    input.dispatchEvent(event);
  }

  function dispatchFileDrop(target, transfer, type) {
    let event;
    try {
      event = new DragEvent(type, {
        bubbles: true,
        cancelable: true,
        composed: true,
        dataTransfer: transfer
      });
    } catch (_error) {
      event = new Event(type, { bubbles: true, cancelable: true, composed: true });
    }
    defineFileEventProperties(event, target, transfer);
    target.dispatchEvent(event);
  }

  function defineFileEventProperties(event, target, transfer) {
    defineReadonly(event, "target", target);
    defineReadonly(event, "currentTarget", target);
    defineReadonly(event, "dataTransfer", transfer);
  }

  function defineReadonly(object, property, value) {
    try {
      Object.defineProperty(object, property, { configurable: true, value });
    } catch (_error) {
      // Native event implementations may prevent overriding some properties.
    }
  }

  function getFileDropTargets(input) {
    const candidates = [
      input.closest("label"),
      input.closest("fieldset"),
      input.closest("[class*='upload' i], [class*='dropzone' i], [class*='drop-zone' i], [class*='attachment' i], [class*='resume' i], [class*='cv' i], [data-testid*='upload' i], [data-testid*='drop' i], [data-testid*='file' i]"),
      findNearestUploadZone(input),
      input.parentElement,
      input
    ].filter(Boolean);
    return Array.from(new Set(candidates));
  }

  function findNearestUploadZone(input) {
    const zones = Array.from(document.querySelectorAll([
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
      "[data-testid*='file' i]"
    ].join(","))).filter((element) => isResumeUploadZoneText(element.textContent || element.getAttribute?.("aria-label") || ""));

    if (zones.length === 0) return null;
    const inputForm = input.closest("form");
    return zones
      .map((zone) => ({
        zone,
        score: (zone.contains(input) ? 1000 : 0) +
          (inputForm && zone.closest("form") === inputForm ? 200 : 0) +
          (input.parentElement && zone.parentElement === input.parentElement ? 120 : 0)
      }))
      .filter((candidate) => candidate.score > 0)
      .sort((left, right) => right.score - left.score)[0]?.zone || zones[0];
  }

  function isResumeUploadZoneText(value) {
    const text = normalizeText(value);
    if (!text || /(cover letter|motivation letter|portfolio|photo|avatar|image|transcript|certificate)/.test(text)) return false;
    return (/\b(resume|cv|curriculum vitae)\b/.test(text) &&
      /\b(upload|attach|attachment|file|drop|drag|browse|choose)\b/.test(text)) ||
      isGenericApplicationFileUploadText(text);
  }

  function isGenericApplicationFileUploadText(value) {
    const text = normalizeText(value);
    if (!text || /(cover letter|motivation letter|dropbox|google drive|drive|manual|manually|paste|photo|avatar|image|portfolio|certificate|transcript)/.test(text)) {
      return false;
    }
    const hasGenericUpload = /\b(choose|select|upload|attach|browse|drop|drag)\b.*\bfile\b|\bfile\b.*\b(drop|upload|attach|browse|choose|select)\b/.test(text);
    if (!hasGenericUpload) return false;
    return /\b(easy apply|autocomplete your application|application|personal information|apply|mb size limit|size limit|pdf|doc|docx|rtf|txt)\b/.test(text);
  }

  function normalizeText(value) {
    return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  }

  function dispatchKey(element, key) {
    const keyCode = key === "ArrowDown" ? 40 : key === " " ? 32 : key === "Escape" ? 27 : 13;
    element.focus?.();
    element.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, composed: true, key, code: key, keyCode, which: keyCode }));
    element.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, composed: true, key, code: key, keyCode, which: keyCode }));
  }

  function dispatchRealisticMouseClick(element) {
    const target = getHitTarget(element);
    const rect = target.getBoundingClientRect();
    const clientX = Math.max(rect.left + Math.min(rect.width / 2, rect.width - 2), rect.left + 1);
    const clientY = Math.max(rect.top + Math.min(rect.height / 2, rect.height - 2), rect.top + 1);
    const base = { bubbles: true, cancelable: true, composed: true, view: window, clientX, clientY, screenX: clientX, screenY: clientY, button: 0, buttons: 1 };

    dispatchPointer(target, "pointerover", base);
    dispatchPointer(target, "pointerenter", { ...base, bubbles: false });
    target.dispatchEvent(new MouseEvent("mouseover", base));
    target.dispatchEvent(new MouseEvent("mouseenter", { ...base, bubbles: false }));
    dispatchPointer(target, "pointermove", base);
    target.dispatchEvent(new MouseEvent("mousemove", base));
    dispatchPointer(target, "pointerdown", base);
    target.dispatchEvent(new MouseEvent("mousedown", base));
    dispatchPointer(target, "pointerup", { ...base, buttons: 0 });
    target.dispatchEvent(new MouseEvent("mouseup", { ...base, buttons: 0 }));
    target.dispatchEvent(new MouseEvent("click", { ...base, buttons: 0, detail: 1 }));
    target.click?.();
  }

  function dispatchMouseDown(element) {
    const target = getHitTarget(element);
    const rect = target.getBoundingClientRect();
    const clientX = Math.max(rect.left + Math.min(rect.width / 2, rect.width - 2), rect.left + 1);
    const clientY = Math.max(rect.top + Math.min(rect.height / 2, rect.height - 2), rect.top + 1);
    const base = {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX,
      clientY,
      screenX: clientX,
      screenY: clientY,
      button: 0,
      buttons: 1
    };

    dispatchPointer(target, "pointerdown", base);
    target.dispatchEvent(new MouseEvent("mousedown", base));
  }

  function dispatchClick(element) {
    const target = getHitTarget(element);
    const rect = target.getBoundingClientRect();
    const clientX = Math.max(rect.left + Math.min(rect.width / 2, rect.width - 2), rect.left + 1);
    const clientY = Math.max(rect.top + Math.min(rect.height / 2, rect.height - 2), rect.top + 1);
    target.dispatchEvent(new MouseEvent("click", {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX,
      clientY,
      screenX: clientX,
      screenY: clientY,
      button: 0,
      buttons: 0,
      detail: 1
    }));
  }

  function callReactHandlers(element, names, eventType = "mousedown", extra = {}) {
    const reactTarget = findReactProps(element, names);
    if (!reactTarget) return false;

    const rect = element.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;
    const key = extra.key || "";
    const code = key === " " ? "Space" : key;
    const keyCode = key === "ArrowDown" ? 40 : key === " " ? 32 : key === "Escape" ? 27 : key === "Enter" ? 13 : 0;
    const event = {
      target: element,
      currentTarget: reactTarget.element,
      type: eventType,
      button: 0,
      buttons: 1,
      key,
      code,
      keyCode,
      which: keyCode,
      clientX,
      clientY,
      screenX: clientX,
      screenY: clientY,
      isTrusted: true,
      defaultPrevented: false,
      propagationStopped: false,
      nativeEvent: {
        target: element,
        currentTarget: reactTarget.element,
        type: eventType,
        button: 0,
        buttons: 1,
        key,
        code,
        keyCode,
        which: keyCode,
        clientX,
        clientY,
        dataTransfer: extra.dataTransfer || undefined
      },
      dataTransfer: extra.dataTransfer || undefined,
      preventDefault() {
        this.defaultPrevented = true;
      },
      stopPropagation() {
        this.propagationStopped = true;
      },
      isDefaultPrevented() {
        return this.defaultPrevented;
      },
      isPropagationStopped() {
        return this.propagationStopped;
      },
      persist() {}
    };

    let called = false;
    names.forEach((name) => {
      if (typeof reactTarget.props[name] === "function") {
        reactTarget.props[name](event);
        called = true;
      }
    });
    return called;
  }

  function findReactProps(element, names) {
    let current = element;
    while (current && current !== document.documentElement) {
      const key = Object.keys(current).find((item) => item.startsWith("__reactProps$"));
      if (key && names.some((name) => typeof current[key]?.[name] === "function")) {
        return { element: current, props: current[key] };
      }
      current = current.parentElement;
    }
    return null;
  }

  function dispatchPointer(target, type, init) {
    if (typeof PointerEvent === "function") {
      target.dispatchEvent(new PointerEvent(type, { ...init, pointerType: "mouse" }));
    }
  }

  function getHitTarget(element) {
    const rect = element.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const hit = document.elementFromPoint(x, y);
    if (hit && (element === hit || element.contains(hit))) return hit;
    return element;
  }

  function cssEscape(value) {
    if (window.CSS?.escape) return CSS.escape(value);
    return String(value || "").replace(/["\\]/g, "\\$&");
  }

  function queryDeep(selector, root = document) {
    const stack = [root];
    const seen = new Set();

    while (stack.length) {
      const current = stack.shift();
      if (!current || seen.has(current)) continue;
      seen.add(current);

      if (current.matches?.(selector)) return current;
      const match = current.querySelector?.(selector);
      if (match) return match;

      current.querySelectorAll?.("*").forEach((element) => {
        if (element.shadowRoot) stack.push(element.shadowRoot);
      });
    }

    return null;
  }

  function getRoot() {
    return document.documentElement || document.head || document.body;
  }

  function markReady() {
    getRoot()?.setAttribute("data-auto-bid-page-helper", "ready");
    document.documentElement?.setAttribute("data-auto-bid-page-helper", "ready");
  }
})();
