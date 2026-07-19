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

    const element = document.querySelector(`[data-auto-bid-bridge-token="${cssEscape(command.token)}"]`);
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
        clientY
      },
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

  function getRoot() {
    return document.documentElement || document.head || document.body;
  }

  function markReady() {
    getRoot()?.setAttribute("data-auto-bid-page-helper", "ready");
    document.documentElement?.setAttribute("data-auto-bid-page-helper", "ready");
  }
})();
