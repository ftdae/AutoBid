(() => {
  if (window.__autoBidHotkeyListenerLoaded) return;
  window.__autoBidHotkeyListenerLoaded = true;
  const HOTKEY_SOURCE = "auto-bid-hotkey";

  document.addEventListener("keydown", (event) => {
    const messageType = getAutoBidHotkeyMessage(event);
    if (!messageType) return;
    event.preventDefault();
    event.stopPropagation();
    triggerHotkey(messageType).catch(() => {});
  }, true);

  window.addEventListener("message", (event) => {
    const data = event.data || {};
    if (data.source !== HOTKEY_SOURCE || !isAllowedHotkeyMessage(data.type)) return;
    sendRuntimeMessage(data.type).then((sent) => {
      if (!sent) traceHotkey("runtime-unavailable", { type: data.type });
    });
  }, true);

  async function triggerHotkey(messageType) {
    traceHotkey("pressed", { type: messageType });
    if (await sendRuntimeMessage(messageType)) return;

    try {
      const target = window.top && window.top !== window ? window.top : window;
      target.postMessage({ source: HOTKEY_SOURCE, type: messageType }, "*");
      traceHotkey("relayed-to-top-frame", { type: messageType });
    } catch (_error) {
      document.dispatchEvent(new CustomEvent("autoBid:hotkey", { detail: { type: messageType } }));
    }
  }

  function sendRuntimeMessage(messageType) {
    if (!isAllowedHotkeyMessage(messageType)) return Promise.resolve(false);
    const runtime = globalThis.chrome?.runtime;
    if (!runtime?.sendMessage) return Promise.resolve(false);

    return new Promise((resolve) => {
      try {
        runtime.sendMessage({ type: messageType }, () => {
          const error = runtime.lastError;
          if (error) {
            traceHotkey("runtime-send-failed", { type: messageType, message: error.message || String(error) });
            resolve(false);
            return;
          }
          traceHotkey("runtime-sent", { type: messageType });
          resolve(true);
        });
      } catch (error) {
        traceHotkey("runtime-send-threw", { type: messageType, message: error.message || String(error) });
        resolve(false);
      }
    });
  }

  function isAllowedHotkeyMessage(messageType) {
    return messageType === "HOTKEY_TRIGGER";
  }

  function traceHotkey(stage, detail = {}) {
    try {
      console.info(`[AutoBid] hotkey:${stage}`, detail);
    } catch (_error) {
      // Ignore console failures on restricted extension pages.
    }
  }

  function getAutoBidHotkeyMessage(event) {
    if (event.defaultPrevented || event.repeat) return false;
    const key = String(event.key || "").toLowerCase();
    const code = String(event.code || "").toLowerCase();
    const isQ = key === "q" || code === "keyq";
    if (isQ && event.ctrlKey && !event.shiftKey && !event.altKey && !event.metaKey) return "HOTKEY_TRIGGER";
    return "";
  }
})();
