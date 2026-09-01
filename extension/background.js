const DEFAULT_API_BASE = "http://localhost:7003";
const DEV_AUTH_BYPASS = true;
const DEV_PROFILE_ID = "abp_dev_default";
const DEV_USER = {
  id: "abu_dev_local",
  first_name: "Dev",
  last_name: "User",
  email: "dev@autobid.local",
  name: "Dev User",
  timezone: "UTC"
};
const AUTO_BID_TRIGGER_DEDUP_MS = 1500;
const AUTO_BID_WINDOW_STORAGE_KEY = "autoBidWindowId";
const EXECUTION_LOG_STORAGE_KEY = "autoBidExecutionLogsV1";
const EXECUTION_LOG_MAX_RUNS = 100;
const GPT_WORKER_TAB_STORAGE_KEY = "autoBidGptWorkerTabId";
const LEGACY_GPT_BATCH_STATE_STORAGE_KEY = "autoBidGptBatchStateV1";
const GPT_BATCH_STATES_STORAGE_KEY = "autoBidGptBatchStatesV2";
const GPT_BATCH_PAUSED_STORAGE_KEY = "autoBidGptBatchPaused";
const RUNTIME_GPT_QUEUE_STORAGE_KEY = "autoBidRuntimeGptQueueV2";
const GPT_WORKER_URL = "https://chatgpt.com/?autobid_worker=1";
// Keep one application per prompt, with three persistent workers running in
// parallel. The durable request queue itself is unlimited.
const RUNTIME_GPT_PROMPT_BATCH_SIZE = 1;
const RUNTIME_GPT_MAX_WORKERS = 3;
const RUNTIME_GPT_REQUEST_TTL_MS = 24 * 60 * 60 * 1000;
const RUNTIME_GPT_LEASE_MS = 5 * 60 * 1000;
const RUNTIME_GPT_RETRY_DELAY_MS = 5000;
const RUNTIME_GPT_DELIVERY_ATTEMPTS = 3;
const RUNTIME_GPT_DELIVERY_RETRY_MS = 1000;
const RUNTIME_GPT_DELIVERY_ALARM_PREFIX = "autobid-gpt-delivery:";
const AUTO_BID_WINDOW_SIZE = { width: 500, height: 760 };
const AUTO_BID_WINDOW_EDGE_GAP = 24;
const nativeInputQueues = new Map();
const backgroundAutomationHolds = new Map();
const runtimeGptDeliveryPromises = new Map();
const autoBidTriggerTimes = new Map();
const runtimeGptRequests = new Map();
let gptBatchPumpPromise = null;
let gptBatchPumpRequested = false;
let runtimeGptBatchStateMutationPromise = Promise.resolve();
let runtimeGptRequestsLoaded = false;
let runtimeGptRequestsLoadPromise = null;
let runtimeGptPersistPromise = Promise.resolve();
let nextRuntimeGptQueueSequence = 1;
let executionLogWritePromise = Promise.resolve();

chrome.runtime.onInstalled.addListener(async () => {
  const settings = await chrome.storage.local.get(["apiBase"]);
  if (!settings.apiBase) {
    await chrome.storage.local.set({ apiBase: DEFAULT_API_BASE });
  }
  await enforceRuntimeGptWorkerLimit().catch((error) => {
    console.warn("[AutoBid GPT Pool] Could not trim legacy worker tabs after reload", error);
  });
});

chrome.runtime.onStartup?.addListener(() => {
  enforceRuntimeGptWorkerLimit().catch((error) => {
    console.warn("[AutoBid GPT Pool] Could not trim legacy worker tabs at startup", error);
  });
});

chrome.commands.onCommand.addListener((command) => {
  if (command === "trigger-auto-bid") {
    triggerAutoBidInActiveTab().catch((error) => {
      console.error("Auto Bid shortcut failed", error);
    });
    return;
  }
});

chrome.action.onClicked.addListener((tab) => {
  openAutoBidSurface(tab).catch((error) => {
    console.error("Auto Bid surface failed", error);
  });
});

async function openAutoBidSurface(tab) {
  if (Number.isInteger(tab?.id) && isInjectablePageUrl(tab.url)) {
    try {
      await togglePanelInTab(tab.id);
      return { surface: "panel", tabId: tab.id };
    } catch (error) {
      console.warn("Auto Bid panel injection failed; opening extension window", error);
    }
  }
  return openAutoBidWindow();
}

function isInjectablePageUrl(url) {
  return /^(https?|file):/i.test(String(url || ""));
}

async function triggerAutoBidInActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  return triggerAutoBidInTab(tab.id, "command");
}

async function togglePanelInTab(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["panel-host.js"]
  });
}

async function openAutoBidWindow() {
  const existing = await getStoredAutoBidWindow();
  if (existing?.id) {
    await chrome.windows.update(existing.id, {
      focused: true,
      state: "normal",
      drawAttention: true
    });
    return { reused: true, windowId: existing.id };
  }

  const current = await chrome.windows.getLastFocused().catch(() => chrome.windows.getCurrent()).catch(() => null);
  const position = getAutoBidWindowPosition(current);
  const created = await chrome.windows.create({
    url: chrome.runtime.getURL("popup.html?surface=window"),
    type: "popup",
    focused: true,
    width: AUTO_BID_WINDOW_SIZE.width,
    height: AUTO_BID_WINDOW_SIZE.height,
    left: position.left,
    top: position.top
  });

  if (created?.id) {
    await chrome.storage.local.set({ [AUTO_BID_WINDOW_STORAGE_KEY]: created.id });
  }

  return { reused: false, windowId: created?.id || null };
}

async function getStoredAutoBidWindow() {
  const stored = await chrome.storage.local.get([AUTO_BID_WINDOW_STORAGE_KEY]).catch(() => ({}));
  const windowId = Number(stored[AUTO_BID_WINDOW_STORAGE_KEY]);
  if (!Number.isInteger(windowId)) return null;

  try {
    return await chrome.windows.get(windowId);
  } catch (_error) {
    await chrome.storage.local.remove([AUTO_BID_WINDOW_STORAGE_KEY]).catch(() => {});
    return null;
  }
}

function getAutoBidWindowPosition(currentWindow) {
  const width = AUTO_BID_WINDOW_SIZE.width;
  const height = AUTO_BID_WINDOW_SIZE.height;
  const currentLeft = Number(currentWindow?.left || 0);
  const currentTop = Number(currentWindow?.top || 0);
  const currentWidth = Number(currentWindow?.width || 1280);
  const currentHeight = Number(currentWindow?.height || 900);

  return {
    left: Math.max(0, Math.round(currentLeft + currentWidth - width - AUTO_BID_WINDOW_EDGE_GAP)),
    top: Math.max(0, Math.round(currentTop + Math.min(80, Math.max(AUTO_BID_WINDOW_EDGE_GAP, currentHeight - height - AUTO_BID_WINDOW_EDGE_GAP))))
  };
}

chrome.windows.onRemoved.addListener((windowId) => {
  chrome.storage.local.get([AUTO_BID_WINDOW_STORAGE_KEY])
    .then((stored) => {
      if (Number(stored[AUTO_BID_WINDOW_STORAGE_KEY]) === windowId) {
        return chrome.storage.local.remove([AUTO_BID_WINDOW_STORAGE_KEY]);
      }
      return null;
    })
    .catch(() => {});
});

chrome.tabs.onActivated.addListener(({ tabId }) => {
  retryCompletedRuntimeGptDeliveriesForTab(tabId).catch(() => {});
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (changeInfo.status !== "complete") return;
  retryCompletedRuntimeGptDeliveriesForTab(tabId).catch(() => {});
});

chrome.tabs.onRemoved.addListener((tabId) => {
  backgroundAutomationHolds.delete(tabId);
  handleRuntimeGptBatchTabRemoved(tabId).catch((error) => {
    console.warn("[AutoBid Batch] Could not recover a closed ChatGPT batch tab", error);
  });
});

chrome.alarms?.onAlarm?.addListener((alarm) => {
  const name = String(alarm?.name || "");
  if (!name.startsWith(RUNTIME_GPT_DELIVERY_ALARM_PREFIX)) return;
  const requestId = name.slice(RUNTIME_GPT_DELIVERY_ALARM_PREFIX.length);
  ensureRuntimeGptRequestsLoaded()
    .then(() => {
      const request = runtimeGptRequests.get(requestId);
      if (request?.status === "complete" && !request.delivered_at) {
        return queueRuntimeGptAnswerDelivery(request);
      }
      return null;
    })
    .catch((error) => {
      console.warn("[AutoBid] Scheduled background answer delivery failed", {
        request_id: requestId,
        error: error.message || String(error)
      });
    });
});

async function injectAutoBidScripts(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId, allFrames: true },
    files: ["page-helper.js"],
    world: "MAIN"
  });
  await chrome.scripting.executeScript({
    target: { tabId, allFrames: true },
    files: [
      "content-modules/ats-adapters.js",
      "content-modules/deterministic-defaults.js",
      "content.js"
    ]
  });
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  handleMessage(message, sender)
    .then((data) => sendResponse({ ok: true, data }))
    .catch((error) => sendResponse({ ok: false, error: error.message || String(error) }));
  return true;
});

async function handleMessage(message, sender) {
  switch (message?.type) {
    case "HOTKEY_TRIGGER":
      return triggerAutoBidFromSender(sender);
    case "GPT_HOTKEY_TRIGGER":
      return runGptAnswerWorker(message.payload || {});
    case "GPT_STOP_HOTKEY_TRIGGER":
      return stopGptAnswerWorker();
    case "GET_SETTINGS":
      return getSettings();
    case "TRIGGER_ACTIVE_AUTOBID":
      return triggerAutoBidInActiveTab();
    case "OPEN_OPTIONS":
      await chrome.runtime.openOptionsPage();
      return { opened: true };
    case "AUTOBID_LOG_UPSERT":
      return upsertExecutionLog(message.payload || {}, sender);
    case "AUTOBID_AUTOFILL_STATE":
      return updateAutofillTabState(message.payload || {}, sender);
    case "GET_EXECUTION_LOGS":
      return getExecutionLogs(message.payload || {});
    case "CLEAR_EXECUTION_LOGS":
      return clearExecutionLogs();
    case "OUTLOOK_STATUS":
      return getOutlookStatus();
    case "OUTLOOK_CONNECT":
      return connectOutlook();
    case "OUTLOOK_DISCONNECT":
      return disconnectOutlook();
    case "OUTLOOK_LIST_VERIFICATION":
      return listOutlookVerificationMessages(message.payload || {});
    case "OUTLOOK_FIND_VERIFICATION":
      return findOutlookVerificationCode(message.payload || {});
    case "OUTLOOK_MARK_READ":
      return markOutlookMessageRead(message.payload || {});
    case "DEV_SESSION":
      return ensureDevSession();
    case "SAVE_SETTINGS":
      return saveSettings(message.payload || {});
    case "SAVE_SHEET_SETTINGS":
      return saveSheetSettings(message.payload || {});
    case "OPEN_SHEET_ROWS":
      return openSheetRows(message.payload || {});
    case "RUN_GPT_ANSWER":
      return runGptAnswerWorker(message.payload || {});
    case "STOP_GPT_ANSWER":
      return stopGptAnswerWorker();
    case "GPT_ANSWER_REQUEST":
      return createRuntimeGptAnswerRequest(message.payload || {}, sender);
    case "GPT_ANSWER_STATUS":
      return getRuntimeGptAnswerStatus(message.payload || {});
    case "GPT_ANSWER_CANCEL":
      return cancelRuntimeGptAnswerRequest(message.payload || {});
    case "AUTOBID_GPT_FETCH_PENDING_REQUESTS":
      return fetchPendingRuntimeGptRequests(message.payload || {}, sender);
    case "AUTOBID_GPT_SAVE_REQUEST_ANSWERS":
      return saveRuntimeGptRequestAnswers(message.payload || {});
    case "AUTOBID_GPT_BATCH_COMPLETE":
      return completeRuntimeGptBatch(message.payload || {}, sender);
    case "AUTOBID_GPT_WORKER_READY":
      return handleRuntimeGptWorkerReady(message.payload || {}, sender);
    case "GPT_ANSWER_APPLIED":
      return acknowledgeRuntimeGptAnswersApplied(message.payload || {}, sender);
    case "AUTOBID_GPT_FAIL_REQUEST":
      return failRuntimeGptRequest(message.payload || {});
    case "SHEET_CONTEXT":
      return getSheetContextForPage(message.payload || {}, sender);
    case "SHEET_SUBMIT_QUESTIONS":
      return submitSheetQuestions(message.payload || {}, sender);
    case "SHEET_FETCH_ANSWERS":
      return fetchSheetAnswers(message.payload || {}, sender);
    case "SHEET_FETCH_RESUME_FILE":
      return fetchSheetResumeFile(message.payload || {}, sender);
    case "AUTOBID_GPT_FETCH_PENDING_ROWS":
      return fetchPendingGptRows();
    case "AUTOBID_GPT_SAVE_ANSWERS":
      return saveGptAnswers(message.payload || {});
    case "AUTOBID_GPT_WORKER_STATE":
      await chrome.storage.local.set({ autoBidGptWorkerState: message.payload || {} });
      return { ok: true };
    case "SIGNUP":
      return authRequest("/auth/signup", message.payload || {});
    case "LOGIN":
      return authRequest("/auth/login", message.payload || {});
    case "LOGOUT":
      await chrome.storage.local.remove(["token", "user", "selectedProfileId"]);
      return { ok: true };
    case "LIST_PROFILES":
      return listProfiles();
    case "SAVE_PROFILE":
      return saveProfile(message.payload || {});
    case "DELETE_PROFILE":
      return deleteProfile(message.profileId);
    case "SELECT_PROFILE":
      await chrome.storage.local.set({ selectedProfileId: message.profileId || null });
      return getSettings();
    case "GET_PROFILE_STATIC_FIELDS":
      return getSelectedProfileStaticFields();
    case "ASSIST":
      return assist(message.payload || {});
    case "NATIVE_CLICK":
      return queueNativeClick(sender, message.payload || {});
    case "NATIVE_TYPE":
      return queueNativeType(sender, message.payload || {});
    case "NATIVE_FILE_UPLOAD":
      return nativeFileUpload(sender, message.payload || {});
    case "NATIVE_FILE_CHOOSER_UPLOAD":
      return nativeFileChooserUpload(sender, message.payload || {});
    case "DRAFT_STATUS":
      return apiFetch(`/drafts/${encodeURIComponent(message.draftId)}/status`, {
        method: "POST",
        body: { status: message.status }
      });
    default:
      throw new Error("Unknown Auto Bid message");
  }
}

function upsertExecutionLog(payload, sender) {
  executionLogWritePromise = executionLogWritePromise
    .catch(() => {})
    .then(async () => {
      const runId = String(payload.run_id || payload.runId || "").slice(0, 120);
      if (!runId) throw new Error("Execution log run ID is required");

      const stored = await chrome.storage.local.get([EXECUTION_LOG_STORAGE_KEY]).catch(() => ({}));
      const current = Array.isArray(stored[EXECUTION_LOG_STORAGE_KEY]) ? stored[EXECUTION_LOG_STORAGE_KEY] : [];
      const existing = current.find((run) => run.run_id === runId) || {};
      const tabId = Number.isInteger(sender?.tab?.id) ? sender.tab.id : existing.tab_id || null;
      const entries = Array.isArray(payload.entries)
        ? payload.entries.slice(-240).map(normalizeExecutionLogEntry)
        : Array.isArray(existing.entries) ? existing.entries : [];
      const nextRun = {
        ...existing,
        run_id: runId,
        tab_id: tabId,
        frame_id: Number.isInteger(sender?.frameId) ? sender.frameId : existing.frame_id || 0,
        url: String(payload.url || existing.url || sender?.tab?.url || "").slice(0, 4000),
        title: String(payload.title || existing.title || sender?.tab?.title || "").slice(0, 500),
        ats: normalizeExecutionAts(payload.ats || existing.ats),
        status: normalizeExecutionStatus(payload.status || existing.status),
        started_at: payload.started_at || existing.started_at || new Date().toISOString(),
        updated_at: payload.updated_at || new Date().toISOString(),
        completed_at: payload.completed_at || existing.completed_at || null,
        summary: normalizeExecutionSummary(payload.summary || existing.summary),
        entries
      };
      const next = [nextRun, ...current.filter((run) => run.run_id !== runId)]
        .sort((left, right) => String(right.updated_at || "").localeCompare(String(left.updated_at || "")))
        .slice(0, EXECUTION_LOG_MAX_RUNS);
      await chrome.storage.local.set({ [EXECUTION_LOG_STORAGE_KEY]: next });
      return nextRun;
    });
  return executionLogWritePromise;
}

async function getExecutionLogs(payload = {}) {
  await executionLogWritePromise.catch(() => {});
  const stored = await chrome.storage.local.get([EXECUTION_LOG_STORAGE_KEY]).catch(() => ({}));
  const runs = Array.isArray(stored[EXECUTION_LOG_STORAGE_KEY]) ? stored[EXECUTION_LOG_STORAGE_KEY] : [];
  const limit = Math.min(100, Math.max(1, Number(payload.limit || 30)));
  return { runs: runs.slice(0, limit) };
}

async function clearExecutionLogs() {
  await executionLogWritePromise.catch(() => {});
  await chrome.storage.local.remove([EXECUTION_LOG_STORAGE_KEY]);
  return { cleared: true };
}

function normalizeExecutionLogEntry(entry) {
  return {
    at: String(entry?.at || new Date().toISOString()).slice(0, 40),
    event: String(entry?.event || "event").slice(0, 160),
    data: sanitizeLogValue(entry?.data, 0)
  };
}

function sanitizeLogValue(value, depth) {
  if (depth > 4) return "[truncated]";
  if (value == null || typeof value === "boolean" || typeof value === "number") return value;
  if (typeof value === "string") return value.slice(0, 4000);
  if (Array.isArray(value)) return value.slice(0, 80).map((item) => sanitizeLogValue(item, depth + 1));
  if (typeof value === "object") {
    return Object.fromEntries(Object.entries(value).slice(0, 80).map(([key, item]) => [
      key,
      /(password|secret|token|authorization|private[_-]?key|base64)/i.test(key) ? "[redacted]" : sanitizeLogValue(item, depth + 1)
    ]));
  }
  return String(value).slice(0, 1000);
}

function normalizeExecutionAts(value) {
  return {
    id: String(value?.id || "common").slice(0, 80),
    name: String(value?.name || "Common form").slice(0, 120)
  };
}

function normalizeExecutionStatus(value) {
  const status = String(value || "running").toLowerCase();
  return ["running", "completed", "failed", "cancelled"].includes(status) ? status : "running";
}

function normalizeExecutionSummary(value) {
  return {
    filled: Math.max(0, Number(value?.filled || 0)),
    missed: Math.max(0, Number(value?.missed || 0)),
    submitted: Boolean(value?.submitted),
    message: String(value?.message || "").slice(0, 1000)
  };
}

async function getOutlookStatus() {
  return apiFetch("/outlook/connection");
}

async function connectOutlook() {
  const redirectUri = chrome.identity.getRedirectURL("outlook");
  const authorization = await apiFetch("/outlook/oauth/start", {
    method: "POST",
    body: { redirect_uri: redirectUri }
  });
  if (!authorization?.authorization_url) throw new Error("Auto Bid server did not return a Microsoft authorization URL");

  const finalUrl = await chrome.identity.launchWebAuthFlow({
    url: authorization.authorization_url,
    interactive: true
  });
  if (!finalUrl) throw new Error("Microsoft sign-in was cancelled");

  const callback = new URL(finalUrl);
  const oauthError = callback.searchParams.get("error_description") || callback.searchParams.get("error");
  if (oauthError) throw new Error(oauthError);
  const code = callback.searchParams.get("code");
  const state = callback.searchParams.get("state");
  if (!code || !state) throw new Error("Microsoft sign-in did not return the expected authorization response");

  return apiFetch("/outlook/oauth/callback", {
    method: "POST",
    body: { code, state, redirect_uri: redirectUri }
  });
}

async function disconnectOutlook() {
  return apiFetch("/outlook/connection", { method: "DELETE" });
}

async function listOutlookVerificationMessages(payload = {}) {
  const params = new URLSearchParams();
  params.set("top", String(Math.min(40, Math.max(1, Number(payload.top || 20)))));
  if (payload.domain) params.set("domain", String(payload.domain));
  return apiFetch(`/outlook/messages?${params}`);
}

async function findOutlookVerificationCode(payload = {}) {
  const params = new URLSearchParams();
  params.set("top", String(Math.min(10, Math.max(1, Number(payload.top || 10)))));
  if (payload.domain) params.set("domain", String(payload.domain));
  return apiFetch(`/outlook/latest-code?${params}`);
}

async function markOutlookMessageRead(payload = {}) {
  const messageId = String(payload.messageId || payload.message_id || "");
  if (!messageId) throw new Error("Outlook message ID is required");
  return apiFetch(`/outlook/messages/${encodeURIComponent(messageId)}/read`, { method: "POST" });
}

async function triggerAutoBidFromSender(sender) {
  const tabId = sender?.tab?.id;
  if (!Number.isInteger(tabId)) return triggerAutoBidInActiveTab();
  return triggerAutoBidInTab(tabId, "hotkey-listener");
}

async function triggerAutoBidInTab(tabId, source) {
  const now = Date.now();
  const lastTriggeredAt = autoBidTriggerTimes.get(tabId) || 0;
  if (now - lastTriggeredAt < AUTO_BID_TRIGGER_DEDUP_MS) {
    return { triggered: false, deduped: true, source };
  }

  autoBidTriggerTimes.set(tabId, now);
  console.info("[AutoBid] trigger", { tabId, source });
  chrome.storage.local.set({ [GPT_BATCH_PAUSED_STORAGE_KEY]: false }).catch(() => {});
  ensureRuntimeGptBatchWorker({
    force: true,
    prewarm: true,
    allowSheetFallback: false,
    source: "first-autofill-hotkey"
  }).catch((error) => {
    console.warn("[AutoBid GPT Pool] First-hotkey prewarm failed", {
      tab_id: tabId,
      error: error.message || String(error)
    });
  });
  const backgroundHold = holdBackgroundTabAutomation(tabId, "autofill-bootstrap").catch((error) => {
    console.warn("[AutoBid Background] Could not keep the application tab active", {
      tab_id: tabId,
      error: error.message || String(error)
    });
  });
  try {
    await injectAutoBidScripts(tabId);
    await backgroundHold;
  } catch (error) {
    await releaseBackgroundTabAutomation(tabId, "autofill-bootstrap").catch(() => {});
    throw error;
  }
  return { triggered: true, gptWorker: { scheduled: true, reason: "first-hotkey-prewarm" } };
}

async function updateAutofillTabState(payload = {}, sender = {}) {
  const tabId = Number.isInteger(sender?.tab?.id) ? sender.tab.id : null;
  if (!Number.isInteger(tabId)) return { updated: false, reason: "missing-tab" };
  const frameId = Number.isInteger(sender?.frameId) ? sender.frameId : 0;
  const runId = String(payload.run_id || payload.runId || "legacy");
  const holdReason = `autofill-run:${frameId}:${runId}`;
  if (payload.running !== false) {
    await holdBackgroundTabAutomation(tabId, holdReason);
    await releaseBackgroundTabAutomation(tabId, "autofill-bootstrap").catch(() => {});
    return { updated: true, running: true, tabId, frameId };
  }
  await releaseBackgroundTabAutomation(tabId, holdReason);
  return { updated: true, running: false, tabId, frameId };
}

function queueNativeClick(sender, payload) {
  const tabId = sender?.tab?.id;
  if (!Number.isInteger(tabId)) throw new Error("Native click requires an active browser tab");
  if ((sender.frameId || 0) !== 0) throw new Error("Native dropdown clicks are currently supported in the top page only");

  return queueNativeInput(tabId, () => dispatchNativeClick(tabId, payload));
}

function queueNativeType(sender, payload) {
  const tabId = sender?.tab?.id;
  if (!Number.isInteger(tabId)) throw new Error("Native typing requires an active browser tab");
  if ((sender.frameId || 0) !== 0) throw new Error("Native typing is currently supported in the top page only");

  return queueNativeInput(tabId, () => dispatchNativeType(tabId, payload));
}

function queueNativeInput(tabId, operation) {
  const previous = nativeInputQueues.get(tabId) || Promise.resolve();
  const next = previous.catch(() => {}).then(operation);
  nativeInputQueues.set(tabId, next);
  const cleanup = () => {
    if (nativeInputQueues.get(tabId) === next) nativeInputQueues.delete(tabId);
  };
  next.then(cleanup, cleanup);
  return next;
}

async function dispatchNativeClick(tabId, payload) {
  const x = Number(payload.x);
  const y = Number(payload.y);
  if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || y < 0) {
    throw new Error("Invalid page click coordinates");
  }
  const [execution] = await chrome.scripting.executeScript({
    target: { tabId },
    world: "MAIN",
    func: clickPagePoint,
    args: [{ x, y }]
  });
  if (!execution?.result?.clicked) throw new Error(execution?.result?.reason || "Page click did not reach a control");
  return execution.result;
}

function clickPagePoint(payload) {
  const target = document.elementFromPoint(Number(payload.x), Number(payload.y));
  const control = target?.closest?.("button, [role='button'], [role='option'], [role='radio'], [role='checkbox'], input, label, a") || target;
  if (!control) return { clicked: false, reason: "No control was found at the requested position" };
  control.focus?.({ preventScroll: true });
  control.click?.();
  return { clicked: true, method: "page-click" };
}

async function dispatchNativeType(tabId, payload) {
  const text = String(payload.text || "").slice(0, 200);
  if (!text) throw new Error("Page typing requires text");
  const [execution] = await chrome.scripting.executeScript({
    target: { tabId },
    world: "MAIN",
    func: typeIntoFocusedPageControl,
    args: [{ text, commit: payload.commit !== false }]
  });
  if (!execution?.result?.typed) throw new Error(execution?.result?.reason || "Page typing did not reach a field");
  return execution.result;
}

function typeIntoFocusedPageControl(payload) {
  let control = document.activeElement;
  while (control?.shadowRoot?.activeElement) control = control.shadowRoot.activeElement;
  if (!control || !control.matches?.("input, textarea, [contenteditable='true'], [role='combobox']")) {
    return { typed: false, reason: "No editable control is focused" };
  }

  const text = String(payload.text || "");
  if (control.matches("input, textarea")) {
    const setter = Object.getOwnPropertyDescriptor(control.constructor.prototype, "value")?.set;
    if (setter) setter.call(control, text);
    else control.value = text;
  } else {
    control.textContent = text;
  }
  control.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText", data: text }));
  if (payload.commit) {
    control.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, composed: true, key: "Enter", code: "Enter" }));
    control.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, composed: true, key: "Enter", code: "Enter" }));
    control.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
  }
  return { typed: true, method: "page-input" };
}

function nativeFileUpload(sender, payload) {
  void sender;
  void payload;
  throw new Error("Debugger-free mode uses the page file-input bridge instead");
}

function nativeFileChooserUpload(sender, payload) {
  void sender;
  void payload;
  throw new Error("Debugger-free mode cannot intercept a browser file chooser");
}

async function holdBackgroundTabAutomation(tabId, reason) {
  if (!Number.isInteger(tabId)) return { held: false, reason: "missing-tab" };
  const holdReason = String(reason || "background-automation");
  const reasons = backgroundAutomationHolds.get(tabId) || new Set();
  reasons.add(holdReason);
  backgroundAutomationHolds.set(tabId, reasons);

  await chrome.tabs.update(tabId, { autoDiscardable: false }).catch(() => null);
  console.info("[AutoBid Background] debugger-free tab protection enabled", {
    tab_id: tabId,
    reason: holdReason,
    holds: reasons.size,
    auto_discardable: false
  });
  return { held: true, tabId, reason: holdReason, debuggerFree: true };
}

async function releaseBackgroundTabAutomation(tabId, reason) {
  if (!Number.isInteger(tabId)) return { released: false, reason: "missing-tab" };
  const releasedReason = String(reason || "background-automation");
  const reasons = backgroundAutomationHolds.get(tabId);
  if (!reasons) return { released: false, reason: "not-held" };
  reasons.delete(releasedReason);
  if (reasons.size > 0) {
    return { released: true, remaining: reasons.size };
  }

  backgroundAutomationHolds.delete(tabId);
  if (releasedReason !== "gpt-worker") {
    await chrome.tabs.update(tabId, { autoDiscardable: true }).catch(() => null);
  }
  console.info("[AutoBid Background] debugger-free tab protection released", { tab_id: tabId });
  return { released: true, remaining: 0 };
}

async function getSettings() {
  const settings = await chrome.storage.local.get(["apiBase", "token", "user", "selectedProfileId", "devProfile", "sheetSettings"]);
  return {
    apiBase: settings.apiBase || DEFAULT_API_BASE,
    token: settings.token || null,
    user: DEV_AUTH_BYPASS ? (settings.user || DEV_USER) : (settings.user || null),
    selectedProfileId: DEV_AUTH_BYPASS ? (settings.selectedProfileId || DEV_PROFILE_ID) : (settings.selectedProfileId || null),
    devAuthBypass: DEV_AUTH_BYPASS,
    sheetSettings: normalizeSheetSettings(settings.sheetSettings || {})
  };
}

async function saveSettings(payload) {
  const apiBase = String(payload.apiBase || DEFAULT_API_BASE).replace(/\/+$/, "");
  await chrome.storage.local.set({ apiBase });
  if (DEV_AUTH_BYPASS) await ensureDevSession();
  return getSettings();
}

async function saveSheetSettings(payload) {
  const sheetSettings = normalizeSheetSettings(payload);
  await chrome.storage.local.set({ sheetSettings });
  return getSettings();
}

async function openSheetRows(payload) {
  const sheetSettings = normalizeSheetSettings(payload.sheetSettings || payload);
  if (!sheetSettings.sheetName) throw new Error("Sheet tab name is required");
  if (!sheetSettings.startRow || !sheetSettings.endRow) throw new Error("Start and end rows are required");
  await chrome.storage.local.set({ sheetSettings });

  const data = await apiFetch("/sheets/jobs", {
    method: "POST",
    body: sheetSettings
  });
  const jobs = normalizeSheetJobsResponse(data);
  const sheetJobs = {};
  const sheetTabJobs = {};
  const openedJobs = [];
  const skippedJobs = [];
  const failedJobs = [];

  for (const job of jobs) {
    const url = normalizeJobOpenUrl(job.url || findJobUrlInSheetJob(job));
    if (!url) {
      skippedJobs.push({ rowNumber: job.rowNumber, reason: "missing-or-invalid-url" });
      continue;
    }

    const context = {
      ...sheetSettings,
      rowNumber: job.rowNumber,
      url,
      values: job.values || {},
      raw: job.raw || []
    };
    addSheetJobContext(sheetJobs, context);

    try {
      const tab = await chrome.tabs.create({ url, active: false });
      if (Number.isInteger(tab?.id)) {
        openedJobs.push({ ...job, url, tabId: tab.id });
        sheetTabJobs[String(tab.id)] = {
          ...context,
          openedTabId: tab.id,
          openedAt: new Date().toISOString()
        };
      } else {
        failedJobs.push({ rowNumber: job.rowNumber, url, error: "Chrome did not return a tab id" });
      }
    } catch (error) {
      failedJobs.push({ rowNumber: job.rowNumber, url, error: error.message || String(error) });
    }
  }

  await chrome.storage.local.set({ sheetJobs, sheetTabJobs });
  return {
    opened: openedJobs.length,
    total: jobs.length,
    skipped: skippedJobs.length,
    failed: failedJobs.length,
    jobs: openedJobs,
    skippedJobs,
    failedJobs
  };
}

async function createRuntimeGptAnswerRequest(payload = {}, sender = {}) {
  await ensureRuntimeGptRequestsLoaded();
  await chrome.storage.local.set({ [GPT_BATCH_PAUSED_STORAGE_KEY]: false });
  cleanupRuntimeGptRequests();
  const fields = Array.isArray(payload.payload?.fields) ? payload.payload.fields : [];
  if (fields.length === 0) throw new Error("Runtime GPT request requires at least one field.");

  const requestId = createRuntimeGptRequestId();
  const now = Date.now();
  const requestedTimeoutMs = Math.max(30000, Number(payload.timeout_ms || 0));
  const maxAttempts = Math.max(1, Number(payload.max_attempts || 2));
  const context = payload.context || {};
  const page = payload.page || {};
  const request = {
    id: requestId,
    request_id: requestId,
    queue_sequence: nextRuntimeGptQueueSequence++,
    status: "pending",
    created_at: now,
    updated_at: now,
    expires_at: now + Math.max(RUNTIME_GPT_REQUEST_TTL_MS, requestedTimeoutMs + 30000),
    lease_expires_at: 0,
    available_at: now,
    attempt_count: 0,
    max_attempts: maxAttempts,
    terminal_error: false,
    batch_id: "",
    worker_tab_id: null,
    tab_id: Number.isInteger(sender?.tab?.id) ? sender.tab.id : null,
    frame_id: Number.isInteger(sender?.frameId) ? sender.frameId : 0,
    client_run_id: String(payload.client_run_id || payload.clientRunId || ""),
    context,
    page,
    payload: payload.payload,
    answers: [],
    error: "",
    delivered_at: 0
  };

  runtimeGptRequests.set(requestId, request);
  await persistRuntimeGptRequests();
  if (Number.isInteger(request.tab_id)) {
    await holdBackgroundTabAutomation(request.tab_id, `gpt-request:${requestId}`).catch((error) => {
      console.warn("[AutoBid Background] Could not hold the originating application tab", {
        request_id: requestId,
        tab_id: request.tab_id,
        error: error.message || String(error)
      });
    });
  }
  console.info("[AutoBid Parallel] queued durably", {
    request_id: requestId,
    sequence: request.queue_sequence,
    position: getRuntimeGptQueuePosition(requestId),
    tab_id: request.tab_id,
    frame_id: request.frame_id
  });
  ensureRuntimeGptBatchWorker().catch(async (error) => {
    const workerStartError = error.message || String(error);
    request.error = workerStartError;
    request.updated_at = Date.now();
    await persistRuntimeGptRequests().catch(() => {});
    console.warn("[AutoBid] GPT request remains durable after worker startup failure", {
      request_id: requestId,
      error: workerStartError
    });
  });

  return {
    request_id: requestId,
    status: request.status,
    queued: true,
    queue_position: getRuntimeGptQueuePosition(requestId),
    fields: fields.length,
    worker_start_scheduled: true
  };
}

async function getRuntimeGptAnswerStatus(payload = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const cleaned = cleanupRuntimeGptRequests();
  if (cleaned) await persistRuntimeGptRequests();
  const requestId = String(payload.request_id || payload.requestId || "");
  const request = runtimeGptRequests.get(requestId);
  if (!request) return { request_id: requestId, status: "missing", error: "Runtime GPT request not found." };
  if (request.status === "pending" && Number(request.available_at || 0) <= Date.now()) {
    ensureRuntimeGptBatchWorker().catch((error) => {
      console.warn("[AutoBid Batch] Could not restart the queued GPT batch", error);
    });
  }
  return publicRuntimeGptRequestStatus(request);
}

async function cancelRuntimeGptAnswerRequest(payload = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const requestId = String(payload.request_id || payload.requestId || "");
  const request = runtimeGptRequests.get(requestId);
  if (!request) return { request_id: requestId, status: "missing" };
  if (request.status !== "complete") {
    request.status = "cancelled";
    request.error = String(payload.reason || "cancelled");
    request.updated_at = Date.now();
  }
  await persistRuntimeGptRequests();
  await clearRuntimeGptDeliveryRetry(requestId);
  await releaseBackgroundTabAutomation(request.tab_id, `gpt-request:${requestId}`).catch(() => {});
  return publicRuntimeGptRequestStatus(request);
}

async function fetchPendingRuntimeGptRequests(payload = {}, sender = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const cleaned = cleanupRuntimeGptRequests();
  const now = Date.now();
  const queue = getOrderedRuntimeGptQueue();
  const head = queue[0] || null;
  const runnableQueue = queue.filter((request) => request.status === "pending" && Number(request.available_at || 0) <= now);
  const workerTabId = Number.isInteger(sender?.tab?.id) ? sender.tab.id : null;
  const requestedBatchId = String(payload.batch_id || payload.batchId || "");
  const batchId = requestedBatchId || `manual_${workerTabId || Date.now()}`;
  const limit = Math.min(
    RUNTIME_GPT_PROMPT_BATCH_SIZE,
    Math.max(1, Number(payload.limit || RUNTIME_GPT_PROMPT_BATCH_SIZE))
  );
  const batchStates = await getStoredRuntimeGptBatchStates();
  const activeState = batchStates.find((state) => (
    requestedBatchId && state.batch_id === requestedBatchId
  ) || (
    !requestedBatchId && Number.isInteger(workerTabId) && state.tab_id === workerTabId
  ));

  if (requestedBatchId && !activeState) {
    throw new Error("This ChatGPT tab does not own the active AutoBid batch.");
  }
  if (Number.isInteger(activeState?.tab_id) && Number.isInteger(workerTabId) && activeState.tab_id !== workerTabId) {
    throw new Error("This ChatGPT tab does not own the active AutoBid batch.");
  }

  let requests = queue.filter((request) => (
    request.status === "processing" &&
    request.batch_id === batchId &&
    (!workerTabId || request.worker_tab_id === workerTabId)
  )).slice(0, limit);

  if (requests.length === 0 && runnableQueue.length > 0) {
    requests = [];
    for (const request of runnableQueue) {
      if (requests.length >= limit) break;
      request.status = "processing";
      request.batch_id = batchId;
      request.worker_tab_id = workerTabId;
      request.updated_at = now;
      request.lease_expires_at = now + RUNTIME_GPT_LEASE_MS;
      requests.push(request);
    }

    if (requests.length > 0) {
      console.info("[AutoBid Batch] claimed", {
        batch_id: batchId,
        request_ids: requests.map((request) => request.id),
        sequences: requests.map((request) => request.queue_sequence),
        queue_length: queue.length
      });
    }
  }

  if (activeState?.batch_id === batchId) {
    await upsertStoredRuntimeGptBatchState({
      ...activeState,
      status: requests.length ? "processing" : activeState.status,
      request_ids: requests.map((request) => request.id),
      updated_at: Date.now()
    });
  }
  if (cleaned || requests.length > 0) await persistRuntimeGptRequests();

  return {
    batch_id: batchId,
    requests: requests.map(runtimeGptRequestForWorker),
    queue_length: queue.length,
    head_status: head?.status || "empty",
    retry_at: head?.available_at || 0
  };
}

async function saveRuntimeGptRequestAnswers(payload = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const requestId = String(payload.request_id || payload.requestId || "");
  const request = runtimeGptRequests.get(requestId);
  if (!request) return { request_id: requestId, status: "missing", saved_answers: 0 };
  if (request.status === "complete") {
    return {
      request_id: requestId,
      status: request.status,
      saved_answers: request.answers?.length || 0,
      answers: request.answers || [],
      idempotent: true
    };
  }
  if (request.status === "cancelled") {
    return { request_id: requestId, status: request.status, saved_answers: 0 };
  }
  const batchId = String(payload.batch_id || payload.batchId || "");
  if (batchId && request.batch_id !== batchId) {
    throw new Error(`Runtime GPT request ${requestId} belongs to a different batch.`);
  }

  const answers = normalizeRuntimeGptAnswers(payload.answers || []);
  if (answers.length === 0) throw new Error(`Runtime GPT request ${requestId} has no usable answers.`);
  request.status = "complete";
  request.answers = answers;
  request.error = "";
  request.updated_at = Date.now();
  request.lease_expires_at = 0;
  request.available_at = 0;
  request.batch_id = "";
  request.worker_tab_id = null;
  request.completed_at = Date.now();
  request.expires_at = Date.now() + RUNTIME_GPT_REQUEST_TTL_MS;

  await persistRuntimeGptRequests();
  console.info("[AutoBid FIFO] completed", {
    request_id: requestId,
    sequence: request.queue_sequence,
    answers: answers.length
  });
  queueRuntimeGptAnswerDelivery(request).catch((error) => {
    console.warn("[AutoBid] Background GPT answer delivery remains pending", {
      request_id: request.id,
      error: error.message || String(error)
    });
  });

  return {
    request_id: requestId,
    status: request.status,
    saved_answers: answers.length,
    answers,
    delivered: false,
    delivery_scheduled: true
  };
}

function queueRuntimeGptAnswerDelivery(request) {
  if (!request?.id) return Promise.resolve(false);
  const existing = runtimeGptDeliveryPromises.get(request.id);
  if (existing) return existing;

  const delivery = notifyRuntimeGptAnswersReady(request)
    .then((delivered) => {
      if (!delivered && request.status === "complete" && !request.delivered_at && !request.delivery_suppressed) {
        scheduleRuntimeGptDeliveryRetry(request.id);
      }
      return delivered;
    })
    .catch((error) => {
      if (request.status === "complete" && !request.delivered_at && !request.delivery_suppressed) {
        scheduleRuntimeGptDeliveryRetry(request.id);
      }
      throw error;
    })
    .finally(() => {
      if (runtimeGptDeliveryPromises.get(request.id) === delivery) {
        runtimeGptDeliveryPromises.delete(request.id);
      }
    });
  runtimeGptDeliveryPromises.set(request.id, delivery);
  return delivery;
}

function scheduleRuntimeGptDeliveryRetry(requestId) {
  const alarmName = `${RUNTIME_GPT_DELIVERY_ALARM_PREFIX}${requestId}`;
  if (chrome.alarms?.create) {
    chrome.alarms.create(alarmName, { when: Date.now() + RUNTIME_GPT_RETRY_DELAY_MS });
  }
}

function clearRuntimeGptDeliveryRetry(requestId) {
  if (!requestId || !chrome.alarms?.clear) return Promise.resolve(false);
  return chrome.alarms.clear(`${RUNTIME_GPT_DELIVERY_ALARM_PREFIX}${requestId}`).catch(() => false);
}

async function notifyRuntimeGptAnswersReady(request) {
  if (!Number.isInteger(request?.tab_id)) return false;
  const holdReason = `gpt-request:${request.id}`;
  try {
    await holdBackgroundTabAutomation(request.tab_id, holdReason);
    const tab = await chrome.tabs.get(request.tab_id);
    const requestedUrl = String(request.page?.url || request.payload?.page?.url || "");
    if (requestedUrl && scoreSheetContextUrlMatch(normalizeUrlForMatch(requestedUrl), normalizeUrlForMatch(tab.url || "")) < 600) {
      console.warn("[AutoBid] Deferred GPT answer delivery because the originating tab navigated away", {
        request_id: request.id,
        requested_url: requestedUrl,
        current_url: tab.url || ""
      });
      request.delivery_suppressed = true;
      request.updated_at = Date.now();
      await persistRuntimeGptRequests();
      await clearRuntimeGptDeliveryRetry(request.id);
      await releaseBackgroundTabAutomation(request.tab_id, holdReason).catch(() => {});
      return false;
    }
    let response = null;
    for (let attempt = 1; attempt <= RUNTIME_GPT_DELIVERY_ATTEMPTS; attempt += 1) {
      response = await chrome.tabs.sendMessage(request.tab_id, {
        type: "AUTO_BID_GPT_ANSWERS_READY",
        payload: {
          request_id: request.id,
          answers: request.answers || [],
          fields: Array.isArray(request.payload?.fields) ? request.payload.fields : [],
          page_url: request.page?.url || request.payload?.page?.url || "",
          client_run_id: request.client_run_id || "",
          delivery_attempt: attempt
        }
      }, { frameId: Number.isInteger(request.frame_id) ? request.frame_id : 0 });
      if (response?.settled) {
        await acknowledgeRuntimeGptAnswersApplied({
          request_id: request.id,
          result: {
            filled: Number(response.filled || 0),
            missed: Number(response.missed || 0),
            exhausted: Number(response.exhausted || 0),
            background_delivery: true
          }
        }, {
          tab: { id: request.tab_id },
          frameId: Number.isInteger(request.frame_id) ? request.frame_id : 0
        });
        return true;
      }
      if (attempt < RUNTIME_GPT_DELIVERY_ATTEMPTS) await sleep(RUNTIME_GPT_DELIVERY_RETRY_MS);
    }
    return false;
  } catch (error) {
    console.warn("[AutoBid] Could not push GPT answers to originating tab", {
      request_id: request.id,
      tab_id: request.tab_id,
      frame_id: request.frame_id,
      error: error.message || String(error)
    });
    return false;
  }
}

async function retryCompletedRuntimeGptDeliveriesForTab(tabId) {
  if (!Number.isInteger(tabId)) return { attempted: 0 };
  await ensureRuntimeGptRequestsLoaded();
  const requests = Array.from(runtimeGptRequests.values())
    .filter((request) => request.status === "complete" && !request.delivered_at && request.tab_id === tabId)
    .sort((left, right) => Number(left.queue_sequence || 0) - Number(right.queue_sequence || 0));

  for (const request of requests) {
    await queueRuntimeGptAnswerDelivery(request);
  }
  return { attempted: requests.length };
}

async function acknowledgeRuntimeGptAnswersApplied(payload = {}, sender = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const requestId = String(payload.request_id || payload.requestId || "");
  const request = runtimeGptRequests.get(requestId);
  if (!request) return { request_id: requestId, status: "missing" };
  if (Number.isInteger(sender?.tab?.id) && Number.isInteger(request.tab_id) && sender.tab.id !== request.tab_id) {
    return { request_id: requestId, status: request.status, acknowledged: false, reason: "wrong-tab" };
  }
  if (Number.isInteger(sender?.frameId) && Number.isInteger(request.frame_id) && sender.frameId !== request.frame_id) {
    return { request_id: requestId, status: request.status, acknowledged: false, reason: "wrong-frame" };
  }

  request.delivered_at = Date.now();
  request.delivery_result = payload.result || null;
  request.updated_at = Date.now();
  await persistRuntimeGptRequests();
  await clearRuntimeGptDeliveryRetry(requestId);
  await releaseBackgroundTabAutomation(request.tab_id, `gpt-request:${requestId}`).catch(() => {});
  return { request_id: requestId, status: request.status, acknowledged: true };
}

async function failRuntimeGptRequest(payload = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const requestId = String(payload.request_id || payload.requestId || "");
  const request = runtimeGptRequests.get(requestId);
  if (!request) return { request_id: requestId, status: "missing" };
  if (request.status === "complete" || request.status === "cancelled") {
    return publicRuntimeGptRequestStatus(request);
  }
  const batchId = String(payload.batch_id || payload.batchId || "");
  if (batchId && request.batch_id !== batchId) {
    return { ...publicRuntimeGptRequestStatus(request), ignored: true, reason: "stale-batch" };
  }
  failOrRequeueRuntimeGptRequest(
    request,
    payload.error || "Runtime GPT request failed.",
    RUNTIME_GPT_RETRY_DELAY_MS
  );
  await persistRuntimeGptRequests();
  if (request.terminal_error) {
    await clearRuntimeGptDeliveryRetry(requestId);
    await releaseBackgroundTabAutomation(request.tab_id, `gpt-request:${requestId}`).catch(() => {});
  }
  console.warn(request.terminal_error
    ? "[AutoBid Parallel] request exhausted ChatGPT attempts"
    : "[AutoBid Parallel] request returned to durable retry queue", {
    request_id: requestId,
    sequence: request.queue_sequence,
    attempt: request.attempt_count,
    retry_at: request.available_at,
    terminal: request.terminal_error,
    error: request.error
  });
  return publicRuntimeGptRequestStatus(request);
}

function failOrRequeueRuntimeGptRequest(request, reason, retryDelayMs = RUNTIME_GPT_RETRY_DELAY_MS) {
  request.error = String(reason || "Runtime GPT request failed.");
  request.updated_at = Date.now();
  request.lease_expires_at = 0;
  request.batch_id = "";
  request.worker_tab_id = null;
  request.attempt_count = Number(request.attempt_count || 0) + 1;
  request.terminal_error = request.attempt_count >= Math.max(1, Number(request.max_attempts || 2));
  request.status = request.terminal_error ? "error" : "pending";
  request.available_at = request.terminal_error ? 0 : Date.now() + Math.max(0, Number(retryDelayMs || 0));
  return request;
}

function runtimeGptRequestForWorker(request) {
  const context = request.context || {};
  const payload = request.payload || {};
  return {
    requestId: request.id,
    request_id: request.id,
    queueSequence: request.queue_sequence,
    queue_sequence: request.queue_sequence,
    attemptCount: request.attempt_count || 0,
    attempt_count: request.attempt_count || 0,
    maxAttempts: request.max_attempts || 2,
    max_attempts: request.max_attempts || 2,
    runtime: true,
    rowNumber: context.rowNumber || null,
    row: context.rowNumber || null,
    sheetName: context.sheetName || "",
    spreadsheetId: context.spreadsheetId || "",
    url: context.url || request.page?.url || payload.page?.url || "",
    values: context.values || payload.row?.values || {},
    raw: Array.isArray(context.raw) ? context.raw : [],
    page: request.page || payload.page || {},
    questions: payload,
    created_at: new Date(request.created_at).toISOString()
  };
}

function publicRuntimeGptRequestStatus(request) {
  return {
    request_id: request.id,
    status: request.status,
    queue_sequence: request.queue_sequence,
    queue_position: getRuntimeGptQueuePosition(request.id),
    attempt_count: request.attempt_count || 0,
    max_attempts: request.max_attempts || 2,
    terminal_error: Boolean(request.terminal_error),
    answers: request.answers || [],
    error: request.error || "",
    delivered: Boolean(request.delivered_at),
    created_at: new Date(request.created_at).toISOString(),
    updated_at: new Date(request.updated_at).toISOString()
  };
}

function normalizeRuntimeGptAnswers(answers) {
  return (Array.isArray(answers) ? answers : [])
    .map((answer) => ({
      field_id: String(answer.field_id || answer.id || ""),
      question: String(answer.question || ""),
      option: String(answer.option || ""),
      value: String(answer.value ?? answer.answer ?? "").trim(),
      source: "runtime-gpt"
    }))
    .filter((answer) => answer.field_id && answer.value && !isRejectedRuntimeAnswerPlaceholder(answer.value));
}

function isRejectedRuntimeAnswerPlaceholder(value) {
  return /^(?:not specified|unspecified|unknown|not provided|not available|no information(?: provided| available)?|information unavailable|to be determined|tbd)$/i
    .test(String(value || "").trim());
}

function getOrderedRuntimeGptQueue() {
  return Array.from(runtimeGptRequests.values())
    .filter((request) => request.status === "pending" || request.status === "processing")
    .sort((left, right) => {
      const leftSequence = Number(left.queue_sequence || Number.MAX_SAFE_INTEGER);
      const rightSequence = Number(right.queue_sequence || Number.MAX_SAFE_INTEGER);
      return leftSequence - rightSequence || Number(left.created_at || 0) - Number(right.created_at || 0) || String(left.id).localeCompare(String(right.id));
    });
}

function getRuntimeGptQueuePosition(requestId) {
  const index = getOrderedRuntimeGptQueue().findIndex((request) => request.id === requestId);
  return index >= 0 ? index + 1 : 0;
}

function cleanupRuntimeGptRequests() {
  const now = Date.now();
  let changed = false;
  for (const [requestId, request] of runtimeGptRequests.entries()) {
    if (request.status === "cancelled" || (request.status === "complete" && request.delivered_at && Number(request.expires_at || 0) <= now)) {
      runtimeGptRequests.delete(requestId);
      clearRuntimeGptDeliveryRetry(requestId).catch(() => {});
      releaseBackgroundTabAutomation(request.tab_id, `gpt-request:${requestId}`).catch(() => {});
      changed = true;
      continue;
    }
    if (request.status === "error" && !request.terminal_error) {
      request.status = "pending";
      request.available_at = now;
      request.lease_expires_at = 0;
      request.batch_id = "";
      request.worker_tab_id = null;
      changed = true;
      continue;
    }
    if (request.status === "processing" && Number(request.lease_expires_at || 0) <= now) {
      request.status = "pending";
      request.available_at = Math.min(Number(request.available_at || now), now);
      request.lease_expires_at = 0;
      request.batch_id = "";
      request.worker_tab_id = null;
      changed = true;
    }
  }
  return changed;
}

async function ensureRuntimeGptRequestsLoaded() {
  if (runtimeGptRequestsLoaded) return;
  if (runtimeGptRequestsLoadPromise) return runtimeGptRequestsLoadPromise;

  runtimeGptRequestsLoadPromise = (async () => {
    const stored = await chrome.storage.local.get([RUNTIME_GPT_QUEUE_STORAGE_KEY]).catch(() => ({}));
    const queueState = stored[RUNTIME_GPT_QUEUE_STORAGE_KEY] || {};
    const requests = Array.isArray(queueState.requests) ? queueState.requests : [];
    let maxSequence = 0;

    requests.forEach((rawRequest, index) => {
      const id = String(rawRequest?.id || rawRequest?.request_id || "");
      if (!id || rawRequest?.status === "cancelled") return;
      const queueSequence = Number(rawRequest.queue_sequence || index + 1);
      maxSequence = Math.max(maxSequence, queueSequence);
      runtimeGptRequests.set(id, {
        ...rawRequest,
        id,
        request_id: id,
        queue_sequence: queueSequence,
        status: rawRequest.status === "processing" && !rawRequest.batch_id ? "pending" : (rawRequest.status || "pending"),
        created_at: Number(rawRequest.created_at || Date.now()),
        updated_at: Number(rawRequest.updated_at || rawRequest.created_at || Date.now()),
        lease_expires_at: Number(rawRequest.lease_expires_at || 0),
        available_at: Number(rawRequest.available_at || 0),
        attempt_count: Number(rawRequest.attempt_count || 0),
        max_attempts: Math.max(1, Number(rawRequest.max_attempts || 2)),
        terminal_error: Boolean(rawRequest.terminal_error),
        batch_id: String(rawRequest.batch_id || ""),
        worker_tab_id: Number.isInteger(rawRequest.worker_tab_id) ? rawRequest.worker_tab_id : null,
        answers: Array.isArray(rawRequest.answers) ? rawRequest.answers : []
      });
    });

    nextRuntimeGptQueueSequence = Math.max(
      Number(queueState.next_sequence || 1),
      maxSequence + 1
    );
    runtimeGptRequestsLoaded = true;
    if (cleanupRuntimeGptRequests()) await persistRuntimeGptRequests();
  })().finally(() => {
    runtimeGptRequestsLoadPromise = null;
  });

  return runtimeGptRequestsLoadPromise;
}

function persistRuntimeGptRequests() {
  const snapshot = {
    version: 2,
    next_sequence: nextRuntimeGptQueueSequence,
    requests: Array.from(runtimeGptRequests.values())
  };
  runtimeGptPersistPromise = runtimeGptPersistPromise
    .catch(() => {})
    .then(() => chrome.storage.local.set({ [RUNTIME_GPT_QUEUE_STORAGE_KEY]: snapshot }));
  return runtimeGptPersistPromise;
}

function createRuntimeGptRequestId() {
  if (globalThis.crypto?.randomUUID) return `abg_${globalThis.crypto.randomUUID()}`;
  return `abg_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

async function runGptAnswerWorker(payload = {}) {
  await chrome.storage.local.set({ [GPT_BATCH_PAUSED_STORAGE_KEY]: false });
  return ensureRuntimeGptBatchWorker({
    force: true,
    allowSheetFallback: true,
    source: payload.mode || "manual"
  });
}

async function stopGptAnswerWorker() {
  await chrome.storage.local.set({ [GPT_BATCH_PAUSED_STORAGE_KEY]: true });
  const states = await getStoredRuntimeGptBatchStates();
  if (states.length === 0) {
    await cleanupLegacyGptWorkerTab();
    await chrome.storage.local.set({ autoBidGptWorkerState: { running: false, stoppedAt: new Date().toISOString() } });
    return { ok: true, stopped: false, stoppedWorkers: 0, reason: "no-worker-tab" };
  }

  const responses = await Promise.all(states.map(async (state) => {
    let response = null;
    try {
      response = await chrome.tabs.sendMessage(state.tab_id, { type: "AUTOBID_GPT_STOP" });
    } catch (_error) {
      response = null;
    }
    await releaseRuntimeGptBatchRequests(state.batch_id, "GPT batch stopped by the user.", RUNTIME_GPT_RETRY_DELAY_MS);
    await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
    if (Number.isInteger(state.tab_id)) await chrome.tabs.remove(state.tab_id).catch(() => {});
    return { batchId: state.batch_id, tabId: state.tab_id, response };
  }));
  await replaceStoredRuntimeGptBatchStates([]);
  await chrome.storage.local.set({ autoBidGptWorkerState: { running: false, stoppedAt: new Date().toISOString() } });
  return { ok: true, stopped: true, stoppedWorkers: states.length, workers: responses };
}

function ensureRuntimeGptBatchWorker(options = {}) {
  gptBatchPumpRequested = true;
  if (gptBatchPumpPromise) return gptBatchPumpPromise;

  gptBatchPumpPromise = (async () => {
    let result = { ok: true, scheduled: false, reason: "queue-empty" };
    do {
      gptBatchPumpRequested = false;
      result = await ensureRuntimeGptBatchWorkerInternal(options);
    } while (gptBatchPumpRequested);
    return result;
  })().finally(() => {
    gptBatchPumpPromise = null;
  });
  return gptBatchPumpPromise;
}

async function ensureRuntimeGptBatchWorkerInternal(options = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const control = await chrome.storage.local.get([GPT_BATCH_PAUSED_STORAGE_KEY]).catch(() => ({}));
  if (control[GPT_BATCH_PAUSED_STORAGE_KEY] && !options.force) {
    return { ok: true, scheduled: false, reason: "paused" };
  }
  const cleaned = cleanupRuntimeGptRequests();
  if (cleaned) await persistRuntimeGptRequests();

  await cleanupLegacyGptWorkerTab();
  const trimmedWorkers = await enforceRuntimeGptWorkerLimit();
  const recoveryResults = await recoverRuntimeGptBatchWorkers();
  const activeStates = await getStoredRuntimeGptBatchStates();
  const runnableRequests = getOrderedRuntimeGptQueue()
    .filter((request) => request.status === "pending" && Number(request.available_at || 0) <= Date.now());
  const idleStates = activeStates.filter(isRuntimeGptWorkerIdle);
  const workerJobs = [];
  let requestIndex = 0;

  for (const state of idleStates) {
    if (requestIndex >= runnableRequests.length) break;
    const requests = runnableRequests.slice(
      requestIndex,
      requestIndex + RUNTIME_GPT_PROMPT_BATCH_SIZE
    );
    requestIndex += requests.length;
    workerJobs.push(() => assignRuntimeGptRequestsToWorker(state, requests, options));
  }

  let newWorkerCount = 0;
  while (
    requestIndex < runnableRequests.length &&
    activeStates.length + newWorkerCount < RUNTIME_GPT_MAX_WORKERS
  ) {
    const requests = runnableRequests.slice(
      requestIndex,
      requestIndex + RUNTIME_GPT_PROMPT_BATCH_SIZE
    );
    requestIndex += requests.length;
    newWorkerCount += 1;
    workerJobs.push(() => startRuntimeGptBatchWorker(requests, options));
  }

  if (workerJobs.length === 0 && options.force && activeStates.length === 0) {
    workerJobs.push(() => startRuntimeGptBatchWorker([], options));
  }

  if (workerJobs.length === 0) {
    const nextRetryAt = getOrderedRuntimeGptQueue()
      .filter((request) => request.status === "pending")
      .reduce((earliest, request) => Math.min(earliest, Number(request.available_at || 0) || Infinity), Infinity);
    return {
      ok: true,
      scheduled: false,
      activeWorkers: activeStates.length,
      recoveredWorkers: recoveryResults.recovered,
      trimmedWorkers,
      reason: activeStates.length ? "all-runnable-requests-owned" : (Number.isFinite(nextRetryAt) ? "requests-waiting-for-retry" : "queue-empty"),
      retryAt: Number.isFinite(nextRetryAt) ? nextRetryAt : 0
    };
  }

  const workerResults = await Promise.allSettled(workerJobs.map((run) => run()));
  const started = workerResults
    .filter((result) => result.status === "fulfilled")
    .map((result) => result.value);
  const failures = workerResults
    .filter((result) => result.status === "rejected")
    .map((result) => result.reason?.message || String(result.reason));
  const states = await getStoredRuntimeGptBatchStates();

  return {
    ok: failures.length === 0,
    scheduled: started.length > 0,
    activeWorkers: states.length,
    startedWorkers: started.length,
    recoveredWorkers: recoveryResults.recovered,
    trimmedWorkers,
    batches: started,
    errors: failures
  };
}

function isRuntimeGptWorkerIdle(state) {
  return state?.status === "idle" && (!Array.isArray(state.request_ids) || state.request_ids.length === 0);
}

async function enforceRuntimeGptWorkerLimit() {
  const states = await getStoredRuntimeGptBatchStates();
  if (states.length <= RUNTIME_GPT_MAX_WORKERS) return 0;

  const ordered = [...states].sort((left, right) => (
    Number(left.created_at || 0) - Number(right.created_at || 0)
  ));
  const excess = ordered.slice(RUNTIME_GPT_MAX_WORKERS);
  for (const state of excess) {
    await releaseRuntimeGptBatchRequests(
      state.batch_id,
      "The GPT worker pool was reduced to three persistent tabs.",
      0
    );
    await removeStoredRuntimeGptBatchState(state.batch_id);
    if (Number.isInteger(state.tab_id)) {
      await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
      await chrome.tabs.sendMessage(state.tab_id, { type: "AUTOBID_GPT_STOP" }).catch(() => {});
      await chrome.tabs.remove(state.tab_id).catch(() => {});
    }
  }

  console.info("[AutoBid GPT Pool] removed excess worker tabs", { removed: excess.length });
  return excess.length;
}

async function startRuntimeGptBatchWorker(requests, options = {}) {
  const batchId = createRuntimeGptBatchId();
  const now = Date.now();
  const requestIds = requests.map((request) => request.id);

  for (const request of requests) {
    request.status = "processing";
    request.batch_id = batchId;
    request.worker_tab_id = null;
    request.updated_at = now;
    request.lease_expires_at = now + RUNTIME_GPT_LEASE_MS;
  }
  if (requests.length > 0) await persistRuntimeGptRequests();

  const workerUrl = `${GPT_WORKER_URL}&autobid_batch=${encodeURIComponent(batchId)}`;
  let created = null;
  try {
    created = await chrome.tabs.create({ url: workerUrl, active: false });
    if (!Number.isInteger(created?.id)) throw new Error("Could not open a ChatGPT batch tab.");

    for (const request of requests) request.worker_tab_id = created.id;
    if (requests.length > 0) await persistRuntimeGptRequests();

    const state = {
      version: 3,
      batch_id: batchId,
      tab_id: created.id,
      status: "loading",
      request_ids: requestIds,
      allow_sheet_fallback: Boolean(options.allowSheetFallback),
      source: String(options.source || "runtime"),
      created_at: now,
      updated_at: Date.now()
    };
    await upsertStoredRuntimeGptBatchState(state);
    const response = options.prewarm && requestIds.length === 0
      ? await prewarmRuntimeGptBatchWorker(state)
      : await resumeRuntimeGptBatchWorker(state, { inject: true });
    console.info("[AutoBid Parallel] opened ChatGPT worker", {
      batch_id: batchId,
      tab_id: created.id,
      request_ids: requestIds
    });
    return { batchId, tabId: created.id, requestIds, response };
  } catch (error) {
    await releaseRuntimeGptBatchRequests(batchId, "Could not start the ChatGPT worker.", 0);
    await removeStoredRuntimeGptBatchState(batchId);
    if (Number.isInteger(created?.id)) await releaseBackgroundTabAutomation(created.id, "gpt-worker").catch(() => {});
    if (Number.isInteger(created?.id)) await chrome.tabs.remove(created.id).catch(() => {});
    throw error;
  }
}

async function prewarmRuntimeGptBatchWorker(state) {
  await protectGptWorkerTab(state.tab_id);
  await waitForTabLoaded(state.tab_id);
  await injectGptAnswerWorker(state.tab_id);
  const response = await chrome.tabs.sendMessage(state.tab_id, { type: "AUTOBID_GPT_PING" });
  await upsertStoredRuntimeGptBatchState({
    ...state,
    version: 3,
    status: "idle",
    request_ids: [],
    allow_sheet_fallback: false,
    updated_at: Date.now()
  });
  await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
  const states = await getStoredRuntimeGptBatchStates();
  await chrome.storage.local.set({
    autoBidGptWorkerState: {
      running: states.length > 0,
      mode: "persistent-worker-pool",
      activeWorkers: states.length,
      busyWorkers: states.filter((worker) => !isRuntimeGptWorkerIdle(worker)).length,
      maxParallelWorkers: RUNTIME_GPT_MAX_WORKERS,
      prewarmed: true,
      updatedAt: new Date().toISOString()
    }
  });
  console.info("[AutoBid GPT Pool] prewarmed an idle ChatGPT worker", {
    worker_id: state.batch_id,
    tab_id: state.tab_id
  });
  return { ...response, prewarmed: true };
}

async function assignRuntimeGptRequestsToWorker(state, requests, options = {}) {
  const ownedRequests = requests.filter((request) => request?.status === "pending");
  if (ownedRequests.length === 0) return { reused: true, scheduled: false, tabId: state.tab_id };

  const now = Date.now();
  const requestIds = ownedRequests.map((request) => request.id);
  for (const request of ownedRequests) {
    request.status = "processing";
    request.batch_id = state.batch_id;
    request.worker_tab_id = state.tab_id;
    request.updated_at = now;
    request.lease_expires_at = now + RUNTIME_GPT_LEASE_MS;
  }
  await persistRuntimeGptRequests();

  const assignedState = {
    ...state,
    version: 3,
    status: "loading",
    request_ids: requestIds,
    allow_sheet_fallback: false,
    source: String(options.source || state.source || "runtime"),
    updated_at: now
  };
  await upsertStoredRuntimeGptBatchState(assignedState);

  try {
    const response = await resumeRuntimeGptBatchWorker(assignedState);
    console.info("[AutoBid GPT Pool] reused ChatGPT worker", {
      worker_id: state.batch_id,
      tab_id: state.tab_id,
      request_ids: requestIds
    });
    return {
      batchId: state.batch_id,
      tabId: state.tab_id,
      requestIds,
      reused: true,
      response
    };
  } catch (error) {
    await releaseRuntimeGptBatchRequests(
      state.batch_id,
      "Could not reuse the persistent ChatGPT worker.",
      0
    );
    await removeStoredRuntimeGptBatchState(state.batch_id);
    await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
    if (Number.isInteger(state.tab_id)) await chrome.tabs.remove(state.tab_id).catch(() => {});
    throw error;
  }
}

async function recoverRuntimeGptBatchWorkers() {
  const states = await getStoredRuntimeGptBatchStates();
  const results = await Promise.all(states.map(async (state) => {
    const tab = Number.isInteger(state.tab_id) ? await chrome.tabs.get(state.tab_id).catch(() => null) : null;
    if (!tab || !isGptTabUrl(tab.url || "")) {
      await releaseRuntimeGptBatchRequests(state.batch_id, "Recovered an orphaned GPT batch.", 0);
      await removeStoredRuntimeGptBatchState(state.batch_id);
      return { recovered: false, removed: true };
    }

    await holdBackgroundTabAutomation(state.tab_id, "gpt-worker").catch((error) => {
      console.warn("[AutoBid GPT Pool] Could not restore the background lifecycle hold", {
        tab_id: state.tab_id,
        error: error.message || String(error)
      });
    });

    const requestIds = Array.isArray(state.request_ids) ? state.request_ids : [];
    const hasInFlightRequest = requestIds.some((requestId) => {
      const request = runtimeGptRequests.get(String(requestId));
      return request?.status === "processing" && request.batch_id === state.batch_id;
    });
    const ping = await sendTabMessageWithReinject(state.tab_id, { type: "AUTOBID_GPT_PING" }).catch(() => null);
    if (ping?.runningBatch) return { recovered: false, running: true };

    if (requestIds.length > 0 && !hasInFlightRequest) {
      await completeRuntimeGptBatch({ batch_id: state.batch_id, request_ids: requestIds }, { tab: { id: state.tab_id } });
      return { recovered: false, completed: true };
    }

    if (requestIds.length > 0 && hasInFlightRequest) {
      await resumeRuntimeGptBatchWorker(state);
      return { recovered: true, resumed: true };
    }

    await upsertStoredRuntimeGptBatchState({
      ...state,
      version: 3,
      status: "idle",
      request_ids: [],
      updated_at: Date.now()
    });
    await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
    return { recovered: true, idle: true };
  }));

  return {
    recovered: results.filter((result) => result.recovered).length,
    removed: results.filter((result) => result.removed).length,
    completed: results.filter((result) => result.completed).length
  };
}

async function resumeRuntimeGptBatchWorker(state, options = {}) {
  await protectGptWorkerTab(state.tab_id);
  await waitForTabLoaded(state.tab_id);
  if (options.inject) await injectGptAnswerWorker(state.tab_id);
  const runningState = {
    ...state,
    status: state.request_ids?.length ? "processing" : "running",
    updated_at: Date.now()
  };
  await upsertStoredRuntimeGptBatchState(runningState);
  const response = await sendTabMessageWithReinject(state.tab_id, {
    type: "AUTOBID_GPT_RUN_BATCH",
    batchId: state.batch_id,
    batchSize: Math.max(1, Math.min(
      RUNTIME_GPT_PROMPT_BATCH_SIZE,
      state.request_ids?.length || RUNTIME_GPT_PROMPT_BATCH_SIZE
    )),
    collectionDelayMs: 0,
    allowSheetFallback: Boolean(state.allow_sheet_fallback)
  });
  if (response?.busy || response?.started === false) {
    throw new Error("The persistent ChatGPT worker reported that it is still busy.");
  }
  const states = await getStoredRuntimeGptBatchStates();
  await chrome.storage.local.set({
    autoBidGptWorkerState: {
      running: states.length > 0,
      mode: "persistent-worker-pool",
      activeWorkers: states.length,
      busyWorkers: states.filter((worker) => !isRuntimeGptWorkerIdle(worker)).length,
      maxParallelWorkers: RUNTIME_GPT_MAX_WORKERS,
      updatedAt: new Date().toISOString()
    }
  });
  return response;
}

async function protectGptWorkerTab(tabId) {
  try {
    await holdBackgroundTabAutomation(tabId, "gpt-worker");
  } catch (error) {
    console.warn("[AutoBid] Could not keep the GPT worker active in the background", error);
  }
}

async function getStoredRuntimeGptBatchStates() {
  const stored = await chrome.storage.local.get([
    GPT_BATCH_STATES_STORAGE_KEY,
    LEGACY_GPT_BATCH_STATE_STORAGE_KEY
  ]).catch(() => ({}));
  const states = stored[GPT_BATCH_STATES_STORAGE_KEY];
  if (Array.isArray(states)) return states.filter((state) => state?.batch_id);

  const legacyState = stored[LEGACY_GPT_BATCH_STATE_STORAGE_KEY];
  if (!legacyState?.batch_id) return [];
  const migrated = [{ ...legacyState, version: 2 }];
  await chrome.storage.local.set({ [GPT_BATCH_STATES_STORAGE_KEY]: migrated });
  await chrome.storage.local.remove([LEGACY_GPT_BATCH_STATE_STORAGE_KEY]);
  return migrated;
}

function mutateStoredRuntimeGptBatchStates(mutator) {
  const operation = runtimeGptBatchStateMutationPromise
    .catch(() => {})
    .then(async () => {
      const current = await getStoredRuntimeGptBatchStates();
      const next = await mutator([...current]);
      const normalized = Array.isArray(next) ? next.filter((state) => state?.batch_id) : current;
      if (normalized.length > 0) {
        await chrome.storage.local.set({ [GPT_BATCH_STATES_STORAGE_KEY]: normalized });
      } else {
        await chrome.storage.local.remove([GPT_BATCH_STATES_STORAGE_KEY]);
      }
      await chrome.storage.local.remove([LEGACY_GPT_BATCH_STATE_STORAGE_KEY]);
      return normalized;
    });
  runtimeGptBatchStateMutationPromise = operation.then(() => undefined, () => undefined);
  return operation;
}

function replaceStoredRuntimeGptBatchStates(states) {
  return mutateStoredRuntimeGptBatchStates(() => states);
}

function upsertStoredRuntimeGptBatchState(state) {
  return mutateStoredRuntimeGptBatchStates((states) => {
    const index = states.findIndex((candidate) => candidate.batch_id === state.batch_id);
    if (index >= 0) states[index] = state;
    else states.push(state);
    return states;
  });
}

function removeStoredRuntimeGptBatchState(batchId) {
  return mutateStoredRuntimeGptBatchStates((states) => (
    states.filter((state) => state.batch_id !== batchId)
  ));
}

async function cleanupLegacyGptWorkerTab() {
  const stored = await chrome.storage.local.get([GPT_WORKER_TAB_STORAGE_KEY]).catch(() => ({}));
  const tabId = Number(stored[GPT_WORKER_TAB_STORAGE_KEY]);
  await chrome.storage.local.remove([GPT_WORKER_TAB_STORAGE_KEY]).catch(() => {});
  if (!Number.isInteger(tabId)) return;
  const tab = await chrome.tabs.get(tabId).catch(() => null);
  if (!tab || !isAutoBidGptWorkerUrl(tab.url || "")) return;
  await chrome.tabs.remove(tabId).catch(() => {});
}

function isAutoBidGptWorkerUrl(url) {
  try {
    const parsed = new URL(url);
    return isGptTabUrl(url) && parsed.searchParams.get("autobid_worker") === "1";
  } catch {
    return false;
  }
}

function createRuntimeGptBatchId() {
  if (globalThis.crypto?.randomUUID) return `abb_${globalThis.crypto.randomUUID()}`;
  return `abb_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

async function completeRuntimeGptBatch(payload = {}, sender = {}) {
  await ensureRuntimeGptRequestsLoaded();
  const senderTabId = Number.isInteger(sender?.tab?.id) ? sender.tab.id : null;
  const states = await getStoredRuntimeGptBatchStates();
  const requestedBatchId = String(payload.batch_id || payload.batchId || "");
  const state = states.find((candidate) => (
    requestedBatchId && candidate.batch_id === requestedBatchId
  ) || (
    !requestedBatchId && Number.isInteger(senderTabId) && candidate.tab_id === senderTabId
  ));
  const batchId = requestedBatchId || state?.batch_id || "";

  if (!state?.batch_id) {
    return { completed: false, ignored: true, reason: "stale-batch", batch_id: batchId };
  }
  if (Number.isInteger(senderTabId) && Number.isInteger(state.tab_id) && senderTabId !== state.tab_id) {
    throw new Error("Only the active ChatGPT batch tab can complete this batch.");
  }

  const requestIds = Array.from(new Set([
    ...(Array.isArray(state.request_ids) ? state.request_ids : []),
    ...(Array.isArray(payload.request_ids) ? payload.request_ids : [])
  ].map(String).filter(Boolean)));
  const completedRequestIds = [];
  const queuedRequestIds = [];
  let changed = false;

  for (const requestId of requestIds) {
    const request = runtimeGptRequests.get(requestId);
    if (!request) continue;
    if (request.status === "complete" || request.status === "cancelled") {
      completedRequestIds.push(requestId);
      continue;
    }
    if (request.status === "processing" && request.batch_id === batchId) {
      failOrRequeueRuntimeGptRequest(
        request,
        "The ChatGPT batch ended before this answer was stored.",
        RUNTIME_GPT_RETRY_DELAY_MS
      );
      changed = true;
    }
    if (request.status === "pending") queuedRequestIds.push(requestId);
  }

  if (changed) await persistRuntimeGptRequests();
  const control = await chrome.storage.local.get([GPT_BATCH_PAUSED_STORAGE_KEY]).catch(() => ({}));
  if (control[GPT_BATCH_PAUSED_STORAGE_KEY]) {
    await removeStoredRuntimeGptBatchState(batchId);
  } else {
    await upsertStoredRuntimeGptBatchState({
      ...state,
      version: 3,
      status: "idle",
      request_ids: [],
      allow_sheet_fallback: false,
      updated_at: Date.now(),
      completed_at: Date.now()
    });
  }
  await releaseBackgroundTabAutomation(state.tab_id, "gpt-worker").catch(() => {});
  const remainingStates = await getStoredRuntimeGptBatchStates();
  await chrome.storage.local.set({
    autoBidGptWorkerState: {
      running: remainingStates.length > 0,
      mode: "persistent-worker-pool",
      batchId,
      completedRequests: completedRequestIds.length,
      queuedRequests: queuedRequestIds.length,
      activeWorkers: remainingStates.length,
      busyWorkers: remainingStates.filter((worker) => !isRuntimeGptWorkerIdle(worker)).length,
      maxParallelWorkers: RUNTIME_GPT_MAX_WORKERS,
      updatedAt: new Date().toISOString()
    }
  });

  console.info("[AutoBid Batch] durable batch complete", {
    batch_id: batchId,
    completed_request_ids: completedRequestIds,
    queued_request_ids: queuedRequestIds
  });

  return {
    completed: true,
    batch_id: batchId,
    completed_request_ids: completedRequestIds,
    queued_request_ids: queuedRequestIds,
    worker_reusable: !control[GPT_BATCH_PAUSED_STORAGE_KEY],
    tab_id: state.tab_id
  };
}

async function handleRuntimeGptWorkerReady(payload = {}, sender = {}) {
  const senderTabId = Number.isInteger(sender?.tab?.id) ? sender.tab.id : null;
  if (!Number.isInteger(senderTabId)) return { ready: false, reason: "missing-worker-tab" };

  const states = await getStoredRuntimeGptBatchStates();
  const state = states.find((candidate) => candidate.tab_id === senderTabId);
  if (!state?.batch_id) return { ready: false, reason: "retired-worker" };

  const reportedBatchId = String(payload.batch_id || payload.batchId || "");
  if (reportedBatchId && reportedBatchId !== state.batch_id) {
    return { ready: false, reason: "stale-worker-ready" };
  }

  if (!isRuntimeGptWorkerIdle(state)) {
    await completeRuntimeGptBatch({
      batch_id: state.batch_id,
      request_ids: state.request_ids || [],
      error: "The worker became ready before its completion acknowledgement was received."
    }, sender);
  }

  const scheduled = await ensureRuntimeGptBatchWorker();
  return { ready: true, worker_id: state.batch_id, scheduled };
}

async function releaseRuntimeGptBatchRequests(batchId, reason, retryDelayMs = RUNTIME_GPT_RETRY_DELAY_MS) {
  if (!batchId) return { released: 0 };
  await ensureRuntimeGptRequestsLoaded();
  const now = Date.now();
  let released = 0;

  for (const request of runtimeGptRequests.values()) {
    if (request.status !== "processing" || request.batch_id !== batchId) continue;
    failOrRequeueRuntimeGptRequest(request, reason || "ChatGPT batch interrupted.", retryDelayMs);
    request.updated_at = now;
    released += 1;
  }

  if (released) {
    await persistRuntimeGptRequests();
    console.warn("[AutoBid Batch] released interrupted requests", {
      batch_id: batchId,
      released,
      reason
    });
  }
  return { released };
}

async function handleRuntimeGptBatchTabRemoved(tabId) {
  const states = await getStoredRuntimeGptBatchStates();
  const state = states.find((candidate) => Number(candidate.tab_id) === Number(tabId));
  if (!state?.batch_id) return;

  await releaseRuntimeGptBatchRequests(
    state.batch_id,
    "The ChatGPT batch tab closed before completion.",
    RUNTIME_GPT_RETRY_DELAY_MS
  );
  await removeStoredRuntimeGptBatchState(state.batch_id);
  const remainingStates = await getStoredRuntimeGptBatchStates();
  await chrome.storage.local.set({
    autoBidGptWorkerState: {
      running: remainingStates.length > 0,
      mode: "persistent-worker-pool",
      batchId: state.batch_id,
      activeWorkers: remainingStates.length,
      maxParallelWorkers: RUNTIME_GPT_MAX_WORKERS,
      interruptedAt: new Date().toISOString()
    }
  });

  globalThis.setTimeout(() => {
    ensureRuntimeGptBatchWorker().catch(() => {});
  }, RUNTIME_GPT_RETRY_DELAY_MS);
}

function isGptTabUrl(url) {
  try {
    const host = new URL(url).hostname;
    return host === "chatgpt.com" || host === "chat.openai.com";
  } catch {
    return false;
  }
}

async function waitForTabLoaded(tabId, timeoutMs = 30000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const tab = await chrome.tabs.get(tabId).catch(() => null);
    if (!tab) throw new Error("ChatGPT worker tab was closed.");
    if (tab.status === "complete") return tab;
    await sleep(300);
  }
  return chrome.tabs.get(tabId);
}

async function injectGptAnswerWorker(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["gpt-answer-worker.js"]
  });
}

async function sendTabMessageWithReinject(tabId, message) {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (_error) {
    await injectGptAnswerWorker(tabId);
    await sleep(300);
    return chrome.tabs.sendMessage(tabId, message);
  }
}

async function fetchPendingGptRows() {
  const settings = await chrome.storage.local.get(["sheetSettings"]);
  const sheetSettings = normalizeSheetSettings(settings.sheetSettings || {});
  if (!sheetSettings.sheetName || !sheetSettings.startRow || !sheetSettings.endRow) {
    return { rows: [] };
  }
  return apiFetch("/sheets/pending-questions", {
    method: "POST",
    body: sheetSettings
  });
}

async function saveGptAnswers(payload) {
  return apiFetch("/sheets/save-answers", {
    method: "POST",
    body: {
      spreadsheet_id: payload.spreadsheetId || payload.spreadsheet_id,
      sheet_name: payload.sheetName || payload.sheet_name,
      row_number: payload.rowNumber || payload.row_number,
      answers: payload.answers || [],
      payload: payload.payload || null
    }
  });
}

async function getSheetContextForPage(payload, sender) {
  const url = String(payload.url || payload.page?.url || sender?.tab?.url || "");
  const alternateUrl = String(sender?.tab?.url || "");
  const tabId = Number.isInteger(sender?.tab?.id) ? String(sender.tab.id) : "";
  const forceRefresh = Boolean(payload.forceRefresh || payload.force_refresh);
  const settings = await chrome.storage.local.get(["sheetSettings", "sheetJobs", "sheetTabJobs"]);
  const sheetSettings = normalizeSheetSettings(settings.sheetSettings || {});
  const sheetJobs = settings.sheetJobs || {};
  const tabContext = tabId ? settings.sheetTabJobs?.[tabId] : null;
  const normalizedPageUrl = normalizeUrlForMatch(url || alternateUrl);
  const tabContextScore = tabContext?.url
    ? scoreSheetContextUrlMatch(normalizeUrlForMatch(tabContext.url), normalizedPageUrl)
    : 0;
  if (!forceRefresh && tabContext?.rowNumber && tabContextScore >= 600) {
    return withSheetMatchSource(tabContext, "tab");
  }

  if (!forceRefresh) {
    const context = findSheetContext(sheetJobs, url) || findSheetContext(sheetJobs, alternateUrl);
    if (context) return context;
  }

  if (!sheetSettings.sheetName || !sheetSettings.startRow || !sheetSettings.endRow) return null;
  const data = await apiFetch("/sheets/jobs", {
    method: "POST",
    body: sheetSettings
  });
  const jobs = Array.isArray(data.jobs) ? data.jobs : [];
  const nextJobs = {};
  const refreshedContexts = [];
  jobs.forEach((job) => {
    if (!job.url) return;
    const context = {
      ...sheetSettings,
      rowNumber: job.rowNumber,
      url: job.url,
      values: job.values || {},
      raw: job.raw || []
    };
    refreshedContexts.push(context);
    addSheetJobContext(nextJobs, context);
  });

  const rowContext = tabContext?.rowNumber
    ? refreshedContexts.find((context) => Number(context.rowNumber) === Number(tabContext.rowNumber))
    : null;
  const resolvedContext = rowContext
    ? withSheetMatchSource(rowContext, "tab-refresh")
    : findSheetContext(nextJobs, url) || findSheetContext(nextJobs, alternateUrl);
  const nextTabJobs = { ...(settings.sheetTabJobs || {}) };
  if (tabId && resolvedContext?.rowNumber) nextTabJobs[tabId] = resolvedContext;
  await chrome.storage.local.set({ sheetJobs: nextJobs, sheetTabJobs: nextTabJobs });
  return resolvedContext;
}

async function submitSheetQuestions(payload, sender) {
  const context = payload.context || await getSheetContextForPage({ url: payload.page?.url }, sender);
  if (!context?.rowNumber) throw new Error("No matching Google Sheet row found for this page");
  return apiFetch("/sheets/questions", {
    method: "POST",
    body: {
      spreadsheet_id: context.spreadsheetId,
      sheet_name: context.sheetName,
      row_number: context.rowNumber,
      payload: payload.payload || {}
    }
  });
}

async function fetchSheetAnswers(payload, sender) {
  const context = payload.context || await getSheetContextForPage({ url: payload.page?.url }, sender);
  if (!context?.rowNumber) return { answers: [], raw: "", row_number: null };
  return apiFetch("/sheets/answers", {
    method: "POST",
    body: {
      spreadsheet_id: context.spreadsheetId,
      sheet_name: context.sheetName,
      row_number: context.rowNumber
    }
  });
}

async function fetchSheetResumeFile(payload, sender) {
  let context = payload.context || await getSheetContextForPage({
    url: payload.page?.url,
    force_refresh: payload.force_refresh
  }, sender);
  if (!context?.rowNumber) return { base64: "", filename: "", mime_type: "", row_number: null };

  let result = await requestSheetResumeFile(context, payload);
  if (result?.base64 || payload.force_refresh) {
    return withResumeFetchMetadata(result, context, Boolean(payload.force_refresh));
  }

  context = await getSheetContextForPage({ url: payload.page?.url, force_refresh: true }, sender) || context;
  result = await requestSheetResumeFile(context, payload);
  return withResumeFetchMetadata(result, context, true);
}

function requestSheetResumeFile(context, payload) {
  return apiFetch("/sheets/resume-file", {
    method: "POST",
    body: {
      spreadsheet_id: context.spreadsheetId,
      sheet_name: context.sheetName,
      row_number: context.rowNumber,
      resume_url: getContextResumeFileUrl(context),
      row_values: context.values || {},
      raw: Array.isArray(context.raw) ? context.raw : [],
      accept: payload.accept || []
    }
  });
}

function withResumeFetchMetadata(result, context, refreshed) {
  const payload = result || {};
  const hasResumeUrl = Boolean(getContextResumeFileUrl(context));
  return {
    ...payload,
    row_number: payload.row_number || payload.rowNumber || context?.rowNumber || null,
    refreshed,
    reason: payload.base64 ? "" : hasResumeUrl ? "resume-file-empty" : "resume-link-missing"
  };
}

function getContextResumeFileUrl(context) {
  const values = context?.values || {};
  const raw = Array.isArray(context?.raw) ? context.raw : [];
  const candidates = [
    values.resume_link,
    values.resume_url,
    values.resume_pdf,
    values.generated_resume,
    values.generated_resume_link,
    values.tailored_resume_link,
    values.cv_link,
    values.pdf_resume,
    values.column_j,
    raw[9]
  ];

  return candidates.map(findFirstHttpUrl).find(Boolean) || "";
}

function findFirstHttpUrl(value) {
  const match = String(value || "").match(/https?:\/\/[^\s"'<>),]+/i);
  return match ? match[0] : "";
}

async function ensureDevSession() {
  const settings = await getSettings();
  if (settings.user && settings.selectedProfileId && (!DEV_AUTH_BYPASS || settings.token)) return settings;

  const apiBase = (settings.apiBase || DEFAULT_API_BASE).replace(/\/+$/, "");
  try {
    const response = await fetch(`${apiBase}/api/auto-bid/auth/dev-session`, {
      method: "POST",
      headers: { "Content-Type": "application/json" }
    });
    const json = await response.json().catch(() => ({}));

    if (!response.ok || json.errors) {
      throw new Error(json.errors?.[0]?.message || `Dev session failed with ${response.status}`);
    }

    await chrome.storage.local.set({
      token: json.data.token,
      user: json.data.user,
      selectedProfileId: json.data.profile?.id || DEV_PROFILE_ID,
      devProfile: json.data.profile || getDefaultDevProfile()
    });
  } catch (_error) {
    await chrome.storage.local.set({
      user: DEV_USER,
      selectedProfileId: DEV_PROFILE_ID,
      devProfile: await getLocalDevProfile()
    });
  }

  return getSettings();
}

async function authRequest(path, payload) {
  const data = await apiFetch(path, { method: "POST", body: payload, skipAuth: true });
  await chrome.storage.local.set({ token: data.token, user: data.user });
  return data;
}

async function listProfiles() {
  if (!DEV_AUTH_BYPASS) return apiFetch("/profiles");
  try {
    return await apiFetch("/profiles");
  } catch (_error) {
    return [await getLocalDevProfile()];
  }
}

async function saveProfile(payload) {
  const profile = normalizeProfilePayload(payload);
  const creatingProfile = !profile.id;
  const path = profile.id ? `/profiles/${encodeURIComponent(profile.id)}` : "/profiles";
  const method = profile.id ? "PATCH" : "POST";
  delete profile.id;

  let saved;
  try {
    saved = await apiFetch(path, { method, body: profile });
  } catch (error) {
    if (!DEV_AUTH_BYPASS) throw error;
    throw new Error(`Profile was not saved. Start the AutoBid server and PostgreSQL, then try again. ${error.message || String(error)}`);
  }

  if (!saved?.id) throw new Error("Profile API returned an invalid saved profile.");
  if (creatingProfile && saved.id === DEV_PROFILE_ID) {
    throw new Error("Profile was not saved. The backend is running without PostgreSQL; restart it after PostgreSQL is ready and try again.");
  }
  if (DEV_AUTH_BYPASS) {
    await chrome.storage.local.set({ devProfile: saved, user: DEV_USER });
  }
  await chrome.storage.local.set({ selectedProfileId: saved.id });
  return saved;
}

async function deleteProfile(profileId) {
  if (!profileId) throw new Error("Profile id is required");
  let result;
  try {
    result = await apiFetch(`/profiles/${encodeURIComponent(profileId)}`, { method: "DELETE" });
  } catch (error) {
    if (!DEV_AUTH_BYPASS) throw error;
    const profile = getDefaultDevProfile();
    await chrome.storage.local.set({ devProfile: profile, selectedProfileId: profile.id, user: DEV_USER });
    result = { ok: true };
  }
  const settings = await getSettings();
  if (settings.selectedProfileId === profileId) {
    if (DEV_AUTH_BYPASS) await chrome.storage.local.set({ selectedProfileId: DEV_PROFILE_ID });
    else await chrome.storage.local.remove("selectedProfileId");
  }
  return result;
}

async function assist(payload) {
  const settings = DEV_AUTH_BYPASS ? await ensureDevSession() : await getSettings();
  if (!settings.selectedProfileId) throw new Error("Select a profile in the Auto Bid popup first");
  const started = await apiFetch("/assist/start", {
    method: "POST",
    body: { ...payload, profile_id: settings.selectedProfileId }
  });
  const data = await waitForAssistJob(started.job_id);
  const profile = await loadSelectedProfileForStatic(settings).catch(() => null);
  return mergeProfileStaticAnswers(data, payload.fields || [], profile?.static_fields || {});
}

async function waitForAssistJob(jobId) {
  if (!jobId) throw new Error("Assist job did not start");

  const startedAt = Date.now();
  const timeoutMs = 4 * 60 * 1000;
  while (Date.now() - startedAt < timeoutMs) {
    const job = await apiFetch(`/assist/jobs/${encodeURIComponent(jobId)}`);
    if (job.status === "complete") return job.result || {};
    if (job.status === "error") throw new Error(job.error || "Assist job failed");
    await sleep(700);
  }

  throw new Error("Assist job timed out");
}

async function getSelectedProfileStaticFields() {
  const settings = DEV_AUTH_BYPASS ? await ensureDevSession() : await getSettings();
  if (!settings.selectedProfileId) throw new Error("Select a profile in the Auto Bid popup first");
  const profile = await loadSelectedProfileForStatic(settings);
  return {
    profile_id: profile?.id || settings.selectedProfileId,
    static_fields: profile?.static_fields || {}
  };
}

async function apiFetch(path, options = {}) {
  let settings = await getSettings();
  if (DEV_AUTH_BYPASS && !options.skipAuth && (!settings.token || !settings.selectedProfileId)) {
    settings = await ensureDevSession();
  }
  const url = `${settings.apiBase.replace(/\/+$/, "")}/api/auto-bid${path}`;
  const headers = { "Content-Type": "application/json" };

  if (!options.skipAuth) {
    if (settings.token) headers.Authorization = `Bearer ${settings.token}`;
    else if (!DEV_AUTH_BYPASS) throw new Error("Login required");
  }

  const response = await fetch(url, {
    method: options.method || "GET",
    headers,
    body: options.body === undefined ? undefined : JSON.stringify(options.body)
  });
  const json = await response.json().catch(() => ({}));

  if (!response.ok || json.errors) {
    throw new Error(json.errors?.[0]?.message || `Request failed with ${response.status}`);
  }

  return json.data;
}

async function loadSelectedProfileForStatic(settings) {
  const localProfile = DEV_AUTH_BYPASS ? await getLocalDevProfile().catch(() => null) : null;
  try {
    const profiles = await apiFetch("/profiles");
    const profile = profiles.find((item) => item.id === settings.selectedProfileId) || profiles[0] || null;
    const shouldMergeLocal = DEV_AUTH_BYPASS && localProfile && profile && localProfile.id === profile.id;
    const mergedProfile = shouldMergeLocal ? mergeProfileStaticFields(profile, localProfile) : (profile || localProfile);
    if (DEV_AUTH_BYPASS && mergedProfile) await chrome.storage.local.set({ devProfile: mergedProfile });
    return mergedProfile;
  } catch (error) {
    if (DEV_AUTH_BYPASS) return localProfile || getLocalDevProfile();
    throw error;
  }
}

function mergeProfileStaticAnswers(data, fields, staticFields) {
  const answers = Array.isArray(data?.answers) ? [...data.answers] : [];
  const answeredFieldIds = new Set(answers.map((answer) => answer.field_id));
  const staticMerge = {
    available_keys: Object.keys(staticFields || {}).filter((key) => String(staticFields[key] || "").trim()),
    matched: [],
    filled: [],
    missing: []
  };

  for (const field of Array.isArray(fields) ? fields : []) {
    if (!field?.id || answeredFieldIds.has(field.id) || String(field.value || "").trim()) continue;
    const key = matchStaticFieldKey(field);
    if (!key) continue;
    staticMerge.matched.push({
      field_id: field.id,
      key,
      label: field.label || ""
    });
    const value = getStaticFieldValue(staticFields, key);
    if (value === undefined || value === null || !String(value).trim()) {
      staticMerge.missing.push({
        field_id: field.id,
        key,
        label: field.label || ""
      });
      continue;
    }

    answers.push({
      field_id: field.id,
      value: String(value),
      source: "static-local",
      cache_scope: "profile",
      confidence: 1,
      warning: null
    });
    staticMerge.filled.push({
      field_id: field.id,
      key,
      label: field.label || ""
    });
    answeredFieldIds.add(field.id);
  }

  return { ...(data || {}), answers, static_merge: staticMerge };
}

function mergeProfileStaticFields(serverProfile, localProfile) {
  if (!serverProfile && !localProfile) return null;
  if (!serverProfile) return localProfile;
  if (!localProfile) return serverProfile;

  const serverFields = serverProfile.static_fields || {};
  const localFields = localProfile.static_fields || {};
  return {
    ...serverProfile,
    static_fields: {
      ...serverFields,
      ...Object.fromEntries(
        Object.entries(localFields).filter(([_key, value]) => value !== undefined && value !== null && String(value).trim())
      )
    }
  };
}

function matchStaticFieldKey(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  if (isLanguageYesNoField(field)) return null;
  if (isCombinedLocationField(field)) return "location";
  if (isSemanticBasedInQuestion(field)) return null;
  if (isPlainFullNameField(field)) return "full_name";
  if (isPhoneStaticField(field)) return "phone";
  const addressComponentKey = matchAddressComponentFieldKey(field);
  if (addressComponentKey) return addressComponentKey;
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name"]],
    ["email", ["email", "e mail", "mail"]],
    ["location", ["location", "address", "current city", "current location"]],
    ["city", ["city"]],
    ["country", ["country", "residence", "current residence", "where is your current residence", "where are you based"]],
    ["linkedin", ["linkedin"]],
    ["github", ["github"]],
    ["portfolio", ["portfolio"]],
    ["website", ["website", "personal site", "web site"]],
    ["languages", ["languages", "spoken languages", "language proficiency", "fluent languages", "languages spoken"]],
    ["expected_rate", ["hourly rate", "rate", "expected rate", "expected salary", "salary expectation", "salary expectations", "expected compensation", "desired salary", "desired compensation", "gross monthly", "monthly salary", "salary", "compensation"]],
    ["work_authorization", ["authorized", "authorization", "legally work", "eligible to work"]],
    ["sponsorship", ["sponsor", "sponsorship", "visa"]],
    ["availability", ["availability", "available", "start date"]],
    ["notice_period", ["notice period", "current notice", "notice"]]
  ];

  for (const [key, needles] of patterns) {
    if (needles.some((needle) => includesNormalizedPhrase(text, needle))) return key;
  }

  return null;
}

function includesNormalizedPhrase(text, phrase) {
  const normalizedPhrase = normalizeText(phrase);
  return Boolean(normalizedPhrase) && ` ${text} `.includes(` ${normalizedPhrase} `);
}

function isPlainFullNameField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const candidates = [field.label, field.name, field.autocomplete]
    .map(normalizeText)
    .filter(Boolean);
  return candidates.some((candidate) =>
    ["name", "your name", "applicant name", "candidate name", "full name", "preferred name"].includes(candidate) ||
    /\bfirst(?:\s+and|\s*\/)\s*last\s+name\b|\blast(?:\s+and|\s*\/)\s*first\s+name\b/.test(candidate)
  );
}

function isCombinedLocationField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const text = normalizeText([field.label, field.name, field.placeholder].filter(Boolean).join(" "));
  if (!/(location|where.*based|based.*in|residence)/.test(text)) return false;
  const parts = ["city", "state", "country"].filter((part) => new RegExp(`\\b${part}\\b`).test(text));
  return parts.length >= 2 || /what location.*based|where.*(?:located|based|reside)/.test(text);
}

function isTextLikeStaticField(field) {
  return !["checkbox", "radio", "select", "combobox", "button-group", "file", "hidden", "password", "submit", "button", "reset"].includes(field.type);
}

function isLanguageYesNoField(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  const options = (field.options || []).map(normalizeText);
  return /(speak|language|fluent|fluency|proficien|native speaker|bilingual|multilingual)/.test(text) &&
    options.includes("yes") &&
    options.includes("no");
}

function isSemanticBasedInQuestion(field) {
  const text = normalizeText([field?.label, field?.name, field?.placeholder].filter(Boolean).join(" "));
  return /(currently.*based.*in|based.*in|currently.*located.*in|located.*in|currently.*living.*in|currently.*residing.*in|resident.*in)/.test(text);
}

function isPhoneStaticField(field) {
  const type = normalizeText(field?.type || "");
  const autocomplete = normalizeText(field?.autocomplete || "");
  const text = normalizeText([field?.name, field?.label, field?.placeholder].filter(Boolean).join(" "));
  if (type === "tel" || /^(tel|phone|mobile)$/.test(autocomplete)) return true;
  if (/\b(phone|telephone|cell)(?:\s+number)?\b/.test(text)) return true;
  return /\bmobile\b/.test(text) && (
    /\bmobile\s+(?:phone|number|contact)\b/.test(text) ||
    /^(?:your\s+)?mobile(?:\s+number)?$/.test(text)
  );
}

function getStaticFieldValue(staticFields, key) {
  const aliases = {
    postal_code: ["postal_code", "postalcode", "post_code", "postcode", "zip_code", "zipcode", "zip"],
    state_region: ["state_region", "state_province_region", "state_province", "state", "province", "region", "administrative_area"],
    expected_rate: ["expected_rate", "expected_salary", "salary_expectation", "salary_expectations", "monthly_salary", "monthly_salary_expectation", "desired_salary", "desired_compensation"],
    notice_period: ["notice_period", "current_notice_period", "availability_notice"],
    languages: ["languages", "language", "spoken_languages", "language_proficiency", "fluent_languages", "languages_spoken"],
    work_authorization: ["work_authorization", "right_to_work", "work_status", "employment_status"]
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

function normalizeText(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function matchAddressComponentFieldKey(field) {
  const autocomplete = normalizeText(field.autocomplete || "");
  const name = normalizeText(field.name || "");
  const prompt = simplifyAddressPrompt(field.label || field.placeholder || "");

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

  return null;
}

function simplifyAddressPrompt(value) {
  return normalizeText(value)
    .replace(/^(?:what is|please enter|please select|enter|select|choose|provide)\s+(?:your\s+)?(?:current\s+)?/, "")
    .replace(/\s+(?:required|optional)$/, "")
    .trim();
}

function normalizeSheetSettings(value) {
  const startRow = Math.max(2, Number(value.startRow || value.start_row || 2));
  const endRow = Math.max(startRow, Number(value.endRow || value.end_row || startRow));
  return {
    spreadsheetId: normalizeSpreadsheetId(value.spreadsheetId || value.spreadsheet_id || ""),
    sheetName: String(value.sheetName || value.sheet_name || "").trim(),
    startRow,
    endRow
  };
}

function normalizeSheetJobsResponse(data) {
  const candidates = Array.isArray(data?.jobs)
    ? data.jobs
    : Array.isArray(data?.rows)
      ? data.rows
      : Array.isArray(data)
        ? data
        : [];

  return candidates.map((job, index) => ({
    ...job,
    rowNumber: Number(job.rowNumber || job.row_number || job.row || 0) || null,
    url: normalizeJobOpenUrl(job.url || findJobUrlInSheetJob(job)),
    values: job.values || job.rowValues || job.row_values || {},
    raw: Array.isArray(job.raw) ? job.raw : [],
    index
  }));
}

function findJobUrlInSheetJob(job = {}) {
  const values = job.values || job.rowValues || job.row_values || {};
  const raw = Array.isArray(job.raw) ? job.raw : [];
  const direct = [
    job.url,
    job.jobUrl,
    job.job_url,
    job.applyUrl,
    job.apply_url,
    job.applicationUrl,
    job.application_url,
    job.link,
    values.url,
    values.job_url,
    values.jobUrl,
    values.apply_url,
    values.applyUrl,
    values.application_url,
    values.applicationUrl,
    values.link
  ].map(findFirstUrlInText).find(Boolean);
  if (direct) return direct;

  for (const value of raw) {
    const url = findFirstUrlInText(value);
    if (url) return url;
  }

  for (const value of Object.values(values)) {
    const url = findFirstUrlInText(value);
    if (url) return url;
  }

  return "";
}

function normalizeJobOpenUrl(value) {
  const url = findFirstUrlInText(value);
  if (!url) return "";

  try {
    const parsed = new URL(url);
    if (!["http:", "https:"].includes(parsed.protocol)) return "";
    return parsed.href;
  } catch {
    return "";
  }
}

function findFirstUrlInText(value) {
  const match = String(value || "").trim().match(/https?:\/\/[^\s"'<>),]+/i);
  return match ? match[0] : "";
}

function normalizeSpreadsheetId(value) {
  const text = String(value || "").trim();
  const match = text.match(/\/spreadsheets\/d\/([^/?#]+)/);
  return match ? match[1] : text;
}

function addSheetJobContext(sheetJobs, context) {
  const key = normalizeUrlForMatch(context.url);
  if (!key) return;

  const existing = sheetJobs[key];
  if (!existing) {
    sheetJobs[key] = [context];
    return;
  }

  if (Array.isArray(existing)) existing.push(context);
  else sheetJobs[key] = [existing, context];
}

function findSheetContext(sheetJobs, pageUrl) {
  const normalized = normalizeUrlForMatch(pageUrl);
  if (!normalized) return null;

  const exact = pickBestSheetContext(sheetJobs[normalized], normalized);
  if (exact) return withSheetMatchSource(exact, "url-exact");

  const candidates = Object.entries(sheetJobs)
    .flatMap(([url, value]) => normalizeSheetContextList(value).map((context) => ({
      context,
      score: scoreSheetContextUrlMatch(url, normalized)
    })))
    .filter((candidate) => candidate.score >= 600)
    .sort((left, right) => right.score - left.score);

  return candidates[0] ? withSheetMatchSource(candidates[0].context, "url-fuzzy") : null;
}

function normalizeSheetContextList(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function pickBestSheetContext(value, normalizedPageUrl) {
  const contexts = normalizeSheetContextList(value);
  if (contexts.length <= 1) return contexts[0] || null;

  return contexts
    .map((context) => ({
      context,
      score: scoreSheetContextUrlMatch(normalizeUrlForMatch(context.url), normalizedPageUrl)
    }))
    .sort((left, right) => right.score - left.score || Number(left.context.rowNumber || 0) - Number(right.context.rowNumber || 0))[0]?.context || null;
}

function scoreSheetContextUrlMatch(sheetUrl, pageUrl) {
  if (!sheetUrl || !pageUrl) return 0;
  if (sheetUrl === pageUrl) return 1000;
  if (pageUrl.startsWith(`${sheetUrl}/`) || sheetUrl.startsWith(`${pageUrl}/`)) return 800;
  if (pageUrl.includes(sheetUrl) || sheetUrl.includes(pageUrl)) return 600;
  return sameUrlHost(sheetUrl, pageUrl) ? 100 : 0;
}

function sameUrlHost(left, right) {
  try {
    return new URL(left).hostname === new URL(right).hostname;
  } catch {
    return false;
  }
}

function withSheetMatchSource(context, source) {
  return context ? { ...context, matchSource: source } : null;
}

function normalizeUrlForMatch(url) {
  try {
    const parsed = new URL(String(url || ""));
    const path = parsed.pathname.replace(/\/+$/, "") || "/";
    return `${parsed.origin}${path}`.toLowerCase();
  } catch {
    return String(url || "").split("?")[0].replace(/\/+$/, "").toLowerCase();
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeProfilePayload(payload) {
  return {
    id: payload.id || undefined,
    name: payload.name || "Default profile",
    static_fields: payload.static_fields || {},
    resume_text: payload.resume_text || "",
    preferences: payload.preferences || {}
  };
}

async function getLocalDevProfile() {
  const settings = await chrome.storage.local.get(["devProfile"]);
  return normalizeLocalDevProfile(settings.devProfile || getDefaultDevProfile());
}

function getDefaultDevProfile() {
  return normalizeLocalDevProfile({
    id: DEV_PROFILE_ID,
    user_id: DEV_USER.id,
    name: "Development profile",
    static_fields: {},
    resume_text: "",
    preferences: {
      tone: "direct, confident, concise",
      bid_style: "short proposal"
    },
    profile_version: 1
  });
}

function normalizeLocalDevProfile(profile) {
  return {
    id: profile.id || DEV_PROFILE_ID,
    user_id: profile.user_id || DEV_USER.id,
    name: profile.name || "Development profile",
    static_fields: profile.static_fields || {},
    resume_text: profile.resume_text || "",
    preferences: profile.preferences || {},
    profile_version: Number(profile.profile_version || 1),
    created_at: profile.created_at || new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
}
