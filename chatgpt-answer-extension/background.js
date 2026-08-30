const DEFAULT_SETTINGS = {
  appsScriptUrl: "",
  secret: "",
  sheetName: "",
  startRow: 2,
  endRow: 2,
  maxRowsPerRun: 10,
  running: false
};

chrome.runtime.onInstalled.addListener(async () => {
  const existing = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  await chrome.storage.local.set({ ...DEFAULT_SETTINGS, ...compactSettings(existing) });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  handleMessage(message, sender)
    .then((data) => sendResponse({ ok: true, data }))
    .catch((error) => sendResponse({ ok: false, error: error.message || String(error) }));
  return true;
});

async function handleMessage(message) {
  switch (message?.type) {
    case "GET_SETTINGS":
      return getSettings();
    case "SAVE_SETTINGS":
      return saveSettings(message.payload || {});
    case "RUN_ONCE":
      return sendToChatGptTab("AUTOBID_CHATGPT_RUN_ONCE");
    case "START":
      await chrome.storage.local.set({ running: true });
      return sendToChatGptTab("AUTOBID_CHATGPT_START");
    case "STOP":
      await chrome.storage.local.set({ running: false });
      return sendToChatGptTab("AUTOBID_CHATGPT_STOP").catch(() => ({ message: "Stopped" }));
    case "FETCH_PENDING_ROWS":
      return fetchPendingRows();
    case "SAVE_ANSWERS":
      return saveAnswers(message.payload || {});
    case "WORKER_STATE":
      await chrome.storage.local.set({ running: Boolean(message.payload?.running) });
      return getSettings();
    default:
      throw new Error("Unknown AutoBid Answer Worker message");
  }
}

async function getSettings() {
  const values = await chrome.storage.local.get(Object.keys(DEFAULT_SETTINGS));
  return {
    ...DEFAULT_SETTINGS,
    ...compactSettings(values),
    startRow: Math.max(2, Number(values.startRow || DEFAULT_SETTINGS.startRow)),
    endRow: Math.max(Math.max(2, Number(values.startRow || DEFAULT_SETTINGS.startRow)), Number(values.endRow || values.startRow || DEFAULT_SETTINGS.endRow)),
    maxRowsPerRun: Math.max(1, Math.min(50, Number(values.maxRowsPerRun || DEFAULT_SETTINGS.maxRowsPerRun))),
    running: Boolean(values.running)
  };
}

async function saveSettings(payload) {
  const startRow = Math.max(2, Number(payload.startRow || 2));
  const settings = {
    appsScriptUrl: String(payload.appsScriptUrl || "").trim(),
    secret: String(payload.secret || "").trim(),
    sheetName: String(payload.sheetName || "").trim(),
    startRow,
    endRow: Math.max(startRow, Number(payload.endRow || startRow)),
    maxRowsPerRun: Math.max(1, Math.min(50, Number(payload.maxRowsPerRun || 10)))
  };
  await chrome.storage.local.set(settings);
  return getSettings();
}

async function sendToChatGptTab(type) {
  const tab = await findChatGptTab();
  if (!tab?.id) throw new Error("Open a ChatGPT tab first.");

  await ensureContentScript(tab.id);
  return chrome.tabs.sendMessage(tab.id, { type });
}

async function findChatGptTab() {
  const [active] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (active && isChatGptUrl(active.url)) return active;

  const tabs = await chrome.tabs.query({ url: ["https://chatgpt.com/*", "https://chat.openai.com/*"] });
  return tabs[0] || null;
}

function isChatGptUrl(url) {
  return /^https:\/\/(chatgpt\.com|chat\.openai\.com)\//i.test(String(url || ""));
}

async function ensureContentScript(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: "AUTOBID_CHATGPT_PING" });
  } catch {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
  }
}

async function fetchPendingRows() {
  const settings = await getSettings();
  const data = await callAppsScript("autobidListJobs", {
    sheetName: settings.sheetName,
    startRow: settings.startRow,
    endRow: settings.endRow
  });

  const rows = (Array.isArray(data.jobs) ? data.jobs : [])
    .map(normalizeJobRow)
    .filter(isPendingQuestionRow)
    .slice(0, settings.maxRowsPerRun);

  return { rows };
}

async function saveAnswers(payload) {
  const settings = await getSettings();
  const sheetName = String(payload.sheetName || settings.sheetName || "").trim();
  const rowNumber = Number(payload.rowNumber || payload.row || 0);
  const answers = Array.isArray(payload.answers) ? payload.answers : [];
  if (!sheetName || !rowNumber) throw new Error("sheetName and rowNumber are required to save answers.");
  if (answers.length === 0) throw new Error("No answers were generated.");

  return callAppsScript("autobidSaveAnswers", {
    sheetName,
    rowNumber,
    answers,
    payload: {
      answers,
      source: "chatgpt-answer-extension",
      generated_at: new Date().toISOString()
    }
  });
}

async function callAppsScript(action, payload) {
  const settings = await getSettings();
  if (!settings.appsScriptUrl) throw new Error("Apps Script Web App URL is required.");
  if (!settings.secret) throw new Error("Extension secret is required.");
  if (!settings.sheetName && !payload.sheetName) throw new Error("Sheet tab is required.");

  const response = await fetch(settings.appsScriptUrl, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      action,
      secret: settings.secret,
      ...payload
    })
  });
  const text = await response.text();
  let data;

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Apps Script returned non-JSON response: ${text.slice(0, 180)}`);
  }

  if (!response.ok || data.status === "error" || data.error) {
    throw new Error(data.message || data.error || `Apps Script request failed with ${response.status}`);
  }

  return data;
}

function normalizeJobRow(row) {
  const values = row.values || {};
  const rawQuestions = String(values.autobid_questions || row.autobid_questions || "").trim();
  const rawAnswers = getAnswerCellValue(row, values);
  const status = normalizeText(values.autobid_status || row.autobid_status || "");
  return {
    sheetName: row.sheetName || row.sheet_name || "",
    rowNumber: Number(row.rowNumber || row.row_number || row.row || 0),
    url: String(row.url || row.applyUrl || row.apply_url || "").trim(),
    values,
    status,
    rawQuestions,
    rawAnswers,
    questions: parseJson(rawQuestions)
  };
}

function isPendingQuestionRow(row) {
  if (!row.rowNumber || !row.rawQuestions || !row.questions?.fields?.length) return false;
  if (hasAnswerPayload(row.rawAnswers)) return false;
  if (["answers_ready", "submitted", "complete", "completed"].includes(row.status)) return false;
  if (["answers_in_progress", "answering", "processing"].includes(row.status)) return false;
  if (row.status === "questions_pending") return true;
  return !row.rawAnswers;
}

function getAnswerCellValue(row, values) {
  const trusted = [
    values.autobid_answers,
    values.auto_bid_answers,
    values.answers,
    row.autobid_answers,
    row.auto_bid_answers,
    row.answers
  ].map((value) => String(value || "").trim()).find(Boolean);
  if (trusted) return trusted;

  const raw = Array.isArray(row.raw) ? row.raw : [];
  const columnR = String(raw[17] || "").trim();
  return looksLikeAnswerPayload(columnR) ? columnR : "";
}

function hasAnswerPayload(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  return looksLikeAnswerPayload(text) || !/^\s*[\[{]/.test(text);
}

function looksLikeAnswerPayload(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) return parsed.length > 0;
    if (Array.isArray(parsed?.answers)) return parsed.answers.length > 0;
    if (parsed && typeof parsed === "object") {
      return Boolean(parsed.field_id || parsed.value || Object.keys(parsed).length > 0);
    }
  } catch {
    return /\b(field_id|answers?|answer|value)\b/i.test(text);
  }
  return false;
}

function parseJson(text) {
  try {
    return text ? JSON.parse(text) : null;
  } catch {
    return null;
  }
}

function normalizeText(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function compactSettings(values) {
  return Object.fromEntries(
    Object.entries(values || {}).filter(([_key, value]) => value !== undefined && value !== null)
  );
}
