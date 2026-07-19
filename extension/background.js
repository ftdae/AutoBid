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
const NATIVE_DEBUGGER_IDLE_MS = 5000;
const nativeInputQueues = new Map();
const nativeDebuggerSessions = new Map();

chrome.runtime.onInstalled.addListener(async () => {
  const settings = await chrome.storage.local.get(["apiBase"]);
  if (!settings.apiBase) {
    await chrome.storage.local.set({ apiBase: DEFAULT_API_BASE });
  }
});

chrome.commands.onCommand.addListener((command) => {
  if (command !== "trigger-auto-bid") return;
  triggerAutoBidInActiveTab().catch((error) => {
    console.error("Auto Bid shortcut failed", error);
  });
});

async function triggerAutoBidInActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  await injectAutoBidScripts(tab.id);
}

async function injectAutoBidScripts(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId, allFrames: true },
    files: ["page-helper.js"],
    world: "MAIN"
  });
  await chrome.scripting.executeScript({
    target: { tabId, allFrames: true },
    files: [
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
    case "GET_SETTINGS":
      return getSettings();
    case "DEV_SESSION":
      return ensureDevSession();
    case "SAVE_SETTINGS":
      return saveSettings(message.payload || {});
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
    case "DRAFT_STATUS":
      return apiFetch(`/drafts/${encodeURIComponent(message.draftId)}/status`, {
        method: "POST",
        body: { status: message.status }
      });
    default:
      throw new Error("Unknown Auto Bid message");
  }
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
    throw new Error("Invalid native click coordinates");
  }

  const target = await acquireNativeDebugger(tabId);
  try {
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x,
      y,
      button: "none",
      buttons: 0
    });
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mousePressed",
      x,
      y,
      button: "left",
      buttons: 1,
      clickCount: 1
    });
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x,
      y,
      button: "left",
      buttons: 0,
      clickCount: 1
    });
    scheduleNativeDebuggerDetach(tabId);
    return { clicked: true };
  } catch (error) {
    await detachNativeDebugger(tabId);
    const message = error?.message || String(error);
    if (/another debugger|already attached|cannot attach/i.test(message)) {
      throw new Error("Auto Bid could not control the dropdown. Close DevTools for this tab and try again.");
    }
    throw new Error(`Native dropdown click failed: ${message}`);
  }
}

async function dispatchNativeType(tabId, payload) {
  const text = String(payload.text || "").slice(0, 200);
  if (!text) throw new Error("Native typing requires text");

  const target = await acquireNativeDebugger(tabId);
  try {
    await dispatchNativeKey(target, "keyDown", "a", "KeyA", 65, 2);
    await dispatchNativeKey(target, "keyUp", "a", "KeyA", 65, 2);
    await dispatchNativeKey(target, "keyDown", "Backspace", "Backspace", 8);
    await dispatchNativeKey(target, "keyUp", "Backspace", "Backspace", 8);
    await chrome.debugger.sendCommand(target, "Input.insertText", { text });
    if (payload.commit !== false) {
      await dispatchNativeKey(target, "keyDown", "Enter", "Enter", 13);
      await dispatchNativeKey(target, "keyUp", "Enter", "Enter", 13);
    }
    scheduleNativeDebuggerDetach(tabId);
    return { typed: true };
  } catch (error) {
    await detachNativeDebugger(tabId);
    const message = error?.message || String(error);
    if (/another debugger|already attached|cannot attach/i.test(message)) {
      throw new Error("Auto Bid could not type into the field. Close DevTools for this tab and try again.");
    }
    throw new Error(`Native typing failed: ${message}`);
  }
}

function dispatchNativeKey(target, type, key, code, windowsVirtualKeyCode, modifiers = 0) {
  return chrome.debugger.sendCommand(target, "Input.dispatchKeyEvent", {
    type,
    key,
    code,
    windowsVirtualKeyCode,
    nativeVirtualKeyCode: windowsVirtualKeyCode,
    modifiers
  });
}

async function acquireNativeDebugger(tabId) {
  const existing = nativeDebuggerSessions.get(tabId);
  if (existing?.attached) {
    scheduleNativeDebuggerDetach(tabId);
    return existing.target;
  }

  const target = { tabId };
  try {
    await chrome.debugger.attach(target, "1.3");
  } catch (error) {
    const message = error?.message || String(error);
    if (/another debugger|already attached|cannot attach/i.test(message)) {
      throw new Error("Auto Bid could not control the dropdown. Close DevTools for this tab and try again.");
    }
    throw new Error(`Native dropdown click failed: ${message}`);
  }

  nativeDebuggerSessions.set(tabId, { target, attached: true, detachTimer: null });
  scheduleNativeDebuggerDetach(tabId);
  return target;
}

function scheduleNativeDebuggerDetach(tabId) {
  const session = nativeDebuggerSessions.get(tabId);
  if (!session) return;
  if (session.detachTimer) clearTimeout(session.detachTimer);
  session.detachTimer = setTimeout(() => {
    detachNativeDebugger(tabId).catch(() => {});
  }, NATIVE_DEBUGGER_IDLE_MS);
}

async function detachNativeDebugger(tabId) {
  const session = nativeDebuggerSessions.get(tabId);
  if (!session) return;
  if (session.detachTimer) clearTimeout(session.detachTimer);
  nativeDebuggerSessions.delete(tabId);
  if (session.attached) await chrome.debugger.detach(session.target).catch(() => {});
}

async function getSettings() {
  const settings = await chrome.storage.local.get(["apiBase", "token", "user", "selectedProfileId", "devProfile"]);
  return {
    apiBase: settings.apiBase || DEFAULT_API_BASE,
    token: settings.token || null,
    user: DEV_AUTH_BYPASS ? (settings.user || DEV_USER) : (settings.user || null),
    selectedProfileId: DEV_AUTH_BYPASS ? (settings.selectedProfileId || DEV_PROFILE_ID) : (settings.selectedProfileId || null),
    devAuthBypass: DEV_AUTH_BYPASS
  };
}

async function saveSettings(payload) {
  const apiBase = String(payload.apiBase || DEFAULT_API_BASE).replace(/\/+$/, "");
  await chrome.storage.local.set({ apiBase });
  if (DEV_AUTH_BYPASS) await ensureDevSession();
  return getSettings();
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
  const path = profile.id ? `/profiles/${encodeURIComponent(profile.id)}` : "/profiles";
  const method = profile.id ? "PATCH" : "POST";
  const id = profile.id;
  delete profile.id;
  let saved;
  try {
    saved = await apiFetch(path, { method, body: profile });
  } catch (error) {
    if (!DEV_AUTH_BYPASS) throw error;
    saved = normalizeLocalDevProfile({ ...profile, id: id || DEV_PROFILE_ID });
    await chrome.storage.local.set({ devProfile: saved, selectedProfileId: saved.id, user: DEV_USER });
  }
  if (DEV_AUTH_BYPASS) {
    await chrome.storage.local.set({ devProfile: saved, user: DEV_USER });
  }
  const settings = await getSettings();
  if (!settings.selectedProfileId || settings.selectedProfileId === id) {
    await chrome.storage.local.set({ selectedProfileId: saved.id });
  }
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
  const data = await apiFetch("/assist", {
    method: "POST",
    body: { ...payload, profile_id: settings.selectedProfileId }
  });
  const profile = await loadSelectedProfileForStatic(settings).catch(() => null);
  return mergeProfileStaticAnswers(data, payload.fields || [], profile?.static_fields || {});
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
  if (isPlainFullNameField(field)) return "full_name";
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name"]],
    ["email", ["email", "e mail", "mail"]],
    ["phone", ["phone", "mobile", "telephone", "cell"]],
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
    if (needles.some((needle) => text.includes(needle))) return key;
  }

  return null;
}

function isPlainFullNameField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const candidates = [field.label, field.name, field.autocomplete]
    .map(normalizeText)
    .filter(Boolean);
  return candidates.some((candidate) => ["name", "your name", "applicant name", "candidate name", "full name"].includes(candidate));
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

function getStaticFieldValue(staticFields, key) {
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

function normalizeText(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
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
