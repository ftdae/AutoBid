let mode = "login";
let settings = {};
let profiles = [];
let currentProfile = null;
let executionRuns = [];
let outlookConnection = null;
let outlookMessages = [];
const surface = new URLSearchParams(location.search).get("surface") || "popup";
let activeView = surface === "options" ? "settings" : "dashboard";
const CONTEXT_RECOVERY_STORAGE_KEY = "autoBidContextRecoveryAt";
const CONTEXT_RECOVERY_COOLDOWN_MS = 10000;
let contextRecoveryStarted = false;

const els = {
  sessionLabel: document.getElementById("sessionLabel"),
  logoutButton: document.getElementById("logoutButton"),
  apiBase: document.getElementById("apiBase"),
  saveSettingsButton: document.getElementById("saveSettingsButton"),
  spreadsheetId: document.getElementById("spreadsheetId"),
  sheetName: document.getElementById("sheetName"),
  sheetStartRow: document.getElementById("sheetStartRow"),
  sheetEndRow: document.getElementById("sheetEndRow"),
  saveSheetButton: document.getElementById("saveSheetButton"),
  openSheetRowsButton: document.getElementById("openSheetRowsButton"),
  dashboardOpenRowsButton: document.getElementById("dashboardOpenRowsButton"),
  dashboardSheetName: document.getElementById("dashboardSheetName"),
  dashboardSheetRows: document.getElementById("dashboardSheetRows"),
  sheetStatus: document.getElementById("sheetStatus"),
  authSection: document.getElementById("authSection"),
  profileSection: document.getElementById("profileSection"),
  loginTab: document.getElementById("loginTab"),
  signupTab: document.getElementById("signupTab"),
  signupNames: document.getElementById("signupNames"),
  firstName: document.getElementById("firstName"),
  lastName: document.getElementById("lastName"),
  email: document.getElementById("email"),
  password: document.getElementById("password"),
  authButton: document.getElementById("authButton"),
  profileSelect: document.getElementById("profileSelect"),
  newProfileButton: document.getElementById("newProfileButton"),
  profileName: document.getElementById("profileName"),
  resumeText: document.getElementById("resumeText"),
  tone: document.getElementById("tone"),
  bidStyle: document.getElementById("bidStyle"),
  saveProfileButton: document.getElementById("saveProfileButton"),
  deleteProfileButton: document.getElementById("deleteProfileButton"),
  closePanelButton: document.getElementById("closePanelButton"),
  runAutofillButton: document.getElementById("runAutofillButton"),
  dashboardRunState: document.getElementById("dashboardRunState"),
  latestRunSummary: document.getElementById("latestRunSummary"),
  latestRunMessage: document.getElementById("latestRunMessage"),
  dashboardOpenLogButton: document.getElementById("dashboardOpenLogButton"),
  executionLogs: document.getElementById("executionLogs"),
  refreshLogsButton: document.getElementById("refreshLogsButton"),
  clearLogsButton: document.getElementById("clearLogsButton"),
  openOptionsButton: document.getElementById("openOptionsButton"),
  outlookState: document.getElementById("outlookState"),
  outlookIdentity: document.getElementById("outlookIdentity"),
  outlookStatus: document.getElementById("outlookStatus"),
  connectOutlookButton: document.getElementById("connectOutlookButton"),
  refreshOutlookButton: document.getElementById("refreshOutlookButton"),
  outlookMessagesSection: document.getElementById("outlookMessagesSection"),
  outlookMessages: document.getElementById("outlookMessages"),
  status: document.getElementById("status")
};

document.addEventListener("DOMContentLoaded", init);
els.saveSettingsButton.addEventListener("click", saveSettings);
els.saveSheetButton.addEventListener("click", saveSheetSettings);
els.openSheetRowsButton.addEventListener("click", openSheetRows);
els.dashboardOpenRowsButton.addEventListener("click", openSheetRows);
els.loginTab.addEventListener("click", () => setMode("login"));
els.signupTab.addEventListener("click", () => setMode("signup"));
els.authButton.addEventListener("click", submitAuth);
els.logoutButton.addEventListener("click", logout);
els.profileSelect.addEventListener("change", selectProfile);
els.newProfileButton.addEventListener("click", newProfile);
els.saveProfileButton.addEventListener("click", saveProfile);
els.deleteProfileButton.addEventListener("click", deleteProfile);
els.closePanelButton?.addEventListener("click", closePanel);
els.runAutofillButton.addEventListener("click", runAutofill);
els.dashboardOpenLogButton.addEventListener("click", () => setActiveView("log"));
els.refreshLogsButton.addEventListener("click", () => loadExecutionLogs(true));
els.clearLogsButton.addEventListener("click", clearExecutionLogs);
els.openOptionsButton?.addEventListener("click", () => {
  send("OPEN_OPTIONS").catch((error) => setStatus(error.message));
});
els.connectOutlookButton.addEventListener("click", connectOutlook);
els.refreshOutlookButton.addEventListener("click", () => loadOutlookMessages(true));
document.querySelectorAll("[data-view]").forEach((button) => {
  button.addEventListener("click", () => setActiveView(button.dataset.view));
});

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local" || !changes.autoBidExecutionLogsV1) return;
  const next = changes.autoBidExecutionLogsV1.newValue;
  executionRuns = Array.isArray(next) ? next : [];
  renderExecutionLogs();
  renderDashboard();
});

async function init() {
  try {
    await send("DEV_SESSION");
    clearContextRecoveryMarker();
    await refreshSettings();
    if (settings.devAuthBypass || settings.token) await loadProfiles();
    await loadExecutionLogs(false);
    render();
    setActiveView(activeView);
  } catch (error) {
    setStatus(error.message);
    render();
  }
}

async function refreshSettings() {
  settings = await send("GET_SETTINGS");
  els.apiBase.value = settings.apiBase || "";
  renderSheetSettings();
}

async function saveSettings() {
  await withBusyButton(els.saveSettingsButton, "Saving...", async () => {
    await send("SAVE_SETTINGS", { apiBase: els.apiBase.value });
    await refreshSettings();
    setStatus("Settings saved");
  });
}

async function saveSheetSettings() {
  try {
    await send("SAVE_SHEET_SETTINGS", readSheetSettings());
    await refreshSettings();
    setSheetStatus("Sheet settings saved");
  } catch (error) {
    setSheetStatus(error.message);
  }
}

async function openSheetRows() {
  const buttons = [els.openSheetRowsButton, els.dashboardOpenRowsButton];
  buttons.forEach((button) => {
    button.disabled = true;
    button.dataset.previousText = button.textContent;
    button.textContent = "Opening...";
  });
  setSheetStatus("Reading job URLs from the sheet...");

  try {
    const result = await send("OPEN_SHEET_ROWS", readSheetSettings());
    await refreshSettings();
    const opened = result.opened || 0;
    const failed = result.failed || 0;
    const skipped = result.skipped || 0;
    if (opened) {
      const suffix = failed || skipped ? ` (${failed} failed, ${skipped} skipped)` : "";
      setSheetStatus(`Opened ${opened} job page${opened === 1 ? "" : "s"}${suffix}`);
    } else if (failed || skipped) {
      const firstError = result.failedJobs?.[0]?.error || result.skippedJobs?.[0]?.reason || "URLs could not be opened";
      setSheetStatus(`No tabs opened. ${failed} failed, ${skipped} skipped. ${firstError}`);
    } else {
      setSheetStatus("No job URLs found in that row range");
    }
  } catch (error) {
    setSheetStatus(error.message);
  } finally {
    buttons.forEach((button) => {
      button.disabled = false;
      button.textContent = button.dataset.previousText || "Open";
    });
  }
}

async function runAutofill() {
  await withBusyButton(els.runAutofillButton, "Starting...", async () => {
    setRunState("running", "Running");
    await send("TRIGGER_ACTIVE_AUTOBID");
    setStatus("Autofill started in the active job tab");
    window.setTimeout(() => loadExecutionLogs(false), 800);
  });
}

function setMode(nextMode) {
  mode = nextMode;
  renderAuthMode();
}

function renderAuthMode() {
  const signup = mode === "signup";
  els.loginTab.classList.toggle("active", !signup);
  els.signupTab.classList.toggle("active", signup);
  els.signupNames.classList.toggle("hidden", !signup);
  els.authButton.textContent = signup ? "Create account" : "Login";
}

async function submitAuth() {
  const payload = {
    email: els.email.value.trim(),
    password: els.password.value
  };
  if (mode === "signup") {
    payload.first_name = els.firstName.value.trim();
    payload.last_name = els.lastName.value.trim();
  }

  await withBusyButton(els.authButton, mode === "signup" ? "Creating..." : "Connecting...", async () => {
    await send(mode === "signup" ? "SIGNUP" : "LOGIN", payload);
    await refreshSettings();
    await loadProfiles();
    render();
    setStatus("Connected");
  });
}

async function logout() {
  await send("LOGOUT");
  await send("DEV_SESSION");
  settings = await send("GET_SETTINGS");
  await loadProfiles();
  render();
  setStatus("Development session reset");
}

async function loadProfiles() {
  profiles = await send("LIST_PROFILES");
  currentProfile = profiles.find((profile) => profile.id === settings.selectedProfileId) || profiles[0] || null;
  if (currentProfile && currentProfile.id !== settings.selectedProfileId) {
    settings = await send("SELECT_PROFILE", null, { profileId: currentProfile.id });
  }
}

async function selectProfile() {
  const profileId = els.profileSelect.value;
  settings = await send("SELECT_PROFILE", null, { profileId });
  currentProfile = profiles.find((profile) => profile.id === profileId) || null;
  renderProfileForm();
  renderOutlookConnection();
}

function newProfile() {
  currentProfile = null;
  renderProfileSelect();
  renderProfileForm();
  renderOutlookConnection();
  els.profileName.focus();
}

async function saveProfile() {
  await withBusyButton(els.saveProfileButton, "Saving...", async () => {
    const staticFields = {};
    document.querySelectorAll("[data-static-key]").forEach((input) => {
      const value = input.value.trim();
      if (value) staticFields[input.dataset.staticKey] = value;
    });

    const saved = await send("SAVE_PROFILE", {
      id: currentProfile?.id,
      name: els.profileName.value.trim() || "Default profile",
      static_fields: staticFields,
      resume_text: els.resumeText.value.trim(),
      preferences: {
        tone: els.tone.value.trim(),
        bid_style: els.bidStyle.value.trim()
      }
    });
    settings = await send("SELECT_PROFILE", null, { profileId: saved.id });
    await refreshSettings();
    await loadProfiles();
    currentProfile = profiles.find((profile) => profile.id === saved.id) || saved;
    render();
    setStatus("Profile saved");
  });
}

async function deleteProfile() {
  if (!currentProfile) return;
  await send("DELETE_PROFILE", null, { profileId: currentProfile.id });
  await refreshSettings();
  await loadProfiles();
  render();
  setStatus("Profile deleted");
}

async function loadExecutionLogs(showMessage) {
  try {
    const result = await send("GET_EXECUTION_LOGS", { limit: 50 });
    executionRuns = Array.isArray(result?.runs) ? result.runs : [];
    renderExecutionLogs();
    renderDashboard();
    if (showMessage) setStatus("Execution log refreshed");
  } catch (error) {
    if (showMessage) setStatus(error.message);
  }
}

async function clearExecutionLogs() {
  await send("CLEAR_EXECUTION_LOGS");
  executionRuns = [];
  renderExecutionLogs();
  renderDashboard();
  setStatus("Execution log cleared");
}

function renderExecutionLogs() {
  els.executionLogs.replaceChildren();
  if (executionRuns.length === 0) {
    els.executionLogs.append(createElement("div", "empty-state", "No autofill runs recorded yet."));
    return;
  }

  executionRuns.forEach((run, index) => {
    const details = document.createElement("details");
    details.className = "run-log";
    if (index === 0) details.open = true;

    const summary = document.createElement("summary");
    const dot = createElement("span", `run-status-dot ${run.status || "running"}`);
    const titleWrap = document.createElement("div");
    titleWrap.className = "run-title-wrap";
    titleWrap.append(
      createElement("div", "run-title", run.title || safeHost(run.url) || "Application form"),
      createElement("div", "run-meta", `${run.ats?.name || "Common form"} · ${formatDateTime(run.updated_at)}`)
    );
    const count = createElement("span", "run-meta", `${run.summary?.filled || 0} filled`);
    summary.append(dot, titleWrap, count);
    details.append(summary);

    const events = document.createElement("div");
    events.className = "run-events";
    (run.entries || []).forEach((entry) => {
      const row = document.createElement("div");
      row.className = "run-event";
      row.append(
        createElement("time", "run-meta", formatTime(entry.at)),
        createElement("code", "", `${entry.event}${formatLogData(entry.data)}`)
      );
      events.append(row);
    });
    if (!run.entries?.length) events.append(createElement("div", "run-meta", "Waiting for the first event..."));
    details.append(events);
    els.executionLogs.append(details);
  });
}

function renderDashboard() {
  const latest = executionRuns[0];
  if (!latest) {
    setRunState("idle", "Idle");
    els.latestRunMessage.textContent = "No autofill runs recorded yet.";
    renderLatestSummary(null);
    return;
  }

  setRunState(latest.status || "running", capitalize(latest.status || "running"));
  const host = safeHost(latest.url);
  els.latestRunMessage.textContent = `${latest.ats?.name || "Common form"} on ${host || "current page"}, ${formatDateTime(latest.updated_at)}.`;
  renderLatestSummary(latest);
}

function renderLatestSummary(run) {
  const values = [
    [String(run?.summary?.filled || 0), "Filled"],
    [String(run?.summary?.missed || 0), "Skipped"],
    [run?.ats?.name || "Common", "Form engine"]
  ];
  els.latestRunSummary.replaceChildren(...values.map(([value, label]) => {
    const item = document.createElement("div");
    item.append(createElement("strong", "", value), createElement("span", "", label));
    return item;
  }));
}

function setRunState(status, label) {
  els.dashboardRunState.className = `state-badge ${status}`;
  els.dashboardRunState.textContent = label;
}

async function refreshOutlookConnection(loadMessages = false) {
  setOutlookStatus("Checking connection...");
  try {
    outlookConnection = await send("OUTLOOK_STATUS");
    renderOutlookConnection();
    if (loadMessages && outlookConnection?.connected) await loadOutlookMessages(false);
    else setOutlookStatus(outlookConnection?.configured ? "" : formatOutlookSetupMessage(outlookConnection));
  } catch (error) {
    outlookConnection = { connected: false, configured: false, error: error.message };
    renderOutlookConnection();
    setOutlookStatus(error.message);
  }
}

async function connectOutlook() {
  await withBusyButton(els.connectOutlookButton, "Connecting...", async () => {
    setOutlookStatus("Complete sign-in in the Microsoft window...");
    outlookConnection = await send("OUTLOOK_CONNECT");
    renderOutlookConnection();
    await loadOutlookMessages(false);
    const connected = getOutlookConnections().find((item) => item.profile_id === currentProfile?.id);
    setOutlookStatus(connected
      ? `${connected.email || "Outlook mailbox"} is connected to ${currentProfile?.name || "this profile"}.`
      : "Outlook connected");
  });
}

async function disconnectOutlook(connection, button) {
  const label = connection.email || connection.display_name || "this Outlook mailbox";
  if (!window.confirm(`Disconnect ${label} from ${connection.profile_name || "Auto Bid"}?`)) return;
  button.disabled = true;
  try {
    outlookConnection = await send("OUTLOOK_DISCONNECT", { connection_id: connection.id });
    outlookMessages = outlookMessages.filter((message) => message.connection_id !== connection.id);
    renderOutlookConnection();
    renderOutlookMessages();
    setOutlookStatus(`${label} disconnected`);
  } catch (error) {
    button.disabled = false;
    setOutlookStatus(error.message);
  }
}

async function loadOutlookMessages(showMessage) {
  els.refreshOutlookButton.disabled = true;
  setOutlookStatus("Reading recent Inbox and Junk messages...");
  try {
    const result = await send("OUTLOOK_LIST_VERIFICATION", { top: 25 });
    outlookMessages = Array.isArray(result?.messages) ? result.messages : [];
    renderOutlookMessages();
    setOutlookStatus(outlookMessages.length ? `Found ${outlookMessages.length} likely verification message${outlookMessages.length === 1 ? "" : "s"}.` : "No recent verification messages found.");
    if (showMessage) setStatus("Outlook inbox refreshed");
  } catch (error) {
    setOutlookStatus(error.message);
  } finally {
    els.refreshOutlookButton.disabled = false;
  }
}

function renderOutlookConnection() {
  const connections = getOutlookConnections();
  const connected = connections.length > 0;
  const configured = outlookConnection?.configured !== false;
  els.outlookState.className = `state-badge ${connected ? "connected" : "idle"}`;
  els.outlookState.textContent = connected
    ? `${connections.length} connected`
    : configured ? "Not connected" : "Needs setup";
  const selectedConnection = connections.find((item) => item.profile_id === currentProfile?.id);
  els.connectOutlookButton.classList.remove("hidden");
  els.connectOutlookButton.disabled = !configured || !currentProfile;
  els.connectOutlookButton.textContent = !currentProfile
    ? "Select a profile first"
    : selectedConnection
      ? `Replace Outlook for ${currentProfile.name}`
      : `Connect Outlook for ${currentProfile.name}`;
  els.refreshOutlookButton.classList.toggle("hidden", !connected);
  els.outlookIdentity.classList.toggle("hidden", !connected);
  els.outlookMessagesSection.classList.toggle("hidden", !connected);
  if (connected) {
    els.outlookIdentity.replaceChildren(...connections.map((connection) => {
      const account = document.createElement("article");
      account.className = `connection-account${connection.profile_id === currentProfile?.id ? " selected" : ""}`;
      const identity = document.createElement("div");
      identity.className = "connection-account-identity";
      identity.append(
        createElement("strong", "", connection.display_name || connection.email || "Microsoft account"),
        createElement("span", "", connection.email || ""),
        createElement("span", "connection-profile", connection.profile_name
          ? `Profile: ${connection.profile_name}`
          : "Legacy connection — reconnect it to a profile")
      );
      const disconnectButton = createElement("button", "danger compact-button", "Disconnect");
      disconnectButton.addEventListener("click", () => disconnectOutlook(connection, disconnectButton));
      account.append(identity, disconnectButton);
      return account;
    }));
  } else {
    els.outlookIdentity.replaceChildren();
  }
}

function getOutlookConnections() {
  if (Array.isArray(outlookConnection?.connections)) return outlookConnection.connections;
  return outlookConnection?.connected && outlookConnection?.id ? [outlookConnection] : [];
}

function formatOutlookSetupMessage(connection = {}) {
  const missing = Array.isArray(connection.missing) && connection.missing.length
    ? connection.missing.join(" and ")
    : "MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET";
  const redirect = connection.redirect_uri || "the Chrome identity redirect shown after reloading the extension";
  return `Backend setup required: add ${missing} to .env, register this Web redirect URI in Microsoft Entra, then restart the server:\n${redirect}`;
}

function renderOutlookMessages() {
  els.outlookMessages.replaceChildren();
  if (!outlookConnection?.connected) return;
  if (outlookMessages.length === 0) {
    els.outlookMessages.append(createElement("div", "empty-state", "No recent verification messages."));
    return;
  }

  outlookMessages.forEach((message) => {
    const item = document.createElement("article");
    item.className = "message-item";
    item.append(
      createElement("h3", "", message.subject || "Verification message"),
      createElement("div", "message-meta", [
        message.mailbox_email ? `Mailbox: ${message.mailbox_email}` : "",
        message.from?.name || message.from?.address || "Unknown sender",
        formatDateTime(message.received_at)
      ].filter(Boolean).join(" · "))
    );
    if (message.preview) item.append(createElement("p", "message-meta", message.preview));

    const actions = document.createElement("div");
    actions.className = "row wrap-row";
    (message.codes || []).slice(0, 3).forEach((code) => {
      const button = createElement("button", "verification-code", code);
      button.title = "Copy verification code";
      button.addEventListener("click", async () => {
        await navigator.clipboard.writeText(code);
        setStatus(`Copied ${code}`);
      });
      actions.append(button);
    });
    (message.links || []).slice(0, 2).forEach((url) => {
      const link = document.createElement("a");
      link.href = url;
      link.target = "_blank";
      link.rel = "noreferrer";
      link.textContent = "Open verification link";
      actions.append(link);
    });
    if (message.outlook_url) {
      const outlookLink = document.createElement("a");
      outlookLink.href = message.outlook_url;
      outlookLink.target = "_blank";
      outlookLink.rel = "noreferrer";
      outlookLink.textContent = "Open in Outlook";
      actions.append(outlookLink);
    }
    if (!message.is_read) {
      const readButton = createElement("button", "", "Mark read");
      readButton.addEventListener("click", async () => {
        await send("OUTLOOK_MARK_READ", {
          messageId: message.id,
          connection_id: message.connection_id || ""
        });
        message.is_read = true;
        renderOutlookMessages();
      });
      actions.append(readButton);
    }
    if (actions.childElementCount) item.append(actions);
    els.outlookMessages.append(item);
  });
}

function render() {
  const connected = settings.devAuthBypass || Boolean(settings.token);
  els.sessionLabel.textContent = settings.devAuthBypass
    ? "Development mode"
    : settings.user
      ? settings.user.email
      : "Not connected";
  els.logoutButton.classList.toggle("hidden", !connected);
  els.authSection.classList.toggle("hidden", connected);
  els.profileSection.classList.toggle("hidden", !connected);
  renderAuthMode();
  renderProfileSelect();
  renderProfileForm();
  renderDashboard();
  renderExecutionLogs();
  renderOutlookConnection();
}

function renderProfileSelect() {
  els.profileSelect.innerHTML = "";
  if (profiles.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No profiles yet";
    els.profileSelect.append(option);
    return;
  }

  if (!currentProfile) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "New profile";
    option.selected = true;
    els.profileSelect.append(option);
  }

  for (const profile of profiles) {
    const option = document.createElement("option");
    option.value = profile.id;
    option.textContent = profile.name;
    option.selected = profile.id === currentProfile?.id;
    els.profileSelect.append(option);
  }
}

function renderProfileForm() {
  const profile = currentProfile || { name: "", static_fields: {}, resume_text: "", preferences: {} };
  els.profileName.value = profile.name || "";
  els.resumeText.value = profile.resume_text || "";
  els.tone.value = profile.preferences?.tone || "";
  els.bidStyle.value = profile.preferences?.bid_style || "";
  document.querySelectorAll("[data-static-key]").forEach((input) => {
    input.value = profile.static_fields?.[input.dataset.staticKey] || "";
  });
  els.deleteProfileButton.disabled = !currentProfile;
}

function renderSheetSettings() {
  const sheet = settings.sheetSettings || {};
  els.spreadsheetId.value = sheet.spreadsheetId || "";
  els.sheetName.value = sheet.sheetName || "";
  els.sheetStartRow.value = sheet.startRow || 2;
  els.sheetEndRow.value = sheet.endRow || sheet.startRow || 2;
  els.dashboardSheetName.textContent = sheet.sheetName ? `Sheet: ${sheet.sheetName}` : "No sheet selected";
  els.dashboardSheetRows.textContent = `Rows ${sheet.startRow || 2}-${sheet.endRow || sheet.startRow || 2}`;
}

function readSheetSettings() {
  return {
    spreadsheetId: els.spreadsheetId.value.trim(),
    sheetName: els.sheetName.value.trim(),
    startRow: Number(els.sheetStartRow.value || 2),
    endRow: Number(els.sheetEndRow.value || els.sheetStartRow.value || 2)
  };
}

function setActiveView(view) {
  activeView = view || "dashboard";
  document.querySelectorAll("[data-view]").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === activeView);
  });
  document.querySelectorAll("[data-view-panel]").forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.viewPanel === activeView);
  });
  if (activeView === "log") loadExecutionLogs(false);
  if (activeView === "outlook") refreshOutlookConnection(true);
}

function setStatus(message) {
  els.status.textContent = message;
  window.setTimeout(() => {
    if (els.status.textContent === message) els.status.textContent = "";
  }, 3500);
}

function setSheetStatus(message) {
  els.sheetStatus.textContent = message;
  setStatus(message);
}

function setOutlookStatus(message) {
  els.outlookStatus.textContent = message || "";
}

function closePanel() {
  window.parent?.postMessage({ source: "auto-bid-panel", type: "CLOSE_PANEL" }, "*");
}

async function withBusyButton(button, busyText, operation) {
  const previousText = button.textContent;
  button.disabled = true;
  button.textContent = busyText;
  try {
    return await operation();
  } catch (error) {
    setStatus(error.message);
    throw error;
  } finally {
    button.disabled = false;
    button.textContent = previousText;
  }
}

function createElement(tag, className = "", text = "") {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined && text !== null) element.textContent = text;
  return element;
}

function safeHost(url) {
  try {
    return new URL(url).hostname.replace(/^www\./, "");
  } catch (_error) {
    return "";
  }
}

function formatDateTime(value) {
  const date = new Date(value || 0);
  if (!Number.isFinite(date.getTime())) return "unknown time";
  return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" }).format(date);
}

function formatTime(value) {
  const date = new Date(value || 0);
  if (!Number.isFinite(date.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, { hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(date);
}

function formatLogData(value) {
  if (!value || (typeof value === "object" && Object.keys(value).length === 0)) return "";
  try {
    const text = JSON.stringify(value);
    return ` ${text.length > 1200 ? `${text.slice(0, 1197)}...` : text}`;
  } catch (_error) {
    return "";
  }
}

function capitalize(value) {
  const text = String(value || "");
  return text ? `${text[0].toUpperCase()}${text.slice(1)}` : "";
}

function send(type, payload, extra = {}) {
  return new Promise((resolve, reject) => {
    try {
      chrome.runtime.sendMessage({ type, payload, ...extra }, (response) => {
        if (chrome.runtime.lastError) {
          reject(createRuntimeMessageError(chrome.runtime.lastError.message));
          return;
        }
        if (!response?.ok) {
          reject(new Error(response?.error || "Request failed"));
          return;
        }
        resolve(response.data);
      });
    } catch (error) {
      reject(createRuntimeMessageError(error?.message || String(error)));
    }
  });
}

function createRuntimeMessageError(message) {
  const text = String(message || "Extension request failed");
  if (!/extension context invalidated/i.test(text)) return new Error(text);

  const reloadScheduled = scheduleContextRecovery();
  return new Error(reloadScheduled
    ? "Auto Bid was reloaded. Refreshing this panel..."
    : "Auto Bid was reloaded. Close and reopen this panel.");
}

function scheduleContextRecovery() {
  if (contextRecoveryStarted) return true;
  contextRecoveryStarted = true;

  let previousRecoveryAt = 0;
  try {
    previousRecoveryAt = Number(sessionStorage.getItem(CONTEXT_RECOVERY_STORAGE_KEY) || 0);
  } catch (_error) {
    // A reload is still safe when session storage is unavailable.
  }

  const now = Date.now();
  if (now - previousRecoveryAt < CONTEXT_RECOVERY_COOLDOWN_MS) return false;

  try {
    sessionStorage.setItem(CONTEXT_RECOVERY_STORAGE_KEY, String(now));
  } catch (_error) {
    // The reload itself is the recovery mechanism.
  }

  window.setTimeout(() => {
    try {
      location.reload();
    } catch (_error) {
      window.parent?.postMessage({ source: "auto-bid-panel", type: "RELOAD_PANEL" }, "*");
    }
  }, 80);
  return true;
}

function clearContextRecoveryMarker() {
  contextRecoveryStarted = false;
  try {
    sessionStorage.removeItem(CONTEXT_RECOVERY_STORAGE_KEY);
  } catch (_error) {
    // Nothing else is required after a successful runtime request.
  }
}
