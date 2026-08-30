const els = {
  stateLabel: document.getElementById("stateLabel"),
  appsScriptUrl: document.getElementById("appsScriptUrl"),
  secret: document.getElementById("secret"),
  sheetName: document.getElementById("sheetName"),
  startRow: document.getElementById("startRow"),
  endRow: document.getElementById("endRow"),
  maxRowsPerRun: document.getElementById("maxRowsPerRun"),
  saveButton: document.getElementById("saveButton"),
  runOnceButton: document.getElementById("runOnceButton"),
  startButton: document.getElementById("startButton"),
  stopButton: document.getElementById("stopButton"),
  status: document.getElementById("status")
};

document.addEventListener("DOMContentLoaded", init);
els.saveButton.addEventListener("click", saveSettings);
els.runOnceButton.addEventListener("click", () => command("RUN_ONCE"));
els.startButton.addEventListener("click", () => command("START"));
els.stopButton.addEventListener("click", () => command("STOP"));

async function init() {
  const settings = await send("GET_SETTINGS");
  render(settings);
}

function render(settings) {
  els.stateLabel.textContent = settings.running ? "Loop running" : "Idle";
  els.appsScriptUrl.value = settings.appsScriptUrl || "";
  els.secret.value = settings.secret || "";
  els.sheetName.value = settings.sheetName || "";
  els.startRow.value = settings.startRow || 2;
  els.endRow.value = settings.endRow || settings.startRow || 2;
  els.maxRowsPerRun.value = settings.maxRowsPerRun || 10;
}

async function saveSettings() {
  const settings = readSettings();
  const saved = await send("SAVE_SETTINGS", settings);
  render(saved);
  setStatus("Saved");
}

async function command(type) {
  await saveSettings();
  const result = await send(type);
  setStatus(result?.message || "Sent to ChatGPT tab");
  const settings = await send("GET_SETTINGS");
  render(settings);
}

function readSettings() {
  const startRow = Math.max(2, Number(els.startRow.value || 2));
  const endRow = Math.max(startRow, Number(els.endRow.value || startRow));
  return {
    appsScriptUrl: els.appsScriptUrl.value.trim(),
    secret: els.secret.value.trim(),
    sheetName: els.sheetName.value.trim(),
    startRow,
    endRow,
    maxRowsPerRun: Math.max(1, Math.min(50, Number(els.maxRowsPerRun.value || 10)))
  };
}

function setStatus(message) {
  els.status.textContent = message;
  window.setTimeout(() => {
    if (els.status.textContent === message) els.status.textContent = "";
  }, 3500);
}

function send(type, payload) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type, payload }, (response) => {
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
  }).catch((error) => {
    setStatus(error.message);
    throw error;
  });
}
