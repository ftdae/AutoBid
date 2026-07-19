let mode = "login";
let settings = {};
let profiles = [];
let currentProfile = null;

const els = {
  sessionLabel: document.getElementById("sessionLabel"),
  logoutButton: document.getElementById("logoutButton"),
  apiBase: document.getElementById("apiBase"),
  saveSettingsButton: document.getElementById("saveSettingsButton"),
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
  status: document.getElementById("status")
};

document.addEventListener("DOMContentLoaded", init);
els.saveSettingsButton.addEventListener("click", saveSettings);
els.loginTab.addEventListener("click", () => setMode("login"));
els.signupTab.addEventListener("click", () => setMode("signup"));
els.authButton.addEventListener("click", submitAuth);
els.logoutButton.addEventListener("click", logout);
els.profileSelect.addEventListener("change", selectProfile);
els.newProfileButton.addEventListener("click", newProfile);
els.saveProfileButton.addEventListener("click", saveProfile);
els.deleteProfileButton.addEventListener("click", deleteProfile);

async function init() {
  await send("DEV_SESSION");
  await refreshSettings();
  if (settings.devAuthBypass || settings.token) await loadProfiles();
  render();
}

async function refreshSettings() {
  settings = await send("GET_SETTINGS");
  els.apiBase.value = settings.apiBase || "";
}

async function saveSettings() {
  await send("SAVE_SETTINGS", { apiBase: els.apiBase.value });
  await refreshSettings();
  setStatus("Settings saved");
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

  await send(mode === "signup" ? "SIGNUP" : "LOGIN", payload);
  await refreshSettings();
  await loadProfiles();
  render();
  setStatus("Connected");
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
}

function newProfile() {
  currentProfile = null;
  renderProfileForm();
}

async function saveProfile() {
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

  await refreshSettings();
  await loadProfiles();
  currentProfile = profiles.find((profile) => profile.id === saved.id) || saved;
  render();
  setStatus("Profile saved");
}

async function deleteProfile() {
  if (!currentProfile) return;
  await send("DELETE_PROFILE", null, { profileId: currentProfile.id });
  await refreshSettings();
  await loadProfiles();
  render();
  setStatus("Profile deleted");
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

function setStatus(message) {
  els.status.textContent = message;
  window.setTimeout(() => {
    if (els.status.textContent === message) els.status.textContent = "";
  }, 2800);
}

function send(type, payload, extra = {}) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type, payload, ...extra }, (response) => {
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
