import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { extractVerificationCodes } from "../server/outlook/microsoft-graph.js";

const backgroundSource = readFileSync(new URL("../extension/background.js", import.meta.url), "utf8");
const contentSource = readFileSync(new URL("../extension/content.js", import.meta.url), "utf8");
const outlookSource = readFileSync(new URL("../server/outlook/microsoft-graph.js", import.meta.url), "utf8");
const popupSource = readFileSync(new URL("../extension/popup.js", import.meta.url), "utf8");
const schemaSource = readFileSync(new URL("../server/db/schema.js", import.meta.url), "utf8");

test("Outlook code extraction requires verification context", () => {
  assert.deepEqual(extractVerificationCodes("Your Greenhouse verification code is 482913."), ["482913"]);
  assert.deepEqual(extractVerificationCodes("Use confirmation code AB-1234 to continue."), ["AB1234"]);
  assert.deepEqual(extractVerificationCodes("Application 123456 was received."), []);
});

test("Outlook connection diagnostics expose missing backend setup and the exact redirect URI", () => {
  assert.match(outlookSource, /getOutlookConfiguration/);
  assert.match(outlookSource, /missing\.push\("MICROSOFT_CLIENT_ID"\)/);
  assert.match(outlookSource, /missing\.push\("MICROSOFT_CLIENT_SECRET"\)/);
  assert.match(backgroundSource, /redirect_uri: chrome\.identity\.getRedirectURL\("outlook"\)/);
  assert.match(popupSource, /register this Web redirect URI in Microsoft Entra/);
});

test("verification email lookup is bounded to the submission and page context", () => {
  assert.match(outlookSource, /requestedSince \? requestedSince - 30_000/);
  assert.match(outlookSource, /buildVerificationContext\(options\)/);
  assert.match(outlookSource, /contextScore: scoreVerificationContext/);
  assert.match(backgroundSource, /page_url: String\(payload\.page_url/);
  assert.match(backgroundSource, /since: normalizeOutlookMonitorSince/);
  assert.match(backgroundSource, /Ambiguous verification email was not assigned/);
});

test("verification monitoring survives background tabs and caps delivery attempts", () => {
  assert.match(backgroundSource, /OUTLOOK_VERIFICATION_MONITORS_STORAGE_KEY/);
  assert.match(backgroundSource, /OUTLOOK_VERIFICATION_ALARM/);
  assert.match(backgroundSource, /ensureOutlookVerificationMonitorPump/);
  assert.match(backgroundSource, /outlookVerificationMonitorScanPromise/);
  assert.match(backgroundSource, /Number\(monitor\.delivery_attempts \|\| 0\) < 3/);
  assert.match(backgroundSource, /AUTO_BID_OUTLOOK_VERIFICATION_READY/);
});

test("Outlook mailboxes are profile-bound and verification messages stay in their mailbox", () => {
  assert.match(schemaSource, /auto_bid_outlook_mailboxes/);
  assert.match(outlookSource, /profile_id: normalizeIdentifier\(profileId\)/);
  assert.match(outlookSource, /connection_id: connection\.id/);
  assert.match(outlookSource, /mailbox_email: connection\.email/);
  assert.match(backgroundSource, /outlookMessageMatchesMonitorMailbox/);
  assert.match(contentSource, /profile_id: activeAutoBidProfileId/);
  assert.match(contentSource, /mailbox_email: activeAutoBidProfileEmail/);
  assert.match(popupSource, /getOutlookConnections/);
});

test("a delivered code is filled and the verification action is clicked only once", () => {
  assert.match(contentSource, /completedOutlookVerificationMessages/);
  assert.match(contentSource, /outlook-verification:manual-value-preserved/);
  assert.match(contentSource, /completeOutlookVerification\(payload\)/);
  assert.match(contentSource, /applyOutlookVerificationCode/);
  assert.match(contentSource, /waitForOutlookVerificationSubmitButton/);
  assert.match(contentSource, /resend\|send again\|new code/);
  assert.match(contentSource, /The code was filled and the verification button was clicked once/);
});

test("the background monitor delivers one newly received code to its originating tab", async () => {
  const storage = {};
  const delivered = [];
  const listeners = [];
  const event = () => ({ addListener(listener) { listeners.push(listener); } });
  const runtimeOnMessage = { listeners: [], addListener(listener) { this.listeners.push(listener); } };
  const clone = (value) => value === undefined ? undefined : structuredClone(value);
  globalThis.chrome = {
    runtime: {
      lastError: null,
      onInstalled: event(),
      onStartup: event(),
      onMessage: runtimeOnMessage,
      openOptionsPage: async () => {}
    },
    commands: { onCommand: event() },
    action: { onClicked: event() },
    windows: { onRemoved: event(), getLastFocused: async () => null, getCurrent: async () => null },
    identity: {
      getRedirectURL: () => "https://autobid-test.chromiumapp.org/outlook",
      launchWebAuthFlow: async () => ""
    },
    alarms: { onAlarm: event(), create: async () => {}, clear: async () => true },
    tabs: {
      onActivated: event(),
      onUpdated: event(),
      onRemoved: event(),
      async update(tabId, changes) { return { id: tabId, ...changes }; },
      async sendMessage(tabId, message) {
        if (message.type === "AUTO_BID_OUTLOOK_VERIFICATION_READY") {
          delivered.push({ tabId, message: clone(message) });
          return { ok: true, applied: true, clicked: true, settled: true };
        }
        return { ok: true };
      },
      async query() { return []; }
    },
    scripting: { executeScript: async () => [] },
    storage: {
      local: {
        async get(keys) {
          const list = Array.isArray(keys) ? keys : [keys];
          return Object.fromEntries(list.filter((key) => key in storage).map((key) => [key, clone(storage[key])]));
        },
        async set(values) { Object.assign(storage, clone(values)); },
        async remove(keys) { for (const key of Array.isArray(keys) ? keys : [keys]) delete storage[key]; }
      }
    }
  };

  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (url) => {
    const path = new URL(String(url)).pathname;
    let data;
    if (path.endsWith("/auth/dev-session")) {
      data = {
        token: "test-token",
        user: { id: "abu_dev_local", email: "test@example.com" },
        profile: { id: "abp_dev_default", static_fields: {} }
      };
    } else if (path.endsWith("/outlook/connection")) {
      data = {
        connected: true,
        configured: true,
        missing: [],
        connections: [{
          id: "connection-1",
          profile_id: "abp_dev_default",
          profile_name: "Development profile",
          email: "test@example.com"
        }, {
          id: "connection-2",
          profile_id: "abp_other",
          profile_name: "Other profile",
          email: "other@example.com"
        }]
      };
    } else if (path.endsWith("/outlook/messages")) {
      data = {
        messages: [{
          id: "wrong-mailbox-message",
          connection_id: "connection-2",
          mailbox_email: "other@example.com",
          profile_id: "abp_other",
          subject: "Greenhouse verification code",
          preview: "Confirm your application",
          from: { name: "Greenhouse", address: "no-reply@greenhouse.io" },
          received_at: new Date(Date.now() + 1000).toISOString(),
          codes: ["999999"]
        }, {
          id: "message-1",
          connection_id: "connection-1",
          mailbox_email: "test@example.com",
          profile_id: "abp_dev_default",
          subject: "Greenhouse verification code",
          preview: "Confirm your application",
          from: { name: "Greenhouse", address: "no-reply@greenhouse.io" },
          received_at: new Date().toISOString(),
          codes: ["482913"]
        }]
      };
    } else if (/\/outlook\/messages\/message-1\/read$/.test(path)) {
      data = { id: "message-1", is_read: true };
    } else {
      throw new Error(`Unexpected test request: ${url}`);
    }
    return { ok: true, status: 200, json: async () => ({ data, errors: null }) };
  };

  try {
    await import(`../extension/background.js?outlook-monitor=${Date.now()}`);
    const listener = runtimeOnMessage.listeners[0];
    const send = (message) => new Promise((resolve, reject) => {
      listener(message, {
        tab: { id: 42, url: "https://job-boards.greenhouse.io/acme/jobs/123", title: "Backend Engineer - Acme" },
        frameId: 0
      }, (response) => response?.ok ? resolve(response.data) : reject(new Error(response?.error || "message failed")));
    });

    const started = await send({
      type: "OUTLOOK_MONITOR_START",
      payload: {
        since: new Date(Date.now() - 1000).toISOString(),
        page_url: "https://job-boards.greenhouse.io/acme/jobs/123",
        title: "Backend Engineer - Acme"
      }
    });
    assert.equal(started.started, true);
    const armed = await send({ type: "OUTLOOK_MONITOR_ARM", payload: { monitor_id: started.monitor_id } });
    assert.equal(armed.armed, true);

    const deadline = Date.now() + 2000;
    while (delivered.length === 0 && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 10));
    }
    assert.equal(delivered.length, 1);
    assert.equal(delivered[0].tabId, 42);
    assert.equal(delivered[0].message.payload.message.id, "message-1");
    assert.equal(delivered[0].message.payload.message.connection_id, "connection-1");
    assert.equal(delivered[0].message.payload.message.mailbox_email, "test@example.com");
    assert.equal((storage.autoBidOutlookVerificationMonitorsV1 || []).length, 0);
  } finally {
    globalThis.fetch = originalFetch;
    delete globalThis.chrome;
  }
});
