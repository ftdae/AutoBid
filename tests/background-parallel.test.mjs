import assert from "node:assert/strict";
import test from "node:test";

function createEvent() {
  const listeners = [];
  return {
    listeners,
    addListener(listener) {
      listeners.push(listener);
    },
    removeListener(listener) {
      const index = listeners.indexOf(listener);
      if (index >= 0) listeners.splice(index, 1);
    }
  };
}

function installChromeMock() {
  const storage = {};
  const tabs = new Map();
  const runningWorkerTabs = new Set();
  const workerRunMessages = [];
  let nextTabId = 1000;

  const clone = (value) => value === undefined ? undefined : structuredClone(value);
  const runtimeOnMessage = createEvent();
  const tabsOnRemoved = createEvent();

  globalThis.chrome = {
    runtime: {
      lastError: null,
      onInstalled: createEvent(),
      onMessage: runtimeOnMessage,
      openOptionsPage: async () => {}
    },
    commands: { onCommand: createEvent() },
    action: { onClicked: createEvent() },
    windows: {
      onRemoved: createEvent(),
      getLastFocused: async () => null,
      getCurrent: async () => null
    },
    tabs: {
      onActivated: createEvent(),
      onUpdated: createEvent(),
      onRemoved: tabsOnRemoved,
      async create({ url }) {
        const tab = { id: nextTabId++, url, status: "complete", active: false };
        tabs.set(tab.id, tab);
        return clone(tab);
      },
      async get(tabId) {
        const tab = tabs.get(tabId);
        if (!tab) throw new Error("No tab");
        return clone(tab);
      },
      async update(tabId, changes) {
        const tab = tabs.get(tabId);
        if (!tab) throw new Error("No tab");
        Object.assign(tab, changes);
        return clone(tab);
      },
      async remove(tabId) {
        tabs.delete(tabId);
        runningWorkerTabs.delete(tabId);
      },
      async sendMessage(tabId, message) {
        if (!tabs.has(tabId)) throw new Error("No tab");
        if (message?.type === "AUTOBID_GPT_PING") {
          return { message: "ready", runningBatch: runningWorkerTabs.has(tabId) };
        }
        if (message?.type === "AUTOBID_GPT_RUN_BATCH") {
          runningWorkerTabs.add(tabId);
          workerRunMessages.push({ tabId, ...clone(message) });
          return { message: "started", batchId: message.batchId, started: true };
        }
        return { ok: true };
      },
      async query() {
        return [];
      }
    },
    storage: {
      local: {
        async get(keys) {
          const list = Array.isArray(keys) ? keys : [keys];
          return Object.fromEntries(list.filter((key) => key in storage).map((key) => [key, clone(storage[key])]));
        },
        async set(values) {
          for (const [key, value] of Object.entries(values)) storage[key] = clone(value);
        },
        async remove(keys) {
          for (const key of Array.isArray(keys) ? keys : [keys]) delete storage[key];
        }
      }
    },
    scripting: {
      executeScript: async () => []
    }
  };

  return { storage, runtimeOnMessage, tabs, runningWorkerTabs, workerRunMessages };
}

function waitFor(predicate, timeoutMs = 5000) {
  const started = Date.now();
  return new Promise((resolve, reject) => {
    const check = () => {
      if (predicate()) return resolve();
      if (Date.now() - started >= timeoutMs) return reject(new Error("Timed out waiting for parallel workers"));
      setTimeout(check, 10);
    };
    check();
  });
}

test("the first autofill hotkey prewarms one idle GPT tab without sending an empty prompt", async () => {
  const mock = installChromeMock();
  mock.tabs.set(77, {
    id: 77,
    url: "https://jobs.example.test/first",
    status: "complete",
    active: true
  });
  await import(`../extension/background.js?first-hotkey-prewarm=${Date.now()}`);
  const listener = mock.runtimeOnMessage.listeners[0];
  const send = (message, sender) => new Promise((resolve, reject) => {
    listener(message, sender, (response) => {
      if (response?.ok) resolve(response.data);
      else reject(new Error(response?.error || "Background request failed"));
    });
  });

  const result = await send({ type: "HOTKEY_TRIGGER" }, { tab: { id: 77 }, frameId: 0 });
  assert.equal(result.triggered, true);
  assert.equal(result.gptWorker.reason, "first-hotkey-prewarm");

  await waitFor(() => {
    const states = mock.storage.autoBidGptBatchStatesV2 || [];
    return states.length === 1 && states[0].status === "idle";
  });

  const [worker] = mock.storage.autoBidGptBatchStatesV2;
  assert.notEqual(worker.tab_id, 77);
  assert.equal(mock.tabs.size, 2);
  assert.equal(mock.workerRunMessages.filter((message) => message.tabId === worker.tab_id).length, 0);
  assert.equal(mock.tabs.get(worker.tab_id)?.autoDiscardable, false);
});

test("50 runtime GPT requests use at most three persistent tabs and each tab accepts more work", async () => {
  const mock = installChromeMock();
  await import(`../extension/background.js?parallel-test=${Date.now()}`);
  const listener = mock.runtimeOnMessage.listeners[0];
  assert.equal(typeof listener, "function");

  const send = (message, sender) => new Promise((resolve, reject) => {
    listener(message, sender, (response) => {
      if (response?.ok) resolve(response.data);
      else reject(new Error(response?.error || "Background request failed"));
    });
  });

  const queued = await Promise.all(Array.from({ length: 50 }, (_, index) => send({
    type: "GPT_ANSWER_REQUEST",
    payload: {
      context: {},
      page: { url: `https://jobs.example.test/${index}` },
      payload: { fields: [{ field_id: `field_${index}`, required: true }] },
      client_run_id: `run_${index}`,
      timeout_ms: 90000
    }
  }, { tab: { id: index + 1 }, frameId: 0 })));

  assert.equal(new Set(queued.map((request) => request.request_id)).size, 50);
  await waitFor(() => {
    const requests = mock.storage.autoBidRuntimeGptQueueV2?.requests || [];
    const states = mock.storage.autoBidGptBatchStatesV2 || [];
    return requests.length === 50 &&
      requests.filter((request) => request.status === "processing").length === 3 &&
      requests.filter((request) => request.status === "pending").length === 47 &&
      states.length === 3 &&
      states.every((state) => state.request_ids.length === 1);
  });

  const initialRequests = mock.storage.autoBidRuntimeGptQueueV2.requests;
  const initialStates = mock.storage.autoBidGptBatchStatesV2;
  assert.equal(initialRequests.find((request) => request.payload.fields[0].field_id === "field_0")?.client_run_id, "run_0");
  assert.equal(initialStates.length, 3);
  assert.equal(mock.tabs.size, 3);
  assert.equal(new Set(initialRequests.filter((request) => request.status === "processing").map((request) => request.batch_id)).size, 3);
  for (const state of initialStates) {
    assert.equal(mock.tabs.get(state.tab_id)?.autoDiscardable, false);
  }

  const reusedWorker = structuredClone(initialStates[0]);
  for (const requestId of reusedWorker.request_ids) {
    const request = initialRequests.find((item) => item.id === requestId);
    await send({
      type: "AUTOBID_GPT_SAVE_REQUEST_ANSWERS",
      payload: {
        request_id: requestId,
        batch_id: reusedWorker.batch_id,
        answers: [{ field_id: request.payload.fields[0].field_id, value: "answer" }]
      }
    });
  }

  mock.runningWorkerTabs.delete(reusedWorker.tab_id);
  await send({
    type: "AUTOBID_GPT_BATCH_COMPLETE",
    payload: {
      batch_id: reusedWorker.batch_id,
      request_ids: reusedWorker.request_ids,
      saved_request_ids: reusedWorker.request_ids
    }
  }, { tab: { id: reusedWorker.tab_id } });
  await send({
    type: "AUTOBID_GPT_WORKER_READY",
    payload: { batch_id: reusedWorker.batch_id }
  }, { tab: { id: reusedWorker.tab_id } });

  await waitFor(() => {
    const requests = mock.storage.autoBidRuntimeGptQueueV2?.requests || [];
    const states = mock.storage.autoBidGptBatchStatesV2 || [];
    const state = states.find((item) => item.batch_id === reusedWorker.batch_id);
    return states.length === 3 &&
      state?.request_ids?.length === 1 &&
      !state.request_ids.some((requestId) => reusedWorker.request_ids.includes(requestId)) &&
      requests.filter((request) => request.status === "complete").length === 1 &&
      requests.filter((request) => request.status === "processing").length === 3;
  });

  assert.equal(mock.tabs.size, 3);
  assert.ok(mock.workerRunMessages.filter((message) => message.tabId === reusedWorker.tab_id).length >= 2);
});

test("an autofill ChatGPT request becomes terminal after one failed attempt so OpenAI can continue", async () => {
  const mock = installChromeMock();
  await import(`../extension/background.js?single-chatgpt-attempt=${Date.now()}`);
  const listener = mock.runtimeOnMessage.listeners[0];
  const send = (message, sender = {}) => new Promise((resolve, reject) => {
    listener(message, sender, (response) => {
      if (response?.ok) resolve(response.data);
      else reject(new Error(response?.error || "Background request failed"));
    });
  });

  const queued = await send({
    type: "GPT_ANSWER_REQUEST",
    payload: {
      context: {},
      page: { url: "https://jobs.example.test/fallback" },
      payload: { fields: [{ field_id: "required_answer", required: true }] },
      timeout_ms: 90000,
      max_attempts: 1
    }
  }, { tab: { id: 41 }, frameId: 0 });

  await waitFor(() => {
    const request = mock.storage.autoBidRuntimeGptQueueV2?.requests?.find((item) => item.id === queued.request_id);
    return request?.status === "processing";
  });
  const processing = mock.storage.autoBidRuntimeGptQueueV2.requests.find((item) => item.id === queued.request_id);

  await send({
    type: "AUTOBID_GPT_FAIL_REQUEST",
    payload: {
      request_id: processing.id,
      batch_id: processing.batch_id,
      error: "ChatGPT omitted the field"
    }
  });

  const status = await send({
    type: "GPT_ANSWER_STATUS",
    payload: { request_id: processing.id }
  });
  assert.equal(status.status, "error");
  assert.equal(status.terminal_error, true);
  assert.equal(status.attempt_count, 1);
});
