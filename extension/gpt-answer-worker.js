(() => {
  const WORKER_BUILD_ID = "2026-09-01-persistent-chat-three-workers-v6";
  if (window.__autoBidGptAnswerWorkerBuildId === WORKER_BUILD_ID) return;
  window.__autoBidGptAnswerWorkerBuildId = WORKER_BUILD_ID;
  window.__autoBidGptAnswerWorkerLoaded = true;

  const STATUS_ID = "autobid-gpt-answer-status";
  const RESPONSE_TIMEOUT_MS = 120000;
  const RESPONSE_NO_TEXT_TIMEOUT_MS = 30000;
  const RESPONSE_STALL_TIMEOUT_MS = 30000;
  const RESPONSE_BACKGROUND_NO_TEXT_TIMEOUT_MS = 90000;
  const RESPONSE_BACKGROUND_STALL_TIMEOUT_MS = 90000;
  const RESPONSE_STABLE_MS = 2500;
  const RESPONSE_JSON_STABLE_MS = 900;
  const RESPONSE_POLL_MS = 500;
  const SAVE_ANSWERS_TIMEOUT_MS = 5000;
  const SAVE_ANSWERS_RETRIES = 2;
  const SAVE_ANSWERS_RETRY_MS = 1000;
  const LOOP_DELAY_MS = 5000;
  const ROW_DELAY_MS = 1500;
  const ROW_FAILURE_COOLDOWN_MS = 5000;
  const ROW_COMPLETION_COOLDOWN_MS = 30000;
  const RUN_ONCE_MAX_ROWS = 200;
  const ROW_CHAT_ATTEMPTS = 2;
  // One application per prompt keeps result parsing deterministic. Three
  // persistent browser workers run those prompts in parallel in their existing chats.
  const MAX_REQUESTS_PER_PROMPT = 1;

  let running = false;
  let stopRequested = false;
  let loopPromise = null;
  let oneShotPromise = null;
  let batchPromise = null;
  let currentBatchId = "";
  let runAgainRequested = false;
  const recentFailedRows = new Map();
  const recentCompletedRows = new Map();

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type === "AUTOBID_GPT_PING") {
      sendResponse({ message: "ready", runningBatch: Boolean(batchPromise), batchId: currentBatchId });
      return false;
    }

    if (message?.type === "AUTOBID_GPT_RUN_ONCE") {
      stopRequested = false;
      startRunOnce();
      sendResponse({ message: "GPT answer worker started" });
      return false;
    }

    if (message?.type === "AUTOBID_GPT_RUN_BATCH") {
      stopRequested = false;
      const started = startRuntimeBatch(message);
      sendResponse({
        message: started ? "GPT batch worker started" : "GPT batch worker is busy",
        batchId: message.batchId || "",
        started,
        busy: !started
      });
      return false;
    }

    if (message?.type === "AUTOBID_GPT_START") {
      running = true;
      stopRequested = false;
      send("AUTOBID_GPT_WORKER_STATE", { running: true }).catch(() => {});
      if (!loopPromise) loopPromise = runLoop().finally(() => { loopPromise = null; });
      sendResponse({ message: "GPT answer worker started" });
      return false;
    }

    if (message?.type === "AUTOBID_GPT_STOP") {
      requestStop();
      sendResponse({ message: "Stopped" });
      return false;
    }

    return false;
  });

  function startRuntimeBatch(message) {
    if (batchPromise) {
      showStatus("This ChatGPT tab is already processing its batch.", "success");
      return false;
    }

    currentBatchId = String(message.batchId || message.batch_id || "");
    batchPromise = processRuntimeBatch(message)
      .catch((error) => {
        if (!isStopError(error)) {
          console.error("[AutoBid GPT Batch Worker]", error);
          showStatus(error.message || String(error), "error");
        }
      })
      .finally(() => {
        const completedBatchId = currentBatchId;
        batchPromise = null;
        currentBatchId = "";
        send("AUTOBID_GPT_WORKER_READY", { batch_id: completedBatchId }).catch((error) => {
          console.warn("[AutoBid GPT Batch Worker] Could not report worker readiness", error);
        });
      });
    return true;
  }

  async function processRuntimeBatch(message = {}) {
    const batchId = String(message.batchId || message.batch_id || "");
    const batchSize = Math.min(
      MAX_REQUESTS_PER_PROMPT,
      Math.max(1, Number(message.batchSize || MAX_REQUESTS_PER_PROMPT))
    );
    const collectionDelayMs = Math.max(0, Number(message.collectionDelayMs || 0));
    const allowSheetFallback = Boolean(message.allowSheetFallback);
    let rows = [];
    let source = "runtime";
    const savedRequestIds = [];
    const failedRequestIds = [];
    let batchError = "";

    try {
      if (collectionDelayMs) {
        showStatus(`Collecting up to ${batchSize} AutoBid requests...`, "success");
        await sleepInterruptible(collectionDelayMs);
      }

      const runtimeData = await send("AUTOBID_GPT_FETCH_PENDING_REQUESTS", {
        batch_id: batchId,
        limit: batchSize
      });
      rows = Array.isArray(runtimeData?.requests) ? runtimeData.requests : [];

      if (rows.length === 0 && allowSheetFallback) {
        const sheetData = await send("AUTOBID_GPT_FETCH_PENDING_ROWS").catch(() => ({ rows: [] }));
        rows = (Array.isArray(sheetData?.rows) ? sheetData.rows : [])
          .filter((row) => row?.questions?.fields?.length)
          .slice(0, batchSize);
        source = "sheet";
      }

      if (rows.length === 0) {
        showStatus("No pending AutoBid answer requests found.", "success");
        return;
      }

      showStatus(`Answering a batch of ${rows.length} job application${rows.length === 1 ? "" : "s"}...`, "success");
      const responseText = await askChatGptBatch(buildBatchAnswerPrompt(rows), rows);
      const results = extractBatchAnswers(responseText, rows);

      await Promise.all(rows.map(async (row) => {
        const requestId = getBatchRequestId(row);
        const result = results.get(requestId);
        if (!result?.answers?.length) {
          const error = result?.error || new Error(`ChatGPT omitted ${getRowLabel(row)} from the batch response.`);
          if (isRuntimeRequest(row)) {
            await markRuntimeRequestFailed(row, error, batchId).catch(() => {});
            failedRequestIds.push(requestId);
          }
          return;
        }

        try {
          await saveAnswersWithRetry(row, result.answers, batchId);
          savedRequestIds.push(requestId);
        } catch (error) {
          if (isRuntimeRequest(row)) {
            await markRuntimeRequestFailed(row, error, batchId).catch(() => {});
            failedRequestIds.push(requestId);
          }
        }
      }));

      showStatus(
        `Stored ${savedRequestIds.length} of ${rows.length} batched AutoBid result${rows.length === 1 ? "" : "s"}.`,
        failedRequestIds.length ? "error" : "success"
      );
    } catch (error) {
      if (!isStopError(error)) {
        batchError = error.message || String(error);
        await Promise.all(rows.filter(isRuntimeRequest).map(async (row) => {
          const requestId = getBatchRequestId(row);
          await markRuntimeRequestFailed(row, error, batchId).catch(() => {});
          failedRequestIds.push(requestId);
        }));
        console.error("[AutoBid GPT Batch Worker] Batch failed", { batchId, error });
        showStatus(`Batch failed; its requests remain queued for a worker retry: ${batchError}`, "error");
      }
    } finally {
      const requestIds = rows.filter(isRuntimeRequest).map(getBatchRequestId);
      await acknowledgeBatchCompletionWithRetry({
        batch_id: batchId,
        request_ids: requestIds,
        saved_request_ids: savedRequestIds,
        failed_request_ids: Array.from(new Set(failedRequestIds)),
        source,
        error: batchError
      }).catch((error) => {
        console.warn("[AutoBid GPT Batch Worker] Could not acknowledge batch completion", error);
      });
    }
  }

  async function acknowledgeBatchCompletionWithRetry(payload) {
    let lastError = null;
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        return await send("AUTOBID_GPT_BATCH_COMPLETE", payload);
      } catch (error) {
        lastError = error;
        if (attempt < 3) await sleep(400);
      }
    }
    throw lastError || new Error("Could not acknowledge the completed AutoBid batch.");
  }

  function startRunOnce() {
    if (oneShotPromise) {
      runAgainRequested = true;
      showStatus("AutoBid answer worker is already running. Added another pass.", "success");
      return;
    }
    oneShotPromise = drainPendingRows({ maxRows: RUN_ONCE_MAX_ROWS, stopWhenEmpty: true })
      .then((result) => {
        if (result?.limitReached && !stopRequested) runAgainRequested = true;
        return result;
      })
      .catch((error) => {
        if (isStopError(error)) {
          showStatus("AutoBid answer worker stopped.", "success");
          return;
        }
        console.error("[AutoBid GPT Answer Worker]", error);
        showStatus(error.message || String(error), "error");
      })
      .finally(() => {
        oneShotPromise = null;
        if (runAgainRequested && !stopRequested) {
          runAgainRequested = false;
          startRunOnce();
        }
      });
  }

  async function runLoop() {
    showStatus("AutoBid GPT answer loop running.", "success");

    while (running) {
      try {
        const result = await processNextPendingRow();
        if (result.stopped || stopRequested) break;
        await sleepInterruptible(result.processed ? ROW_DELAY_MS : LOOP_DELAY_MS);
      } catch (error) {
        if (isStopError(error)) break;
        console.error("[AutoBid GPT Answer Worker]", error);
        showStatus(error.message || String(error), "error");
        try {
          await sleepInterruptible(LOOP_DELAY_MS);
        } catch (sleepError) {
          if (isStopError(sleepError)) break;
          throw sleepError;
        }
      }
    }
  }

  async function drainPendingRows({ maxRows, stopWhenEmpty }) {
    let processed = 0;
    let attempts = 0;
    while (processed < maxRows && attempts < maxRows && !stopRequested) {
      attempts += 1;
      const result = await processNextPendingRow();
      if (result.stopped || stopRequested) break;
      if (!result.processed) {
        if (result.reason === "rows-failed-or-cooling-down") {
          try {
            await sleepInterruptible(ROW_FAILURE_COOLDOWN_MS || LOOP_DELAY_MS);
          } catch (sleepError) {
            if (isStopError(sleepError)) break;
            throw sleepError;
          }
          continue;
        }
        if (stopWhenEmpty) break;
        try {
          await sleepInterruptible(LOOP_DELAY_MS);
        } catch (sleepError) {
          if (isStopError(sleepError)) break;
          throw sleepError;
        }
        continue;
      }
      processed += 1;
      try {
        await sleepInterruptible(ROW_DELAY_MS);
      } catch (sleepError) {
        if (isStopError(sleepError)) break;
        throw sleepError;
      }
    }

    return { processed, limitReached: processed >= maxRows || attempts >= maxRows };
  }

  async function processNextPendingRow() {
    throwIfStopped();
    clearExpiredRowMemory();
    showStatus("Checking pending AutoBid answer requests...", "success");
    const runtimeData = await send("AUTOBID_GPT_FETCH_PENDING_REQUESTS").catch((error) => {
      console.warn("[AutoBid GPT Answer Worker] Runtime queue fetch failed", error);
      return { requests: [] };
    });
    let rows = Array.isArray(runtimeData?.requests) ? runtimeData.requests : [];
    let source = "runtime";

    if (rows.length === 0) {
      const data = await send("AUTOBID_GPT_FETCH_PENDING_ROWS").catch((error) => {
        console.warn("[AutoBid GPT Answer Worker] Sheet queue fetch failed", error);
        return { rows: [] };
      });
      rows = Array.isArray(data?.rows) ? data.rows : [];
      source = "sheet";
    }

    if (rows.length === 0) {
      showStatus("No pending AutoBid answer requests found.", "success");
      return { processed: false, reason: "no-rows" };
    }

    let skipped = 0;
    let lastError = null;

    for (const row of rows) {
      throwIfStopped();
      if (!row?.questions?.fields?.length) {
        skipped += 1;
        continue;
      }

      const key = rowKey(row);
      if (recentCompletedRows.has(key) || recentFailedRows.has(key)) {
        skipped += 1;
        continue;
      }

      showStatus(`Answering ${getRowLabel(row)}...`, "success");

      try {
        let answers = [];
        let lastAttemptError = null;
        for (let attempt = 1; attempt <= ROW_CHAT_ATTEMPTS; attempt += 1) {
          try {
            const responseText = await askChatGpt(buildAnswerPrompt(row), row.questions.fields);
            answers = extractAnswers(responseText, row.questions.fields);
            lastAttemptError = null;
            break;
          } catch (error) {
            if (isStopError(error)) throw error;
            lastAttemptError = error;
            clickStopGenerating();
            if (attempt < ROW_CHAT_ATTEMPTS) {
              const reason = error?.message || "ChatGPT did not return a complete answer";
              showStatus(`${getRowLabel(row)} attempt ${attempt} failed: ${reason} Starting a fresh chat retry...`, "error");
              await sleepInterruptible(700);
            }
          }
        }

        if (lastAttemptError) throw lastAttemptError;

        showStatus(`Saving ${answers.length} answer${answers.length === 1 ? "" : "s"} for ${getRowLabel(row)}...`, "success");
        await saveAnswersWithRetry(row, answers);

        recentCompletedRows.set(key, Date.now() + ROW_COMPLETION_COOLDOWN_MS);
        showStatus(`Saved ${answers.length} answer${answers.length === 1 ? "" : "s"} for ${getRowLabel(row)}. Checking next request...`, "success");
        return { processed: true, source, rowNumber: row.rowNumber, requestId: row.requestId || row.request_id || "", answers: answers.length };
      } catch (error) {
        if (isStopError(error)) {
          showStatus("AutoBid answer loop stopped.", "success");
          return { processed: false, stopped: true };
        }
        lastError = error;
        recentFailedRows.set(key, Date.now() + ROW_FAILURE_COOLDOWN_MS);
        if (isRuntimeRequest(row)) {
          await markRuntimeRequestFailed(row, error).catch(() => {});
        }
        console.error("[AutoBid GPT Answer Worker] Request failed", getRowLabel(row), error);
        showStatus(`${getRowLabel(row)} failed, continuing to next request: ${error.message || String(error)}`, "error");
      }
    }

    if (lastError) {
      return { processed: false, reason: "rows-failed-or-cooling-down", skipped, error: lastError.message || String(lastError) };
    }

    showStatus(skipped ? "Pending rows are cooling down. Checking again soon." : "No pending AutoBid questions found.", "success");
    return { processed: false, reason: "no-eligible-rows", skipped };
  }

  function clearExpiredRowMemory() {
    const now = Date.now();
    for (const [key, expiresAt] of recentFailedRows.entries()) {
      if (expiresAt <= now) recentFailedRows.delete(key);
    }
    for (const [key, expiresAt] of recentCompletedRows.entries()) {
      if (expiresAt <= now) recentCompletedRows.delete(key);
    }
  }

  function rowKey(row) {
    if (row.requestId || row.request_id) return `runtime:${row.requestId || row.request_id}`;
    return `${row.sheetName || ""}:${row.rowNumber || ""}`;
  }

  function getRowLabel(row) {
    if (row.requestId || row.request_id) return `runtime request ${String(row.requestId || row.request_id).slice(0, 12)}`;
    return `sheet row ${row.rowNumber}`;
  }

  function getBatchRequestId(row) {
    const runtimeRequestId = String(row?.requestId || row?.request_id || "");
    if (runtimeRequestId) return runtimeRequestId;
    return `sheet:${row?.sheetName || "sheet"}:${row?.rowNumber || "unknown"}`;
  }

  function isRuntimeRequest(row) {
    return Boolean(row?.runtime || row?.requestId || row?.request_id);
  }

  function requestStop() {
    stopRequested = true;
    running = false;
    clickStopGenerating();
    send("AUTOBID_GPT_WORKER_STATE", { running: false }).catch(() => {});
    showStatus("Stopping AutoBid GPT answer worker...", "success");
  }

  function throwIfStopped() {
    if (!stopRequested) return;
    const error = new Error("AutoBid GPT answer worker stopped.");
    error.name = "AutoBidStopError";
    throw error;
  }

  function isStopError(error) {
    return error?.name === "AutoBidStopError" || /worker stopped/i.test(error?.message || "");
  }

  async function askChatGpt(prompt, fields) {
    throwIfStopped();
    const beforeCount = getAssistantMessages().length;
    const composer = await waitForComposer();
    throwIfStopped();
    await setComposerText(composer, prompt);
    throwIfStopped();
    await clickSend(composer);
    return waitForAssistantResponse(beforeCount, (text) => hasUsableAnswerJson(text, fields));
  }

  async function askChatGptBatch(prompt, rows) {
    throwIfStopped();
    const beforeCount = getAssistantMessages().length;
    const composer = await waitForComposer();
    throwIfStopped();
    await setComposerText(composer, prompt);
    throwIfStopped();
    await clickSend(composer);
    return waitForAssistantResponse(beforeCount, (text) => hasUsableBatchAnswerJson(text, rows));
  }

  async function saveAnswersWithRetry(row, answers, batchId = "") {
    const runtimeRequestId = row.requestId || row.request_id || "";
    const payload = runtimeRequestId ? {
      request_id: runtimeRequestId,
      batch_id: batchId,
      answers,
      payload: {
        answers,
        source: "auto-bid-extension-runtime",
        generated_at: new Date().toISOString()
      }
    } : {
      spreadsheetId: row.spreadsheetId,
      sheetName: row.sheetName,
      rowNumber: row.rowNumber,
      answers,
      payload: {
        answers,
        source: "auto-bid-extension",
        generated_at: new Date().toISOString()
      }
    };
    const messageType = runtimeRequestId ? "AUTOBID_GPT_SAVE_REQUEST_ANSWERS" : "AUTOBID_GPT_SAVE_ANSWERS";

    for (let attempt = 1; attempt <= SAVE_ANSWERS_RETRIES + 1; attempt += 1) {
      throwIfStopped();
      try {
        const result = await withTimeout(
          send(messageType, payload),
          SAVE_ANSWERS_TIMEOUT_MS,
          `Timed out saving ${getRowLabel(row)}`
        );
        if (runtimeRequestId && result?.status !== "complete" && result?.status !== "cancelled") {
          throw new Error(`Background did not durably store ${getRowLabel(row)}.`);
        }
        console.info("[AutoBid GPT Answer Worker] Saved answers", { rowNumber: row.rowNumber, requestId: runtimeRequestId, answers: answers.length, result });
        return result;
      } catch (error) {
        if (attempt > SAVE_ANSWERS_RETRIES) throw error;
        showStatus(`Saving ${getRowLabel(row)} is slow. Retrying...`, "error");
        await sleepInterruptible(SAVE_ANSWERS_RETRY_MS);
      }
    }

    return null;
  }

  function markRuntimeRequestFailed(row, error, batchId = "") {
    const requestId = row.requestId || row.request_id || "";
    if (!requestId) return Promise.resolve(null);
    return send("AUTOBID_GPT_FAIL_REQUEST", {
      request_id: requestId,
      batch_id: batchId,
      error: error.message || String(error)
    });
  }

  async function waitForComposer() {
    const started = Date.now();

    while (Date.now() - started < 30000) {
      throwIfStopped();
      const composer = findComposer();
      if (composer && isVisible(composer)) return composer;
      await sleepInterruptible(300);
    }

    throw new Error("Could not find the ChatGPT message box. Open a normal ChatGPT chat tab and try again.");
  }

  function findComposer() {
    return document.querySelector("#prompt-textarea") ||
      document.querySelector("[data-testid='prompt-textarea']") ||
      Array.from(document.querySelectorAll("textarea, [contenteditable='true']")).find((element) => {
        if (!isVisible(element)) return false;
        const label = `${element.getAttribute("aria-label") || ""} ${element.getAttribute("placeholder") || ""}`;
        return /message|prompt|ask|chatgpt/i.test(label) || element.closest("form");
      }) ||
      null;
  }

  async function setComposerText(composer, text) {
    composer.focus();

    if (composer.matches("textarea, input")) {
      const setter = Object.getOwnPropertyDescriptor(composer.constructor.prototype, "value")?.set;
      if (setter) setter.call(composer, text);
      else composer.value = text;
      composer.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
      return;
    }

    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(composer);
    selection.removeAllRanges();
    selection.addRange(range);
    document.execCommand("insertText", false, text);
    composer.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
  }

  async function clickSend(composer) {
    throwIfStopped();
    const button = findSendButton();
    if (button && !button.disabled && button.getAttribute("aria-disabled") !== "true") {
      button.click();
      return { sent: true, method: "dom-send-button" };
    }

    const form = composer.closest("form");
    if (form?.requestSubmit) {
      form.requestSubmit();
      return { sent: true, method: "dom-request-submit" };
    }

    composer.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 }));
    composer.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 }));
    return { sent: true, method: "dom-enter-fallback" };
  }

  function findSendButton() {
    return document.querySelector("button[data-testid='send-button']") ||
      document.querySelector("button[aria-label*='Send' i]") ||
      document.querySelector("form button[type='submit']") ||
      Array.from(document.querySelectorAll("button")).find((button) => {
        if (!isVisible(button)) return false;
        const text = `${button.textContent || ""} ${button.getAttribute("aria-label") || ""}`;
        return /\bsend\b/i.test(text);
      }) ||
      null;
  }

  async function waitForAssistantResponse(beforeCount, hasUsableResponse) {
    const started = Date.now();
    let lastText = "";
    let stableSince = 0;
    let lastProgressAt = started;
    let noTextWindowStartedAt = started;
    let lastVisibilityState = document.visibilityState;

    while (Date.now() - started < RESPONSE_TIMEOUT_MS) {
      throwIfStopped();
      const now = Date.now();
      const visibilityState = document.visibilityState;
      const backgrounded = visibilityState !== "visible";
      const noTextTimeoutMs = backgrounded
        ? RESPONSE_BACKGROUND_NO_TEXT_TIMEOUT_MS
        : RESPONSE_NO_TEXT_TIMEOUT_MS;
      const stallTimeoutMs = backgrounded
        ? RESPONSE_BACKGROUND_STALL_TIMEOUT_MS
        : RESPONSE_STALL_TIMEOUT_MS;

      if (visibilityState !== lastVisibilityState) {
        // Chrome can heavily throttle timers in an inactive tab. Give the DOM a
        // fresh observation window whenever the tab changes visibility.
        lastVisibilityState = visibilityState;
        lastProgressAt = now;
        noTextWindowStartedAt = now;
      }

      const text = getLatestAssistantResponseText(beforeCount);

      if (text) {
        if (text !== lastText) {
          lastText = text;
          stableSince = now;
          lastProgressAt = now;
        } else if (hasUsableResponse(text) && now - stableSince >= RESPONSE_JSON_STABLE_MS) {
          return text;
        } else if (!isGenerating() && now - stableSince >= RESPONSE_STABLE_MS) {
          return text;
        } else if (now - lastProgressAt >= stallTimeoutMs) {
          const finalText = getLatestAssistantResponseText(beforeCount);
          if (hasUsableResponse(finalText)) return finalText;
          clickStopGenerating();
          throw new Error(`ChatGPT answer made no progress for ${Math.round(stallTimeoutMs / 1000)} seconds${backgrounded ? " while the tab was in the background" : ""}.`);
        }
      } else if (now - noTextWindowStartedAt >= noTextTimeoutMs) {
        const finalText = getLatestAssistantResponseText(beforeCount);
        if (hasUsableResponse(finalText)) return finalText;
        clickStopGenerating();
        throw new Error(`ChatGPT did not start a visible answer within ${Math.round(noTextTimeoutMs / 1000)} seconds${backgrounded ? " while the tab was in the background" : ""}.`);
      }

      await waitForDomActivity(RESPONSE_POLL_MS);
    }

    const finalText = getLatestAssistantResponseText(beforeCount);
    if (hasUsableResponse(finalText)) return finalText;
    clickStopGenerating();
    throw new Error("Timed out waiting for ChatGPT answer.");
  }

  function getLatestAssistantResponseText(beforeCount) {
    const messages = getAssistantMessages();
    const next = messages.length > beforeCount ? messages[messages.length - 1] : null;
    return cleanResponseText(next?.innerText || next?.textContent || "");
  }

  function getAssistantMessages() {
    return Array.from(document.querySelectorAll("[data-message-author-role='assistant']"))
      .filter(isVisible);
  }

  function isGenerating() {
    return Boolean(
      document.querySelector("button[data-testid='stop-button']") ||
      document.querySelector("button[aria-label*='Stop' i]")
    );
  }

  function clickStopGenerating() {
    const button = document.querySelector("button[data-testid='stop-button']") ||
      document.querySelector("button[aria-label*='Stop' i]");
    if (button && isVisible(button)) button.click();
  }

  function buildAnswerPrompt(row) {
    const payload = buildAnswerPayload(row, 12000);

    return [
      "You are the AutoBid job application answer worker.",
      "Generate honest, concise answers for the application fields in the JSON input.",
      "Return ONLY valid JSON. Do not use markdown or code fences.",
      "",
      "Required output shape:",
      "{\"answers\":[{\"field_id\":\"exact field_id\",\"question\":\"plain question text\",\"option\":\"checkbox option when present\",\"value\":\"answer to fill\"}]}",
      "",
      ...getAnswerRules(),
      "",
      "JSON input:",
      JSON.stringify(payload)
    ].join("\n");
  }

  function buildBatchAnswerPrompt(rows) {
    const payload = {
      batch_size: rows.length,
      requests: rows.map((row) => ({
        request_id: getBatchRequestId(row),
        ...buildAnswerPayload(row, 8000)
      }))
    };

    return [
      "You are the AutoBid job application batch answer worker.",
      "Generate honest, concise answers for EVERY request and EVERY field in the JSON input.",
      "Keep requests isolated. Never copy facts or answers from one request into another.",
      "Return ONLY valid JSON. Do not use markdown or code fences.",
      "",
      "Required output shape:",
      "{\"requests\":[{\"request_id\":\"exact request_id\",\"answers\":[{\"field_id\":\"exact field_id\",\"question\":\"plain question text\",\"option\":\"checkbox option when present\",\"value\":\"answer to fill\"}]}]}",
      "",
      "Batch rules:",
      "- Return exactly one requests entry for every input request_id.",
      "- Preserve every request_id and field_id exactly as provided.",
      ...getAnswerRules(),
      "",
      "JSON input:",
      JSON.stringify(payload)
    ].join("\n");
  }

  function getAnswerRules() {
    return [
      "Rules:",
      "- Use tailored_resume_content as the primary candidate/resume source when present.",
      "- Use job_description as the primary role/source-of-truth when present.",
      "- Use row values for company, stack, salary, location, and other job-row metadata when present.",
      "- Do not invent employers, degrees, certifications, locations, compensation, legal eligibility, languages, or years of experience.",
      "- Use field.question as the plain question text. For checkbox fields with field.option, answer Yes only when that option should be selected.",
      "- For select, radio, checkbox, and button-group fields with provided options, value must exactly equal one provided option.",
      "- If the question asks about English, answer Yes. If it asks about another language and the profile does not prove it, answer No.",
      "- Ignore decorative/browser fallback text such as 'SVGs not supported by this browser'.",
      "- Keep textarea answers specific but short, usually 2 to 5 sentences.",
      "- Never answer with placeholder text such as 'Not specified', 'Unknown', 'Not provided', 'TBD', or similar.",
      "- Use 'N/A' or 'Not applicable' only when the field's own question/help text explicitly permits it or the field is genuinely conditional and inapplicable.",
      "- For required narrative questions, write a useful first-person answer grounded in the supplied resume/profile. If an exact claim is unsupported, describe the closest supported experience honestly instead of returning a placeholder.",
      "- Answer every included field when possible. If one field truly cannot be answered honestly, omit only that field; keep every other valid answer."
    ];
  }

  function buildAnswerPayload(row, contextLimit) {
    const questions = row.questions || {};
    const rowValues = questions.row?.values || row.values || {};
    const gptContext = questions.gpt_context || {};
    const tailoredResumeContent = limitPromptText(
      gptContext.tailored_resume_content ||
      gptContext.column_g ||
      questions.row?.tailored_resume_content ||
      questions.row?.column_g ||
      findRowValueByAlias(rowValues, ["tailored_resume_content", "tailored resume content", "tailor_resume_content", "tailor resume content", "tailored_resume", "tailored resume", "resume_content", "resume content", "candidate_profile", "candidate profile", "profile", "column_g"]),
      contextLimit
    );
    const jobDescription = limitPromptText(
      gptContext.job_description ||
      gptContext.column_m ||
      questions.row?.job_description ||
      questions.row?.column_m ||
      findRowValueByAlias(rowValues, ["job_description", "job description", "jd", "job_desc", "job desc", "description", "job_posting", "job posting", "job_details", "job details", "column_m"]),
      contextLimit
    );
    return {
      sheet_row: row.rowNumber,
      apply_url: row.url || questions.row?.url || questions.page?.url || "",
      profile: questions.profile || {},
      tailored_resume_content: tailoredResumeContent,
      job_description: jobDescription,
      row_values: compactRowValues(rowValues),
      page: questions.page || {},
      fields: (questions.fields || []).map(sanitizePromptField)
    };
  }

  function compactRowValues(values) {
    const compact = {};
    let remaining = 6000;
    for (const [key, value] of Object.entries(values || {})) {
      const normalizedKey = normalizeKey(key);
      if (/tailored resume|resume content|job description|job posting|column g|column m/.test(normalizedKey)) continue;
      const text = sanitizePromptText(value);
      if (!text || remaining <= 0) continue;
      const limited = text.slice(0, Math.min(1000, remaining));
      compact[key] = limited;
      remaining -= limited.length;
    }
    return compact;
  }

  function sanitizePromptField(field) {
    return {
      ...field,
      question: sanitizePromptText(field.question || field.label || ""),
      option: sanitizePromptText(field.option || ""),
      label: sanitizePromptText(field.label || field.question || ""),
      raw_label: sanitizePromptText(field.raw_label || ""),
      help_text: sanitizePromptText(field.help_text || ""),
      placeholder: sanitizePromptText(field.placeholder || ""),
      options: Array.isArray(field.options)
        ? field.options.map(sanitizePromptText).filter(Boolean)
        : []
    };
  }

  function sanitizePromptText(value) {
    return String(value || "")
      .replace(/SVGs?\s+not\s+supported\s+by\s+this\s+browser\.?\s*C?\d*/ig, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function limitPromptText(value, maxLength) {
    const text = sanitizePromptText(value);
    return text.length > maxLength ? `${text.slice(0, maxLength - 3)}...` : text;
  }

  function findRowValueByAlias(values, aliases) {
    const normalizedAliases = new Set(aliases.map(normalizeKey));
    for (const [key, value] of Object.entries(values || {})) {
      const text = String(value || "").trim();
      if (text && normalizedAliases.has(normalizeKey(key))) return text;
    }
    return "";
  }

  function normalizeKey(value) {
    return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  }

  function extractAnswers(responseText, fields) {
    const parsed = parseJsonFromText(responseText);
    return extractAnswersFromPayload(parsed, fields);
  }

  function extractAnswersFromPayload(parsed, fields, options = {}) {
    const fieldList = (fields || []).filter(Boolean);
    const fieldsById = new Map(fieldList.map((field) => [String(field.field_id || field.id || ""), field]));
    const fieldIds = new Set(Array.from(fieldsById.keys()).filter(Boolean));
    const rawAnswers = Array.isArray(parsed?.answers)
      ? parsed.answers
      : parsed && typeof parsed === "object"
        ? Object.entries(parsed).map(([field_id, value]) => ({ field_id, value }))
        : [];

    const answers = rawAnswers
      .map((answer) => {
        const field = resolveAnswerField(answer, fieldList, fieldsById);
        if (!field) return null;
        const fieldId = String(field.field_id || field.id || "");
        return {
          field_id: fieldId,
          question: sanitizePromptText(answer.question || field.question || field.label || ""),
          option: sanitizePromptText(answer.option || field.option || ""),
          value: sanitizePromptText(answer.value ?? answer.answer ?? "")
        };
      })
      .filter((answer) => answer?.field_id && fieldIds.has(answer.field_id) && answer.value)
      .filter((answer) => {
        const field = fieldsById.get(answer.field_id);
        return !isRejectedPlaceholderAnswer(answer.value, field);
      })
      .filter((answer, index, items) => items.findIndex((item) => item.field_id === answer.field_id) === index);

    if (answers.length === 0) {
      throw new Error("ChatGPT response did not contain usable AutoBid answers.");
    }

    const answeredFieldIds = new Set(answers.map((answer) => answer.field_id));
    const missingFieldIds = Array.from(fieldIds).filter((fieldId) => !answeredFieldIds.has(fieldId));
    if (missingFieldIds.length > 0 && options.requireComplete !== false) {
      console.warn("[AutoBid GPT Answer Worker] ChatGPT returned a partial answer set", {
        answered: answers.length,
        missing_field_ids: missingFieldIds
      });
    }

    return answers;
  }

  function isRejectedPlaceholderAnswer(value, field = {}) {
    const normalized = normalizeKey(value);
    if (!/^(?:not specified|unspecified|unknown|not provided|not available|no information(?: provided| available)?|information unavailable|to be determined|tbd)$/.test(normalized)) {
      return false;
    }
    const options = Array.isArray(field?.options) ? field.options.map(normalizeKey) : [];
    return !options.includes(normalized);
  }

  function resolveAnswerField(answer, fields, fieldsById) {
    const requestedId = String(answer?.field_id || answer?.id || "").trim();
    if (requestedId && fieldsById.has(requestedId)) return fieldsById.get(requestedId);

    const identities = new Set([
      requestedId,
      answer?.question,
      answer?.label
    ].map(normalizeKey).filter(Boolean));
    if (identities.size === 0) return null;

    const matches = fields.filter((field) => {
      const fieldIdentities = [
        field.question,
        field.label,
        field.name
      ].map(normalizeKey).filter(Boolean);
      return fieldIdentities.some((identity) => identities.has(identity));
    });
    return matches.length === 1 ? matches[0] : null;
  }

  function extractBatchAnswers(responseText, rows) {
    const parsed = parseJsonFromText(responseText);
    let rawRequests = Array.isArray(parsed?.requests)
      ? parsed.requests
      : Array.isArray(parsed?.results)
        ? parsed.results
        : [];

    if (rawRequests.length === 0 && rows.length === 1 && Array.isArray(parsed?.answers)) {
      rawRequests = [{ request_id: getBatchRequestId(rows[0]), answers: parsed.answers }];
    }
    if (rawRequests.length === 0 && parsed && typeof parsed === "object") {
      rawRequests = rows.map((row) => {
        const requestId = getBatchRequestId(row);
        const value = parsed[requestId];
        if (Array.isArray(value)) return { request_id: requestId, answers: value };
        if (value && typeof value === "object") return { request_id: requestId, ...value };
        return null;
      }).filter(Boolean);
    }

    const groupsById = new Map(rawRequests.map((request) => [
      String(request?.request_id || request?.requestId || request?.id || ""),
      request
    ]).filter(([requestId]) => requestId));
    const results = new Map();

    for (const row of rows) {
      const requestId = getBatchRequestId(row);
      const group = groupsById.get(requestId);
      if (!group) {
        results.set(requestId, { answers: [], error: new Error(`ChatGPT omitted request ${requestId}.`) });
        continue;
      }
      try {
        results.set(requestId, {
          answers: extractAnswersFromPayload(
            { answers: group.answers || group.values || [] },
            row.questions?.fields || [],
            { requireComplete: false }
          ),
          error: null
        });
      } catch (error) {
        results.set(requestId, { answers: [], error });
      }
    }

    return results;
  }

  function hasUsableAnswerJson(responseText, fields) {
    try {
      const parsed = parseJsonFromText(responseText);
      return extractAnswersFromPayload(parsed, fields, { requireComplete: false }).length > 0;
    } catch {
      return false;
    }
  }

  function hasUsableBatchAnswerJson(responseText, rows) {
    try {
      const results = extractBatchAnswers(responseText, rows);
      return rows.every((row) => {
        const result = results.get(getBatchRequestId(row));
        return Boolean(result?.answers?.length) && !result.error;
      });
    } catch {
      return false;
    }
  }

  function parseJsonFromText(text) {
    const source = String(text || "").trim();
    const candidates = getJsonCandidates(source);

    for (const candidate of candidates) {
      const parsed = tryParseJson(candidate);
      if (parsed && (Array.isArray(parsed.requests) || Array.isArray(parsed.results))) return parsed;
    }

    for (const candidate of candidates) {
      const parsed = tryParseJson(candidate);
      if (parsed && Array.isArray(parsed.answers)) return parsed;
    }

    for (const candidate of candidates) {
      const parsed = tryParseJson(candidate);
      if (parsed) return parsed;
    }

    throw new Error("ChatGPT response did not include usable JSON.");
  }

  function getJsonCandidates(source) {
    const candidates = [];
    const fencedMatches = Array.from(source.matchAll(/```(?:json)?\s*([\s\S]*?)```/ig))
      .map((match) => match[1].trim())
      .filter(Boolean)
      .reverse();
    candidates.push(...fencedMatches);

    const starts = [];
    for (let index = 0; index < source.length; index += 1) {
      if (source[index] === "{") starts.push(index);
    }

    for (const start of starts.reverse()) {
      const end = findJsonObjectEnd(source, start);
      if (end > start) candidates.push(source.slice(start, end + 1));
    }

    if (source) candidates.push(source);
    return Array.from(new Set(candidates));
  }

  function findJsonObjectEnd(source, start) {
    let depth = 0;
    let inString = false;
    let escaped = false;

    for (let index = start; index < source.length; index += 1) {
      const char = source[index];

      if (inString) {
        if (escaped) {
          escaped = false;
        } else if (char === "\\") {
          escaped = true;
        } else if (char === "\"") {
          inString = false;
        }
        continue;
      }

      if (char === "\"") {
        inString = true;
      } else if (char === "{") {
        depth += 1;
      } else if (char === "}") {
        depth -= 1;
        if (depth === 0) return index;
      }
    }

    return -1;
  }

  function tryParseJson(candidate) {
    try {
      return JSON.parse(candidate);
    } catch (_error) {
      return null;
    }
  }

  function cleanResponseText(text) {
    return String(text || "").replace(/\s+/g, " ").trim();
  }

  function isVisible(element) {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = window.getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
  }

  function showStatus(message, kind) {
    let status = document.getElementById(STATUS_ID);
    if (!status) {
      status = document.createElement("div");
      status.id = STATUS_ID;
      document.documentElement.append(status);
    }

    status.textContent = message;
    status.style.cssText = [
      "position:fixed",
      "right:18px",
      "top:18px",
      "z-index:2147483647",
      "max-width:min(390px, calc(100vw - 36px))",
      "padding:10px 12px",
      "border-radius:8px",
      "font:13px/1.4 system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif",
      "box-shadow:0 14px 40px rgba(19,35,29,.22)",
      kind === "error" ? "background:#fff7ed;color:#9a3412;border:1px solid #fdba74" : "background:#ecfdf5;color:#065f46;border:1px solid #86efac"
    ].join(";");
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
    });
  }

  function withTimeout(promise, timeoutMs, message) {
    let timeoutId = null;
    const timeout = new Promise((_resolve, reject) => {
      timeoutId = window.setTimeout(() => reject(new Error(message)), timeoutMs);
    });
    return Promise.race([promise, timeout]).finally(() => {
      if (timeoutId) window.clearTimeout(timeoutId);
    });
  }

  function sleep(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }

  function waitForDomActivity(timeoutMs) {
    throwIfStopped();
    return new Promise((resolve) => {
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        observer.disconnect();
        window.clearTimeout(timer);
        resolve();
      };
      const observer = new MutationObserver(finish);
      const timer = window.setTimeout(finish, Math.max(50, Number(timeoutMs || 0)));
      observer.observe(document.documentElement, {
        childList: true,
        subtree: true,
        characterData: true
      });
    }).then(() => {
      throwIfStopped();
    });
  }

  async function sleepInterruptible(ms) {
    const started = Date.now();
    while (Date.now() - started < ms) {
      throwIfStopped();
      await sleep(Math.min(250, ms - (Date.now() - started)));
    }
  }
})();
