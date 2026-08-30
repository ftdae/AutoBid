(() => {
  if (window.__autoBidChatGptWorkerLoaded) return;
  window.__autoBidChatGptWorkerLoaded = true;

  const STATUS_ID = "autobid-chatgpt-answer-status";
  const RESPONSE_TIMEOUT_MS = 180000;
  const RESPONSE_STABLE_MS = 2500;
  const RESPONSE_JSON_STABLE_MS = 900;
  const RESPONSE_POLL_MS = 500;
  const SAVE_ANSWERS_TIMEOUT_MS = 12000;
  const SAVE_ANSWERS_RETRIES = 1;
  const SAVE_ANSWERS_RETRY_MS = 1500;
  const LOOP_DELAY_MS = 5000;
  const ROW_DELAY_MS = 1500;
  const ROW_FAILURE_COOLDOWN_MS = 60000;
  const ROW_COMPLETION_COOLDOWN_MS = 30000;
  const RUN_ONCE_MAX_ROWS = 50;

  let running = false;
  let loopPromise = null;
  let oneShotPromise = null;
  const recentFailedRows = new Map();
  const recentCompletedRows = new Map();

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type === "AUTOBID_CHATGPT_PING") {
      sendResponse({ message: "ready" });
      return false;
    }

    if (message?.type === "AUTOBID_CHATGPT_RUN_ONCE") {
      startRunOnce();
      sendResponse({ message: "Answer worker started" });
      return false;
    }

    if (message?.type === "AUTOBID_CHATGPT_START") {
      running = true;
      send("WORKER_STATE", { running: true }).catch(() => {});
      if (!loopPromise) loopPromise = runLoop().finally(() => { loopPromise = null; });
      sendResponse({ message: "Answer loop started" });
      return false;
    }

    if (message?.type === "AUTOBID_CHATGPT_STOP") {
      running = false;
      send("WORKER_STATE", { running: false }).catch(() => {});
      showStatus("AutoBid answer loop stopped.", "success");
      sendResponse({ message: "Stopped" });
      return false;
    }

    return false;
  });

  function startRunOnce() {
    if (oneShotPromise) return;
    oneShotPromise = drainPendingRows({ maxRows: RUN_ONCE_MAX_ROWS, stopWhenEmpty: true })
      .catch((error) => {
        console.error("[AutoBid Answer Worker]", error);
        showStatus(error.message || String(error), "error");
      })
      .finally(() => {
        oneShotPromise = null;
      });
  }

  async function runLoop() {
    showStatus("AutoBid answer loop running.", "success");

    while (running) {
      try {
        const result = await processNextPendingRow();
        await sleep(result.processed ? ROW_DELAY_MS : LOOP_DELAY_MS);
      } catch (error) {
        console.error("[AutoBid Answer Worker]", error);
        showStatus(error.message || String(error), "error");
        await sleep(LOOP_DELAY_MS);
      }
    }
  }

  async function drainPendingRows({ maxRows, stopWhenEmpty }) {
    let processed = 0;
    while (processed < maxRows) {
      const result = await processNextPendingRow();
      if (!result.processed) {
        if (stopWhenEmpty) break;
        await sleep(LOOP_DELAY_MS);
        continue;
      }
      processed += 1;
      await sleep(ROW_DELAY_MS);
    }

    return { processed };
  }

  async function processNextPendingRow() {
    clearExpiredRowMemory();
    showStatus("Checking Google Sheet for pending AutoBid questions...", "success");
    const data = await send("FETCH_PENDING_ROWS");
    const rows = Array.isArray(data?.rows) ? data.rows : [];

    if (rows.length === 0) {
      showStatus("No pending AutoBid questions found.", "success");
      return { processed: false, reason: "no-rows" };
    }

    let skipped = 0;
    let lastError = null;

    for (const row of rows) {
      if (!row?.questions?.fields?.length) {
        skipped += 1;
        continue;
      }

      const key = rowKey(row);
      if (recentCompletedRows.has(key) || recentFailedRows.has(key)) {
        skipped += 1;
        continue;
      }

      showStatus(`Answering sheet row ${row.rowNumber}...`, "success");

      try {
        const responseText = await askChatGpt(buildAnswerPrompt(row), row.questions.fields);
        const answers = extractAnswers(responseText, row.questions.fields);

        showStatus(`Saving ${answers.length} answer${answers.length === 1 ? "" : "s"} for row ${row.rowNumber}...`, "success");
        await saveAnswersWithRetry(row, answers);

        recentCompletedRows.set(key, Date.now() + ROW_COMPLETION_COOLDOWN_MS);
        showStatus(`Saved ${answers.length} answer${answers.length === 1 ? "" : "s"} for row ${row.rowNumber}. Checking next row...`, "success");
        return { processed: true, rowNumber: row.rowNumber, answers: answers.length };
      } catch (error) {
        lastError = error;
        recentFailedRows.set(key, Date.now() + ROW_FAILURE_COOLDOWN_MS);
        console.error("[AutoBid Answer Worker] Row failed", row.rowNumber, error);
        showStatus(`Row ${row.rowNumber} failed, continuing to next pending row: ${error.message || String(error)}`, "error");
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
    return `${row.sheetName || ""}:${row.rowNumber || ""}`;
  }

  async function askChatGpt(prompt, fields) {
    const beforeCount = getAssistantMessages().length;
    const composer = await waitForComposer();
    await setComposerText(composer, prompt);
    await sleep(250);
    await clickSend(composer);
    return waitForAssistantResponse(beforeCount, fields);
  }

  async function saveAnswersWithRetry(row, answers) {
    const payload = {
      sheetName: row.sheetName,
      rowNumber: row.rowNumber,
      answers
    };

    for (let attempt = 1; attempt <= SAVE_ANSWERS_RETRIES + 1; attempt += 1) {
      try {
        const result = await withTimeout(
          send("SAVE_ANSWERS", payload),
          SAVE_ANSWERS_TIMEOUT_MS,
          `Timed out saving row ${row.rowNumber} to Google Sheet`
        );
        console.info("[AutoBid Answer Worker] Saved answers", { rowNumber: row.rowNumber, answers: answers.length, result });
        return result;
      } catch (error) {
        if (attempt > SAVE_ANSWERS_RETRIES) throw error;
        showStatus(`Saving row ${row.rowNumber} is slow. Retrying...`, "error");
        await sleep(SAVE_ANSWERS_RETRY_MS);
      }
    }

    return null;
  }

  async function waitForComposer() {
    const started = Date.now();

    while (Date.now() - started < 30000) {
      const composer = findComposer();
      if (composer && isVisible(composer)) return composer;
      await sleep(300);
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
    const started = Date.now();

    while (Date.now() - started < 15000) {
      const button = findSendButton();
      if (button && !button.disabled && button.getAttribute("aria-disabled") !== "true") {
        button.click();
        return;
      }
      await sleep(250);
    }

    composer.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 }));
    composer.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 }));
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

  async function waitForAssistantResponse(beforeCount, fields) {
    const started = Date.now();
    let lastText = "";
    let stableSince = 0;

    while (Date.now() - started < RESPONSE_TIMEOUT_MS) {
      const messages = getAssistantMessages();
      const next = messages.length > beforeCount ? messages[messages.length - 1] : null;
      const text = cleanResponseText(next?.innerText || next?.textContent || "");

      if (text) {
        if (text !== lastText) {
          lastText = text;
          stableSince = Date.now();
        } else if (hasUsableAnswerJson(text, fields) && Date.now() - stableSince >= RESPONSE_JSON_STABLE_MS) {
          return text;
        } else if (!isGenerating() && Date.now() - stableSince >= RESPONSE_STABLE_MS) {
          return text;
        }
      }

      await sleep(RESPONSE_POLL_MS);
    }

    throw new Error("Timed out waiting for ChatGPT answer.");
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

  function buildAnswerPrompt(row) {
    const questions = row.questions || {};
    const payload = {
      sheet_row: row.rowNumber,
      apply_url: row.url || questions.row?.url || questions.page?.url || "",
      row_values: questions.row?.values || row.values || {},
      page: questions.page || {},
      fields: (questions.fields || []).map(sanitizePromptField)
    };

    return [
      "You are the AutoBid job application answer worker.",
      "Generate honest, concise answers for the application fields in the JSON input.",
      "Return ONLY valid JSON. Do not use markdown or code fences.",
      "",
      "Required output shape:",
      "{\"answers\":[{\"field_id\":\"exact field_id\",\"question\":\"plain question text\",\"option\":\"checkbox option when present\",\"value\":\"answer to fill\"}]}",
      "",
      "Rules:",
      "- Use the candidate profile, tailored resume, job description, and row values when present.",
      "- Do not invent employers, degrees, certifications, locations, compensation, legal eligibility, languages, or years of experience.",
      "- Use field.question as the plain question text. For checkbox fields with field.option, answer Yes only when that option should be selected.",
      "- For select, radio, checkbox, and button-group fields, use one of the provided options when possible.",
      "- If the question asks about English, answer Yes. If it asks about another language and the profile does not prove it, answer No.",
      "- Ignore decorative/browser fallback text such as 'SVGs not supported by this browser'.",
      "- Keep textarea answers specific but short, usually 2 to 5 sentences.",
      "- Answer every included field unless doing so would be dishonest.",
      "",
      "JSON input:",
      JSON.stringify(payload)
    ].join("\n");
  }

  function sanitizePromptField(field) {
    return {
      ...field,
      question: sanitizePromptText(field.question || field.label || ""),
      option: sanitizePromptText(field.option || ""),
      label: sanitizePromptText(field.label || field.question || ""),
      raw_label: sanitizePromptText(field.raw_label || ""),
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

  function extractAnswers(responseText, fields) {
    const parsed = parseJsonFromText(responseText);
    const fieldsById = new Map((fields || []).map((field) => [String(field.field_id || field.id || ""), field]));
    const fieldIds = new Set(fieldsById.keys());
    const rawAnswers = Array.isArray(parsed?.answers)
      ? parsed.answers
      : parsed && typeof parsed === "object"
        ? Object.entries(parsed).map(([field_id, value]) => ({ field_id, value }))
        : [];

    const answers = rawAnswers
      .map((answer) => {
        const fieldId = String(answer.field_id || answer.id || "");
        const field = fieldsById.get(fieldId) || {};
        return {
          field_id: fieldId,
          question: String(answer.question || field.question || field.label || "").trim(),
          option: String(answer.option || field.option || "").trim(),
          value: sanitizePromptText(answer.value ?? answer.answer ?? "")
        };
      })
      .filter((answer) => answer.field_id && fieldIds.has(answer.field_id));

    if (answers.length === 0) {
      throw new Error("ChatGPT response did not contain usable AutoBid answers.");
    }

    return answers;
  }

  function hasUsableAnswerJson(responseText, fields) {
    try {
      const parsed = parseJsonFromText(responseText);
      const fieldIds = new Set((fields || []).map((field) => String(field.field_id || field.id || "")).filter(Boolean));
      const answers = Array.isArray(parsed?.answers)
        ? parsed.answers
        : parsed && typeof parsed === "object"
          ? Object.entries(parsed).map(([field_id, value]) => ({ field_id, value }))
          : [];
      return answers.some((answer) => {
        const fieldId = String(answer.field_id || answer.id || "");
        return fieldId && (!fieldIds.size || fieldIds.has(fieldId));
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
})();
