import { createHash, createSign } from "node:crypto";
import { lookup as dnsLookup } from "node:dns";
import { mkdir, writeFile } from "node:fs/promises";
import { Agent as HttpsAgent, request as httpsRequest } from "node:https";
import os from "node:os";
import path from "node:path";
import {
  GOOGLE_APPS_SCRIPT_CONNECT_HOST,
  GOOGLE_APPS_SCRIPT_SECRET,
  GOOGLE_APPS_SCRIPT_TIMEOUT_MS,
  GOOGLE_APPS_SCRIPT_WEB_APP_URL,
  GOOGLE_PRIVATE_KEY,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_SHEETS_SPREADSHEET_ID
} from "../config.js";
import { logBackendEvent } from "../utils/logger.js";

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
const AUTO_BID_COLUMNS = {
  questions: "autobid_questions",
  answers: "autobid_answers",
  status: "autobid_status",
  updatedAt: "autobid_updated_at"
};
const TAILORED_RESUME_COLUMN = 7;
const RESUME_FILE_COLUMN = 10;
const JOB_DESCRIPTION_COLUMN = 13;
const JOB_URL_HEADERS = [
  "job_url",
  "job url",
  "joburl",
  "job link",
  "apply_url",
  "apply url",
  "applyurl",
  "apply link",
  "application_url",
  "application url",
  "applicationurl",
  "application link",
  "url",
  "link"
];
const RESUME_FILE_VALUE_KEYS = [
  "resume_link",
  "resume_url",
  "resume_pdf",
  "generated_resume",
  "generated_resume_link",
  "tailored_resume",
  "tailored_resume_link",
  "cv_link",
  "pdf_resume",
  "column_j"
];

let tokenCache = null;
const appsScriptAgent = new HttpsAgent({
  keepAlive: true,
  lookup: lookupAppsScriptAddress
});

export async function listSheetJobs({ spreadsheetId, sheetName, startRow, endRow }) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow, endRow });
  if (shouldUseAppsScriptBridge()) {
    const data = await callAppsScriptBridge("autobidListJobs", {
      sheetName: normalized.sheetName,
      startRow: normalized.startRow,
      endRow: normalized.endRow
    });
    return normalizeBridgeJobs(data.jobs || data.rows || []);
  }

  const values = await getValues(normalized.spreadsheetId, `${quoteSheetName(normalized.sheetName)}!A1:ZZ${normalized.endRow}`);
  const rows = values.values || [];
  const headers = normalizeHeaders(rows[0] || []);

  return rows
    .map((row, index) => ({
      rowNumber: index + 1,
      values: withFixedColumnValues(rowToObject(headers, row), row),
      raw: row,
      url: findJobUrl(headers, row)
    }))
    .filter((row) => row.rowNumber >= normalized.startRow && row.rowNumber <= normalized.endRow && row.url);
}

export async function writeQuestionPayload({ spreadsheetId, sheetName, rowNumber, payload }) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow: rowNumber, endRow: rowNumber });
  if (shouldUseAppsScriptBridge()) {
    const data = await callAppsScriptBridge("autobidSaveQuestions", {
      sheetName: normalized.sheetName,
      rowNumber,
      row: rowNumber,
      payload
    });
    return {
      ok: true,
      updated_at: data.updated_at || data.updatedAt || new Date().toISOString()
    };
  }

  const headers = await ensureAutoBidHeaders(normalized.spreadsheetId, normalized.sheetName);
  const now = new Date().toISOString();
  const updates = [
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.questions] + 1, rowNumber, JSON.stringify(payload)),
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.status] + 1, rowNumber, "questions_pending"),
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.updatedAt] + 1, rowNumber, now)
  ];

  await batchUpdateValues(normalized.spreadsheetId, updates);
  return { ok: true, updated_at: now };
}

export async function listPendingQuestionRows({ spreadsheetId, sheetName, startRow, endRow }) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow, endRow });
  if (shouldUseAppsScriptBridge()) {
    const data = await callAppsScriptBridge("autobidListPendingQuestions", {
      sheetName: normalized.sheetName,
      startRow: normalized.startRow,
      endRow: normalized.endRow
    });
    return normalizePendingQuestionRows(data.rows || data.pendingRows || [], normalized);
  }

  const headers = await ensureAutoBidHeaders(normalized.spreadsheetId, normalized.sheetName);
  const lastColumn = Math.max(
    ...Object.values(headers).map((index) => Number(index) + 1),
    RESUME_FILE_COLUMN,
    JOB_DESCRIPTION_COLUMN
  );
  const values = await getValues(
    normalized.spreadsheetId,
    `${quoteSheetName(normalized.sheetName)}!A1:${columnName(lastColumn)}${normalized.endRow}`
  );
  const rows = values.values || [];
  const headerRow = rows[0] || [];
  const normalizedHeaders = normalizeHeaders(headerRow);

  return rows
    .map((row, index) => {
      const rowNumber = index + 1;
      if (rowNumber < normalized.startRow || rowNumber > normalized.endRow) return null;

      const questionsRaw = String(row[headers[AUTO_BID_COLUMNS.questions]] || "").trim();
      const answersRaw = String(row[headers[AUTO_BID_COLUMNS.answers]] || "").trim();
      if (!questionsRaw || answersRaw) return null;

      const questions = parseQuestionPayload(questionsRaw);
      if (!questions?.fields?.length) return null;

      const rowValues = withFixedColumnValues(rowToObject(normalizedHeaders, row), row);
      return {
        spreadsheetId: normalized.spreadsheetId,
        sheetName: normalized.sheetName,
        rowNumber,
        row: rowNumber,
        url: findJobUrl(normalizedHeaders, row) || questions.row?.url || questions.page?.url || "",
        values: rowValues,
        raw: row,
        questions
      };
    })
    .filter(Boolean);
}

export async function readAnswerPayload({ spreadsheetId, sheetName, rowNumber }) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow: rowNumber, endRow: rowNumber });
  if (shouldUseAppsScriptBridge()) {
    const data = await callAppsScriptBridge("autobidReadAnswers", {
      sheetName: normalized.sheetName,
      rowNumber,
      row: rowNumber
    });
    const raw = data.raw || data.answers_raw || data.answersRaw || "";
    return {
      row_number: Number(data.row_number || data.rowNumber || rowNumber),
      raw,
      answers: Array.isArray(data.answers) ? normalizeAnswerArray(data.answers) : parseSheetAnswers(raw)
    };
  }

  const headers = await ensureAutoBidHeaders(normalized.spreadsheetId, normalized.sheetName);
  const column = headers[AUTO_BID_COLUMNS.answers] + 1;
  const range = `${quoteSheetName(normalized.sheetName)}!${columnName(column)}${rowNumber}:${columnName(column)}${rowNumber}`;
  const data = await getValues(normalized.spreadsheetId, range);
  const raw = data.values?.[0]?.[0] || "";
  return {
    row_number: rowNumber,
    raw,
    answers: parseSheetAnswers(raw)
  };
}

export async function writeAnswerPayload({ spreadsheetId, sheetName, rowNumber, answers = [], payload = null }) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow: rowNumber, endRow: rowNumber });
  const normalizedPayload = {
    answers: normalizeAnswerArray(answers),
    ...(payload && typeof payload === "object" ? payload : {})
  };
  normalizedPayload.answers = normalizeAnswerArray(normalizedPayload.answers || answers);

  if (shouldUseAppsScriptBridge()) {
    const data = await callAppsScriptBridge("autobidSaveAnswers", {
      sheetName: normalized.sheetName,
      rowNumber,
      row: rowNumber,
      answers: normalizedPayload.answers,
      payload: normalizedPayload
    });
    return {
      ok: true,
      saved_answers: Number(data.saved_answers || data.savedAnswers || normalizedPayload.answers.length),
      updated_at: data.updated_at || data.updatedAt || new Date().toISOString()
    };
  }

  const headers = await ensureAutoBidHeaders(normalized.spreadsheetId, normalized.sheetName);
  const now = new Date().toISOString();
  const updates = [
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.answers] + 1, rowNumber, JSON.stringify(normalizedPayload)),
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.status] + 1, rowNumber, "answers_ready"),
    cellUpdate(normalized.sheetName, headers[AUTO_BID_COLUMNS.updatedAt] + 1, rowNumber, now)
  ];

  await batchUpdateValues(normalized.spreadsheetId, updates);
  return {
    ok: true,
    saved_answers: normalizedPayload.answers.length,
    updated_at: now
  };
}

export async function readResumeFilePayload({
  spreadsheetId,
  sheetName,
  rowNumber,
  accept = [],
  resumeUrl = "",
  resume_url = "",
  rowValues = {},
  row_values = {},
  raw = []
}) {
  const normalized = normalizeSheetRequest({ spreadsheetId, sheetName, startRow: rowNumber, endRow: rowNumber });
  if (!shouldUseAppsScriptBridge()) {
    throw new Error("Resume file attach currently requires the Apps Script bridge so Drive links can be read safely.");
  }

  const values = rowValues && Object.keys(rowValues).length ? rowValues : row_values;
  const explicitResumeUrl = findResumeUrlFromRequest({
    resumeUrl,
    resume_url,
    rowValues: values,
    raw
  });

  const data = await callAppsScriptBridge("autobidReadResumeFile", {
    sheetName: normalized.sheetName,
    rowNumber,
    row: rowNumber,
    resumeUrl: explicitResumeUrl,
    resume_url: explicitResumeUrl,
    rowValues: values || {},
    row_values: values || {},
    raw: Array.isArray(raw) ? raw : [],
    accept
  });

  return writeResumePayloadToLocalFile(normalizeResumeFilePayload(data, rowNumber));
}

function shouldUseAppsScriptBridge() {
  return Boolean(String(GOOGLE_APPS_SCRIPT_WEB_APP_URL || "").trim());
}

async function callAppsScriptBridge(action, payload) {
  if (!GOOGLE_APPS_SCRIPT_SECRET) {
    throw new Error("GOOGLE_APPS_SCRIPT_SECRET is required when GOOGLE_APPS_SCRIPT_WEB_APP_URL is configured.");
  }

  const startedAt = Date.now();
  logBackendEvent("APPS_SCRIPT_REQUEST", {
    action,
    host: safeHost(GOOGLE_APPS_SCRIPT_WEB_APP_URL),
    connect_via: GOOGLE_APPS_SCRIPT_CONNECT_HOST,
    sheet_name: payload?.sheetName,
    row: payload?.rowNumber || payload?.row,
    start_row: payload?.startRow,
    end_row: payload?.endRow
  });

  let response;
  try {
    response = await requestAppsScript(GOOGLE_APPS_SCRIPT_WEB_APP_URL, {
      method: "POST",
      body: JSON.stringify({
        action,
        secret: GOOGLE_APPS_SCRIPT_SECRET,
        ...payload
      }),
      timeoutMs: GOOGLE_APPS_SCRIPT_TIMEOUT_MS
    });
  } catch (error) {
    logBackendEvent("APPS_SCRIPT_REQUEST_ERROR", {
      action,
      host: safeHost(GOOGLE_APPS_SCRIPT_WEB_APP_URL),
      duration_ms: Date.now() - startedAt,
      error
    }, { level: "error" });
    throw new Error(`Could not connect to Apps Script for ${action}: ${deepestErrorMessage(error)}`, {
      cause: error
    });
  }

  const text = response.text;
  let data = {};

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Apps Script returned non-JSON response: ${extractAppsScriptHtmlError(text)}`);
  }

  if (!response.ok || data.status === "error" || data.error) {
    throw new Error(data.message || data.error || `Apps Script request failed with ${response.status}`);
  }

  logBackendEvent("APPS_SCRIPT_RESPONSE", {
    action,
    status: response.status,
    final_host: safeHost(response.url),
    duration_ms: Date.now() - startedAt,
    jobs_count: Array.isArray(data.jobs) ? data.jobs.length : undefined,
    rows_count: Array.isArray(data.rows) ? data.rows.length : undefined
  });

  return data;
}

function requestAppsScript(url, {
  method = "GET",
  body = "",
  timeoutMs = 30_000,
  redirects = 0,
  deadlineAt = Date.now() + timeoutMs
} = {}) {
  const target = new URL(url);
  if (target.protocol !== "https:") {
    return Promise.reject(new Error("Apps Script URL must use HTTPS."));
  }

  const remainingMs = deadlineAt - Date.now();
  if (remainingMs <= 0) {
    const error = new Error(`Apps Script request timed out after ${timeoutMs} ms.`);
    error.code = "ETIMEDOUT";
    return Promise.reject(error);
  }

  return new Promise((resolve, reject) => {
    const headers = method === "GET" || !body
      ? {}
      : {
          "Content-Type": "application/json",
          "Content-Length": Buffer.byteLength(body)
        };
    const request = httpsRequest(target, {
      method,
      headers,
      agent: appsScriptAgent
    }, (response) => {
      const status = Number(response.statusCode || 0);
      const location = response.headers.location;

      if ([301, 302, 303, 307, 308].includes(status) && location) {
        request.setTimeout(0);
        response.resume();
        if (redirects >= 5) {
          reject(new Error("Apps Script returned too many redirects."));
          return;
        }

        const redirectedUrl = new URL(location, target);
        if (!isAllowedAppsScriptHost(redirectedUrl.hostname)) {
          reject(new Error(`Apps Script redirected to an unexpected host: ${redirectedUrl.hostname}`));
          return;
        }

        const preserveMethod = status === 307 || status === 308;
        requestAppsScript(redirectedUrl, {
          method: preserveMethod ? method : "GET",
          body: preserveMethod ? body : "",
          timeoutMs,
          redirects: redirects + 1,
          deadlineAt
        }).then(resolve, reject);
        return;
      }

      const chunks = [];
      response.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
      response.on("end", () => resolve({
        ok: status >= 200 && status < 300,
        status,
        text: Buffer.concat(chunks).toString("utf8"),
        url: target.toString()
      }));
      response.on("error", reject);
    });

    request.setTimeout(remainingMs, () => {
      const error = new Error(`Apps Script request timed out after ${timeoutMs} ms.`);
      error.code = "ETIMEDOUT";
      request.destroy(error);
    });
    request.on("error", reject);
    if (body && method !== "GET") request.write(body);
    request.end();
  });
}

function lookupAppsScriptAddress(hostname, options, callback) {
  const routeHost = hostname === "script.google.com" && GOOGLE_APPS_SCRIPT_CONNECT_HOST
    ? GOOGLE_APPS_SCRIPT_CONNECT_HOST
    : hostname;
  dnsLookup(routeHost, options, callback);
}

function isAllowedAppsScriptHost(hostname) {
  const value = String(hostname || "").toLowerCase();
  return value === "script.google.com" || value === "script.googleusercontent.com" || value.endsWith(".script.googleusercontent.com");
}

function safeHost(value) {
  try {
    return new URL(value).hostname;
  } catch {
    return "invalid-url";
  }
}

function deepestErrorMessage(error) {
  let current = error;
  let message = "Unknown connection error";
  const visited = new Set();
  while (current && !visited.has(current)) {
    visited.add(current);
    if (current.message) message = current.message;
    current = current.cause;
  }
  return message;
}

function extractAppsScriptHtmlError(text) {
  const raw = String(text || "");
  const divMatch = raw.match(/<div[^>]*>([\s\S]*?)<\/div>/i);
  const titleMatch = raw.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const value = divMatch?.[1] || titleMatch?.[1] || raw.slice(0, 500);
  return decodeHtmlEntities(value).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 500);
}

function decodeHtmlEntities(value) {
  return String(value || "")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, "\"")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function normalizeBridgeJobs(jobs) {
  return jobs
    .map((job) => {
      const raw = Array.isArray(job.raw) ? job.raw : [];
      const values = job.values || job.rowValues || job.row_values || {};
      return {
        rowNumber: Number(job.rowNumber || job.row_number || job.row || 0),
        values: withFixedColumnValues(values, raw),
        raw,
        url: findFirstUrlInText(job.url) ||
          findFirstUrlInText(job.jobUrl) ||
          findFirstUrlInText(job.job_url) ||
          findFirstUrlInText(job.applyUrl) ||
          findFirstUrlInText(job.apply_url) ||
          findFirstUrlInText(job.applicationUrl) ||
          findFirstUrlInText(job.application_url) ||
          findFirstUrlInText(job.link) ||
          findFirstUrlInSheetValues(values, raw)
      };
    })
    .filter((job) => job.rowNumber && job.url);
}

function normalizePendingQuestionRows(rows, normalized) {
  return rows
    .map((row) => {
      const raw = Array.isArray(row.raw) ? row.raw : [];
      const questions = normalizeQuestionPayload(row.questions || row.payload || parseQuestionPayload(row.questions_raw || row.questionsRaw || row.raw_questions || ""));
      return {
        spreadsheetId: row.spreadsheetId || row.spreadsheet_id || normalized.spreadsheetId,
        sheetName: row.sheetName || row.sheet_name || normalized.sheetName,
        rowNumber: Number(row.rowNumber || row.row_number || row.row || 0),
        row: Number(row.rowNumber || row.row_number || row.row || 0),
        url: String(row.url || row.applyUrl || row.apply_url || questions.row?.url || questions.page?.url || "").trim(),
        values: withFixedColumnValues(row.values || {}, raw),
        raw,
        questions
      };
    })
    .filter((row) => row.rowNumber && row.questions?.fields?.length);
}

function parseQuestionPayload(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  try {
    return normalizeQuestionPayload(JSON.parse(text));
  } catch {
    return null;
  }
}

function normalizeQuestionPayload(value) {
  if (!value || typeof value !== "object") return null;
  const fields = Array.isArray(value.fields)
    ? value.fields
      .map((field) => ({
        ...field,
        field_id: String(field.field_id || field.id || ""),
        id: String(field.id || field.field_id || ""),
        question: String(field.question || field.label || ""),
        option: String(field.option || ""),
        label: String(field.label || field.question || ""),
        type: String(field.type || ""),
        options: Array.isArray(field.options) ? field.options.map((option) => String(option || "")).filter(Boolean) : []
      }))
      .filter((field) => field.field_id || field.id)
    : [];
  return { ...value, fields };
}

function withFixedColumnValues(values, raw) {
  const result = { ...(values || {}) };
  const columnG = getRawColumnValue(raw, TAILORED_RESUME_COLUMN);
  const columnJ = getRawColumnValue(raw, RESUME_FILE_COLUMN);
  const columnM = getRawColumnValue(raw, JOB_DESCRIPTION_COLUMN);

  if (columnG && !String(result.column_g || "").trim()) result.column_g = columnG;
  if (columnJ && !String(result.column_j || "").trim()) result.column_j = columnJ;
  if (columnM && !String(result.column_m || "").trim()) result.column_m = columnM;

  return result;
}

function getRawColumnValue(raw, oneBasedColumn) {
  if (!Array.isArray(raw)) return "";
  return String(raw[oneBasedColumn - 1] || "").trim();
}

function normalizeSheetRequest({ spreadsheetId, sheetName, startRow, endRow }) {
  const resolvedSpreadsheetId = normalizeSpreadsheetId(spreadsheetId || GOOGLE_SHEETS_SPREADSHEET_ID || "");
  const resolvedSheetName = String(sheetName || "").trim();
  const firstRow = Math.max(2, Number(startRow || 2));
  const lastRow = Math.max(firstRow, Number(endRow || firstRow));

  if (!resolvedSpreadsheetId && !shouldUseAppsScriptBridge()) throw new Error("Google spreadsheet id is required. Set GOOGLE_SHEETS_SPREADSHEET_ID or provide it in the extension.");
  if (!resolvedSheetName) throw new Error("Google sheet tab name is required.");
  return {
    spreadsheetId: resolvedSpreadsheetId,
    sheetName: resolvedSheetName,
    startRow: firstRow,
    endRow: lastRow
  };
}

function normalizeSpreadsheetId(value) {
  const text = String(value || "").trim();
  const match = text.match(/\/spreadsheets\/d\/([^/?#]+)/);
  return match ? match[1] : text;
}

async function ensureAutoBidHeaders(spreadsheetId, sheetName) {
  const headerRange = `${quoteSheetName(sheetName)}!A1:ZZ1`;
  const data = await getValues(spreadsheetId, headerRange);
  const headerRow = data.values?.[0] || [];
  const headers = normalizeHeaders(headerRow);
  let nextColumn = headerRow.length + 1;
  const updates = [];

  for (const header of Object.values(AUTO_BID_COLUMNS)) {
    if (headers[header] !== undefined) continue;
    headers[header] = nextColumn - 1;
    updates.push(cellUpdate(sheetName, nextColumn, 1, header));
    nextColumn += 1;
  }

  if (updates.length > 0) await batchUpdateValues(spreadsheetId, updates);
  return headers;
}

async function getValues(spreadsheetId, range) {
  const token = await getAccessToken();
  const response = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  return parseGoogleResponse(response);
}

async function batchUpdateValues(spreadsheetId, updates) {
  const token = await getAccessToken();
  const response = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      valueInputOption: "RAW",
      data: updates.map((update) => ({
        range: update.range,
        majorDimension: "ROWS",
        values: [[update.value]]
      }))
    })
  });
  return parseGoogleResponse(response);
}

async function getAccessToken() {
  if (tokenCache && tokenCache.expiresAt > Date.now() + 60_000) return tokenCache.token;
  if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY) {
    throw new Error("Google Sheets service account is not configured. Set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY.");
  }

  const now = Math.floor(Date.now() / 1000);
  const header = base64UrlJson({ alg: "RS256", typ: "JWT" });
  const claim = base64UrlJson({
    iss: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: SHEETS_SCOPE,
    aud: TOKEN_URL,
    iat: now,
    exp: now + 3600
  });
  const unsigned = `${header}.${claim}`;
  const signature = createSign("RSA-SHA256").update(unsigned).sign(GOOGLE_PRIVATE_KEY, "base64url");
  const assertion = `${unsigned}.${signature}`;

  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion
    })
  });
  const data = await parseGoogleResponse(response);
  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + Number(data.expires_in || 3600) * 1000
  };
  return tokenCache.token;
}

async function parseGoogleResponse(response) {
  const text = await response.text();
  const data = text ? JSON.parse(text) : {};
  if (!response.ok) {
    throw new Error(data.error?.message || data.error_description || `Google Sheets request failed with ${response.status}`);
  }
  return data;
}

function normalizeHeaders(row) {
  const headers = {};
  row.forEach((value, index) => {
    const key = normalizeHeader(value);
    if (key && headers[key] === undefined) headers[key] = index;
  });
  return headers;
}

function rowToObject(headers, row) {
  return Object.fromEntries(
    Object.entries(headers).map(([key, index]) => [key, row[index] || ""])
  );
}

function findJobUrl(headers, row) {
  for (const header of JOB_URL_HEADERS) {
    const index = headers[normalizeHeader(header)];
    const value = index === undefined ? "" : row[index];
    const url = findFirstUrlInText(value);
    if (url) return url;
  }

  return findFirstUrlInSheetValues({}, row);
}

function parseSheetAnswers(raw) {
  const text = String(raw || "").trim();
  if (!text) return [];

  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) return normalizeAnswerArray(parsed);
    if (Array.isArray(parsed.answers)) return normalizeAnswerArray(parsed.answers);
    if (parsed && typeof parsed === "object") {
      return Object.entries(parsed).map(([field_id, value]) => ({
        field_id,
        value: String(value ?? ""),
        source: "sheet"
      })).filter((answer) => answer.field_id && answer.value);
    }
  } catch {
    return [];
  }

  return [];
}

function normalizeAnswerArray(answers) {
  return answers
    .map((answer) => ({
      field_id: String(answer.field_id || answer.id || ""),
      question: String(answer.question || ""),
      option: String(answer.option || ""),
      value: String(answer.value ?? answer.answer ?? ""),
      source: "sheet"
    }))
    .filter((answer) => answer.field_id);
}

function normalizeResumeFilePayload(data, rowNumber) {
  const base64 = String(data.base64 || data.file_base64 || data.content || "");
  return {
    row_number: Number(data.row_number || data.rowNumber || rowNumber),
    filename: String(data.filename || data.name || "resume.pdf"),
    mime_type: String(data.mime_type || data.mimeType || "application/pdf"),
    size: Number(data.size || data.bytes || 0),
    source_url: String(data.source_url || data.sourceUrl || data.url || ""),
    base64
  };
}

function findResumeUrlFromRequest({ resumeUrl, resume_url, rowValues = {}, raw = [] }) {
  const direct = findFirstUrlInText(resumeUrl) || findFirstUrlInText(resume_url);
  if (direct) return direct;

  for (const key of RESUME_FILE_VALUE_KEYS) {
    const url = findFirstUrlInText(rowValues?.[key]);
    if (url) return url;
  }

  const columnJ = getRawColumnValue(raw, RESUME_FILE_COLUMN);
  const columnJUrl = findFirstUrlInText(columnJ);
  if (columnJUrl) return columnJUrl;

  return "";
}

function findFirstUrlInText(value) {
  const match = String(value || "").match(/https?:\/\/[^\s"'<>),]+/i);
  return match ? match[0] : "";
}

function findFirstUrlInSheetValues(values = {}, raw = []) {
  for (const value of Object.values(values || {})) {
    const url = findFirstUrlInText(value);
    if (url) return url;
  }

  for (const value of raw || []) {
    const url = findFirstUrlInText(value);
    if (url) return url;
  }

  return "";
}

async function writeResumePayloadToLocalFile(payload) {
  const base64 = String(payload.base64 || "").replace(/^data:[^,]+,/, "");
  if (!base64) return payload;

  const bytes = Buffer.from(base64, "base64");
  if (!bytes.length) return payload;

  const hash = createHash("sha256")
    .update(String(payload.source_url || ""))
    .update(String(payload.filename || ""))
    .update(bytes)
    .digest("hex")
    .slice(0, 24);
  const filename = ensureResumeFilenameExtension(sanitizeLocalFilename(payload.filename || "resume.pdf"), payload.mime_type);
  const directory = path.join(os.tmpdir(), "autobid-resumes", hash);
  const localPath = path.join(directory, filename);

  await mkdir(directory, { recursive: true });
  await writeFile(localPath, bytes);

  return {
    ...payload,
    size: payload.size || bytes.length,
    local_path: localPath
  };
}

function sanitizeLocalFilename(value) {
  return String(value || "resume.pdf")
    .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "_")
    .replace(/\s+/g, " ")
    .trim() || "resume.pdf";
}

function ensureResumeFilenameExtension(filename, mimeType) {
  if (/\.[a-z0-9]{2,8}$/i.test(filename)) return filename;
  const extension = mimeTypeToExtension(mimeType);
  return `${filename}.${extension}`;
}

function mimeTypeToExtension(mimeType) {
  const value = String(mimeType || "").toLowerCase();
  if (value.includes("jpeg")) return "jpg";
  if (value.includes("png")) return "png";
  if (value.includes("msword")) return "doc";
  if (value.includes("wordprocessingml")) return "docx";
  if (value.includes("rtf")) return "rtf";
  return "pdf";
}

function cellUpdate(sheetName, column, row, value) {
  return {
    range: `${quoteSheetName(sheetName)}!${columnName(column)}${row}:${columnName(column)}${row}`,
    value
  };
}

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replace(/'/g, "''")}'`;
}

function columnName(index) {
  let value = Number(index);
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function isHttpUrl(value) {
  return /^https?:\/\/\S+/i.test(String(value || "").trim());
}

function base64UrlJson(value) {
  return Buffer.from(JSON.stringify(value)).toString("base64url");
}
