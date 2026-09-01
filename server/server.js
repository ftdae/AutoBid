import http from "node:http";
import { createToken, hashPassword, normalizeEmail, readToken, verifyPassword } from "./auth/security.js";
import { ensureQuestions, loadCacheAnswers, saveAiAnswers } from "./assist/cache.js";
import { normalizeFields, normalizePage, shouldAnswerWithAi } from "./assist/field-policy.js";
import { generateAiAnswers } from "./assist/ai.js";
import { DEV_AUTH_BYPASS, DEV_USER_EMAIL, OPENAI_ROUTE_ENABLED, PORT } from "./config.js";
import { pool } from "./db/pool.js";
import { ensureSchema } from "./db/schema.js";
import { readJson, sendJson, setCorsHeaders } from "./http/json.js";
import { buildStaticAnswers } from "./profiles/static-fields.js";
import {
  completeOutlookAuthorization,
  createOutlookAuthorization,
  disconnectOutlook,
  findLatestVerificationCode,
  getOutlookConnection,
  listVerificationMessages,
  markOutlookMessageRead
} from "./outlook/microsoft-graph.js";
import { listPendingQuestionRows, listSheetJobs, readAnswerPayload, readResumeFilePayload, writeAnswerPayload, writeQuestionPayload } from "./sheets/google-sheets.js";
import { id } from "./utils/id.js";
import { beginHttpRequestLog, getHttpRequestId, logBackendEvent, logHttpRequestBody } from "./utils/logger.js";
import { hashText, normalizeUrl, safeDomain, trimForPrompt } from "./utils/text.js";
import { serializeProfile, serializeUser } from "./users/serializers.js";

let databaseAvailable = false;
let databaseRecoveryPromise = null;
let lastDatabaseRecoveryAttemptAt = 0;
const ASSIST_JOB_PENDING_TTL_MS = 10 * 60 * 1000;
const ASSIST_JOB_RESULT_TTL_MS = 20 * 60 * 1000;
const assistJobs = new Map();

try {
  await ensureSchema(pool);
  databaseAvailable = true;
} catch (error) {
  if (!DEV_AUTH_BYPASS) throw error;
  console.warn(`[AutoBid] PostgreSQL unavailable; starting in development fallback mode: ${error.message || String(error)}`);
}

const DEV_FALLBACK_USER = {
  id: "abu_dev_local",
  first_name: "Dev",
  last_name: "User",
  email: DEV_USER_EMAIL,
  password: "",
  timezone: "UTC",
  active: true,
  created_at: new Date(),
  updated_at: new Date()
};

const DEV_FALLBACK_PROFILE = {
  id: "abp_dev_default",
  user_id: DEV_FALLBACK_USER.id,
  name: "Development profile",
  static_fields: {},
  resume_text: "",
  preferences: {},
  profile_version: 1,
  active: true,
  created_at: new Date(),
  updated_at: new Date()
};

const server = http.createServer(async (req, res) => {
  beginHttpRequestLog(req, res);
  try {
    await handleRequest(req, res);
  } catch (error) {
    const status = Number(error.status || error.statusCode || 500);
    if (status >= 500) {
      logBackendEvent("HTTP_HANDLER_ERROR", { error }, {
        requestId: getHttpRequestId(res),
        level: "error"
      });
    }
    sendJson(res, status, {
      data: null,
      errors: [{ code: error.code || "server_error", message: error.message || "Internal server error" }]
    });
  }
});

server.listen(PORT, () => {
  logBackendEvent("BACKEND_STARTED", {
    url: `http://localhost:${PORT}`,
    database_available: databaseAvailable
  });
});

async function handleRequest(req, res) {
  setCorsHeaders(res);

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
  const pathname = url.pathname.replace(/\/+$/, "") || "/";

  if (!pathname.startsWith("/api/auto-bid")) {
    sendJson(res, 404, { data: null, errors: [{ code: "not_found", message: "Route not found" }] });
    return;
  }

  const body = await readJson(req);
  logHttpRequestBody(res, body);

  if (req.method === "POST" && pathname === "/api/auto-bid/auth/signup") {
    return signup(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/auth/login") {
    return login(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/auth/dev-session") {
    return devSession(res);
  }

  const user = await authenticate(req);
  if (!user) {
    sendJson(res, 401, {
      data: null,
      errors: [{ code: "auth_required", message: "Authentication required" }]
    });
    return;
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/auth/session") {
    sendJson(res, 200, {
      data: { token: createToken(user.id), user: serializeUser(user) },
      errors: null
    });
    return;
  }

  if (req.method === "GET" && pathname === "/api/auto-bid/profiles") {
    await recoverDatabaseConnection();
    if (!databaseAvailable && DEV_AUTH_BYPASS) {
      sendJson(res, 200, { data: [serializeProfile(DEV_FALLBACK_PROFILE)], errors: null });
      return;
    }

    const { rows } = await pool.query(
      "select * from auto_bid_profiles where user_id = $1 and active = true order by updated_at desc",
      [user.id]
    );
    sendJson(res, 200, { data: rows.map(serializeProfile), errors: null });
    return;
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/profiles") {
    return createProfile(res, user, body);
  }

  const profileMatch = pathname.match(/^\/api\/auto-bid\/profiles\/([^/]+)$/);
  if (profileMatch && req.method === "PATCH") {
    return updateProfile(res, user, profileMatch[1], body);
  }
  if (profileMatch && req.method === "DELETE") {
    return deleteProfile(res, user, profileMatch[1]);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/assist/start") {
    return startAssistJob(res, user, body);
  }

  const assistJobMatch = pathname.match(/^\/api\/auto-bid\/assist\/jobs\/([^/]+)$/);
  if ((req.method === "GET" || req.method === "POST") && assistJobMatch) {
    return getAssistJob(res, user, assistJobMatch[1]);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/assist") {
    return assist(res, user, body);
  }

  if (req.method === "GET" && pathname === "/api/auto-bid/outlook/connection") {
    const connection = await getOutlookConnection(pool, user.id);
    sendJson(res, 200, { data: connection, errors: null });
    return;
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/outlook/oauth/start") {
    const authorization = createOutlookAuthorization({ userId: user.id, redirectUri: body.redirect_uri });
    sendJson(res, 200, { data: authorization, errors: null });
    return;
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/outlook/oauth/callback") {
    const connection = await completeOutlookAuthorization(pool, {
      userId: user.id,
      code: body.code,
      state: body.state,
      redirectUri: body.redirect_uri
    });
    sendJson(res, 200, { data: connection, errors: null });
    return;
  }

  if (req.method === "DELETE" && pathname === "/api/auto-bid/outlook/connection") {
    const connection = await disconnectOutlook(pool, user.id);
    sendJson(res, 200, { data: connection, errors: null });
    return;
  }

  if (req.method === "GET" && pathname === "/api/auto-bid/outlook/messages") {
    const messages = await listVerificationMessages(pool, user.id, {
      top: url.searchParams.get("top"),
      domain: url.searchParams.get("domain")
    });
    sendJson(res, 200, { data: messages, errors: null });
    return;
  }

  if (req.method === "GET" && pathname === "/api/auto-bid/outlook/latest-code") {
    const result = await findLatestVerificationCode(pool, user.id, {
      top: url.searchParams.get("top"),
      domain: url.searchParams.get("domain")
    });
    sendJson(res, 200, { data: result, errors: null });
    return;
  }

  const outlookReadMatch = pathname.match(/^\/api\/auto-bid\/outlook\/messages\/([^/]+)\/read$/);
  if (req.method === "POST" && outlookReadMatch) {
    const result = await markOutlookMessageRead(pool, user.id, decodeURIComponent(outlookReadMatch[1]));
    sendJson(res, 200, { data: result, errors: null });
    return;
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/jobs") {
    return listJobsFromSheet(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/questions") {
    return saveSheetQuestions(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/pending-questions") {
    return listPendingSheetQuestions(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/answers") {
    return getSheetAnswers(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/save-answers") {
    return saveSheetAnswers(res, body);
  }

  if (req.method === "POST" && pathname === "/api/auto-bid/sheets/resume-file") {
    return getSheetResumeFile(res, body);
  }

  const draftMatch = pathname.match(/^\/api\/auto-bid\/drafts\/([^/]+)\/status$/);
  if (draftMatch && req.method === "POST") {
    return updateDraftStatus(res, user, draftMatch[1], body);
  }

  sendJson(res, 404, { data: null, errors: [{ code: "not_found", message: "Route not found" }] });
}

async function signup(res, body) {
  const email = normalizeEmail(body.email);
  if (!body.first_name || !body.last_name || !email || !body.password || String(body.password).length < 8) {
    sendJson(res, 400, {
      data: null,
      errors: [{ code: "validation_error", message: "First name, last name, valid email, and 8 character password are required" }]
    });
    return;
  }

  try {
    const { rows } = await pool.query(
      `insert into auto_bid_users
        (id, first_name, last_name, email, password, timezone, active)
       values ($1, $2, $3, $4, $5, $6, true)
       returning *`,
      [
        id("abu"),
        String(body.first_name).trim(),
        String(body.last_name).trim(),
        email,
        hashPassword(String(body.password)),
        body.timezone || "UTC"
      ]
    );
    const user = rows[0];
    sendJson(res, 200, {
      data: { token: createToken(user.id), user: serializeUser(user) },
      errors: null
    });
  } catch (error) {
    if (error.code === "23505") {
      sendJson(res, 400, {
        data: null,
        errors: [{ code: "email_exists", message: "Email already exists" }]
      });
      return;
    }
    throw error;
  }
}

async function login(res, body) {
  const email = normalizeEmail(body.email);
  const { rows } = await pool.query(
    "select * from auto_bid_users where email = $1 and active = true limit 1",
    [email]
  );
  const user = rows[0];

  if (!user || !verifyPassword(String(body.password || ""), user.password)) {
    sendJson(res, 401, {
      data: null,
      errors: [{ code: "invalid_credentials", message: "Invalid credentials" }]
    });
    return;
  }

  sendJson(res, 200, {
    data: { token: createToken(user.id), user: serializeUser(user) },
    errors: null
  });
}

async function devSession(res) {
  if (!DEV_AUTH_BYPASS) {
    sendJson(res, 404, { data: null, errors: [{ code: "not_found", message: "Route not found" }] });
    return;
  }

  await recoverDatabaseConnection();
  if (!databaseAvailable) {
    sendJson(res, 200, {
      data: {
        token: createToken(DEV_FALLBACK_USER.id),
        user: serializeUser(DEV_FALLBACK_USER),
        profile: serializeProfile(DEV_FALLBACK_PROFILE)
      },
      errors: null
    });
    return;
  }

  const { user, profile } = await ensureDevAccount();
  sendJson(res, 200, {
    data: {
      token: createToken(user.id),
      user: serializeUser(user),
      profile: serializeProfile(profile)
    },
    errors: null
  });
}

async function ensureDevAccount() {
  if (!databaseAvailable) {
    return { user: DEV_FALLBACK_USER, profile: DEV_FALLBACK_PROFILE };
  }

  const userId = "abu_dev_local";
  const profileId = "abp_dev_default";

  const { rows: userRows } = await pool.query(
    `insert into auto_bid_users
      (id, first_name, last_name, email, password, timezone, active)
     values ($1, 'Dev', 'User', $2, $3, 'UTC', true)
     on conflict (id) do update
       set active = true,
           email = excluded.email,
           updated_at = now()
     returning *`,
    [userId, DEV_USER_EMAIL, hashPassword("dev-password-not-used")]
  );

  const { rows: profileRows } = await pool.query(
    `insert into auto_bid_profiles
      (id, user_id, name, static_fields, resume_text, preferences, profile_version, active)
     values ($1, $2, 'Development profile', '{}'::jsonb, '', '{}'::jsonb, 1, true)
     on conflict (id) do update
       set active = true,
           user_id = excluded.user_id,
           updated_at = now()
     returning *`,
    [profileId, userId]
  );

  return { user: userRows[0], profile: profileRows[0] };
}

async function createProfile(res, user, body) {
  await recoverDatabaseConnection();
  if (!databaseAvailable && DEV_AUTH_BYPASS) {
    sendProfileDatabaseUnavailable(res);
    return;
  }

  const { rows } = await pool.query(
    `insert into auto_bid_profiles
      (id, user_id, name, static_fields, resume_text, preferences, profile_version, active)
     values ($1, $2, $3, $4::jsonb, $5, $6::jsonb, 1, true)
     returning *`,
    [
      id("abp"),
      user.id,
      String(body.name || "Default profile").trim(),
      JSON.stringify(body.static_fields && typeof body.static_fields === "object" ? body.static_fields : {}),
      String(body.resume_text || ""),
      JSON.stringify(body.preferences && typeof body.preferences === "object" ? body.preferences : {})
    ]
  );

  sendJson(res, 200, { data: serializeProfile(rows[0]), errors: null });
}

async function updateProfile(res, user, profileId, body) {
  await recoverDatabaseConnection();
  if (!databaseAvailable && DEV_AUTH_BYPASS) {
    sendProfileDatabaseUnavailable(res);
    return;
  }

  const existing = await loadOwnedProfile(profileId, user.id);
  if (!existing) {
    sendJson(res, 404, { data: null, errors: [{ code: "profile_not_found", message: "Profile not found" }] });
    return;
  }

  const next = {
    name: body.name !== undefined ? String(body.name).trim() : existing.name,
    static_fields: body.static_fields !== undefined && body.static_fields && typeof body.static_fields === "object" ? body.static_fields : existing.static_fields,
    resume_text: body.resume_text !== undefined ? String(body.resume_text || "") : existing.resume_text,
    preferences: body.preferences !== undefined && body.preferences && typeof body.preferences === "object" ? body.preferences : existing.preferences
  };

  const { rows } = await pool.query(
    `update auto_bid_profiles
       set name = $1,
           static_fields = $2::jsonb,
           resume_text = $3,
           preferences = $4::jsonb,
           profile_version = profile_version + 1,
           updated_at = now()
     where id = $5 and user_id = $6 and active = true
     returning *`,
    [
      next.name,
      JSON.stringify(next.static_fields || {}),
      next.resume_text || "",
      JSON.stringify(next.preferences || {}),
      profileId,
      user.id
    ]
  );

  sendJson(res, 200, { data: serializeProfile(rows[0]), errors: null });
}

async function deleteProfile(res, user, profileId) {
  await recoverDatabaseConnection();
  if (!databaseAvailable && DEV_AUTH_BYPASS) {
    sendProfileDatabaseUnavailable(res);
    return;
  }

  await pool.query(
    "update auto_bid_profiles set active = false, updated_at = now() where id = $1 and user_id = $2",
    [profileId, user.id]
  );
  sendJson(res, 200, { data: { ok: true }, errors: null });
}

function sendProfileDatabaseUnavailable(res) {
  sendJson(res, 503, {
    data: null,
    errors: [{
      code: "database_unavailable",
      message: "PostgreSQL is unavailable. The profile was not saved."
    }]
  });
}

async function recoverDatabaseConnection() {
  if (databaseAvailable) return true;
  if (databaseRecoveryPromise) return databaseRecoveryPromise;

  const now = Date.now();
  if (now - lastDatabaseRecoveryAttemptAt < 2000) return false;
  lastDatabaseRecoveryAttemptAt = now;

  databaseRecoveryPromise = ensureSchema(pool)
    .then(() => {
      databaseAvailable = true;
      console.info("[AutoBid] PostgreSQL connection recovered.");
      return true;
    })
    .catch((error) => {
      console.warn(`[AutoBid] PostgreSQL recovery failed: ${error.message || String(error)}`);
      return false;
    })
    .finally(() => {
      databaseRecoveryPromise = null;
    });

  return databaseRecoveryPromise;
}

async function assist(res, user, body) {
  const data = await buildAssistResponse(user, body, {
    requestId: getHttpRequestId(res),
    mode: "synchronous"
  });
  sendJson(res, 200, { data, errors: null });
}

function startAssistJob(res, user, body) {
  cleanupAssistJobs();
  const jobId = id("abaj");
  const parentRequestId = getHttpRequestId(res);
  const job = {
    id: jobId,
    user_id: user.id,
    status: "pending",
    created_at: Date.now(),
    updated_at: Date.now(),
    data: null,
    error: null
  };
  assistJobs.set(jobId, job);
  logBackendEvent("ASSIST_JOB_QUEUED", {
    job_id: jobId,
    parent_request_id: parentRequestId,
    profile_id: body.profile_id || "",
    page_url: body.page?.url || "",
    fields: Array.isArray(body.fields) ? body.fields.length : 0
  }, { requestId: jobId });

  Promise.resolve()
    .then(() => {
      logBackendEvent("ASSIST_JOB_STARTED", {
        job_id: jobId,
        parent_request_id: parentRequestId
      }, { requestId: jobId });
      return buildAssistResponse(user, body, {
        requestId: jobId,
        parentRequestId,
        jobId,
        mode: "background"
      });
    })
    .then((data) => {
      job.status = "complete";
      job.data = data;
      job.updated_at = Date.now();
      logBackendEvent("ASSIST_JOB_COMPLETED", {
        job_id: jobId,
        duration_ms: job.updated_at - job.created_at,
        answers: Array.isArray(data?.answers) ? data.answers.length : 0,
        warnings: Array.isArray(data?.warnings) ? data.warnings : []
      }, { requestId: jobId });
    })
    .catch((error) => {
      job.status = "error";
      job.error = error.message || String(error);
      job.updated_at = Date.now();
      logBackendEvent("ASSIST_JOB_FAILED", {
        job_id: jobId,
        duration_ms: job.updated_at - job.created_at,
        error
      }, { requestId: jobId, level: "error" });
    });

  sendJson(res, 202, {
    data: {
      job_id: jobId,
      status: "pending"
    },
    errors: null
  });
}

function getAssistJob(res, user, jobId) {
  cleanupAssistJobs();
  const job = assistJobs.get(jobId);
  if (!job || job.user_id !== user.id) {
    sendJson(res, 404, { data: null, errors: [{ code: "assist_job_not_found", message: "Assist job not found" }] });
    return;
  }

  if (job.status === "complete") {
    sendJson(res, 200, {
      data: {
        job_id: job.id,
        status: job.status,
        result: job.data
      },
      errors: null
    });
    return;
  }

  if (job.status === "error") {
    sendJson(res, 200, {
      data: {
        job_id: job.id,
        status: job.status,
        error: job.error || "Assist job failed"
      },
      errors: null
    });
    return;
  }

  sendJson(res, 200, {
    data: {
      job_id: job.id,
      status: job.status
    },
    errors: null
  });
}

function cleanupAssistJobs() {
  const now = Date.now();
  for (const [jobId, job] of assistJobs.entries()) {
    const ttl = job.status === "pending" ? ASSIST_JOB_PENDING_TTL_MS : ASSIST_JOB_RESULT_TTL_MS;
    if (now - job.updated_at > ttl) assistJobs.delete(jobId);
  }
}

async function buildAssistResponse(user, body, logContext = {}) {
  const profile = await loadOwnedProfile(body.profile_id, user.id);
  if (!profile) {
    const error = new Error("Profile not found");
    error.code = "profile_not_found";
    error.status = 404;
    throw error;
  }

  const page = normalizePage(body.page || {});
  const normalizedUrl = normalizeUrl(page.url);
  const domain = page.domain || safeDomain(page.url);
  const jobHash = hashText([domain, normalizedUrl, page.title, page.job_title, trimForPrompt(page.text, 6000)].join("\n"));
  const fields = normalizeFields(Array.isArray(body.fields) ? body.fields : [], domain, normalizedUrl);

  if (databaseAvailable) await ensureQuestions(pool, fields, domain, normalizedUrl);

  const staticAnswers = buildStaticAnswers(fields, profile.static_fields || {});
  const cache = databaseAvailable
    ? await loadCacheAnswers(pool, fields, profile, jobHash, staticAnswers)
    : { answers: new Map(), hits: 0 };
  const fieldsForAi = OPENAI_ROUTE_ENABLED
    ? fields.filter((field) => shouldAnswerWithAi(field) && !staticAnswers.has(field.id) && !cache.answers.has(field.id))
    : [];
  const warnings = [];
  logBackendEvent("ASSIST_PLAN", {
    job_id: logContext.jobId || null,
    mode: logContext.mode || "unknown",
    provider: OPENAI_ROUTE_ENABLED ? "openai" : "disabled",
    route_position: OPENAI_ROUTE_ENABLED ? 2 : null,
    openai_enabled: OPENAI_ROUTE_ENABLED,
    page: {
      url: page.url,
      title: page.title,
      job_title: page.job_title,
      domain
    },
    fields_received: fields.length,
    static_answers: staticAnswers.size,
    cache_hits: cache.hits,
    openai_fields: fieldsForAi.map((field) => ({
      field_id: field.id,
      label: field.label,
      type: field.type,
      required: field.required,
      options: field.options
    }))
  }, { requestId: logContext.requestId });
  const aiAnswers = fieldsForAi.length
    ? await generateAiAnswers(fieldsForAi, profile, page, jobHash, warnings, logContext)
    : [];

  if (databaseAvailable) await saveAiAnswers(pool, aiAnswers, fields, profile, jobHash);

  const answers = [
    ...Array.from(staticAnswers.values()),
    ...Array.from(cache.answers.values()),
    ...aiAnswers.map((answer) => ({
      field_id: answer.field_id,
      value: answer.value,
      source: "ai",
      provider: answer.provider || null,
      model: answer.model || null,
      estimated_request_cost_usd: answer.estimated_request_cost_usd ?? null,
      cache_scope: answer.cache_scope,
      confidence: answer.confidence,
      warning: answer.warning || null
    }))
  ];

  const fieldSnapshot = fields.map((field) => ({
    id: field.id,
    label: field.label,
    name: field.name,
    type: field.type,
    required: field.required,
    options: field.options,
    question_hash: field.question_hash,
    cache_scope: field.cache_scope
  }));
  const draftId = id("abd");

  if (databaseAvailable) {
    await pool.query(
      `insert into auto_bid_application_drafts
        (id, user_id, profile_id, domain, url, normalized_url, job_hash, form_hash, field_snapshot, answers_json, status)
       values ($1, $2, $3, $4, $5, $6, $7, $8, $9::jsonb, $10::jsonb, 'draft')`,
      [
        draftId,
        user.id,
        profile.id,
        domain,
        page.url,
        normalizedUrl,
        jobHash,
        hashText(fields.map((field) => field.question_hash).sort().join("|")),
        JSON.stringify(fieldSnapshot),
        JSON.stringify(answers)
      ]
    );
  }

  const result = {
    draft_id: draftId,
    profile: {
      id: profile.id,
      name: profile.name,
      profile_version: profile.profile_version
    },
    answers,
    cache: {
      hits: cache.hits,
      misses: fieldsForAi.length
    },
    warnings
  };
  logBackendEvent("ASSIST_RESULT", {
    job_id: logContext.jobId || null,
    draft_id: draftId,
    answers: result.answers,
    cache: result.cache,
    warnings
  }, { requestId: logContext.requestId });
  return result;
}

async function updateDraftStatus(res, user, draftId, body) {
  const status = String(body.status || "");
  if (!["draft", "filled", "submitted"].includes(status)) {
    sendJson(res, 400, { data: null, errors: [{ code: "validation_error", message: "Invalid draft status" }] });
    return;
  }

  if (!databaseAvailable && DEV_AUTH_BYPASS) {
    sendJson(res, 200, { data: { ok: true, fallback: true }, errors: null });
    return;
  }

  await pool.query(
    "update auto_bid_application_drafts set status = $1, updated_at = now() where id = $2 and user_id = $3",
    [status, draftId, user.id]
  );
  sendJson(res, 200, { data: { ok: true }, errors: null });
}

async function listJobsFromSheet(res, body) {
  const sheet = readSheetRequest(body);
  const jobs = await listSheetJobs({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    startRow: sheet.startRow,
    endRow: sheet.endRow
  });
  sendJson(res, 200, { data: { jobs }, errors: null });
}

async function saveSheetQuestions(res, body) {
  const sheet = readSheetRequest(body);
  const result = await writeQuestionPayload({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    rowNumber: sheet.rowNumber,
    payload: body.payload || {}
  });
  sendJson(res, 200, { data: result, errors: null });
}

async function listPendingSheetQuestions(res, body) {
  const sheet = readSheetRequest(body);
  const rows = await listPendingQuestionRows({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    startRow: sheet.startRow,
    endRow: sheet.endRow
  });
  sendJson(res, 200, { data: { rows }, errors: null });
}

async function getSheetAnswers(res, body) {
  const sheet = readSheetRequest(body);
  const result = await readAnswerPayload({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    rowNumber: sheet.rowNumber
  });
  sendJson(res, 200, { data: result, errors: null });
}

async function saveSheetAnswers(res, body) {
  const sheet = readSheetRequest(body);
  const result = await writeAnswerPayload({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    rowNumber: sheet.rowNumber,
    answers: body.answers || [],
    payload: body.payload || null
  });
  sendJson(res, 200, { data: result, errors: null });
}

async function getSheetResumeFile(res, body) {
  const sheet = readSheetRequest(body);
  const result = await readResumeFilePayload({
    spreadsheetId: sheet.spreadsheetId,
    sheetName: sheet.sheetName,
    rowNumber: sheet.rowNumber,
    resumeUrl: body.resume_url || body.resumeUrl || "",
    rowValues: body.row_values || body.rowValues || {},
    raw: Array.isArray(body.raw) ? body.raw : [],
    accept: body.accept || []
  });
  sendJson(res, 200, { data: result, errors: null });
}

function readSheetRequest(body = {}) {
  return {
    spreadsheetId: body.spreadsheet_id || body.spreadsheetId,
    sheetName: body.sheet_name || body.sheetName,
    startRow: body.start_row || body.startRow,
    endRow: body.end_row || body.endRow,
    rowNumber: body.row_number || body.rowNumber
  };
}

async function loadOwnedProfile(profileId, userId) {
  if (!databaseAvailable && DEV_AUTH_BYPASS) {
    return profileId === DEV_FALLBACK_PROFILE.id ? DEV_FALLBACK_PROFILE : null;
  }

  const { rows } = await pool.query(
    "select * from auto_bid_profiles where id = $1 and user_id = $2 and active = true limit 1",
    [profileId, userId]
  );
  return rows[0] || null;
}

async function authenticate(req) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : header;
  const userId = readToken(token);
  if (!userId) {
    if (DEV_AUTH_BYPASS) return (await ensureDevAccount()).user;
    return null;
  }

  if (!databaseAvailable && DEV_AUTH_BYPASS && userId === DEV_FALLBACK_USER.id) {
    return DEV_FALLBACK_USER;
  }

  const { rows } = await pool.query(
    "select * from auto_bid_users where id = $1 and active = true limit 1",
    [userId]
  );
  return rows[0] || null;
}
