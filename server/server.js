import http from "node:http";
import { createToken, hashPassword, normalizeEmail, readToken, verifyPassword } from "./auth/security.js";
import { ensureQuestions, loadCacheAnswers, saveAiAnswers } from "./assist/cache.js";
import { normalizeFields, normalizePage, shouldAnswerWithAi } from "./assist/field-policy.js";
import { generateAiAnswers } from "./assist/openai.js";
import { DEV_AUTH_BYPASS, DEV_USER_EMAIL, PORT } from "./config.js";
import { pool } from "./db/pool.js";
import { ensureSchema } from "./db/schema.js";
import { readJson, sendJson, setCorsHeaders } from "./http/json.js";
import { buildStaticAnswers } from "./profiles/static-fields.js";
import { id } from "./utils/id.js";
import { hashText, normalizeUrl, safeDomain, trimForPrompt } from "./utils/text.js";
import { serializeProfile, serializeUser } from "./users/serializers.js";

await ensureSchema(pool);

const server = http.createServer(async (req, res) => {
  try {
    await handleRequest(req, res);
  } catch (error) {
    console.error(error);
    sendJson(res, 500, {
      data: null,
      errors: [{ code: "server_error", message: error.message || "Internal server error" }]
    });
  }
});

server.listen(PORT, () => {
  console.log(`AutoBid server listening on http://localhost:${PORT}`);
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

  if (req.method === "POST" && pathname === "/api/auto-bid/assist") {
    return assist(res, user, body);
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
  await pool.query(
    "update auto_bid_profiles set active = false, updated_at = now() where id = $1 and user_id = $2",
    [profileId, user.id]
  );
  sendJson(res, 200, { data: { ok: true }, errors: null });
}

async function assist(res, user, body) {
  const profile = await loadOwnedProfile(body.profile_id, user.id);
  if (!profile) {
    sendJson(res, 404, { data: null, errors: [{ code: "profile_not_found", message: "Profile not found" }] });
    return;
  }

  const page = normalizePage(body.page || {});
  const normalizedUrl = normalizeUrl(page.url);
  const domain = page.domain || safeDomain(page.url);
  const jobHash = hashText([domain, normalizedUrl, page.title, page.job_title, trimForPrompt(page.text, 6000)].join("\n"));
  const fields = normalizeFields(Array.isArray(body.fields) ? body.fields : [], domain, normalizedUrl);

  await ensureQuestions(pool, fields, domain, normalizedUrl);

  const staticAnswers = buildStaticAnswers(fields, profile.static_fields || {});
  const cache = await loadCacheAnswers(pool, fields, profile, jobHash, staticAnswers);
  const fieldsForAi = fields.filter((field) => shouldAnswerWithAi(field) && !staticAnswers.has(field.id) && !cache.answers.has(field.id));
  const warnings = [];
  const aiAnswers = fieldsForAi.length ? await generateAiAnswers(fieldsForAi, profile, page, jobHash, warnings) : [];

  await saveAiAnswers(pool, aiAnswers, fields, profile, jobHash);

  const answers = [
    ...Array.from(staticAnswers.values()),
    ...Array.from(cache.answers.values()),
    ...aiAnswers.map((answer) => ({
      field_id: answer.field_id,
      value: answer.value,
      source: "ai",
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

  sendJson(res, 200, {
    data: {
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
    },
    errors: null
  });
}

async function updateDraftStatus(res, user, draftId, body) {
  const status = String(body.status || "");
  if (!["draft", "filled", "submitted"].includes(status)) {
    sendJson(res, 400, { data: null, errors: [{ code: "validation_error", message: "Invalid draft status" }] });
    return;
  }

  await pool.query(
    "update auto_bid_application_drafts set status = $1, updated_at = now() where id = $2 and user_id = $3",
    [status, draftId, user.id]
  );
  sendJson(res, 200, { data: { ok: true }, errors: null });
}

async function loadOwnedProfile(profileId, userId) {
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
  const { rows } = await pool.query(
    "select * from auto_bid_users where id = $1 and active = true limit 1",
    [userId]
  );
  return rows[0] || null;
}
