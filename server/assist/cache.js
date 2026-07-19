import { id } from "../utils/id.js";
import { normalizeText } from "../utils/text.js";

export async function ensureQuestions(pool, fields, domain, normalizedUrl) {
  for (const field of fields) {
    const { rows } = await pool.query(
      "select question_hash, cache_scope from auto_bid_questions where question_hash = $1 limit 1",
      [field.question_hash]
    );
    const existing = rows[0];
    if (existing) {
      field.cache_scope = existing.cache_scope;
      continue;
    }

    await pool.query(
      `insert into auto_bid_questions
        (id, question_hash, domain, url_pattern, normalized_label, field_type, options_json, required, cache_scope)
       values ($1, $2, $3, $4, $5, $6, $7::jsonb, $8, $9)
       on conflict (question_hash) do nothing`,
      [
        id("abq"),
        field.question_hash,
        domain,
        normalizedUrl,
        normalizeText(field.label),
        field.type,
        JSON.stringify(field.options),
        field.required,
        field.cache_scope
      ]
    );
  }
}

export async function loadCacheAnswers(pool, fields, profile, jobHash, staticAnswers) {
  const answers = new Map();
  let hits = 0;

  for (const field of fields) {
    if (staticAnswers.has(field.id)) continue;

    const params = [field.question_hash, field.cache_scope];
    let scopeWhere = "profile_id is null";

    if (field.cache_scope === "profile") {
      params.push(profile.id, profile.profile_version);
      scopeWhere = `profile_id = $${params.length - 1} and profile_version = $${params.length}`;
    } else if (field.cache_scope === "profile_job") {
      params.push(profile.id, profile.profile_version, jobHash);
      scopeWhere = `profile_id = $${params.length - 2} and profile_version = $${params.length - 1} and job_hash = $${params.length}`;
    }

    const { rows } = await pool.query(
      `select *
         from auto_bid_answer_cache
        where question_hash = $1
          and cache_scope = $2
          and (${scopeWhere})
          and (expires_at is null or expires_at > now())
        order by created_at desc
        limit 1`,
      params
    );
    const cached = rows[0];

    if (!cached) continue;
    hits += 1;
    answers.set(field.id, {
      field_id: field.id,
      value: cached.answer,
      source: "cache",
      cache_scope: cached.cache_scope,
      confidence: cached.confidence === null || cached.confidence === undefined ? null : Number(cached.confidence),
      warning: null
    });
  }

  return { answers, hits };
}

export async function saveAiAnswers(pool, answers, fields, profile, jobHash) {
  const fieldsById = new Map(fields.map((field) => [field.id, field]));

  for (const answer of answers) {
    const field = fieldsById.get(answer.field_id);
    if (!field || !answer.value.trim()) continue;

    await pool.query(
      "update auto_bid_questions set cache_scope = $1, updated_at = now() where question_hash = $2",
      [answer.cache_scope, field.question_hash]
    );

    await pool.query(
      `insert into auto_bid_answer_cache
        (id, question_hash, cache_scope, profile_id, profile_version, job_hash, answer, confidence, source)
       values ($1, $2, $3, $4, $5, $6, $7, $8, 'ai')`,
      [
        id("aba"),
        field.question_hash,
        answer.cache_scope,
        answer.cache_scope === "global" ? null : profile.id,
        answer.cache_scope === "global" ? null : profile.profile_version,
        answer.cache_scope === "profile_job" ? jobHash : null,
        answer.value,
        answer.confidence
      ]
    );
  }
}
