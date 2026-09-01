import {
  CACHE_SCOPES,
  OPENAI_API_KEY,
  OPENAI_MODEL,
  OPENAI_ROUTE_ENABLED
} from "../config.js";
import { logBackendEvent } from "../utils/logger.js";
import { extractOutputText, trimForPrompt } from "../utils/text.js";

const SYSTEM_PROMPT = [
  "Generate honest job application form answers from the supplied profile and page context.",
  "Only answer the supplied fields; they were pre-filtered to exclude basic profile fields.",
  "Do not invent employers, degrees, certifications, locations, compensation, legal eligibility, or years of experience.",
  "Use concise first-person answers unless the field asks for a different format.",
  "For select, radio, checkbox, combobox, and button-group fields with provided options, return exactly one provided option.",
  "Never return placeholders such as Not specified, Unknown, Not provided, or TBD.",
  "Use N/A or Not applicable only when the field instructions explicitly permit it or the question is genuinely conditional and inapplicable.",
  "Use cache_scope global only for profile-independent consent or terms answers, profile for reusable profile facts, and profile_job for job-specific answers like fit, cover letters, proposals, and bids.",
  "Answer every supplied field when an honest answer can be supported by the supplied context."
].join(" ");

const OPENAI_TIMEOUT_MS = 90_000;

export async function generateAiAnswers(fields, profile, page, jobHash, warnings, logContext = {}) {
  if (!OPENAI_ROUTE_ENABLED) {
    warnings.push("OpenAI autofill routing is disabled.");
    logBackendEvent("OPENAI_SKIPPED", {
      job_id: logContext.jobId || null,
      reason: "route-disabled",
      fields: fields.length
    }, { requestId: logContext.requestId });
    return [];
  }

  const payload = buildAiPayload(fields, profile, page, jobHash);
  const model = OPENAI_MODEL;

  if (!OPENAI_API_KEY || !model) {
    warnings.push("OpenAI is not configured. Set OPENAI_API_KEY and OPENAI_MODEL.");
    logBackendEvent("OPENAI_SKIPPED", {
      job_id: logContext.jobId || null,
      reason: "not-configured",
      fields: fields.length
    }, { requestId: logContext.requestId, level: "warn" });
    return [];
  }

  const estimates = estimateTokenUsage(payload);
  const route = {
    provider: "openai",
    model,
    input_tokens_estimate: estimates.input_tokens,
    output_tokens_estimate: estimates.output_tokens,
    estimated_request_cost_usd: null
  };
  const startedAt = Date.now();
  logBackendEvent("OPENAI_REQUEST", {
    job_id: logContext.jobId || null,
    model,
    input_tokens_estimate: route.input_tokens_estimate,
    output_tokens_estimate: route.output_tokens_estimate,
    payload
  }, { requestId: logContext.requestId });

  try {
    const { parsed, usage } = await generateOpenAiAnswers(payload, model);
    const answers = normalizeAiResponse(parsed, fields, route, warnings);
    if (answers.length === 0) warnings.push(`OpenAI ${model} returned no usable answers.`);
    logBackendEvent("OPENAI_RESPONSE", {
      job_id: logContext.jobId || null,
      model,
      status: "success",
      duration_ms: Date.now() - startedAt,
      answers,
      provider_warnings: parsed?.warnings || [],
      usage
    }, { requestId: logContext.requestId });
    return answers;
  } catch (error) {
    warnings.push(`OpenAI ${model} request failed: ${error.message || String(error)}`);
    logBackendEvent("OPENAI_RESPONSE", {
      job_id: logContext.jobId || null,
      model,
      status: "error",
      duration_ms: Date.now() - startedAt,
      error
    }, { requestId: logContext.requestId, level: "error" });
    return [];
  }
}

function buildAiPayload(fields, profile, page, jobHash) {
  return {
    profile: {
      name: profile.name,
      static_fields: profile.static_fields,
      resume_text: trimForPrompt(profile.resume_text, 7000),
      preferences: profile.preferences,
      profile_version: profile.profile_version
    },
    page: {
      url: page.url,
      title: page.title,
      job_title: page.job_title,
      text: trimForPrompt(page.text, 9000)
    },
    job_hash: jobHash,
    fields: fields.map((field) => ({
      field_id: field.id,
      label: field.label,
      type: field.type,
      required: field.required,
      options: field.options,
      default_cache_scope: field.cache_scope
    }))
  };
}

async function generateOpenAiAnswers(payload, model) {
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    signal: AbortSignal.timeout(OPENAI_TIMEOUT_MS),
    body: JSON.stringify({
      model,
      store: false,
      instructions: SYSTEM_PROMPT,
      input: JSON.stringify(payload),
      text: {
        format: {
          type: "json_schema",
          name: "auto_bid_answers",
          strict: true,
          schema: openAiAnswerSchema()
        }
      },
      max_output_tokens: Math.max(1500, payload.fields.length * 500)
    })
  });

  if (!response.ok) {
    throw new Error(`status ${response.status}: ${await limitedResponseText(response)}`);
  }

  const data = await response.json();
  const outputText = extractOutputText(data);
  if (!outputText) throw new Error("no parseable output");
  return {
    parsed: parseJsonOutput(outputText),
    usage: data.usage || null
  };
}

function normalizeAiResponse(parsed, fields, route, warnings) {
  if (Array.isArray(parsed?.warnings)) warnings.push(...parsed.warnings.map(String));
  const fieldIds = new Set(fields.map((field) => field.id));
  return Array.isArray(parsed?.answers)
    ? parsed.answers
      .filter((answer) => fieldIds.has(answer.field_id) && typeof answer.value === "string" && answer.value.trim())
      .filter((answer) => !isRejectedAiPlaceholder(answer.value))
      .map((answer) => ({
        field_id: String(answer.field_id),
        value: String(answer.value),
        cache_scope: CACHE_SCOPES.includes(answer.cache_scope) ? answer.cache_scope : "profile_job",
        confidence: typeof answer.confidence === "number" ? answer.confidence : null,
        warning: answer.warning ? String(answer.warning) : null,
        provider: route.provider,
        model: route.model,
        estimated_request_cost_usd: route.estimated_request_cost_usd
      }))
    : [];
}

function isRejectedAiPlaceholder(value) {
  return /^(?:not specified|unspecified|unknown|not provided|not available|no information(?: provided| available)?|information unavailable|to be determined|tbd)$/i
    .test(String(value || "").trim());
}

function openAiAnswerSchema() {
  return {
    type: "object",
    additionalProperties: false,
    properties: {
      answers: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            field_id: { type: "string" },
            value: { type: "string" },
            cache_scope: { type: "string", enum: CACHE_SCOPES },
            confidence: { type: "number" },
            warning: { type: ["string", "null"] }
          },
          required: ["field_id", "value", "cache_scope", "confidence", "warning"]
        }
      },
      warnings: {
        type: "array",
        items: { type: "string" }
      }
    },
    required: ["answers", "warnings"]
  };
}

function estimateTokenUsage(payload) {
  const inputChars = SYSTEM_PROMPT.length + JSON.stringify(payload).length;
  const fieldCount = Array.isArray(payload.fields) ? payload.fields.length : 0;
  return {
    input_tokens: Math.ceil(inputChars / 4),
    output_tokens: Math.max(180, fieldCount * 90)
  };
}

function parseJsonOutput(text) {
  const trimmed = String(text || "").trim();
  try {
    return JSON.parse(trimmed);
  } catch {
    const match = trimmed.match(/\{[\s\S]*\}/);
    if (!match) throw new Error("output was not valid JSON");
    return JSON.parse(match[0]);
  }
}

async function limitedResponseText(response) {
  const text = await response.text().catch(() => "");
  return text.slice(0, 500) || "empty response";
}
