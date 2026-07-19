import { CACHE_SCOPES, OPENAI_API_KEY, OPENAI_MODEL } from "../config.js";
import { extractOutputText, trimForPrompt } from "../utils/text.js";

export async function generateAiAnswers(fields, profile, page, jobHash, warnings) {
  if (!OPENAI_API_KEY) {
    warnings.push("OPENAI_API_KEY or GPT_API_KEY is not configured, so dynamic answers were skipped.");
    return [];
  }

  const payload = {
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

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      input: [
        {
          role: "system",
          content: [{
            type: "input_text",
            text: [
              "Generate honest job application form answers from the supplied profile and page context.",
              "Do not invent employers, degrees, certifications, locations, compensation, legal eligibility, or years of experience.",
              "Use concise first-person answers unless the field asks for a different format.",
              "For select, radio, checkbox, and button-group fields, choose one of the provided options when possible.",
              "Use cache_scope global only for profile-independent consent or terms answers, profile for reusable profile facts, and profile_job for job-specific answers like fit, cover letters, proposals, and bids."
            ].join(" ")
          }]
        },
        {
          role: "user",
          content: [{ type: "input_text", text: JSON.stringify(payload) }]
        }
      ],
      text: {
        format: {
          type: "json_schema",
          name: "auto_bid_answers",
          strict: true,
          schema: {
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
              warnings: { type: "array", items: { type: "string" } }
            },
            required: ["answers", "warnings"]
          }
        }
      }
    })
  });

  if (!response.ok) {
    warnings.push(`OpenAI request failed with status ${response.status}.`);
    return [];
  }

  const data = await response.json();
  const outputText = extractOutputText(data);
  if (!outputText) {
    warnings.push("OpenAI returned no parseable output.");
    return [];
  }

  try {
    const parsed = JSON.parse(outputText);
    if (Array.isArray(parsed.warnings)) warnings.push(...parsed.warnings.map(String));
    const fieldIds = new Set(fields.map((field) => field.id));
    return Array.isArray(parsed.answers)
      ? parsed.answers
        .filter((answer) => fieldIds.has(answer.field_id) && typeof answer.value === "string")
        .map((answer) => ({
          field_id: String(answer.field_id),
          value: String(answer.value),
          cache_scope: CACHE_SCOPES.includes(answer.cache_scope) ? answer.cache_scope : "profile_job",
          confidence: typeof answer.confidence === "number" ? answer.confidence : null,
          warning: answer.warning ? String(answer.warning) : null
        }))
      : [];
  } catch {
    warnings.push("OpenAI output was not valid JSON.");
    return [];
  }
}
