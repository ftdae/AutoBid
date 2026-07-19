import {
  hashText,
  normalizeDomain,
  normalizeFieldType,
  normalizeText,
  safeDomain
} from "../utils/text.js";

export function normalizePage(page) {
  return {
    url: String(page.url || ""),
    domain: normalizeDomain(page.domain || safeDomain(page.url)),
    title: String(page.title || ""),
    job_title: String(page.job_title || ""),
    text: String(page.text || "")
  };
}

export function normalizeFields(fields, domain, normalizedUrl) {
  return fields.map((field, index) => {
    const label = String(field.label || field.placeholder || field.name || field.id || "");
    const type = normalizeFieldType(field.type);
    const options = Array.isArray(field.options) ? field.options.map((option) => String(option).trim()).filter(Boolean) : [];
    const normalizedLabel = normalizeText([label, field.name || "", field.placeholder || "", field.autocomplete || ""].filter(Boolean).join(" "));
    const questionHash = hashText([domain, normalizedUrl, normalizedLabel, type, options.map(normalizeText).join("|")].join("|"));

    return {
      id: String(field.id || `field_${index}`),
      label,
      name: String(field.name || ""),
      placeholder: String(field.placeholder || ""),
      autocomplete: String(field.autocomplete || ""),
      type,
      required: Boolean(field.required),
      options,
      value: String(field.value || ""),
      question_hash: questionHash,
      cache_scope: inferCacheScope(label, type)
    };
  });
}

export function shouldAnswerWithAi(field) {
  if (["file", "hidden", "password", "submit", "button", "reset"].includes(field.type)) return false;
  if (field.value.trim()) return false;
  if (isSensitiveOptionalField(field)) return false;
  if (field.required || field.type === "textarea") return true;
  if (["select", "radio", "checkbox", "combobox", "button-group"].includes(field.type)) {
    return field.options.length > 0 || isLikelyApplicationQuestion(field);
  }
  return isLikelyApplicationQuestion(field);
}

export function inferCacheScope(label, type) {
  const text = normalizeText(label);
  if (/(terms|privacy|consent|agree|acknowledge|confirm|certify|accurate|contact me|email me)/.test(text)) return "global";
  if (/(cover letter|proposal|bid|why|fit|motivation|interested|describe|tell us|experience with|project|role)/.test(text)) return "profile_job";
  if (type === "textarea") return "profile_job";
  return "profile";
}

function isLikelyApplicationQuestion(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  if (!text) return false;
  if (/(search|filter|captcha|recaptcha|password|one time|otp|verification code|coupon|promo|upload|attach|resume|cv)/.test(text)) return false;
  return /(linkedin|github|portfolio|website|personal site|profile url|headline|summary|bio|phone|mobile|city|country|location|address|residence|where are you based|salary|compensation|rate|notice|availability|available|start date|authorization|authorized|eligible|legally work|work permit|sponsor|sponsorship|visa|relocat|remote|experience|years|level|skill|technology|framework|language|speak|fluent|fluency|bilingual|multilingual|aws|api|java|spring|kotlin|sql|database|why|motivation|interested|describe|tell us|question|cover letter|proposal|bid|source|heard about|referral|terms|privacy|policy|consent|agree|accept|acknowledge|confirm|certify|accurate|contact me|future job opportunit|future opportunit|talent community|job alert|recruiting communication|recruitment communication)/.test(text);
}

function isSensitiveOptionalField(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  if (!text) return false;
  return /(gender|race|ethnicity|ethnic|disability|veteran|protected veteran|sexual orientation|pronoun|he him|she her|they them|xe xem|ze hir|ey em|hir hir|fae faer|hu hu|use name only|custom|date of birth|birth date|national id|social security|ssn|passport)/.test(text);
}
