import { normalizeText } from "../utils/text.js";

export function buildStaticAnswers(fields, staticFields) {
  const answers = new Map();

  for (const field of fields) {
    const key = matchStaticFieldKey(field);
    if (!key) continue;
    const value = getStaticFieldValue(staticFields, key);
    if (value === undefined || value === null || String(value).trim() === "") continue;

    answers.set(field.id, {
      field_id: field.id,
      value: String(value),
      source: "static",
      cache_scope: "profile",
      confidence: 1,
      warning: null
    });
  }

  return answers;
}

export function matchStaticFieldKey(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  if (isLanguageYesNoField(field)) return null;
  if (isPlainFullNameField(field)) return "full_name";
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name"]],
    ["email", ["email", "e mail", "mail"]],
    ["phone", ["phone", "mobile", "telephone", "cell"]],
    ["location", ["location", "address", "current city", "current location"]],
    ["city", ["city"]],
    ["country", ["country", "residence", "current residence", "where is your current residence", "where are you based"]],
    ["linkedin", ["linkedin"]],
    ["github", ["github"]],
    ["portfolio", ["portfolio"]],
    ["website", ["website", "personal site", "web site"]],
    ["languages", ["languages", "spoken languages", "language proficiency", "fluent languages", "languages spoken"]],
    ["expected_rate", ["hourly rate", "rate", "expected rate", "expected salary", "salary expectation", "salary expectations", "expected compensation", "desired salary", "desired compensation", "gross monthly", "monthly salary", "salary", "compensation"]],
    ["work_authorization", ["authorized", "authorization", "legally work", "eligible to work"]],
    ["sponsorship", ["sponsor", "sponsorship", "visa"]],
    ["availability", ["availability", "available", "start date"]],
    ["notice_period", ["notice period", "current notice", "notice"]]
  ];

  for (const [key, needles] of patterns) {
    if (needles.some((needle) => text.includes(needle))) return key;
  }

  return null;
}

export function getStaticFieldValue(staticFields, key) {
  const aliases = {
    expected_rate: ["expected_rate", "expected_salary", "salary_expectation", "salary_expectations", "monthly_salary", "monthly_salary_expectation", "desired_salary", "desired_compensation"],
    notice_period: ["notice_period", "current_notice_period", "availability_notice"],
    languages: ["languages", "language", "spoken_languages", "language_proficiency", "fluent_languages", "languages_spoken"]
  };

  if (key === "full_name") {
    const fullName = staticFields?.full_name;
    if (fullName !== undefined && fullName !== null && String(fullName).trim()) return fullName;
    const composed = [staticFields?.first_name, staticFields?.last_name].filter(Boolean).join(" ").trim();
    if (composed) return composed;
  }

  for (const candidate of [key, ...(aliases[key] || [])]) {
    const value = staticFields?.[candidate];
    if (value !== undefined && value !== null && String(value).trim()) return value;
  }

  return staticFields?.[key];
}

function isPlainFullNameField(field) {
  if (!isTextLikeStaticField(field)) return false;
  const candidates = [field.label, field.name, field.autocomplete]
    .map(normalizeText)
    .filter(Boolean);
  return candidates.some((candidate) => ["name", "your name", "applicant name", "candidate name", "full name"].includes(candidate));
}

function isTextLikeStaticField(field) {
  return !["checkbox", "radio", "select", "combobox", "button-group", "file", "hidden", "password", "submit", "button", "reset"].includes(field.type);
}

function isLanguageYesNoField(field) {
  const text = normalizeText([field.autocomplete, field.name, field.label, field.placeholder].join(" "));
  const options = (field.options || []).map(normalizeText);
  return /(speak|language|fluent|fluency|proficien|native speaker|bilingual|multilingual)/.test(text) &&
    options.includes("yes") &&
    options.includes("no");
}
