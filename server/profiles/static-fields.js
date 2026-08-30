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
  const optionText = normalizeText((field.options || []).join(" "));
  if (isLanguageYesNoField(field)) return null;
  if (isPlainFullNameField(field)) return "full_name";
  const addressComponentKey = matchAddressComponentFieldKey(field);
  if (addressComponentKey) return addressComponentKey;
  if (/(work status|right to work|employment status)/.test(text) &&
      /(national|citizen|work permit|residence permit|third country)/.test(optionText)) {
    return "work_authorization";
  }
  const patterns = [
    ["first_name", ["given name", "first name", "firstname", "first_name"]],
    ["last_name", ["family name", "last name", "lastname", "surname", "last_name"]],
    ["full_name", ["full name", "your name", "applicant name"]],
    ["email", ["email", "e mail", "mail"]],
    ["phone", ["phone", "mobile", "telephone", "cell"]],
    ["location", ["location", "address", "current city", "current location"]],
    ["city", ["city"]],
    ["nationality", ["nationality", "citizenship", "country of citizenship", "citizen of"]],
    ["country", ["country", "residence", "current residence", "where is your current residence", "where are you based"]],
    ["linkedin", ["linkedin"]],
    ["github", ["github"]],
    ["portfolio", ["portfolio"]],
    ["website", ["website", "personal site", "web site"]],
    ["languages", ["languages", "spoken languages", "language proficiency", "fluent languages", "languages spoken"]],
    ["expected_rate", ["hourly rate", "rate", "expected rate", "expected salary", "salary expectation", "salary expectations", "expected compensation", "desired salary", "desired compensation", "gross monthly", "monthly salary", "salary", "compensation"]],
    ["work_authorization", ["authorized", "authorization", "legally work", "eligible to work", "right to work", "work status", "employment status"]],
    ["sponsorship", ["sponsor", "sponsorship", "visa"]],
    ["availability", ["availability", "available", "start date"]],
    ["notice_period", ["notice period", "current notice", "notice"]]
  ];

  for (const [key, needles] of patterns) {
    if (needles.some((needle) => includesNormalizedPhrase(text, needle))) return key;
  }

  return null;
}

function includesNormalizedPhrase(text, phrase) {
  const normalizedPhrase = normalizeText(phrase);
  return Boolean(normalizedPhrase) && ` ${text} `.includes(` ${normalizedPhrase} `);
}

export function getStaticFieldValue(staticFields, key) {
  const aliases = {
    postal_code: ["postal_code", "postalcode", "post_code", "postcode", "zip_code", "zipcode", "zip"],
    state_region: ["state_region", "state_province_region", "state_province", "state", "province", "region", "administrative_area"],
    expected_rate: ["expected_rate", "expected_salary", "salary_expectation", "salary_expectations", "monthly_salary", "monthly_salary_expectation", "desired_salary", "desired_compensation"],
    notice_period: ["notice_period", "current_notice_period", "availability_notice"],
    languages: ["languages", "language", "spoken_languages", "language_proficiency", "fluent_languages", "languages_spoken"],
    work_authorization: ["work_authorization", "right_to_work", "work_status", "employment_status", "nationality", "citizenship"],
    nationality: ["nationality", "citizenship", "citizen_of", "country_of_citizenship"]
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

function matchAddressComponentFieldKey(field) {
  const autocomplete = normalizeText(field.autocomplete || "");
  const name = normalizeText(field.name || "");
  const prompt = simplifyAddressPrompt(field.label || field.placeholder || "");

  if (autocomplete === "postal code") return "postal_code";
  if (autocomplete === "address level1") return "state_region";

  const postalAliases = new Set([
    "postal code",
    "postalcode",
    "post code",
    "postcode",
    "zip code",
    "zipcode",
    "zip",
    "pin code",
    "postal code zip code",
    "zip code postal code",
    "postal zip code",
    "zip postal code"
  ]);
  if (postalAliases.has(prompt) || postalAliases.has(name)) return "postal_code";

  const regionAliases = new Set([
    "state",
    "province",
    "region",
    "county",
    "prefecture",
    "administrative area",
    "state province",
    "state region",
    "province region",
    "state province region"
  ]);
  if (regionAliases.has(prompt) || regionAliases.has(name)) return "state_region";

  return null;
}

function simplifyAddressPrompt(value) {
  return normalizeText(value)
    .replace(/^(?:what is|please enter|please select|enter|select|choose|provide)\s+(?:your\s+)?(?:current\s+)?/, "")
    .replace(/\s+(?:required|optional)$/, "")
    .trim();
}
