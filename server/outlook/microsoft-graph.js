import crypto from "node:crypto";
import {
  APP_SECRET,
  MICROSOFT_CLIENT_ID,
  MICROSOFT_CLIENT_SECRET,
  MICROSOFT_OUTLOOK_SCOPES,
  MICROSOFT_TENANT_ID
} from "../config.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const TOKEN_REFRESH_MARGIN_MS = 2 * 60 * 1000;
const MESSAGE_LOOKBACK_MS = 7 * 24 * 60 * 60 * 1000;
const TOKEN_KEY = crypto.createHash("sha256").update(`${APP_SECRET}:outlook-token-v1`).digest();

export function isOutlookConfigured() {
  return Boolean(MICROSOFT_CLIENT_ID && MICROSOFT_CLIENT_SECRET);
}

export function createOutlookAuthorization({ userId, redirectUri }) {
  ensureConfigured();
  const normalizedRedirect = validateRedirectUri(redirectUri);
  const state = signState({
    sub: userId,
    redirect_uri: normalizedRedirect,
    nonce: crypto.randomBytes(18).toString("base64url"),
    exp: Date.now() + OAUTH_STATE_TTL_MS
  });
  const params = new URLSearchParams({
    client_id: MICROSOFT_CLIENT_ID,
    response_type: "code",
    redirect_uri: normalizedRedirect,
    response_mode: "query",
    scope: MICROSOFT_OUTLOOK_SCOPES.join(" "),
    state,
    prompt: "select_account"
  });
  return {
    authorization_url: `${microsoftOAuthBase()}/authorize?${params}`,
    redirect_uri: normalizedRedirect,
    expires_at: new Date(Date.now() + OAUTH_STATE_TTL_MS).toISOString()
  };
}

export async function completeOutlookAuthorization(pool, { userId, code, state, redirectUri }) {
  ensureConfigured();
  if (!code) throw httpError(400, "Microsoft did not return an authorization code");
  const statePayload = verifyState(state);
  const normalizedRedirect = validateRedirectUri(redirectUri);
  if (statePayload.sub !== userId || statePayload.redirect_uri !== normalizedRedirect) {
    throw httpError(403, "Outlook authorization state does not match this Auto Bid session");
  }

  const token = await exchangeToken({
    grant_type: "authorization_code",
    code,
    redirect_uri: normalizedRedirect
  });
  if (!token.refresh_token) throw httpError(502, "Microsoft did not return a refresh token; reconnect and grant offline access");
  const profile = await graphFetchWithToken(token.access_token, "/me?$select=id,displayName,mail,userPrincipalName");
  const expiresAt = new Date(Date.now() + Math.max(60, Number(token.expires_in || 3600)) * 1000);

  const { rows } = await pool.query(
    `insert into auto_bid_outlook_connections
      (user_id, microsoft_user_id, tenant_id, email, display_name, access_token_encrypted,
       refresh_token_encrypted, token_expires_at, scopes, active, updated_at)
     values ($1, $2, $3, $4, $5, $6, $7, $8, $9, true, now())
     on conflict (user_id) do update set
       microsoft_user_id = excluded.microsoft_user_id,
       tenant_id = excluded.tenant_id,
       email = excluded.email,
       display_name = excluded.display_name,
       access_token_encrypted = excluded.access_token_encrypted,
       refresh_token_encrypted = excluded.refresh_token_encrypted,
       token_expires_at = excluded.token_expires_at,
       scopes = excluded.scopes,
       active = true,
       updated_at = now()
     returning *`,
    [
      userId,
      profile.id,
      tokenTenantId(token.id_token),
      profile.mail || profile.userPrincipalName || "",
      profile.displayName || "",
      encryptSecret(token.access_token),
      encryptSecret(token.refresh_token),
      expiresAt,
      normalizeScopes(token.scope || MICROSOFT_OUTLOOK_SCOPES)
    ]
  );
  return serializeConnection(rows[0]);
}

export async function getOutlookConnection(pool, userId) {
  const { rows } = await pool.query(
    "select * from auto_bid_outlook_connections where user_id = $1 and active = true limit 1",
    [userId]
  );
  return rows[0] ? serializeConnection(rows[0]) : { connected: false, configured: isOutlookConfigured() };
}

export async function disconnectOutlook(pool, userId) {
  await pool.query("delete from auto_bid_outlook_connections where user_id = $1", [userId]);
  return { connected: false, configured: isOutlookConfigured() };
}

export async function listVerificationMessages(pool, userId, options = {}) {
  const top = Math.min(40, Math.max(1, Number(options.top || 20)));
  const domain = normalizeHost(options.domain || "");
  const accessToken = await getValidAccessToken(pool, userId);
  const listTop = Math.min(50, Math.max(25, top * 2));
  const query = new URLSearchParams({
    "$top": String(listTop),
    "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,webLink,categories",
    "$orderby": "receivedDateTime desc"
  });
  const headers = { Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"' };
  const folders = await Promise.all(["inbox", "junkemail"].map(async (folder) => {
    try {
      const result = await graphFetch(pool, userId, `/me/mailFolders/${folder}/messages?${query}`, { headers, accessToken });
      return (result.value || []).map((message) => ({ ...message, folder }));
    } catch (error) {
      if (folder === "junkemail") return [];
      throw error;
    }
  }));

  const cutoff = Date.now() - MESSAGE_LOOKBACK_MS;
  const candidates = folders.flat()
    .filter((message) => new Date(message.receivedDateTime || 0).getTime() >= cutoff)
    .map((message) => ({ message, score: scoreVerificationMessage(message, domain) }))
    .filter((item) => item.score >= 5)
    .sort((left, right) => right.score - left.score || new Date(right.message.receivedDateTime) - new Date(left.message.receivedDateTime))
    .slice(0, top);

  const detailed = await Promise.all(candidates.map(async ({ message, score }) => {
    try {
      const detail = await graphFetch(pool, userId, `/me/messages/${encodeURIComponent(message.id)}?$select=id,subject,from,receivedDateTime,isRead,bodyPreview,body,webLink,categories`, {
        headers,
        accessToken
      });
      return serializeMessage({ ...message, ...detail }, score);
    } catch (_error) {
      return serializeMessage(message, score);
    }
  }));

  return { messages: detailed, scanned: folders.flat().length };
}

export async function markOutlookMessageRead(pool, userId, messageId) {
  if (!messageId) throw httpError(400, "Outlook message ID is required");
  await graphFetch(pool, userId, `/me/messages/${encodeURIComponent(messageId)}`, {
    method: "PATCH",
    body: { isRead: true }
  });
  return { id: messageId, is_read: true };
}

export async function findLatestVerificationCode(pool, userId, options = {}) {
  const result = await listVerificationMessages(pool, userId, { ...options, top: Math.min(10, Number(options.top || 10)) });
  const message = result.messages.find((item) => item.codes?.length);
  return {
    code: message?.codes?.[0] || "",
    message: message || null
  };
}

async function getValidAccessToken(pool, userId) {
  ensureConfigured();
  const { rows } = await pool.query(
    "select * from auto_bid_outlook_connections where user_id = $1 and active = true limit 1",
    [userId]
  );
  const connection = rows[0];
  if (!connection) throw httpError(409, "Connect an Outlook mailbox in Auto Bid first");
  if (new Date(connection.token_expires_at).getTime() > Date.now() + TOKEN_REFRESH_MARGIN_MS) {
    return decryptSecret(connection.access_token_encrypted);
  }

  const token = await exchangeToken({
    grant_type: "refresh_token",
    refresh_token: decryptSecret(connection.refresh_token_encrypted),
    scope: MICROSOFT_OUTLOOK_SCOPES.join(" ")
  });
  const refreshToken = token.refresh_token || decryptSecret(connection.refresh_token_encrypted);
  const expiresAt = new Date(Date.now() + Math.max(60, Number(token.expires_in || 3600)) * 1000);
  await pool.query(
    `update auto_bid_outlook_connections
       set access_token_encrypted = $2,
           refresh_token_encrypted = $3,
           token_expires_at = $4,
           scopes = $5,
           updated_at = now()
     where user_id = $1`,
    [userId, encryptSecret(token.access_token), encryptSecret(refreshToken), expiresAt, normalizeScopes(token.scope || connection.scopes)]
  );
  return token.access_token;
}

async function graphFetch(pool, userId, path, options = {}) {
  let accessToken = options.accessToken || await getValidAccessToken(pool, userId);
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await graphFetchWithToken(accessToken, path, options);
    } catch (error) {
      if (error.status !== 401 || attempt > 0) throw error;
      await pool.query("update auto_bid_outlook_connections set token_expires_at = now() where user_id = $1", [userId]);
      accessToken = await getValidAccessToken(pool, userId);
    }
  }
  throw httpError(502, "Microsoft Graph request failed");
}

async function graphFetchWithToken(accessToken, path, options = {}) {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Accept: "application/json",
    ...(options.headers || {})
  };
  if (options.body !== undefined) headers["Content-Type"] = "application/json";
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    method: options.method || "GET",
    headers,
    body: options.body === undefined ? undefined : JSON.stringify(options.body)
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw httpError(response.status, data.error?.message || `Microsoft Graph request failed with ${response.status}`);
  }
  return data;
}

async function exchangeToken(values) {
  const params = new URLSearchParams({
    client_id: MICROSOFT_CLIENT_ID,
    client_secret: MICROSOFT_CLIENT_SECRET,
    ...values
  });
  const response = await fetch(`${microsoftOAuthBase()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || !data.access_token) {
    throw httpError(response.status || 502, data.error_description || data.error || "Microsoft token exchange failed");
  }
  return data;
}

function microsoftOAuthBase() {
  return `https://login.microsoftonline.com/${encodeURIComponent(MICROSOFT_TENANT_ID)}/oauth2/v2.0`;
}

function serializeConnection(connection) {
  return {
    connected: true,
    configured: isOutlookConfigured(),
    email: connection.email || "",
    display_name: connection.display_name || "",
    microsoft_user_id: connection.microsoft_user_id || "",
    scopes: connection.scopes || [],
    token_expires_at: new Date(connection.token_expires_at).toISOString(),
    updated_at: new Date(connection.updated_at).toISOString()
  };
}

function serializeMessage(message, score) {
  const bodyText = stripHtml(message.body?.content || message.bodyPreview || "");
  const combined = `${message.subject || ""}\n${bodyText}`;
  return {
    id: message.id,
    subject: message.subject || "",
    from: {
      name: message.from?.emailAddress?.name || "",
      address: message.from?.emailAddress?.address || ""
    },
    received_at: message.receivedDateTime || null,
    is_read: Boolean(message.isRead),
    folder: message.folder || "",
    preview: String(message.bodyPreview || bodyText).replace(/\s+/g, " ").trim().slice(0, 500),
    codes: extractVerificationCodes(combined),
    links: extractVerificationLinks(message.body?.content || ""),
    outlook_url: message.webLink || "",
    score
  };
}

function scoreVerificationMessage(message, domain) {
  const subject = normalizeText(message.subject);
  const preview = normalizeText(message.bodyPreview);
  const sender = normalizeText(`${message.from?.emailAddress?.name || ""} ${message.from?.emailAddress?.address || ""}`);
  const text = `${subject} ${preview} ${sender}`;
  let score = 0;
  if (/\b(verify|verification|confirm|confirmation|activate|activation)\b/.test(subject)) score += 7;
  if (/\b(code|otp|pin|passcode|one time|one-time|security code)\b/.test(subject)) score += 7;
  if (/\b(verify|verification|confirm|activate|code|otp|pin|passcode)\b/.test(preview)) score += 4;
  if (/\b(application|candidate|career|job|recruit|workday|greenhouse|lever|ashby|workable|successfactors)\b/.test(text)) score += 3;
  if (!message.isRead) score += 1;
  if (domain && text.includes(domain.replace(/\./g, " "))) score += 6;
  const ageHours = Math.max(0, (Date.now() - new Date(message.receivedDateTime || 0).getTime()) / 3_600_000);
  if (ageHours <= 1) score += 4;
  else if (ageHours <= 24) score += 2;
  return score;
}

function extractVerificationCodes(value) {
  const text = String(value || "");
  const codes = [];
  const contextual = /(?:verification|security|confirmation|one[- ]time|otp|pin|passcode|code)(?:\s+(?:is|of))?[^a-z0-9]{0,16}([a-z0-9](?:[a-z0-9-]{2,10}[a-z0-9]))/gi;
  for (const match of text.matchAll(contextual)) {
    const code = match[1].replace(/[^a-z0-9]/gi, "");
    if (/\d/.test(code) && code.length >= 4 && code.length <= 10) codes.push(code);
  }
  for (const match of text.matchAll(/\b\d{6}\b/g)) codes.push(match[0]);
  return Array.from(new Set(codes)).slice(0, 5);
}

function extractVerificationLinks(html) {
  const links = [];
  for (const match of String(html || "").matchAll(/href=["'](https?:\/\/[^"']+)["']/gi)) {
    const url = decodeHtmlEntities(match[1]);
    if (!/(verify|confirm|activate|account|application|candidate|token|code)/i.test(url)) continue;
    if (/(unsubscribe|preferences|privacy|tracking)/i.test(url)) continue;
    links.push(url);
  }
  return Array.from(new Set(links)).slice(0, 4);
}

function stripHtml(value) {
  return decodeHtmlEntities(String(value || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " "))
    .replace(/\s+/g, " ")
    .trim();
}

function decodeHtmlEntities(value) {
  return String(value || "")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
}

function encryptSecret(value) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", TOKEN_KEY, iv);
  const encrypted = Buffer.concat([cipher.update(String(value || ""), "utf8"), cipher.final()]);
  return ["v1", iv.toString("base64url"), cipher.getAuthTag().toString("base64url"), encrypted.toString("base64url")].join(".");
}

function decryptSecret(value) {
  const [version, ivValue, tagValue, encryptedValue] = String(value || "").split(".");
  if (version !== "v1" || !ivValue || !tagValue || !encryptedValue) throw httpError(500, "Stored Outlook token is invalid");
  const decipher = crypto.createDecipheriv("aes-256-gcm", TOKEN_KEY, Buffer.from(ivValue, "base64url"));
  decipher.setAuthTag(Buffer.from(tagValue, "base64url"));
  return Buffer.concat([decipher.update(Buffer.from(encryptedValue, "base64url")), decipher.final()]).toString("utf8");
}

function signState(payload) {
  const encoded = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const signature = crypto.createHmac("sha256", APP_SECRET).update(`outlook.${encoded}`).digest("base64url");
  return `${encoded}.${signature}`;
}

function verifyState(value) {
  const [encoded, providedSignature] = String(value || "").split(".");
  if (!encoded || !providedSignature) throw httpError(400, "Invalid Outlook authorization state");
  const expectedSignature = crypto.createHmac("sha256", APP_SECRET).update(`outlook.${encoded}`).digest("base64url");
  const provided = Buffer.from(providedSignature);
  const expected = Buffer.from(expectedSignature);
  if (provided.length !== expected.length || !crypto.timingSafeEqual(provided, expected)) {
    throw httpError(403, "Outlook authorization state signature is invalid");
  }
  let payload;
  try {
    payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
  } catch (_error) {
    throw httpError(400, "Outlook authorization state is malformed");
  }
  if (!payload.exp || Date.now() > Number(payload.exp)) throw httpError(403, "Outlook authorization request expired");
  return payload;
}

function validateRedirectUri(value) {
  let url;
  try {
    url = new URL(String(value || ""));
  } catch (_error) {
    throw httpError(400, "Invalid Outlook OAuth redirect URL");
  }
  if (url.protocol !== "https:" || !url.hostname.endsWith(".chromiumapp.org") || !url.pathname.startsWith("/outlook")) {
    throw httpError(400, "Outlook OAuth redirect must be the Chrome extension identity redirect");
  }
  return url.toString();
}

function tokenTenantId(idToken) {
  try {
    const payload = JSON.parse(Buffer.from(String(idToken || "").split(".")[1], "base64url").toString("utf8"));
    return payload.tid || "";
  } catch (_error) {
    return "";
  }
}

function normalizeScopes(value) {
  if (Array.isArray(value)) return value.map(String).filter(Boolean);
  return String(value || "").split(/[\s,]+/).filter(Boolean);
}

function normalizeText(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9.]+/g, " ").trim();
}

function normalizeHost(value) {
  return String(value || "").toLowerCase().replace(/^https?:\/\//, "").replace(/^www\./, "").split(/[/:]/)[0];
}

function ensureConfigured() {
  if (!isOutlookConfigured()) throw httpError(503, "Set MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET on the Auto Bid server");
}

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}
