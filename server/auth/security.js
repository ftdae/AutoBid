import { createHmac, randomBytes, scryptSync, timingSafeEqual } from "node:crypto";
import { APP_SECRET, TOKEN_MAX_AGE_MS, TOKEN_VERSION } from "../config.js";

export function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

export function createToken(userId) {
  const issuedAt = Math.floor(Date.now() / 1000);
  const payload = {
    sub: userId,
    iat: issuedAt,
    exp: Math.floor((Date.now() + TOKEN_MAX_AGE_MS) / 1000)
  };
  const encodedPayload = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const signature = createHmac("sha256", APP_SECRET).update(`${TOKEN_VERSION}.${encodedPayload}`).digest("base64url");
  return `${TOKEN_VERSION}.${encodedPayload}.${signature}`;
}

export function readToken(token) {
  if (!token) return null;
  try {
    const [version, encodedPayload, signature] = token.split(".");
    if (version !== TOKEN_VERSION || !encodedPayload || !signature) return null;
    const expected = createHmac("sha256", APP_SECRET).update(`${TOKEN_VERSION}.${encodedPayload}`).digest("base64url");
    const actualBuffer = Buffer.from(signature);
    const expectedBuffer = Buffer.from(expected);
    if (actualBuffer.length !== expectedBuffer.length || !timingSafeEqual(actualBuffer, expectedBuffer)) return null;
    const payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString());
    if (payload.exp <= Math.floor(Date.now() / 1000)) return null;
    return typeof payload.sub === "string" ? payload.sub : null;
  } catch {
    return null;
  }
}

export function hashPassword(password) {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

export function verifyPassword(password, stored) {
  const [salt, hash] = String(stored || "").split(":");
  if (!salt || !hash) return false;
  return scryptSync(password, salt, 64).toString("hex") === hash;
}
