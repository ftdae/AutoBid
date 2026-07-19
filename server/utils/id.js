import { randomBytes } from "node:crypto";

export function id(prefix) {
  return `${prefix}_${randomBytes(12).toString("hex")}`;
}
