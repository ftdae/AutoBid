import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const popupSource = await readFile(new URL("../extension/popup.js", import.meta.url), "utf8");
const panelSource = await readFile(new URL("../extension/panel-host.js", import.meta.url), "utf8");

test("an extension reload refreshes a stale AutoBid panel instead of exposing the raw runtime error", () => {
  assert.match(popupSource, /extension context invalidated/i);
  assert.match(popupSource, /scheduleContextRecovery/);
  assert.match(popupSource, /location\.reload\(\)/);
  assert.match(popupSource, /CONTEXT_RECOVERY_COOLDOWN_MS = 10000/);
});

test("the panel host can reload its embedded extension page as a fallback", () => {
  assert.match(panelSource, /type === "RELOAD_PANEL"/);
  assert.match(panelSource, /iframe\.src = url\.href/);
});
