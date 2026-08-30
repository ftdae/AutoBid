import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const contentSource = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");

test("resume upload waits for a new server-confirmed component attachment", () => {
  assert.match(contentSource, /const RESUME_SERVER_CONFIRM_TIMEOUT_MS = 6000;/);
  assert.match(contentSource, /hasNewServerConfirmedResumeUpload/);
  assert.match(contentSource, /getResumeServerConfirmationSnapshot/);
  assert.match(contentSource, /fileid\|file id\|attachmenttype\|attachment type\|uploaded/);
});

test("resume metadata lookup crosses open shadow-root component hosts", () => {
  assert.match(contentSource, /function getResumeUploadComponentElements/);
  assert.match(contentSource, /current = root\?\.host \|\| null/);
  assert.match(contentSource, /spl-dropzone/);
});

test("native file uploads are serialized and reconnect once after debugger detachment", () => {
  assert.match(backgroundSource, /queueNativeInput\(tabId, \(\) => dispatchNativeFileUpload/);
  assert.match(backgroundSource, /dispatchNativeFileUpload\(tabId, payload, reconnectAttempt \+ 1\)/);
  assert.match(backgroundSource, /dispatchNativeFileChooserUpload\(tabId, payload, reconnectAttempt \+ 1\)/);
});
