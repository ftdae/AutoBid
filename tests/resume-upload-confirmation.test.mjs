import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const contentSource = await readFile(new URL("../extension/content.js", import.meta.url), "utf8");
const backgroundSource = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");
const pageHelperSource = await readFile(new URL("../extension/page-helper.js", import.meta.url), "utf8");
const atsAdaptersSource = await readFile(new URL("../extension/content-modules/ats-adapters.js", import.meta.url), "utf8");
const sheetsSource = await readFile(new URL("../server/sheets/google-sheets.js", import.meta.url), "utf8");

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

test("resume upload never sends a backend-local filesystem path to Chrome", () => {
  assert.doesNotMatch(backgroundSource, /DOM\.setFileInputFiles/);
  assert.match(backgroundSource, /function setDebuggerFileInputFromBytes/);
  assert.match(backgroundSource, /const file = new File\(\[bytes\], filename/);
  assert.match(backgroundSource, /const transfer = new DataTransfer\(\)/);
  assert.match(backgroundSource, /transport: "bytes"/);
  assert.doesNotMatch(sheetsSource, /autobid-resumes|local_path|writeResumePayloadToLocalFile/);
});

test("known resume inputs try the debugger-free byte upload before the native fallback", () => {
  const setFileSource = contentSource.slice(
    contentSource.indexOf("async function setFileInputValue(input"),
    contentSource.indexOf("async function setFileInputValueNatively")
  );
  assert.ok(setFileSource.indexOf("setFileInputValueInPage") < setFileSource.indexOf("setFileInputValueNatively"));
  assert.match(contentSource, /NATIVE_FILE_UPLOAD[\s\S]*mime_type:[\s\S]*base64/);
  assert.match(contentSource, /NATIVE_FILE_CHOOSER_UPLOAD[\s\S]*mime_type:[\s\S]*base64/);
  assert.doesNotMatch(contentSource, /reason: "no-local-path"/);
});

test("page upload is idempotent and gives managed Dropzone exactly one change event", () => {
  const uploadSource = pageHelperSource.slice(
    pageHelperSource.indexOf("function uploadFile(input"),
    pageHelperSource.indexOf("function fileFromBase64")
  );
  assert.match(pageHelperSource, /PAGE_HELPER_BUILD_ID/);
  assert.match(pageHelperSource, /event\.stopImmediatePropagation\(\)/);
  assert.match(uploadSource, /data-auto-bid-file-upload-key/);
  assert.match(uploadSource, /if \(!isManagedDropzoneInput\(input\)\) dispatchFileInput\(input, transfer, "input"\)/);
  assert.match(uploadSource, /dispatchFileInput\(input, transfer, "change"\)/);
  assert.doesNotMatch(uploadSource, /dispatchFileDrop|callReactHandlers/);
});

test("Teamtailor waits for Dropzone completion and does not cascade upload transports", () => {
  assert.match(contentSource, /const RESUME_MANAGED_UPLOAD_TIMEOUT_MS = 30000;/);
  assert.match(contentSource, /function getManagedResumeUploadStatus/);
  assert.match(contentSource, /data-forms--inputs--upload-preview-target='urlInput'/);
  assert.match(contentSource, /already been processed or was rejected/);
  assert.match(contentSource, /if \(pageResult\.attempted\) return false/);
  assert.match(contentSource, /if \(nativeResult\.attempted\) return false/);
  assert.match(contentSource, /resume:managed-upload-resumed/);
  assert.match(contentSource, /getManagedResumeUploadContext\(input\)\.managed \? \["change"\] : \["input", "change"\]/);
  assert.match(backgroundSource, /duplicateSuppressed: true/);
  assert.match(contentSource, /function getTeamtailorResumeFileInputs/);
  assert.match(contentSource, /data-forms--inputs--upload-required-value/);
  assert.match(contentSource, /data-forms--inputs--upload-accepted-files-value/);
  assert.match(atsAdaptersSource, /input\.dz-hidden-input/);
  assert.match(atsAdaptersSource, /#job-application-form\[data-controller~='careersite--form'\]/);
});

test("BambooHR required resume input wins over its optional cover-letter input", () => {
  assert.match(contentSource, /atsAdapters\?\.describe\?\.\(\)\.id === "bamboohr"/);
  assert.match(contentSource, /input\.getAttribute\("aria-required"\) === "true"\)\) score \+= 125/);
  assert.match(contentSource, /if \(requiredInputs\.length === 1\) return requiredInputs/);
  assert.match(contentSource, /const requiredResumeInputs = getResumeFileInputs\(\)\.filter/);
  assert.match(contentSource, /requiredResumeInputs\.some\(\(input\) => !isResumeInputAttached\(input\)\)/);
  assert.match(contentSource, /function getBambooResumeFileInputFallback/);
  assert.match(contentSource, /resume:bamboo-input-mapped/);
  assert.match(contentSource, /getFieldContainer\(input\)\?\.textContent/);
});
