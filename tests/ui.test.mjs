import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const html = await readFile(path.join(root, "bundle", "index.html"), "utf8");
const app = await readFile(path.join(root, "bundle", "app.js"), "utf8");
const styles = await readFile(path.join(root, "bundle", "styles.css"), "utf8");

test("chat starts without exposing the presentation workspace", () => {
  assert.match(html, /class="studio chat-only"/);
  assert.match(html, /id="viewer-panel"[^>]+hidden/);
  assert.match(html, /id="close-viewer"/);
  assert.match(app, /appendDeckLink\(\)/);
  assert.match(app, /state\.workspaceOpen = true/);
});

test("only the model chooser remains in the composer", () => {
  assert.doesNotMatch(html, /advanced-settings|brief-slides|brief-language|brief-tone/);
  assert.match(html, /id="model-settings" class="model-picker"/);
  assert.doesNotMatch(html, />Auto</);
  assert.match(html, /PPT template \(optional\)/);
});

test("model label refreshes from the Anna account after returning from settings", () => {
  assert.match(app, /async function refreshCurrentModel/);
  assert.match(app, /pendingModelRefresh/);
  assert.match(app, /addEventListener\("focus"/);
});

test("top bar contains only actionable controls", () => {
  assert.match(html, /id="home-button" class="brand"/);
  assert.match(html, /id="new-project"/);
  assert.doesNotMatch(html, /project-label|mode-badge|quota-badge/);
});

test("task sidebar supports switching independently persisted projects", () => {
  assert.match(html, /class="task-sidebar"/);
  assert.match(html, /id="task-list"/);
  assert.match(app, /TASK_LIST_KEY/);
  assert.match(app, /async function switchTask/);
  assert.match(app, /messages: clone\(state\.messages\)/);
});

test("slide thumbnails reserve their full aspect-ratio height", () => {
  assert.match(styles, /\.slide-filmstrip\s*\{[\s\S]*grid-auto-rows:\s*max-content/);
  assert.match(styles, /\.filmstrip-item\s*\{[\s\S]*height:\s*auto[\s\S]*aspect-ratio:\s*16\s*\/\s*9/);
});

test("selected text shapes expose a direct manual editor", () => {
  assert.match(html, /id="manual-edit-panel"/);
  assert.match(html, /id="manual-text"/);
  assert.match(html, /id="apply-manual-edit"/);
  assert.match(app, /async function applyManualEdit/);
  assert.match(app, /action: "setTextStyle"/);
});

test("generation warnings are summarized instead of printed as a list", () => {
  assert.match(app, /function summarizeGenerationWarnings/);
  assert.match(app, /已自动清理 \$\{invalidReferences\} 个无效来源引用/);
  assert.doesNotMatch(app, /result\.warnings\.join\("；"\)/);
});
