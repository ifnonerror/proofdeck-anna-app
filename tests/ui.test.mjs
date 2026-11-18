import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const html = await readFile(path.join(root, "bundle", "index.html"), "utf8");
const app = await readFile(path.join(root, "bundle", "app.js"), "utf8");

test("chat starts without exposing the presentation workspace", () => {
  assert.match(html, /class="studio chat-only"/);
  assert.match(html, /id="viewer-panel"[^>]+hidden/);
  assert.match(html, /id="close-viewer"/);
  assert.match(app, /appendDeckLink\(\)/);
  assert.match(app, /state\.workspaceOpen = true/);
});

test("only the model chooser remains in the composer", () => {
  assert.doesNotMatch(html, /advanced-settings|brief-slides|brief-language|brief-tone/);
  assert.match(html, /class="model-picker"[\s\S]+id="brief-model"/);
});
