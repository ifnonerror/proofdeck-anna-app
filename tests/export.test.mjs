import test from "node:test";
import assert from "node:assert/strict";

import {
  isUnavailableFileBridge,
  safeFilename,
} from "../bundle/export.js";

test("safeFilename keeps readable unicode titles and removes unsafe punctuation", () => {
  assert.equal(safeFilename("Q3 / 华东区：增长复盘"), "Q3-华东区-增长复盘.pptx");
  assert.equal(safeFilename(""), "proofdeck.pptx");
});

test("isUnavailableFileBridge recognizes local harness stubs only", () => {
  assert.equal(isUnavailableFileBridge({ code: "not_implemented" }), true);
  assert.equal(isUnavailableFileBridge(new Error("endpoint not yet available")), true);
  assert.equal(isUnavailableFileBridge({ code: "quota_exceeded" }), false);
  assert.equal(isUnavailableFileBridge(new Error("文件上传失败（HTTP 500）")), false);
});
