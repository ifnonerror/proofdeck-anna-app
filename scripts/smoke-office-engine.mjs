import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";

import { FIXTURE_DECK } from "../bundle/fixture.js";
import { clone } from "../bundle/model.js";
import {
  applyPptxPatch,
  buildFromTemplate,
  renderPresentation,
} from "../engine/office-engine.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const templatePath = path.join(root, "examples", "proofdeck-q3-sample.pptx");
const outputPath = path.join(root, "out", "proofdeck-template-engine.pptx");

await mkdir(path.dirname(outputPath), { recursive: true });
await rm(outputPath, { force: true });

const template = new Uint8Array(await readFile(templatePath));
const built = await buildFromTemplate(template, clone(FIXTURE_DECK));
const rendered = await renderPresentation(built.bytes);
assert.equal(rendered.length, FIXTURE_DECK.slides.length);
assert.ok(rendered.every((slide) => slide.svg.startsWith("<svg")));

const target = rendered[0].shapes.find((shape) => shape.text);
assert.ok(target);
const patched = await applyPptxPatch(built.bytes, [{
  slideIndex: 0,
  shapeId: target.id,
  action: "setTextStyle",
  fontSize: 18,
  bold: true,
  color: "#C95E38",
}]);
assert.equal(patched.applied.length, 1);
await writeFile(outputPath, Buffer.from(patched.bytes));

console.log(JSON.stringify({
  file: outputPath,
  bytes: patched.bytes.byteLength,
  slides: patched.slides.length,
  shapeAwareSvg: patched.slides.every((slide) => slide.svg.includes("data-pptx-shape-id")),
}));
