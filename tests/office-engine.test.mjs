import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { clone } from "../bundle/model.js";
import { FIXTURE_DECK } from "../bundle/fixture.js";
import {
  applyPptxPatch,
  buildFromTemplate,
  inspectTemplate,
  renderPresentation,
} from "../engine/office-engine.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const templatePath = path.join(root, "examples", "proofdeck-q3-sample.pptx");

test("template engine clones existing slides, saves, and renders shape-aware SVG", async () => {
  const template = new Uint8Array(await readFile(templatePath));
  const info = await inspectTemplate(template);
  assert.ok(info.slideCount >= 8);
  assert.ok(info.layouts.length >= 1);

  const deck = clone(FIXTURE_DECK);
  const built = await buildFromTemplate(template, deck);
  const rendered = await renderPresentation(built.bytes);
  assert.equal(rendered.length, deck.slides.length);
  assert.ok(rendered.every((slide) => slide.svg.startsWith("<svg")));
  assert.ok(rendered.some((slide) => slide.svg.includes("data-pptx-shape-id")));
});

test("single-element patch preserves the rest of the template deck", async () => {
  const template = new Uint8Array(await readFile(templatePath));
  const built = await buildFromTemplate(template, clone(FIXTURE_DECK));
  const rendered = await renderPresentation(built.bytes);
  const target = rendered[1].shapes.find((shape) => shape.text);
  assert.ok(target);

  const patched = await applyPptxPatch(built.bytes, [{
    slideIndex: 1,
    shapeId: target.id,
    action: "setTextStyle",
    fontSize: 30,
    bold: true,
    color: "#C95E38",
  }]);

  assert.equal(patched.applied.length, 1);
  assert.equal(patched.slides.length, FIXTURE_DECK.slides.length);
  assert.ok(patched.slides[1].svg.startsWith("<svg"));
});
