import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as pptx from "@office-kit/pptx";

import { clone, createDeck, createSlide } from "../bundle/model.js";
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

test("free-form template text is replaced without leaving stale body copy", async () => {
  const pres = pptx.createPresentation();
  const addTextBox = (slide, text, y, size, height = 0.8) => {
    const shape = pptx.addSlideTextBox(slide, {
      x: pptx.inches(0.7),
      y: pptx.inches(y),
      w: pptx.inches(11.9),
      h: pptx.inches(height),
      text,
    });
    pptx.setShapeTextFormat(shape, { size });
    return shape;
  };

  const cover = pptx.addBlankSlide(pres);
  addTextBox(cover, "Old cover title", 0.8, 36);
  addTextBox(cover, "Old cover subtitle", 3.0, 22);

  const body = pptx.addBlankSlide(pres);
  addTextBox(body, "Old section title", 0.3, 30, 0.6);
  addTextBox(body, "OLD BODY COPY THAT MUST DISAPPEAR", 1.35, 19, 3.2);
  addTextBox(body, "OLD SECONDARY ANALYSIS", 5.0, 17, 0.9);
  addTextBox(body, "CONFIDENTIAL FOOTER", 7.1, 9, 0.3);

  const closing = pptx.addBlankSlide(pres);
  addTextBox(closing, "Old closing", 1.8, 72);
  addTextBox(closing, "Old closing subtitle", 4.0, 22);

  const template = await pptx.savePresentation(pres);
  const deck = createDeck({
    title: "Replacement deck",
    slides: [
      { ...createSlide("cover"), title: "New cover", subtitle: "New subtitle" },
      {
        ...createSlide("bullets"),
        title: "New body title",
        bullets: [
          { text: "First new point", sourceRefs: [] },
          { text: "Second new point", sourceRefs: [] },
        ],
      },
      { ...createSlide("closing"), title: "New closing", subtitle: "New close" },
    ],
  });

  const built = await buildFromTemplate(template, deck);
  const rendered = await renderPresentation(built.bytes);
  const allText = rendered.flatMap((slide) => slide.shapes.map((shape) => shape.text)).join("\n");
  assert.doesNotMatch(allText, /OLD BODY COPY|OLD SECONDARY ANALYSIS|Old cover title|Old closing/);
  assert.match(allText, /New body title/);
  assert.match(allText, /First new point/);
  assert.match(allText, /CONFIDENTIAL FOOTER/);
});
