import test from "node:test";
import assert from "node:assert/strict";

import {
  createDeck,
  createSlide,
  normalizeDeck,
  numberSource,
  THEMES,
} from "../bundle/model.js";

test("numberSource creates stable source ids", () => {
  const units = numberSource("First fact.\n\nSecond fact.\nThird fact.");
  assert.deepEqual(units.map((unit) => unit.id), ["S1", "S2", "S3"]);
  assert.equal(units[1].text, "Second fact.");
});

test("createDeck uses safe defaults", () => {
  const deck = createDeck();
  assert.equal(deck.theme, "executive");
  assert.equal(deck.language, "zh-CN");
  assert.deepEqual(deck.slides, []);
  assert.ok(THEMES.includes(deck.theme));
});

test("createSlide creates a usable comparison slide", () => {
  const slide = createSlide("comparison");
  assert.equal(slide.kind, "comparison");
  assert.equal(slide.leftTitle, "Option A");
  assert.equal(slide.rightTitle, "Option B");
  assert.ok(slide.id.startsWith("slide-"));
});

test("normalizeDeck drops invalid source references and constrains content", () => {
  const warnings = [];
  const deck = normalizeDeck({
    title: "Evidence-backed decision",
    theme: "unknown",
    language: "fr",
    sourceUnits: [{ id: "S1", text: "Revenue increased." }],
    slides: [
      {
        id: "one",
        kind: "cover",
        title: "Revenue is growing",
        notes: "Open with the result.",
      },
      {
        id: "two",
        kind: "bullets",
        title: "Growth came from new customers",
        bullets: [
          { text: "Revenue increased", sourceRefs: ["S1", "S9"] },
        ],
      },
      {
        id: "three",
        kind: "closing",
        title: "Approve next step",
      },
    ],
  }, warnings);

  assert.equal(deck.theme, "executive");
  assert.equal(deck.language, "zh-CN");
  assert.deepEqual(deck.slides[1].bullets[0].sourceRefs, ["S1"]);
  assert.match(warnings[0], /S9/);
});
