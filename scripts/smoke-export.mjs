import { createRequire } from "node:module";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";

import { FIXTURE_DECK } from "../bundle/fixture.js";
import { createPptxArrayBuffer } from "../bundle/export.js";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");
const JSZip = require("jszip");
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(root, "out");
const outputPath = path.join(outputDir, "proofdeck-fixture.pptx");

await mkdir(outputDir, { recursive: true });
await rm(outputPath, { force: true });

const buffer = await createPptxArrayBuffer(FIXTURE_DECK, PptxGenJS);
await writeFile(outputPath, Buffer.from(buffer));
const zip = await JSZip.loadAsync(buffer);
const files = Object.keys(zip.files);
const slideFiles = files.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
const noteFiles = files.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name));

assert.equal(slideFiles.length, FIXTURE_DECK.slides.length, "slide count must match deck");
assert.equal(noteFiles.length, FIXTURE_DECK.slides.length, "every slide must have speaker notes");
assert.ok(files.includes("ppt/presentation.xml"));
assert.ok(files.includes("[Content_Types].xml"));

const presentation = await zip.file("ppt/presentation.xml").async("string");
assert.match(presentation, /<p:sldSz cx="\d+" cy="\d+"/);
const width = Number(presentation.match(/<p:sldSz cx="(\d+)"/)?.[1]);
const height = Number(presentation.match(/cy="(\d+)"/)?.[1]);
assert.ok(Math.abs(width / height - 16 / 9) < 0.01, "presentation must be 16:9");

const allSlideXml = (await Promise.all(
  slideFiles.map((name) => zip.file(name).async("string"))
)).join("\n");
assert.match(allSlideXml, /Q3 增长达标/);
assert.match(allSlideXml, /Sources:/);
assert.match(allSlideXml, /<p:sp>/);

const allNoteXml = (await Promise.all(
  noteFiles.map((name) => zip.file(name).async("string"))
)).join("\n");
assert.match(allNoteXml, /开场 30 秒/);

const fileStat = await readFile(outputPath);
console.log(JSON.stringify({
  file: outputPath,
  bytes: fileStat.byteLength,
  slides: slideFiles.length,
  notes: noteFiles.length,
  ratio: Number((width / height).toFixed(4)),
}));
