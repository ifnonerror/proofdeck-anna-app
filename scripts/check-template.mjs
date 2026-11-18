import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

import { FIXTURE_DECK } from "../bundle/fixture.js";
import { clone } from "../bundle/model.js";
import { buildFromTemplate, inspectTemplate, renderPresentation } from "../engine/office-engine.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const inputPath = path.resolve(process.argv[2] || "");
if (!process.argv[2]) {
  throw new Error("Usage: node scripts/check-template.mjs <template.pptx>");
}

const sourceBytes = new Uint8Array(await readFile(inputPath));
const info = await inspectTemplate(sourceBytes);
const built = await buildFromTemplate(sourceBytes, clone(FIXTURE_DECK));
const rendered = await renderPresentation(built.bytes);
const outputDir = path.join(root, "out");
await mkdir(outputDir, { recursive: true });
const outputPath = path.join(outputDir, "template-reuse-check.pptx");
await writeFile(outputPath, built.bytes);

const sourceZip = await JSZip.loadAsync(sourceBytes);
const outputZip = await JSZip.loadAsync(built.bytes);
const hash = async (zip, name) => createHash("sha256")
  .update(await zip.file(name).async("uint8array"))
  .digest("hex");
const parts = (zip, pattern) => Object.keys(zip.files).filter((name) => pattern.test(name)).sort();
const preserved = async (pattern) => {
  const names = parts(sourceZip, pattern);
  let unchanged = 0;
  for (const name of names) {
    if (outputZip.file(name) && await hash(sourceZip, name) === await hash(outputZip, name)) unchanged += 1;
  }
  return { total: names.length, unchanged };
};

const result = {
  input: inputPath,
  output: outputPath,
  template: {
    slides: info.slideCount,
    layouts: info.layouts.length,
    fonts: info.fonts,
    theme: info.theme,
  },
  output: {
    bytes: built.bytes.length,
    slides: rendered.length,
    titledSlides: rendered.filter((slide) => slide.title).length,
    shapeAwareSvg: rendered.every((slide) => slide.svg.includes("data-pptx-shape-id")),
  },
  preserved: {
    themes: await preserved(/^ppt\/theme\/theme\d+\.xml$/),
    masters: await preserved(/^ppt\/slideMasters\/slideMaster\d+\.xml$/),
    layouts: await preserved(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/),
  },
};

console.log(JSON.stringify(result, null, 2));
