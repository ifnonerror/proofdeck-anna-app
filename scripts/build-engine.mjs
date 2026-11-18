import { build } from "esbuild";
import { copyFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const vendor = path.join(root, "bundle", "vendor");
await mkdir(vendor, { recursive: true });

await build({
  entryPoints: [
    path.join(root, "engine", "office-engine.js"),
    path.join(root, "engine", "document-engine.js"),
  ],
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  minify: true,
  sourcemap: false,
  outdir: vendor,
  entryNames: "[name]",
  metafile: false,
});

await copyFile(
  path.join(root, "node_modules", "pdfjs-dist", "build", "pdf.worker.min.mjs"),
  path.join(vendor, "pdf.worker.mjs")
);
console.log("Office engine bundle prepared");
