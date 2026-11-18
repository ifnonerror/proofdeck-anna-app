import mammoth from "mammoth/mammoth.browser.js";
import { GlobalWorkerOptions, getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

const decoder = new TextDecoder();

function cleanText(value) {
  return String(value || "")
    .replace(/\u0000/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{4,}/g, "\n\n\n")
    .trim();
}

export function configurePdfWorker(url) {
  GlobalWorkerOptions.workerSrc = String(url);
}

export async function extractDocumentText(fileName, mimeType, bytes) {
  const name = String(fileName || "").toLowerCase();
  const data = new Uint8Array(bytes);
  if (name.endsWith(".pdf") || mimeType === "application/pdf") {
    const pdf = await getDocument({ data, isEvalSupported: false }).promise;
    const pages = [];
    for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
      const page = await pdf.getPage(pageIndex);
      const content = await page.getTextContent();
      pages.push(content.items.map((item) => item.str).join(" "));
    }
    return cleanText(pages.join("\n\n"));
  }
  if (name.endsWith(".docx")) {
    const arrayBuffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
    const result = await mammoth.extractRawText({ arrayBuffer });
    return cleanText(result.value);
  }
  if (/\.(txt|md|markdown|csv)$/i.test(name) || String(mimeType || "").startsWith("text/")) {
    return cleanText(decoder.decode(data));
  }
  throw new Error("目前支持 PDF、DOCX、TXT 和 Markdown 资料。");
}
