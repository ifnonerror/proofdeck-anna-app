import { collectRefs } from "./model.js";
import { layoutSlide } from "./layout.js";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PAGE_W = 13.333;
const PAGE_H = 7.5;

function box(region) {
  return {
    x: +(region.x / 1280 * PAGE_W).toFixed(3),
    y: +(region.y / 720 * PAGE_H).toFixed(3),
    w: +(region.w / 1280 * PAGE_W).toFixed(3),
    h: +(region.h / 720 * PAGE_H).toFixed(3),
  };
}

function pointSize(canvasPx) {
  return Math.max(8, Math.round(canvasPx * PAGE_H / 720 * 72));
}

function refSuffix(refs) {
  return refs?.length ? `  [${refs.join(" ")}]` : "";
}

function itemRuns(items) {
  return (items || []).filter((item) => item.text).map((item) => ({
    text: item.text + refSuffix(item.sourceRefs),
    options: { bullet: { indent: 18 }, breakLine: true },
  }));
}

export function buildPresentation(deck, PptxGenJS) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "PROOFDECK_WIDE", width: PAGE_W, height: PAGE_H });
  pptx.layout = "PROOFDECK_WIDE";
  pptx.title = deck.title || "ProofDeck";
  pptx.subject = deck.subtitle || "Source-grounded presentation";
  pptx.author = "ProofDeck";
  pptx.company = "ProofDeck";

  for (const slide of deck.slides || []) {
    const layout = layoutSlide(slide, deck.theme, deck.customTheme);
    const t = layout.tokens;
    const out = pptx.addSlide();
    out.background = { color: t.background };

    if (slide.kind !== "cover" && slide.kind !== "closing") {
      const rule = box({ x: 66, y: 132, w: 1148, h: 4 });
      out.addShape(pptx.ShapeType.rect, {
        ...rule,
        fill: { color: t.accent },
        line: { color: t.accent, transparency: 100 },
      });
    }

    if (slide.kind === "cover" || slide.kind === "closing") {
      const kicker = box(layout.regions.kicker);
      out.addText(slide.kind === "cover" ? "PROOFDECK · DECISION BRIEF" : "NEXT DECISION", {
        ...kicker,
        fontFace: t.bodyFont,
        fontSize: pointSize(t.metaSize),
        bold: true,
        color: t.accent,
        charSpacing: 2,
      });
      const title = box(layout.regions.title);
      out.addText(slide.title || "", {
        ...title,
        fontFace: t.titleFont,
        fontSize: pointSize(t.titleSize),
        bold: true,
        color: t.foreground,
        valign: "top",
        fit: "shrink",
        breakLine: false,
      });
      if (slide.subtitle) {
        const subtitle = box(layout.regions.subtitle);
        out.addText(slide.subtitle, {
          ...subtitle,
          fontFace: t.bodyFont,
          fontSize: pointSize(t.bodySize),
          color: t.muted,
          valign: "top",
          fit: "shrink",
        });
      }
    } else {
      const title = box(layout.regions.title);
      out.addText(slide.title || "", {
        ...title,
        fontFace: t.titleFont,
        fontSize: pointSize(t.titleSize),
        bold: true,
        color: t.foreground,
        valign: "top",
        fit: "shrink",
      });

      if (slide.kind === "comparison") {
        const makeColumn = (region, columnTitle, items) => {
          const b = box(region);
          out.addShape(pptx.ShapeType.roundRect, {
            ...b,
            fill: { color: t.card },
            line: { color: t.accent, width: 1 },
            radius: 0.12,
          });
          out.addText([
            {
              text: columnTitle || "",
              options: {
                fontSize: pointSize(16),
                bold: true,
                color: t.muted,
                breakLine: true,
              },
            },
            ...itemRuns(items),
          ], {
            x: b.x + 0.16,
            y: b.y + 0.16,
            w: b.w - 0.32,
            h: b.h - 0.32,
            fontFace: t.bodyFont,
            fontSize: pointSize(t.bodySize),
            color: t.foreground,
            valign: "top",
            fit: "shrink",
          });
        };
        makeColumn(layout.regions.left, slide.leftTitle, slide.leftItems);
        makeColumn(layout.regions.right, slide.rightTitle, slide.rightItems);
      } else if (slide.kind === "metrics") {
        const body = box(layout.regions.body);
        const metrics = slide.metrics || [];
        const gap = 0.2;
        const count = Math.max(1, metrics.length);
        const cardW = (body.w - gap * (count - 1)) / count;
        metrics.forEach((metric, index) => {
          const x = body.x + index * (cardW + gap);
          out.addShape(pptx.ShapeType.roundRect, {
            x,
            y: body.y,
            w: cardW,
            h: body.h * 0.68,
            fill: { color: t.card },
            line: { color: t.accent, width: 1.2 },
            radius: 0.12,
          });
          out.addText([
            {
              text: metric.value || "",
              options: {
                fontSize: pointSize(43),
                bold: true,
                color: t.foreground,
                align: "center",
                breakLine: true,
              },
            },
            {
              text: metric.label || "",
              options: {
                fontSize: pointSize(15),
                color: t.muted,
                align: "center",
                breakLine: true,
              },
            },
            {
              text: refSuffix(metric.sourceRefs),
              options: { fontSize: pointSize(10), color: t.muted, align: "center" },
            },
          ], {
            x,
            y: body.y,
            w: cardW,
            h: body.h * 0.68,
            fontFace: t.bodyFont,
            valign: "middle",
            fit: "shrink",
          });
        });
      } else {
        const body = box(layout.regions.body);
        const lines = itemRuns(slide.bullets);
        if (lines.length) {
          out.addText(lines, {
            ...body,
            fontFace: t.bodyFont,
            fontSize: pointSize(t.bodySize),
            color: t.foreground,
            valign: "top",
            fit: "shrink",
            paraSpaceAfterPt: 10,
          });
        }
      }
    }

    const refs = collectRefs(slide);
    if (refs.length) {
      const footer = box(layout.regions.footer);
      out.addText(`Sources: ${refs.join(" · ")}`, {
        ...footer,
        fontFace: t.bodyFont,
        fontSize: pointSize(t.metaSize),
        color: t.muted,
        align: "right",
        valign: "bottom",
      });
    }

    if (slide.notes) out.addNotes(slide.notes);
  }

  return pptx;
}

export async function createPptxArrayBuffer(deck, PptxGenJS) {
  const pptx = buildPresentation(deck, PptxGenJS);
  return pptx.write({ outputType: "arraybuffer" });
}

export function safeFilename(title) {
  const cleaned = String(title || "proofdeck")
    .replace(/[^a-zA-Z0-9\u3400-\u9fff-_]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 72);
  return `${cleaned || "proofdeck"}.pptx`;
}

export function browserDownload(buffer, filename) {
  const blob = new Blob([buffer], { type: PPTX_MIME });
  const href = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = href;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(href), 3000);
}

export function isUnavailableFileBridge(error) {
  const code = String(error?.code || error?.name || "").toLowerCase();
  const message = String(error?.message || error || "").toLowerCase();
  return code === "not_implemented"
    || code === "unknown_method"
    || message.includes("endpoint not yet available")
    || message.includes("not implemented");
}

export async function deliverPptxBuffer({ anna, buffer, filename, deckId = "deck" }) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  if (!anna?.files?.upload_init || !anna?.files?.download) {
    browserDownload(bytes, filename);
    return { mode: "local", filename, bytes: bytes.byteLength };
  }

  const path = `exports/${deckId || Date.now()}/${Date.now()}-${filename}`;
  try {
    const upload = await anna.files.upload_init({
      path,
      content_type: PPTX_MIME,
      size: bytes.byteLength,
    });
    const response = await fetch(upload.put_url, {
      method: "PUT",
      headers: upload.headers || { "Content-Type": PPTX_MIME },
      body: bytes,
    });
    if (!response.ok) throw new Error(`文件上传失败（HTTP ${response.status}）`);
    const etag = response.headers.get("etag") || upload.etag;
    await anna.files.upload_finalize({
      path,
      ...(etag ? { etag } : {}),
      size_bytes: bytes.byteLength,
    });
    await anna.files.download({ path, filename });
    return { mode: "anna", filename, bytes: bytes.byteLength, path };
  } catch (error) {
    if (!isUnavailableFileBridge(error)) throw error;
    browserDownload(bytes, filename);
    return {
      mode: "local",
      filename,
      bytes: bytes.byteLength,
      bridgeFallback: true,
    };
  }
}

export async function exportPresentation({ anna, deck, PptxGenJS }) {
  const buffer = await createPptxArrayBuffer(deck, PptxGenJS);
  const filename = safeFilename(deck.title);
  return deliverPptxBuffer({ anna, buffer, filename, deckId: deck.id });
}
