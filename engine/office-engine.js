import * as pptx from "@office-kit/pptx";
import { renderSlideToSvg } from "@office-kit/pptx-preview";

function sourceSuffix(refs) {
  return refs?.length ? ` [${refs.join(" ")}]` : "";
}

function textLines(items) {
  return (items || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`).join("\n");
}

function shapeInfo(pres, shape) {
  const text = pptx.hasShapeText(shape) ? pptx.getShapeText(shape) : "";
  let format = null;
  try {
    format = pptx.getShapeRunFormatEffective(pres, shape, 0, 0);
  } catch {
    // Some non-text or uncommon shapes do not expose a first run.
  }
  return {
    id: pptx.getShapeId(shape),
    name: pptx.getShapeName(shape),
    kind: pptx.getShapeKind(shape),
    placeholderType: pptx.getShapePlaceholderType(shape),
    text,
    bounds: pptx.getShapeBounds(shape),
    fontSize: Number(format?.size || 0),
  };
}

function slideInfo(pres, slide, index) {
  const shapes = pptx.getSlideShapes(slide).map((shape) => shapeInfo(pres, shape));
  return {
    index,
    title: pptx.getSlideTitle(slide) || inferTitle(pres, slide),
    notes: pptx.getSlideNotes(slide) || "",
    layout: pptx.getSlideLayoutName(pptx.getSlideLayout(slide)),
    shapes,
  };
}

function titleShape(pres, slide) {
  const shapes = pptx.getSlideShapes(slide).filter((shape) => pptx.hasShapeText(shape));
  const placeholder = shapes.find((shape) => {
    const type = String(pptx.getShapePlaceholderType(shape) || "").toLowerCase();
    return type === "title" || type === "ctrtitle";
  });
  if (placeholder) return placeholder;
  return shapes
    .map((shape) => ({
      shape,
      info: shapeInfo(pres, shape),
    }))
    .filter(({ info }) => info.text)
    .sort((a, b) => {
      const sizeDiff = b.info.fontSize - a.info.fontSize;
      if (Math.abs(sizeDiff) > 1) return sizeDiff;
      return (a.info.bounds?.y || 0) - (b.info.bounds?.y || 0);
    })[0]?.shape || null;
}

function inferTitle(pres, slide) {
  const shape = titleShape(pres, slide);
  return shape && pptx.hasShapeText(shape) ? pptx.getShapeText(shape) : null;
}

function pickTextShapes(pres, slide) {
  const size = pptx.getSlideSize(pres);
  const title = titleShape(pres, slide);
  const all = pptx.getSlideShapes(slide)
    .filter((shape) => pptx.hasShapeText(shape))
    .map((shape) => ({ shape, info: shapeInfo(pres, shape) }))
    .filter(({ info }) => info.text && info.bounds);
  const content = all.filter(({ shape, info }) => {
    if (shape === title) return false;
    const y = info.bounds?.y || 0;
    return y > size.height * 0.08 && y < size.height * 0.88;
  });
  return { title, content };
}

function setText(shape, value) {
  if (!shape || !value) return;
  pptx.setShapeText(shape, value, {
    newlines: "paragraph",
    preserveFormatting: false,
  });
  try {
    pptx.setShapeTextAutoFit(shape, "normal");
  } catch {
    // Not every text-bearing object supports autofit.
  }
}

function replaceTokens(slide, data) {
  return pptx.replaceTokensInSlide(slide, {
    title: data.title || "",
    subtitle: data.subtitle || "",
    body: textLines(data.bullets),
    left_title: data.leftTitle || "",
    right_title: data.rightTitle || "",
    left_body: textLines(data.leftItems),
    right_body: textLines(data.rightItems),
    notes: data.notes || "",
  });
}

function fillTemplateSlide(pres, slide, data) {
  if (replaceTokens(slide, data) > 0) {
    if (data.notes) pptx.setSlideNotes(slide, data.notes);
    return slideInfo(pres, slide, 0);
  }

  const title = titleShape(pres, slide);
  const content = pickTextShapes(pres, slide).content;
  if (title) setText(title, data.title);

  if (data.kind === "cover" || data.kind === "closing" || data.kind === "section") {
    const subtitle = content.find(({ info }) => /subtitle/i.test(String(info.placeholderType || "")))?.shape
      || content[0]?.shape;
    if (subtitle) setText(subtitle, data.subtitle || textLines(data.bullets));
  } else if (data.kind === "comparison") {
    const size = pptx.getSlideSize(pres);
    const left = content.filter(({ info }) => (info.bounds?.x || 0) < size.width / 2);
    const right = content.filter(({ info }) => (info.bounds?.x || 0) >= size.width / 2);
    if (left.length && right.length) {
      setText(left[0]?.shape, [data.leftTitle, ...(data.leftItems || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`)].join("\n"));
      setText(right[0]?.shape, [data.rightTitle, ...(data.rightItems || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`)].join("\n"));
    } else if (content[0]) {
      setText(content[0].shape, `${data.leftTitle || "A"}\n${textLines(data.leftItems)}\n\n${data.rightTitle || "B"}\n${textLines(data.rightItems)}`);
    }
  } else if (data.kind === "metrics") {
    const metrics = data.metrics || [];
    if (content.length >= metrics.length * 2) {
      metrics.forEach((metric, index) => {
        setText(content[index * 2]?.shape, metric.value);
        setText(content[index * 2 + 1]?.shape, `${metric.label}${sourceSuffix(metric.sourceRefs)}`);
      });
    } else {
      const lines = metrics.map((metric) => `${metric.value}  ${metric.label}${sourceSuffix(metric.sourceRefs)}`);
      setText(content[0]?.shape, lines.join("\n"));
    }
  } else if (content[0]) {
    setText(content[0].shape, textLines(data.bullets));
  }

  const used = new Set([title, content[0]?.shape]);
  for (const { shape } of content.slice(1)) {
    if (!used.has(shape) && pptx.getShapePlaceholderType(shape)) {
      pptx.setShapeText(shape, "");
    }
  }
  if (data.notes) pptx.setSlideNotes(slide, data.notes);
  return slideInfo(pres, slide, 0);
}

function chooseDonor(slides, index, kind) {
  if (!slides.length) return null;
  if (kind === "cover") return slides[0];
  if (kind === "closing") return slides[slides.length - 1];
  const count = Math.max(1, slides.length - 2);
  return slides[1 + ((index - 1) % count)] || slides[slides.length - 1];
}

export async function inspectTemplate(bytes) {
  const pres = await pptx.loadPresentation(new Uint8Array(bytes));
  const slides = pptx.getSlides(pres);
  return {
    slideCount: slides.length,
    size: pptx.getSlideSize(pres),
    theme: pptx.getPresentationTheme(pres),
    fonts: pptx.getPresentationFonts(pres),
    layouts: [...new Set(pptx.getSlideLayouts(pres).map((layout) => pptx.getSlideLayoutName(layout)))],
    slides: slides.map((slide, index) => ({
      ...slideInfo(pres, slide, index),
      layout: pptx.getSlideLayoutName(pptx.getSlideLayout(slide)),
    })),
  };
}

export async function buildFromTemplate(templateBytes, deck) {
  const pres = await pptx.loadPresentation(new Uint8Array(templateBytes));
  const original = [...pptx.getSlides(pres)];
  if (!original.length) throw new Error("上传的模板没有可复用页面。");

  const generated = [];
  for (let index = 0; index < deck.slides.length; index += 1) {
    const data = deck.slides[index];
    const donor = chooseDonor(original, index, data.kind);
    const slide = pptx.duplicateSlide(pres, donor);
    generated.push(fillTemplateSlide(pres, slide, data));
  }
  original.forEach((slide) => pptx.removeSlide(pres, slide));
  pptx.compactPackage(pres);
  const bytes = await pptx.savePresentation(pres);
  return {
    bytes,
    slides: generated,
  };
}

export async function renderPresentation(bytes) {
  const pres = await pptx.loadPresentation(new Uint8Array(bytes));
  return pptx.getSlides(pres).map((slide, index) => ({
    index,
    title: pptx.getSlideTitle(slide) || inferTitle(pres, slide),
    notes: pptx.getSlideNotes(slide) || "",
    svg: renderSlideToSvg(pres, slide),
    shapes: pptx.getSlideShapes(slide).map((shape) => shapeInfo(pres, shape)),
  }));
}

export async function applyPptxPatch(bytes, ops) {
  const pres = await pptx.loadPresentation(new Uint8Array(bytes));
  const slides = pptx.getSlides(pres);
  const applied = [];
  for (const op of ops || []) {
    const slide = slides[Number(op.slideIndex)];
    if (!slide) continue;
    const shape = pptx.getSlideShapes(slide).find((candidate) => pptx.getShapeId(candidate) === Number(op.shapeId));
    if (!shape) continue;
    if (op.action === "replaceText" && typeof op.text === "string") {
      pptx.setShapeText(shape, op.text, { preserveFormatting: false });
    } else if (op.action === "setTextStyle") {
      const format = {};
      if (op.fontSize) format.size = Math.max(6, Math.min(120, Number(op.fontSize)));
      if (op.color && /^#[0-9a-f]{6}$/i.test(op.color)) format.color = op.color;
      if (op.font) format.font = String(op.font).slice(0, 80);
      if (typeof op.bold === "boolean") format.bold = op.bold;
      if (typeof op.italic === "boolean") format.italic = op.italic;
      if (typeof op.underline === "boolean") format.underline = op.underline;
      if (Object.keys(format).length) pptx.setShapeTextFormat(shape, format);
    } else if (op.action === "move") {
      const bounds = pptx.getShapeBounds(shape);
      if (bounds) {
        pptx.setShapePosition(shape, Math.round(bounds.x + Number(op.dx || 0)), Math.round(bounds.y + Number(op.dy || 0)));
      }
    } else if (op.action === "resize") {
      const bounds = pptx.getShapeBounds(shape);
      if (bounds) {
        pptx.setShapeSize(shape, Math.max(1, Math.round(bounds.w + Number(op.dw || 0))), Math.max(1, Math.round(bounds.h + Number(op.dh || 0))));
      }
    } else if (op.action === "setFill" && /^#[0-9a-f]{6}$/i.test(String(op.color || ""))) {
      pptx.setShapeFill(shape, op.color);
    }
    applied.push(op);
  }
  pptx.touchModified(pres);
  const nextBytes = await pptx.savePresentation(pres);
  return {
    bytes: nextBytes,
    applied,
    slides: await renderPresentation(nextBytes),
  };
}

