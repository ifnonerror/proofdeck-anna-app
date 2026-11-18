import * as pptx from "@office-kit/pptx";
import { renderSlideToSvg } from "@office-kit/pptx-preview";

function sourceSuffix(refs) {
  return refs?.length ? ` [${refs.join(" ")}]` : "";
}

function textLines(items) {
  return (items || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`).join("\n");
}

function normalizedText(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLocaleLowerCase();
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

function clearText(shape) {
  if (!shape || !pptx.hasShapeText(shape)) return;
  try {
    pptx.setShapeText(shape, "", { preserveFormatting: false });
  } catch {
    // Some uncommon text-bearing shapes are read-only.
  }
}

function fillTargets(targets, lines) {
  const values = (lines || []).filter(Boolean);
  if (!targets.length || !values.length) return;
  const area = (target) => (target.info.bounds?.w || 0) * (target.info.bounds?.h || 0);
  const largest = targets.reduce((best, target) => area(target) > area(best) ? target : best, targets[0]);
  const comparable = targets.filter((target) => area(target) >= area(largest) * 0.55);
  if (values.length > 1 && comparable.length >= 2) {
    const count = Math.min(values.length, comparable.length);
    for (let index = 0; index < count; index += 1) {
      setText(comparable[index].shape, values[index]);
    }
    return;
  }
  setText(largest.shape, values.join("\n"));
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

function fillTemplateSlide(pres, slide, data, fixedTexts) {
  if (replaceTokens(slide, data) > 0) {
    pptx.setSlideNotes(slide, data.notes || "");
    return slideInfo(pres, slide, 0);
  }

  const title = titleShape(pres, slide);
  const content = pickTextShapes(pres, slide).content;
  const replaceable = content.filter(({ info }) => !fixedTexts.has(normalizedText(info.text)));
  const titleInfo = title ? shapeInfo(pres, title) : null;
  const minimumBodySize = Math.min(18, Math.max(11, Number(titleInfo?.fontSize || 0) * 0.4));
  const targets = replaceable
    .filter(({ info }) => Number(info.fontSize || 0) >= minimumBodySize)
    .sort((a, b) => {
      const yDiff = (a.info.bounds?.y || 0) - (b.info.bounds?.y || 0);
      return Math.abs(yDiff) > 25400
        ? yDiff
        : (a.info.bounds?.x || 0) - (b.info.bounds?.x || 0);
    });

  if (title) setText(title, data.title);
  replaceable.forEach(({ shape }) => clearText(shape));

  if (data.kind === "cover" || data.kind === "closing" || data.kind === "section") {
    const subtitle = targets.find(({ info }) => /subtitle/i.test(String(info.placeholderType || "")))
      || targets[0];
    if (subtitle?.shape) setText(subtitle.shape, data.subtitle || textLines(data.bullets));
  } else if (data.kind === "comparison") {
    const size = pptx.getSlideSize(pres);
    const left = targets.filter(({ info }) => (info.bounds?.x || 0) < size.width / 2);
    const right = targets.filter(({ info }) => (info.bounds?.x || 0) >= size.width / 2);
    if (left.length && right.length) {
      setText(left[0]?.shape, [data.leftTitle, ...(data.leftItems || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`)].join("\n"));
      setText(right[0]?.shape, [data.rightTitle, ...(data.rightItems || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`)].join("\n"));
    } else if (targets[0]) {
      setText(targets[0].shape, `${data.leftTitle || "A"}\n${textLines(data.leftItems)}\n\n${data.rightTitle || "B"}\n${textLines(data.rightItems)}`);
    }
  } else if (data.kind === "metrics") {
    const lines = (data.metrics || []).map((metric) => `${metric.value}  ${metric.label}${sourceSuffix(metric.sourceRefs)}`);
    fillTargets(targets, lines);
  } else {
    fillTargets(targets, (data.bullets || []).map((item) => `${item.text}${sourceSuffix(item.sourceRefs)}`));
  }

  pptx.setSlideNotes(slide, data.notes || "");
  return slideInfo(pres, slide, 0);
}

function findFixedTexts(pres, slides) {
  const counts = new Map();
  for (const slide of slides) {
    const title = titleShape(pres, slide);
    for (const shape of pptx.getSlideShapes(slide)) {
      if (shape === title || !pptx.hasShapeText(shape)) continue;
      const key = normalizedText(pptx.getShapeText(shape));
      if (!key) continue;
      counts.set(key, (counts.get(key) || 0) + 1);
    }
  }
  const threshold = Math.max(2, Math.ceil(slides.length * 0.35));
  return new Set([...counts].filter(([, count]) => count >= threshold).map(([key]) => key));
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
  const fixedTexts = findFixedTexts(pres, original);

  const generated = [];
  for (let index = 0; index < deck.slides.length; index += 1) {
    const data = deck.slides[index];
    const donor = chooseDonor(original, index, data.kind);
    const slide = pptx.duplicateSlide(pres, donor);
    generated.push(fillTemplateSlide(pres, slide, data, fixedTexts));
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

