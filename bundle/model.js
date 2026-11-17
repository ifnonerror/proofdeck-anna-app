export const THEMES = ["executive", "minimal", "signal"];
export const SLIDE_KINDS = [
  "cover",
  "agenda",
  "section",
  "bullets",
  "comparison",
  "timeline",
  "metrics",
  "closing",
];

export function uid(prefix = "slide") {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

export function numberSource(text) {
  return String(text || "")
    .split(/\n\s*\n|\r?\n/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part, index) => ({ id: `S${index + 1}`, text: part }));
}

export function createSlide(kind = "bullets") {
  return {
    id: uid(),
    kind: SLIDE_KINDS.includes(kind) ? kind : "bullets",
    title: kind === "section" ? "New section" : "New slide",
    subtitle: "",
    bullets: [],
    leftTitle: kind === "comparison" ? "Option A" : "",
    rightTitle: kind === "comparison" ? "Option B" : "",
    leftItems: [],
    rightItems: [],
    metrics: [],
    notes: "",
  };
}

export function createDeck(overrides = {}) {
  return {
    id: uid("deck"),
    title: "Untitled deck",
    subtitle: "",
    theme: "executive",
    language: "zh-CN",
    sourceText: "",
    sourceUnits: [],
    slides: [],
    updatedAt: new Date().toISOString(),
    ...overrides,
  };
}

export function clampText(value, max) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text.length > max ? text.slice(0, max).trim() : text;
}

export function normalizeRefs(refs, validIds) {
  const valid = new Set(validIds);
  return [...new Set((Array.isArray(refs) ? refs : [])
    .map((ref) => String(ref || "").trim().toUpperCase())
    .filter((ref) => /^S\d+$/.test(ref) && valid.has(ref)))];
}

export function normalizeDeck(input, warnings = []) {
  const validIds = (input.sourceUnits || []).map((unit) => unit.id);
  const slides = Array.isArray(input.slides) ? input.slides : [];
  return createDeck({
    id: input.id || uid("deck"),
    title: clampText(input.title || "Untitled deck", 100),
    subtitle: clampText(input.subtitle || "", 140),
    theme: THEMES.includes(input.theme) ? input.theme : "executive",
    language: input.language === "en" ? "en" : "zh-CN",
    sourceText: String(input.sourceText || ""),
    sourceUnits: validIds.map((id, index) => ({
      id,
      text: clampText(input.sourceUnits[index].text || "", 6000),
    })),
    slides: slides.map((slide, index) => {
      const kind = SLIDE_KINDS.includes(slide.kind)
        ? slide.kind
        : index === 0
          ? "cover"
          : index === slides.length - 1
            ? "closing"
            : "bullets";
      const refItems = (items, field) =>
        (Array.isArray(items) ? items : []).slice(0, 5).map((item) => {
          const suppliedRefs = Array.isArray(item?.sourceRefs) ? item.sourceRefs : [];
          const normalizedRefs = normalizeRefs(suppliedRefs, validIds);
          const next = {
            text: clampText(typeof item === "string" ? item : item.text, 100),
            sourceRefs: normalizedRefs,
          };
          const invalidRefs = suppliedRefs
            .map((ref) => String(ref || "").trim().toUpperCase())
            .filter((ref) => !normalizeRefs([ref], validIds).length);
          if (invalidRefs.length) {
            warnings.push(
              `${field} on slide ${index + 1} had invalid source references: ${invalidRefs.join(", ")}`
            );
          }
          return next;
        }).filter((item) => item.text);
      return {
        id: String(slide.id || uid()),
        kind,
        title: clampText(slide.title || `Slide ${index + 1}`, 90),
        subtitle: clampText(slide.subtitle || "", 140),
        bullets: refItems(slide.bullets, "bullet"),
        leftTitle: clampText(slide.leftTitle || "", 60),
        rightTitle: clampText(slide.rightTitle || "", 60),
        leftItems: refItems(slide.leftItems, "left item"),
        rightItems: refItems(slide.rightItems, "right item"),
        metrics: (Array.isArray(slide.metrics) ? slide.metrics : []).slice(0, 4).map((metric) => {
          const suppliedRefs = Array.isArray(metric?.sourceRefs) ? metric.sourceRefs : [];
          const invalidRefs = suppliedRefs
            .map((ref) => String(ref || "").trim().toUpperCase())
            .filter((ref) => !normalizeRefs([ref], validIds).length);
          if (invalidRefs.length) {
            warnings.push(
              `metric on slide ${index + 1} had invalid source references: ${invalidRefs.join(", ")}`
            );
          }
          return {
            value: clampText(metric?.value || "", 24),
            label: clampText(metric?.label || "", 70),
            sourceRefs: normalizeRefs(suppliedRefs, validIds),
          };
        }),
        notes: clampText(slide.notes || "", 700),
      };
    }),
    updatedAt: new Date().toISOString(),
  });
}

export function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

export function isDeck(value) {
  return Boolean(
    value &&
      typeof value === "object" &&
      typeof value.title === "string" &&
      Array.isArray(value.slides) &&
      value.slides.length > 0
  );
}

export function collectRefs(slide) {
  const refs = [];
  const addItems = (items) => {
    for (const item of items || []) {
      for (const ref of item.sourceRefs || []) {
        if (!refs.includes(ref)) refs.push(ref);
      }
    }
  };
  addItems(slide.bullets);
  addItems(slide.leftItems);
  addItems(slide.rightItems);
  addItems(slide.metrics);
  return refs;
}
