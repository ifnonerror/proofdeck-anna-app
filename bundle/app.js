import {
  THEMES,
  createDeck,
  createSlide,
  clone,
  isDeck,
  numberSource,
} from "./model.js";
import { FIXTURE_DECK, SAMPLE_BRIEF, deckFromPayload } from "./fixture.js";
import {
  formatGenerationError,
  generateDeckWithAnna,
  modelPreferenceSatisfied,
} from "./ai.js";
import { renderSlidePreview } from "./preview.js";
import { exportPresentation } from "./export.js";

const STORAGE_KEY = "proofdeck:current-project";
const $ = (selector, root = document) => root.querySelector(selector);

const state = {
  anna: null,
  deck: clone(FIXTURE_DECK),
  activeSlideId: FIXTURE_DECK.slides[0].id,
  busy: false,
  saveTimer: null,
  lastPayload: null,
};

let pptxLoader = null;

function loadPptxGenJS() {
  if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
  if (pptxLoader) return pptxLoader;
  pptxLoader = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "./vendor/pptxgen.bundle.js";
    script.async = true;
    script.addEventListener("load", () => {
      if (window.PptxGenJS) resolve(window.PptxGenJS);
      else reject(new Error("PptxGenJS 资源加载后不可用，请刷新重试。"));
    }, { once: true });
    script.addEventListener("error", () => {
      reject(new Error("PptxGenJS 资源加载失败，请检查网络后重试。"));
    }, { once: true });
    document.head.appendChild(script);
  });
  return pptxLoader;
}

function activeSlide() {
  return state.deck.slides.find((slide) => slide.id === state.activeSlideId)
    || state.deck.slides[0]
    || null;
}

function setText(node, text) {
  if (node) node.textContent = text;
}

function setStatus(text, tone = "") {
  const node = $("#generation-status");
  if (!node) return;
  node.textContent = text;
  node.dataset.tone = tone;
}

function setSaveStatus(text, tone = "") {
  const node = $("#save-status");
  if (!node) return;
  node.textContent = text;
  node.dataset.tone = tone;
}

function currentBrief() {
  return {
    topic: $("#brief-topic").value.trim(),
    audience: $("#brief-audience").value.trim(),
    purpose: $("#brief-purpose").value.trim(),
    language: $("#brief-language").value,
    slideCount: Number($("#brief-slides").value),
    tone: $("#brief-tone").value,
    modelHint: $("#brief-model").value,
    theme: $("#theme-select").value,
    sourceText: $("#source-text").value.trim(),
  };
}

function fillBrief(brief) {
  $("#brief-topic").value = brief.topic || "";
  $("#brief-audience").value = brief.audience || "";
  $("#brief-purpose").value = brief.purpose || "";
  $("#brief-language").value = brief.language === "en" ? "en" : "zh-CN";
  $("#brief-slides").value = String(brief.slideCount || brief.slides || 8);
  $("#brief-tone").value = brief.tone || "executive";
  $("#brief-model").value = brief.modelHint || "auto";
  $("#source-text").value = brief.sourceText || "";
  $("#theme-select").value = THEMES.includes(brief.theme) ? brief.theme : "executive";
}

function button(label, className, onClick) {
  const node = document.createElement("button");
  node.type = "button";
  node.className = className;
  node.textContent = label;
  node.addEventListener("click", onClick);
  return node;
}

function input(value, onInput, placeholder = "") {
  const node = document.createElement("input");
  node.type = "text";
  node.value = value || "";
  node.placeholder = placeholder;
  node.addEventListener("input", () => onInput(node.value));
  return node;
}

function textarea(value, onInput, rows = 3, placeholder = "") {
  const node = document.createElement("textarea");
  node.rows = rows;
  node.value = value || "";
  node.placeholder = placeholder;
  node.addEventListener("input", () => onInput(node.value));
  return node;
}

function sourceChips(refs) {
  const wrap = document.createElement("span");
  wrap.className = "chip-row";
  for (const ref of refs || []) {
    const chip = document.createElement("span");
    chip.className = "source-chip";
    chip.textContent = ref;
    wrap.appendChild(chip);
  }
  return wrap;
}

function touchDeck({ rerender = false } = {}) {
  state.deck.updatedAt = new Date().toISOString();
  renderPreview();
  renderOutline();
  scheduleSave();
  if (rerender) renderSlideEditor();
}

function renderSources() {
  const box = $("#sources");
  box.replaceChildren();
  const units = numberSource($("#source-text").value);
  setText($("#source-count"), `${units.length} unit${units.length === 1 ? "" : "s"}`);
  if (!units.length) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "No source material. ProofDeck can still draft from the brief.";
    box.appendChild(empty);
    return;
  }
  for (const unit of units) {
    const row = document.createElement("article");
    row.className = "source-unit";
    const id = document.createElement("span");
    id.className = "source-id";
    id.textContent = unit.id;
    const text = document.createElement("p");
    text.textContent = unit.text;
    row.append(id, text);
    box.appendChild(row);
  }
}

function renderOutline() {
  const list = $("#outline");
  list.replaceChildren();
  state.deck.slides.forEach((slide, index) => {
    const item = document.createElement("li");
    item.className = slide.id === state.activeSlideId ? "active" : "";
    const select = button(`${String(index + 1).padStart(2, "0")} · ${slide.title || "Untitled"}`, "outline-select", () => {
      state.activeSlideId = slide.id;
      renderAll();
    });
    item.appendChild(select);
    const kind = document.createElement("span");
    kind.className = "kind-chip";
    kind.textContent = slide.kind;
    item.appendChild(kind);
    list.appendChild(item);
  });
}

function renderItemEditor(container, items, emptyLabel, onChange) {
  container.replaceChildren();
  if (!items.length) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = emptyLabel;
    container.appendChild(empty);
  }
  items.forEach((item, index) => {
    const row = document.createElement("div");
    row.className = "edit-item";
    const main = document.createElement("div");
    main.className = "edit-item-main";
    const textInput = input(item.text, (value) => {
      item.text = value;
      touchDeck();
    }, "Content");
    const refsInput = input((item.sourceRefs || []).join(" "), (value) => {
      const valid = new Set(state.deck.sourceUnits.map((unit) => unit.id));
      item.sourceRefs = [...new Set(value.toUpperCase().split(/[\s,，、]+/)
        .filter((ref) => /^S\d+$/.test(ref) && valid.has(ref)))];
      touchDeck();
    }, "S1 S2");
    main.append(textInput, refsInput);
    row.appendChild(main);
    row.appendChild(button("×", "icon-button danger", () => {
      items.splice(index, 1);
      touchDeck({ rerender: true });
    }));
    container.appendChild(row);
  });
}

function renderSlideEditor() {
  const host = $("#slide-editor");
  host.replaceChildren();
  const slide = activeSlide();
  if (!slide) return;

  const heading = document.createElement("div");
  heading.className = "editor-heading";
  const label = document.createElement("h3");
  label.textContent = `Slide ${state.deck.slides.indexOf(slide) + 1}`;
  heading.appendChild(label);
  const actions = document.createElement("div");
  actions.className = "editor-actions";
  actions.appendChild(button("Move up", "small-button", () => moveSlide(-1)));
  actions.appendChild(button("Move down", "small-button", () => moveSlide(1)));
  actions.appendChild(button("Delete", "small-button danger", deleteSlide));
  heading.appendChild(actions);
  host.appendChild(heading);

  const fields = document.createElement("div");
  fields.className = "field-grid";
  const titleLabel = document.createElement("label");
  titleLabel.textContent = "Assertion title";
  titleLabel.appendChild(input(slide.title, (value) => {
    slide.title = value;
    touchDeck();
  }));
  const kindLabel = document.createElement("label");
  kindLabel.textContent = "Layout";
  const kindSelect = document.createElement("select");
  for (const kind of ["cover", "agenda", "section", "bullets", "comparison", "timeline", "metrics", "closing"]) {
    const option = document.createElement("option");
    option.value = kind;
    option.textContent = kind;
    option.selected = kind === slide.kind;
    kindSelect.appendChild(option);
  }
  kindSelect.addEventListener("change", () => {
    slide.kind = kindSelect.value;
    touchDeck({ rerender: true });
  });
  kindLabel.appendChild(kindSelect);
  fields.append(titleLabel, kindLabel);
  host.appendChild(fields);

  if (slide.kind === "cover" || slide.kind === "closing" || slide.kind === "section") {
    const subtitleLabel = document.createElement("label");
    subtitleLabel.textContent = "Subtitle";
    subtitleLabel.appendChild(textarea(slide.subtitle, (value) => {
      slide.subtitle = value;
      touchDeck();
    }, 2));
    host.appendChild(subtitleLabel);
  }

  if (slide.kind === "comparison") {
    const columns = document.createElement("div");
    columns.className = "comparison-editors";
    for (const side of ["left", "right"]) {
      const section = document.createElement("section");
      const title = document.createElement("label");
      title.textContent = `${side === "left" ? "Left" : "Right"} heading`;
      title.appendChild(input(slide[`${side}Title`], (value) => {
        slide[`${side}Title`] = value;
        touchDeck();
      }));
      section.appendChild(title);
      const items = document.createElement("div");
      renderItemEditor(items, slide[`${side}Items`], "No items", () => {});
      section.appendChild(items);
      section.appendChild(button("Add item", "small-button", () => {
        slide[`${side}Items`].push({ text: "", sourceRefs: [] });
        touchDeck({ rerender: true });
      }));
      columns.appendChild(section);
    }
    host.appendChild(columns);
  } else if (slide.kind === "metrics") {
    const metricList = document.createElement("div");
    metricList.className = "metrics-editor";
    slide.metrics.forEach((metric, index) => {
      const row = document.createElement("div");
      row.className = "metric-edit";
      row.appendChild(input(metric.value, (value) => {
        metric.value = value;
        touchDeck();
      }, "Value"));
      row.appendChild(input(metric.label, (value) => {
        metric.label = value;
        touchDeck();
      }, "Label"));
      row.appendChild(button("×", "icon-button danger", () => {
        slide.metrics.splice(index, 1);
        touchDeck({ rerender: true });
      }));
      metricList.appendChild(row);
    });
    host.appendChild(metricList);
    host.appendChild(button("Add metric", "small-button", () => {
      slide.metrics.push({ value: "", label: "", sourceRefs: [] });
      touchDeck({ rerender: true });
    }));
  } else if (slide.kind !== "cover" && slide.kind !== "closing") {
    const bulletList = document.createElement("div");
    bulletList.className = "bullet-editor";
    renderItemEditor(bulletList, slide.bullets, "No bullets", () => {});
    host.appendChild(bulletList);
    host.appendChild(button("Add key point", "small-button", () => {
      slide.bullets.push({ text: "", sourceRefs: [] });
      touchDeck({ rerender: true });
    }));
  }

  const notes = document.createElement("label");
  notes.className = "notes-field";
  notes.textContent = "Speaker notes";
  notes.appendChild(textarea(slide.notes, (value) => {
    slide.notes = value;
    touchDeck();
  }, 4, "What should the presenter say on this slide?"));
  host.appendChild(notes);

  const refs = document.createElement("div");
  refs.className = "editor-source-line";
  refs.appendChild(sourceChips(Array.from(new Set(
    [...slide.bullets, ...slide.leftItems, ...slide.rightItems, ...slide.metrics]
      .flatMap((item) => item.sourceRefs || [])
  ))));
  host.appendChild(refs);
}

function renderPreview() {
  const slide = activeSlide();
  renderSlidePreview($("#slide-preview"), slide, state.deck);
  setText($("#slide-position"), `${Math.max(0, activeIndex()) + 1} / ${state.deck.slides.length}`);
  setText($("#preview-caption"), `${state.deck.theme} theme`);
}

function renderAll() {
  const slide = activeSlide();
  if (!slide && state.deck.slides.length) state.activeSlideId = state.deck.slides[0].id;
  setText($("#deck-title-label"), state.deck.title);
  setText($("#mode-badge"), state.anna ? "Anna connected" : "Fixture preview");
  $("#deck-title").value = state.deck.title || "";
  $("#deck-subtitle").value = state.deck.subtitle || "";
  $("#theme-select").value = state.deck.theme;
  renderSources();
  renderOutline();
  renderSlideEditor();
  renderPreview();
  $("#generate").disabled = state.busy || !state.anna;
}

function activeIndex() {
  return state.deck.slides.findIndex((slide) => slide.id === state.activeSlideId);
}

function moveSlide(direction) {
  const index = activeIndex();
  const next = index + direction;
  if (index < 0 || next < 0 || next >= state.deck.slides.length) return;
  const [slide] = state.deck.slides.splice(index, 1);
  state.deck.slides.splice(next, 0, slide);
  state.activeSlideId = slide.id;
  touchDeck({ rerender: true });
}

function deleteSlide() {
  const index = activeIndex();
  if (index < 0 || state.deck.slides.length <= 1) return;
  state.deck.slides.splice(index, 1);
  state.activeSlideId = state.deck.slides[Math.max(0, index - 1)].id;
  touchDeck({ rerender: true });
}

async function saveDraft() {
  if (!state.anna?.storage?.set) {
    setSaveStatus("Local only");
    return;
  }
  try {
    await state.anna.storage.set({
      key: STORAGE_KEY,
      value: clone(state.deck),
      metadata: { updatedAt: state.deck.updatedAt },
    });
    setSaveStatus("Saved");
  } catch (error) {
    setSaveStatus("Save failed", "error");
    setStatus(error?.message || String(error), "error");
  }
}

function scheduleSave() {
  setSaveStatus("Saving…");
  clearTimeout(state.saveTimer);
  state.saveTimer = setTimeout(saveDraft, 700);
}

async function restoreDraft() {
  if (!state.anna?.storage?.get) return;
  try {
    const result = await state.anna.storage.get({ key: STORAGE_KEY });
    if (isDeck(result?.value)) {
      state.deck = result.value;
      state.activeSlideId = state.deck.slides[0]?.id || null;
      setSaveStatus("Restored");
    }
  } catch {
    // A missing draft is expected for a new installation.
  }
}

function applyEntryPayload(payload) {
  if (!payload || typeof payload !== "object") return;
  state.lastPayload = payload;
  fillBrief({
    topic: payload.topic || payload.brief || "",
    audience: payload.audience || "",
    purpose: payload.purpose || "",
    language: payload.language || "zh-CN",
    slideCount: payload.slides || payload.slideCount || 8,
    tone: payload.tone || "executive",
    theme: payload.theme || state.deck.theme,
    sourceText: payload.sourceText || payload.source || "",
  });
  const example = deckFromPayload(payload).deck;
  state.deck = example;
  state.activeSlideId = example.slides[0]?.id || null;
  setStatus("Brief received from Anna. Review it, then click Generate.");
  renderAll();
  scheduleSave();
}

async function connectAnna() {
  try {
    const { AnnaAppRuntime } = await import("/static/anna-apps/_sdk/latest/index.js");
    state.anna = await AnnaAppRuntime.connect();
    await state.anna.window.set_title({ title: "ProofDeck" }).catch(() => {});
    await restoreDraft();
    if (state.anna.entryPayload && Object.keys(state.anna.entryPayload).length) {
      applyEntryPayload(state.anna.entryPayload);
    }
    state.anna.on?.("entry_payload", applyEntryPayload);
    setStatus("Anna connected. Generation will use your current quota.");
  } catch (error) {
    console.warn("[ProofDeck] Anna runtime unavailable; using standalone mode.", error);
    state.anna = null;
    setStatus("Standalone preview: editing and PPTX download work, but AI generation needs Anna.");
  }
}

async function generate() {
  if (!state.anna) {
    setStatus("未连接 Anna AI；请从平台安装并打开此 App 后再生成。", "error");
    return;
  }
  const brief = currentBrief();
  if (!brief.topic) {
    setStatus("请先填写主题或需求。", "error");
    $("#brief-topic").focus();
    return;
  }
  const sourceUnits = numberSource(brief.sourceText);
  state.busy = true;
  $("#generate").disabled = true;
  setStatus(`Anna AI 正在生成 ${brief.slideCount} 页演示…`);
  try {
    const result = await generateDeckWithAnna(state.anna, brief, sourceUnits);
    if (!modelPreferenceSatisfied(brief.modelHint, result.model)) {
      result.warnings.push(
        `模型偏好 ${brief.modelHint} 被账号默认模型 ${result.model || "未知"} 覆盖`
      );
    }
    result.deck.theme = brief.theme;
    result.deck.sourceText = brief.sourceText;
    result.deck.language = brief.language;
    state.deck = result.deck;
    state.activeSlideId = state.deck.slides[0]?.id || null;
    renderAll();
    const quota = Number(result.quotaConsumed || 0);
    setText($("#quota-badge"), quota ? `Quota ${quota.toFixed(3)}` : `Model ${result.model || "Anna AI"}`);
    $("#quota-badge").hidden = false;
    setStatus(
      result.warnings.length
        ? `已生成。注意：${result.warnings.join("；")}`
        : `已生成 ${result.deck.slides.length} 页。模型：${result.model || "Anna AI"}`,
      result.warnings.length ? "warning" : "success"
    );
    await saveDraft();
  } catch (error) {
    setStatus(formatGenerationError(error), "error");
  } finally {
    state.busy = false;
    $("#generate").disabled = !state.anna;
  }
}

function loadExample() {
  fillBrief(SAMPLE_BRIEF);
  state.deck = clone(FIXTURE_DECK);
  state.activeSlideId = state.deck.slides[0].id;
  $("#quota-badge").hidden = true;
  setStatus("已载入离线示例。编辑和导出不消耗模型额度。");
  renderAll();
  scheduleSave();
}

async function exportDeck() {
  const buttonNode = $("#export");
  buttonNode.disabled = true;
  setText($("#export-status"), "正在生成 PPTX…");
  try {
    const Pptx = await loadPptxGenJS();
    const result = await exportPresentation({ anna: state.anna, deck: state.deck, PptxGenJS: Pptx });
    setText($("#export-status"), result.mode === "anna" ? "已交给 Anna 下载。" : "PPTX 已开始下载。");
  } catch (error) {
    setText($("#export-status"), error?.message || String(error));
  } finally {
    buttonNode.disabled = false;
  }
}

function newProject() {
  const blank = createDeck({
    title: "Untitled deck",
    sourceUnits: [],
    slides: [
      createSlide("cover"),
      createSlide("bullets"),
      createSlide("closing"),
    ],
  });
  state.deck = blank;
  state.activeSlideId = blank.slides[0].id;
  fillBrief({
    topic: "",
    audience: "",
    purpose: "",
    language: "zh-CN",
    slideCount: 8,
    tone: "executive",
    modelHint: "auto",
    theme: "executive",
    sourceText: "",
  });
  setStatus("New project created.");
  renderAll();
  scheduleSave();
}

function bind() {
  $("#new-project").addEventListener("click", newProject);
  $("#load-example").addEventListener("click", loadExample);
  $("#generate").addEventListener("click", generate);
  $("#export").addEventListener("click", exportDeck);
  $("#source-text").addEventListener("input", renderSources);
  $("#theme-select").addEventListener("change", () => {
    state.deck.theme = $("#theme-select").value;
    touchDeck();
  });
  $("#deck-title").addEventListener("input", (event) => {
    state.deck.title = event.target.value;
    setText($("#deck-title-label"), state.deck.title);
    touchDeck();
  });
  $("#deck-subtitle").addEventListener("input", (event) => {
    state.deck.subtitle = event.target.value;
    touchDeck();
  });
  $("#add-slide").addEventListener("click", () => {
    const slide = createSlide($("#new-slide-kind").value);
    const index = activeIndex();
    state.deck.slides.splice(index + 1, 0, slide);
    state.activeSlideId = slide.id;
    touchDeck({ rerender: true });
  });
  $("#tab-edit").addEventListener("click", () => setMobileTab("edit"));
  $("#tab-preview").addEventListener("click", () => setMobileTab("preview"));
  window.addEventListener("resize", renderPreview);
}

function setMobileTab(tab) {
  document.body.dataset.mobileTab = tab;
  $("#tab-edit").classList.toggle("active", tab === "edit");
  $("#tab-preview").classList.toggle("active", tab === "preview");
}

async function main() {
  bind();
  fillBrief(SAMPLE_BRIEF);
  renderAll();
  setStatus("正在连接 Anna…");
  await connectAnna();
  if (!state.anna) {
    state.deck = clone(FIXTURE_DECK);
    state.activeSlideId = state.deck.slides[0].id;
  }
  renderAll();
}

main();
