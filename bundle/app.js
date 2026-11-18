import { FIXTURE_DECK, SAMPLE_BRIEF } from "./fixture.js";
import {
  clone,
  createDeck,
  createSlide,
  isDeck,
  normalizeDeck,
  numberSource,
} from "./model.js";
import {
  formatGenerationError,
  generateDeckWithAnna,
  modelPreferenceSatisfied,
  planPptxEdit,
} from "./ai.js";
import {
  createPptxArrayBuffer,
  deliverPptxBuffer,
  safeFilename,
} from "./export.js";

const $ = (selector, root = document) => root.querySelector(selector);
const STORAGE_KEY = "proofdeck:current-project:v2";
const DB_NAME = "proofdeck-studio";
const DB_STORE = "projects";

const state = {
  anna: null,
  deck: clone(FIXTURE_DECK),
  prompt: SAMPLE_BRIEF.topic,
  activeSlideIndex: 0,
  renderedSlides: [],
  pptxBytes: null,
  templateBytes: null,
  templateInfo: null,
  sources: [],
  selectedShape: null,
  generated: false,
  busy: false,
  saveTimer: null,
  sampleBytes: null,
};

let pptxLoader = null;
let officeEnginePromise = null;
let documentEnginePromise = null;
let projectDbPromise = null;

function loadPptxGenJS() {
  if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
  if (pptxLoader) return pptxLoader;
  pptxLoader = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "./vendor/pptxgen.bundle.js";
    script.async = true;
    script.addEventListener("load", () => {
      if (window.PptxGenJS) resolve(window.PptxGenJS);
      else reject(new Error("PPTX 生成器加载失败，请刷新重试。"));
    }, { once: true });
    script.addEventListener("error", () => reject(new Error("PPTX 生成器加载失败，请检查网络后重试。")), { once: true });
    document.head.appendChild(script);
  });
  return pptxLoader;
}

function loadOfficeEngine() {
  if (!officeEnginePromise) officeEnginePromise = import("./vendor/office-engine.js");
  return officeEnginePromise;
}

function loadDocumentEngine() {
  if (!documentEnginePromise) {
    documentEnginePromise = import("./vendor/document-engine.js").then((engine) => {
      engine.configurePdfWorker(new URL("./vendor/pdf.worker.mjs", window.location.href).href);
      return engine;
    });
  }
  return documentEnginePromise;
}

function setText(node, value) {
  if (node) node.textContent = value;
}

function setStatus(text, tone = "") {
  const node = $("#generation-status");
  if (!node) return;
  node.textContent = text;
  node.dataset.tone = tone;
}

function setExportStatus(text, tone = "") {
  const node = $("#export-status");
  if (!node) return;
  node.textContent = text;
  node.dataset.tone = tone;
}

function appendMessage(author, text, tone = "") {
  const log = $("#conversation");
  const article = document.createElement("article");
  article.className = `message ${author === "You" ? "user" : "assistant"} ${tone}`.trim();
  const label = document.createElement("span");
  label.className = "message-author";
  label.textContent = author;
  const body = document.createElement("p");
  body.textContent = text;
  article.append(label, body);
  log.appendChild(article);
  log.scrollTop = log.scrollHeight;
}

function sourceText() {
  return state.sources.map((doc) => `[${doc.name}]\n${doc.text}`).join("\n\n");
}

function inferSlideCount(prompt) {
  const selected = $("#brief-slides").value;
  if (selected !== "auto") return Number(selected);
  const match = String(prompt).match(/(\d{1,2})\s*(?:页|张|slides?|pages?)/i);
  const value = match ? Number(match[1]) : 8;
  return Math.min(10, Math.max(6, value));
}

function inferLanguage(prompt) {
  const selected = $("#brief-language").value;
  if (selected !== "auto") return selected;
  const han = (String(prompt).match(/[\u3400-\u9fff]/g) || []).length;
  const latin = (String(prompt).match(/[a-z]/gi) || []).length;
  return han > latin ? "zh-CN" : "en";
}

function currentBrief() {
  const prompt = $("#chat-prompt").value.trim();
  return {
    topic: prompt,
    audience: "由提示词中的受众信息确定",
    purpose: "由提示词中的目标信息确定",
    language: inferLanguage(prompt),
    slideCount: inferSlideCount(prompt),
    tone: $("#brief-tone").value,
    modelHint: $("#brief-model").value,
    theme: $("#theme-select").value || "executive",
    sourceText: sourceText(),
  };
}

function setBusy(busy, label = "") {
  state.busy = busy;
  $("#generate").disabled = busy || !state.anna;
  $("#attach-template").disabled = busy;
  $("#attach-sources").disabled = busy;
  $("#apply-annotation").disabled = busy || !state.selectedShape || !state.generated || !state.anna;
  document.body.classList.toggle("busy", busy);
  if (label) setStatus(label);
}

function openProjectDb() {
  if (projectDbPromise) return projectDbPromise;
  projectDbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  return projectDbPromise;
}

async function idbGet(key) {
  const db = await openProjectDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(key);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function idbSet(key, value) {
  const db = await openProjectDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function svgElement(svgText) {
  const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
  if (doc.querySelector("parsererror")) throw new Error("幻灯片预览解析失败。");
  const svg = document.importNode(doc.documentElement, true);
  svg.removeAttribute("width");
  svg.removeAttribute("height");
  svg.setAttribute("role", "img");
  return svg;
}

function findShape(shapeId, slideIndex = state.activeSlideIndex) {
  return state.renderedSlides[slideIndex]?.shapes?.find((shape) => Number(shape.id) === Number(shapeId)) || null;
}

function updateAnnotationPanel() {
  const selected = state.selectedShape;
  setText($("#selected-shape-name"), selected ? (selected.name || `Element ${selected.id}`) : "Click any element in the slide");
  setText(
    $("#selected-shape-text"),
    selected
      ? (selected.text || `${selected.kind || "Shape"} · ${Math.round((selected.bounds?.w || 0) / 12700)} × ${Math.round((selected.bounds?.h || 0) / 12700)} pt`)
      : "Then describe the change for Anna."
  );
  $("#apply-annotation").disabled = !selected || !state.generated || state.busy || !state.anna;
  $("#annotation-input").disabled = !selected || !state.generated || state.busy;
}

function selectShape(shapeId, element = null) {
  const shape = findShape(shapeId);
  state.selectedShape = shape
    ? { ...shape, shapeId: Number(shape.id), slideIndex: state.activeSlideIndex }
    : null;
  const preview = $("#slide-preview");
  preview.querySelectorAll(".shape-selected").forEach((node) => node.classList.remove("shape-selected"));
  const target = element?.closest?.("[data-pptx-shape-id]")
    || preview.querySelector(`[data-pptx-shape-id="${Number(shapeId)}"]`);
  target?.classList.add("shape-selected");
  updateAnnotationPanel();
}

function renderSlidePreview() {
  const host = $("#slide-preview");
  const slide = state.renderedSlides[state.activeSlideIndex];
  if (!slide) {
    const empty = document.createElement("div");
    empty.className = "preview-empty";
    empty.textContent = "Upload a template or generate a deck to begin.";
    host.replaceChildren(empty);
    return;
  }
  try {
    const svg = svgElement(slide.svg);
    svg.setAttribute("aria-label", slide.title || `Slide ${state.activeSlideIndex + 1}`);
    host.replaceChildren(svg);
    if (state.selectedShape?.slideIndex === state.activeSlideIndex) {
      selectShape(state.selectedShape.shapeId);
    }
  } catch (error) {
    const empty = document.createElement("div");
    empty.className = "preview-empty";
    empty.textContent = error?.message || String(error);
    host.replaceChildren(empty);
  }
}

function renderFilmstrip() {
  const strip = $("#slide-filmstrip");
  strip.replaceChildren();
  state.renderedSlides.forEach((slide, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `filmstrip-item ${index === state.activeSlideIndex ? "active" : ""}`;
    button.setAttribute("aria-label", `Open slide ${index + 1}`);
    const image = document.createElement("img");
    image.alt = "";
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(slide.svg)}`;
    const number = document.createElement("span");
    number.textContent = String(index + 1).padStart(2, "0");
    button.append(image, number);
    button.addEventListener("click", () => {
      state.activeSlideIndex = index;
      state.selectedShape = null;
      renderFilmstrip();
      renderSlidePreview();
      updateSelectionLabel();
      updateAnnotationPanel();
    });
    strip.appendChild(button);
  });
}

function updateSelectionLabel() {
  setText($("#slide-position"), `${state.renderedSlides.length ? state.activeSlideIndex + 1 : 0} / ${state.renderedSlides.length}`);
  $("#previous-slide").disabled = state.activeSlideIndex <= 0;
  $("#next-slide").disabled = state.activeSlideIndex >= state.renderedSlides.length - 1;
}

function renderAttachments() {
  const host = $("#attachment-chips");
  host.replaceChildren();
  if (state.templateInfo) {
    const chip = document.createElement("div");
    chip.className = "attachment template-attachment";
    const icon = document.createElement("span");
    icon.className = "attachment-icon";
    icon.textContent = "T";
    const copy = document.createElement("span");
    copy.textContent = state.templateInfo.sourceName;
    const detail = document.createElement("small");
    detail.textContent = `${state.templateInfo.slideCount} slides · ${state.templateInfo.layouts.length} layouts`;
    copy.appendChild(detail);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "attachment-remove";
    remove.textContent = "×";
    remove.setAttribute("aria-label", "Remove template");
    remove.addEventListener("click", removeTemplate);
    chip.append(icon, copy, remove);
    host.appendChild(chip);
  }
  state.sources.forEach((doc, index) => {
    const chip = document.createElement("div");
    chip.className = "attachment";
    const icon = document.createElement("span");
    icon.className = "attachment-icon source";
    icon.textContent = doc.name.split(".").pop().slice(0, 1).toUpperCase();
    const copy = document.createElement("span");
    copy.textContent = doc.name;
    const detail = document.createElement("small");
    detail.textContent = `${doc.chars.toLocaleString()} chars`;
    copy.appendChild(detail);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "attachment-remove";
    remove.textContent = "×";
    remove.setAttribute("aria-label", `Remove ${doc.name}`);
    remove.addEventListener("click", () => {
      state.sources.splice(index, 1);
      renderAttachments();
      scheduleSave();
    });
    chip.append(icon, copy, remove);
    host.appendChild(chip);
  });
  setText($("#project-label"), state.deck.title || "Untitled project");
  setText($("#deck-title-label"), state.deck.title || "Untitled deck");
  $("#mode-badge").textContent = state.anna ? (state.templateInfo ? "Template mode" : "Anna connected") : "Local studio";
}

function renderTemplateSummary() {
  const box = $("#template-summary");
  if (!state.templateInfo) {
    box.hidden = true;
    return;
  }
  box.hidden = false;
  setText($("#template-name"), state.templateInfo.sourceName);
  setText(
    $("#template-detail"),
    `${state.templateInfo.slideCount} 个模板页 · ${state.templateInfo.layouts.length} 个版式 · 母版与组件已载入`
  );
  $("#template-swatch").style.background = "linear-gradient(135deg, #16264A, #E0A458)";
  $("#theme-select").disabled = true;
  $("#theme-select").title = "This project uses the uploaded template theme.";
}

function renderAll() {
  renderAttachments();
  renderTemplateSummary();
  renderFilmstrip();
  renderSlidePreview();
  updateSelectionLabel();
  updateAnnotationPanel();
  $("#export").disabled = !state.pptxBytes || !state.generated;
  $("#theme-select").disabled = Boolean(state.templateInfo);
}

async function renderBytes(bytes, generated = state.generated) {
  const engine = await loadOfficeEngine();
  state.pptxBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  state.renderedSlides = await engine.renderPresentation(state.pptxBytes);
  state.generated = generated;
  state.activeSlideIndex = Math.min(state.activeSlideIndex, Math.max(0, state.renderedSlides.length - 1));
  state.selectedShape = null;
  renderAll();
}

async function regenerateThemePreview() {
  if (!state.deck?.slides?.length || state.templateInfo) return;
  const Pptx = await loadPptxGenJS();
  const buffer = await createPptxArrayBuffer(state.deck, Pptx);
  await renderBytes(buffer, state.generated);
  scheduleSave();
}

async function getSampleBytes() {
  if (state.sampleBytes) return state.sampleBytes;
  const response = await fetch("./assets/sample.pptx");
  if (!response.ok) throw new Error("示例文件加载失败。");
  state.sampleBytes = new Uint8Array(await response.arrayBuffer());
  return state.sampleBytes;
}

async function loadExample(announce = true) {
  state.deck = clone(FIXTURE_DECK);
  state.prompt = SAMPLE_BRIEF.topic;
  state.sources = [{ name: "Q3 华东区经营复盘.txt", text: SAMPLE_BRIEF.sourceText, chars: SAMPLE_BRIEF.sourceText.length }];
  state.templateBytes = null;
  state.templateInfo = null;
  state.activeSlideIndex = 0;
  state.generated = true;
  $("#chat-prompt").value = state.prompt;
  $("#theme-select").value = state.deck.theme;
  await renderBytes(await getSampleBytes(), true);
  if (announce) appendMessage("Anna", "已载入一个可直接浏览和修改的示例。点击页面元素即可继续批注。");
  scheduleSave();
}

function newProject() {
  state.deck = createDeck({
    title: "Untitled deck",
    slides: [
      { ...createSlide("cover"), title: "New presentation" },
      { ...createSlide("bullets"), title: "Key message" },
      { ...createSlide("closing"), title: "Decision" },
    ],
  });
  state.prompt = "";
  state.sources = [];
  state.templateBytes = null;
  state.templateInfo = null;
  state.renderedSlides = [];
  state.pptxBytes = null;
  state.selectedShape = null;
  state.activeSlideIndex = 0;
  state.generated = false;
  $("#chat-prompt").value = "";
  $("#conversation").replaceChildren();
  appendMessage("Anna", "New project started. Upload a template or describe the deck you want to create.");
  renderAll();
  scheduleSave();
}

async function importTemplate(file) {
  if (!file) return;
  if (!/\.pptx$/i.test(file.name)) {
    setStatus("模板必须是 .pptx 文件。", "error");
    return;
  }
  if (file.size > 20 * 1024 * 1024) {
    setStatus("模板文件不能超过 20 MB。", "error");
    return;
  }
  setBusy(true, "正在读取模板结构、母版和组件…");
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const engine = await loadOfficeEngine();
    const info = await engine.inspectTemplate(bytes);
    if (!info.slideCount) throw new Error("模板中没有可用页面。");
    state.templateBytes = bytes;
    state.templateInfo = { ...info, sourceName: file.name };
    state.deck.templateName = file.name;
    state.activeSlideIndex = 0;
    await renderBytes(bytes, false);
    appendMessage("Anna", `已读取模板「${file.name}」：${info.slideCount} 页、${info.layouts.length} 个版式。生成时会复用原模板页面和组件。`);
    setStatus("模板已载入。描述要生成的内容，然后点击 Generate deck。", "success");
  } catch (error) {
    setStatus(error?.message || String(error), "error");
  } finally {
    setBusy(false);
    $("#template-file").value = "";
  }
}

function removeTemplate() {
  state.templateBytes = null;
  state.templateInfo = null;
  state.generated = false;
  state.renderedSlides = [];
  state.pptxBytes = null;
  state.selectedShape = null;
  state.deck.templateName = "";
  renderAll();
  scheduleSave();
  setStatus("已移除模板。下一次生成会使用 ProofDeck 内置版式。");
}

async function importSources(files) {
  if (!files?.length) return;
  setBusy(true, "正在提取资料文本…");
  try {
    const engine = await loadDocumentEngine();
    let total = state.sources.reduce((sum, doc) => sum + doc.text.length, 0);
    for (const file of files) {
      if (file.size > 15 * 1024 * 1024) {
        appendMessage("Anna", `跳过「${file.name}」：文件超过 15 MB。`, "error");
        continue;
      }
      const text = await engine.extractDocumentText(file.name, file.type, new Uint8Array(await file.arrayBuffer()));
      const available = Math.max(0, 60000 - total);
      const clipped = text.slice(0, Math.min(20000, available)).trim();
      if (!clipped) {
        appendMessage("Anna", `「${file.name}」没有提取到可用文本。`, "error");
        continue;
      }
      state.sources.push({ name: file.name, text: clipped, chars: clipped.length });
      total += clipped.length;
      appendMessage("Anna", `已读取「${file.name}」，提取 ${clipped.length.toLocaleString()} 个字符。`);
    }
    renderAttachments();
    scheduleSave();
    setStatus("资料已加入。直接描述你要生成的演示文稿。", "success");
  } catch (error) {
    setStatus(`资料读取失败：${error?.message || String(error)}`, "error");
  } finally {
    setBusy(false);
    $("#source-files").value = "";
  }
}

async function generate() {
  const prompt = $("#chat-prompt").value.trim();
  if (!prompt) {
    setStatus("先用一句话说明你要做什么。", "error");
    $("#chat-prompt").focus();
    return;
  }
  if (!state.anna) {
    setStatus("当前是本地预览模式；请从 Anna 平台打开 App 后生成。", "error");
    return;
  }
  appendMessage("You", prompt);
  setBusy(true, "Anna 正在规划内容并填充模板…");
  try {
    const brief = currentBrief();
    const sourceUnits = numberSource(brief.sourceText);
    const result = await generateDeckWithAnna(state.anna, brief, sourceUnits);
    if (!modelPreferenceSatisfied(brief.modelHint, result.model)) {
      result.warnings.push(`模型偏好 ${brief.modelHint} 被账号默认模型 ${result.model || "未知"} 覆盖`);
    }
    result.deck.theme = brief.theme;
    result.deck.sourceText = brief.sourceText;
    result.deck.language = brief.language;
    state.deck = result.deck;
    state.activeSlideIndex = 0;
    state.selectedShape = null;

    let buffer;
    if (state.templateBytes) {
      const engine = await loadOfficeEngine();
      const built = await engine.buildFromTemplate(state.templateBytes, state.deck);
      buffer = built.bytes;
    } else {
      const Pptx = await loadPptxGenJS();
      buffer = await createPptxArrayBuffer(state.deck, Pptx);
    }
    await renderBytes(buffer, true);
    const quota = Number(result.quotaConsumed || 0);
    setText($("#quota-badge"), quota ? `Quota ${quota.toFixed(3)}` : `Model ${result.model || "Anna AI"}`);
    $("#quota-badge").hidden = false;
    const warning = result.warnings.length ? ` ${result.warnings.join("；")}` : "";
    appendMessage(
      "Anna",
      `已生成 ${result.deck.slides.length} 页。模型：${result.model || "Anna AI"}。${warning} 点击任意页面元素即可继续批注修改。`,
      result.warnings.length ? "warning" : ""
    );
    setStatus(
      result.warnings.length ? `已生成，需注意：${result.warnings.join("；")}` : "生成完成。点击页面元素继续修改。",
      result.warnings.length ? "warning" : "success"
    );
    scheduleSave();
  } catch (error) {
    const message = formatGenerationError(error);
    appendMessage("Anna", message, "error");
    setStatus(message, "error");
  } finally {
    setBusy(false);
  }
}

async function applyAnnotation() {
  const instruction = $("#annotation-input").value.trim();
  if (!state.anna) {
    setStatus("当前没有连接 Anna AI。", "error");
    return;
  }
  if (!state.selectedShape) {
    setStatus("请先在幻灯片中选中一个元素。", "error");
    return;
  }
  if (!instruction) {
    setStatus("请输入希望如何修改这个元素。", "error");
    return;
  }
  const target = { ...state.selectedShape };
  const activeSlide = state.renderedSlides[state.activeSlideIndex];
  const sourceUnits = state.deck.sourceUnits || numberSource(sourceText());
  const refs = [...String(target.text || "").matchAll(/\bS\d+\b/g)].map((match) => match[0]);
  const relevantSources = sourceUnits.filter((unit) => refs.includes(unit.id)).slice(0, 8);
  setBusy(true, "Anna 正在根据批注修改选中元素…");
  try {
    const planned = await planPptxEdit(state.anna, {
      instruction,
      deckTitle: state.deck.title,
      language: state.deck.language,
      slide: {
        index: state.activeSlideIndex,
        title: activeSlide?.title || state.deck.slides[state.activeSlideIndex]?.title || "",
        notes: activeSlide?.notes || state.deck.slides[state.activeSlideIndex]?.notes || "",
      },
      target,
      sources: relevantSources,
      template: state.templateInfo
        ? { name: state.templateInfo.sourceName, layouts: state.templateInfo.layouts }
        : null,
    });
    const engine = await loadOfficeEngine();
    const patched = await engine.applyPptxPatch(state.pptxBytes, planned.operations);
    state.pptxBytes = patched.bytes;
    state.renderedSlides = patched.slides;
    state.generated = true;
    $("#annotation-input").value = "";
    renderFilmstrip();
    renderSlidePreview();
    selectShape(target.shapeId);
    setStatus(planned.reply, "success");
    appendMessage("Anna", planned.reply);
    const quota = Number(planned.quotaConsumed || 0);
    if (quota) {
      setText($("#quota-badge"), `Quota ${quota.toFixed(3)}`);
      $("#quota-badge").hidden = false;
    }
    scheduleSave();
  } catch (error) {
    const message = formatGenerationError(error);
    setStatus(message, "error");
  } finally {
    setBusy(false);
  }
}

async function exportDeck() {
  if (!state.pptxBytes) {
    setStatus("还没有可以导出的演示文稿。", "error");
    return;
  }
  $("#export").disabled = true;
  setExportStatus("正在准备 PPTX…");
  try {
    const result = await deliverPptxBuffer({
      anna: state.anna,
      buffer: state.pptxBytes,
      filename: safeFilename(state.deck.title),
      deckId: state.deck.id,
    });
    setExportStatus(result.mode === "anna" ? "PPTX 已交给 Anna 下载。" : "PPTX 已开始下载。", "success");
  } catch (error) {
    setExportStatus(error?.message || String(error), "error");
  } finally {
    $("#export").disabled = false;
  }
}

async function saveDraft() {
  const record = {
    deck: clone(state.deck),
    prompt: $("#chat-prompt").value,
    sources: state.sources,
    templateInfo: state.templateInfo,
    templateBytes: state.templateBytes,
    pptxBytes: state.pptxBytes,
    generated: state.generated,
    savedAt: new Date().toISOString(),
  };
  try {
    await idbSet("current", record);
  } catch {
    // The app remains usable when IndexedDB is unavailable.
  }
  if (!state.anna?.storage?.set) return;
  try {
    await state.anna.storage.set({
      key: STORAGE_KEY,
      value: record.deck,
      metadata: { updatedAt: record.deck.updatedAt || record.savedAt },
    });
  } catch {
    // The local copy remains available.
  }
}

function scheduleSave() {
  clearTimeout(state.saveTimer);
  state.saveTimer = setTimeout(saveDraft, 600);
}

async function restoreDraft() {
  try {
    const record = await idbGet("current");
    if (record?.deck && isDeck(record.deck)) {
      state.deck = normalizeDeck(record.deck);
      state.prompt = String(record.prompt || "");
      state.sources = Array.isArray(record.sources) ? record.sources : [];
      state.templateInfo = record.templateInfo || null;
      state.templateBytes = record.templateBytes ? new Uint8Array(record.templateBytes) : null;
      state.pptxBytes = record.pptxBytes ? new Uint8Array(record.pptxBytes) : null;
      state.generated = Boolean(record.generated);
      $("#chat-prompt").value = state.prompt;
      if (state.pptxBytes) {
        await renderBytes(state.pptxBytes, state.generated);
      } else if (state.generated && state.deck?.slides?.length) {
        const engine = await loadOfficeEngine();
        const built = state.templateBytes
          ? await engine.buildFromTemplate(state.templateBytes, state.deck)
          : null;
        const buffer = built?.bytes || await createPptxArrayBuffer(state.deck, await loadPptxGenJS());
        await renderBytes(buffer, true);
      } else if (state.templateBytes) {
        await renderBytes(state.templateBytes, false);
      }
      return true;
    }
  } catch {
    // Fall through to the Anna draft or demo.
  }

  if (state.anna?.storage?.get) {
    try {
      const result = await state.anna.storage.get({ key: STORAGE_KEY });
      if (isDeck(result?.value)) {
        state.deck = normalizeDeck(result.value);
        const buffer = await createPptxArrayBuffer(state.deck, await loadPptxGenJS());
        await renderBytes(buffer, true);
        return true;
      }
    } catch {
      // A missing Anna draft is expected.
    }
  }
  return false;
}

function applyEntryPayload(payload) {
  if (!payload || typeof payload !== "object") return;
  const prompt = String(payload.topic || payload.brief || "").trim();
  if (prompt) {
    state.prompt = prompt;
    $("#chat-prompt").value = prompt;
  }
  const incomingSource = String(payload.sourceText || payload.source || "").trim();
  if (incomingSource) {
    state.sources = [{ name: "Anna chat context.txt", text: incomingSource, chars: incomingSource.length }];
  }
  if (payload.slides || payload.slideCount) $("#brief-slides").value = String(payload.slides || payload.slideCount);
  if (payload.language) $("#brief-language").value = payload.language === "en" ? "en" : "zh-CN";
  setStatus("Brief received from Anna. Review it, then generate.");
  renderAttachments();
}

async function connectAnna() {
  try {
    const { AnnaAppRuntime } = await import("/static/anna-apps/_sdk/latest/index.js");
    state.anna = await AnnaAppRuntime.connect();
    await state.anna.window.set_title({ title: "ProofDeck" }).catch(() => {});
    state.anna.on?.("entry_payload", applyEntryPayload);
    if (state.anna.entryPayload && Object.keys(state.anna.entryPayload).length) {
      applyEntryPayload(state.anna.entryPayload);
    }
    setStatus("Anna connected. Generation and edits use your current quota.");
  } catch (error) {
    console.warn("[ProofDeck] Anna runtime unavailable.", error);
    state.anna = null;
    setStatus("Local studio: template import, preview, editing, and export work; AI generation requires Anna.");
  }
  renderAttachments();
}

function bind() {
  $("#generate").addEventListener("click", generate);
  $("#export").addEventListener("click", exportDeck);
  $("#new-project").addEventListener("click", newProject);
  $("#load-example").addEventListener("click", () => loadExample(true));
  $("#attach-template").addEventListener("click", () => $("#template-file").click());
  $("#attach-sources").addEventListener("click", () => $("#source-files").click());
  $("#template-file").addEventListener("change", (event) => importTemplate(event.target.files?.[0]));
  $("#source-files").addEventListener("change", (event) => importSources([...event.target.files]));
  $("#previous-slide").addEventListener("click", () => {
    state.activeSlideIndex = Math.max(0, state.activeSlideIndex - 1);
    state.selectedShape = null;
    renderFilmstrip();
    renderSlidePreview();
    updateSelectionLabel();
    updateAnnotationPanel();
  });
  $("#next-slide").addEventListener("click", () => {
    state.activeSlideIndex = Math.min(state.renderedSlides.length - 1, state.activeSlideIndex + 1);
    state.selectedShape = null;
    renderFilmstrip();
    renderSlidePreview();
    updateSelectionLabel();
    updateAnnotationPanel();
  });
  $("#present").addEventListener("click", () => {
    if (!document.fullscreenElement) $("#viewer-panel").requestFullscreen?.();
    else document.exitFullscreen?.();
  });
  $("#apply-annotation").addEventListener("click", applyAnnotation);
  document.querySelectorAll("[data-annotation]").forEach((button) => {
    button.addEventListener("click", () => {
      $("#annotation-input").value = button.dataset.annotation || "";
      applyAnnotation();
    });
  });
  $("#annotation-input").addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === "Enter") applyAnnotation();
  });
  $("#slide-preview").addEventListener("click", (event) => {
    if (!state.generated) return;
    const target = event.target.closest?.("[data-pptx-shape-id]");
    if (!target) return;
    const shapeId = Number(target.getAttribute("data-pptx-shape-id"));
    if (Number.isFinite(shapeId)) selectShape(shapeId, target);
  });
  $("#theme-select").addEventListener("change", async () => {
    if (state.templateInfo) return;
    state.deck.theme = $("#theme-select").value;
    await regenerateThemePreview();
  });
}

async function main() {
  bind();
  try {
    await loadExample(false);
  } catch (error) {
    setStatus(`示例加载失败：${error?.message || String(error)}`, "error");
  }
  await connectAnna();
  const restored = await restoreDraft();
  if (restored) renderAll();
  setBusy(false);
  renderAll();
}

main();
