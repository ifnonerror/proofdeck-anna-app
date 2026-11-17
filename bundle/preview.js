import { collectRefs } from "./model.js";
import { layoutSlide, toCssRegion } from "./layout.js";

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function addTextWithRefs(parent, text, refs) {
  parent.append(document.createTextNode(text || ""));
  for (const ref of refs || []) {
    const chip = el("span", "source-ref", ref);
    parent.appendChild(chip);
  }
}

function regionDiv(region, scale, className) {
  const div = el("div", `slide-region ${className || ""}`);
  Object.assign(div.style, toCssRegion(region, scale));
  return div;
}

export function renderSlidePreview(container, slide, deck) {
  if (!slide) {
    container.replaceChildren(el("div", "empty-preview", "No slide selected"));
    return;
  }
  const width = container.clientWidth || 640;
  const scale = width / 1280;
  const layout = layoutSlide(slide, deck.theme);
  const t = layout.tokens;
  const root = el("div", `slide-canvas theme-${layout.themeName}`);
  root.style.setProperty("--slide-bg", `#${t.background}`);
  root.style.setProperty("--slide-fg", `#${t.foreground}`);
  root.style.setProperty("--slide-muted", `#${t.muted}`);
  root.style.setProperty("--slide-accent", `#${t.accent}`);
  root.style.setProperty("--slide-card", `#${t.card}`);
  root.style.setProperty("--slide-scale", String(scale));

  if (slide.kind !== "cover" && slide.kind !== "closing") {
    const line = el("div", "accent-rule");
    Object.assign(line.style, toCssRegion({ x: 66, y: 132, w: 1148, h: 4 }, scale));
    root.appendChild(line);
  }

  if (slide.kind === "cover" || slide.kind === "closing") {
    const kicker = regionDiv(layout.regions.kicker, scale, "slide-kicker");
    kicker.textContent = slide.kind === "cover" ? "PROOFDECK · DECISION BRIEF" : "NEXT DECISION";
    root.appendChild(kicker);

    const title = regionDiv(layout.regions.title, scale, "slide-title cover-title");
    title.style.fontSize = `${t.titleSize * scale}px`;
    title.textContent = slide.title || "";
    root.appendChild(title);

    const subtitle = regionDiv(layout.regions.subtitle, scale, "slide-subtitle");
    subtitle.style.fontSize = `${t.bodySize * scale}px`;
    subtitle.textContent = slide.subtitle || "";
    root.appendChild(subtitle);
  } else {
    const title = regionDiv(layout.regions.title, scale, "slide-title");
    title.style.fontSize = `${t.titleSize * scale}px`;
    title.textContent = slide.title || "";
    root.appendChild(title);
  }

  if (slide.kind === "comparison") {
    const makeColumn = (region, columnTitle, items) => {
      const box = regionDiv(region, scale, "comparison-card");
      const h = el("h4", "", columnTitle || "");
      box.appendChild(h);
      const list = el("ul", "slide-list");
      for (const item of items || []) {
        const li = el("li");
        addTextWithRefs(li, item.text, item.sourceRefs);
        list.appendChild(li);
      }
      box.appendChild(list);
      return box;
    };
    root.appendChild(makeColumn(layout.regions.left, slide.leftTitle, slide.leftItems));
    root.appendChild(makeColumn(layout.regions.right, slide.rightTitle, slide.rightItems));
  } else if (slide.kind === "metrics") {
    const body = regionDiv(layout.regions.body, scale, "metric-grid");
    for (const metric of slide.metrics || []) {
      const card = el("div", "metric-card");
      const value = el("strong", "metric-value", metric.value || "");
      value.style.fontSize = `${42 * scale}px`;
      const label = el("span", "metric-label", metric.label || "");
      const refs = el("div", "metric-refs");
      refs.append(...(metric.sourceRefs || []).map((ref) => el("span", "source-ref", ref)));
      card.append(value, label, refs);
      body.appendChild(card);
    }
    root.appendChild(body);
  } else if (slide.kind !== "cover" && slide.kind !== "closing") {
    const body = regionDiv(layout.regions.body, scale, "slide-body");
    const list = el("ul", "slide-list");
    (slide.bullets || []).forEach((item, index) => {
      const li = el("li");
      if (slide.kind === "timeline") {
        const number = el("span", "step-number", String(index + 1).padStart(2, "0"));
        li.appendChild(number);
      }
      addTextWithRefs(li, item.text, item.sourceRefs);
      list.appendChild(li);
    });
    body.appendChild(list);
    root.appendChild(body);
  }

  const used = collectRefs(slide);
  if (used.length) {
    const footer = regionDiv(layout.regions.footer, scale, "slide-source-footer");
    footer.textContent = `Sources: ${used.join(" · ")}`;
    root.appendChild(footer);
  }

  container.replaceChildren(root);
  container.style.aspectRatio = `${layout.canvas.w} / ${layout.canvas.h}`;
}
