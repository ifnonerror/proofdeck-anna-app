export const CANVAS = { w: 1280, h: 720 };

export const THEME_TOKENS = {
  executive: {
    background: "0F1B33",
    foreground: "F4F6FB",
    muted: "9FB0CC",
    accent: "E0A458",
    card: "16264A",
    titleSize: 40,
    bodySize: 23,
    metaSize: 11,
    titleFont: "Microsoft YaHei",
    bodyFont: "Microsoft YaHei",
  },
  minimal: {
    background: "FCFAF5",
    foreground: "1F2937",
    muted: "6B7280",
    accent: "D66536",
    card: "F2EDE4",
    titleSize: 38,
    bodySize: 23,
    metaSize: 11,
    titleFont: "Microsoft YaHei",
    bodyFont: "Microsoft YaHei",
  },
  signal: {
    background: "0D1419",
    foreground: "EDF7FF",
    muted: "8EA4B5",
    accent: "4CC2FF",
    card: "17232C",
    titleSize: 40,
    bodySize: 23,
    metaSize: 11,
    titleFont: "Microsoft YaHei",
    bodyFont: "Microsoft YaHei",
  },
};

export function normalizeThemeTokens(value) {
  if (!value || typeof value !== "object") return null;
  const color = (input, fallback) => (
    /^[0-9a-f]{6}$/i.test(String(input || "")) ? String(input).toUpperCase() : fallback
  );
  const font = (input, fallback) => String(input || fallback).replace(/["\r\n]/g, "").slice(0, 80);
  return {
    name: String(value.name || "Imported theme").slice(0, 100),
    background: color(value.background, THEME_TOKENS.executive.background),
    foreground: color(value.foreground, THEME_TOKENS.executive.foreground),
    muted: color(value.muted, THEME_TOKENS.executive.muted),
    accent: color(value.accent, THEME_TOKENS.executive.accent),
    card: color(value.card, THEME_TOKENS.executive.card),
    titleSize: 40,
    bodySize: 23,
    metaSize: 11,
    titleFont: font(value.titleFont, THEME_TOKENS.executive.titleFont),
    bodyFont: font(value.bodyFont, THEME_TOKENS.executive.bodyFont),
  };
}

export function themeToken(name, customTheme = null) {
  if (name === "custom") return normalizeThemeTokens(customTheme) || THEME_TOKENS.executive;
  return THEME_TOKENS[name] || THEME_TOKENS.executive;
}

export function regionsFor(kind) {
  if (kind === "cover") {
    return {
      kicker: { x: 76, y: 165, w: 1128, h: 35 },
      title: { x: 76, y: 220, w: 1080, h: 150 },
      subtitle: { x: 76, y: 400, w: 1020, h: 85 },
      footer: { x: 76, y: 650, w: 1128, h: 25 },
    };
  }
  if (kind === "closing") {
    return {
      kicker: { x: 76, y: 185, w: 1128, h: 35 },
      title: { x: 76, y: 255, w: 1080, h: 120 },
      subtitle: { x: 76, y: 395, w: 1020, h: 85 },
      footer: { x: 76, y: 650, w: 1128, h: 25 },
    };
  }
  if (kind === "comparison") {
    return {
      title: { x: 66, y: 45, w: 1148, h: 82 },
      left: { x: 66, y: 158, w: 554, h: 470 },
      right: { x: 660, y: 158, w: 554, h: 470 },
      footer: { x: 66, y: 650, w: 1148, h: 25 },
    };
  }
  if (kind === "metrics") {
    return {
      title: { x: 66, y: 45, w: 1148, h: 82 },
      body: { x: 66, y: 165, w: 1148, h: 430 },
      footer: { x: 66, y: 650, w: 1148, h: 25 },
    };
  }
  return {
    title: { x: 66, y: 45, w: 1148, h: 82 },
    body: { x: 66, y: 160, w: 1148, h: 455 },
    footer: { x: 66, y: 650, w: 1148, h: 25 },
  };
}

export function layoutSlide(slide, themeName, customTheme = null) {
  const resolvedTheme = themeName === "custom" && !customTheme ? "executive" : themeName;
  return {
    canvas: { ...CANVAS },
    themeName: resolvedTheme === "custom" ? "custom" : (THEME_TOKENS[resolvedTheme] ? resolvedTheme : "executive"),
    tokens: themeToken(resolvedTheme, customTheme),
    regions: regionsFor(slide.kind),
  };
}

export function toCssRegion(region, scale = 1) {
  return {
    left: `${region.x * scale}px`,
    top: `${region.y * scale}px`,
    width: `${region.w * scale}px`,
    height: `${region.h * scale}px`,
  };
}
