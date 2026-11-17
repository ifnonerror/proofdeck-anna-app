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
  },
};

export function themeToken(name) {
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

export function layoutSlide(slide, themeName) {
  return {
    canvas: { ...CANVAS },
    themeName: THEME_TOKENS[themeName] ? themeName : "executive",
    tokens: themeToken(themeName),
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
