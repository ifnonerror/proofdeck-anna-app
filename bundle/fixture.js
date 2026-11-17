import { createDeck, normalizeDeck, numberSource } from "./model.js";

export const SAMPLE_BRIEF = {
  topic: "把 Q3 华东区业务复盘整理成管理层决策汇报",
  audience: "公司管理层",
  purpose: "确认增长质量并批准 Q4 试点",
  language: "zh-CN",
  slideCount: 8,
  tone: "executive",
  theme: "executive",
  sourceText:
    "Q3 华东区营收 4,820 万元，同比增长 18%，其中新客户贡献 41%。\n\n" +
    "续约率从 Q2 的 87% 降至 82%，主要流失集中在 20 人以下小微客户。\n\n" +
    "销售团队反馈：报价审批平均 4.2 天，是丢单原因第二位（28%）。\n\n" +
    "竞品 A 在 9 月推出按席位年付方案，价格低约 15%。\n\n" +
    "建议 Q4 试点“预审报价单”机制，将审批压缩到 1 个工作日。",
};

const sourceUnits = [
  { id: "S1", text: "Q3 华东区营收 4,820 万元，同比增长 18%，其中新客户贡献 41%。" },
  { id: "S2", text: "续约率从 Q2 的 87% 降至 82%，主要流失集中在 20 人以下小微客户。" },
  { id: "S3", text: "销售团队反馈：报价审批平均 4.2 天，是丢单原因第二位（28%）。" },
  { id: "S4", text: "竞品 A 在 9 月推出按席位年付方案，价格低约 15%。" },
  { id: "S5", text: "建议 Q4 试点“预审报价单”机制，将审批压缩到 1 个工作日。" },
];

export const FIXTURE_DECK = normalizeDeck(createDeck({
  id: "fixture-deck",
  title: "Q3 区域增长复盘与 Q4 决策建议",
  subtitle: "华东区 · 管理层业务评审会 · 2026-10",
  theme: "executive",
  language: "zh-CN",
  sourceText: SAMPLE_BRIEF.sourceText,
  sourceUnits,
  slides: [
    {
      id: "s1",
      kind: "cover",
      title: "Q3 增长达标，但续约与报价流程正在侵蚀优势",
      subtitle: "华东区业务评审 · 2026-10",
      notes: "开场 30 秒先给结论：增长达标，但有两个风险需要今天拍板。",
    },
    {
      id: "s2",
      kind: "agenda",
      title: "今天要决定的三件事",
      bullets: [
        { text: "增长来源是否健康", sourceRefs: ["S1"] },
        { text: "续约下滑是否需要专项干预", sourceRefs: ["S2"] },
        { text: "报价审批是否启动 Q4 试点", sourceRefs: ["S3", "S5"] },
      ],
      notes: "第三项需要当场拍板。",
    },
    {
      id: "s3",
      kind: "metrics",
      title: "增长由新客户驱动，结构比数字更重要",
      metrics: [
        { value: "4,820万", label: "Q3 华东营收", sourceRefs: ["S1"] },
        { value: "+18%", label: "同比增长", sourceRefs: ["S1"] },
        { value: "41%", label: "新客户贡献", sourceRefs: ["S1"] },
      ],
      notes: "新客户占比高意味着续约风险会被放大。",
    },
    {
      id: "s4",
      kind: "bullets",
      title: "续约率下滑 5 个百分点，集中在小微客户",
      bullets: [
        { text: "续约率从 87% 降至 82%", sourceRefs: ["S2"] },
        { text: "流失集中在 20 人以下客户", sourceRefs: ["S2"] },
        { text: "小微客户对价格与响应速度最敏感", sourceRefs: ["S2", "S4"] },
      ],
      notes: "先看价格与服务响应，不要直接归因为产品。",
    },
    {
      id: "s5",
      kind: "comparison",
      title: "丢单主因是流程，不是产品",
      leftTitle: "内部流程",
      rightTitle: "外部竞争",
      leftItems: [
        { text: "报价审批平均 4.2 天", sourceRefs: ["S3"] },
        { text: "审批慢是丢单原因第二位", sourceRefs: ["S3"] },
      ],
      rightItems: [
        { text: "竞品价格低约 15%", sourceRefs: ["S4"] },
        { text: "流程速度可以部分抵消价格差", sourceRefs: ["S5"] },
      ],
      notes: "流程问题可自行改善，价格问题只能部分应对。",
    },
    {
      id: "s6",
      kind: "timeline",
      title: "Q4 试点路径",
      bullets: [
        { text: "10 月：确定试点范围和负责人", sourceRefs: ["S5"] },
        { text: "11 月：审批流程压至 1 个工作日", sourceRefs: ["S3", "S5"] },
        { text: "12 月：复盘成交率与续约变化", sourceRefs: ["S2"] },
      ],
      notes: "每个阶段明确指标和负责人。",
    },
    {
      id: "s7",
      kind: "bullets",
      title: "需要的两项决策",
      bullets: [
        { text: "批准 Q4 报价预审试点", sourceRefs: ["S5"] },
        { text: "指定小微客户续约专项负责人", sourceRefs: ["S2"] },
      ],
      notes: "收尾时明确 owner 和启动日期。",
    },
    {
      id: "s8",
      kind: "closing",
      title: "决策请求",
      subtitle: "批准试点 · 明确负责人 · 12 月复盘",
      notes: "感谢与会者，确认下一步行动人。",
    },
  ],
}));

export function deckFromPayload(payload) {
  if (!payload || typeof payload !== "object") return FIXTURE_DECK;
  const brief = {
    ...SAMPLE_BRIEF,
    ...payload,
    sourceText: String(payload.sourceText || SAMPLE_BRIEF.sourceText),
  };
  const units = numberSource(brief.sourceText);
  const example = normalizeDeck(createDeck({
    id: FIXTURE_DECK.id,
    title: brief.topic || FIXTURE_DECK.title,
    subtitle: [brief.audience, brief.purpose].filter(Boolean).join(" · "),
    theme: brief.theme || "executive",
    language: brief.language === "en" ? "en" : "zh-CN",
    sourceText: brief.sourceText,
    sourceUnits: units,
    slides: FIXTURE_DECK.slides.map((slide) => ({ ...slide })),
  }));
  return { deck: example, brief };
}
