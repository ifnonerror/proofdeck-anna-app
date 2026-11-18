import test from "node:test";
import assert from "node:assert/strict";

import {
  buildDeckRequest,
  buildPptxEditRequest,
  extractJSON,
  formatGenerationError,
  generateDeckWithAnna,
  modelPreferenceSatisfied,
  outputTokenBudget,
  parsePptxEditReply,
  parseDeckReply,
} from "../bundle/ai.js";
import { numberSource } from "../bundle/model.js";

const brief = {
  topic: "Q3 business review",
  audience: "Leadership",
  purpose: "Approve the next step",
  language: "en",
  slideCount: 3,
  tone: "executive",
  sourceText: "Revenue grew 18%.\n\nRenewal declined.",
};

function makeReply(text, extra = {}) {
  return {
    role: "assistant",
    content: { type: "text", text },
    model: "mock-model",
    stopReason: "endTurn",
    usage: { inputTokens: 100, outputTokens: 120, totalTokens: 220 },
    _meta: { quotaConsumed: 0.02 },
    ...extra,
  };
}

const validDeck = {
  title: "Q3 decision",
  subtitle: "Leadership review",
  slides: [
    { kind: "cover", title: "Growth is healthy, but renewal needs action", bullets: [], notes: "Open." },
    {
      kind: "bullets",
      title: "Renewal is the main risk",
      bullets: [
        { text: "Revenue grew 18%", sourceRefs: ["S1"] },
        { text: "Renewal declined", sourceRefs: ["S2"] },
      ],
      notes: "Explain.",
    },
    { kind: "closing", title: "Approve a renewal task force", bullets: [], notes: "Close." },
  ],
};

test("buildDeckRequest uses one bounded call and includes source units", () => {
  const request = buildDeckRequest(brief, numberSource(brief.sourceText));
  assert.equal(request.maxTokens, 2800);
  assert.equal(request.temperature, 0.5);
  assert.equal(Object.hasOwn(request, "thinking"), false);
  assert.equal(request.messages.length, 1);
  assert.match(request.messages[0].content, /S2/);
});

test("buildDeckRequest forwards a selected model preference", () => {
  const request = buildDeckRequest({ ...brief, modelHint: "claude-sonnet-4.5" }, []);
  assert.deepEqual(request.modelPreferences.hints, [
    { family: "claude", name: "Claude Sonnet 4.5" },
  ]);
  assert.equal(request.modelPreferences.intelligencePriority, 0.8);
});

test("extractJSON accepts fenced and prefixed JSON", () => {
  assert.deepEqual(extractJSON('```json\n{"a":1}\n```'), { a: 1 });
  assert.deepEqual(extractJSON('Result follows: {"a":2}'), { a: 2 });
});

test("parseDeckReply validates, sanitizes, and reports metadata", () => {
  const units = numberSource(brief.sourceText);
  const reply = makeReply(JSON.stringify(validDeck));
  const result = parseDeckReply(reply, 3, units);
  assert.equal(result.deck.slides.length, 3);
  assert.equal(result.deck.slides[1].bullets[0].sourceRefs[0], "S1");
  assert.equal(result.model, "mock-model");
  assert.equal(result.quotaConsumed, 0.02);
});

test("parseDeckReply removes invalid source references with a warning", () => {
  const bad = structuredClone(validDeck);
  bad.slides[1].bullets[0].sourceRefs = ["S99"];
  const result = parseDeckReply(makeReply(JSON.stringify(bad)), 3, numberSource(brief.sourceText));
  assert.deepEqual(result.deck.slides[1].bullets[0].sourceRefs, []);
  assert.match(result.warnings[0], /S99/);
});

test("parseDeckReply rejects truncation and invalid JSON", () => {
  assert.throws(
    () => parseDeckReply(makeReply("{}", { stopReason: "maxTokens" }), 3, []),
    /长度限制/
  );
  assert.throws(() => parseDeckReply(makeReply("not json"), 3, []), /无法解析/);
});

test("parseDeckReply preserves usable content when the requested page count differs", () => {
  const twoSlideReply = JSON.stringify({
    ...validDeck,
    slides: [
      { kind: "bullets", title: "Growth needs a renewal plan" },
      { kind: "bullets", title: "Approve the next step" },
    ],
  });
  const result = parseDeckReply(makeReply(twoSlideReply), 3, []);
  assert.equal(result.deck.slides.length, 2);
  assert.equal(result.deck.slides[0].kind, "cover");
  assert.equal(result.deck.slides[1].kind, "closing");
  assert.match(result.warnings.join(" "), /请求 3 页/);
});

test("generateDeckWithAnna calls llm.complete exactly once", async () => {
  let calls = 0;
  const mockAnna = {
    llm: {
      async complete(request) {
        calls += 1;
        assert.equal(request.maxTokens, 2800);
        return makeReply(JSON.stringify(validDeck));
      },
    },
  };
  const result = await generateDeckWithAnna(mockAnna, brief, numberSource(brief.sourceText));
  assert.equal(calls, 1);
  assert.equal(result.deck.title, "Q3 decision");
});

test("generateDeckWithAnna retries only mandatory-reasoning compatibility errors", async () => {
  let calls = 0;
  const mockAnna = {
    llm: {
      async complete(request) {
        calls += 1;
        if (calls === 1) {
          assert.equal(Object.hasOwn(request, "thinking"), false);
          throw new Error("upstream LLM call failed: Reasoning is mandatory for this endpoint and cannot be disabled.");
        }
        assert.deepEqual(request.thinking, { enabled: true, effort: "minimal" });
        return makeReply(JSON.stringify(validDeck));
      },
    },
  };
  const result = await generateDeckWithAnna(mockAnna, brief, numberSource(brief.sourceText));
  assert.equal(calls, 2);
  assert.equal(result.deck.slides.length, 3);
});

test("outputTokenBudget scales with requested deck length", () => {
  assert.equal(outputTokenBudget(6), 2800);
  assert.equal(outputTokenBudget(8), 4200);
  assert.equal(outputTokenBudget(10), 5600);
});

test("modelPreferenceSatisfied detects account routing overrides", () => {
  assert.equal(modelPreferenceSatisfied("claude-sonnet-4.5", "anthropic/claude-sonnet-4.5"), true);
  assert.equal(modelPreferenceSatisfied("gemini-3.5-flash", "google/gemini-3.5-flash"), true);
  assert.equal(modelPreferenceSatisfied("claude-sonnet-4.5", "xiaomi/mimo-v2.5"), false);
  assert.equal(modelPreferenceSatisfied("auto", "xiaomi/mimo-v2.5"), true);
});

test("formatGenerationError turns quota errors into an actionable message", () => {
  const error = new Error(
    'HTTP 429: {"detail":{"errorCode":"APP_QUOTA_EXCEEDED","data":{"requiredCu":0.9172,"remainingCu":0}}}'
  );
  assert.match(formatGenerationError(error), /额度不足/);
  assert.match(formatGenerationError(error), /0\.917 CU/);
});

test("formatGenerationError hides raw HTML for a model gateway outage", () => {
  const error = new Error("HTTP 502: <!DOCTYPE html><title>Bad gateway</title>");
  assert.equal(
    formatGenerationError(error),
    "Anna 模型网关暂时不可用（HTTP 502）。请几分钟后重试，或切换到其他模型。"
  );
});

test("PPTX edit replies can only patch the selected shape", () => {
  const target = { slideIndex: 2, shapeId: 7 };
  const reply = makeReply(JSON.stringify({
    reply: "Adjusted the title.",
    operations: [
      { slideIndex: 2, shapeId: 7, action: "setTextStyle", fontSize: 28, color: "#C75B39", dx: 999, dy: 999 },
      { slideIndex: 1, shapeId: 3, action: "replaceText", text: "Wrong target" },
      { slideIndex: 2, shapeId: 7, action: "move", dx: 100, dy: -20 },
    ],
  }));
  const result = parsePptxEditReply(reply, target);
  assert.equal(result.operations.length, 2);
  assert.equal(result.operations[0].fontSize, 28);
  assert.equal(result.operations[0].dx, undefined);
  assert.equal(result.operations[0].dy, undefined);
  assert.equal(result.operations[1].dx, 100);
});

test("PPTX edit request asks for a targeted JSON patch", () => {
  const request = buildPptxEditRequest({
    instruction: "Make the title larger",
    target: { slideIndex: 0, shapeId: 2, text: "Title", fontSize: 24 },
  });
  assert.equal(request.maxTokens, 900);
  assert.match(request.systemPrompt, /只允许操作/);
  assert.match(request.messages[0].content, /Make the title larger/);
});
