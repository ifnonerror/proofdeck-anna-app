import test from "node:test";
import assert from "node:assert/strict";

import { deckFromPayload } from "../bundle/fixture.js";

test("entry payload keeps source units so its example references remain valid", () => {
  const result = deckFromPayload({
    topic: "Turn a launch brief into a decision deck",
    audience: "Leadership",
    purpose: "Approve the launch",
    sourceText: "Launch is scheduled for May.\n\nThe pilot reduced support time by 18%.",
  });

  assert.deepEqual(
    result.deck.sourceUnits.map((unit) => unit.id),
    ["S1", "S2"]
  );
  assert.deepEqual(result.deck.slides[1].bullets[0].sourceRefs, ["S1"]);
});
