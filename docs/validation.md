# ProofDeck validation record

Test date: 2026-11-18

Environment:

- Anna staging: `https://staging.anna.partners`
- Developer account: `liuzk`
- Developer handle: `@liuzk2025`
- App: ProofDeck (`proofdeck-ai`)
- App ID: `274`

## Build and package gates

- `npm test`: 20 tests passed
- `npm run smoke`: 8 slides, 8 speaker-note pages, valid 16:9 package
- `npm run validate`: Anna schema 4 strict validation passed
- Production host-mediated download was exercised in the installed Anna App
- Downloaded decks were opened in Microsoft PowerPoint and exported to 1280x720 slide images

## Real model runs

| Model | Scenario | Result | Cost |
| --- | --- | --- | --- |
| `xiaomi/mimo-v2.5` | Six-slide Chinese decision deck | Passed, valid PPTX and notes | 0.276 CU |
| `xiaomi/mimo-v2.5` | Ten-slide Chinese analytical deck | Passed, 10 rendered slides and 10 notes | 0.363 CU |
| `xiaomi/mimo-v2.5` | Eight-slide English persuasive deck | Passed, source references retained | 0.324 CU |
| `xiaomi/mimo-v2.5` | Six-slide product roadmap | Passed, valid 6-slide PPTX | 0.270 CU |
| `xiaomi/mimo-v2.5` | Final v0.1.7 mobile checkout review | Passed, valid 6-slide PPTX | 0.247 CU |
| `anthropic/claude-sonnet-4.5` | Eight-slide financial and cash-runway review | Passed, valid 8-slide PPTX | 24.579 CU |
| `google/gemini-3.5-flash` | Eight-slide randomized-trial research summary | Passed after output-cap increase | 28.301 CU |
| `qwen3.8-flash` | Six-slide project-risk report | Passed, 5 source units | 0.543 CU |
| `xiaomi/mimo-v2.5` | Final v0.1.7 onboarding rollout decision | Passed, valid 6-slide PPTX | 0.256 CU |
| `xiaomi/mimo-v2.5` | Six-slide training deck with no source material | Passed, zero fabricated citations | 0.233 CU |

Total model testing consumed approximately `82.18 CU`. Remaining quota after the final run was approximately `917.82 CU`.

## Defects found by cross-model testing

1. A model returned seven slides when eight were requested, and the App discarded the otherwise valid deck. The parser now preserves usable slides, reports the mismatch, and fixes the cover/closing roles.
2. Gemini produced longer JSON than MiMo at the same page count and hit the former 3,200-token cap. The cap now scales to 2,800, 4,200, and 5,600 tokens for 6, 8, and 10 slides.
3. GPT-5 Mini, Kimi K2.5, and MiniMax M2.7 require reasoning on their OpenRouter endpoints. ProofDeck now performs exactly one compatibility retry with minimal reasoning only when the provider explicitly rejects a no-reasoning request. During testing these endpoints were still returning HTTP 502 from the staging gateway, so they are not reported as successful model runs.

## Model-routing caveat

The final model route is controlled by the account-level model selected in Anna's `/llm` settings. `modelPreferences` in the request act as hints, while the account default can override them. ProofDeck displays the actual model returned by Anna and warns when the returned model family differs from the App's requested preference.
