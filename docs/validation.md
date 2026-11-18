# ProofDeck validation record

Test date: 2026-11-18

Environment:

- Anna staging: `https://staging.anna.partners`
- Developer account: `liuzk`
- Developer handle: `@liuzk2025`
- App: ProofDeck
- App ID: `274`
- Slug: `proofdeck-ai`
- Review candidate: `v0.2.0` (`version_id=663`)

## Automated gates

- `npm test`: 24 tests passed.
- `npm run smoke`: produced an eight-slide fallback deck with eight notes and a valid 16:9 package.
- OfficeKit smoke: produced an eight-slide deck from a real `.pptx`, with shape-aware SVG output.
- `npm run validate`: Anna schema 4 strict validation passed.
- `npm run build`: browser engines and the static Anna bundle built successfully.

The OfficeKit tests verify that:

- A template is parsed and its slides and layouts are discovered.
- Donor slides are cloned from the uploaded `.pptx`.
- The saved file remains a valid PowerPoint package.
- Rendered slides contain `data-pptx-shape-id` attributes.
- A patch changes only the requested slide and shape.

## Browser acceptance

Browser acceptance used Microsoft Edge through `playwright-core` and both mock and real Anna harnesses.

| Scenario | Result |
| --- | --- |
| Initial sample deck | Passed; 8 slides rendered from PPTX |
| PPTX template upload | Passed; slides, layouts, masters, and components reported as loaded |
| PDF source upload | Passed; selectable text extracted in the browser |
| DOCX source upload | Passed; document text extracted in the browser |
| Six-slide AI generation from template | Passed with real Anna model |
| Mock eight-slide template generation | Passed |
| Shape inspection and click-to-select | Passed |
| Selected-element AI patch | Passed with real Anna model |
| Export after AI patch | Passed; valid PPTX with 8 slides and 8 notes |
| Draft regeneration when IndexedDB lacks PPTX bytes | Passed; restored 8 slides and re-enabled export |
| 720px minimum layout | Passed; no horizontal overflow |
| Page errors during end-to-end flows | None |

An automated browser pass covered 12 checks and all passed:

```text
PASS sample deck renders 8 slides
PASS PPTX template structure loads
PASS PDF source extracted
PASS DOCX source extracted
PASS mock Anna generates 8 template slides
PASS exported PPTX has 8 slides and 8 notes: 8/8
PASS generation flow has no page errors
PASS slide shape becomes selectable
PASS AI annotation patches selected shape
PASS annotation flow has no page errors
PASS edited deck remains exportable
PASS 720px layout has no horizontal overflow
```

## Real Anna model runs

The real acceptance run used Anna's staging model bridge with the `proofdeck` development app and a real PPTX template.

| Model | Scenario | Result | Known quota |
| --- | --- | --- | --- |
| `xiaomi/mimo-v2.5` | Six-slide Chinese management brief from uploaded DOCX and PPTX template | Passed; 6 slides rendered, template mode retained | `0.277 CU` |
| `xiaomi/mimo-v2.5` | Selected title: larger font and template accent color | Passed; model returned a valid single-shape patch | `0.059 CU` |
| `xiaomi/mimo-v2.5` | Repeat selected-element patch followed by export | Passed; valid PPTX, 8 slides, 8 notes | `0.059 CU` |

Known successful real-model usage for this release was `0.395 CU`. An earlier attempt to force `qwen3.8-flash` produced no response before the harness was stopped after 220 seconds; its billing outcome was not observable from the client and is therefore not counted as a successful run. The account default subsequently routed the request to `xiaomi/mimo-v2.5`.

Quota consumption from earlier generation-only validation was approximately `82.18 CU`. The known successful total across development and final acceptance is therefore approximately `82.575 CU`; remaining quota is approximately `917.425 CU` before accounting for the unresolved canceled Qwen attempt.

## Historical cross-model generation checks

The generation and parser path is shared with the earlier generation-only release. These runs remain useful evidence for model compatibility:

| Model | Scenario | Result | Quota |
| --- | --- | --- | --- |
| `xiaomi/mimo-v2.5` | Chinese business, roadmap, and strategy decks; 6/8/10 pages | Passed; PPTX and notes validated | 0.27–0.36 CU per run |
| `anthropic/claude-sonnet-4.5` | Eight-page English financial analysis | Passed; 8 slides and 8 notes | 24.579 CU |
| `google/gemini-3.5-flash` | Eight-page clinical-evidence seminar | Passed after increasing the output cap | 28.301 CU |
| `qwen3.8-flash` | Six-page Chinese project-risk report | Passed; 5 source units | 0.543 CU |
| `xiaomi/mimo-v2.5` | Six-page training deck with no source material | Passed; no fabricated source references | 0.233 CU |

GPT-5 Mini, Kimi K2.5, and MiniMax M2.7 exposed mandatory-reasoning errors during the earlier test window. ProofDeck retains a single narrow compatibility retry with minimal reasoning only when the provider explicitly requires reasoning.

## Publishing state

`v0.2.0` was uploaded and cut successfully:

- Remote working draft revision: `8`.
- Bundle: 25 files, approximately 3.49 MB.
- Immutable review version: `v0.2.0`, `version_id=663`.
- App status after submission: `pending_review`.
- Review pin: `v0.2.0`.

The Anna web login requires a Cloudflare Turnstile human check in a fresh browser session. The real model acceptance therefore used the same saved Anna PAT and real `anna-runtime` bridge from the App's local Harness. The platform package is installed and pinned for review; a reviewer can install the `0.0.0-draft` version from the Developer Console after human verification.

## Known limitations

1. The app supports `.pptx`, not legacy `.ppt`.
2. Standard PowerPoint shapes, media, tables, and themed backgrounds are cloned. SmartArt, OLE embeds, embedded video, and custom animations depend on OfficeKit fidelity and require validation with each production template.
3. Scanned PDFs require an OCR layer; PDF.js extracts selectable text only.
4. One source file is capped at 15 MB and extracted project text at 60,000 characters.
5. Account-level model routing can override the App's model preference. The UI reports the model actually returned by Anna.
6. The selected-element patch intentionally supports a bounded set of operations: text, typography, color, movement, size, and fill. It does not regenerate a whole slide from an annotation.
