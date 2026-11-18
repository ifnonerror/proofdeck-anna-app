# ProofDeck validation record

Test date: 2026-11-18

Environment:

- Anna staging: `https://staging.anna.partners`
- Developer account: `liuzk`
- Developer handle: `@liuzk2025`
- App: ProofDeck
- App ID: `274`
- Slug: `proofdeck-ai`
- Review candidate: `v0.2.3` (`version_id=667`)

## Automated gates

- `npm test`: 31 tests passed.
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
- Free-form non-placeholder body text is cleared instead of leaving stale donor copy.
- Repeated fixed text and slide chrome outside the content area remain intact.

## Browser acceptance

Browser acceptance used Microsoft Edge through `playwright-core` and both mock and real Anna harnesses.

| Scenario | Result |
| --- | --- |
| Initial app state | Passed; single-column chat with no viewer, no deck link, and no advanced settings bar |
| Model-only composer controls | Passed; the model selector is visible in the composer |
| PPTX template upload | Passed; slides, layouts, masters, and components reported as loaded |
| Template upload before generation | Passed; the app stays in chat and does not expose a PPT area |
| PDF source upload | Passed; selectable text extracted in the browser |
| DOCX source upload | Passed; document text extracted in the browser |
| Six-slide AI generation from template | Passed with real Anna model |
| Mock eight-slide template generation | Passed |
| Generated presentation link | Passed; title and slide count appear in the conversation |
| Opening the presentation link | Passed; reveals the left-chat/right-presentation workspace |
| Returning to chat | Passed; closes the viewer and preserves the generated link |
| Reload with a completed deck | Passed; restores into chat with the link instead of opening the viewer |
| Slide thumbnail row layout | Passed; each 16:9 thumbnail reserves its full height and never overlaps the next slide |
| Compact generation status | Passed; source-reference cleanup is summarized instead of printing every warning below the composer |
| Multi-task sidebar | Passed; separate tasks retain their own prompt, template, conversation and generated deck |
| Task switching | Passed; switching between two projects restores the correct content without mixing state |
| Task-list reload | Passed; both tasks and the active task survive a page reload |
| Shape inspection and click-to-select | Passed |
| Selected-element AI patch | Passed with real Anna model |
| Export after AI patch | Passed; valid PPTX with 8 slides and 8 notes |
| Draft regeneration when IndexedDB lacks PPTX bytes | Passed; restored 8 slides and re-enabled export |
| 720px minimum layout | Passed; no horizontal overflow |
| Page errors during end-to-end flows | None |

The final automated browser pass covered the new flow and all checks passed:

```text
PASS initial state has no presentation workspace
PASS only the model chooser remains in the composer
PASS template import stays in chat
PASS generation stays in chat
PASS generated presentation link appears
PASS presentation link opens 8 template slides
PASS stale RRSI body copy is absent
PASS back to chat closes the workspace
PASS reload restores into chat with the presentation link
PASS slide thumbnails have 8px gaps and no overlap
PASS warning text is compact
PASS task sidebar creates and selects separate projects
PASS prompt, attachment and deck state remain isolated per task
PASS 720px layout has no horizontal overflow
PASS no page errors
```

## Supplied RRSI template acceptance

The supplied 14-slide deck `RRSI 论文解读：智能体脚手架的正则化递归自进化.pptx` was tested without modifying the source file.

- Slide size: 16:9, `12192000 × 6858000` EMU.
- Source slides: 14.
- Layouts: 13.
- Masters: 2.
- Themes: 3.
- Media parts: 15.
- Fonts: Arial.
- Body content uses direct text boxes rather than placeholders.

`npm run smoke:template -- "<template path>"` produced an eight-slide deck from the supplied template. All eight slides rendered with shape-aware SVG. Theme, master, and layout parts were byte-for-byte unchanged:

```text
themes:  3 / 3 unchanged
masters: 2 / 2 unchanged
layouts: 13 / 13 unchanged
OfficeKit package validation: no issues
```

The previous implementation left captions and body sections from the paper on newly generated slides. The revision now clears donor body copy outside repeated fixed text while preserving images, master artwork, backgrounds, and fixed slide chrome. Browser inspection confirmed that `图 1(a)`, `标准进化循环`, and `风险警示` no longer leaked into the generated deck.

## Real Anna model runs

The real acceptance run used Anna's staging model bridge with the `proofdeck` development app and a real PPTX template.

| Model | Scenario | Result | Known quota |
| --- | --- | --- | --- |
| `xiaomi/mimo-v2.5` | Six-slide Chinese management brief from uploaded DOCX and PPTX template | Passed; 6 slides rendered, template mode retained | `0.277 CU` |
| `xiaomi/mimo-v2.5` | Selected title: larger font and template accent color | Passed; model returned a valid single-shape patch | `0.059 CU` |
| `xiaomi/mimo-v2.5` | Repeat selected-element patch followed by export | Passed; valid PPTX, 8 slides, 8 notes | `0.059 CU` |
| `xiaomi/mimo-v2.5` | Post-thumbnail-fix six-slide RRSI report from the supplied template | Passed; 6 slides, valid export, no thumbnail overlap | `0.234 CU` |

Known successful real-model usage for template generation and element editing was `0.629 CU`. A pre-fix `v0.2.1` acceptance also completed successfully for `0.241 CU`; it was used only to reproduce the thumbnail bug. An earlier attempt to force `qwen3.8-flash` produced no response before the harness was stopped after 220 seconds; its billing outcome was not observable from the client and is therefore not counted as a successful run. The account default subsequently routed the request to `xiaomi/mimo-v2.5`.

Quota consumption from earlier generation-only validation was approximately `82.18 CU`. Counting successful final-template and edit runs, the known successful total across development and acceptance is approximately `83.050 CU`; remaining quota is approximately `916.950 CU` before accounting for the unresolved canceled Qwen attempt.

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

`v0.2.3` was uploaded and cut successfully:

- Remote working draft revision: `11`.
- Bundle: 25 files, approximately 3.43 MB.
- Immutable review version: `v0.2.3`, `version_id=667`.
- App status after submission: `pending_review`.
- Review pin: `v0.2.3`.

The Anna web login requires a Cloudflare Turnstile human check in a fresh browser session. The real model acceptance therefore used the same saved Anna PAT and real `anna-runtime` bridge from the App's local Harness. The platform package is installed and pinned for review; a reviewer can install the `0.0.0-draft` version from the Developer Console after human verification.

## Known limitations

1. The app supports `.pptx`, not legacy `.ppt`.
2. Standard PowerPoint shapes, media, tables, and themed backgrounds are cloned. SmartArt, OLE embeds, embedded video, and custom animations depend on OfficeKit fidelity and require validation with each production template.
3. Scanned PDFs require an OCR layer; PDF.js extracts selectable text only.
4. One source file is capped at 15 MB and extracted project text at 60,000 characters.
5. Account-level model routing can override the App's model preference. The UI reports the model actually returned by Anna.
6. The selected-element patch intentionally supports a bounded set of operations: text, typography, color, movement, size, and fill. It does not regenerate a whole slide from an annotation.
