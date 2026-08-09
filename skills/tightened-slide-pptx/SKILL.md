---
name: tightened-slide-pptx
description: Create fully editable native PowerPoint decks (.pptx) with the Tightened Slide visual system, registered layout discipline, and strict editability validation. Use when Codex must build technical, analysis, product, launch, framework, or data-driven presentations where every text box, shape, connector, chart, table, and diagram node must remain individually editable. Do not use for HTML decks or flattened slide images.
---

# Tightened Slide PPTX

Build native PowerPoint files. Treat PNG renders as QA artifacts only, never as the slide implementation.

## Output Contract

- Deliver `.pptx` as the primary artifact.
- Create every visible title, paragraph, label, rule, box, node, connector, chart, and table as an editable PowerPoint object.
- Never place a completed slide screenshot or full-slide PNG/JPEG into the PPTX.
- Use bitmap images only for genuine photos, illustrations, screenshots, or supplied evidence.
- Keep explanatory text and diagram labels outside bitmap images.
- Preserve the requested language, page count, narrative, and visual system.

## Required Presentation Workflow

Use the available `Presentations` skill for every deck. Follow its `@oai/artifact-tool` implementation, render, inspection, speaker-note, and delivery requirements.

Before coding:

1. Read `references/layouts.md`.
2. Read `references/themes.md`.
3. Read `references/qa.md`.
4. Read `references/artifact-tool-patterns.md` when creating the implementation module.

## Clarify Only What Matters

If inputs are incomplete, ask at most three questions:

1. Audience and presentation setting.
2. Target duration or slide count.
3. Source material, language, required data/assets, and hard constraints.

Default to Korean only when the source is Korean. Otherwise follow the source language. Default to International Klein Blue when no theme is specified.

## Plan Before Authoring

Write a planning table in a temporary `.txt` file:

```text
page -> layout id -> narrative job -> native object plan -> image slot if any
```

Give every slide one primary claim. Compress copy before reducing type size.

## Native Object Rules

- Use textboxes for all visible text.
- Use native rectangles, ellipses, stars, lines, and connectors for simple schematics.
- Create connectors before nodes when connectors attach to shapes.
- Use native PowerPoint charts and tables for real quantitative data.
- Use Graphviz only for genuinely complex relational diagrams, and keep labels editable when practical.
- Use image generation or image search only for actual visual assets, never to render a complete slide.
- Do not hide a flattened slide behind invisible editable text.
- Name important objects so `presentation.inspect()` can locate them.
- Add `[Sources]` blocks to speaker notes for externally sourced claims and assets.

## Visual Rules

- Use a 16:9 canvas, normally `1600 × 900` in artifact-tool units.
- Use one accent color per deck.
- Use flat rectangles, hairline rules, and deliberate whitespace.
- Do not use gradients, shadows, glass, neon, or decorative rounded cards.
- Keep body titles on the left/top axis unless using a registered statement or split layout.
- Keep large titles light; use stronger weights only for compact labels and emphasis.
- Use a Korean-capable font for Korean copy and a stable sans-serif fallback.
- Use consistent visual encoding for supported, observed, risk, and unsupported states.

## Build Sequence

1. Load workspace dependencies.
2. Initialize the Presentation artifact-tool workspace.
3. Create an `.mjs` implementation in the temporary workspace.
4. Create the presentation with `Presentation.create({ slideSize: { width: 1600, height: 900 } })`.
5. Implement every slide with native objects.
6. Inspect the in-memory deck for slides, textboxes, shapes, images, notes, charts, and tables.
7. Export each slide render and layout JSON to the temporary workspace.
8. Export the final PPTX with `PresentationFile.exportPptx()`.
9. Render the exported PPTX again with the Presentation skill's renderer.
10. Run overflow checks and `scripts/validate-editable-pptx.mjs`.
11. Fix every error and visually inspect every slide at full size.

## Validation Commands

Run the Presentation skill's overflow test, then run:

```bash
node scripts/validate-editable-pptx.mjs path/to/final.pptx
```

The validator must confirm that slides contain native objects and are not implemented as full-slide images. Legitimate full-bleed photography may produce a warning and requires visual review.

## Delivery

- Deliver the final `.pptx` exactly once.
- State whether photographs or other non-editable assets are present.
- State that the deck passed rendered-slide review, overflow checks, and editability validation.
- Do not deliver a flattened PPTX unless the user explicitly requests a locked or image-only version.
