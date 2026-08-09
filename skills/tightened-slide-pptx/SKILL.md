---
name: tightened-slide-pptx
description: Create editable PowerPoint decks (.pptx) with the Tightened Slide visual system, registered layout discipline, content-density planning, bounded generated visual assets, and strict editability validation. Use for technical, analysis, product, launch, framework, or data-driven presentations; for Slide Prompter prompts that must become editable PPTX; and for hybrid decks where visible text, data, labels, and core diagrams remain native while genuine photos or illustrations may be raster assets. Do not use for HTML decks or flattened image-only slides.
---

# Tightened Slide PPTX

Build editable PowerPoint files. Treat completed-slide PNG renders as QA artifacts only, never as the slide implementation.

## Output Contract

- Deliver `.pptx` as the primary artifact.
- Create every visible title, paragraph, label, rule, box, connector, claim, chart, table, and core diagram node as an editable PowerPoint object.
- Never place a completed slide screenshot or full-slide PNG/JPEG into the PPTX.
- Use bounded bitmap images for genuine photos, illustrations, screenshots, scientific cutaways, physical scenes, or supplied evidence.
- Keep explanatory text and diagram labels outside bitmap images.
- Preserve the requested language, page count, narrative, and visual system.

## Required Presentation Workflow

Use the available `Presentations` skill for every deck. Follow its `@oai/artifact-tool` implementation, render, inspection, speaker-note, and delivery requirements.

Before coding:

1. Read `references/layouts.md`.
2. Read `references/themes.md`.
3. Read `references/content-density.md`.
4. Read `references/images.md` when the plan includes photos, screenshots, search, or ImageGen.
5. Read `references/slide-prompter.md` when the user supplies a prompt from Slide Prompter or requests Slide Prompter compatibility.
6. Read `references/qa.md`.
7. Read `references/artifact-tool-patterns.md` when creating the implementation module.

## Clarify Only What Matters

If inputs are incomplete, ask at most three questions:

1. Audience and presentation setting.
2. Target duration or slide count.
3. Source material, language, required data/assets, and hard constraints.

Default to Korean only when the source is Korean. Otherwise follow the source language. Default to International Klein Blue when no theme is specified.

## Plan Before Authoring

Write a planning table in a temporary `.txt` file:

```text
page -> layout id -> narrative job -> density tier -> evidence modules -> native object plan -> image slot if any
```

Give every slide one primary claim. Give standard body slides two to four useful evidence modules instead of leaving them as a title plus one sentence. Merge adjacent slides that answer the same question before reducing type size or inflating page count. Follow `references/content-density.md`.

## Native Object Rules

- Use textboxes for all visible text.
- Use native rectangles, ellipses, stars, lines, and connectors for simple schematics.
- Create connectors before nodes when connectors attach to shapes.
- Use native PowerPoint charts and tables for real quantitative data.
- Use Graphviz only for genuinely complex relational diagrams, and keep labels editable when practical.
- Use image generation or image search only for actual visual assets, never to render a complete slide.
- Generate visual assets before slide assembly, inspect the crop at the intended aspect ratio, and keep factual labels as native text.
- Do not hide a flattened slide behind invisible editable text.
- Name important objects so `presentation.inspect()` can locate them.
- Add `[Sources]` blocks to speaker notes for externally sourced claims and assets, including generated-asset provenance and prompt-record locations.

## Visual Rules

- Use a 16:9 canvas, normally `1600 × 900` in artifact-tool units.
- Use one accent color per deck.
- Use flat rectangles, hairline rules, and deliberate whitespace.
- Do not use gradients, shadows, glass, neon, or decorative rounded cards as native slide chrome. Restrained lighting or glow inside a genuine photo or generated illustration is allowed when it supports the subject and deck palette.
- Keep body titles on the left/top axis unless using a registered statement or split layout.
- Keep large titles light; use stronger weights only for compact labels and emphasis.
- Use a Korean-capable font for Korean copy and a stable sans-serif fallback.
- Use consistent visual encoding for supported, observed, risk, and unsupported states.

## Build Sequence

1. Load workspace dependencies.
2. Normalize the user brief, including any Slide Prompter prompt.
3. Cluster content into slide questions and run the density/compression pass.
4. Register a layout id, density tier, native object plan, and optional image slot for every slide.
5. Source or generate each approved visual asset and record its prompt, source, aspect ratio, and provenance.
6. Initialize the Presentation artifact-tool workspace.
7. Create an `.mjs` implementation in the temporary workspace.
8. Create the presentation with `Presentation.create({ slideSize: { width: 1600, height: 900 } })`.
9. Assemble every slide with native objects plus bounded visual assets where planned.
10. Inspect the in-memory deck for slides, textboxes, shapes, images, notes, charts, and tables.
11. Export each slide render and layout JSON to the temporary workspace.
12. Export the final PPTX with `PresentationFile.exportPptx()`.
13. Render the exported PPTX again with the Presentation skill's renderer.
14. Run overflow checks and `scripts/validate-editable-pptx.mjs`.
15. Fix every error and visually inspect every slide at full size.

## Validation Commands

Run the Presentation skill's overflow test, then run:

```bash
node scripts/validate-editable-pptx.mjs path/to/final.pptx
```

The validator must confirm that slides contain native objects and are not implemented as full-slide images. It reports picture counts; verify image alt metadata with the in-memory `presentation.inspect()` result before export. Legitimate full-bleed photography may produce a warning and requires visual review.

## Delivery

- Deliver the final `.pptx` exactly once.
- State whether photographs or other non-editable assets are present.
- State that the deck passed rendered-slide review, overflow checks, and editability validation.
- Do not deliver a flattened PPTX unless the user explicitly requests a locked or image-only version.
