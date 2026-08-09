# Editable PPTX QA

## Editability

- Every visible line of copy is a native textbox.
- Every box, rule, marker, node, and connector is a native object.
- Charts and tables remain native when their data must be edited.
- No completed slide is embedded as a single image.
- No invisible text or shapes are used to disguise a flattened slide.
- Photographs, screenshots, and illustrations are the only expected bitmap assets.

## Visual Review

- Render every exported slide at 16:9.
- Inspect every slide individually at full size.
- Check hierarchy, title wrapping, body fit, label clarity, diagram alignment, and visual consistency.
- Fix all unintended overlaps, clipping, and objects outside the canvas.
- Confirm that supporting labels clear footers and page markers.

## Structural Review

- Run `presentation.inspect()` before export.
- Export layout JSON for every slide.
- Run the Presentation skill's slide overflow test on the final PPTX.
- Run `scripts/validate-editable-pptx.mjs` on the final PPTX.
- Reopen or re-import the final PPTX and verify that native objects remain present after export.

## Content Review

- Do not invent results, citations, people, or quantitative claims.
- Mark illustrative values explicitly.
- Keep audience-facing copy concise.
- Add `[Sources]` notes for every external claim and asset.
