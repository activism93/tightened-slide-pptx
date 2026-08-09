# Editable PPTX QA

## Editability

- Every visible line of copy is a native textbox.
- Every box, rule, marker, node, and connector is a native object.
- Charts and tables remain native when their data must be edited.
- No completed slide is embedded as a single image.
- No invisible text or shapes are used to disguise a flattened slide.
- Photographs, screenshots, and illustrations are the only expected bitmap assets.
- Every bitmap has useful in-memory alt metadata and a source or generated-asset provenance entry in `[Sources]` notes. If the export path drops alt metadata, keep the object identifiable through its name and notes rather than claiming exported alt-text support.

## Content Density

- Every D2/D3 slide has one claim and at least two useful evidence modules.
- Adjacent slides do not repeat the same definition, equation, visual basis, or conclusion.
- Sparse D0/D1 slides are deliberate, not filler inserted to raise the page count.
- Compression does not reduce body text below 16 pt or combine unrelated mental models.
- The final slide count follows the content clusters, not the original number of headings.

## Image Review

- Generated assets contain no baked-in audience-facing text, numbers, logos, page chrome, or watermark.
- Factual labels, process order, equations, and specifications remain native.
- Image crops preserve the subject at final slide size.
- Conceptual scientific visuals are marked `not to scale` when appropriate.
- Full-bleed images are genuine photos or illustrations, never completed slide renders.

## Visual Review

- Render every exported slide at 16:9.
- Inspect every slide individually at full size.
- Check hierarchy, title wrapping, body fit, label clarity, diagram alignment, and visual consistency.
- Check that hybrid slides read as one composition rather than an image pasted beside unrelated cards.
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
