# Image and ImageGen Assets

Use images to supply physical realism, atmosphere, material texture, or scientific depth. Keep claims and editable information native.

## Choose the Right Surface

| Need | Use |
|---|---|
| Exact labels, equations, data, process order, or editable geometry | Native PowerPoint objects |
| Real person, product, facility, UI, event, or documentary evidence | Supplied or authoritative sourced image |
| Generic physical scene, conceptual apparatus, material cutaway, or aesthetic scientific illustration | ImageGen |
| Real chart or quantitative comparison | Native chart or table |
| Complete slide composition | Never use an image |

Use a hybrid slide when an image improves intuition but the audience must still read exact labels, numbers, or causal steps.

## Image Slot Contract

Define every slot before generation:

```text
role | aspect ratio | bounding box | subject position | safe area | crop mode | native labels | provenance
```

Keep ordinary images bounded. A hero image may be large, but it must not contain the completed title, labels, page number, or slide chrome.

## ImageGen Prompt Pattern

Include:

1. use case and audience;
2. physical or conceptual subject;
3. composition and required subject position;
4. aspect ratio and intended crop;
5. deck palette and realism level;
6. exclusions: no text, labels, numbers, logos, watermark, UI chrome, or slide frame;
7. scientific constraints and misleading implications to avoid.

For a text-left/image-right layout, explicitly place the subject on the right and preserve calm negative space on the left. Generate individual assets, never a full slide.

## Factual Boundary

- Treat generated scientific visuals as conceptual and mark `not to scale` when appropriate.
- Do not use generated images as evidence of real products, branded equipment, historical events, or measured results.
- Keep exact component names, dimensions, specifications, and process steps outside the image.
- Do not imply causal mechanisms that the editable labels contradict.

## Provenance

Store prompt records in a temporary `.txt` file during generation. When source and example files are published together, include a concise prompt manifest beside the assets.

Add speaker-note provenance such as:

```text
[Sources]
- AI-generated visual · OpenAI ImageGen · <asset path> · prompt recorded in <prompt record>
```

Use meaningful alt text that describes the visual without making unverified claims.

## Crop and Review

- Inspect the asset before embedding.
- Render it at the final slide size and inspect the crop again.
- Replace assets with clipped subjects, ambiguous mechanisms, baked-in text, fake logos, or inconsistent visual style.
- Do not reuse the same generated image on multiple slides unless it is an intentional background system.
