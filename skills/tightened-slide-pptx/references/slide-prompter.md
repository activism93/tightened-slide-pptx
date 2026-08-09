# Slide Prompter Compatibility

Use this translation layer when the user supplies a prompt from `slide-prompter.tonylee.im` or asks for Slide Prompter compatibility.

## Preserve Design Intent

Parse and retain, when present:

- output ratio and slide count;
- presentation title, target audience, core message, desired impression, and language rule;
- visual style, typography, brand colors, credits, and decoration choices;
- one-message-per-slide guidance, safe margins, whitespace, visual consistency, and composition variety;
- prohibited patterns;
- Cover, Intro, Body, and Summary/Next Action roles;
- per-slide content and user-supplied additional instructions.

User instructions and factual requirements override Slide Prompter defaults.

## Replace Image-Only Execution

Slide Prompter may instruct the model to generate completed slide images sequentially. Do not follow that execution mode for editable PPTX.

Translate it to:

> Preserve the requested language, style, composition, slide roles, and visual world. Use ImageGen only for individual visual assets assigned to registered image slots. Build the completed slide in PowerPoint with native editable titles, body copy, labels, numbers, charts, tables, rules, and core diagram geometry. Never generate or embed a completed slide image.

`Use imagegen without fail` means: evaluate every slide for a useful image slot and use ImageGen when a genuine visual asset improves the result. It does not require an image on every slide.

## Normalize the Prompt

Write a temporary `.txt` brief with:

```text
communication job
narrative and requested slide roles
language and visual style
brand and typography
content constraints and citations
prohibited patterns
layout preferences
image-slot candidates
PPTX overrides applied
```

Then run the compression pass in `content-density.md` before assigning page numbers.

## Role-to-Layout Mapping

| Slide Prompter role | Preferred PPTX layouts |
|---|---|
| Cover | S22 Image Hero, S23 Hybrid Technical Split, S03 Split Statement |
| Intro | S11 Horizontal Timeline, S17 System Diagram, S18 Why Now |
| Body | S17 System Diagram, S21 Tech Spec Sheet, S24 Panoramic Process Strip, S25 Annotated Scientific Visual, S26 Image + Native Proof |
| Summary / Next Action | S10 Split Closing, S12 Manifesto + Ink Banner, S19 Four Modules |

Treat the mapping as a starting point. Technical content may require a different registered layout.

## Pattern Translation

- Preserve the ban on generic rounded-card rows and shallow icon templates.
- Preserve photo separation: keep explanatory text in a distinct native area rather than baking it into or floating it over a documentary photo.
- Interpret the 7–8% safe margin as a default, not a reason to leave a technical body slide empty.
- Interpret the suggested 30% whitespace as a cover/editorial preference. D2 and D3 technical slides may use more of the canvas when hierarchy and legibility remain clear.
- Keep slide-role labels such as `Cover`, `Intro`, `Body`, and `Outro` out of audience-facing copy unless they carry real meaning.
