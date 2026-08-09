# Artifact-tool Native Authoring Patterns

Use the available Presentation skill's current API documentation as the source of truth.

## Full deck

```js
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const presentation = Presentation.create({
  slideSize: { width: 1600, height: 900 },
});

const slide = presentation.slides.add();
slide.background.fill = "#FAFAF8";
```

## Editable textbox

```js
const title = slide.shapes.add({
  geometry: "textbox",
  name: "slide-title",
  position: { left: 80, top: 96, width: 1200, height: 80 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
title.text = "One clear claim";
title.text.style = {
  fontSize: 48,
  color: "#0A0A0A",
  autoFit: "shrinkText",
  insets: { top: 0, right: 0, bottom: 0, left: 0 },
};
```

## Editable flat shape

```js
slide.shapes.add({
  geometry: "rect",
  name: "evidence-block",
  position: { left: 80, top: 220, width: 360, height: 160 },
  fill: "#F0F0EE",
  line: { style: "solid", fill: "none", width: 0 },
});
```

## Editable connector

Create connectors before nodes when routing matters:

```js
slide.shapes.connect(sourceNode, targetNode, {
  kind: "elbow",
  fromSide: "right",
  toSide: "left",
  line: { style: "solid", fill: "#737373", width: 2 },
  head: { type: "triangle", width: "sm", length: "sm" },
});
```

## Speaker notes

```js
slide.speakerNotes.textFrame.setText(
  "[Sources]\n- User-provided source material.\n- No external sources used.",
);
```

## Export

```js
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(finalPptx);
```

Do not call `slide.images.add()` with a rendered slide image.
