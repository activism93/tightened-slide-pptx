# Tightened Slide PPTX

A Codex skill for building fully editable native PowerPoint decks with Tightened Slide layout discipline.

The final deck keeps text, shapes, lines, connectors, charts, tables, and diagram nodes editable. PNG renders are used only for QA and are never inserted as completed slides.

## Repository structure

```text
skills/tightened-slide-pptx/   Codex skill
prompts/                       Reusable editable-PPTX prompt
```

## Install

From a GitHub repository:

```bash
npx skills add activism93/tightened-slide-pptx
```

Manual installation:

```bash
mkdir -p ~/.codex/skills
cp -R skills/tightened-slide-pptx ~/.codex/skills/
```

Restart Codex after installation.

## Use

Start the request with:

```text
With $tightened-slide-pptx skill,
```

Then provide the audience, slide count, source content, language, style, and constraints. See `prompts/editable-pptx-base.txt`.

## Editability contract

- No completed slide screenshots inside the PPTX
- Native editable textboxes and diagram objects
- Native charts and tables when data must be edited
- Bitmap assets only for genuine visual media
- Rendered-slide, overflow, and editability validation before delivery

## Requirements

- Codex with the Presentation artifact workflow available
- Node.js
- The standard `unzip` command for the bundled editability validator

## Publish to GitHub

Repository URL: `https://github.com/activism93/tightened-slide-pptx`

To publish a local clone manually:

```bash
git init -b main
git add .
git commit -m "Add editable Tightened Slide PPTX skill"
git remote add origin git@github.com:activism93/tightened-slide-pptx.git
git push -u origin main
```

## License

MIT License. See `LICENSE`.
