# Tightened Slide PPTX

A Codex skill for building fully editable native PowerPoint decks with Tightened Slide layout discipline, content-density planning, and bounded ImageGen assets.

The final deck keeps text, shapes, lines, connectors, charts, tables, and diagram nodes editable. PNG renders are used only for QA and are never inserted as completed slides.

## Repository structure

```text
skills/tightened-slide-pptx/   Codex skill
prompts/                       Reusable editable-PPTX prompt
examples/                      Editable PPTX examples and generation source
```

## Example deck

[`examples/euv-wavelength-beyond-platform-economics-ko.pptx`](examples/euv-wavelength-beyond-platform-economics-ko.pptx) is a dense 14-slide Korean semiconductor-engineering deck covering Pattern Transfer, DUV extension, EUV reflective optics, Sn LPP sources, High-NA EUV, and platform economics. Visible text, technical labels, and core schematics remain editable PowerPoint objects; seven bounded, text-free ImageGen visuals add depth to selected slides without flattening the deck. Cited claims and image provenance are recorded in speaker notes.

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

Prompts from [Slide Prompter](https://slide-prompter.tonylee.im/) are also supported. The skill preserves their visual direction and slide roles, compresses thin content into denser technical pages, and translates image-only instructions into editable native PowerPoint layouts with optional bounded ImageGen assets.

## Editability contract

- No completed slide screenshots inside the PPTX
- Native editable textboxes and diagram objects
- Native charts and tables when data must be edited
- Bitmap assets only for genuine visual media
- Two to four useful evidence modules on standard body slides
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
