import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const TMP = `${REPO}/work/euv-deck-dense`;
const OUT = `${REPO}/examples/euv-wavelength-beyond-platform-economics-ko.pptx`;
const ASSET_DIR = path.join(HERE, "assets", "euv");
const W = 1600;
const H = 900;
const ASSET_FILES = {
  hero: "s01-euv-chamber-hero.png",
  absorption: "s08-euv-absorption.png",
  opticalTrain: "s09-euv-optical-train.png",
  mosiMirror: "s10-mosi-multilayer-mirror.png",
  snLpp: "s14-sn-lpp-sequence.png",
  highNa: "s17-high-na-anamorphic-optics.png",
  stage: "s18-stage-dynamics.png",
};
const ASSETS = {};

const C = {
  white: "#FFFFFF",
  paper: "#F7FAFC",
  ink: "#0B1F33",
  navy: "#0A2A43",
  blue: "#1456A0",
  cyan: "#16A9CF",
  cyan2: "#5BD0E6",
  blueSoft: "#EAF1FA",
  cyanSoft: "#DFF6FB",
  grey1: "#EEF2F5",
  grey2: "#D7E0E8",
  grey3: "#758493",
  grey4: "#A8B3BD",
  red: "#D94B4B",
  orange: "#E78C35",
  green: "#2E8B70",
};

const KFONT = "Apple SD Gothic Neo";
const EFONT = "Arial";

async function loadRasterAssets() {
  for (const [key, filename] of Object.entries(ASSET_FILES)) {
    ASSETS[key] = new Uint8Array(await fs.readFile(path.join(ASSET_DIR, filename)));
  }
}

function addRaster(slide, key, x, y, w, h, alt, options = {}) {
  if (!ASSETS[key]) throw new Error(`Missing raster asset: ${key}`);
  return slide.images.add({
    blob: ASSETS[key],
    contentType: "image/png",
    alt: `[AI visual] ${alt}`,
    prompt: `OpenAI ImageGen prompt recorded in examples/assets/euv/prompts.txt (${ASSET_FILES[key]})`,
    fit: options.fit ?? "cover",
    position: { left: x, top: y, width: w, height: h },
    crop: options.crop ?? { left: 0.02, top: 0.02, right: 0.02, bottom: 0.02 },
    geometry: options.rounded === false ? "rect" : "roundRect",
    ...(options.rounded === false ? {} : { borderRadius: options.radius ?? 14 }),
  });
}

function addText(slide, text, x, y, w, h, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: options.size ?? 20,
    bold: options.bold ?? false,
    color: options.color ?? C.ink,
    alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "top",
    autoFit: options.autoFit ?? "shrinkText",
    wrap: options.wrap ?? "square",
    lineSpacing: options.lineSpacing ?? 1.12,
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
    typeface: options.font ?? KFONT,
  };
  return shape;
}

function addRect(slide, x, y, w, h, fill = "none", stroke = "none", sw = 0, name) {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: stroke, width: sw },
  });
}

function addLine(slide, x1, y1, x2, y2, color = C.ink, width = 1, dashed = false) {
  return slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2), top: Math.min(y1, y2),
      width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
      horizontalFlip: x2 < x1, verticalFlip: y2 < y1,
    },
    fill: "none",
    line: { style: dashed ? "dash" : "solid", fill: color, width },
  });
}

function addCircle(slide, cx, cy, r, fill, stroke = "none", sw = 0) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: cx - r, top: cy - r, width: r * 2, height: r * 2 },
    fill,
    line: { style: "solid", fill: stroke, width: sw },
  });
}

function addEllipse(slide, x, y, w, h, fill, stroke = "none", sw = 0) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: stroke, width: sw },
  });
}

function label(slide, text, x, y, w, options = {}) {
  return addText(slide, text.toUpperCase(), x, y, w, options.h ?? 22, {
    size: options.size ?? 11,
    color: options.color ?? C.grey3,
    bold: options.bold ?? true,
    align: options.align ?? "left",
    valign: "middle",
    font: EFONT,
  });
}

function arrow(slide, x, y, w = 42, color = C.blue, size = 30) {
  return addText(slide, "→", x, y, w, 36, { size, color, font: EFONT, bold: true, align: "center", valign: "middle" });
}

function slideBase(pres, no, section, title, kicker = "") {
  const actualNo = pres.slides.items.length + 1;
  const s = pres.slides.add();
  s.background.fill = C.white;
  label(s, section, 64, 34, 650, { color: C.blue, size: 11 });
  label(s, `${String(actualNo).padStart(2, "0")} / 14`, 1390, 34, 146, { align: "right", size: 11 });
  if (kicker) label(s, kicker, 64, 69, 820, { size: 10, color: C.grey3 });
  addText(s, title, 64, kicker ? 94 : 76, 1472, 76, { size: kicker ? 43 : 47, bold: true, lineSpacing: 1.0, name: `slide-${actualNo}-title` });
  return s;
}

function sectionRule(slide, x, y, w, title, right = "", color = C.blue) {
  addLine(slide, x, y, x + w, y, color, 3);
  label(slide, title, x, y + 9, w * 0.62, { color, size: 11 });
  if (right) label(slide, right, x + w * 0.55, y + 9, w * 0.45, { color: C.grey3, size: 10, align: "right" });
}

function card(slide, x, y, w, h, title, body, options = {}) {
  addRect(slide, x, y, w, h, options.fill ?? C.paper, options.stroke ?? C.grey2, options.sw ?? 1);
  if (options.bar !== false) addRect(slide, x, y, 5, h, options.accent ?? C.blue);
  label(slide, title, x + 18, y + 14, w - 36, { color: options.accent ?? C.blue, size: options.titleSize ?? 10 });
  addText(slide, body, x + 18, y + 44, w - 36, h - 56, {
    size: options.size ?? 18, color: options.color ?? C.ink, bold: options.bold ?? false,
    lineSpacing: options.lineSpacing ?? 1.16,
  });
}

function footer(slide, refs) {
  addLine(slide, 64, 844, 1536, 844, C.grey2, 1);
  addText(slide, `Reference · ${refs}`, 64, 853, 1472, 25, { size: 9.5, color: C.grey3, font: EFONT, valign: "middle" });
}

function notes(slide, sources, extra = "") {
  const body = ["[Sources]", ...sources.map((u) => `- ${u}`)];
  if (extra) body.push(`- ${extra}`);
  body.push("- All visible text remains native and editable. Raster illustrations are bounded visual assets, not flattened full-slide compositions.");
  slide.speakerNotes.textFrame.setText(body.join("\n"));
  slide.speakerNotes.setVisible(true);
}

function flowBox(slide, x, y, w, h, top, bottom, accent = C.blue, dark = false) {
  addRect(slide, x, y, w, h, dark ? C.navy : C.paper, dark ? C.navy : C.grey2, 1);
  addRect(slide, x, y, w, 6, accent);
  label(slide, top, x + 14, y + 18, w - 28, { color: dark ? C.cyan2 : accent, size: 10, align: "center" });
  addText(slide, bottom, x + 14, y + 50, w - 28, h - 64, { size: 17, color: dark ? C.white : C.ink, align: "center", valign: "middle", lineSpacing: 1.1 });
}

function stackLayers(slide, x, y, w, layerH, layers) {
  for (let i = 0; i < layers.length; i += 1) {
    const [name, fill, color = C.ink] = layers[i];
    addRect(slide, x, y + i * layerH, w, layerH, fill, C.white, 1);
    addText(slide, name, x + 8, y + i * layerH + 2, w - 16, layerH - 4, { size: 12, bold: true, color, align: "center", valign: "middle" });
  }
}

function photoCrossSection(slide, x, y, w, stage) {
  const h = 42;
  addRect(slide, x, y + h * 2, w, h, C.grey2, C.white, 1);
  addText(slide, "Wafer / Substrate", x + 4, y + h * 2 + 2, w - 8, h - 4, { size: 11, align: "center", valign: "middle" });
  if (stage === "etch" || stage === "strip") {
    addRect(slide, x, y + h, 56, h, C.blueSoft, C.white, 1);
    addRect(slide, x + w - 56, y + h, 56, h, C.blueSoft, C.white, 1);
    addText(slide, "Target", x + 4, y + h + 2, 48, h - 4, { size: 9, align: "center", valign: "middle" });
    addText(slide, "Film", x + w - 52, y + h + 2, 48, h - 4, { size: 9, align: "center", valign: "middle" });
  } else {
    addRect(slide, x, y + h, w, h, C.blueSoft, C.white, 1);
    addText(slide, "Target Film", x + 4, y + h + 2, w - 8, h - 4, { size: 11, align: "center", valign: "middle" });
  }
  if (stage === "coat" || stage === "expose") {
    addRect(slide, x, y, w, h, C.cyanSoft, C.white, 1);
    addText(slide, "PR", x + 4, y + 2, w - 8, h - 4, { size: 11, align: "center", valign: "middle" });
    if (stage === "expose") addRect(slide, x + 56, y, w - 112, h, "#B9ECF5", C.cyan, 1);
  } else if (stage === "develop" || stage === "etch") {
    addRect(slide, x, y, 56, h, C.cyanSoft, C.white, 1);
    addRect(slide, x + w - 56, y, 56, h, C.cyanSoft, C.white, 1);
    addText(slide, "PR", x + 4, y + 2, 48, h - 4, { size: 9, align: "center", valign: "middle" });
    addText(slide, "PR", x + w - 52, y + 2, 48, h - 4, { size: 9, align: "center", valign: "middle" });
  } else {
    addText(slide, "PR stripped", x, y + 9, w, 24, { size: 10, color: C.grey3, align: "center" });
  }
}

function slide01(pres) {
  const s = pres.slides.add();
  s.background.fill = C.white;
  addRaster(s, "hero", 650, 0, 950, 620, "Generic EUV reflective-optics chamber with cyan beam and wafer stage", {
    rounded: false,
    crop: { left: 0.02, top: 0.01, right: 0, bottom: 0.02 },
  });
  addRect(s, 0, 0, 722, 620, C.navy);
  addRect(s, 690, 0, 32, 620, C.navy);
  label(s, "Semiconductor Engineering Presentation", 64, 42, 580, { color: C.cyan2, size: 11 });
  label(s, "01 / 14", 1390, 42, 146, { color: C.white, align: "right", size: 11 });
  addText(s, "EUV,\n파장 너머의 문제", 64, 112, 590, 150, { size: 54, bold: true, color: C.white, lineSpacing: 0.98 });
  addText(s, "Pattern Transfer에서\nPlatform Economics까지", 64, 286, 530, 66, { size: 23, color: C.cyan2, lineSpacing: 1.08 });
  addRect(s, 64, 390, 530, 4, C.cyan);
  label(s, "Key Message", 64, 418, 220, { color: C.cyan2, size: 10 });
  addText(s, "PHOTO는 target film으로 pattern을 전사한다.", 64, 452, 560, 34, { size: 23, bold: true, color: C.white });
  addText(s, "미세화 이후의 승부는 wavelength가 아니라\noptics · PR · source · scanner의 system integration이다.", 64, 502, 570, 68, { size: 18.5, color: C.grey2, lineSpacing: 1.16 });

  label(s, "Storyline", 64, 642, 240, { color: C.blue, size: 10 });
  const items = ["Pattern Transfer", "Resolution Limit", "DUV Extension", "EUV Transition", "Platform Economics"];
  const xs = [64, 366, 668, 970, 1272];
  for (let i = 0; i < 4; i += 1) {
    addLine(s, xs[i] + 238, 729, xs[i + 1] - 14, 729, C.cyan, 3);
    addCircle(s, xs[i + 1] - 14, 729, 4, C.cyan);
  }
  items.forEach((t, i) => {
    addRect(s, xs[i], 680, 238, 98, i === 3 ? C.blue : C.paper, i === 3 ? C.blue : C.grey2, 1);
    label(s, `0${i + 1}`, xs[i] + 14, 694, 36, { color: i === 3 ? C.cyan2 : C.blue, size: 9 });
    addText(s, t, xs[i] + 14, 721, 210, 40, { size: 17, bold: true, color: i === 3 ? C.white : C.ink, align: "center", valign: "middle" });
  });
  addText(s, "Resolution · defect · throughput · cost를 함께 만족할 때 EUV는 양산 platform이 된다.", 64, 800, 1472, 30, { size: 17, color: C.blue, bold: true, align: "center" });
  footer(s, "ASML · Rayleigh criterion; ASML · EUV lithography systems");
  notes(s, [
    "https://www.asml.com/en/technology/lithography-principles/rayleigh-criterion",
    "https://www.asml.com/en/products/euv-lithography-systems",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s01-euv-chamber-hero.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide02(pres) {
  const s = slideBase(pres, 2, "Pattern Transfer", "PHOTO의 본질은 PR 형상을 target film으로 옮기는 것이다");
  sectionRule(s, 64, 170, 970, "Positive PR Sequence", "PR → target film");
  const stepX = [64, 256, 448, 640, 832];
  for (let i = 0; i < 4; i += 1) arrow(s, stepX[i] + 152, 291, 40, C.cyan);
  const steps = ["PR coating", "Exposure", "Develop", "Etch", "PR strip"];
  steps.forEach((t, i) => label(s, t, stepX[i], 203, 168, { align: "center", color: i >= 2 ? C.blue : C.grey3, size: 9.5 }));
  photoCrossSection(s, 64, 247, 168, "coat");
  photoCrossSection(s, 256, 247, 168, "expose");
  photoCrossSection(s, 448, 247, 168, "develop");
  photoCrossSection(s, 640, 247, 168, "etch");
  photoCrossSection(s, 832, 247, 168, "strip");
  addText(s, "노광된 positive PR 영역의 용해도 ↑", 240, 393, 440, 34, { size: 16, color: C.blue, align: "center" });
  addText(s, "PR이 없는 곳의 target film을 제거", 598, 393, 440, 34, { size: 16, color: C.blue, align: "center" });

  sectionRule(s, 1082, 170, 454, "What is the target?", "Not the bulk wafer");
  card(s, 1082, 219, 454, 216, "Target film examples", "Oxide · Nitride · Metal\nPoly-Si · Hard mask", { fill: C.navy, stroke: C.navy, accent: C.cyan, color: C.white, size: 24, lineSpacing: 1.35 });

  sectionRule(s, 64, 486, 1472, "Photo + Resist", "직관적 역할 구분");
  card(s, 64, 535, 454, 212, "Photo", "빛에 반응해 chemical solubility가 바뀐다.", { accent: C.cyan, size: 21 });
  card(s, 552, 535, 454, 212, "Resist", "Develop 이후 etch 등에서 후속 pattern-transfer mask가 된다.", { accent: C.blue, size: 21 });
  card(s, 1040, 535, 496, 212, "Transfer result", "PR에 생긴 opening이 target film의 형상으로 남는다.", { accent: C.navy, size: 21 });
  addText(s, "‘photo-sensitive material + resist material’이라는 엄밀한 화학 정의가 아니라 역할을 이해하기 위한 표현", 64, 776, 1472, 30, { size: 15, color: C.grey3, align: "right" });
  footer(s, "Tokyo Electron · The principle of Semiconductor; ASML · How microchips are made");
  notes(s, [
    "https://www.tel.com/museum/exhibition/process/process2.html?page=2",
    "https://www.asml.com/en/technology/all-about-microchips/how-microchips-are-made",
  ]);
}

function slide03(pres) {
  const s = slideBase(pres, 3, "Resolution Limit", "Resolution은 λ·NA·k1이 함께 정한다", "RAYLEIGH CRITERION");
  addRect(s, 64, 181, 520, 128, C.navy);
  addText(s, "CD ≈ k₁ × λ / NA", 94, 207, 460, 58, { size: 40, bold: true, color: C.white, font: EFONT, align: "center", valign: "middle" });
  addText(s, "λ  wavelength     NA  acceptance of optical system     k₁  process factor", 94, 274, 460, 22, { size: 13, color: C.cyan2, font: EFONT, align: "center" });
  card(s, 64, 340, 520, 108, "Reading the number", "NA ↑ → imaging capability ↑ → resolution limit CD ↓", { accent: C.cyan, size: 21, bold: true });
  card(s, 64, 468, 520, 126, "DUV vs EUV", "DUV: projection lens NA\nEUV: projection mirror optics NA", { accent: C.blue, size: 20 });

  sectionRule(s, 636, 181, 900, "Acceptance Cone", "고각 회절광 = 고공간주파수 정보");
  addLine(s, 1086, 308, 1086, 674, C.grey3, 2);
  addRect(s, 1014, 648, 144, 28, C.grey1, C.grey2, 1);
  label(s, "Wafer Pattern", 1014, 651, 144, { align: "center", size: 9 });
  addCircle(s, 1086, 633, 7, C.blue);
  addLine(s, 1086, 633, 798, 356, C.cyan, 3);
  addLine(s, 1086, 633, 930, 344, C.grey3, 2);
  addLine(s, 1086, 633, 1242, 344, C.grey3, 2);
  addLine(s, 1086, 633, 1374, 356, C.cyan, 3);
  addLine(s, 930, 344, 1242, 344, C.grey3, 3);
  addLine(s, 798, 356, 1374, 356, C.cyan, 4);
  addText(s, "High NA · wide acceptance", 752, 316, 270, 32, { size: 18, color: C.cyan, bold: true, align: "center" });
  addText(s, "Low NA · narrow acceptance", 1160, 316, 300, 32, { size: 18, color: C.grey3, bold: true, align: "center" });
  addText(s, "0th", 1054, 573, 64, 24, { size: 14, color: C.grey3, font: EFONT, align: "center" });
  addText(s, "±1st diffraction orders", 899, 693, 374, 30, { size: 16, color: C.blue, font: EFONT, align: "center" });
  addRect(s, 636, 746, 900, 64, C.cyanSoft);
  addText(s, "NA는 ‘빛을 더 잘 꺾는 정도’가 아니라 optical system이 받아들일 수 있는 angular range다.", 658, 761, 856, 34, { size: 20, bold: true, align: "center", valign: "middle" });
  footer(s, "ASML · The Rayleigh criterion for resolution");
  notes(s, ["https://www.asml.com/en/technology/lithography-principles/rayleigh-criterion"]);
}

function slide04(pres) {
  const s = slideBase(pres, 4, "Resolution Limit", "k₁을 낮출수록 같은 장비를 한계에 가깝게 쓴다", "OPC = LOW-k₁ ENABLER");
  card(s, 64, 178, 438, 166, "k₁ ↓", "같은 λ와 NA에서 더 작은 feature를 요구한다.", { accent: C.blue, size: 22, bold: true });
  card(s, 64, 362, 438, 202, "The cost of low-k₁", "Image contrast · process margin ↓\nDose · focus · mask · PR variation sensitivity ↑", { accent: C.orange, size: 20 });
  addRect(s, 64, 590, 438, 150, C.navy);
  addText(s, "낮은 k₁ = 무조건 좋은 공정\n이 아니라 optical limit에 가까운 공정", 88, 618, 390, 86, { size: 23, color: C.white, bold: true, align: "center", valign: "middle", lineSpacing: 1.15 });

  sectionRule(s, 554, 178, 982, "Mask → Optical Image → Wafer", "보정 전 / 보정 후");
  addLine(s, 832, 380, 890, 380, C.grey2, 2);
  addLine(s, 832, 625, 890, 625, C.blue, 2);
  arrow(s, 838, 355, 46, C.grey3);
  arrow(s, 838, 600, 46, C.blue);
  addRect(s, 580, 277, 252, 204, C.paper, C.grey2, 1);
  label(s, "Literal mask", 598, 294, 216, { align: "center" });
  addRect(s, 630, 342, 54, 72, C.ink);
  addRect(s, 728, 342, 54, 72, C.ink);
  addRect(s, 890, 277, 252, 204, C.paper, C.grey2, 1);
  label(s, "Printed result", 908, 294, 216, { align: "center" });
  addCircle(s, 960, 380, 35, C.grey3);
  addCircle(s, 1070, 380, 35, C.grey3);
  addText(s, "rounded / shortened", 912, 437, 208, 24, { size: 14, color: C.red, align: "center" });

  addRect(s, 580, 522, 252, 204, C.blueSoft, C.blue, 1);
  label(s, "OPC-corrected mask", 598, 539, 216, { color: C.blue, align: "center" });
  addRect(s, 612, 593, 82, 72, C.blue);
  addRect(s, 718, 593, 82, 72, C.blue);
  addRect(s, 602, 614, 208, 30, C.blue);
  addRect(s, 890, 522, 252, 204, C.cyanSoft, C.cyan, 1);
  label(s, "Printed result", 908, 539, 216, { color: C.blue, align: "center" });
  addRect(s, 938, 591, 58, 76, C.blue);
  addRect(s, 1036, 591, 58, 76, C.blue);
  addText(s, "target에 가까운 geometry", 908, 681, 216, 24, { size: 14, color: C.blue, align: "center" });
  card(s, 1190, 522, 346, 204, "OPC", "Diffraction과 imaging error를 미리 계산해 mask geometry를 의도적으로 보정하는 computational lithography", { accent: C.cyan, size: 19 });
  addText(s, "OPC는 ‘mask를 이상하게 그리는 기술’이 아니라 low-k₁ lithography를 가능하게 한 핵심 RET다.", 554, 763, 982, 46, { size: 20, color: C.blue, bold: true, align: "right" });
  footer(s, "ASML · Pushing k1 further; ASML · Computational lithography");
  notes(s, [
    "https://www.asml.com/en/technology/lithography-principles/pushing-k1-further",
    "https://www.asml.com/en/products/computational-lithography",
  ]);
}

function slide05(pres) {
  const s = slideBase(pres, 5, "DUV Extension", "ArF immersion은 193 nm를 유지하면서 NA를 1.35까지 높였다", "ARF = SOURCE, NOT LENS");
  sectionRule(s, 64, 178, 1078, "Optical Path", "Transmissive reticle + refractive lens");
  const xs = [64, 252, 464, 676, 910];
  for (let i = 0; i < 4; i += 1) arrow(s, xs[i] + 150, 330, 42, i === 3 ? C.cyan : C.blue);
  flowBox(s, 64, 246, 150, 190, "SOURCE", "ArF excimer\n193 nm", C.cyan, true);
  flowBox(s, 252, 246, 170, 190, "ILLUMINATION", "Angle · pupil\ncontrol", C.blue);
  flowBox(s, 464, 246, 170, 190, "RETICLE", "Transmissive\npattern", C.blue);
  flowBox(s, 676, 246, 192, 190, "PROJECTION", "Refractive\nlens", C.blue);
  flowBox(s, 910, 246, 232, 190, "LAST GAP", "Water\n↓\nWafer PR", C.cyan, true);
  addText(s, "Immersion water는 projection lens의 마지막 optical element와 wafer PR 사이에 있다.", 64, 472, 1078, 42, { size: 20, color: C.blue, bold: true });

  sectionRule(s, 1192, 178, 344, "Production scale", "Tool-level");
  addRect(s, 1192, 228, 344, 208, C.navy);
  label(s, "Water-based ArF immersion", 1214, 248, 300, { color: C.cyan2, align: "center" });
  addText(s, "NA 1.35", 1214, 292, 300, 50, { size: 42, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "38–40 nm", 1214, 354, 300, 50, { size: 39, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  addText(s, "production resolution", 1214, 408, 300, 22, { size: 13, color: C.grey2, font: EFONT, align: "center" });

  card(s, 64, 568, 464, 178, "ArF", "Argon Fluoride excimer laser source\nλ = 193 nm", { accent: C.cyan, size: 22 });
  card(s, 568, 568, 464, 178, "Immersion", "Water의 refractive index를 이용해 optical-system NA를 높인다.", { accent: C.blue, size: 21 });
  card(s, 1072, 568, 464, 178, "Interpretation", "38–40 nm는 node name이 아니라 tool-level single-patterning resolution이다.", { accent: C.navy, size: 19 });
  footer(s, "ASML · Lenses & mirrors; ASML · Light & lasers; TWINSCAN NXT:2000i / 2050i");
  notes(s, [
    "https://www.asml.com/technology/lithography-principles/lenses-and-mirrors",
    "https://www.asml.com/en/technology/lithography-principles/light-and-lasers",
    "https://www.asml.com/en/products/duv-lithography-systems/twinscan-nxt2000i",
    "https://www.asml.com/en/en/products/duv-lithography-systems/twinscan-nxt2050i",
  ]);
}

function slide06(pres) {
  const s = slideBase(pres, 6, "Resolution Limit", "193 ÷ 13.5 ≈ 14.3, 그러나 실제 resolution은 14배가 아니다", "λ ALONE DOES NOT SET RESOLUTION");
  addRect(s, 64, 184, 1472, 96, C.navy);
  addText(s, "CD ≈ k₁ × λ / NA", 92, 207, 500, 48, { size: 37, bold: true, color: C.white, font: EFONT });
  addText(s, "Wavelength advantage is filtered through NA and k₁.", 680, 216, 820, 32, { size: 22, color: C.cyan2, font: EFONT, align: "right" });

  const cols = [64, 560, 1056];
  const data = [
    ["ArF immersion DUV", "193 nm", "NA ≈ 1.35", "38–40 nm", C.grey3],
    ["Low-NA EUV · NXE", "13.5 nm", "NA = 0.33", "≈ 13 nm", C.blue],
    ["High-NA EUV · EXE", "13.5 nm", "NA = 0.55", "≈ 8 nm", C.cyan],
  ];
  data.forEach(([name, wave, na, res, accent], i) => {
    addRect(s, cols[i], 324, 432, 330, C.paper, C.grey2, 1);
    addRect(s, cols[i], 324, 432, 8, accent);
    label(s, name, cols[i] + 24, 350, 384, { color: accent, size: 11, align: "center" });
    addText(s, wave, cols[i] + 24, 398, 384, 52, { size: 40, bold: true, font: EFONT, align: "center" });
    addText(s, na, cols[i] + 24, 478, 384, 38, { size: 26, color: accent, font: EFONT, align: "center" });
    addLine(s, cols[i] + 64, 538, cols[i] + 368, 538, C.grey2, 1);
    addText(s, res, cols[i] + 24, 564, 384, 50, { size: 38, bold: true, color: accent, font: EFONT, align: "center" });
    label(s, "Production resolution", cols[i] + 24, 620, 384, { align: "center", size: 9 });
  });
  addText(s, "같은 13.5 nm에서도 NA 0.33 → 0.55가 imaging capability를 바꾼다.", 64, 704, 1472, 42, { size: 24, color: C.blue, bold: true, align: "center" });
  addText(s, "비교 수치는 각 ASML platform의 공식 production-resolution 표기이며 공정 node 명칭과 동일하지 않다.", 64, 766, 1472, 28, { size: 15, color: C.grey3, align: "center" });
  footer(s, "ASML · NXT:2050i; NXE:3400C; EXE:5200B");
  notes(s, [
    "https://www.asml.com/en/en/products/duv-lithography-systems/twinscan-nxt2050i",
    "https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe3400c",
    "https://www.asml.com/products/euv-lithography-systems/twinscan-exe-5200b",
  ]);
}

function slide07(pres) {
  const s = slideBase(pres, 7, "DUV Extension", "Multi-patterning은 density를 늘리지만 공정·overlay·비용도 늘린다", "LELE ≠ SELF-ALIGNED SPACER PATTERNING");
  sectionRule(s, 64, 180, 376, "LELE", "Multiple lithography");
  card(s, 64, 230, 376, 198, "Litho → Etch → Litho → Etch", "두 번의 independent exposure를 분해해 pitch를 줄인다.\nOverlay 관리가 핵심 부담이다.", { accent: C.grey3, size: 18 });
  addText(s, "Mask A", 90, 458, 120, 26, { size: 15, color: C.blue, bold: true });
  addText(s, "Mask B", 286, 458, 120, 26, { size: 15, color: C.cyan, bold: true });
  for (let i = 0; i < 4; i += 1) {
    addRect(s, 86 + i * 82, 504, 34, 108, i % 2 === 0 ? C.blue : C.cyan);
  }
  addLine(s, 64, 635, 440, 635, C.grey2, 1);
  addText(s, "Independent exposure 간 overlay opportunity", 64, 650, 376, 54, { size: 18, color: C.orange, bold: true, align: "center" });

  sectionRule(s, 488, 180, 1048, "SADP → SAQP", "Self-aligned spacer sequence");
  const bx = [488, 688, 888, 1088, 1288];
  for (let i = 0; i < 4; i += 1) arrow(s, bx[i] + 154, 344, 36, C.cyan, 26);
  const titles = ["Mandrel", "Conformal spacer", "Anisotropic etch", "Mandrel removal", "Pattern transfer"];
  titles.forEach((t, i) => label(s, t, bx[i], 218, 162, { align: "center", color: i >= 2 ? C.blue : C.grey3, size: 9 }));
  bx.forEach((x) => { addRect(s, x, 404, 162, 20, C.grey2); addText(s, "target", x, 407, 162, 16, { size: 9, color: C.grey3, align: "center" }); });
  addRect(s, bx[0] + 54, 292, 54, 112, C.ink);
  addRect(s, bx[1] + 42, 284, 12, 120, C.cyan);
  addRect(s, bx[1] + 54, 292, 54, 112, C.ink);
  addRect(s, bx[1] + 108, 284, 12, 120, C.cyan);
  addRect(s, bx[2] + 42, 292, 12, 112, C.cyan);
  addRect(s, bx[2] + 54, 292, 54, 112, C.ink);
  addRect(s, bx[2] + 108, 292, 12, 112, C.cyan);
  addRect(s, bx[3] + 42, 292, 12, 112, C.cyan);
  addRect(s, bx[3] + 108, 292, 12, 112, C.cyan);
  addRect(s, bx[4] + 42, 338, 12, 66, C.blue);
  addRect(s, bx[4] + 108, 338, 12, 66, C.blue);
  addText(s, "sacrificial core", bx[0], 444, 162, 28, { size: 13, color: C.grey3, align: "center" });
  addText(s, "sidewall only", bx[2], 444, 162, 28, { size: 13, color: C.cyan, align: "center" });
  addText(s, "density ≈ 2×", bx[4], 444, 162, 28, { size: 15, color: C.blue, bold: true, align: "center" });

  addRect(s, 488, 514, 1048, 98, C.blueSoft);
  label(s, "SAQP", 512, 530, 100, { color: C.blue });
  addText(s, "Spacer-derived structure를 다시 mandrel처럼 사용해 spacer patterning을 한 번 더 수행 → density를 다시 증가", 612, 530, 898, 46, { size: 19, bold: true });
  addText(s, "‘공짜 resolution’이 아니다", 488, 650, 420, 36, { size: 26, color: C.orange, bold: true });
  addText(s, "Deposition · etch · mandrel removal · cut/block mask · metrology가 추가되어 cycle time, overlay management, process complexity, cost가 증가한다.", 488, 700, 1048, 70, { size: 20, lineSpacing: 1.2 });
  footer(s, "SPIE · Self-aligned double patterning; imec · SAQP / EUV N5 BEOL integration");
  notes(s, [
    "https://doi.org/10.1117/12.772905",
    "https://www.imec-int.com/en/imec-magazine/imec-magazine-march-2017/first-euv-lithography-high-volume-manufacturing-solution-for-n5-beol",
  ]);
}

function slide08(pres) {
  const s = slideBase(pres, 8, "EUV Transition", "13.5 nm는 기존 DUV lens에 그대로 넣을 수 없다", "PHOTON ENERGY CHANGES THE OPTICAL PLATFORM");
  addRect(s, 64, 180, 430, 254, C.navy);
  label(s, "Photon energy", 88, 205, 382, { color: C.cyan2, align: "center" });
  addText(s, "E = hc / λ", 88, 249, 382, 48, { size: 38, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "13.5 nm  →  ≈ 92 eV", 88, 318, 382, 45, { size: 30, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  addText(s, "high-energy EUV photon", 88, 383, 382, 22, { size: 14, color: C.grey2, font: EFONT, align: "center" });

  sectionRule(s, 530, 180, 470, "Question", "기존 DUV optics 재사용?");
  addText(s, "13.5 nm 빛을\ntransmissive lens에\n넣으면 되는가?", 530, 232, 290, 132, { size: 28, bold: true, lineSpacing: 1.06 });
  addRect(s, 846, 248, 136, 88, C.red);
  addText(s, "NO", 846, 256, 136, 72, { size: 44, bold: true, color: C.white, font: EFONT, align: "center", valign: "middle" });
  addText(s, "공기와 대부분의 optical material이 EUV를 강하게 흡수한다.", 530, 379, 470, 50, { size: 17.5, color: C.blue, bold: true });

  addRaster(s, "absorption", 1042, 180, 494, 254, "EUV photon absorption and electron cascade contrasted with multilayer reflection", {
    crop: { left: 0.05, top: 0.08, right: 0.05, bottom: 0.08 },
  });
  label(s, "Energy transfer · conceptual", 1070, 397, 438, { color: C.white, align: "right", size: 9 });

  sectionRule(s, 64, 482, 1472, "What changes?", "Vacuum + reflective optics");
  const bx = [64, 408, 752, 1096];
  for (let i = 0; i < 3; i += 1) arrow(s, bx[i] + 280, 614, 52, C.cyan);
  flowBox(s, bx[0], 535, 280, 160, "AIR", "Strong absorption\n→ high vacuum", C.red);
  flowBox(s, bx[1], 535, 280, 160, "TRANSMISSION", "Lens / transmissive mask\n→ unusable", C.red);
  flowBox(s, bx[2], 535, 280, 160, "REFLECTION", "Mo/Si multilayer\nmirror + mask", C.blue, true);
  flowBox(s, bx[3], 535, 280, 160, "SYSTEM", "Entire optical path\nredesigned", C.cyan, true);
  addRect(s, 64, 744, 1472, 70, C.cyanSoft);
  addText(s, "흡수는 energy 소멸이 아니라 electron excitation · ionization · secondary electron · heat · chemical reaction으로의 전환이다.", 88, 762, 1424, 36, { size: 18, align: "center", valign: "middle" });
  footer(s, "ASML · Lenses & mirrors; EUV resist review · 92 eV photon / secondary electrons");
  notes(s, [
    "https://www.asml.com/technology/lithography-principles/lenses-and-mirrors",
    "https://pmc.ncbi.nlm.nih.gov/articles/PMC9049984/",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s08-euv-absorption.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide09(pres) {
  const s = slideBase(pres, 9, "EUV Transition", "EUV는 source부터 wafer까지 전 구간이 reflective optics in vacuum이다", "SYSTEM ARCHITECTURE");
  addRect(s, 64, 178, 1472, 76, C.navy);
  addText(s, "Sn plasma → Collector → Intermediate Focus → Illumination → Reflective mask → Projection → Wafer PR", 88, 199, 1424, 34, { size: 21, bold: true, color: C.white, font: EFONT, align: "center", valign: "middle" });

  addRaster(s, "opticalTrain", 64, 276, 1472, 286, "Panoramic EUV source-to-wafer reflective optical train in a vacuum chamber", {
    crop: { left: 0.01, top: 0.23, right: 0.01, bottom: 0.23 },
  });
  addRect(s, 88, 292, 286, 38, C.navy);
  label(s, "Reflective optical train · not to scale", 102, 300, 258, { color: C.white, size: 9 });

  sectionRule(s, 64, 592, 1472, "Roles are not interchangeable", "각 요소의 역할 분리");
  card(s, 64, 632, 350, 128, "Collector mirror", "Sn plasma의 EUV를 모아 optical system으로 넘긴다.", { accent: C.cyan, size: 16.5 });
  card(s, 438, 632, 350, 128, "Illumination optics", "Reticle의 angle · intensity · pupil condition을 만든다.", { accent: C.blue, size: 16.5 });
  card(s, 812, 632, 350, 128, "Reflective mask", "Mo/Si reflector + absorber pattern이 반사광에 정보를 싣는다.", { accent: C.cyan, size: 16.5 });
  card(s, 1186, 632, 350, 128, "Projection optics", "Mirror system이 mask image를 wafer PR에 축소 투영한다.", { accent: C.blue, size: 16.5 });
  addText(s, "Reflective mask ≠ projection mirror     ·     Collector ≠ illumination optics", 64, 788, 1472, 30, { size: 18, color: C.red, bold: true, font: EFONT, align: "center" });
  footer(s, "ASML · EUV lithography systems; ASML · Lenses & mirrors; imec · High-NA EUVL");
  notes(s, [
    "https://www.asml.com/en/products/euv-lithography-systems",
    "https://www.asml.com/technology/lithography-principles/lenses-and-mirrors",
    "https://www.imec-int.com/en/articles/high-na-euvl-next-major-step-lithography",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s09-euv-optical-train.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide10(pres) {
  const s = slideBase(pres, 10, "EUV Reflective Optics", "Mo/Si는 반사 coating이고, mirror geometry가 image를 만든다", "COATING ≠ MAGNIFICATION");
  sectionRule(s, 64, 180, 666, "Mo/Si multilayer coating", "13.5 nm reflectivity");
  addRect(s, 64, 231, 666, 414, C.paper, C.grey2, 1);
  for (let i = 0; i < 12; i += 1) {
    const y = 258 + i * 26;
    addRect(s, 138, y, 430, 13, i % 2 === 0 ? C.blue : C.cyanSoft, C.white, 0.5);
    addText(s, i % 2 === 0 ? "Mo" : "Si", 584, y - 1, 64, 15, { size: 10, color: i % 2 === 0 ? C.blue : C.grey3, font: EFONT, valign: "middle" });
  }
  addLine(s, 106, 258, 106, 309, C.orange, 2);
  addLine(s, 94, 258, 118, 258, C.orange, 2);
  addLine(s, 94, 309, 118, 309, C.orange, 2);
  addText(s, "1 period", 72, 319, 90, 22, { size: 12, color: C.orange, font: EFONT, align: "center" });
  addText(s, "d = tMo + tSi ≈ 6.9 nm", 138, 588, 430, 32, { size: 23, bold: true, color: C.blue, font: EFONT, align: "center" });
  addText(s, "6.9 nm는 Mo 한 층이 아니라 Mo+Si pair의 반복 두께", 88, 658, 618, 42, { size: 18, color: C.orange, bold: true, align: "center" });

  sectionRule(s, 782, 180, 754, "Mirror figure & optical geometry", "Focus · aberration · magnification");
  addRaster(s, "mosiMirror", 782, 231, 754, 414, "Curved Mo-Si multilayer mirror cutaway with cyan incident and reflected EUV beam", {
    fit: "contain",
    crop: { left: 0, top: 0, right: 0, bottom: 0 },
  });
  addRect(s, 806, 538, 706, 82, C.navy);
  label(s, "Curvature · position · arrangement", 830, 552, 658, { color: C.cyan2, align: "center" });
  addText(s, "Focusing · aberration correction · demagnification", 830, 579, 658, 26, { size: 17.5, color: C.white, font: EFONT, bold: true, align: "center" });
  label(s, "Conceptual cutaway · not to scale", 806, 626, 706, { color: C.grey3, align: "right", size: 9 });
  addRect(s, 782, 674, 754, 78, C.cyanSoft);
  addText(s, "Optical coating period는 wafer circuit pitch와 아무 관련이 없다.", 806, 695, 706, 36, { size: 21, color: C.blue, bold: true, align: "center" });
  footer(s, "Montcalm et al., SPIE 3331 (1998), DOI 10.1117/12.309600; imec · EUV mask stack");
  notes(s, [
    "https://doi.org/10.1117/12.309600",
    "https://www.osti.gov/servlets/purl/310916",
    "https://www.imec-int.com/en/articles/high-na-euvl-next-major-step-lithography",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s10-mosi-multilayer-mirror.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide11(pres) {
  const s = slideBase(pres, 11, "EUV Reflective Optics", "Mo/Si interface의 약한 반사가 같은 phase로 겹쳐 강해진다", "BRAGG-LIKE CONSTRUCTIVE INTERFERENCE");
  sectionRule(s, 64, 178, 930, "Geometry", "θ = glancing angle to multilayer plane");
  for (let i = 0; i < 6; i += 1) addRect(s, 64, 585 + i * 20, 930, 10, i % 2 === 0 ? C.blue : C.cyanSoft);
  addLine(s, 258, 276, 500, 585, C.blue, 4);
  addLine(s, 500, 585, 742, 276, C.blue, 4);
  addLine(s, 382, 276, 616, 605, C.cyan, 3);
  addLine(s, 616, 605, 850, 276, C.cyan, 3);
  addLine(s, 500, 585, 616, 605, C.orange, 2, true);
  addText(s, "wave A", 218, 244, 120, 24, { size: 15, color: C.blue, font: EFONT });
  addText(s, "wave B", 364, 244, 120, 24, { size: 15, color: C.cyan, font: EFONT });
  addText(s, "d sinθ", 500, 502, 112, 24, { size: 15, color: C.orange, font: EFONT, bold: true, align: "center" });
  addText(s, "+ d sinθ", 632, 493, 126, 24, { size: 15, color: C.orange, font: EFONT, bold: true, align: "center" });
  addLine(s, 500, 585, 636, 585, C.grey3, 1, true);
  addText(s, "θ", 459, 550, 36, 24, { size: 18, color: C.blue, font: EFONT, bold: true });
  addText(s, "Path difference = 2d sinθ", 64, 742, 930, 38, { size: 26, bold: true, color: C.blue, font: EFONT, align: "center" });

  sectionRule(s, 1046, 178, 490, "Condition", "Same phase");
  addRect(s, 1046, 228, 490, 142, C.navy);
  addText(s, "mλ = 2d sinθ", 1070, 262, 442, 48, { size: 39, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "θ: beam ↔ multilayer plane", 1070, 322, 442, 25, { size: 14, color: C.cyan2, font: EFONT, align: "center" });
  card(s, 1046, 394, 490, 130, "Angle convention", "Surface-normal angle α를 쓰면\nmλ = 2d cosα", { accent: C.cyan, size: 21 });
  card(s, 1046, 548, 490, 130, "Near-normal intuition", "d ≈ λ / 2\n13.5 / 2 ≈ 6.75 nm", { accent: C.blue, size: 21 });
  addText(s, "실제 ≈ 6.9 nm는 optical constants · absorption · phase shift · interface structure를 반영한 optimized period", 1046, 712, 490, 74, { size: 17, color: C.grey3, lineSpacing: 1.18 });
  footer(s, "Underwood & Barbee, Applied Optics 20 (1981), DOI 10.1364/AO.20.003027; Montcalm et al. (1998)");
  notes(s, [
    "https://doi.org/10.1364/AO.20.003027",
    "https://doi.org/10.1117/12.309600",
    "https://www.osti.gov/servlets/purl/310916",
  ]);
}

function slide12(pres) {
  const s = slideBase(pres, 12, "EUV Reflective Optics", "EUV mirror는 수십 년의 thin-film·interface·optical theory가 만든 기술이다", "HISTORICAL CONTEXT");
  addLine(s, 128, 454, 1472, 454, C.grey2, 4);
  const events = [
    [160, "1972", "Design principle", "Spiller\nLow-loss reflection coatings", C.grey3, true],
    [470, "1980", "Controlled fabrication", "Spiller et al.\nsoft-X-ray mirrors", C.blue, false],
    [780, "1981", "Theory formalized", "Underwood & Barbee\nBragg diffractors", C.cyan, true],
    [1090, "1998", "EUVL performance", "Montcalm et al.\n≈67.5% @ 13.4 nm", C.blue, false],
    [1400, "2001", "Interface engineering", "Bajt et al.\n≈70% @ 13.5 nm", C.cyan, true],
  ];
  events.forEach(([x, year, stage, desc, color, top], i) => {
    addCircle(s, x, 454, 13, color, C.white, 3);
    const y = top ? 230 : 516;
    addLine(s, x, top ? 442 : 466, x, top ? 373 : 516, color, 2);
    addText(s, year, x - 78, y, 156, 38, { size: 28, bold: true, color, font: EFONT, align: "center" });
    label(s, stage, x - 120, y + 46, 240, { color, align: "center", size: 9 });
    addText(s, desc, x - 132, y + 80, 264, 76, { size: 17, bold: i >= 3, align: "center", lineSpacing: 1.15 });
  });
  addRect(s, 64, 720, 1472, 84, C.navy);
  addText(s, "어느 날 갑자기 등장한 거울이 아니라, weak interface reflection을 제조 가능한 multilayer mirror로 바꾼 누적 혁신", 94, 742, 1412, 44, { size: 22, bold: true, color: C.white, align: "center", valign: "middle" });
  footer(s, "Spiller 1972 · 10.1063/1.1654189 | Spiller et al. 1980 · 10.1063/1.91759 | Underwood & Barbee 1981 · 10.1364/AO.20.003027 | Montcalm 1998 · 10.1117/12.309600 | Bajt 2001 · 10.1117/12.450946");
  notes(s, [
    "https://doi.org/10.1063/1.1654189",
    "https://doi.org/10.1063/1.91759",
    "https://doi.org/10.1364/AO.20.003027",
    "https://doi.org/10.1117/12.309600",
    "https://doi.org/10.1117/12.450946",
    "https://www.osti.gov/biblio/802924",
  ], "1972는 현대 EUV lithography용 Mo/Si mirror 완성 시점이 아니라 low-loss multilayer-like reflection design principle의 역사적 선행으로 서술함.");
}

function slide13(pres) {
  const s = slideBase(pres, 13, "EUV Source", "Sn은 photon을 만들고, Mo/Si는 만들어진 photon을 반사한다", "SOURCE FUEL ≠ OPTICAL COATING");
  addRect(s, 64, 182, 1472, 88, C.navy);
  addText(s, "Sn  →  highly ionized plasma  →  13.5 nm photon  →  Mo/Si optics  →  scanner", 88, 206, 1424, 38, { size: 27, bold: true, color: C.white, font: EFONT, align: "center" });
  sectionRule(s, 64, 318, 708, "Sn · source fuel", "Photon generation");
  card(s, 64, 368, 708, 168, "Why Sn?", "Highly ionized plasma가 13.5 nm 근처에 강한 emission을 만들고, Mo/Si reflective band와 정합된다.", { accent: C.cyan, size: 21 });
  card(s, 64, 560, 708, 174, "Historical choice", "초기 후보에는 Xe · Li · Sn 등이 있었다. HVM source power와 conversion-efficiency development를 거치며 Sn LPP가 주류가 됐다.", { accent: C.blue, size: 19 });

  sectionRule(s, 828, 318, 708, "Mo/Si · optical coating", "Photon transport");
  card(s, 828, 368, 708, 168, "What Mo/Si does", "이미 생성된 13.5 nm EUV를 각 mirror surface에서 최대한 보존해 다음 optical element로 전달한다.", { accent: C.blue, size: 21 });
  card(s, 828, 560, 708, 174, "Sn trade-off", "Metal debris · vapor · collector contamination → mass-limited droplet, debris mitigation, collector protection이 source architecture의 핵심이 된다.", { accent: C.orange, size: 19 });
  addText(s, "Sn 효율을 Xe 대비 ‘5–10×’로 단정하지 않는다: 여기서는 검증된 primary source가 뒷받침하는 정성 비교만 사용한다.", 64, 778, 1472, 32, { size: 15, color: C.grey3, align: "center" });
  footer(s, "J. Optics · EUV source material review (2022); JMM · LPP source technology; AIP · Sn debris mitigation");
  notes(s, [
    "https://doi.org/10.1088/2040-8986/ac5a7e",
    "https://doi.org/10.1117/1.JMM.11.2.021111",
    "https://doi.org/10.1063/5.0200896",
    "https://www.imec-int.com/en/articles/high-na-euvl-next-major-step-lithography",
  ]);
}

function slide14(pres) {
  const s = slideBase(pres, 14, "EUV Source", "Droplet 하나는 shaping·plasma·emission을 거쳐 EUV pulse가 된다", "SN LASER-PRODUCED PLASMA");
  addRaster(s, "snLpp", 64, 220, 1120, 486, "Tin laser-produced-plasma sequence from droplet shaping to EUV collection", {
    crop: { left: 0.01, top: 0.02, right: 0.01, bottom: 0.02 },
  });
  const stageNames = ["Sn droplet", "Moving target", "Pre-pulse", "Main CO₂", "Sn plasma", "13.5 nm", "Collector", "Scanner"];
  stageNames.forEach((name, i) => {
    const x = 76 + i * 137;
    addRect(s, x, 238, 128, 44, i >= 4 ? C.navy : C.blue);
    addText(s, `${String(i + 1).padStart(2, "0")}  ${name}`, x + 6, 247, 116, 26, { size: 11.5, bold: true, color: i === 3 || i === 4 ? C.cyan2 : C.white, font: EFONT, align: "center", valign: "middle" });
  });
  label(s, "Conceptual sequence · not to scale", 82, 670, 1068, { color: C.white, align: "right", size: 9 });

  addRect(s, 1216, 220, 320, 486, C.navy);
  label(s, "Operating scale", 1240, 244, 272, { color: C.cyan2, align: "center" });
  addText(s, "≈ 25 µm", 1240, 290, 272, 42, { size: 33, bold: true, color: C.white, font: EFONT, align: "center" });
  label(s, "Tin droplet", 1240, 336, 272, { color: C.grey2, align: "center" });
  addLine(s, 1256, 380, 1496, 380, C.blue, 1);
  addText(s, "≈ 70 m/s", 1240, 404, 272, 42, { size: 33, bold: true, color: C.white, font: EFONT, align: "center" });
  label(s, "Droplet velocity", 1240, 450, 272, { color: C.grey2, align: "center" });
  addLine(s, 1256, 494, 1496, 494, C.blue, 1);
  addText(s, "50,000× / s", 1240, 518, 272, 42, { size: 31, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  label(s, "Laser interactions", 1240, 564, 272, { color: C.grey2, align: "center" });
  addText(s, "moving target", 1240, 626, 272, 30, { size: 20, bold: true, color: C.white, font: EFONT, align: "center" });
  label(s, "Not a continuously lit lamp", 1240, 662, 272, { color: C.cyan2, align: "center", size: 9 });

  addRect(s, 64, 730, 1472, 82, C.cyanSoft);
  addText(s, "Laser energy가 13.5 nm로 직접 변환되는 것이 아니다.", 88, 744, 1424, 27, { size: 21, color: C.blue, bold: true, align: "center" });
  addText(s, "Laser가 Sn을 plasma화하고, highly ionized Sn plasma의 electronic transitions가 EUV photon을 방출한다.", 88, 776, 1424, 25, { size: 17, align: "center" });
  footer(s, "ASML · Light & lasers (25 µm, 70 m/s, 50,000 interactions/s)");
  notes(s, ["https://www.asml.com/en/technology/lithography-principles/light-and-lasers"], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s14-sn-lpp-sequence.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide15(pres) {
  const s = slideBase(pres, 15, "EUV Source", "Power는 dose를 빨리 채우고, stability는 dose를 정확히 맞춘다", "THROUGHPUT ≠ DOSE ACCURACY");
  addRect(s, 64, 184, 704, 100, C.blue);
  addText(s, "POWER", 88, 202, 180, 30, { size: 16, bold: true, color: C.cyan2, font: EFONT });
  addText(s, "Dose를 빨리 채우는 능력", 88, 238, 656, 34, { size: 27, bold: true, color: C.white });
  addRect(s, 832, 184, 704, 100, C.navy);
  addText(s, "STABILITY", 856, 202, 180, 30, { size: 16, bold: true, color: C.cyan2, font: EFONT });
  addText(s, "Dose를 정확히 맞추는 능력", 856, 238, 656, 34, { size: 27, bold: true, color: C.white });

  sectionRule(s, 64, 334, 704, "Average source power ↑", "Exposure time ↓");
  addLine(s, 118, 541, 712, 541, C.grey2, 3);
  addRect(s, 138, 462, 96, 79, C.grey3);
  addRect(s, 314, 430, 96, 111, C.blue);
  addRect(s, 490, 388, 96, 153, C.cyan);
  addText(s, "P₁", 138, 553, 96, 28, { size: 14, color: C.grey3, font: EFONT, align: "center" });
  addText(s, "P₂", 314, 553, 96, 28, { size: 14, color: C.blue, font: EFONT, align: "center" });
  addText(s, "P₃", 490, 553, 96, 28, { size: 14, color: C.cyan, font: EFONT, align: "center" });
  addText(s, "Same target dose", 604, 455, 140, 28, { size: 15, color: C.ink, font: EFONT, align: "center" });
  addText(s, "→ shorter exposure time", 604, 495, 140, 48, { size: 17, color: C.blue, bold: true, font: EFONT, align: "center" });
  card(s, 64, 614, 704, 146, "Process margin", "Throughput을 유지하면서 higher-dose resist/process를 선택할 여지가 생겨 stochastic defect trade-off에 유리할 수 있다.", { accent: C.cyan, size: 18 });

  sectionRule(s, 832, 334, 704, "Dose control", "Pulse-to-pulse + scanner sync");
  const pulses = [84, 116, 92, 132, 104, 74, 121, 96, 110, 88, 126, 102];
  addLine(s, 878, 541, 1492, 541, C.grey2, 2);
  addLine(s, 878, 418, 1492, 418, C.blue, 1, true);
  label(s, "Target", 1410, 394, 80, { color: C.blue, align: "right", size: 9 });
  pulses.forEach((h, i) => addRect(s, 894 + i * 48, 541 - h, 24, h, Math.abs(h - 110) > 22 ? C.orange : C.blue));
  addText(s, "Average power가 같아도 pulse variation · dose control · synchronization이 다르면 dose accuracy는 달라진다.", 860, 574, 648, 58, { size: 18, align: "center", lineSpacing: 1.16 });
  addRect(s, 832, 654, 704, 106, C.cyanSoft);
  addText(s, "Higher source power ≠ automatically lower dose error", 856, 675, 656, 32, { size: 22, bold: true, color: C.red, font: EFONT, align: "center" });
  addText(s, "Roadmap 수치는 official primary source의 시점·맥락과 함께만 사용", 856, 714, 656, 24, { size: 14, color: C.grey3, align: "center" });
  footer(s, "ASML 2025 Annual Report · source-power development context; ASML · Light & lasers");
  notes(s, [
    "https://www.asml.com/en/investors/annual-report/2025/strategy-and-stories",
    "https://www.asml.com/en/technology/lithography-principles/light-and-lasers",
  ], "1000 W는 2025년 4월 demonstration context이며 commercial HVM source로 단정하지 않음. Triple-pulse architecture는 이 슬라이드에서 사용하지 않음.");
}

function slide16(pres) {
  const s = slideBase(pres, 16, "High-NA EUV", "High-NA는 13.5 nm를 유지한 채 NA를 0.33에서 0.55로 높인다", "NXE = 0.33 NA · EXE = 0.55 NA");
  addLine(s, 800, 185, 800, 744, C.grey2, 2);
  sectionRule(s, 64, 185, 672, "Low-NA EUV · NXE", "NA 0.33");
  sectionRule(s, 864, 185, 672, "High-NA EUV · EXE", "NA 0.55");
  addRect(s, 258, 605, 284, 26, C.grey1, C.grey2, 1);
  addRect(s, 1058, 605, 284, 26, C.grey1, C.grey2, 1);
  addCircle(s, 400, 589, 6, C.blue);
  addCircle(s, 1200, 589, 6, C.cyan);
  addLine(s, 400, 589, 286, 334, C.blue, 3);
  addLine(s, 400, 589, 514, 334, C.blue, 3);
  addLine(s, 1200, 589, 934, 334, C.cyan, 3);
  addLine(s, 1200, 589, 1466, 334, C.cyan, 3);
  addLine(s, 286, 334, 514, 334, C.blue, 5);
  addLine(s, 934, 334, 1466, 334, C.cyan, 5);
  addText(s, "0.33", 320, 420, 160, 50, { size: 42, bold: true, color: C.blue, font: EFONT, align: "center" });
  addText(s, "0.55", 1120, 420, 160, 50, { size: 42, bold: true, color: C.cyan, font: EFONT, align: "center" });
  addText(s, "≈ 13 nm", 248, 656, 304, 44, { size: 34, bold: true, color: C.blue, font: EFONT, align: "center" });
  addText(s, "≈ 8 nm", 1048, 656, 304, 44, { size: 34, bold: true, color: C.cyan, font: EFONT, align: "center" });
  label(s, "Production resolution", 248, 704, 304, { align: "center" });
  label(s, "Production resolution", 1048, 704, 304, { align: "center" });
  addRect(s, 64, 765, 1472, 54, C.navy);
  addText(s, "EXE는 NXE 대비 single exposure에서 약 1.7× smaller feature를 print할 수 있다.", 88, 778, 1424, 30, { size: 20, bold: true, color: C.white, align: "center" });
  footer(s, "ASML · NXE:3400C (NA 0.33, 13 nm); EXE:5200B (NA 0.55, 8 nm, 1.7× smaller feature)");
  notes(s, [
    "https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe3400c",
    "https://www.asml.com/products/euv-lithography-systems/twinscan-exe-5200b",
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
  ], "Single exposure는 critical pattern을 multiple lithography decomposition으로 나누지 않고 한 번의 exposure에서 구현하는 resolution capability라는 의미로 사용함.");
}

function slide17(pres) {
  const s = slideBase(pres, 17, "High-NA EUV", "Anamorphic optics는 reticle angle 문제를 4×/8× projection으로 푼다", "MAGNIFICATION COMES FROM MIRROR GEOMETRY");
  addRect(s, 64, 182, 1472, 70, C.cyanSoft);
  addText(s, "Mo/Si coating = EUV reflectivity     ·     Projection mirror curvature / position / arrangement = focusing + magnification", 88, 202, 1424, 32, { size: 19, bold: true, color: C.blue, font: EFONT, align: "center" });
  addRaster(s, "highNa", 64, 284, 1008, 504, "Low-NA symmetric projection compared with High-NA anamorphic projection and half-field", {
    crop: { left: 0.01, top: 0.01, right: 0.01, bottom: 0.01 },
  });
  addRect(s, 84, 302, 468, 50, C.blue);
  addText(s, "NXE · LOW-NA 0.33 · ISOMORPHIC", 96, 314, 444, 26, { size: 14, bold: true, color: C.white, font: EFONT, align: "center" });
  addRect(s, 584, 302, 468, 50, C.cyan);
  addText(s, "EXE · HIGH-NA 0.55 · ANAMORPHIC", 596, 314, 444, 26, { size: 14, bold: true, color: C.navy, font: EFONT, align: "center" });
  label(s, "Conceptual optics · not to scale", 84, 754, 968, { color: C.grey3, align: "right", size: 9 });

  card(s, 1104, 284, 432, 132, "NXE · 4× / 4×", "양 방향이 같은 isomorphic demagnification", { accent: C.blue, size: 18, bold: true });
  card(s, 1104, 438, 432, 132, "EXE · 4× / 8×", "한 방향을 더 축소해 reticle incidence angle을 관리", { accent: C.cyan, size: 18, bold: true });
  card(s, 1104, 592, 432, 196, "Half-field trade-off", "한 방향 8× demagnification의 대가로 exposure field가 절반이 된다. 동일 wafer에는 더 많은 field와 더 빠른 stage가 필요하다.", { accent: C.orange, size: 17.5 });
  footer(s, "ASML · 5 things you should know about High-NA EUV lithography; ZEISS · High-NA EUV lithography");
  notes(s, [
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
    "https://www.zeiss.com/semiconductor-manufacturing-technology/inspiring-technology/high-na-euv-lithography.html",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s17-high-na-anamorphic-optics.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide18(pres) {
  const s = slideBase(pres, 18, "High-NA EUV", "Half-field는 stage dynamics를 새로운 throughput 병목으로 만든다", "OPTICS → FIELD COUNT → MOTION CONTROL");
  sectionRule(s, 64, 182, 928, "Precision mechatronics", "Half field → more motion");
  addRaster(s, "stage", 64, 226, 928, 462, "Precision reticle and wafer stages with metrology beams and vibration isolation", {
    crop: { left: 0.08, top: 0.06, right: 0.08, bottom: 0.06 },
  });
  addRect(s, 88, 246, 298, 42, C.navy);
  label(s, "Synchronized reticle + wafer stages", 102, 256, 270, { color: C.white, size: 9 });
  addRect(s, 64, 710, 928, 106, C.cyanSoft);
  addRect(s, 92, 735, 118, 48, C.blue);
  addText(s, "FULL FIELD", 98, 747, 106, 24, { size: 12, bold: true, color: C.white, font: EFONT, align: "center" });
  arrow(s, 226, 741, 54, C.cyan, 30);
  addRect(s, 292, 735, 58, 48, C.cyan);
  addRect(s, 358, 735, 58, 48, C.cyan);
  addText(s, "HALF ×2", 292, 786, 124, 20, { size: 10.5, bold: true, color: C.blue, font: EFONT, align: "center" });
  addText(s, "Field count ↑  →  stage speed · settling · overlay control이 throughput을 결정", 454, 739, 514, 52, { size: 18, bold: true, color: C.blue, align: "center", valign: "middle" });

  sectionRule(s, 1044, 182, 492, "EXE:5000 stage scale", "Acceleration");
  addRect(s, 1044, 232, 492, 178, C.navy);
  addText(s, "≈ 8 g", 1070, 268, 190, 54, { size: 44, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "≈ 32 g", 1318, 268, 190, 54, { size: 44, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  label(s, "Wafer stage", 1070, 334, 190, { color: C.grey2, align: "center" });
  label(s, "Reticle stage", 1318, 334, 190, { color: C.grey2, align: "center" });
  card(s, 1044, 442, 492, 142, "The hard part", "수십 g acceleration에서도 imaging · overlay에 필요한 nm-level positioning을 유지해야 한다.", { accent: C.cyan, size: 20 });
  card(s, 1044, 610, 492, 138, "New bottleneck", "Source power가 exposure-time 병목을 줄이면 stage dynamics · overlay · focus · vibration control이 throughput을 제한할 수 있다.", { accent: C.orange, size: 18 });
  addText(s, "High-NA의 resolution은 optics가 만들지만, 생산성은 precision mechatronics가 지킨다.", 1044, 770, 492, 48, { size: 17, color: C.blue, bold: true, align: "center", valign: "middle" });
  footer(s, "ASML · 5 things you should know about High-NA EUV lithography (8g wafer stage, 32g reticle stage)");
  notes(s, ["https://www.asml.com/en/company/stories/2024/5-things-high-na-euv"], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s18-stage-dynamics.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function slide19(pres) {
  const s = slideBase(pres, 19, "Platform Economics", "EUV는 resolution이 아니라 cost-of-patterning이 이길 때 선택된다", "SYSTEM VALUE, NOT WAVELENGTH VALUE");
  const nodes = [
    [182, 270, "Resolution", "single-exposure capability", C.blue],
    [522, 270, "Patterning complexity", "process steps · mask count", C.cyan],
    [1078, 270, "Overlay / defect", "opportunity reduction", C.orange],
    [1418, 270, "Source", "power · stability", C.cyan],
    [1418, 650, "PR", "sensitivity · stochastic", C.blue],
    [1078, 650, "Scanner", "throughput · precision", C.cyan],
    [522, 650, "Cycle time", "queue · metrology · rework", C.orange],
    [182, 650, "Cost", "tool + process + yield", C.blue],
  ];
  nodes.forEach(([x, y]) => addLine(s, 800, 470, x, y, C.grey2, 2));
  addCircle(s, 800, 470, 118, C.navy);
  addText(s, "EUV\nPlatform", 688, 410, 224, 82, { size: 30, bold: true, color: C.white, font: EFONT, align: "center", valign: "middle", lineSpacing: 1.0 });
  addText(s, "Economics", 688, 500, 224, 30, { size: 17, color: C.cyan2, font: EFONT, align: "center" });
  nodes.forEach(([x, y, a, b, color]) => {
    addCircle(s, x, y, 82, C.white, color, 3);
    addText(s, a, x - 70, y - 34, 140, 28, { size: 17, bold: true, color, font: EFONT, align: "center" });
    addText(s, b, x - 68, y + 2, 136, 48, { size: 12, color: C.grey3, font: EFONT, align: "center", valign: "middle", lineSpacing: 1.06 });
  });
  addRect(s, 420, 758, 760, 64, C.navy);
  addText(s, "DUV + multi-patterning이 더 경제적인 layer라면 EUV가 필수는 아니다.", 444, 774, 712, 32, { size: 20, bold: true, color: C.white, align: "center" });
  footer(s, "Synthesis from ASML platform specifications and imec multi-patterning integration context");
  notes(s, [
    "https://www.asml.com/en/products/euv-lithography-systems",
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
    "https://www.imec-int.com/en/imec-magazine/imec-magazine-march-2017/first-euv-lithography-high-volume-manufacturing-solution-for-n5-beol",
  ], "Cost-of-patterning 판단은 source들의 기술·공정 trade-off를 통합한 발표용 synthesis이며 특정 layer에 대한 경제성 수치를 제시하지 않음.");
}

function denseResolution(pres) {
  const s = slideBase(pres, 3, "Resolution Limit", "λ·NA·k₁이 resolution을 정하고 OPC가 low-k₁을 지탱한다", "RAYLEIGH + COMPUTATIONAL LITHOGRAPHY");
  addRect(s, 64, 181, 442, 118, C.navy);
  addText(s, "CD ≈ k₁ × λ / NA", 86, 204, 398, 46, { size: 34, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "λ wavelength   ·   NA acceptance   ·   k₁ process factor", 86, 258, 398, 23, { size: 11.5, color: C.cyan2, font: EFONT, align: "center" });
  card(s, 64, 324, 442, 158, "NA ↑", "더 넓은 angular range가 high-angle diffraction order를 수집한다. Imaging capability ↑, minimum CD ↓", { accent: C.cyan, size: 18.5 });
  card(s, 64, 504, 442, 184, "k₁ ↓", "같은 λ·NA로 더 작은 feature를 요구한다. Image contrast와 process margin은 줄고 dose·focus·mask·PR variation 민감도는 커진다.", { accent: C.orange, size: 17.5 });
  addRect(s, 64, 710, 442, 104, C.cyanSoft);
  addText(s, "낮은 k₁은 ‘더 좋은 숫자’가 아니라\noptical limit에 더 가까운 운전점이다.", 84, 733, 402, 60, { size: 20, bold: true, color: C.blue, align: "center", valign: "middle" });

  sectionRule(s, 550, 181, 438, "Acceptance cone", "NA = system concept");
  addRect(s, 550, 228, 438, 254, C.paper, C.grey2, 1);
  addRect(s, 722, 426, 94, 18, C.grey2);
  label(s, "Wafer", 722, 397, 94, { align: "center", size: 8.5 });
  addCircle(s, 769, 416, 5, C.blue);
  addLine(s, 769, 416, 634, 268, C.cyan, 3);
  addLine(s, 769, 416, 904, 268, C.cyan, 3);
  addLine(s, 769, 416, 704, 266, C.grey3, 2);
  addLine(s, 769, 416, 834, 266, C.grey3, 2);
  addLine(s, 634, 268, 904, 268, C.cyan, 4);
  addLine(s, 704, 281, 834, 281, C.grey3, 3);
  label(s, "HIGH NA", 584, 239, 120, { color: C.cyan, align: "center" });
  label(s, "LOW NA", 834, 286, 116, { color: C.grey3, align: "center" });
  addText(s, "고각 회절광 = 고공간주파수 정보", 574, 452, 390, 24, { size: 15.5, color: C.blue, bold: true, align: "center" });
  card(s, 550, 504, 438, 184, "DUV / EUV", "DUV에서는 projection lens의 NA, EUV에서는 projection mirror optics의 NA다. ‘빛을 더 잘 꺾는 정도’가 아니다.", { accent: C.blue, size: 17.5 });
  addRect(s, 550, 710, 438, 104, C.navy);
  addText(s, "NA ↑ → resolution 성능 ↑\n→ resolution limit 숫자 ↓", 572, 733, 394, 60, { size: 20, color: C.white, bold: true, align: "center", valign: "middle" });

  sectionRule(s, 1032, 181, 504, "OPC closes the loop", "Mask → image → wafer");
  addRect(s, 1032, 228, 504, 254, C.paper, C.grey2, 1);
  label(s, "Literal mask", 1052, 246, 130, { align: "center" });
  addRect(s, 1082, 292, 32, 58, C.ink);
  addRect(s, 1132, 292, 32, 58, C.ink);
  arrow(s, 1182, 301, 44, C.grey3, 26);
  addCircle(s, 1270, 322, 29, C.grey3);
  addCircle(s, 1342, 322, 29, C.grey3);
  addText(s, "rounded / shortened", 1234, 365, 144, 22, { size: 11.5, color: C.red, font: EFONT, align: "center" });
  label(s, "OPC mask", 1052, 405, 130, { color: C.blue, align: "center" });
  addRect(s, 1066, 435, 50, 24, C.blue);
  addRect(s, 1130, 435, 50, 24, C.blue);
  addRect(s, 1058, 443, 130, 10, C.blue);
  arrow(s, 1188, 420, 40, C.blue, 26);
  addRect(s, 1252, 425, 36, 40, C.blue);
  addRect(s, 1324, 425, 36, 40, C.blue);
  addText(s, "target에 가까움", 1234, 463, 144, 18, { size: 11.5, color: C.blue, font: EFONT, align: "center" });
  card(s, 1032, 504, 504, 184, "Optical Proximity Correction", "Diffraction과 imaging error를 계산해 mask geometry를 의도적으로 보정한다. Low-k₁ lithography를 가능하게 한 핵심 Resolution Enhancement Technique다.", { accent: C.cyan, size: 17 });
  addRect(s, 1032, 710, 504, 104, C.blueSoft);
  addText(s, "OPC는 mask 모양이 아니라\nwafer의 final geometry를 최적화한다.", 1056, 733, 456, 60, { size: 20, bold: true, color: C.blue, align: "center", valign: "middle" });
  footer(s, "ASML · Rayleigh criterion; Pushing k1 further; Computational lithography");
  notes(s, [
    "https://www.asml.com/en/technology/lithography-principles/rayleigh-criterion",
    "https://www.asml.com/en/technology/lithography-principles/pushing-k1-further",
    "https://www.asml.com/en/products/computational-lithography",
  ]);
}

function denseDuv(pres) {
  const s = slideBase(pres, 4, "DUV Extension", "ArF immersion은 NA를 높였지만 wavelength만으로 resolution이 정해지지 않는다", "193 NM · NA 1.35 · TOOL-LEVEL RESOLUTION");
  sectionRule(s, 64, 176, 1104, "ArF immersion optical path", "Water is in the last gap");
  const xs = [64, 248, 432, 616, 800, 984];
  const blocks = [
    ["SOURCE", "ArF excimer\n193 nm", C.cyan, true],
    ["ILLUMINATION", "Angle · pupil\ncontrol", C.blue, false],
    ["RETICLE", "Transmissive\npattern", C.blue, false],
    ["PROJECTION", "Refractive\nlens", C.blue, false],
    ["LAST GAP", "Immersion\nwater", C.cyan, true],
    ["IMAGE", "Wafer PR", C.blue, false],
  ];
  blocks.forEach(([a, b, color, dark], i) => {
    flowBox(s, xs[i], 226, 154, 160, a, b, color, dark);
    if (i < blocks.length - 1) arrow(s, xs[i] + 151, 286, 33, i === 3 ? C.cyan : C.blue, 25);
  });
  addText(s, "Immersion water는 projection lens의 마지막 optical element와 wafer PR 사이에 있다.", 64, 405, 1104, 34, { size: 18, color: C.blue, bold: true, align: "center" });
  addRect(s, 1210, 176, 326, 263, C.navy);
  label(s, "Water-based ArF immersion", 1232, 202, 282, { color: C.cyan2, align: "center" });
  addText(s, "NA 1.35", 1232, 246, 282, 48, { size: 38, color: C.white, bold: true, font: EFONT, align: "center" });
  addText(s, "38–40 nm", 1232, 314, 282, 48, { size: 36, color: C.cyan2, bold: true, font: EFONT, align: "center" });
  label(s, "Tool-level single-patterning", 1232, 374, 282, { color: C.grey2, align: "center", size: 9 });

  sectionRule(s, 64, 488, 1472, "Wavelength advantage is filtered through NA and k₁", "193 / 13.5 ≈ 14.3, resolution ≠ 14.3×");
  const cols = [64, 560, 1056];
  const rows = [
    ["ArF immersion DUV", "193 nm", "NA ≈ 1.35", "38–40 nm", C.grey3],
    ["Low-NA EUV · NXE", "13.5 nm", "NA = 0.33", "≈ 13 nm", C.blue],
    ["High-NA EUV · EXE", "13.5 nm", "NA = 0.55", "≈ 8 nm", C.cyan],
  ];
  rows.forEach(([name, wave, na, res, accent], i) => {
    addRect(s, cols[i], 536, 432, 246, C.paper, C.grey2, 1);
    addRect(s, cols[i], 536, 432, 6, accent);
    label(s, name, cols[i] + 20, 556, 392, { color: accent, align: "center", size: 10 });
    addText(s, wave, cols[i] + 20, 592, 190, 42, { size: 30, bold: true, font: EFONT, align: "center" });
    addText(s, na, cols[i] + 222, 592, 190, 42, { size: 22, color: accent, font: EFONT, align: "center", valign: "middle" });
    addLine(s, cols[i] + 42, 652, cols[i] + 390, 652, C.grey2, 1);
    addText(s, res, cols[i] + 20, 674, 392, 46, { size: 34, bold: true, color: accent, font: EFONT, align: "center" });
    label(s, "Production resolution", cols[i] + 20, 732, 392, { align: "center", size: 8.5 });
  });
  addText(s, "Resolution은 λ뿐 아니라 NA와 k₁의 함수이며, 비교 수치는 node name이 아니라 각 platform의 공식 production-resolution 표기다.", 64, 798, 1472, 28, { size: 14.5, color: C.grey3, align: "center" });
  footer(s, "ASML · Lenses & mirrors; NXT:2000i/2050i; NXE:3400C; EXE:5200B");
  notes(s, [
    "https://www.asml.com/technology/lithography-principles/lenses-and-mirrors",
    "https://www.asml.com/en/technology/lithography-principles/light-and-lasers",
    "https://www.asml.com/en/products/duv-lithography-systems/twinscan-nxt2000i",
    "https://www.asml.com/en/en/products/duv-lithography-systems/twinscan-nxt2050i",
    "https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe3400c",
    "https://www.asml.com/products/euv-lithography-systems/twinscan-exe-5200b",
  ]);
}

function denseMirrorPhysics(pres) {
  const s = slideBase(pres, 8, "EUV Reflective Optics", "Mo/Si period와 constructive interference가 13.5 nm reflectivity를 만든다", "COATING PHYSICS + MIRROR GEOMETRY");
  sectionRule(s, 64, 178, 420, "Mo/Si multilayer", "Optical coating period");
  addRect(s, 64, 226, 420, 432, C.paper, C.grey2, 1);
  for (let i = 0; i < 12; i += 1) {
    const y = 254 + i * 24;
    addRect(s, 136, y, 276, 12, i % 2 === 0 ? C.blue : C.cyanSoft, C.white, 0.5);
    addText(s, i % 2 === 0 ? "Mo" : "Si", 416, y - 1, 42, 14, { size: 9, color: i % 2 === 0 ? C.blue : C.grey3, font: EFONT, valign: "middle" });
  }
  addLine(s, 104, 254, 104, 302, C.orange, 2);
  addLine(s, 94, 254, 114, 254, C.orange, 2);
  addLine(s, 94, 302, 114, 302, C.orange, 2);
  addText(s, "1 period", 74, 309, 60, 34, { size: 11, color: C.orange, font: EFONT, align: "center" });
  addText(s, "d = tMo + tSi", 104, 561, 340, 30, { size: 23, bold: true, color: C.blue, font: EFONT, align: "center" });
  addText(s, "≈ 6.9 nm", 104, 598, 340, 38, { size: 31, bold: true, color: C.cyan, font: EFONT, align: "center" });

  sectionRule(s, 526, 178, 502, "Mirror figure", "Coating ≠ magnification");
  addRaster(s, "mosiMirror", 526, 226, 502, 432, "Curved Mo-Si multilayer mirror cutaway with cyan incident and reflected EUV beam", { fit: "contain", crop: { left: 0, top: 0, right: 0, bottom: 0 } });
  addRect(s, 550, 530, 454, 102, C.navy);
  label(s, "Curvature · position · arrangement", 568, 546, 418, { color: C.cyan2, align: "center" });
  addText(s, "Focus · aberration correction\n· demagnification", 568, 575, 418, 45, { size: 17, color: C.white, bold: true, font: EFONT, align: "center", valign: "middle" });

  sectionRule(s, 1070, 178, 466, "Bragg-like reflection", "θ = glancing angle");
  addRect(s, 1070, 226, 466, 432, C.paper, C.grey2, 1);
  for (let i = 0; i < 5; i += 1) addRect(s, 1094, 510 + i * 18, 418, 9, i % 2 === 0 ? C.blue : C.cyanSoft);
  addLine(s, 1150, 288, 1274, 510, C.blue, 3);
  addLine(s, 1274, 510, 1398, 288, C.blue, 3);
  addLine(s, 1234, 288, 1340, 528, C.cyan, 3);
  addLine(s, 1340, 528, 1446, 288, C.cyan, 3);
  addText(s, "d sinθ", 1260, 436, 94, 20, { size: 13, color: C.orange, bold: true, font: EFONT, align: "center" });
  addText(s, "+ d sinθ", 1340, 430, 94, 20, { size: 13, color: C.orange, bold: true, font: EFONT, align: "center" });
  addText(s, "mλ = 2d sinθ", 1100, 596, 406, 30, { size: 26, color: C.blue, bold: true, font: EFONT, align: "center" });
  addText(s, "α = surface-normal angle이면 mλ = 2d cosα", 1100, 628, 406, 22, { size: 12.5, color: C.grey3, font: EFONT, align: "center" });

  const insights = [
    [64, 456, "Interface", "각 interface의 partial reflection이 같은 phase로 겹친다.", C.blue],
    [536, 496, "Near-normal", "d ≈ λ/2 → 13.5/2 ≈ 6.75 nm", C.cyan],
    [1042, 494, "Reality", "≈6.9 nm는 absorption·phase shift·interface를 반영한 optimized period", C.orange],
  ];
  insights.forEach(([x, w, a, b, accent]) => card(s, x, 690, w, 126, a, b, { accent, size: 16.5 }));
  footer(s, "Underwood & Barbee 1981; Montcalm et al. 1998; imec · EUV mask stack");
  notes(s, [
    "https://doi.org/10.1364/AO.20.003027",
    "https://doi.org/10.1117/12.309600",
    "https://www.osti.gov/servlets/purl/310916",
    "https://www.imec-int.com/en/articles/high-na-euvl-next-major-step-lithography",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s10-mosi-multilayer-mirror.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function denseSource(pres) {
  const s = slideBase(pres, 10, "EUV Source", "Sn은 photon을 만들고, LPP repetition이 source throughput을 만든다", "SOURCE FUEL + LASER-PRODUCED PLASMA");
  addRaster(s, "snLpp", 64, 210, 1000, 420, "Tin laser-produced-plasma sequence from droplet shaping to EUV collection", { crop: { left: 0.01, top: 0.02, right: 0.01, bottom: 0.02 } });
  const stages = ["Sn", "moving", "pre-pulse", "main CO₂", "plasma", "13.5 nm", "collector", "scanner"];
  stages.forEach((name, i) => {
    const x = 74 + i * 122;
    addRect(s, x, 226, 112, 40, i >= 4 ? C.navy : C.blue);
    addText(s, `${i + 1}  ${name}`, x + 5, 235, 102, 22, { size: 10.5, bold: true, color: i === 3 || i === 4 ? C.cyan2 : C.white, font: EFONT, align: "center" });
  });
  label(s, "Conceptual sequence · not to scale", 80, 596, 960, { color: C.white, align: "right", size: 8.5 });

  addRect(s, 1104, 210, 432, 420, C.navy);
  label(s, "Role split", 1128, 230, 384, { color: C.cyan2, align: "center" });
  addText(s, "Sn", 1128, 270, 92, 45, { size: 34, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "source fuel\n13.5 nm photon 생성", 1234, 268, 278, 55, { size: 17, color: C.grey2, align: "left", valign: "middle" });
  addLine(s, 1128, 342, 1512, 342, C.blue, 1);
  addText(s, "Mo/Si", 1128, 366, 110, 45, { size: 28, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  addText(s, "optical coating\n생성된 photon을 반사", 1250, 364, 262, 55, { size: 17, color: C.grey2, align: "left", valign: "middle" });
  addLine(s, 1128, 438, 1512, 438, C.blue, 1);
  addText(s, "≈25 µm", 1128, 460, 118, 34, { size: 24, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "≈70 m/s", 1260, 460, 118, 34, { size: 24, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "50k/s", 1392, 460, 118, 34, { size: 24, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  label(s, "droplet", 1128, 501, 118, { color: C.grey2, align: "center", size: 8.5 });
  label(s, "velocity", 1260, 501, 118, { color: C.grey2, align: "center", size: 8.5 });
  label(s, "interactions", 1392, 501, 118, { color: C.grey2, align: "center", size: 8.5 });
  addText(s, "Xe · Li · Sn 후보 중 Sn LPP가 HVM source power / conversion-efficiency development의 주류가 됐다.", 1130, 544, 380, 62, { size: 16, color: C.white, align: "center", valign: "middle" });

  card(s, 64, 660, 470, 156, "Mechanism", "Laser가 13.5 nm로 직접 바뀌는 것이 아니다. Sn을 plasma화하고 highly ionized Sn의 electronic transitions가 EUV를 방출한다.", { accent: C.cyan, size: 17 });
  card(s, 566, 660, 470, 156, "Why repetition matters", "이 source는 켜 둔 lamp가 아니다. Moving target을 만들고 shaping·plasma화하는 과정을 초당 수만 회 반복한다.", { accent: C.blue, size: 17 });
  card(s, 1068, 660, 468, 156, "Sn trade-off", "Metal debris·vapor·collector contamination 때문에 mass-limited droplet, debris mitigation, collector protection이 필요하다.", { accent: C.orange, size: 17 });
  footer(s, "ASML · Light & lasers; J. Optics source-material review; JMM LPP source technology");
  notes(s, [
    "https://www.asml.com/en/technology/lithography-principles/light-and-lasers",
    "https://doi.org/10.1088/2040-8986/ac5a7e",
    "https://doi.org/10.1117/1.JMM.11.2.021111",
    "https://doi.org/10.1063/5.0200896",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s14-sn-lpp-sequence.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function denseHighNa(pres) {
  const s = slideBase(pres, 12, "High-NA EUV", "High-NA는 0.55 NA와 4×/8× anamorphic optics로 8 nm를 겨냥한다", "SAME 13.5 NM · DIFFERENT ACCEPTANCE + PROJECTION");
  addRect(s, 64, 180, 1472, 58, C.cyanSoft);
  addText(s, "Mo/Si coating = reflectivity   ·   projection mirror geometry = focusing + magnification", 88, 195, 1424, 29, { size: 18, bold: true, color: C.blue, font: EFONT, align: "center" });
  addRaster(s, "highNa", 64, 266, 852, 426, "Low-NA symmetric projection compared with High-NA anamorphic projection and half-field", { crop: { left: 0.01, top: 0.01, right: 0.01, bottom: 0.01 } });
  addRect(s, 84, 282, 384, 44, C.blue);
  label(s, "NXE · 0.33 NA · 4×/4×", 98, 292, 356, { color: C.white, align: "center", size: 9.5 });
  addRect(s, 492, 282, 404, 44, C.cyan);
  label(s, "EXE · 0.55 NA · 4×/8×", 506, 292, 376, { color: C.navy, align: "center", size: 9.5 });
  label(s, "Conceptual optics · not to scale", 84, 658, 812, { align: "right", color: C.grey3, size: 8.5 });

  sectionRule(s, 958, 266, 578, "Platform comparison", "Both λ = 13.5 nm");
  const specs = [
    ["NXE · LOW-NA", "0.33", "≈13 nm", "4× / 4×", C.blue],
    ["EXE · HIGH-NA", "0.55", "≈8 nm", "4× / 8×", C.cyan],
  ];
  specs.forEach(([name, na, res, mag, accent], i) => {
    const y = 314 + i * 152;
    addRect(s, 958, y, 578, 132, i === 1 ? C.navy : C.paper, i === 1 ? C.navy : C.grey2, 1);
    label(s, name, 980, y + 15, 166, { color: i === 1 ? C.cyan2 : accent });
    addText(s, na, 980, y + 50, 166, 44, { size: 34, bold: true, color: i === 1 ? C.white : accent, font: EFONT, align: "center" });
    label(s, "NA", 980, y + 98, 166, { color: i === 1 ? C.grey2 : C.grey3, align: "center", size: 8.5 });
    addText(s, res, 1168, y + 48, 152, 42, { size: 29, bold: true, color: i === 1 ? C.cyan2 : accent, font: EFONT, align: "center" });
    label(s, "resolution", 1168, y + 98, 152, { color: i === 1 ? C.grey2 : C.grey3, align: "center", size: 8.5 });
    addText(s, mag, 1342, y + 50, 170, 40, { size: 23, bold: true, color: i === 1 ? C.white : accent, font: EFONT, align: "center" });
    label(s, "demagnification", 1342, y + 98, 170, { color: i === 1 ? C.grey2 : C.grey3, align: "center", size: 8.5 });
  });
  addRect(s, 958, 632, 578, 60, C.cyanSoft);
  addText(s, "Single exposure에서 ≈1.7× smaller feature", 980, 648, 534, 28, { size: 20, bold: true, color: C.blue, align: "center" });

  card(s, 64, 722, 456, 96, "Why anamorphic?", "Reticle incidence angle을 관리한다.", { accent: C.cyan, size: 17.5, bold: true });
  card(s, 546, 722, 456, 96, "Trade-off", "한 방향 8× → half-size exposure field", { accent: C.orange, size: 17.5, bold: true });
  card(s, 1028, 722, 508, 96, "System implication", "더 많은 field → faster stages · tighter control", { accent: C.blue, size: 17.5, bold: true });
  footer(s, "ASML · NXE:3400C; EXE:5200B; 5 things about High-NA EUV; ZEISS · High-NA EUV");
  notes(s, [
    "https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe3400c",
    "https://www.asml.com/products/euv-lithography-systems/twinscan-exe-5200b",
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
    "https://www.zeiss.com/semiconductor-manufacturing-technology/inspiring-technology/high-na-euv-lithography.html",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s17-high-na-anamorphic-optics.png · prompt recorded in examples/assets/euv/prompts.txt");
}

function denseStageEconomics(pres) {
  const s = slideBase(pres, 13, "Platform Economics", "Half-field 이후에는 stage dynamics까지 cost-of-patterning을 결정한다", "OPTICS → FIELD COUNT → MOTION CONTROL → ECONOMICS");
  addRaster(s, "stage", 64, 196, 704, 352, "Precision reticle and wafer stages with metrology beams and vibration isolation", { crop: { left: 0.08, top: 0.06, right: 0.08, bottom: 0.06 } });
  addRect(s, 84, 214, 300, 42, C.navy);
  label(s, "Synchronized reticle + wafer stages", 98, 224, 272, { color: C.white, align: "center", size: 9 });
  label(s, "Conceptual visualization", 84, 512, 660, { color: C.white, align: "right", size: 8.5 });

  sectionRule(s, 808, 196, 728, "EXE:5000 stage scale", "Half-field → more exposures");
  addRect(s, 808, 244, 728, 128, C.navy);
  addText(s, "≈8 g", 834, 270, 188, 44, { size: 38, bold: true, color: C.white, font: EFONT, align: "center" });
  addText(s, "≈32 g", 1054, 270, 188, 44, { size: 38, bold: true, color: C.cyan2, font: EFONT, align: "center" });
  addText(s, "nm-level", 1274, 270, 236, 44, { size: 34, bold: true, color: C.white, font: EFONT, align: "center" });
  label(s, "wafer stage", 834, 324, 188, { color: C.grey2, align: "center" });
  label(s, "reticle stage", 1054, 324, 188, { color: C.grey2, align: "center" });
  label(s, "positioning accuracy", 1274, 324, 236, { color: C.grey2, align: "center" });
  card(s, 808, 398, 344, 150, "New bottleneck", "Source power가 exposure time을 줄이면 stage dynamics·overlay·focus·vibration control이 throughput을 제한할 수 있다.", { accent: C.orange, size: 17 });
  card(s, 1192, 398, 344, 150, "Control problem", "수십 g acceleration에서도 imaging과 overlay에 필요한 positioning accuracy를 유지해야 한다.", { accent: C.cyan, size: 17 });

  sectionRule(s, 64, 584, 1472, "Cost-of-patterning system view", "EUV is selected when the system wins");
  const econ = [
    [64, "Resolution", "single exposure", C.blue],
    [360, "Complexity", "steps · masks", C.cyan],
    [656, "Risk", "overlay · defect", C.orange],
    [952, "Productivity", "throughput · cycle", C.cyan],
    [1248, "Economics", "tool · process · yield", C.blue],
  ];
  econ.forEach(([x, a, b, accent], i) => {
    addRect(s, x, 632, 264, 112, i === 4 ? C.navy : C.paper, i === 4 ? C.navy : C.grey2, 1);
    addRect(s, x, 632, 264, 5, accent);
    label(s, a, x + 16, 650, 232, { color: i === 4 ? C.cyan2 : accent, align: "center" });
    addText(s, b, x + 16, 686, 232, 34, { size: 17.5, bold: true, color: i === 4 ? C.white : C.ink, font: EFONT, align: "center" });
    if (i < econ.length - 1) arrow(s, x + 263, 664, 33, C.grey3, 24);
  });
  addRect(s, 64, 770, 1472, 52, C.cyanSoft);
  addText(s, "DUV + multi-patterning이 더 경제적인 layer라면 EUV가 필수는 아니다. EUV는 전체 cost-of-patterning이 이길 때 선택된다.", 88, 783, 1424, 28, { size: 18, bold: true, color: C.blue, align: "center" });
  footer(s, "ASML · High-NA EUV; imec · EUV HVM integration and multi-patterning context");
  notes(s, [
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
    "https://www.asml.com/en/products/euv-lithography-systems",
    "https://www.imec-int.com/en/imec-magazine/imec-magazine-march-2017/first-euv-lithography-high-volume-manufacturing-solution-for-n5-beol",
  ], "AI-generated visual · OpenAI ImageGen · examples/assets/euv/s18-stage-dynamics.png · prompt recorded in examples/assets/euv/prompts.txt. Cost-of-patterning은 source들의 trade-off를 통합한 synthesis이며 특정 layer의 경제성 수치를 제시하지 않음.");
}

function slide20(pres) {
  const s = slideBase(pres, 20, "Key Takeaways", "EUV의 가치는 13.5 nm photon을 양산 가능한 pattern으로 바꾸는 데 있다", "REFERENCE MAP");
  const takeaways = [
    ["01", "Pattern Transfer", "PHOTO의 output은 PR image가 아니라 target film에 남은 pattern이다."],
    ["02", "Resolution", "λ · NA · k₁이 함께 정하며, DUV는 OPC와 multi-patterning으로 한계를 확장했다."],
    ["03", "EUV Transition", "13.5 nm의 흡수 특성이 vacuum · reflective mask · mirror optics를 요구했다."],
    ["04", "Platform Economics", "Source · optics · PR · scanner가 resolution · defect · throughput · cost를 함께 만족해야 한다."],
  ];
  takeaways.forEach(([n, a, b], i) => {
    const y = 180 + i * 126;
    addRect(s, 64, y, 746, 102, i === 3 ? C.navy : C.paper, i === 3 ? C.navy : C.grey2, 1);
    addText(s, n, 84, y + 22, 58, 46, { size: 33, bold: true, color: i === 3 ? C.cyan2 : C.blue, font: EFONT, align: "center" });
    addText(s, a, 162, y + 14, 260, 30, { size: 19, bold: true, color: i === 3 ? C.white : C.blue, font: EFONT });
    addText(s, b, 162, y + 49, 620, 42, { size: 16, color: i === 3 ? C.grey2 : C.ink, lineSpacing: 1.14 });
  });

  sectionRule(s, 862, 180, 674, "Primary / official references", "Selected");
  const refs = [
    ["Resolution / k₁ / OPC", "ASML · Rayleigh criterion; Pushing k1 further; Computational lithography"],
    ["DUV", "ASML · NXT:2000i / 2050i; Lenses & mirrors; Light & lasers"],
    ["EUV platforms", "ASML · NXE:3400C; EXE:5200B; 5 things about High-NA EUV"],
    ["Mo/Si mirrors", "Spiller 1972, 1980; Underwood & Barbee 1981; Montcalm 1998; Bajt 2001"],
    ["Source / mask / patterning", "SPIE / JMM source papers; imec High-NA EUVL; imec SAQP integration"],
  ];
  refs.forEach(([a, b], i) => {
    const y = 230 + i * 98;
    addLine(s, 862, y + 78, 1536, y + 78, C.grey2, 1);
    label(s, a, 862, y, 250, { color: i < 3 ? C.blue : C.cyan, size: 9.5 });
    addText(s, b, 1124, y, 412, 66, { size: 15, lineSpacing: 1.16 });
  });
  addRect(s, 862, 744, 674, 74, C.cyanSoft);
  addText(s, "수치·역사·spec은 각 슬라이드 notes의 [Sources] URL로 추적 가능", 886, 764, 626, 34, { size: 16, color: C.blue, bold: true, align: "center" });
  footer(s, "Full URLs are embedded in each slide's [Sources] speaker-note block");
  notes(s, [
    "https://www.asml.com/en/technology/lithography-principles/rayleigh-criterion",
    "https://www.asml.com/en/technology/lithography-principles/pushing-k1-further",
    "https://www.asml.com/en/products/computational-lithography",
    "https://www.asml.com/en/products/duv-lithography-systems/twinscan-nxt2000i",
    "https://www.asml.com/en/en/products/duv-lithography-systems/twinscan-nxt2050i",
    "https://www.asml.com/technology/lithography-principles/lenses-and-mirrors",
    "https://www.asml.com/en/technology/lithography-principles/light-and-lasers",
    "https://www.asml.com/en/products/euv-lithography-systems/twinscan-nxe3400c",
    "https://www.asml.com/products/euv-lithography-systems/twinscan-exe-5200b",
    "https://www.asml.com/en/company/stories/2024/5-things-high-na-euv",
    "https://doi.org/10.1063/1.1654189",
    "https://doi.org/10.1063/1.91759",
    "https://doi.org/10.1364/AO.20.003027",
    "https://doi.org/10.1117/12.309600",
    "https://doi.org/10.1117/12.450946",
  ]);
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.rm(`${TMP}/rendered`, { recursive: true, force: true });
  await fs.rm(`${TMP}/layout`, { recursive: true, force: true });
  await fs.rm(`${TMP}/reimported`, { recursive: true, force: true });
  await fs.mkdir(`${TMP}/rendered`, { recursive: true });
  await fs.mkdir(`${TMP}/layout`, { recursive: true });
  await loadRasterAssets();
  const pres = Presentation.create({ slideSize: { width: W, height: H } });
  [slide01, slide02, denseResolution, denseDuv, slide07, slide08, slide09, denseMirrorPhysics,
   slide12, denseSource, slide15, denseHighNa, denseStageEconomics, slide20]
    .forEach((build) => build(pres));

  const imageInspect = await pres.inspect({ kind: "image", maxChars: 20000 });
  const imageRecords = imageInspect.ndjson.trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
  if (imageRecords.length !== Object.keys(ASSET_FILES).length) {
    throw new Error(`Visual contract failed: expected ${Object.keys(ASSET_FILES).length} bounded images, found ${imageRecords.length}`);
  }
  if (imageRecords.some((record) => !String(record.alt ?? "").startsWith("[AI visual]"))) {
    throw new Error("Visual contract failed: unexpected raster image without AI-visual provenance");
  }

  for (const [index, slide] of pres.slides.items.entries()) {
    const n = String(index + 1).padStart(2, "0");
    await writeBlob(`${TMP}/rendered/slide-${n}.png`, await pres.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${TMP}/layout/slide-${n}.json`, await layout.text());
  }
  await writeBlob(`${TMP}/montage.webp`, await pres.export({ format: "webp", montage: true, scale: 1 }));
  const inspect = await pres.inspect({ kind: "slide,textbox,shape,notes", maxChars: 180000 });
  await fs.writeFile(`${TMP}/inspect.ndjson`, inspect.ndjson);
  const pptx = await PresentationFile.exportPptx(pres);
  await pptx.save(OUT);
  await fs.mkdir(`${TMP}/reimported`, { recursive: true });
  const reimported = await PresentationFile.importPptx(await FileBlob.load(OUT));
  for (const [index, slide] of reimported.slides.items.entries()) {
    const n = String(index + 1).padStart(2, "0");
    await writeBlob(`${TMP}/reimported/slide-${n}.png`, await reimported.export({ slide, format: "png", scale: 1 }));
  }
  const reimportInspect = await reimported.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 180000 });
  await fs.writeFile(`${TMP}/reimported-inspect.ndjson`, reimportInspect.ndjson);
  await fs.rm(`${OUT}.inspect.ndjson`, { force: true });
  console.log(OUT);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
