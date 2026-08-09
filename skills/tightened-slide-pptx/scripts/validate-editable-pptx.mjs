#!/usr/bin/env node
import { execFileSync } from "node:child_process";
import fs from "node:fs";

const file = process.argv[2];
if (!file) {
  console.error("Usage: node validate-editable-pptx.mjs <deck.pptx>");
  process.exit(2);
}
if (!fs.existsSync(file)) {
  console.error(`File not found: ${file}`);
  process.exit(2);
}

function unzipText(entry) {
  return execFileSync("unzip", ["-p", file, entry], { encoding: "utf8", maxBuffer: 32 * 1024 * 1024 });
}

let entries;
try {
  entries = execFileSync("unzip", ["-Z1", file], { encoding: "utf8" }).trim().split("\n");
} catch (error) {
  console.error("Unable to inspect PPTX. Install the standard unzip command and retry.");
  process.exit(2);
}

const presentationXml = unzipText("ppt/presentation.xml");
const sizeMatch = presentationXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
if (!sizeMatch) {
  console.error("Could not determine slide size.");
  process.exit(1);
}
const slideWidth = Number(sizeMatch[1]);
const slideHeight = Number(sizeMatch[2]);

const slideEntries = entries
  .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry))
  .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));

const errors = [];
const warnings = [];
const report = [];

for (const entry of slideEntries) {
  const number = Number(entry.match(/slide(\d+)/)[1]);
  const xml = unzipText(entry);
  const nativeShapes = (xml.match(/<p:sp\b/g) ?? []).length;
  const connectors = (xml.match(/<p:cxnSp\b/g) ?? []).length;
  const graphicFrames = (xml.match(/<p:graphicFrame\b/g) ?? []).length;
  const textRuns = (xml.match(/<a:t(?:\s[^>]*)?>/g) ?? []).length;
  const pictures = [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map((match) => match[0]);
  let fullSlidePictures = 0;

  for (const picture of pictures) {
    const off = picture.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
    const ext = picture.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    if (!off || !ext) continue;
    const x = Number(off[1]);
    const y = Number(off[2]);
    const width = Number(ext[1]);
    const height = Number(ext[2]);
    if (x <= slideWidth * 0.05 && y <= slideHeight * 0.05 && width >= slideWidth * 0.9 && height >= slideHeight * 0.9) {
      fullSlidePictures += 1;
    }
  }

  const nativeObjects = nativeShapes + connectors + graphicFrames;
  if (nativeObjects < 5 && textRuns < 3) {
    errors.push(`Slide ${number}: insufficient native editable objects (${nativeObjects} objects, ${textRuns} text runs).`);
  }
  if (fullSlidePictures > 0 && nativeObjects < 10) {
    errors.push(`Slide ${number}: appears to be flattened into a full-slide image.`);
  } else if (fullSlidePictures > 0) {
    warnings.push(`Slide ${number}: contains a full-bleed image. Confirm that it is a genuine photo/illustration, not a rendered slide.`);
  }

  report.push({ slide: number, nativeShapes, connectors, graphicFrames, textRuns, pictures: pictures.length, fullSlidePictures });
}

if (!slideEntries.length) errors.push("No slides found.");

console.log(JSON.stringify({ file, slideSize: [slideWidth, slideHeight], slides: report, warnings }, null, 2));
if (errors.length) {
  console.error("Editable PPTX validation failed:");
  errors.forEach((error) => console.error(`- ${error}`));
  process.exit(1);
}
console.log(`Editable PPTX validation passed: ${slideEntries.length} slide(s).`);
