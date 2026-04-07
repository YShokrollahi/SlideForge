const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const PptxGenJS = require("pptxgenjs");

const SLIDES_DIR = "./slides";
const OUTPUT_DIR = "./output";
const OUTPUT_FILE = path.join(OUTPUT_DIR, "presentation_editable.pptx");

// PowerPoint wide layout = 13.333 x 7.5 inches
const PPT_W = 13.333;
const PPT_H = 7.5;

// Your HTML slides are fixed at 1280 x 720
const VIEWPORT_W = 1280;
const VIEWPORT_H = 720;

function pxToInX(px) {
  return (px / VIEWPORT_W) * PPT_W;
}

function pxToInY(px) {
  return (px / VIEWPORT_H) * PPT_H;
}

function pxToPt(px) {
  // 96 CSS px ~= 72 pt
  return px * 0.75;
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function safeColor(value, fallback = "000000") {
  if (!value) return fallback;
  const m = String(value).match(/#([0-9a-fA-F]{6})/);
  return m ? m[1].toUpperCase() : fallback;
}

function parseRgbToHex(value, fallback = "000000") {
  if (!value) return fallback;

  const hex = String(value).match(/#([0-9a-fA-F]{6})/);
  if (hex) return hex[1].toUpperCase();

  const rgb = String(value).match(
    /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)/i
  );
  if (!rgb) return fallback;

  const r = Number(rgb[1]).toString(16).padStart(2, "0");
  const g = Number(rgb[2]).toString(16).padStart(2, "0");
  const b = Number(rgb[3]).toString(16).padStart(2, "0");
  return `${r}${g}${b}`.toUpperCase();
}

function parseOpacity(value, fallback = 1) {
  const n = Number(value);
  if (Number.isNaN(n)) return fallback;
  return Math.max(0, Math.min(1, n));
}

function parseBorder(styleString) {
  // Very simple parser for values like: "1px solid rgb(226, 232, 240)"
  if (!styleString || styleString === "none") return null;

  const widthMatch = styleString.match(/(\d+(\.\d+)?)px/);
  const widthPx = widthMatch ? Number(widthMatch[1]) : 1;

  return {
    widthPt: Math.max(0.5, pxToPt(widthPx)),
    color: parseRgbToHex(styleString, "CBD5E1"),
  };
}

async function extractObjects(page) {
  return await page.$$eval("[data-object='true']", (els) => {
    return els.map((el) => {
      const cs = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();

      const firstP = el.querySelector("p");
      const text = (el.innerText || "").trim();

      const directTextStyle = firstP ? window.getComputedStyle(firstP) : cs;

      return {
        objectType: el.getAttribute("data-object-type") || "textbox",
        text,
        html: el.innerHTML,
        x: rect.x,
        y: rect.y,
        w: rect.width,
        h: rect.height,

        backgroundColor: cs.backgroundColor,
        color: directTextStyle.color,
        opacity: cs.opacity,
        borderLeft: cs.borderLeft,
        borderRight: cs.borderRight,
        borderTop: cs.borderTop,
        borderBottom: cs.borderBottom,
        borderRadius: cs.borderRadius,

        fontSize: directTextStyle.fontSize,
        fontWeight: directTextStyle.fontWeight,
        textAlign: directTextStyle.textAlign,
        lineHeight: directTextStyle.lineHeight,
      };
    });
  });
}

function addTextbox(slide, obj) {
  if (!obj.text) return;

  const x = pxToInX(obj.x);
  const y = pxToInY(obj.y);
  const w = pxToInX(obj.w);
  const h = pxToInY(obj.h);

  const fontSizePx = parseFloat(obj.fontSize || "20");
  const fontSizePt = Math.max(8, pxToPt(fontSizePx));

  const fontWeight = String(obj.fontWeight || "400");
  const bold = Number(fontWeight) >= 600 || /bold/i.test(fontWeight);

  let align = "left";
  if (obj.textAlign === "center") align = "center";
  if (obj.textAlign === "right") align = "right";

  const fillColor = parseRgbToHex(obj.backgroundColor, "FFFFFF");
  const bgVisible =
    obj.backgroundColor &&
    !/rgba?\(\s*0,\s*0,\s*0,\s*0\s*\)/i.test(obj.backgroundColor) &&
    obj.backgroundColor !== "transparent";

  slide.addText(obj.text, {
    x,
    y,
    w,
    h,
    fontFace: "Inter",
    fontSize: fontSizePt,
    bold,
    color: parseRgbToHex(obj.color, "0F172A"),
    margin: 0,
    breakLine: false,
    valign: "mid",
    align,
    fit: "shrink",
    fill: bgVisible
      ? {
          color: fillColor,
          transparency: Math.round((1 - parseOpacity(obj.opacity, 1)) * 100),
        }
      : undefined,
    line: undefined,
  });
}

function addShape(slide, obj) {
    const x = pxToInX(obj.x);
    const y = pxToInY(obj.y);
    const w = pxToInX(obj.w);
    const h = pxToInY(obj.h);
  
    const fillColor = parseRgbToHex(obj.backgroundColor, "FFFFFF");
    const transparency = Math.round((1 - parseOpacity(obj.opacity, 1)) * 100);
  
    const border =
      parseBorder(obj.borderTop) ||
      parseBorder(obj.borderRight) ||
      parseBorder(obj.borderBottom) ||
      parseBorder(obj.borderLeft);
  
    slide.addShape("rect", {
      x,
      y,
      w,
      h,
      rectRadius: 0.08,
      fill: {
        color: fillColor,
        transparency,
      },
      line: border
        ? {
            color: border.color,
            width: border.widthPt,
          }
        : {
            color: fillColor,
            transparency: 100,
          },
    });
  }
async function renderEditableSlide(page, pptx, htmlPath) {
  await page.goto(`file://${path.resolve(htmlPath)}`, {
    waitUntil: "networkidle",
  });

  await page.waitForTimeout(800);

  const objects = await extractObjects(page);
  const slide = pptx.addSlide();

  // white background
  slide.background = { color: "FFFFFF" };

  // Draw shapes first, then textboxes on top
  for (const obj of objects) {
    if (!obj.w || !obj.h) continue;

    if (obj.objectType === "shape") {
      addShape(slide, obj);
    }
  }

  for (const obj of objects) {
    if (!obj.w || !obj.h) continue;

    if (obj.objectType === "textbox") {
      addTextbox(slide, obj);
    }
  }

  return { count: objects.length };
}

async function main() {
  ensureDir(OUTPUT_DIR);

  const slides = fs
    .readdirSync(SLIDES_DIR)
    .filter((f) => f.endsWith(".html"))
    .sort();

  if (slides.length === 0) {
    console.error("No HTML slides found in ./slides");
    process.exit(1);
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "SlideForge";
  pptx.company = "SlideForge";
  pptx.subject = "Editable HTML to PowerPoint";
  pptx.title = "Editable Presentation";
  pptx.lang = "en-US";

  const browser = await chromium.launch();
  const page = await browser.newPage({
    viewport: { width: VIEWPORT_W, height: VIEWPORT_H },
    deviceScaleFactor: 1,
  });

  for (const file of slides) {
    const htmlPath = path.join(SLIDES_DIR, file);
    console.log(`Building editable slide from: ${file}`);
    const result = await renderEditableSlide(page, pptx, htmlPath);
    console.log(`  Added ${result.count} objects`);
  }

  await browser.close();

  await pptx.writeFile({ fileName: OUTPUT_FILE });
  console.log(`✅ Editable PPT generated: ${OUTPUT_FILE}`);
}

main().catch((err) => {
  console.error("Build failed:");
  console.error(err);
  process.exit(1);
});