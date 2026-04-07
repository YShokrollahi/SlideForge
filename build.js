const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const PptxGenJS = require("pptxgenjs");

const SLIDES_DIR = "./slides";
const OUTPUT_DIR = "./output";
const TEMP_DIR = path.join(OUTPUT_DIR, "_temp_assets");
const OUTPUT_FILE = path.join(OUTPUT_DIR, "presentation_editable.pptx");

// PowerPoint wide layout = 13.333 x 7.5 inches
const PPT_W = 13.333;
const PPT_H = 7.5;

// HTML viewport
const VIEWPORT_W = 1280;
const VIEWPORT_H = 720;

function pxToInX(px) {
  return (px / VIEWPORT_W) * PPT_W;
}

function pxToInY(px) {
  return (px / VIEWPORT_H) * PPT_H;
}

function pxToPt(px) {
  return px * 0.75; // 96 CSS px ~= 72 pt
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function cleanDir(dir) {
  if (!fs.existsSync(dir)) return;
  for (const f of fs.readdirSync(dir)) {
    fs.rmSync(path.join(dir, f), { recursive: true, force: true });
  }
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
  if (!styleString || styleString === "none") return null;

  const widthMatch = styleString.match(/(\d+(\.\d+)?)px/);
  const widthPx = widthMatch ? Number(widthMatch[1]) : 1;

  return {
    widthPt: Math.max(0.5, pxToPt(widthPx)),
    color: parseRgbToHex(styleString, "CBD5E1"),
  };
}

function parseRadiusPx(value) {
  if (!value) return 0;
  const m = String(value).match(/(\d+(\.\d+)?)px/);
  return m ? Number(m[1]) : 0;
}

function isTransparentColor(value) {
  if (!value) return true;
  const s = String(value).trim().toLowerCase();
  return (
    s === "transparent" ||
    s === "rgba(0, 0, 0, 0)" ||
    s === "rgba(0,0,0,0)"
  );
}

async function waitForFonts(page) {
  try {
    await page.evaluate(async () => {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
    });
  } catch (e) {
    // ignore
  }
}

async function extractObjects(page) {
  return await page.$$eval("[data-object='true']", (els) => {
    return els.map((el, idx) => {
      const cs = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();

      const firstP = el.querySelector("p");
      const paragraphs = Array.from(el.querySelectorAll(":scope > p")).map((p) => {
        const pcs = window.getComputedStyle(p);
        return {
          text: (p.innerText || "").trim(),
          fontSize: pcs.fontSize,
          fontWeight: pcs.fontWeight,
          color: pcs.color,
          lineHeight: pcs.lineHeight,
          textAlign: pcs.textAlign,
          marginTop: pcs.marginTop,
          marginBottom: pcs.marginBottom,
          letterSpacing: pcs.letterSpacing,
        };
      });

      const directTextStyle = firstP ? window.getComputedStyle(firstP) : cs;
      const iconEl = el.querySelector("i");

      const childDivCount = el.querySelectorAll(":scope > div").length;
      const hasDirectParagraphs = el.querySelectorAll(":scope > p").length > 0;

      return {
        idx,
        objectType: el.getAttribute("data-object-type") || "textbox",
        text: (el.innerText || "").trim(),
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
        letterSpacing: directTextStyle.letterSpacing,

        display: cs.display,
        justifyContent: cs.justifyContent,
        alignItems: cs.alignItems,

        hasIcon: !!iconEl,
        iconClass: iconEl ? iconEl.className : "",
        childDivCount,
        hasDirectParagraphs,
        paragraphs,

        shouldRasterize:
          (el.getAttribute("data-object-type") || "textbox") === "icon" ||
          !!iconEl ||
          childDivCount > 0 ||
          cs.display === "flex",
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

  // Multi-paragraph handling
  if (obj.paragraphs && obj.paragraphs.length > 0) {
    const runs = [];
    obj.paragraphs.forEach((p, i) => {
      if (!p.text) return;

      const fontSizePt = Math.max(8, pxToPt(parseFloat(p.fontSize || "20")));
      const fontWeight = String(p.fontWeight || "400");
      const bold = Number(fontWeight) >= 600 || /bold/i.test(fontWeight);

      runs.push({
        text: p.text,
        options: {
          breakLine: i < obj.paragraphs.length - 1,
          fontFace: "Arial",
          fontSize: fontSizePt,
          bold,
          color: parseRgbToHex(p.color, "0F172A"),
          breakBefore: false,
        },
      });
    });

    let align = "left";
    const ta = obj.paragraphs[0]?.textAlign || obj.textAlign;
    if (ta === "center") align = "center";
    if (ta === "right") align = "right";

    slide.addText(runs, {
      x,
      y,
      w,
      h,
      margin: 0,
      valign: "top",
      align,
      fit: "resize",
      paraSpaceAfterPt: 2,
      lineSpacingMultiple: 1.0,
      fill: undefined,
      line: undefined,
    });

    return;
  }

  const fontSizePx = parseFloat(obj.fontSize || "20");
  const fontSizePt = Math.max(8, pxToPt(fontSizePx));

  const fontWeight = String(obj.fontWeight || "400");
  const bold = Number(fontWeight) >= 600 || /bold/i.test(fontWeight);

  let align = "left";
  if (obj.textAlign === "center") align = "center";
  if (obj.textAlign === "right") align = "right";

  const fillColor = parseRgbToHex(obj.backgroundColor, "FFFFFF");
  const bgVisible =
    obj.backgroundColor && !isTransparentColor(obj.backgroundColor);

  slide.addText(obj.text, {
    x,
    y,
    w,
    h,
    fontFace: "Arial",
    fontSize: fontSizePt,
    bold,
    color: parseRgbToHex(obj.color, "0F172A"),
    margin: 0,
    breakLine: false,
    valign: "top",
    align,
    fit: "resize",
    fill: bgVisible
      ? {
          color: fillColor,
          transparency: Math.round((1 - parseOpacity(obj.opacity, 1)) * 100),
        }
      : undefined,
    line: undefined,
  });
}
function touchesSlideBoundary(obj) {
    return (
      obj.x < 0 ||
      obj.y < 0 ||
      obj.x + obj.w > VIEWPORT_W ||
      obj.y + obj.h > VIEWPORT_H
    );
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
  
    const borderRadiusStr = String(obj.borderRadius || "").trim();
    const radiusPx = parseRadiusPx(borderRadiusStr);
  
    const isEllipseLike =
      borderRadiusStr.includes("%") && borderRadiusStr.includes("50");
  
    const hasMixedCornerRadius =
      /^\d+(\.\d+)?px\s+\d+(\.\d+)?px\s+\d+(\.\d+)?px\s+\d+(\.\d+)?px$/.test(borderRadiusStr);
  
    const fillOpts = isTransparentColor(obj.backgroundColor)
      ? { color: "FFFFFF", transparency: 100 }
      : { color: fillColor, transparency };
  
    const lineOpts = border
      ? {
          color: border.color,
          width: border.widthPt,
        }
      : {
          color: fillColor,
          transparency: 100,
          width: 0,
        };
  
    if (isEllipseLike) {
      slide.addShape("ellipse", {
        x,
        y,
        w,
        h,
        fill: fillOpts,
        line: lineOpts,
      });
      return;
    }
  
    // PowerPoint cannot faithfully match CSS per-corner radius with one normal shape
    // So these should be rasterized elsewhere.
    if (hasMixedCornerRadius) {
      slide.addShape("rect", {
        x,
        y,
        w,
        h,
        fill: fillOpts,
        line: lineOpts,
      });
      return;
    }
  
    if (radiusPx > 0) {
      slide.addShape("roundRect", {
        x,
        y,
        w,
        h,
        fill: fillOpts,
        line: lineOpts,
      });
      return;
    }
  
    slide.addShape("rect", {
      x,
      y,
      w,
      h,
      fill: fillOpts,
      line: lineOpts,
    });
  }

function hasMixedCornerRadius(obj) {
const s = String(obj.borderRadius || "").trim();
return /^\d+(\.\d+)?px\s+\d+(\.\d+)?px\s+\d+(\.\d+)?px\s+\d+(\.\d+)?px$/.test(s);
}

async function screenshotObject(page, obj, outPath) {
  const handle = await page.$$(["[data-object='true']"].join(""));
  const el = handle[obj.idx];
  if (!el) return false;

  await el.screenshot({
    path: outPath,
    omitBackground: true,
  });
  return true;
}

function addImage(slide, imgPath, obj) {
  slide.addImage({
    path: imgPath,
    x: pxToInX(obj.x),
    y: pxToInY(obj.y),
    w: pxToInX(obj.w),
    h: pxToInY(obj.h),
  });
}

function shouldKeepShapeEditable(obj) {
  return (
    obj.objectType === "shape" &&
    !obj.shouldRasterize &&
    obj.childDivCount === 0 &&
    !obj.hasIcon
  );
}

function shouldKeepTextboxEditable(obj) {
  return obj.objectType === "textbox" && !obj.shouldRasterize;
}

async function renderEditableSlide(page, pptx, htmlPath, slideIndex) {
  await page.goto(`file://${path.resolve(htmlPath)}`, {
    waitUntil: "networkidle",
  });

  await waitForFonts(page);
  await page.waitForTimeout(500);

  const objects = await extractObjects(page);
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  // 1) Editable simple shapes
  for (const obj of objects) {
    if (!obj.w || !obj.h) continue;
    if (shouldKeepShapeEditable(obj)) {
      addShape(slide, obj);
    }
  }

  // 2) Rasterize complex objects (icons, flex groups, nested graphics)
  for (const obj of objects) {
    if (!obj.w || !obj.h) continue;

    const borderRadiusStr = String(obj.borderRadius || "").trim();
    const isEllipseLike =
      borderRadiusStr.includes("%") && borderRadiusStr.includes("50");
    
    const needsRaster =
      obj.objectType === "icon" ||
      obj.shouldRasterize ||
      (obj.objectType === "shape" && obj.childDivCount > 0) ||
      (obj.objectType === "shape" && hasMixedCornerRadius(obj)) ||
      (obj.objectType === "shape" && touchesSlideBoundary(obj));
      (obj.objectType === "shape" && touchesSlideBoundary(obj))

    if (!needsRaster) continue;

    const imgPath = path.join(
      TEMP_DIR,
      `slide_${String(slideIndex).padStart(2, "0")}_obj_${String(obj.idx).padStart(3, "0")}.png`
    );

    const ok = await screenshotObject(page, obj, imgPath);
    if (ok) addImage(slide, imgPath, obj);
  }

  // 3) Editable textboxes last so text stays on top
  for (const obj of objects) {
    if (!obj.w || !obj.h) continue;
    if (shouldKeepTextboxEditable(obj)) {
      addTextbox(slide, obj);
    }
  }

  return { count: objects.length };
}

async function main() {
  ensureDir(OUTPUT_DIR);
  ensureDir(TEMP_DIR);
  cleanDir(TEMP_DIR);

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

  const browser = await chromium.launch({
    headless: true,
  });

  const page = await browser.newPage({
    viewport: { width: VIEWPORT_W, height: VIEWPORT_H },
    deviceScaleFactor: 1,
  });

  for (let i = 0; i < slides.length; i++) {
    const file = slides[i];
    const htmlPath = path.join(SLIDES_DIR, file);
    console.log(`Building editable slide from: ${file}`);
    const result = await renderEditableSlide(page, pptx, htmlPath, i + 1);
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