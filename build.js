const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const PptxGenJS = require("pptxgenjs");

const SLIDES_DIR = "./slides";
const SCREENSHOT_DIR = "./screenshots";
const OUTPUT_FILE = "./output/presentation.pptx";

async function renderSlide(page, htmlPath, outputPath) {
  await page.goto(`file://${path.resolve(htmlPath)}`, { waitUntil: "networkidle" });
  await page.waitForTimeout(500); // ensure fonts/icons load
  await page.screenshot({ path: outputPath });
}

async function main() {
  if (!fs.existsSync(SCREENSHOT_DIR)) fs.mkdirSync(SCREENSHOT_DIR);
  if (!fs.existsSync("output")) fs.mkdirSync("output");

  const browser = await chromium.launch();
  const page = await browser.newPage({
    viewport: { width: 1280, height: 720 }
  });

  const slides = fs.readdirSync(SLIDES_DIR)
    .filter(f => f.endsWith(".html"))
    .sort();

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  for (const file of slides) {
    const htmlPath = path.join(SLIDES_DIR, file);
    const pngPath = path.join(SCREENSHOT_DIR, file.replace(".html", ".png"));

    console.log(`Rendering: ${file}`);
    await renderSlide(page, htmlPath, pngPath);

    const slide = pptx.addSlide();
    slide.addImage({
      path: pngPath,
      x: 0,
      y: 0,
      w: 13.33,
      h: 7.5
    });
  }

  await browser.close();

  await pptx.writeFile({ fileName: OUTPUT_FILE });
  console.log("✅ PPT generated at:", OUTPUT_FILE);
}

main();