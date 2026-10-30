// Render docs/assets/SFT_social_preview.html to docs/assets/SFT_social_preview.png.
//
// Usage: node tools/render_social_preview.mjs [source.html] [output.png]
//
// Opens the page in Playwright's Chromium at 1280 x 640, scale 1, waits for
// its web fonts and writes the PNG. Playwright is an optional tool, not a
// declared prerequisite: install it with `npm install --no-save playwright`
// and `npx playwright install chromium` (tools/README.md). The page loads its
// fonts from Google Fonts, so rendering needs the network.

import { createRequire } from "node:module";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

// require() also finds a global install through NODE_PATH; import does not.
const { chromium } = createRequire(import.meta.url)("playwright");

const [source = "docs/assets/SFT_social_preview.html",
       output = "docs/assets/SFT_social_preview.png"] = process.argv.slice(2);
const size = { width: 1280, height: 640 };

const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: size, deviceScaleFactor: 1 });
  await page.goto(pathToFileURL(resolve(source)).href, { waitUntil: "networkidle" });
  await page.evaluate(() => document.fonts.ready);
  // A family whose web font did not load falls back silently to another one.
  const loaded = await page.evaluate(() => [...document.fonts]
    .filter((face) => face.status === "loaded")
    .map((face) => face.family.replaceAll('"', "")));
  const missing = ["Barlow", "Barlow Condensed", "JetBrains Mono"]
    .filter((family) => !loaded.includes(family));
  if (missing.length > 0) {
    throw new Error(`fonts not loaded: ${missing.join(", ")}`);
  }
  await page.screenshot({ path: output, clip: { x: 0, y: 0, ...size } });
  console.log(`wrote ${output}`);
} finally {
  await browser.close();
}
