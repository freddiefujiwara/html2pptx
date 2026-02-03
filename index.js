#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const { chromium } = require("playwright");
let PptxGenJS = require("pptxgenjs");
if (PptxGenJS.default) PptxGenJS = PptxGenJS.default;
const { Command } = require("commander");

function isHttpUrl(s) {
  try {
    const u = new URL(s);
    return u.protocol === "http:" || u.protocol === "https:";
  } catch {
    return false;
  }
}

function toPageUrl(input, _fs = fs) {
  if (isHttpUrl(input)) return input;

  const p = path.resolve(input);
  if (!_fs.existsSync(p)) {
    throw new Error(`HTML file not found: ${p}`);
  }
  return pathToFileURL(p).toString(); // file://... に変換
}

async function renderToPng({ pageUrl, pngPath, selector, width, height, scale, timeout, waitUntil }, _chromium = chromium) {
  // console.log('DEBUG: _chromium.launch is', _chromium.launch.toString());
  const browser = await _chromium.launch();
  const page = await browser.newPage({
    viewport: { width, height },
    deviceScaleFactor: scale,
  });

  await page.goto(pageUrl, { waitUntil: waitUntil || "networkidle", timeout: timeout || 30000 });
  await page.waitForTimeout(800);

  if (selector) {
    const loc = page.locator(selector);
    const count = await loc.count();
    // console.log('DEBUG: count is', count);
    if (count === 0) {
      await browser.close();
      throw new Error(`Selector not found: ${selector}`);
    }
    await loc.first().screenshot({ path: pngPath });
  } else {
    await page.screenshot({ path: pngPath, fullPage: true });
  }

  await browser.close();
}

function formatError(err) {
  const message = err.message || String(err);

  if (message.includes("HTML file not found")) {
    return "Error: The HTML file does not exist. Please check the file path.";
  }
  if (message.includes("Selector not found")) {
    return "Error: The CSS selector was not found. Please check your --selector.";
  }
  if (err.name === "TimeoutError") {
    return "Error: The page took too long to load. Please try a larger --timeout.";
  }
  if (err.code === "ENOENT") {
    return "Error: Could not save the file. Please check the folder path and your permissions.";
  }
  if (message.includes("net::ERR_")) {
    return "Error: The URL is not valid or the website could not be reached. Please check your input.";
  }
  if (message.includes("missing required argument 'input'")) {
    return "Error: Please provide an HTML file path or URL.";
  }
  if (message.includes("waitUntil: expected one of")) {
    return "Error: Invalid wait strategy. Please use one of: load, domcontentloaded, networkidle, commit.";
  }

  let displayMessage = message;
  if (displayMessage.startsWith("error: ")) {
    displayMessage = displayMessage.substring(7);
    displayMessage = displayMessage.charAt(0).toUpperCase() + displayMessage.slice(1);
  }

  return `Error: ${displayMessage}`;
}

async function pngToPptx({ pngPath, pptxPath, widescreen = true }, _PptxGenJS = PptxGenJS) {
  const pptx = new _PptxGenJS();

  // 16:9 (13.333 x 7.5 inch) を明示定義
  if (widescreen) {
    pptx.defineLayout({ name: "WIDE_16x9", width: 13.333, height: 7.5 });
    pptx.layout = "WIDE_16x9";
  }

  const slide = pptx.addSlide();

  // スライド全面に貼る
  slide.addImage({
    path: pngPath,
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
  });

  await pptx.writeFile({ fileName: pptxPath });
}

async function main(_toPageUrl = toPageUrl, _renderToPng = renderToPng, _pngToPptx = pngToPptx) {
  const program = new Command();
  program.exitOverride();
  program.configureOutput({
    writeErr: () => {},
  });

  program
    .argument("<input>", "HTML file path or URL (http/https)")
    .option("-o, --out <pptx>", "Output pptx path", "slide.pptx")
    .option("--png <png>", "Intermediate png path (default: out with .png)")
    .option("--selector <css>", "CSS selector to capture. Use empty for full-page.", "")
    .option("--width <n>", "Viewport width", (v) => parseInt(v, 10), 960)
    .option("--height <n>", "Viewport height", (v) => parseInt(v, 10), 540)
    .option("--scale <n>", "Device scale factor", (v) => parseInt(v, 10), 2)
    .option("--timeout <n>", "Navigation timeout in ms", (v) => parseInt(v, 10), 30000)
    .option("--wait <strategy>", "Wait strategy: load, domcontentloaded, networkidle, commit", "networkidle");

  program.parse(process.argv);
  const opts = program.opts();
  const input = program.args[0];

  const pptxPath = path.resolve(opts.out);
  const pngPath = opts.png ? path.resolve(opts.png) : pptxPath.replace(/\.pptx$/i, "") + ".png";

  const pageUrl = _toPageUrl(input);

  let selector = opts.selector;
  // --selector "" を full page 扱いにする
  if (selector != null && selector.trim() === "") selector = null;

  await _renderToPng({
    pageUrl,
    pngPath,
    selector,
    width: opts.width,
    height: opts.height,
    scale: opts.scale,
    timeout: opts.timeout,
    waitUntil: opts.wait,
  });

  await _pngToPptx({ pngPath, pptxPath, widescreen: true });

  console.log(`OK: ${pptxPath}`);
  console.log(`(intermediate) ${pngPath}`);
}

module.exports = {
  isHttpUrl,
  toPageUrl,
  formatError,
  renderToPng,
  pngToPptx,
  main
};

/* v8 ignore start */
if (require.main === module) {
  main().catch((e) => {
    if (e.code && e.code.startsWith("commander.")) {
      if (e.code === "commander.helpDisplayed" || e.code === "commander.help" || e.code === "commander.version") {
        return;
      }
    }
    console.error(formatError(e));
    process.exit(1);
  });
}
/* v8 ignore stop */
