#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const { chromium } = require("playwright");
let PptxGenJS = require("pptxgenjs");
if (PptxGenJS.default) PptxGenJS = PptxGenJS.default;
const { Command } = require("commander");

const isHttpUrl = (s) => {
  try {
    const u = new URL(s);
    return u.protocol === "http:" || u.protocol === "https:";
  } catch {
    return false;
  }
};

const toPageUrl = (input, _fs = fs) => {
  if (isHttpUrl(input)) return input;
  const p = path.resolve(input);
  if (!_fs.existsSync(p)) throw new Error(`HTML file not found: ${p}`);
  return pathToFileURL(p).toString();
};

const renderToPng = async ({ pageUrl, pngPath, selector, width, height, scale, timeout, waitUntil }, _chromium = chromium) => {
  const browser = await _chromium.launch();
  const page = await browser.newPage({ viewport: { width, height }, deviceScaleFactor: scale });

  try {
    await page.goto(pageUrl, { waitUntil: waitUntil || "networkidle", timeout: timeout || 30000 });
    await page.waitForTimeout(800);

    if (selector) {
      const loc = page.locator(selector);
      if ((await loc.count()) === 0) throw new Error(`Selector not found: ${selector}`);
      await loc.first().screenshot({ path: pngPath });
    } else {
      await page.screenshot({ path: pngPath, fullPage: true });
    }
  } finally {
    await browser.close();
  }
};

const formatError = (err) => {
  const message = err.message || String(err);
  const rules = [
    { key: "HTML file not found", msg: "The HTML file does not exist. Please check the file path." },
    { key: "Selector not found", msg: "The CSS selector was not found. Please check your --selector." },
    { test: () => err.name === "TimeoutError", msg: "The page took too long to load. Please try a larger --timeout." },
    { test: () => err.code === "ENOENT", msg: "Could not save the file. Please check the folder path and your permissions." },
    { key: "net::ERR_", msg: "The URL is not valid or the website could not be reached. Please check your input." },
    { key: "missing required argument 'input'", msg: "Please provide an HTML file path or URL." },
    { key: "waitUntil: expected one of", msg: "Invalid wait strategy. Please use one of: load, domcontentloaded, networkidle, commit." },
  ];

  const matched = rules.find((r) => (r.key ? message.includes(r.key) : r.test()));
  if (matched) return `Error: ${matched.msg}`;

  let displayMessage = message.startsWith("error: ") ? message.substring(7) : message;
  displayMessage = displayMessage.charAt(0).toUpperCase() + displayMessage.slice(1);
  return `Error: ${displayMessage}`;
};

const pngToPptx = async ({ pngPath, pptxPath, widescreen = true }, _PptxGenJS = PptxGenJS) => {
  const pptx = new _PptxGenJS();
  if (widescreen) {
    pptx.defineLayout({ name: "WIDE_16x9", width: 13.333, height: 7.5 });
    pptx.layout = "WIDE_16x9";
  }
  pptx.addSlide().addImage({ path: pngPath, x: 0, y: 0, w: 13.333, h: 7.5 });
  await pptx.writeFile({ fileName: pptxPath });
};

const main = async (_toPageUrl = toPageUrl, _renderToPng = renderToPng, _pngToPptx = pngToPptx) => {
  const program = new Command();
  program
    .exitOverride()
    .configureOutput({ writeErr: () => {} })
    .argument("<input>", "HTML file path or URL (http/https)")
    .option("-o, --out <pptx>", "Output pptx path", "slide.pptx")
    .option("--png <png>", "Intermediate png path (default: out with .png)")
    .option("--selector <css>", "CSS selector to capture. Use empty for full-page.", "")
    .option("--width <n>", "Viewport width", (v) => parseInt(v, 10), 960)
    .option("--height <n>", "Viewport height", (v) => parseInt(v, 10), 540)
    .option("--scale <n>", "Device scale factor", (v) => parseInt(v, 10), 2)
    .option("--timeout <n>", "Navigation timeout in ms", (v) => parseInt(v, 10), 30000)
    .option("--wait <strategy>", "Wait strategy: load, domcontentloaded, networkidle, commit", "networkidle")
    .parse(process.argv);

  const opts = program.opts();
  const input = program.args[0];
  const pptxPath = path.resolve(opts.out);
  const pngPath = opts.png ? path.resolve(opts.png) : pptxPath.replace(/\.pptx$/i, "") + ".png";
  const pageUrl = _toPageUrl(input);
  const selector = opts.selector?.trim() || null;

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
  console.log(`OK: ${pptxPath}\n(intermediate) ${pngPath}`);
};

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
