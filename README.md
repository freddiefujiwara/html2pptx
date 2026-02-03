# html2ppt

Convert HTML to PowerPoint (PPTX) easily.

This tool uses [Playwright](https://playwright.dev/) to take a screenshot of your HTML and [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) to put it into a slide.

## Installation

```bash
npm i -g html2ppt
# Install Playwright browser
npx playwright install --with-deps chromium
```

> **Tip:** If you encounter SSL errors during Playwright installation, you can try:
> `NODE_TLS_REJECT_UNAUTHORIZED=0 npx playwright install --with-deps chromium`

## Usage

You can use it with a local HTML file or a URL.

### Local file

```bash
html2ppt ./index.html
```

### URL

```bash
html2ppt https://example.com
```

### Options

- `-o, --out <pptx>`: Change output file name (default: `slide.pptx`)
- `--selector <css>`: Pick a specific part of the page (default: `.slide`). If you want the whole page, use `--selector ""`.
- `--width <n>`: Viewport width (default: `960`)
- `--height <n>`: Viewport height (default: `540`)
- `--scale <n>`: Resolution scale (default: `2`)

## License

MIT
