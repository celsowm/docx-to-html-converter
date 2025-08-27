# DOCX to HTML Converter

Convert `.docx` (Microsoft Word) documents directly to **HTML** in the browser.  
Built on top of [JSZip](https://stuk.github.io/jszip/) and [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser), bundled as a standalone library or installable via npm.

---

## Features

- 📄 Parse DOCX files in the browser (no server required).
- 🖋️ Preserves styles: paragraphs, headings, lists, tables, inline formatting.
- 🎨 Extracts page size and margins as CSS `@page` rules (optional).
- 🖼️ Supports embedded images.
- 🔗 Handles hyperlinks (internal and external).
- 🔧 Available as:
  - **ESM/CJS** package (for bundlers and Node.js).
  - **IIFE** build (global `window.DocxToHtmlConverter` for CDN usage).

---

## Installation

### Via npm
```bash
npm install docx-to-html-converter
```

```js
import { DocxToHtmlConverter } from "docx-to-html-converter";

// arrayBuffer from a File, fetch(), etc.
const buffer = await file.arrayBuffer();

const converter = await DocxToHtmlConverter.create(buffer);
const { html, pageStylesCss } = await converter.convert({
  extractPageStyles: true,
});

document.body.innerHTML = html;
```

### Via CDN
```html
<script src="https://cdn.jsdelivr.net/npm/docx-to-html-converter/dist/index.iife.js"></script>
<script>
  async function demo(file) {
    const buf = await file.arrayBuffer();
    const conv = await window.DocxToHtmlConverter.create(buf);
    const { html } = await conv.convert();
    document.querySelector("#preview").innerHTML = html;
  }
</script>
```

---

## Demo

Clone the repository and open [`index.html`](./index.html) in a browser.  
You can drag & drop a `.docx` file and preview the converted HTML side by side.

---

## API

### `DocxToHtmlConverter.create(arrayBuffer, ParserClass?)`
Create a new converter from a DOCX ArrayBuffer.  
Optionally pass a custom XML parser class (defaults to fast-xml-parser).

### `converter.convert(options)`
Convert the document to HTML.

- `options.extractPageStyles` (boolean, default: `true`)  
  Extracts `@page` CSS with margins and page size.

Returns an object:
```ts
{
  html: string;          // HTML content
  pageStyles?: object;   // Page style object
  pageStylesCss?: string // CSS string with @page
}
```

---

## Development

Build outputs:
- `dist/index.js` → ESM
- `dist/index.cjs` → CommonJS
- `dist/index.iife.js` → Standalone browser build (global)

```bash
# build all targets
npm run build

# clean dist/
npm run clean
```

---

## License

[MIT](./LICENSE) © 2025 Your Name
