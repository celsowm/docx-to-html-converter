import { defineConfig } from "tsup";

export default defineConfig([
  // --- ESM + CJS (npm) ---
  {
    entry: { index: "src/index.js" },
    format: ["esm", "cjs"],
    dts: { entry: "index.d.ts" },
    sourcemap: true,
    minify: true,
    clean: true,
    target: "es2019",
    external: ["jszip", "fast-xml-parser"],
  },

  // --- IIFE standalone (CDN/demo) ---
  {
    entry: { index: "src/index.js" },
    format: ["iife"],
    globalName: "DocxToHtmlConverter",
    sourcemap: true,
    minify: true,
    clean: false,
    target: "es2019",
    noExternal: ["jszip", "fast-xml-parser"],

    // gera dist/index.iife.js
    outExtension({ format }) {
      return { js: format === "iife" ? ".iife.js" : ".js" };
    },

    // ✅ footer como objeto { js: string }
    footer: {
      js: `
        (function(g){
          var m = g.DocxToHtmlConverter;
          if (m && m.default) { g.DocxToHtmlConverter = m.default; }
        })(typeof window !== 'undefined' ? window : globalThis);
      `
    },
  },
]);
