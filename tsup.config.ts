import { defineConfig } from "tsup";

/**
 * Dois builds:
 * 1) npm (esm + cjs) sem embutir deps (jszip, fast-xml-parser ficam como peers)
 * 2) CDN (iife) embutindo tudo (um arquivo só) e expondo a classe direta no window
 */
export default defineConfig([
  // --- Build para npm (ESM + CJS) ---
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

  // --- Build para CDN (IIFE standalone) ---
  {
    entry: { index: "src/index.js" },
    format: ["iife"],
    globalName: "DocxToHtmlConverter", // window.DocxToHtmlConverter
    sourcemap: true,
    minify: true,
    clean: false,
    target: "es2019",
    noExternal: ["jszip", "fast-xml-parser"],

    // força o nome do arquivo ser index.iife.js
    outExtension({ format }) {
      return { js: format === "iife" ? ".iife.js" : ".js" };
    },

    // shim: se por algum motivo vier {default: ...}, mapeia para a classe direta
    footer: {
      iife: `
        (function(g){
          var m = g.DocxToHtmlConverter;
          if (m && m.default) { g.DocxToHtmlConverter = m.default; }
        })(typeof window !== 'undefined' ? window : globalThis);
      `
    },

    // evita tratar default import de peers como namespace
    esbuildOptions(options) {
      // não externalizar ESM automaticamente (mantém shape do default)
      options.esmExternals = false;
    },
  },
]);
