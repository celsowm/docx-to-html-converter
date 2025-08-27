import { defineConfig } from "tsup";

/**
 * 1) Build para npm (esm + cjs): deixa jszip e fast-xml-parser como externos/peers.
 * 2) Build IIFE “standalone” para CDN: embute as dependências num único arquivo.
 */
export default defineConfig([
  // npm (esm + cjs) — SEM bundle das deps
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

  // CDN IIFE — COM bundle das deps (um arquivo só)
  {
    entry: { "index.iife": "src/index.js" },
    format: ["iife"],
    globalName: "DocxToHtmlConverter",
    sourcemap: true,
    minify: true,
    clean: false,
    target: "es2019",
    // força embutir deps:
    noExternal: ["jszip", "fast-xml-parser"],
  },
]);
