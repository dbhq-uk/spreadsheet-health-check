import { defineConfig } from "tsup";

export default defineConfig([
  // ESM build for npm/demo consumers.
  {
    entry: { "spreadsheet-health-check": "src/index.ts" },
    format: ["esm"],
    minify: true,
    treeshake: true,
    platform: "browser",
    noExternal: ["xlsx", "fflate"],
    outDir: "dist",
    clean: true,
    outExtension: () => ({ js: ".mjs" }),
  },
  // IIFE global build the website loads with SRI - sets window.SpreadsheetHealthCheck.
  {
    entry: { "spreadsheet-health-check.global": "src/index.ts" },
    format: ["iife"],
    globalName: "SpreadsheetHealthCheck",
    minify: true,
    treeshake: true,
    platform: "browser",
    noExternal: ["xlsx", "fflate"],
    outDir: "dist",
    clean: false,
    outExtension: () => ({ js: ".js" }),
  },
  // Type declarations only, named index.d.ts to match package.json "types".
  {
    entry: { index: "src/index.ts" },
    format: ["esm"],
    dts: { only: true },
    platform: "browser",
    outDir: "dist",
    clean: false,
  },
]);
