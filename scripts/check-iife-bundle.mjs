// Throwaway sanity check: proves the IIFE global bundle works the way a
// classic (non-module) <script> tag would in a browser - top-level `var`
// attaches to the global/window object. `import()`-ing the file as an ES
// module does NOT reproduce this (module scoping hides top-level `var` from
// the global object), so we execute it with vm.runInThisContext instead,
// which is the closest Node equivalent to a classic script tag.
//
// Not part of the pure engine or the test suite - run manually or via
// `npm run check:iife`.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import vm from "node:vm";

const __dirname = dirname(fileURLToPath(import.meta.url));
const bundlePath = join(__dirname, "..", "dist", "spreadsheet-health-check.global.js");
const fixturePath = join(__dirname, "..", "test", "fixtures", "macros.xlsm");

globalThis.window = globalThis;
const code = readFileSync(bundlePath, "utf8");
vm.runInThisContext(code, { filename: bundlePath });

if (typeof globalThis.window.SpreadsheetHealthCheck?.analyse !== "function") {
  console.error("FAIL: window.SpreadsheetHealthCheck.analyse is not a function");
  process.exit(1);
}

const r = window.SpreadsheetHealthCheck.analyse(new Uint8Array(readFileSync(fixturePath)));
console.log(r.riskBand, r.findings.map((f) => f.id));
console.log("score:", r.score, "verdict:", r.verdict);

if (!r.findings.some((f) => f.id === "key-person.macros")) {
  console.error("FAIL: expected key-person.macros finding not present");
  process.exit(1);
}
console.log("OK: IIFE global bundle works and reports the macro finding.");
