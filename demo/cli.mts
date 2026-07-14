// demo/cli.mts - node demo, not part of the pure engine.
// Imports the built package (dist/) rather than src/ directly: src/ uses
// extensionless relative imports (fine for the bundler and for Vitest, which
// both resolve them), but plain `node --experimental-strip-types` requires
// explicit extensions for ESM resolution. Importing the built bundle also
// means the demo exercises the exact artifact that ships, run `npm run build`
// first.
import { readFileSync } from "node:fs";
import { analyse } from "../dist/spreadsheet-health-check.mjs";
const path = process.argv[2];
if (!path) { console.error("usage: npm run demo -- <file.xlsx>"); process.exit(1); }
const r = analyse(new Uint8Array(readFileSync(path)));
console.log(`${r.riskBand} (${r.score}/100) - ${r.verdict}`);
for (const f of r.findings) console.log(`  [${f.severity}] ${f.title}${f.locations.length ? " - " + f.locations.slice(0,5).join(", ") : ""}`);
