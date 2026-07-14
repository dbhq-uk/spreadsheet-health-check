# Spreadsheet Health Check

A free, client-side risk scanner for Excel workbooks. Point it at a
spreadsheet and it scores the business risk hiding inside - key-person
dependency, live errors, hidden fragility and things that won't scale - with
no upload, ever.

Everything runs in the browser (or in Node, via this package). The file
bytes never leave the machine they were opened on: there is no backend, no
network call, and no analytics on the tool. That guarantee is the point of
this repo - the engine is small, pure and auditable.

Live tool: https://dbhq.uk/spreadsheet-health-check/

## What it checks

Eight checks across four categories:

- **Key-person exposure** - single-author workbooks, hidden macros/VBA,
  hidden or very-hidden sheets.
- **Broken today** - live error values (`#REF!`, `#DIV/0!`, ...), circular
  references with iterative calculation silently switched on.
- **Hidden fragility** - inconsistent formulas within a column (the classic
  "someone typed over one cell" bug), external workbook links.
- **Won't scale** - usage approaching Excel's row/column/reference limits.

Each check returns a `Finding` (or nothing). Findings are scored, grouped
into the four categories, and rolled up into a risk band:

| Band | Score |
| --- | --- |
| Low | 0-24 |
| Moderate | 25-49 |
| High | 50-74 |
| Critical | 75-100 |

## Usage

```ts
import { analyse } from "@dbhq/spreadsheet-health-check";
import { readFileSync } from "node:fs";

const report = analyse(new Uint8Array(readFileSync("workbook.xlsx")));

console.log(report.riskBand); // "Low" | "Moderate" | "High" | "Critical"
console.log(report.score);    // 0-100
console.log(report.verdict);  // one-line plain-English summary
console.log(report.findings); // Finding[]
```

`analyse` takes the raw bytes of an `.xlsx`, `.xlsm`, `.xls` or `.xlsb` file
and returns a `HealthReport` synchronously - no async, no I/O inside the
engine itself.

In the browser, load the self-contained global bundle (SheetJS and fflate
are bundled in, no CDN, no WASM) and call the same function:

```html
<script src="/spreadsheet-health-check.global.js" integrity="sha384-..." crossorigin="anonymous"></script>
<script>
  const report = window.SpreadsheetHealthCheck.analyse(bytes);
</script>
```

## Development

```bash
npm install
npm test            # Vitest, full suite
npm run build       # tsup: ESM + IIFE global bundles + type declarations
npm run demo -- test/fixtures/macros.xlsm   # CLI demo against a fixture
npm run check:iife  # proves the built IIFE bundle sets window.SpreadsheetHealthCheck
```

`npm run build` produces:

- `dist/spreadsheet-health-check.mjs` - ESM, for npm/bundler consumers.
- `dist/spreadsheet-health-check.global.js` - an IIFE that sets
  `window.SpreadsheetHealthCheck` when loaded via a classic `<script>` tag.
  This is the exact file the website serves and pins with an SRI hash.
- `dist/index.d.ts` - type declarations.

The engine itself (`src/`) is pure: no `fetch`, no DOM, no filesystem
access. Only `demo/` and `test/` touch the filesystem.

## Verifying the deployed bundle

dbhq.uk serves this exact bundle, byte-identical to a tagged release, with
Sigstore provenance so "the deployed bytes are the audited source" is
independently checkable rather than asserted.

Every tagged release is built by CI, hashed, and signed with a Sigstore
build-provenance attestation (keyless, recorded in the public Rekor
transparency log). To confirm the bundle dbhq.uk serves is the one built from
this source:

```bash
# 1. Download the released bundle (or save the one dbhq.uk serves)
gh release download v0.1.0 --repo dbhq-uk/spreadsheet-health-check \
  -p spreadsheet-health-check.global.js

# 2. Confirm its hash matches the one published in the release
sha256sum spreadsheet-health-check.global.js
#   c52ee26edb32f819867893781ea55643ca59a8971d2737f39816187bb17bf4b5

# 3. Verify the Sigstore provenance - proves these bytes were built by this
#    repo's release workflow, no human key involved
gh attestation verify spreadsheet-health-check.global.js \
  --repo dbhq-uk/spreadsheet-health-check
```

The bundle served at `https://dbhq.uk/tools/spreadsheet-health-check/` carries
the same SHA-256, so you can hash what the site serves and compare. This is
verifiable by anyone who checks; it is not auto-enforced in every visitor's
browser (no browser mechanism proves a first-party script against an external
attestation). Combined with the runtime check - open devtools and watch the
network tab; the analysis makes zero requests - that is a stronger integrity
story than a bare "we don't upload your file" claim.

## Licence

MIT - see [LICENSE](LICENSE).
