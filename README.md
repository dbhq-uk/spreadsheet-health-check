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

Seventeen checks across four categories:

- **Key-person exposure** - hidden macros/VBA; hidden and very-hidden sheets;
  `INDIRECT`/`OFFSET` references that cannot be traced by reading them;
  formulas too long or too deeply nested to check; formula *sprawl* (hundreds
  of one-off shapes rather than a repeated pattern); sheets locked against
  editing; and a workbook that documents itself nowhere.
- **Broken today** - live error values (`#REF!`, `#DIV/0!`, ...); circular
  references with iterative calculation silently switched on; automatic
  calculation switched off, so the numbers on screen may be stale.
- **Hidden fragility** - inconsistent formulas within a column (the classic
  "someone typed over one cell" bug); external workbook links; business
  constants (a VAT rate, a margin) typed into formula after formula; merged
  cells inside data regions; no validation on any input.
- **Won't scale** - near-identical sheets copied per month/region/client;
  usage approaching Excel's row limits and the legacy `.xls` ceiling.

There is deliberately **no single-author check**. Comparing a workbook's
`creator` to its `lastModifiedBy` only records who saved the file last, so it
fires on almost every workbook and misses a genuinely one-person file that a
colleague once opened. It measures who saved last, not who understands it.
Key-person exposure is inferred from what is *in* the workbook - never from
its metadata.

Every finding carries an `action`: what to actually do about it.

Each check returns a `Finding` (or nothing). Findings are scored, grouped
into the four categories, and rolled up into a risk band. Within a category,
evidence accumulates with diminishing returns, so breadth of evidence raises
the band without any single check dominating it:

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

A file that is not really a workbook is refused rather than described:
`parseWorkbook` throws `NotASpreadsheetError` for anything that is not an
OOXML zip or an OLE2 container, so a renamed JPEG cannot come back with a
confident risk score.

### In the browser: the Web Worker

`analyse` is synchronous and a large workbook takes real time, so on the main
thread it freezes the tab. Use the **worker bundle** instead - it is
self-contained (the engine is inlined) and makes no network calls of any kind:

```js
const worker = new Worker("/spreadsheet-health-check.worker.js");
worker.onmessage = (e) => {
  if (e.data.type === "ready")  { /* engine parsed, warm and waiting */ }
  if (e.data.type === "result") { console.log(e.data.report.riskBand); }
  if (e.data.type === "error")  { console.log(e.data.message); }
};
// transfer the bytes - no copy, and nothing leaves the machine
worker.postMessage({ id: 1, bytes: arrayBuffer }, [arrayBuffer]);
```

There is also `dist/spreadsheet-health-check.global.js`, an IIFE that sets
`window.SpreadsheetHealthCheck` for a classic `<script>` tag. Prefer the
worker in any interactive page.

## Development

```bash
npm install
npm test            # Vitest, full suite
npm run build       # tsup: ESM + IIFE global bundles + type declarations
npm run demo -- test/fixtures/macros.xlsm   # CLI demo against a fixture
npm run scores      # the band calibration table for every fixture
npm run fixtures    # regenerate the test fixtures (reproducible - see below)
npm run example     # regenerate examples/example-workbook.xlsx
npm run check:worker  # boots the built worker bundle and analyses a fixture through it
npm run check:iife    # proves the built IIFE bundle sets window.SpreadsheetHealthCheck
```

Fixture generation is byte-reproducible: `zipSync` is given a pinned `mtime`,
so re-running `npm run fixtures` is a no-op in git unless a fixture's content
genuinely changed. (It used to stamp the current time into every zip entry,
which made three fixtures churn on every run.)

Scoring is calibrated against the fixtures by `test/calibration.test.ts`,
which asserts the band each one lands in - so a change to any check or scoring
constant that moves a workbook into the wrong band fails the build rather than
quietly re-scoring everyone's spreadsheet.

`npm run build` produces:

- `dist/spreadsheet-health-check.mjs` - ESM engine build, for importing
  `analyse()` as a library (attached to every GitHub release; the demo and
  tests run against `src/` directly).
- `dist/spreadsheet-health-check.worker.js` - the self-contained Web Worker
  bundle. **This is the exact file the website loads**, pinned by SHA-256.
- `dist/spreadsheet-health-check.global.js` - an IIFE that sets
  `window.SpreadsheetHealthCheck` when loaded via a classic `<script>` tag.
- `dist/index.d.ts` - type declarations.

The engine itself (`src/`) is pure: no `fetch`, no DOM, no filesystem access.
`browser/worker.ts` is the single, deliberately thin adapter where the engine
meets the browser; only `demo/`, `examples/` and `test/` touch the filesystem.

## Verifying the deployed bundle

dbhq.uk serves this exact bundle, byte-identical to a tagged release, with
Sigstore provenance so "the deployed bytes are the audited source" is
independently checkable rather than asserted.

Every tagged release is built by CI, hashed, and signed with a Sigstore
build-provenance attestation (keyless, recorded in the public Rekor
transparency log). To confirm the bundle dbhq.uk serves is the one built from
this source:

```bash
# 1. Save the exact worker bundle the site serves (cache-busted, so no stale edge copy)
curl -s "https://dbhq.uk/tools/spreadsheet-health-check/$(curl -s https://dbhq.uk/tools/spreadsheet-health-check/engine.manifest.json | jq -r .file)?v=$(date +%s)" \
  -o served.worker.js

# 2. Confirm its hash matches the one published in the release
sha256sum served.worker.js
curl -s https://dbhq.uk/tools/spreadsheet-health-check/engine.manifest.json | jq -r .sha256

# 3. Verify the Sigstore provenance - proves these bytes were built by this
#    repo's release workflow, no human key involved
gh attestation verify served.worker.js --repo dbhq-uk/spreadsheet-health-check
```

The page does the same check itself, in your browser, on every visit: before it
runs the engine it fetches the worker bundle, hashes it with SubtleCrypto and
compares that against the SHA-256 baked into the page at build time. If they
disagree it refuses to analyse anything.

Be honest about the ceiling. That check - like Subresource Integrity, which a
`new Worker()` cannot carry - proves the bytes match what this repo published;
it cannot prove the first party served you an honest page in the first place.
What a sceptic actually leans on is the combination: the published hash, the
Sigstore attestation, a rebuild from source, and the runtime observation that
the network tab stays silent while the file is analysed.

## Licence

MIT - see [LICENSE](LICENSE).
