// The band calibration, pinned to the fixtures.
//
// Removing the single-author check changed every score, and the key-person rebuild replaced
// one metadata signal with seven content signals - so the bands had to be retuned. This file
// is the retune: it asserts the band each fixture lands in, and prints the table so the
// effect of any change to a check or a scoring constant is visible rather than guessed at.
//
//   npm run scores
import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { analyse } from "../src/index";
import type { Band } from "../src/types";

const dir = join(__dirname, "fixtures");
const bytes = (n: string) => new Uint8Array(readFileSync(join(dir, n)));

// What each fixture must score. A fixture in the wrong band means either the check or the
// calibration is wrong - both are a real failure, so this is an assertion, not a report.
const EXPECTED: Record<string, Band> = {
  // Clean workbooks must stay Low. documented.xlsx is the negative control: a substantial,
  // well-built workbook. If an "absence of X" check ever fires on it, it fires on every good
  // workbook - the false-positive trap that got the single-author check deleted.
  "clean.xlsx": "Low",
  "documented.xlsx": "Low",
  "functions.xlsx": "Low",
  "merged.xlsx": "Low",
  "duplicate-sheets.xlsx": "Low",

  // A single medium-severity finding does not move the band on its own - it is shown in the
  // report, but one external link or one iterative-calc flag is not a business risk by itself.
  // (Both scored 14 before this rework too; removing single-author did not change them.)
  "circular.xlsx": "Low",
  "external-links.xlsx": "Low",

  // One real problem, and nothing else: a finding worth showing, not a crisis.
  "errors.xlsx": "Moderate",
  "manual-calc.xlsx": "Moderate",
  "hidden.xlsx": "Moderate",
  "macros.xlsm": "Moderate",
  "inconsistent.xlsx": "Moderate",
  "legacy.xls": "Moderate",
  "hardcoded.xlsx": "Moderate",

  // Evidence stacking up across a category is what High is for.
  "key-person.xlsx": "High",
};

describe("band calibration", () => {
  it("prints the score table", () => {
    const rows = readdirSync(dir).filter(f => /\.xls[xmb]?$/.test(f)).sort();
    const lines = rows.map((name) => {
      const r = analyse(bytes(name));
      const ids = r.findings.map(f => `${f.id.split(".")[1]}(${f.severity[0]}${f.count})`).join(" ");
      return `${name.padEnd(22)} ${String(r.score).padStart(3)}  ${r.riskBand.padEnd(9)} ${ids || "-"}`;
    });
    console.log(["", "fixture                score  band      fired checks", "-".repeat(96), ...lines, ""].join("\n"));
    expect(rows.length).toBeGreaterThan(0);
  });

  for (const [name, band] of Object.entries(EXPECTED)) {
    it(`${name} scores ${band}`, () => {
      expect(analyse(bytes(name)).riskBand).toBe(band);
    });
  }

  it("a clean workbook raises nothing at all", () => {
    expect(analyse(bytes("clean.xlsx")).findings).toEqual([]);
  });

  it("a well-built substantial workbook raises no absence-of-X findings", () => {
    const ids = analyse(bytes("documented.xlsx")).findings.map(f => f.id);
    expect(ids).not.toContain("key-person.undocumented");
    expect(ids).not.toContain("hidden-fragility.missing-validation");
    expect(ids).not.toContain("key-person.formula-sprawl");
  });

  it("no check reports a single-author finding, ever", () => {
    for (const name of readdirSync(dir).filter(f => /\.xls[xmb]?$/.test(f))) {
      expect(analyse(bytes(name)).findings.map(f => f.id)).not.toContain("key-person.single-author");
    }
  });

  it("every finding tells the owner what to do about it", () => {
    for (const name of readdirSync(dir).filter(f => /\.xls[xmb]?$/.test(f))) {
      for (const f of analyse(bytes(name)).findings) {
        expect(f.action.length, `${name} ${f.id} has no action`).toBeGreaterThan(20);
      }
    }
  });
});
