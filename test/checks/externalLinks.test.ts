import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { externalLinks } from "../../src/checks/externalLinks";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("externalLinks", () => {
  it("fires when an external workbook link is present", () => {
    const f = externalLinks(ctx("external-links.xlsx"));
    expect(f?.id).toBe("hidden-fragility.external-links");
    expect(f?.count).toBeGreaterThanOrEqual(1);
  });
  it("does not fire on a self-contained file", () => {
    expect(externalLinks(ctx("clean.xlsx"))).toBeNull();
  });
});
