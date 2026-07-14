import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { scaleLimits } from "../../src/checks/scaleLimits";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("scaleLimits", () => {
  it("fires on a legacy .xls file", () => {
    const f = scaleLimits(ctx("legacy.xls"));
    expect(f?.id).toBe("wont-scale.limits");
  });
  it("does not fire on a small modern xlsx", () => {
    expect(scaleLimits(ctx("clean.xlsx"))).toBeNull();
  });
});
