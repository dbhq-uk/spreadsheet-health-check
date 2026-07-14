import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { hiddenSheets } from "../../src/checks/hiddenSheets";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("hiddenSheets", () => {
  it("fires and lists hidden + veryHidden sheets", () => {
    const f = hiddenSheets(ctx("hidden.xlsx"));
    expect(f?.id).toBe("key-person.hidden-sheets");
    expect(f?.count).toBe(2);
    expect(f?.locations).toEqual(expect.arrayContaining(["Old", "Working"]));
    expect(f?.severity).toBe("high"); // veryHidden present
  });
  it("does not fire when all sheets visible", () => {
    expect(hiddenSheets(ctx("clean.xlsx"))).toBeNull();
  });
});
