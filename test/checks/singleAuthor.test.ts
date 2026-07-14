import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { singleAuthor } from "../../src/checks/singleAuthor";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("singleAuthor", () => {
  it("fires when author == lastModifiedBy", () => {
    const f = singleAuthor(ctx("single-author.xlsx"));
    expect(f?.category).toBe("key-person");
    expect(f?.id).toBe("key-person.single-author");
  });
  it("does not fire when authors differ", () => {
    expect(singleAuthor(ctx("clean.xlsx"))).toBeNull();
  });
});
