import { describe, it, expect } from "vitest";
import { relativeTemplate } from "../src/r1c1";

describe("relativeTemplate", () => {
  it("gives identical templates for the same relative formula in a column", () => {
    expect(relativeTemplate("A2*2", "B2")).toBe(relativeTemplate("A3*2", "B3"));
  });
  it("gives a different template for a different relative formula", () => {
    expect(relativeTemplate("A4*3", "B4")).not.toBe(relativeTemplate("A2*2", "B2"));
  });
  it("does not treat a function name as a cell reference", () => {
    expect(relativeTemplate("LOG10(A2)*2", "B2")).toBe(relativeTemplate("LOG10(A3)*2", "B3"));
  });
  it("does not treat digits inside a quoted string literal as a cell reference", () => {
    expect(relativeTemplate('IF(A2>0,"Q1",0)', "B2")).toBe(relativeTemplate('IF(A3>0,"Q1",0)', "B3"));
  });
  it("keeps a sheet-qualified relative ref consistent across rows", () => {
    // sheet name AB1 itself looks like a cell ref - the old regex broke on exactly this
    expect(relativeTemplate("AB1!A2*2", "B2")).toBe(relativeTemplate("AB1!A3*2", "B3"));
  });
  it("does not treat a quoted sheet name as a cell reference", () => {
    // 'Q1 Data' quoted: the Q1 inside is text, not a ref - relativising it made the
    // same copied-down formula produce a different template on every row.
    expect(relativeTemplate("'Q1 Data'!A2*2", "B2")).toBe(relativeTemplate("'Q1 Data'!A3*2", "B3"));
    expect(relativeTemplate("'Q1 Data'!A2*2", "B2")).toContain("'Q1 Data'!");
  });
});
