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
    expect(relativeTemplate("Data!A2*2", "B2")).toBe(relativeTemplate("Data!A3*2", "B3"));
  });
});
