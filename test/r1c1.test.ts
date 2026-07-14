import { describe, it, expect } from "vitest";
import { relativeTemplate } from "../src/r1c1";

describe("relativeTemplate", () => {
  it("gives identical templates for the same relative formula in a column", () => {
    expect(relativeTemplate("A2*2", "B2")).toBe(relativeTemplate("A3*2", "B3"));
  });
  it("gives a different template for a different relative formula", () => {
    expect(relativeTemplate("A4*3", "B4")).not.toBe(relativeTemplate("A2*2", "B2"));
  });
});
