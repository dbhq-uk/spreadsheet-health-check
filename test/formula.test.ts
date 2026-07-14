import { describe, it, expect } from "vitest";
import { stripStrings, functionsUsed, maxDepth, numericLiterals, isBusinessConstant } from "../src/formula";

describe("numericLiterals", () => {
  it("reads the literals actually written into a formula", () => {
    expect(numericLiterals("A1*1.2")).toEqual([1.2]);
    expect(numericLiterals("(B2+C2)*0.175+250")).toEqual([0.175, 250]);
  });

  // The same trap the r1c1 normaliser had to be taught: digits that belong to something else.
  it("ignores the digits inside cell references", () => {
    expect(numericLiterals("SUM(A1:A100)")).toEqual([]);
    expect(numericLiterals("$B$7+Sheet2!C12")).toEqual([]);
  });
  it("ignores the digits inside function names", () => {
    expect(numericLiterals("LOG10(A2)")).toEqual([]);
  });
  it("ignores the digits inside a quoted sheet name", () => {
    // 'Q1 Data' and '2026 Budget' are sheet names, not numbers - the same trap as
    // string literals, but quoted with apostrophes instead of double quotes.
    expect(numericLiterals("'2026 Budget'!A1*2")).toEqual([2]);
    expect(numericLiterals("SUM('Q1 Data'!A1:A5)")).toEqual([]);
  });
  it("ignores the digits inside string literals", () => {
    // The 1 in "Q1" is not a number; the trailing 0 genuinely is one (isBusinessConstant
    // is what later discards it as arithmetic scaffolding).
    expect(numericLiterals('IF(A1="Q1",B1,0)')).toEqual([0]);
    expect(numericLiterals('CONCATENATE("2026 forecast",A1)')).toEqual([]);
  });
});

describe("isBusinessConstant", () => {
  it("keeps rates, thresholds and day counts", () => {
    for (const n of [1.2, 0.2, 0.175, 20, 365, 1000]) expect(isBusinessConstant(n), String(n)).toBe(true);
  });
  // Small integers are overwhelmingly VLOOKUP/INDEX column indexes and month numbers.
  // Flagging them would fire on well-built workbooks - the single-author mistake.
  it("drops arithmetic scaffolding and lookup column indexes", () => {
    for (const n of [0, 1, 2, 3, 12]) expect(isBusinessConstant(n), String(n)).toBe(false);
  });
});

describe("maxDepth / functionsUsed / stripStrings", () => {
  it("measures bracket nesting", () => {
    expect(maxDepth("A1+1")).toBe(0);
    expect(maxDepth("IF(A1,IF(B1,IF(C1,1,0),0),0)")).toBe(3);
  });
  it("names the functions called", () => {
    expect(functionsUsed("IF(VLOOKUP(A1,B:C,2,0)>0,1,0)")).toEqual(["IF", "VLOOKUP"]);
  });
  it("a bracket or a function name inside a string cannot count", () => {
    expect(maxDepth('CONCAT(A1,"((((")')).toBe(1);
    expect(functionsUsed('IF(A1="SUM(",1,0)')).toEqual(["IF"]);
  });
  it("a bracket or a function-shaped word inside a quoted sheet name cannot count", () => {
    // Excel allows parentheses and spaces in sheet names.
    expect(functionsUsed("SUM('Data (2026)'!A1:A5)")).toEqual(["SUM"]);
    expect(maxDepth("SUM('Data (2026)'!A1:A5)")).toBe(1);
  });
  it("preserves escaped quotes when stripping", () => {
    expect(stripStrings('A1&"say ""hi"" now"&B1')).toBe('A1&""&B1');
  });
});
