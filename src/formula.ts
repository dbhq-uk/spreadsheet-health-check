// Formula tokenising helpers shared by the complexity, sprawl, opaque-reference and
// hardcoded-constant checks.
//
// The hard-won rule (same as src/r1c1.ts): a formula is not a string you can regex
// naively. Digits appear inside string literals ("Q1"), inside function names (LOG10)
// and inside cell references (A1). Every helper here walks the formula with the string
// literals and identifiers skipped, so none of those can masquerade as a number.

/**
 * Replace every quoted run with an empty one, preserving doubled-character escapes.
 * That covers both "string literals" and 'quoted sheet names' ('Q1 Data'!A1) - a sheet
 * name can carry digits, spaces and even parentheses, none of which may leak into the
 * helpers below as numbers, refs or brackets.
 */
export function stripStrings(f: string): string {
  let out = "";
  let i = 0;
  while (i < f.length) {
    const q = f[i];
    if (q === '"' || q === "'") {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === q) {
          if (f[j + 1] === q) { j += 2; continue; }
          break;
        }
        j++;
      }
      out += q + q;
      i = j + 1;
      continue;
    }
    out += f[i];
    i++;
  }
  return out;
}

/** The names of the functions a formula calls, upper-cased. String literals cannot contribute. */
export function functionsUsed(f: string): string[] {
  const names: string[] = [];
  for (const m of stripStrings(f).matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) names.push(m[1].toUpperCase());
  return names;
}

/** Deepest bracket nesting. String literals cannot contribute. */
export function maxDepth(f: string): number {
  let depth = 0, max = 0;
  for (const ch of stripStrings(f)) {
    if (ch === "(") max = Math.max(max, ++depth);
    else if (ch === ")") depth = Math.max(0, depth - 1);
  }
  return max;
}

/**
 * The numeric literals genuinely written into a formula - not the digits that belong to a
 * cell reference (A1, $B$7), a function name (LOG10), a sheet-qualified ref (Data!A1) or a
 * string literal ("Q1").
 */
export function numericLiterals(f: string): number[] {
  const s = stripStrings(f);
  const out: number[] = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    // Skip any identifier-shaped run whole: function names, cell refs, sheet names, named ranges.
    const ident = /^[A-Za-z_$][A-Za-z0-9_.$]*/.exec(s.slice(i));
    if (ident) { i += ident[0].length; continue; }
    const num = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(s.slice(i));
    if (num) {
      // A number directly followed by an identifier char is part of a reference-like token
      // (e.g. the "1" in a stray "1A"); the identifier branch above already ate the normal
      // cases, so this is belt-and-braces.
      const after = s[i + num[0].length];
      if (!after || !/[A-Za-z_$]/.test(after)) out.push(parseFloat(num[0]));
      i += num[0].length;
      continue;
    }
    i++;
  }
  return out;
}

/**
 * Is this literal one an owner would recognise as a business rule buried in a formula?
 *
 * Excludes 0/1/2 (arithmetic scaffolding) and small positive integers up to 12, because
 * those are overwhelmingly VLOOKUP/INDEX column indexes and month numbers - flagging them
 * would fire on well-built workbooks, which is exactly the false-positive trap that got the
 * single-author check deleted. A VAT rate (0.2, 1.2, 20), a day count (365) or a threshold
 * (1000) all survive the filter.
 */
export function isBusinessConstant(n: number): boolean {
  if (!Number.isFinite(n)) return false;
  const abs = Math.abs(n);
  if (abs === 0) return false;
  if (Number.isInteger(n)) return abs > 12;
  return true; // any non-integer literal - rates, multipliers, factors
}
