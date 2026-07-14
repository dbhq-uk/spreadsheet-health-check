import * as XLSX from "xlsx";

// Convert every A1-style cell reference in a formula to an offset relative to origin,
// producing a template identical for the same relative formula at any cell.
// Leaves quoted string literals, function names, and sheet names untouched so they
// cannot masquerade as cell references (which would create false "inconsistent formula" hits).
export function relativeTemplate(formula: string, originRef: string): string {
  const o = XLSX.utils.decode_cell(originRef);
  let out = "";
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // copy the quoted run verbatim, preserving doubled-character escapes: "string
      // literals" and 'quoted sheet names' alike ('Q1 Data'!A1 - the Q1 is text, not a ref)
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) { j += 2; continue; }
          break;
        }
        j++;
      }
      out += formula.slice(i, Math.min(j + 1, formula.length));
      i = j + 1;
      continue;
    }
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(formula.slice(i));
    if (m) {
      const after = formula[i + m[0].length];
      const isFunctionName = after === "(";   // e.g. LOG10( ... )
      const isSheetName = after === "!";       // e.g. AB12!A2
      const isColumnRef = /^[A-Z]{1,3}$/.test(m[2]);
      if (!isFunctionName && !isSheetName && isColumnRef) {
        try {
          const cell = XLSX.utils.decode_cell(`${m[2]}${m[4]}`);
          const dc = m[1] === "$" ? `C${cell.c}` : `c${cell.c - o.c}`;
          const dr = m[3] === "$" ? `R${cell.r}` : `r${cell.r - o.r}`;
          out += `[${dr}${dc}]`;
          i += m[0].length;
          continue;
        } catch { /* fall through to verbatim */ }
      }
      out += m[0];       // emit the identifier/ref-run verbatim
      i += m[0].length;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}
