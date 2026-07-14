import * as XLSX from "xlsx";

// Convert every A1-style cell reference in a formula to an offset relative to origin,
// producing a template that is identical for the same relative formula at any cell.
export function relativeTemplate(formula: string, originRef: string): string {
  const o = XLSX.utils.decode_cell(originRef);
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, ac, col, ar, row) => {
    // skip if this looks like part of a sheet name or function - naive but effective for common cases
    try {
      const cell = XLSX.utils.decode_cell(`${col}${row}`);
      const dc = ac === "$" ? `C${cell.c}` : `c${cell.c - o.c}`;
      const dr = ar === "$" ? `R${cell.r}` : `r${cell.r - o.r}`;
      return `[${dr}${dc}]`;
    } catch {
      return m;
    }
  });
}
