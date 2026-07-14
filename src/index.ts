import { parseWorkbook } from "./parse";
import { ALL_CHECKS } from "./checks";
import { computeScore } from "./scoring";
import { makeVerdict } from "./verdict";
import type { HealthReport } from "./types";
export * from "./types";
export { parseWorkbook } from "./parse";

export function analyse(bytes: Uint8Array): HealthReport {
  const ctx = parseWorkbook(bytes);
  const findings = ALL_CHECKS.map(c => c(ctx)).filter((f): f is NonNullable<typeof f> => f !== null);
  const { score, band, categories } = computeScore(findings);
  return {
    riskBand: band,
    score,
    verdict: makeVerdict(band, categories),
    categories,
    findings,
    meta: {
      sheetCount: ctx.sheets.length,
      fileFormat: ctx.fileFormat,
      author: ctx.props.author,
      lastModifiedBy: ctx.props.lastModifiedBy,
      fileSizeBytes: ctx.fileSizeBytes,
      formulaCount: ctx.formulaCells.length,
    },
  };
}
