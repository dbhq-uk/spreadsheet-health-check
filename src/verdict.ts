import type { Band, CategoryResult } from "./types";

const BAND_OPENER: Record<Band, string> = {
  Low: "This spreadsheet looks reasonably sound",
  Moderate: "This spreadsheet carries some real risk",
  High: "This spreadsheet is a serious risk to the business",
  Critical: "This spreadsheet is a critical risk to the business",
};

export function makeVerdict(band: Band, categories: CategoryResult[]): string {
  const worst = [...categories].filter(c => c.findings.length).sort((a, b) => b.subScore - a.subScore)[0];
  if (!worst) return `${BAND_OPENER[band]} - no significant issues were found.`;
  return `${BAND_OPENER[band]}, driven mainly by ${worst.label.toLowerCase()}.`;
}
