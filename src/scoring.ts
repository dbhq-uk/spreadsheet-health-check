import type { Finding, Band, Category, CategoryResult, Severity } from "./types";

const SEVERITY_PENALTY: Record<Severity, number> = { info: 2, low: 6, medium: 14, high: 26 };
const PER_CHECK_CAP = 30;
const CATEGORY_LABEL: Record<Category, string> = {
  "key-person": "Key-person exposure",
  "broken-today": "Broken today",
  "hidden-fragility": "Hidden fragility",
  "wont-scale": "Won't scale",
};
const CATEGORY_ORDER: Category[] = ["key-person", "broken-today", "hidden-fragility", "wont-scale"];

function penalty(f: Finding): number {
  const base = SEVERITY_PENALTY[f.severity];
  const scaled = base + Math.min(base, (f.count - 1) * 2); // extra weight for volume, capped at 2x base
  return Math.min(scaled, PER_CHECK_CAP);
}

export function bandFor(score: number): Band {
  if (score >= 75) return "Critical";
  if (score >= 50) return "High";
  if (score >= 25) return "Moderate";
  return "Low";
}

export function computeScore(findings: Finding[]): { score: number; band: Band; categories: CategoryResult[] } {
  const categories: CategoryResult[] = CATEGORY_ORDER.map((category) => {
    const catFindings = findings.filter(f => f.category === category);
    const subScore = Math.min(100, catFindings.reduce((s, f) => s + penalty(f), 0));
    return { category, label: CATEGORY_LABEL[category], subScore, findings: catFindings };
  });
  const score = Math.min(100, Math.round(categories.reduce((s, c) => s + c.subScore, 0)));
  return { score, band: bandFor(score), categories };
}
