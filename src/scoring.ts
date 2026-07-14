import type { Finding, Band, Category, CategoryResult, Severity } from "./types";

// high is 26, not 24: a single high-severity finding must clear the Moderate threshold (25).
const SEVERITY_PENALTY: Record<Severity, number> = { info: 2, low: 5, medium: 14, high: 26 };
const PER_CHECK_CAP = 30;
const CATEGORY_CAP = 50;

// Within a category, evidence accumulates with diminishing returns. This matters most for
// key-person exposure, which is deliberately built from many individually-weak signals: a
// straight sum of seven checks would send any macro-enabled workbook straight to Critical,
// while capping the category flat would mean the second, third and fourth pieces of evidence
// counted for nothing. Weighted-by-rank does what a human reader does - the strongest finding
// carries, and each further one adds progressively less.
const RANK_WEIGHT = [1, 0.7, 0.5, 0.35, 0.25, 0.2, 0.15];
const TAIL_WEIGHT = 0.1;

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

function categoryScore(findings: Finding[]): number {
  const ranked = findings.map(penalty).sort((a, b) => b - a);
  const total = ranked.reduce((sum, p, i) => sum + p * (RANK_WEIGHT[i] ?? TAIL_WEIGHT), 0);
  return Math.min(CATEGORY_CAP, Math.round(total));
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
    return { category, label: CATEGORY_LABEL[category], subScore: categoryScore(catFindings), findings: catFindings };
  });
  const score = Math.min(100, Math.round(categories.reduce((s, c) => s + c.subScore, 0)));
  return { score, band: bandFor(score), categories };
}
