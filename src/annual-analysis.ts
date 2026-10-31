export type AnalysisTransaction = {
  transaction_date: string;
  yyyymm: number | null;
  amount: number;
  transaction_type: string;
  included_in_expenses: boolean;
  category: string | null;
};

export function analysisPeriod(row: AnalysisTransaction): string {
  const period = String(row.yyyymm || "");
  if (/^\d{4}(0[1-9]|1[0-2])$/.test(period)) return period;
  const date = row.transaction_date?.slice(0, 7).replace("-", "") || "";
  return /^\d{4}(0[1-9]|1[0-2])$/.test(date) ? date : "";
}

export function summarizeExpenses(rows: AnalysisTransaction[]) {
  const categories = new Map<string, { category: string; total: number; count: number }>();
  for (const row of rows) {
    if (!row.included_in_expenses || row.transaction_type !== "expense" ||
        !Number.isFinite(row.amount) || row.amount >= 0) continue;
    const category = row.category?.trim() || "לא מסווג";
    const entry = categories.get(category) || { category, total: 0, count: 0 };
    entry.total += Math.round(-row.amount * 100);
    entry.count++;
    categories.set(category, entry);
  }
  const ranked = [...categories.values()]
    .map(row => ({ ...row, total: row.total / 100 }))
    .sort((a, b) => b.total - a.total || a.category.localeCompare(b.category, "he"));
  return {
    total: ranked.reduce((sum, row) => sum + Math.round(row.total * 100), 0) / 100,
    count: ranked.reduce((sum, row) => sum + row.count, 0),
    top: ranked.slice(0, 10),
  };
}
