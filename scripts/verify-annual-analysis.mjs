import assert from "node:assert/strict";
import { analysisPeriod, summarizeExpenses } from "../src/annual-analysis.ts";

const expense = (overrides = {}) => ({
  transaction_date: "2025-12-15", yyyymm: 202601, amount: -100,
  transaction_type: "expense", included_in_expenses: true, category: "מזון", ...overrides,
});
assert.equal(analysisPeriod(expense()), "202601", "installment uses its assigned month across year boundaries");
assert.equal(analysisPeriod(expense({ yyyymm: null })), "202512");
assert.equal(analysisPeriod(expense({ yyyymm: 202613, transaction_date: "invalid" })), "");
const rows = [expense(), expense({ amount: -0.1 }), expense({ amount: -0.2 }),
  expense({ included_in_expenses: false, amount: -500 }),
  expense({ transaction_type: "income", amount: 1000 }),
  expense({ transaction_type: "refund", amount: 20 }),
  expense({ amount: 20 }), expense({ amount: NaN }), expense({ category: null, amount: -10 })];
const summary = summarizeExpenses(rows);
assert.equal(summary.total, 110.3);
assert.equal(summary.count, 4);
assert.equal(summary.top[0].total, 100.3);
assert.equal(summary.top[1].category, "לא מסווג");
assert.equal(summarizeExpenses([]).total, 0);
const ranked = summarizeExpenses(Array.from({ length: 12 }, (_, i) => expense({ category: `category-${i}`, amount: -(i + 1) })));
assert.equal(ranked.top.length, 10);
assert.equal(ranked.top[0].total, 12);
assert.equal(ranked.top[9].total, 3);
assert.equal(ranked.total, 78, "total and shares use all categories, including those outside the top ten");
const dated = rows.map(row => ({ row, period: analysisPeriod(row) }));
const months = Array.from({ length: 12 }, (_, i) => `2026${String(i + 1).padStart(2, "0")}`);
assert.equal(months.reduce((sum, month) => sum + Math.round(summarizeExpenses(dated.filter(item => item.period === month).map(item => item.row)).total * 100), 0), Math.round(summary.total * 100));
console.log("Annual analysis checks passed");
