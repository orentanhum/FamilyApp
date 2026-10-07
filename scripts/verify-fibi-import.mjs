import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { billingMonthFromText } from "../src/import-period.ts";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const file = process.argv[2];

if (!file) throw new Error("Usage: npm run test:import -- /path/to/fibi.xls");

const workbook = XLSX.readFile(file, { cellDates: true });
let transactionCount = 0;
const periods = new Set();

for (const sheetName of workbook.SheetNames) {
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
  let header = null;
  let activePeriod = "";
  for (const row of rows) {
    const cells = row.map((value) => String(value ?? "").trim());
    activePeriod = billingMonthFromText(cells.join(" ")) || activePeriod;
    const dateIndex = cells.findIndex((value) => value === "תאריך עסקה");
    const merchantIndex = cells.findIndex((value) => /שם\s*העסק/.test(value));
    const chargeIndex = cells.findIndex((value) => value === "סכום חיוב");
    if (dateIndex >= 0 && merchantIndex >= 0 && chargeIndex >= 0) {
      header = { dateIndex, merchantIndex, chargeIndex };
      continue;
    }
    if (!header) continue;
    const merchant = String(row[header.merchantIndex] ?? "").trim();
    const charge = Number(String(row[header.chargeIndex] ?? "").replace(/[,\s]/g, ""));
    if (!row[header.dateIndex] || !merchant || !charge || /סה"?כ/.test(merchant)) continue;
    transactionCount += 1;
    periods.add(activePeriod);
  }
}

assert.equal(transactionCount, 224);
assert.deepEqual([...periods], ["202610"]);
console.log(`PASS: ${transactionCount} transactions, billing period 202610`);
