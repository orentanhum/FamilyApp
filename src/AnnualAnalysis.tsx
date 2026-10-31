import React, { useState } from "react";
import { analysisPeriod, summarizeExpenses, type AnalysisTransaction } from "./annual-analysis";

const money = (value: number) => new Intl.NumberFormat("he-IL", {
  style: "currency", currency: "ILS",
}).format(value);
const monthNames = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
  "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];

function Ranking({ summary, label }: {
  summary: ReturnType<typeof summarizeExpenses>; label: string;
}) {
  if (!summary.count) return <p className="annualEmpty">אין הוצאות להצגה בתקופה זו.</p>;
  return <div className="tablewrap annualTable"><table aria-label={label}>
    <thead><tr><th>דירוג</th><th>קטגוריה</th><th>תנועות</th><th>סה״כ הוצאות</th><th>אחוז מההוצאות</th></tr></thead>
    <tbody>{summary.top.map((row, index) => <tr key={row.category}>
      <td>{index + 1}</td><th scope="row">{row.category}</th><td>{row.count}</td>
      <td dir="ltr">{money(row.total)}</td>
      <td>{new Intl.NumberFormat("he-IL", { style: "percent", maximumFractionDigits: 1 }).format(row.total / summary.total)}</td>
    </tr>)}</tbody>
  </table></div>;
}

export default function AnnualAnalysis({ rows }: { rows: AnalysisTransaction[] }) {
  const currentYear = new Intl.DateTimeFormat("en", { timeZone: "Asia/Jerusalem", year: "numeric" }).format(new Date());
  const years = [...new Set([currentYear, ...rows.map(row => analysisPeriod(row).slice(0, 4)).filter(Boolean)])].sort().reverse();
  const [year, setYear] = useState(currentYear);
  const dated = rows.map(row => ({ row, period: analysisPeriod(row) }));
  const annual = summarizeExpenses(dated.filter(item => item.period.slice(0, 4) === year).map(item => item.row));
  const monthly = monthNames.map((name, index) => {
    const period = year + String(index + 1).padStart(2, "0");
    const monthRows = dated.filter(item => item.period === period).map(item => item.row);
    return { name, period, hasData: monthRows.length > 0, summary: summarizeExpenses(monthRows) };
  });
  return <div className="annualAnalysis">
    <div className="title annualTitle"><div><h2>ניתוח שנתי</h2>
      <p>10 קטגוריות ההוצאה הגבוהות בשנה ובכל חודש, מהגבוה לנמוך</p></div>
      <label>שנה <select value={year} onChange={event => setYear(event.target.value)}>
        {years.map(value => <option key={value}>{value}</option>)}
      </select></label>
    </div>
    <p className="sub">לפי החודש המשויך לתנועה בסיכומים. כולל רק תנועות הוצאה שנכללות בחישוב; הכנסות וזיכויים אינם נכללים.</p>
    <section className="cards" aria-label="סיכום השנה">
      <article><b>{money(annual.total)}</b><span>סה״כ הוצאות ב־{year}</span></article>
      <article><b>{annual.count}</b><span>תנועות הוצאה</span></article>
      <article><b>{monthly.filter(month => month.hasData).length}</b><span>חודשים עם תנועות במאגר</span></article>
    </section>
    <section className="annualSection" aria-labelledby="annualTopTitle">
      <h3 id="annualTopTitle">טופ 10 קטגוריות — שנת {year}</h3>
      <Ranking summary={annual} label={`טופ 10 קטגוריות בשנת ${year}`} />
    </section>
    <h3>טופ 10 קטגוריות לפי חודש</h3>
    <div className="annualMonths">{monthly.map(month => <section className="annualSection" key={month.period}>
      <div className="title"><h3>{month.name} {year}</h3>
        {month.summary.count > 0 && <span>{money(month.summary.total)} · {month.summary.count} תנועות</span>}
      </div>
      {month.hasData ? <Ranking summary={month.summary} label={`טופ 10 קטגוריות — ${month.name} ${year}`} />
        : <p className="annualEmpty">אין תנועות במאגר לחודש זה.</p>}
    </section>)}</div>
  </div>;
}
