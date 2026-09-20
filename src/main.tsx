import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import "./style.css";
import { supabase } from "./supabase";
const members = ["ליאת", "אורן", "רום", "נועם", "אלון", "אמיר", "ג׳וי"];
type Dup = "new" | "existing" | "possible";
type Tx = {
  date: string;
  description: string;
  amount: number;
  type: string;
  category: string;
  reference: string;
  includedInExpenses: boolean;
  yyyymm?: string;
  sourceType?: string;
  subcategory?: string;
  nature?: string;
  ordinal?: number;
  identity?: string;
  dup?: Dup;
};
type DbTx = {
  id: string;
  transaction_date: string;
  description: string;
  amount: number;
  transaction_type: string;
  category: string | null;
  subcategory: string | null;
  included_in_expenses: boolean;
  source_type: string;
  yyyymm: number | null;
  expense_nature: string | null;
  transaction_identity?: string | null;
};
type Mapping = {
  source_type: string;
  match_text: string;
  category: string;
  subcategory: string | null;
  included_in_expenses: boolean;
};
const months = [
  ["", "הכל"],
  ["202603", "מרץ"],
  ["202604", "אפריל"],
  ["202605", "מאי"],
  ["202606", "יוני"],
  ["202607", "יולי"],
  ["202608", "אוגוסט"],
  ["202609", "ספטמבר"],
];
const natureLabel = (n: string | null) =>
  (
    ({
      recurring: "קבוע",
      variable: "משתנה",
      periodic: "תקופתי",
      one_time: "חד־פעמי",
    }) as any
  )[n || ""] ||
  n ||
  "—";
const money = (n: number) =>
  new Intl.NumberFormat("he-IL", { style: "currency", currency: "ILS" }).format(
    n,
  );
const norm = (s: string) =>
  String(s || "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
function simpleHash(raw: string, p = "x") {
  let h = 2166136261;
  for (let i = 0; i < raw.length; i++) {
    h ^= raw.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return p + "_" + (h >>> 0).toString(16) + "_" + raw.length;
}
function merchantKey(description: string) {
  return norm(description)
    .replace(/[\u200e\u200f]/g, "")
    .replace(/(?:עסקה|עסקת|חיוב|תשלום|אישור|אסמכתא)\s*[:#-]?\s*\d+/g, " ")
    .replace(/\b\d{4,}\b/g, " ")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function daysApart(a: string, b: string) {
  const one = Date.parse(`${a}T00:00:00Z`),
    two = Date.parse(`${b}T00:00:00Z`);
  return Number.isFinite(one) && Number.isFinite(two)
    ? Math.abs(one - two) / 86400000
    : Number.POSITIVE_INFINITY;
}
function installment(ref: string) {
  const s = norm(ref);
  const m = s.match(/(?:תשלום\s*)?(\d+)\s*(?:מתוך|מ\s*-?)\s*(\d+)/);
  return m ? `${m[1]}/${m[2]}` : "";
}
function stableReference(ref: string) {
  return norm(ref)
    .replace(/מזהה כרטיס[^,.;]*/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
function contentKey(r: Tx) {
  return [
    r.sourceType,
    norm(r.description),
    r.amount.toFixed(2),
    installment(r.reference) || stableReference(r.reference),
  ].join("|");
}
function identity(r: Tx) {
  return simpleHash(
    [
      r.sourceType,
      iso(r.date),
      norm(r.description),
      r.amount.toFixed(2),
      installment(r.reference) || stableReference(r.reference),
    ].join("|"),
    "tx4",
  );
}
function fingerprint(r: Tx, file: string) {
  return simpleHash(
    [
      file,
      r.ordinal ?? "",
      iso(r.date),
      r.description,
      r.amount.toFixed(2),
      r.reference,
    ].join("|"),
    "v3",
  );
}
function isCardSettlement(d: string) {
  return /ישראכרט|מקס\s*הבינלאומי|כרטיסי\s*אשראי\s*לי/i.test(d || "");
}
function isInvestmentOrTransfer(d: string) {
  return /פק\s*קרן|פיקדון|רכישת\s*מט["״]?ח|המרת\s*מט["״]?ח|העברה\s+מהחשבון|העברה\s+לחשבון|העברה\s+בין/i.test(
    d || "",
  );
}
function guessCategory(s: string) {
  if (isCardSettlement(s)) return "כרטיסי אשראי";
  if (/משכורת/.test(s)) return "משכורת";
  if (/קצבת\s*ילדים/.test(s)) return "קצבת ילדים";
  if (/משכנת/.test(s)) return "דיור ושירותים";
  if (/ראלי|ריאלי/.test(s)) return "חינוך";
  if (/מכבי|רמב|איזנבוד|אופטיקנה/.test(s)) return "בריאות";
  if (/סלקום אנרג/.test(s)) return "דיור ושירותים";
  if (/פלאפון|בזק/.test(s)) return "תקשורת";
  if (/שופרסל|רמי לוי|סטופמרקט|AM:PM|PM:AM/.test(s)) return "מזון";
  if (/ארקיע|אל-על|אל על|איסתא|מלון/.test(s)) return "נופש ונסיעות";
  if (/YELLOW|סונול|פז /.test(s)) return "רכב ותחבורה";
  if (/מסעד|קפה|מחניודה|פסטורי/.test(s)) return "מסעדות";
  return "לא מסווג";
}
function asNumber(v: any) {
  return typeof v === "number"
    ? v
    : Number(String(v ?? "").replace(/[₪,\s]/g, "")) || 0;
}
function iso(v: any) {
  const pad = (n: number) => String(n).padStart(2, "0"),
    inIsrael = (d: Date) => {
      if (Number.isNaN(d.getTime())) return "";
      const parts = Object.fromEntries(
        new Intl.DateTimeFormat("en", {
          timeZone: "Asia/Jerusalem",
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
        })
          .formatToParts(d)
          .map((part) => [part.type, part.value]),
      );
      return `${parts.year}-${parts.month}-${parts.day}`;
    };
  if (v instanceof Date) return inIsrael(v);
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? `${d.y}-${pad(d.m)}-${pad(d.d)}` : "";
  }
  const s = String(v ?? "").trim();
  if (!s) return "";
  const ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:$|T)/);
  if (ymd) return `${ymd[1]}-${pad(Number(ymd[2]))}-${pad(Number(ymd[3]))}`;
  const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (dmy) return `${dmy[3]}-${pad(Number(dmy[2]))}-${pad(Number(dmy[1]))}`;
  return inIsrael(new Date(s));
}
function billingMonth(name: string) {
  const n = name.match(
    /1788266165702|1788122731881|1788122470543|1788122456822|1788122208047|1788121906677|1788119346635/,
  )?.[0];
  return (
    (
      {
        "1788266165702": "202609",
        "1788122731881": "202603",
        "1788122470543": "202604",
        "1788122456822": "202605",
        "1788122208047": "202606",
        "1788121906677": "202607",
        "1788119346635": "202608",
      } as any
    )[n || ""] || ""
  );
}
function applyMapping(r: Tx, maps: Mapping[]) {
  const d = norm(r.description),
    m = maps
      .filter(
        (x) =>
          (!x.source_type || x.source_type === r.sourceType) &&
          d.includes(norm(x.match_text)),
      )
      .sort(
        (a, b) => (b.match_text || "").length - (a.match_text || "").length,
      )[0];
  return m
    ? {
        ...r,
        category: m.category || r.category,
        subcategory: m.subcategory || undefined,
        includedInExpenses: m.included_in_expenses,
      }
    : r;
}
function parseCard(raw: any[][], file: string, startOrdinal = 0) {
  const out: Tx[] = [];
  let h: string[] | null = null,
    ordinal = startOrdinal;
  for (const row of raw) {
    const c = row.map((v) => String(v ?? "").trim());
    if (
      c.includes("תאריך עסקה") &&
      c.some((v) => /שם\s*העסק/.test(v)) &&
      c.some((v) => /^סכום חיוב$/.test(v))
    ) {
      h = c;
      continue;
    }
    if (!h) continue;
    const idx = (re: RegExp) => h!.findIndex((v) => re.test(v)),
      di = idx(/^תאריך עסקה$/),
      mi = idx(/שם\s*העסק/),
      ci = idx(/^סכום חיוב$/),
      ri = idx(/פירוט/);
    if (di < 0 || mi < 0 || ci < 0) continue;
    const desc = String(row[mi] ?? "").trim(),
      charge = asNumber(row[ci]);
    if (!desc || !row[di] || charge === 0 || /סה"?כ/.test(desc)) continue;
    const amount = -charge;
    out.push({
      date: iso(row[di]),
      description: desc,
      amount,
      type: amount < 0 ? "expense" : "refund",
      category: guessCategory(desc),
      reference: ri >= 0 ? String(row[ri] ?? "") : "",
      includedInExpenses: true,
      yyyymm: billingMonth(file) || "202609",
      sourceType: "credit_card",
      nature: "variable",
      ordinal: ordinal++,
    });
  }
  return out;
}
function bankHeader(raw: any[][]) {
  for (let i = 0; i < Math.min(raw.length, 60); i++) {
    const h = raw[i].map((v) => norm(String(v ?? "")));
    const find = (re: RegExp) => h.findIndex((v) => re.test(v));
    const di = find(/^(תאריך|תאריך ערך|תאריך פעולה|יום)$/),
      xi = find(
        /^(תיאור|תאור|פרטים|פירוט|תיאור הפעולה|תאור הפעולה|פרטי פעולה|פרטי התנועה)$/,
      ),
      de = find(/^(חובה|חיוב|משיכה|סכום חובה)$/),
      cr = find(/^(זכות|זיכוי|הפקדה|סכום זכות)$/),
      amt = find(/^(סכום|סכום פעולה)$/),
      ref = find(/^(אסמכתא|סימוכין|reference)$/);
    if (di >= 0 && xi >= 0 && (de >= 0 || cr >= 0 || amt >= 0))
      return { i, h, di, xi, de, cr, amt, ref };
  }
  return null;
}
function App() {
  const [user, setUser] = useState<any>(),
    [email, setEmail] = useState("orentanhum@gmail.com"),
    [password, setPassword] = useState(""),
    [msg, setMsg] = useState(""),
    [page, setPage] = useState("home"),
    [db, setDb] = useState<DbTx[]>([]),
    [month, setMonth] = useState(""),
    [sourceFilter, setSourceFilter] = useState(""),
    [categoryFilter, setCategoryFilter] = useState(""),
    [subcategoryFilter, setSubcategoryFilter] = useState(""),
    [summaryLevel, setSummaryLevel] = useState<"category" | "subcategory">(
      "category",
    ),
    [search, setSearch] = useState(""),
    [preview, setPreview] = useState<Tx[]>([]),
    [fileName, setFileName] = useState(""),
    [saving, setSaving] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => setUser(data.session?.user || null));
    const { data: s } = supabase.auth.onAuthStateChange((_e, x) =>
      setUser(x?.user || null),
    );
    return () => s.subscription.unsubscribe();
  }, []);
  useEffect(() => {
    if (user) load();
  }, [user]);
  async function familyId() {
    const { data, error } = await supabase
      .from("family_members")
      .select("family_id")
      .eq("user_id", user.id)
      .eq("is_active", true)
      .limit(1)
      .maybeSingle();
    if (error || !data) throw new Error("לא נמצאה משפחה פעילה");
    return data.family_id;
  }
  async function load() {
    try {
      const fid = await familyId(),
        all: DbTx[] = [];
      for (let from = 0; ; from += 1000) {
        const { data, error } = await supabase
          .from("transactions")
          .select(
            "id,transaction_date,description,amount,transaction_type,category,subcategory,included_in_expenses,source_type,yyyymm,expense_nature,transaction_identity",
          )
          .eq("family_id", fid)
          .eq("is_active", true)
          .order("transaction_date", { ascending: false })
          .order("id", { ascending: true })
          .range(from, from + 999);
        if (error) throw error;
        const rows = (data || []) as DbTx[];
        all.push(...rows);
        if (rows.length < 1000) break;
      }
      setDb(all);
    } catch (e: any) {
      setMsg(e.message);
    }
  }
  async function login(e: React.FormEvent) {
    e.preventDefault();
    const { error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });
    if (error) setMsg(error.message);
  }
  async function getMappings(fid: string) {
    const { data, error } = await supabase
      .from("transaction_mappings")
      .select(
        "source_type,match_text,category,subcategory,included_in_expenses",
      )
      .eq("family_id", fid)
      .eq("is_active", true)
      .eq("is_approved", true);
    if (error) throw error;
    return (data || []) as Mapping[];
  }
  async function markDuplicates(rows: Tx[], fid: string, importingFile: string) {
    const prepared = rows.map((r) => ({
      ...r,
      identity: identity(r),
      dup: "new" as Dup,
    }));
    const ids = [...new Set(prepared.map((r) => r.identity!))];
    let existing: any[] = [];
    for (let i = 0; i < ids.length; i += 100) {
      const { data, error } = await supabase
        .from("transactions")
        .select(
          "id,transaction_identity,transaction_date,description,amount,reference,source_type,source_file,yyyymm",
        )
        .eq("family_id", fid)
        .eq("is_active", true)
        .in("transaction_identity", ids.slice(i, i + 100));
      if (error) throw error;
      existing.push(...(data || []));
    }

    const periods = [
      ...new Set(
        prepared
          .map((r) => Number(r.yyyymm || 0))
          .filter((x) => Number.isFinite(x) && x > 0),
      ),
    ];
    const sources = [
      ...new Set(prepared.map((r) => r.sourceType).filter(Boolean)),
    ] as string[];
    if (periods.length && sources.length) {
      for (let from = 0; ; from += 1000) {
        const { data, error } = await supabase
          .from("transactions")
          .select(
            "id,transaction_identity,transaction_date,description,amount,reference,source_type,source_file,yyyymm",
          )
          .eq("family_id", fid)
          .eq("is_active", true)
          .in("yyyymm", periods)
          .in("source_type", sources)
          .range(from, from + 999);
        if (error) throw error;
        existing.push(...(data || []));
        if ((data || []).length < 1000) break;
      }
    }

    const uniqueExisting = [
      ...new Map(existing.map((x) => [x.id, x])).values(),
    ];
    const exact = new Map<string, any[]>();
    const nearby = new Map<string, any[]>();
    for (const x of uniqueExisting) {
      if (x.transaction_identity) {
        const list = exact.get(x.transaction_identity) || [];
        list.push(x);
        exact.set(x.transaction_identity, list);
      }
      const key = contentKey({
        date: x.transaction_date,
        description: x.description,
        amount: Number(x.amount),
        type: "",
        category: "",
        reference: x.reference || "",
        includedInExpenses: true,
        yyyymm: String(x.yyyymm || ""),
        sourceType: x.source_type,
      });
      const list = nearby.get(key) || [];
      list.push(x);
      nearby.set(key, list);
    }
    const used = new Set<string>();
    return prepared.map((r) => {
      const exactMatch = (exact.get(r.identity!) || []).find(
        (x) => !used.has(x.id),
      );
      if (exactMatch) {
        used.add(exactMatch.id);
        return { ...r, dup: "existing" as Dup };
      }
      const nearbyMatch = (nearby.get(contentKey(r)) || [])
        .filter(
          (x) =>
            !used.has(x.id) &&
            x.source_file !== importingFile &&
            Number(x.yyyymm || 0) === Number(r.yyyymm || 0) &&
            daysApart(x.transaction_date, r.date) <= 1,
        )
        .sort(
          (a, b) =>
            daysApart(a.transaction_date, r.date) -
            daysApart(b.transaction_date, r.date),
        )[0];
      if (nearbyMatch) {
        used.add(nearbyMatch.id);
        return { ...r, dup: "existing" as Dup };
      }
      return r;
    });
  }
  async function importFile(file: File) {
    try {
      setFileName(file.name);
      const fid = await familyId(),
        maps = await getMappings(fid);
      const wb = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        cellDates: true,
      });
      let card: Tx[] = [],
        ordinal = 0;
      for (const sn of wb.SheetNames) {
        const raw: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[sn], {
          header: 1,
          defval: "",
          raw: true,
        });
        const bh = bankHeader(raw);
        if (bh) {
          let out: Tx[] = raw
            .slice(bh.i + 1)
            .map((r, i) => {
              const d = String(r[bh.xi] ?? "").trim();
              const date = iso(r[bh.di]);
              let debit = bh.de >= 0 ? Math.abs(asNumber(r[bh.de])) : 0,
                credit = bh.cr >= 0 ? Math.abs(asNumber(r[bh.cr])) : 0,
                amount = debit ? -debit : credit;
              if (!amount && bh.amt >= 0) amount = asNumber(r[bh.amt]);
              const special = isCardSettlement(d) || isInvestmentOrTransfer(d);
              return {
                date,
                description: d,
                amount,
                type: special ? "transfer" : amount < 0 ? "expense" : "income",
                category: isCardSettlement(d)
                  ? "כרטיסי אשראי"
                  : guessCategory(d),
                reference: bh.ref >= 0 ? String(r[bh.ref] || "") : "",
                includedInExpenses: !special,
                yyyymm: date ? date.slice(0, 7).replace("-", "") : "",
                sourceType: "bank",
                nature: special ? "one_time" : "variable",
                ordinal: ordinal + i,
              } as Tx;
            })
            .filter((r) => r.date && r.description && r.amount !== 0);
          out = await markDuplicates(
            out.map((r) => applyMapping(r, maps)),
            fid,
            file.name,
          );
          setPreview(out);
          setMsg(
            `בנק: נמצאו ${out.length} תנועות · ${out.filter((r) => r.dup === "new").length} חדשות · ${out.filter((r) => r.dup === "existing").length} כבר קיימות · ${out.filter((r) => r.dup === "possible").length} לבדיקה`,
          );
          return;
        }
        const part = parseCard(raw, file.name, ordinal);
        ordinal += part.length;
        card.push(...part);
      }
      if (!card.length)
        throw new Error(
          "פורמט הקובץ לא זוהה. נדרשות כותרות תאריך/תיאור וסכום או פורמט כרטיס אשראי.",
        );
      card = await markDuplicates(
        card.map((r) => applyMapping(r, maps)),
        fid,
        file.name,
      );
      setPreview(card);
      setMsg(
        `אשראי: נמצאו ${card.length} · ${card.filter((r) => r.dup === "new").length} חדשות · ${card.filter((r) => r.dup === "existing").length} כבר קיימות · ${card.filter((r) => r.dup === "possible").length} לבדיקה`,
      );
    } catch (e: any) {
      setMsg("שגיאה בניתוח הקובץ: " + e.message);
    }
  }
  async function save() {
    if (!preview.length) return;
    setSaving(true);
    try {
      const fid = await familyId(),
        toSave = preview.filter((r) => r.dup !== "existing");
      let saved = 0;
      for (let i = 0; i < toSave.length; i += 100) {
        const batch = toSave.slice(i, i + 100).map((r) => ({
          family_id: fid,
          member_id: null,
          transaction_date: r.date,
          description: r.description,
          amount: r.amount,
          currency: "ILS",
          transaction_type: r.type,
          category: r.category,
          subcategory: r.subcategory || null,
          included_in_expenses: r.includedInExpenses,
          reference: r.reference || null,
          source_type: r.sourceType,
          source_file: fileName,
          yyyymm: Number(r.yyyymm),
          status: "done",
          is_active: true,
          expense_nature: r.nature || null,
          is_core_living_expense: r.includedInExpenses,
          created_by: user.id,
          updated_by: user.id,
          fingerprint: fingerprint(r, fileName),
          transaction_identity: r.identity || identity(r),
          duplicate_status: r.dup === "possible" ? "possible" : "clear",
        }));
        const { data, error } = await supabase
          .from("transactions")
          .upsert(batch, {
            onConflict: "family_id,fingerprint",
            ignoreDuplicates: true,
          })
          .select("id");
        if (error) throw error;
        saved += data?.length || 0;
      }
      setMsg(
        `נשמרו ${saved} תנועות חדשות; ${preview.filter((r) => r.dup === "existing").length} קיימות דולגו`,
      );
      setPreview([]);
      await load();
    } catch (e: any) {
      setMsg("שגיאה: " + e.message);
    } finally {
      setSaving(false);
    }
  }
  if (user === undefined) return null;
  if (!user)
    return (
      <div dir="rtl" className="authPage">
        <form className="authCard" onSubmit={login}>
          <h1>FamilyApp</h1>
          <h2>כניסה</h2>
          <input value={email} onChange={(e) => setEmail(e.target.value)} />
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />
          <button>כניסה</button>
          {msg && <p>{msg}</p>}
        </form>
      </div>
    );
  const categories = Array.from(
      new Set(db.map((r) => r.category).filter(Boolean) as string[]),
    ).sort((a, b) => a.localeCompare(b, "he")),
    subcategories = Array.from(
      new Set(
        db
          .filter((r) => !categoryFilter || r.category === categoryFilter)
          .map((r) => r.subcategory)
          .filter(Boolean) as string[],
      ),
    ).sort((a, b) => a.localeCompare(b, "he"));
  const filtered = db.filter(
      (r) =>
        (!month || String(r.yyyymm) === month) &&
        (!sourceFilter || r.source_type === sourceFilter) &&
        (!categoryFilter || r.category === categoryFilter) &&
        (!subcategoryFilter || r.subcategory === subcategoryFilter) &&
        (!search ||
          r.description.includes(search) ||
          r.category?.includes(search) ||
          r.subcategory?.includes(search)),
    ),
    expenses = -filtered
      .filter(
        (r) =>
          r.amount < 0 &&
          r.included_in_expenses &&
          r.transaction_type === "expense",
      )
      .reduce((s, r) => s + r.amount, 0),
    income = filtered
      .filter((r) => r.amount > 0 && r.transaction_type === "income")
      .reduce((s, r) => s + r.amount, 0),
    refunds = filtered
      .filter((r) => r.amount > 0 && r.transaction_type === "refund")
      .reduce((s, r) => s + r.amount, 0),
    expenseRows = filtered.filter(
      (r) =>
        r.amount < 0 &&
        r.included_in_expenses &&
        r.transaction_type === "expense",
    ),
    summary = Object.values(
      expenseRows.reduce((a: any, r) => {
        const c = r.category || "לא מסווג",
          s = r.subcategory || "ללא תת קטגוריה",
          k = summaryLevel === "category" ? c : [c, s].join("|");
        if (!a[k])
          a[k] = {
            category: c,
            subcategory: summaryLevel === "subcategory" ? s : "",
            total: 0,
            count: 0,
          };
        a[k].total += Math.abs(r.amount);
        a[k].count++;
        return a;
      }, {}),
    ).sort(
      (a: any, b: any) =>
        b.total - a.total ||
        a.category.localeCompare(b.category, "he") ||
        a.subcategory.localeCompare(b.subcategory, "he"),
    ),
    natureTotals = Object.entries(
      expenseRows.reduce((a: any, r) => {
        const n = r.expense_nature || "variable";
        a[n] = (a[n] || 0) + Math.abs(r.amount);
        return a;
      }, {}),
    ).sort((a: any, b: any) => b[1] - a[1]),
    nowParts = Object.fromEntries(
      new Intl.DateTimeFormat("en", {
        timeZone: "Asia/Jerusalem",
        year: "numeric",
        month: "2-digit",
      })
        .formatToParts(new Date())
        .map((part) => [part.type, part.value]),
    ),
    currentPeriod = Number(`${nowParts.year}${nowParts.month}`),
    currentLabel = new Intl.DateTimeFormat("he-IL", {
      timeZone: "Asia/Jerusalem",
      month: "long",
      year: "numeric",
    }).format(new Date()),
    current = db.filter((r) => r.yyyymm === currentPeriod),
    currentBankCount = current.filter((r) => r.source_type === "bank").length,
    currentCardCount = current.filter(
      (r) => r.source_type === "credit_card",
    ).length,
    currentExpenses = -current
      .filter(
        (r) =>
          r.amount < 0 &&
          r.included_in_expenses &&
          r.transaction_type === "expense",
      )
      .reduce((s, r) => s + r.amount, 0),
    currentIncome = current
      .filter((r) => r.amount > 0 && r.transaction_type === "income")
      .reduce((s, r) => s + r.amount, 0),
    currentUnclassified = current.filter(
      (r) => !r.category || r.category === "לא מסווג",
    ).length,
    currentCardCharges = current.filter(
      (r) => r.source_type === "credit_card" && r.amount < 0,
    ),
    previousMerchantKeys = new Set(
      db
        .filter(
          (r) =>
            r.source_type === "credit_card" &&
            r.amount < 0 &&
            Number(r.yyyymm || 0) < currentPeriod,
        )
        .map((r) => merchantKey(r.description))
        .filter(Boolean),
    ),
    newMerchantCharges = currentCardCharges.filter((r) => {
      const key = merchantKey(r.description);
      return Boolean(key) && !previousMerchantKeys.has(key);
    }),
    duplicateChargeIds = new Set(
      currentCardCharges.flatMap((r, i, rows) =>
        rows.some(
          (other, j) =>
            i !== j &&
            merchantKey(r.description) === merchantKey(other.description) &&
            Math.abs(Math.abs(r.amount) - Math.abs(other.amount)) < 0.01 &&
            daysApart(r.transaction_date, other.transaction_date) <= 3,
        )
          ? [r.id]
          : [],
      ),
    ),
    duplicateCharges = currentCardCharges.filter((r) =>
      duplicateChargeIds.has(r.id),
    );
  const filters = (
    <div className="actions">
      <select value={month} onChange={(e) => setMonth(e.target.value)}>
        {months.map((m) => (
          <option key={m[0]} value={m[0]}>
            {m[1]}
          </option>
        ))}
      </select>
      <select
        value={sourceFilter}
        onChange={(e) => setSourceFilter(e.target.value)}
      >
        <option value="">כל המקורות</option>
        <option value="bank">בנק</option>
        <option value="credit_card">אשראי</option>
      </select>
      <select
        value={categoryFilter}
        onChange={(e) => {
          setCategoryFilter(e.target.value);
          setSubcategoryFilter("");
        }}
      >
        <option value="">כל הקטגוריות</option>
        {categories.map((c) => (
          <option key={c}>{c}</option>
        ))}
      </select>
      <select
        value={subcategoryFilter}
        onChange={(e) => setSubcategoryFilter(e.target.value)}
      >
        <option value="">כל תתי הקטגוריות</option>
        {subcategories.map((s) => (
          <option key={s}>{s}</option>
        ))}
      </select>
      <input
        placeholder="חיפוש..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
    </div>
  );
  const txTable = (rows: DbTx[]) => (
    <div className="tablewrap">
      {rows.length ? (
        <table>
          <thead>
            <tr>
              <th>תאריך</th>
              <th>תיאור</th>
              <th>קטגוריה</th>
              <th>תת קטגוריה</th>
              <th>אופי</th>
              <th>מקור</th>
              <th>סכום</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id}>
                <td>{r.transaction_date}</td>
                <td>{r.description}</td>
                <td>{r.category}</td>
                <td>{r.subcategory || "—"}</td>
                <td>{natureLabel(r.expense_nature)}</td>
                <td>{r.source_type === "bank" ? "בנק" : "אשראי"}</td>
                <td dir="ltr">{money(r.amount)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        <div className="empty">
          <b>אין תנועות לתקופה זו</b>
          <span>תנועות יופיעו כאן לאחר ייבוא ושמירה לתקופת החיוב המתאימה.</span>
        </div>
      )}
    </div>
  );
  return (
    <div dir="rtl">
      <header>
        <div>
          <h1>FamilyApp</h1>
          <span>משפחת תנחום</span>
        </div>
        <nav>
          <button onClick={() => setPage("home")}>דף הבית</button>
          <button onClick={() => setPage("current")}>החודש</button>
          <button onClick={() => setPage("members")}>בני משפחה</button>
          <button onClick={() => setPage("transactions")}>תנועות כספיות</button>
          <button onClick={() => setPage("summary")}>סיכומים</button>
          <button className="secondary" onClick={() => supabase.auth.signOut()}>
            יציאה
          </button>
        </nav>
      </header>
      <main>
        {page === "home" && (
          <>
            <h2>שלום משפחת תנחום 👋</h2>
            <p>ניהול המשפחה במקום אחד</p>
            <section className="cards">
              <article>
                <b>7</b>
                <span>בני משפחה</span>
              </article>
              <article>
                <b>{db.length}</b>
                <span>תנועות במאגר</span>
              </article>
            </section>
          </>
        )}
        {page === "current" && (
          <>
            <div className="title">
              <div>
                <h2>החודש — {currentLabel}</h2>
                <p>
                  {current.length} תנועות שמורות · {currentBankCount} בנק ·{" "}
                  {currentCardCount} אשראי
                </p>
              </div>
              <button onClick={() => input.current?.click()}>
                📥 ייבוא תנועות
              </button>
            </div>
            <section className="cards">
              <article>
                <b>{money(currentIncome)}</b>
                <span>הכנסות</span>
              </article>
              <article>
                <b>{money(currentExpenses)}</b>
                <span>הוצאות</span>
              </article>
              <article>
                <b>{currentUnclassified}</b>
                <span>דורשות סיווג</span>
              </article>
              <article>
                <b>{money(currentIncome - currentExpenses)}</b>
                <span>מאזן עד כה</span>
              </article>
            </section>
            <h3>תנועות {currentLabel}</h3>
            {txTable(current)}
            <section className="suspiciousSection">
              <div className="suspiciousTitle">
                <div>
                  <h3>חיובים חשודים לבדיקה</h3>
                  <p>
                    סימון אוטומטי בלבד — החיובים אינם נמחקים ואינם משתנים.
                  </p>
                </div>
                <div className="suspiciousCounts">
                  <span>{newMerchantCharges.length} בתי עסק חדשים</span>
                  <span>{duplicateCharges.length} חיובים כפולים אפשריים</span>
                </div>
              </div>

              <h3>בתי עסק חדשים החודש</h3>
              <p className="sub">
                חיובי אשראי מבתי עסק שלא נמצאו בחודשים קודמים.
              </p>
              {txTable(newMerchantCharges)}

              <h3>חיובים כפולים אפשריים</h3>
              <p className="sub">
                אותו בית עסק ואותו סכום שחויבו שוב בתוך שלושה ימים.
              </p>
              {txTable(duplicateCharges)}
            </section>
          </>
        )}
        {page === "members" && (
          <div className="members">
            {members.map((m) => (
              <article key={m}>
                <h3>{m}</h3>
              </article>
            ))}
          </div>
        )}
        {page === "transactions" && (
          <>
            <div className="title">
              <div>
                <h2>תנועות כספיות</h2>
                <p>{filtered.length} תנועות</p>
              </div>
              <button onClick={() => input.current?.click()}>📥 ייבוא</button>
            </div>
            {filters}
            <section className="cards">
              <article>
                <b>{money(income)}</b>
                <span>הכנסות</span>
              </article>
              <article>
                <b>{money(expenses)}</b>
                <span>הוצאות</span>
              </article>
              <article>
                <b>{money(income + refunds - expenses)}</b>
                <span>מאזן</span>
              </article>
            </section>
            {txTable(filtered)}
          </>
        )}
        {page === "summary" && (
          <>
            <h2>סיכומים</h2>
            {filters}
            <section className="cards">
              <article>
                <b>{money(income)}</b>
                <span>הכנסות</span>
              </article>
              <article>
                <b>{money(expenses)}</b>
                <span>הוצאות</span>
              </article>
              <article>
                <b>{money(refunds)}</b>
                <span>החזרים וזיכויים</span>
              </article>
              <article>
                <b>{money(income + refunds - expenses)}</b>
                <span>מאזן</span>
              </article>
            </section>
            <h3>הוצאות לפי אופי</h3>
            <section className="cards">
              {natureTotals.map(([n, v]: any) => (
                <article key={n}>
                  <b>{money(v)}</b>
                  <span>{natureLabel(n)}</span>
                </article>
              ))}
            </section>
            <div className="title">
              <div>
                <h3>
                  {summaryLevel === "category"
                    ? "הוצאות לפי קטגוריה ראשית"
                    : "הוצאות לפי תת קטגוריה"}
                </h3>
                <p>ממויין לפי סה״כ ההוצאה, מהגבוה לנמוך</p>
              </div>
              <button
                className="secondary"
                onClick={() =>
                  setSummaryLevel((level) =>
                    level === "category" ? "subcategory" : "category",
                  )
                }
              >
                {summaryLevel === "category"
                  ? "פירוט לפי תת קטגוריה ↓"
                  : "חזרה לקטגוריות ראשיות ↑"}
              </button>
            </div>
            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>קטגוריה</th>
                    {summaryLevel === "subcategory" && (
                      <th>תת קטגוריה</th>
                    )}
                    <th>תנועות</th>
                    <th>סה״כ</th>
                  </tr>
                </thead>
                <tbody>
                  {summary.map((r: any) => (
                    <tr key={`${r.category}|${r.subcategory}`}>
                      <td>
                        <b>{r.category}</b>
                      </td>
                      {summaryLevel === "subcategory" && (
                        <td>{r.subcategory}</td>
                      )}
                      <td>{r.count}</td>
                      <td dir="ltr">{money(r.total)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </main>
      <input
        ref={input}
        hidden
        type="file"
        accept=".xls,.xlsx"
        onChange={(e) => e.target.files?.[0] && importFile(e.target.files[0])}
      />
      {preview.length > 0 && (
        <div className="modal" role="dialog" aria-modal="true">
          <div className="dialog">
            <div className="dialogTitle">
              <div>
                <h2>אישור ייבוא — {fileName}</h2>
                <p>
                  {preview.length} תנועות ·{" "}
                  <b>{preview.filter((r) => r.dup === "new").length} חדשות</b> ·{" "}
                  {preview.filter((r) => r.dup === "existing").length} כבר
                  קיימות · {preview.filter((r) => r.dup === "possible").length}{" "}
                  חשודות
                </p>
              </div>
              <button className="secondary" onClick={() => setPreview([])}>
                סגירה
              </button>
            </div>
            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>מצב</th>
                    <th>תאריך</th>
                    <th>תיאור</th>
                    <th>קטגוריה</th>
                    <th>סכום</th>
                  </tr>
                </thead>
                <tbody>
                  {preview.slice(0, 150).map((r, i) => (
                    <tr key={i}>
                      <td>
                        {r.dup === "existing"
                          ? "✓ כבר קיים"
                          : r.dup === "possible"
                            ? "⚠ לבדיקה"
                            : "חדש"}
                      </td>
                      <td>{r.date}</td>
                      <td>{r.description}</td>
                      <td>{r.category}</td>
                      <td>{money(r.amount)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="actions">
              <button disabled={saving} onClick={save}>
                {saving ? "שומר..." : "שמירת חדשות"}
              </button>
              <button className="secondary" onClick={() => setPreview([])}>
                ביטול
              </button>
            </div>
          </div>
        </div>
      )}
      {msg && <div className="authMessage">{msg}</div>}
    </div>
  );
}
createRoot(document.getElementById("root")!).render(<App />);
