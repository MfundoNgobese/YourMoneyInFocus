import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  ResponsiveContainer, AreaChart, Area, LineChart, Line, BarChart, Bar,
  PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ComposedChart, Treemap, Sankey, Layer, Rectangle, XAxis, YAxis, CartesianGrid,
  Tooltip, ReferenceLine
} from "recharts";
import {
  LayoutDashboard, ListOrdered, PieChart as PieIcon, TrendingUp, Waves,
  Sparkles, Settings as SettingsIcon, Search, Upload, Download,
  ArrowUpRight, ArrowDownRight, Wallet, Coffee, Car, Phone, Fuel,
  ShoppingCart, Clapperboard, Plane, Receipt, Banknote, Calendar, Filter, X, Eye, EyeOff, Tag, PiggyBank
} from "lucide-react";
import * as XLSX from "xlsx";

const TXNS = [];

/* ============================ THEME ============================ */
const C = {
  ink: "#0A0B10", panel: "#14161F", panel2: "#191C27", line: "rgba(255,255,255,0.07)",
  glass: "rgba(255,255,255,0.04)", text: "#ECEDF2", sub: "#9298A6", faint: "#6B7180",
  violet: "#8B7CF6", pink: "#E879C9", amber: "#F6B14A",
  green: "#4ADE80", rose: "#FB7185", blue: "#56B6F7", teal: "#2DD4BF",
  purple: "#A78BFA", cyan: "#22D3EE", orange: "#FB923C", yellow: "#FACC15",
};
const GRAD = `linear-gradient(135deg, ${C.violet} 0%, ${C.pink} 55%, ${C.amber} 100%)`;
const CAT_COLOR = {
  "Food Delivery": "#F472B6", "Ride-hailing": "#8B7CF6", "Groceries": "#34D399",
  "Dining": "#FB923C", "Fuel": "#FACC15", "Subscriptions": "#22D3EE", "Shopping": "#38BDF8",
  "Travel": "#2DD4BF", "Airtime": "#A78BFA", "Cash": "#F59E0B", "Bank Charges": "#F87171",
  "Interest": "#FCA5A5", "Other Spend": "#C084FC",
  "Rates": "#60A5FA", "Mom": "#F9A8D4", "Investments": "#10B981", "TFSA": "#10B981",
  "Car Repayment": "#FBBF24", "Insurance": "#818CF8",
  "People": "#FB7185", "Other Payments": "#E879F9",
  "Card Purchase": "#38BDF8", "Transfer to Card": "#64748B", "Transfer from Card": "#64748B", "Other": "#A78BFA",
  "Tolls": "#F59E0B", "Parking": "#4ADE80", "Home Loan": "#EC4899", "Credit Card Payment": "#FB923C",
  "Digital Payments": "#C084FC", "Debit Order": "#93C5FD", "Transfer": "#64748B",
  "Salary": "#34D399", "Income": "#4ADE80",
};
const CAT_PALETTE = ["#8B7CF6", "#F472B6", "#34D399", "#FB923C", "#FACC15", "#22D3EE", "#38BDF8", "#A78BFA", "#2DD4BF", "#FB7185", "#818CF8", "#60A5FA", "#F59E0B", "#4ADE80", "#E879F9", "#F87171", "#10B981", "#C084FC"];
// Any category without an explicit colour gets a stable vibrant colour (no more grey defaults).
const catColor = (k) => CAT_COLOR[k] || (() => { const s = String(k || "Other"); let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0; return CAT_PALETTE[h % CAT_PALETTE.length]; })();
const TFSA_CAT = "TFSA";
const isEEAccount = (acct) => /ee-?915/i.test(String(acct || ""));

// --- CSV reading (imports the app's own export, and generic bank CSVs) ---
const normDate = (s) => {
  s = String(s || "").trim();
  let m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return `${m[1]}-${m[2]}-${m[3]}`;
  if ((m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/))) {
    let dd = m[1], mm = m[2], yy = m[3]; if (yy.length === 2) yy = "20" + yy;
    return `${yy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
  }
  const MONS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  if ((m = s.match(/^(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})/))) {
    const mo = MONS[m[2].toLowerCase()]; if (mo) return `${m[3]}-${String(mo).padStart(2, "0")}-${String(+m[1]).padStart(2, "0")}`;
  }
  const dt = new Date(s); if (!isNaN(dt)) return dt.toISOString().slice(0, 10);
  return "";
};
const parseCSVRows = (text) => {
  text = String(text).replace(/^\uFEFF/, "");
  const rows = []; let row = [], field = "", inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], nx = text[i + 1];
    if (inQ) {
      if (ch === '"' && nx === '"') { field += '"'; i++; }
      else if (ch === '"') inQ = false;
      else field += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ",") { row.push(field); field = ""; }
      else if (ch === "\r") { /* skip */ }
      else if (ch === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else field += ch;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.length && r.some(c => String(c).trim() !== ""));
};
const csvToTxns = (text) => {
  const rows = parseCSVRows(text);
  if (rows.length < 2) return [];
  const header = rows[0].map(h => String(h).trim().toLowerCase());
  const idx = (names) => { for (const n of names) { const k = header.findIndex(h => h === n || h.includes(n)); if (k >= 0) return k; } return -1; };
  const di = idx(["date"]), ai = idx(["account"]), ni = idx(["description", "details", "narrative", "reference", "merchant"]);
  const ci = idx(["category"]), oi = idx(["money out", "debit", "out (r)", "amount out", "withdrawal"]);
  const ii = idx(["money in", "credit", "in (r)", "amount in", "deposit"]), ami = idx(["amount", "value"]);
  const num = (v) => { const x = parseFloat(String(v).replace(/[^\d.\-]/g, "")); return isFinite(x) ? x : 0; };
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]; if (!row) continue;
    const get = (k) => (k >= 0 && k < row.length ? String(row[k]).trim() : "");
    const d = normDate(get(di)); if (!d) continue;
    const n = get(ni) || "Transaction";
    const acctRaw = get(ai);
    const ee = isEEAccount(acctRaw);
    const a = ee ? "EE" : (acctRaw.toLowerCase().includes("credit") || acctRaw.toUpperCase() === "CC" ? "CC" : "CHQ");
    let o = 0, i = 0;
    if (oi >= 0 || ii >= 0) { o = Math.abs(num(get(oi))); i = Math.abs(num(get(ii))); }
    else if (ami >= 0) { const v = num(get(ami)); if (v < 0) o = Math.abs(v); else i = v; }
    if (o === 0 && i === 0) continue;
    let c = get(ci); if (ee) c = TFSA_CAT; if (!c) c = "Other";
    out.push({ d, a, t: "", n, o, i, c });
  }
  return out;
};

/* ============================ HELPERS ============================ */
const mLabel = (m) => {
  if (!m) return "";
  const [y, mo] = m.split("-");
  return new Date(+y, +mo - 1, 1).toLocaleString("en", { month: "short" }) + " " + y.slice(2);
};
const mShort = (m) => {
  if (!m) return "";
  const [y, mo] = m.split("-");
  return new Date(+y, +mo - 1, 1).toLocaleString("en", { month: "short" });
};
const R = (n, dp = 0) =>
  "R" + (n < 0 ? "-" : "") + Math.abs(n).toLocaleString("en-ZA", { minimumFractionDigits: dp, maximumFractionDigits: dp });
const Rk = (n) => Math.abs(n) >= 1000 ? "R" + (n / 1000).toFixed(1) + "k" : "R" + Math.round(n);

const isExpense = (t) => t.o > 0 && !INTERNAL.has(t.c);
const isIncome = (t) => t.i > 0 && !INTERNAL.has(t.c);
// "lifestyle" = discretionary card-style categories for spend-pattern analysis
const LIFESTYLE = new Set(["Food Delivery", "Ride-hailing", "Groceries", "Dining", "Fuel", "Subscriptions", "Shopping", "Travel", "Airtime", "Other Spend"]);

const txnId = t => `${t.d}|${t.n}|${t.o}|${t.i}`;
// Resolve a transaction's effective category given reclassification overrides.
// Keys are either "id:<txnId>" (single) or "desc:<description>" (cascade to all alike).
const effCategory = (t, ov) => (ov && (ov["id:" + txnId(t)] || ov["desc:" + t.n])) || t.c;

// One collective wallet: money moving between your own accounts (transfers, card payments,
// and contributions into your own investment/TFSA account) nets to zero — it is neither
// income nor expense. Only external inflows (deposits, payments received, interest) count as
// income, and only external outflows (purchases, charges, withdrawals, payments) as expense.
const INTERNAL = new Set([
  "Transfer", "Transfer to Card", "Transfer from Card", "Transfer to Savings", "Transfer from Savings",
  "Internal Transfer", "Card Payment", "TFSA", "Investments"
]);

// Categories/labels that aren't a shop or company — kept out of "Top Merchants".
const NON_MERCHANT_CATS = new Set([
  "Bank Charges", "Interest", "Digital Payments", "Debit Order", "People",
  "Rates", "Insurance", "Home Loan", "Credit Card Payment", "Car Repayment",
  "Cash", "Other Payments", "Mom", "Salary", "Income", "TFSA", "Investments",
  "Transfer", "Transfer to Card", "Transfer from Card", "Card Payment", "Other Payments"
]);
const isRealMerchant = (name, cat) => {
  if (NON_MERCHANT_CATS.has(cat)) return false;
  if (!name || name === "—") return false;
  return !/^(payment|transfer|interest|bank|digital|debit order|eft|deposit|withdrawal|\bfee\b|immediate payment|external payment|banking app|cash )/i.test(name);
};

const merchantName = (n) => {
  if (!n) return "—";
  let s = n.split(" - ")[0].split("(")[0].trim();
  s = s.replace(/DL New Uber Eats|New Uber Eats|DL Uber Eats|DL UBER|DLO\*Uber\*Rides|Uber Eats/i, m => /eats/i.test(m) ? "Uber Eats" : "Uber");
  s = s.replace(/DL Bolt|Bolt/i, "Bolt").replace(/\bUBER\b|Uber/i, "Uber");
  if (/pick n pay|pnp/i.test(s)) s = "Pick n Pay";
  if (/woolworth/i.test(s)) s = "Woolworths";
  if (/checker/i.test(s)) s = "Checkers";
  if (/netflix/i.test(s)) s = "Netflix";
  if (/apple\.com/i.test(s)) s = "Apple";
  if (/takealot/i.test(s)) s = "Takealot";
  if (/tradingview/i.test(s)) s = "TradingView";
  if (/engen/i.test(s)) s = "Engen";
  if (/sasol/i.test(s)) s = "Sasol";
  if (/google/i.test(s)) s = "Google";
  if (/vodacom/i.test(s)) s = "Vodacom";
  else if (/\bmtn\b/i.test(s)) s = "MTN";
  else if (/cell ?c/i.test(s)) s = "Cell C";
  else if (/telkom/i.test(s)) s = "Telkom";
  else if (/airtime|prepaid mobile/i.test(s)) s = "Airtime";
  return s.length > 26 ? s.slice(0, 26) + "…" : s;
};

/* ============================ HOOKS ============================ */
function useCountUp(target, dur = 900) {
  const [v, setV] = useState(0);
  const ref = useRef(0);
  useEffect(() => {
    let raf, start;
    const from = ref.current;
    const step = (ts) => {
      if (!start) start = ts;
      const p = Math.min(1, (ts - start) / dur);
      const e = 1 - Math.pow(1 - p, 3);
      setV(from + (target - from) * e);
      if (p < 1) raf = requestAnimationFrame(step);
      else ref.current = target;
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  }, [target, dur]);
  return v;
}

function useIsNarrow(bp = 860) {
  const [m, setM] = useState(typeof window !== "undefined" ? window.innerWidth < bp : false);
  useEffect(() => {
    const on = () => setM(window.innerWidth < bp);
    on(); window.addEventListener("resize", on);
    return () => window.removeEventListener("resize", on);
  }, [bp]);
  return m;
}

/* ============================ UI PRIMITIVES ============================ */
const Glass = ({ children, className = "", style = {}, pad = true }) => (
  <div className={`rounded-2xl ${className}`} style={{
    background: C.glass, border: `1px solid ${C.line}`, backdropFilter: "blur(14px)",
    WebkitBackdropFilter: "blur(14px)", padding: pad ? 18 : 0, ...style
  }}>{children}</div>
);

const CardTitle = ({ children, right }) => (
  <div className="flex items-center justify-between mb-3">
    <div style={{ color: C.text, fontWeight: 600, fontSize: 14, letterSpacing: 0.2 }}>{children}</div>
    {right}
  </div>
);

function KPI({ label, value, sub, delta, accent = C.violet, money = true, dp = 0, hideable = false }) {
  const v = useCountUp(value);
  const up = delta != null && delta >= 0;
  const [hidden, setHidden] = useState(hideable);
  return (
    <Glass>
      <div className="flex items-start justify-between">
        <div style={{ color: C.sub, fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.6 }}>{label}</div>
        {hideable
          ? <button onClick={() => setHidden(h => !h)} aria-label="Toggle amount"
              style={{ background: "transparent", border: "none", cursor: "pointer", padding: 0, color: C.faint, display: "flex" }}>
              {hidden ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          : <span style={{ width: 8, height: 8, borderRadius: 99, background: accent, boxShadow: `0 0 10px ${accent}` }} />}
      </div>
      <div onClick={hideable ? () => setHidden(h => !h) : undefined}
        style={{
          color: C.text, fontWeight: 700, fontSize: 26, marginTop: 8, fontVariantNumeric: "tabular-nums",
          cursor: hideable ? "pointer" : "default", userSelect: "none",
          filter: hidden ? "blur(9px)" : "none", transition: "filter .25s ease",
        }}>
        {money ? R(v, dp) : v.toFixed(dp)}
      </div>
      <div className="flex items-center gap-2" style={{ marginTop: 4 }}>
        {delta != null && (
          <span className="flex items-center" style={{ color: up ? C.green : C.rose, fontSize: 12, fontWeight: 700 }}>
            {up ? <ArrowUpRight size={13} /> : <ArrowDownRight size={13} />}{Math.abs(delta).toFixed(1)}%
          </span>
        )}
        {sub && <span style={{ color: C.faint, fontSize: 12 }}>{hidden ? "tap to reveal" : sub}</span>}
      </div>
    </Glass>
  );
}

const tip = {
  contentStyle: { background: "#0F1117", border: `1px solid ${C.line}`, borderRadius: 12, color: C.text, fontSize: 12 },
  labelStyle: { color: C.sub }, itemStyle: { color: C.text },
};
const axis = { stroke: C.faint, fontSize: 11, tickLine: false, axisLine: false };

/* ============================ DERIVED DATA ============================ */
function useModel(txns, overrides={}) {
  return useMemo(() => {
    const applied = (overrides && Object.keys(overrides).length)
      ? txns.map(t => { const c = effCategory(t, overrides); return c === t.c ? t : { ...t, c }; })
      : txns;
    // Treat all accounts as one wallet: an outflow from one account matched by an equal
    // inflow into another account (within a few days) is an internal transfer, so it nets
    // to zero — no new money entered or left. These are excluded from income and expense.
    const transferIds = new Set();
    const insByAmt = new Map();
    applied.forEach(t => { if (t.i > 0) { const key = Math.round(t.i * 100); (insByAmt.get(key) || insByAmt.set(key, []).get(key)).push(t); } });
    const usedIn = new Set();
    applied.forEach(o => {
      if (!(o.o > 0)) return;
      const cands = insByAmt.get(Math.round(o.o * 100));
      if (!cands) return;
      let best = null, bestDD = 99;
      for (const inn of cands) {
        if (usedIn.has(inn) || inn.a === o.a) continue;
        const dd = Math.abs((new Date(inn.d) - new Date(o.d)) / 86400000);
        if (dd <= 4 && dd < bestDD) { best = inn; bestDD = dd; }
      }
      if (best) { transferIds.add(txnId(o)); transferIds.add(txnId(best)); usedIn.add(best); }
    });
    const inc = (t) => isIncome(t) && !transferIds.has(txnId(t));
    const exp = (t) => isExpense(t) && !transferIds.has(txnId(t));
    // month axis derived from the ACTUAL data (so imports of any period populate)
    const months = [...new Set(applied.map(t => (t.d || "").slice(0, 7)).filter(Boolean))].sort();
    const byMonth = {};
    months.forEach(m => byMonth[m] = { m, income: 0, expense: 0, cardSpend: 0, salary: 0, fees: 0, savings: 0 });
    const catTot = {}, catMonth = {}, merch = {}, daily = {}, lifeMonth = {};
    let biggest = { o: 0, n: "—", d: "" };
    let wkndOut = 0, wkdayOut = 0, wkndN = 0, wkdayN = 0;
    const salaryHist = [];

    applied.forEach(t => {
      const m = (t.d || "").slice(0, 7);
      const bm = byMonth[m]; if (!bm) return;
      if (inc(t)) { bm.income += t.i; if (t.c === "Salary") { bm.salary += t.i; } }
      if (exp(t)) {
        bm.expense += t.o;
        if (t.a === "CC") bm.cardSpend += t.o;
        if (t.c === "Bank Charges" || t.c === "Interest") bm.fees += t.o;
        catTot[t.c] = (catTot[t.c] || 0) + t.o;
        (catMonth[t.c] = catMonth[t.c] || {})[m] = (catMonth[t.c]?.[m] || 0) + t.o;
        const day = new Date(t.d).getDay();
        if (day === 0 || day === 6) { wkndOut += t.o; wkndN++; } else { wkdayOut += t.o; wkdayN++; }
        daily[t.d] = (daily[t.d] || 0) + t.o;
        if (LIFESTYLE.has(t.c)) (lifeMonth[t.c] = lifeMonth[t.c] || {})[m] = (lifeMonth[t.c]?.[m] || 0) + t.o;
        if (t.a === "CC" && t.o > biggest.o) biggest = t;
        const mn = merchantName(t.n);
        if (isRealMerchant(mn, t.c)) merch[mn] = (merch[mn] || 0) + t.o;
      }
    });
    months.forEach(m => { byMonth[m].savings = byMonth[m].income - byMonth[m].expense; });
    months.forEach(m => { if (byMonth[m].salary > 0) salaryHist.push({ m, v: byMonth[m].salary }); });

    const monthly = months.map(m => byMonth[m]);
    let cum = 0; const cumSavings = monthly.map(r => ({ m: r.m, v: (cum += r.savings) }));
    // rolling 3-mo avg expense
    const roll = monthly.map((r, i) => {
      const s = monthly.slice(Math.max(0, i - 2), i + 1);
      return { m: r.m, expense: r.expense, avg: s.reduce((a, b) => a + b.expense, 0) / s.length };
    });
    const catList = Object.entries(catTot).map(([k, v]) => ({ k, v })).filter(x => isFinite(x.v) && x.v > 0).sort((a, b) => b.v - a.v);
    const merchList = Object.entries(merch).map(([k, v]) => ({ k, v })).filter(x => isFinite(x.v) && x.v > 0).sort((a, b) => b.v - a.v);
    const totals = {
      income: monthly.reduce((a, b) => a + b.income, 0),
      expense: monthly.reduce((a, b) => a + b.expense, 0),
      salary: monthly.reduce((a, b) => a + b.salary, 0),
      cardSpend: monthly.reduce((a, b) => a + b.cardSpend, 0),
    };
    totals.savings = totals.income - totals.expense;
    totals.rate = totals.income ? (totals.savings / totals.income) * 100 : 0;
    const dayCount = Object.keys(daily).length;
    const avgDaily = dayCount ? totals.expense / dayCount : 0;
    return { months, monthly, catTot, catMonth, lifeMonth, merchList, catList, daily, biggest, cumSavings, roll, totals, salaryHist, avgDaily, wknd: { wkndOut, wkdayOut, wkndN, wkdayN } };
  }, [txns, overrides]);
}

/* ============================ CHARTS ============================ */
const H = ({ h = 240, children }) => <div style={{ width: "100%", height: h }}><ResponsiveContainer>{children}</ResponsiveContainer></div>;

function CashFlowChart({ monthly }) {
  const data = monthly.map(r => ({ name: mShort(r.m), income: Math.round(r.income), expense: Math.round(r.expense), net: Math.round(r.savings) }));
  return (
    <H h={260}>
      <ComposedChart data={data} margin={{ left: 4, right: 8, top: 6 }}>
        <defs>
          <linearGradient id="gInc" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.green} stopOpacity={0.5} /><stop offset="100%" stopColor={C.green} stopOpacity={0} /></linearGradient>
        </defs>
        <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
        <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
        <Tooltip {...tip} formatter={(v, n) => [R(v), n]} />
        <Area type="monotone" dataKey="income" stroke={C.green} fill="url(#gInc)" strokeWidth={2} name="Income" />
        <Bar dataKey="expense" fill={C.rose} radius={[4, 4, 0, 0]} name="Expenses" barSize={14} opacity={0.85} />
        <Line type="monotone" dataKey="net" stroke={C.amber} strokeWidth={2.5} dot={false} name="Net savings" />
      </ComposedChart>
    </H>
  );
}

function Donut({ catList, onPick }) {
  const top = catList.slice(0, 9);
  const rest = catList.slice(9).reduce((a, b) => a + b.v, 0);
  const data = [...top, ...(rest ? [{ k: "Other", v: rest }] : [])];
  const labelCount = data.length < 6 ? data.length : Math.ceil(data.length / 2);
  const renderLabel = ({ cx, cy, midAngle, outerRadius, index, name, percent }) => {
    if (index >= labelCount || percent < 0.02) return null;
    const RAD = Math.PI / 180;
    const r = outerRadius + 13;
    const x = cx + r * Math.cos(-midAngle * RAD);
    const y = cy + r * Math.sin(-midAngle * RAD);
    const nm = name.length > 13 ? name.slice(0, 12) + "…" : name;
    return (
      <text x={x} y={y} textAnchor={x > cx ? "start" : "end"} dominantBaseline="central" fontSize={10.5} fill={C.sub}>
        {nm} {(percent * 100).toFixed(0)}%
      </text>
    );
  };
  return (
    <H h={290}>
      <PieChart margin={{ top: 10, bottom: 10, left: 10, right: 10 }}>
        <Pie data={data} dataKey="v" nameKey="k" innerRadius={54} outerRadius={82} paddingAngle={2} stroke="none"
          labelLine={false} label={renderLabel} onClick={(d) => onPick && onPick(d.k)}>
          {data.map((d, i) => <Cell key={i} cursor="pointer" fill={catColor(d.k)} />)}
        </Pie>
        <Tooltip {...tip} formatter={(v, n) => [R(v), n]} />
      </PieChart>
    </H>
  );
}

function HeatGrid({ daily }) {
  // continuous weeks grid — range derived from the actual spending days
  const days = Object.keys(daily);
  if (!days.length) return <div style={{ color: C.faint, fontSize: 12 }}>No spending data to map yet.</div>;
  const min = days.reduce((a, b) => a < b ? a : b);
  const max = days.reduce((a, b) => a > b ? a : b);
  const start = new Date(min); start.setDate(start.getDate() - ((start.getDay() + 6) % 7)); // mon start
  const end = new Date(max);
  const weeks = [];
  let cur = new Date(start);
  const vals = Object.values(daily);
  const mx = Math.max(...vals);
  const shade = (v) => {
    if (!v) return "rgba(255,255,255,0.05)";
    const t = Math.min(1, Math.sqrt(v / mx));
    const a = 0.18 + t * 0.82;
    return `rgba(139,124,246,${a})`;
  };
  while (cur <= end) {
    const w = [];
    for (let i = 0; i < 7; i++) {
      const iso = cur.toISOString().slice(0, 10);
      w.push({ iso, v: daily[iso] || 0, inRange: iso >= min && iso <= max });
      cur.setDate(cur.getDate() + 1);
    }
    weeks.push(w);
  }
  const pitch = 15; // 12px cell + 3px gap
  const monthMarks = [];
  let lastM = "";
  weeks.forEach((w, i) => {
    const firstIn = w.find(d => d.inRange);
    if (firstIn) {
      const mk = firstIn.iso.slice(0, 7);
      if (mk !== lastM) { monthMarks.push({ i, label: mShort(mk) }); lastM = mk; }
    }
  });
  return (
    <div>
      <div style={{ overflowX: "auto", paddingBottom: 4 }}>
        <div style={{ minWidth: weeks.length * pitch }}>
          <div style={{ position: "relative", height: 16 }}>
            {monthMarks.map((m, idx) => (
              <span key={idx} style={{ position: "absolute", left: m.i * pitch, top: 0, color: C.faint, fontSize: 10, whiteSpace: "nowrap" }}>{m.label}</span>
            ))}
          </div>
          <div className="flex" style={{ gap: 3 }}>
            {weeks.map((w, i) => (
              <div key={i} className="flex flex-col" style={{ gap: 3 }}>
                {w.map((d, j) => (
                  <div key={j} title={d.inRange ? `${d.iso}: ${R(d.v)}` : ""}
                    style={{ width: 12, height: 12, borderRadius: 3, background: d.inRange ? shade(d.v) : "transparent" }} />
                ))}
              </div>
            ))}
          </div>
        </div>
      </div>
      <div className="flex items-center gap-2" style={{ marginTop: 10, color: C.faint, fontSize: 11 }}>
        Less
        {[0.05, 0.3, 0.55, 0.8, 1].map((a, i) => <span key={i} style={{ width: 12, height: 12, borderRadius: 3, background: i === 0 ? "rgba(255,255,255,0.05)" : `rgba(139,124,246,${a})` }} />)}
        More
      </div>
    </div>
  );
}

function TopBars({ data, color, fmt = R, h = 260 }) {
  return (
    <H h={h}>
      <BarChart data={data} layout="vertical" margin={{ left: 8, right: 16 }}>
        <XAxis type="number" {...axis} tickFormatter={Rk} hide />
        <YAxis type="category" dataKey="k" {...axis} width={120} tick={{ fill: C.sub, fontSize: 11 }} />
        <Tooltip {...tip} formatter={(v) => [fmt(v), "Spent"]} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
        <Bar dataKey="v" radius={[0, 6, 6, 0]} barSize={16}>
          {data.map((d, i) => <Cell key={i} fill={color || catColor(d.k)} />)}
        </Bar>
      </BarChart>
    </H>
  );
}

function MoneyFlowSankey({ model }) {
  const m = model.totals;
  const billsCats = ["Rates", "Mom", TFSA_CAT, "Debit Order", "Car Repayment", "Insurance", "People", "Home Loan", "Other Payments", "Airtime"];
  const bills = billsCats.reduce((a, c) => a + (model.catTot[c] || 0), 0);
  const card = m.cardSpend;
  const otherChq = m.expense - bills - card;
  const save = Math.max(0, m.savings);
  const nodes = [{ name: "Income" }, { name: "Living & Bills" }, { name: "Card Lifestyle" }, { name: "Other" }, { name: "Saved & Invested" }];
  const links = [
    { source: 0, target: 1, value: Math.round(bills) },
    { source: 0, target: 2, value: Math.round(card) },
    { source: 0, target: 3, value: Math.round(Math.max(0, otherChq)) },
    { source: 0, target: 4, value: Math.round(save) },
  ].filter(l => l.value > 0);
  return (
    <H h={260}>
      <Sankey data={{ nodes, links }} nodePadding={26} nodeWidth={12}
        link={{ stroke: C.violet, strokeOpacity: 0.18 }} margin={{ left: 4, right: 80, top: 6, bottom: 6 }}
        node={<SankeyNode />}>
        <Tooltip {...tip} formatter={(v) => R(v)} />
      </Sankey>
    </H>
  );
}
const SankeyNode = (p) => {
  const { x, y, width, height, index, payload } = p;
  const cols = [C.green, "#60A5FA", C.pink, "#94A3B8", C.amber];
  return (
    <Layer key={index}>
      <Rectangle x={x} y={y} width={width} height={height} fill={cols[index % cols.length]} radius={2} />
      <text x={x + width + 8} y={y + height / 2} textAnchor="start" dominantBaseline="middle" fill={C.sub} fontSize={11}>{payload.name}</text>
    </Layer>
  );
};

function TrendLine({ series, keys, h = 240, fmt = R }) {
  return (
    <H h={h}>
      <LineChart data={series} margin={{ left: 4, right: 8, top: 6 }}>
        <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
        <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={42} />
        <Tooltip {...tip} formatter={(v, n) => [fmt(v), n]} />
        {keys.map(k => <Line key={k.k} type="monotone" dataKey={k.k} stroke={k.c} strokeWidth={2.2} dot={false} name={k.label} />)}
      </LineChart>
    </H>
  );
}

/* ============================ INSIGHTS ============================ */
function buildInsights(model) {
  const out = [];
  const cm = (c, m) => (model.catMonth[c]?.[m] || 0);
  const fdMar = cm("Food Delivery", "2026-03"), fdJun = cm("Food Delivery", "2026-06");
  if (fdMar && fdJun) out.push({
    icon: Coffee, color: C.pink, title: "Uber Eats: March binge, June reset",
    body: `Food delivery hit ${R(fdMar)} in March — the year's peak — then fell to ${R(fdJun)} in June, a ${(100 * (1 - fdJun / fdMar)).toFixed(0)}% drop. March and April together (${R(fdMar + cm("Food Delivery", "2026-04"))}) were your splurge window.`
  });
  const dec = model.monthly.find(r => r.m === "2025-12")?.savings ?? 0;
  const feb = model.monthly.find(r => r.m === "2026-02")?.savings ?? 0;
  if (dec || feb) out.push({
    icon: Wallet, color: feb < 0 ? C.rose : C.green, title: "December vs February",
    body: `December netted ${R(dec)} ${dec >= 0 ? "saved" : "overspent"}, while February ${feb >= 0 ? "saved " + R(feb) : "ran " + R(Math.abs(feb)) + " over"} — February carried heavy grocery (${R(cm("Groceries", "2026-02"))}) and dining loads against a leaner month.`
  });
  // airtime trend
  const air = model.lifeMonth["Airtime"] || {};
  const aVals = model.months.map(m => air[m] || 0);
  const aFirst = aVals.slice(0, 4).reduce((a, b) => a + b, 0) / 4, aLast = aVals.slice(-4).reduce((a, b) => a + b, 0) / 4;
  if (aVals.some(v => v > 0)) out.push({
    icon: Phone, color: C.purple, title: "Airtime is creeping up",
    body: `Airtime averaged ${R(aFirst)}/mo early on and ${R(aLast)}/mo recently — ${aLast > aFirst ? "up " : "down "}${Math.abs(((aLast - aFirst) / (aFirst || 1)) * 100).toFixed(0)}%. May spiked to ${R(air["2026-05"] || 0)}.`
  });
  // fuel
  const fuel = model.catMonth["Fuel"] || {};
  const fuelTot = Object.values(fuel).reduce((a, b) => a + b, 0);
  if (fuelTot > 0) out.push({ icon: Fuel, color: C.yellow, title: "Fuel", body: `${R(fuelTot)} on fuel across ${model.months.length} month${model.months.length === 1 ? "" : "s"}.` });
  // grocery inflation
  const gro = model.catMonth["Groceries"] || {};
  const gVals = model.months.map(m => gro[m] || 0);
  const gMax = Math.max(0, ...gVals), gMaxM = model.months[gVals.indexOf(gMax)];
  if (gMax > 0) out.push({ icon: ShoppingCart, color: C.green, title: "Grocery spend swings hard", body: `Groceries peaked at ${R(gMax)} in ${mLabel(gMaxM)} and averaged ${R(gVals.reduce((a, b) => a + b, 0) / (model.months.length || 1))}/mo. Big Pick n Pay & Woolworths baskets drive the spikes.` });
  // ride-hailing
  const rh = Object.values(model.catMonth["Ride-hailing"] || {}).reduce((a, b) => a + b, 0);
  if (rh > 0) out.push({ icon: Car, color: C.violet, title: "Ride-hailing is your transport", body: `${R(rh)} on Uber + Bolt over ${model.months.length} months (${R(rh / (model.months.length || 1))}/mo) — easily your most frequent expense by count.` });
  // largest
  const b = model.biggest;
  if (b.o > 0) out.push({ icon: Receipt, color: C.blue, title: "Largest single purchase", body: `${R(b.o)} — ${merchantName(b.n)} on ${b.d}.` });
  // weekend vs weekday
  const w = model.wknd;
  const wkndAvg = w.wkndOut / (w.wkndN || 1), wkdayAvg = w.wkdayOut / (w.wkdayN || 1);
  if (w.wkndN + w.wkdayN > 0) out.push({ icon: Calendar, color: C.orange, title: "Weekends cost more per swipe", body: `Average weekend transaction ${R(wkndAvg)} vs ${R(wkdayAvg)} on weekdays. Total weekend spend ${R(w.wkndOut)}.` });
  // avg daily
  if (model.avgDaily > 0) out.push({ icon: Banknote, color: C.teal, title: "Average daily spend", body: `${R(model.avgDaily)} across every day money moved.${model.salaryHist.length ? ` Salary lands ${R(model.totals.salary / model.salaryHist.length)}/mo on average.` : ""}` });
  return out;
}

/* ============================ PAGES ============================ */
function Dashboard({ model, month, setMonth, goCat }) {
  const t = model.totals;
  const narrow = useIsNarrow();
  const two = narrow ? "1fr" : "1fr 1fr";
  return (
    <div className="flex flex-col" style={{ gap: 16 }}>
      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr 1fr" : "repeat(auto-fit,minmax(190px,1fr))", gap: narrow ? 10 : 14 }}>
        <KPI label="Total Income" value={t.income} accent={C.green} sub="12 months" hideable />
        <KPI label="Total Expenses" value={t.expense} accent={C.rose} sub="true outflows" />
        <KPI label="Saved & Invested" value={t.savings} accent={C.amber} sub="income − expenses" hideable />
        <KPI label="Savings Rate" value={t.rate} accent={C.violet} money={false} dp={1} sub="% of income" />
      </div>

      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr" : "1.55fr 1fr", gap: 16 }}>
        <Glass>
          <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>Income · Expenses · Net</span>}>Cash Flow</CardTitle>
          <CashFlowChart monthly={model.monthly} />
        </Glass>
        <Glass>
          <CardTitle>Where money goes</CardTitle>
          <Donut catList={model.catList} onPick={goCat} />
        </Glass>
      </div>

      <div className="grid" style={{ gridTemplateColumns: two, gap: 16 }}>
        <Glass>
          <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>annual flow</span>}>Money Flow</CardTitle>
          <MoneyFlowSankey model={model} />
        </Glass>
        <Glass>
          <CardTitle>Biggest Categories</CardTitle>
          <TopBars data={model.catList.slice(0, 8)} />
        </Glass>
      </div>

      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>daily external spend · {Object.keys(model.daily).length} active days</span>}>Spending Heatmap</CardTitle>
        <HeatGrid daily={model.daily} />
      </Glass>

      <div className="grid" style={{ gridTemplateColumns: two, gap: 16 }}>
        <Glass>
          <CardTitle>Cumulative Savings</CardTitle>
          <H h={220}>
            <AreaChart data={model.cumSavings.map(r => ({ name: mShort(r.m), v: Math.round(r.v) }))} margin={{ left: 4, right: 8, top: 6 }}>
              <defs><linearGradient id="cum" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.amber} stopOpacity={0.45} /><stop offset="100%" stopColor={C.amber} stopOpacity={0} /></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
              <Tooltip {...tip} formatter={(v) => [R(v), "Cumulative"]} />
              <Area type="monotone" dataKey="v" stroke={C.amber} fill="url(#cum)" strokeWidth={2.5} />
            </AreaChart>
          </H>
        </Glass>
        <Glass>
          <CardTitle>Top Merchants</CardTitle>
          <TopBars data={model.merchList.slice(0, 8)} color={C.blue} />
        </Glass>
      </div>
    </div>
  );
}

function Trends({ model }) {
  const narrow = useIsNarrow();
  const series = model.months.map(m => {
    const o = { name: mShort(m) };
    ["Food Delivery", "Ride-hailing", "Groceries", "Dining", "Fuel", "Subscriptions", "Airtime", "Shopping"].forEach(c => o[c] = Math.round(model.catMonth[c]?.[m] || 0));
    return o;
  });
  const fd = series.map(s => ({ name: s.name, v: s["Food Delivery"] }));
  const fdMax = Math.max(0, ...fd.map(d => d.v));
  return (
    <div className="flex flex-col" style={{ gap: 16 }}>
      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>March peak vs June low</span>}>Uber Eats — monthly</CardTitle>
        <H h={240}>
          <BarChart data={fd} margin={{ left: 4, right: 8, top: 6 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
            <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={42} />
            <Tooltip {...tip} formatter={(v) => [R(v), "Uber Eats"]} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
            <Bar dataKey="v" radius={[5, 5, 0, 0]}>
              {fd.map((d, i) => <Cell key={i} fill={d.v === fdMax ? C.pink : "rgba(232,121,201,0.35)"} />)}
            </Bar>
          </BarChart>
        </H>
      </Glass>

      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr" : "1fr 1fr", gap: 16 }}>
        <Glass><CardTitle>Airtime trend</CardTitle><TrendLine series={series} keys={[{ k: "Airtime", c: C.purple, label: "Airtime" }]} /></Glass>
        <Glass><CardTitle>Fuel trend</CardTitle><TrendLine series={series} keys={[{ k: "Fuel", c: C.yellow, label: "Fuel" }]} /></Glass>
      </div>

      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>grocery inflation & swings</span>}>Groceries — monthly</CardTitle>
        <H h={220}>
          <AreaChart data={series} margin={{ left: 4, right: 8, top: 6 }}>
            <defs><linearGradient id="gro" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.green} stopOpacity={0.45} /><stop offset="100%" stopColor={C.green} stopOpacity={0} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
            <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
            <Tooltip {...tip} formatter={(v) => [R(v), "Groceries"]} />
            <Area type="monotone" dataKey="Groceries" stroke={C.green} fill="url(#gro)" strokeWidth={2.5} />
          </AreaChart>
        </H>
      </Glass>

      <Glass>
        <CardTitle>Lifestyle categories — all trends</CardTitle>
        <TrendLine h={300} series={series} keys={[
          { k: "Food Delivery", c: C.pink, label: "Food Delivery" }, { k: "Ride-hailing", c: C.violet, label: "Ride-hailing" },
          { k: "Dining", c: C.orange, label: "Dining" }, { k: "Shopping", c: C.blue, label: "Shopping" },
          { k: "Subscriptions", c: C.cyan, label: "Subscriptions" },
        ]} />
      </Glass>

      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>actual vs rolling 3-mo avg</span>}>Total Expenses — smoothed</CardTitle>
        <H h={240}>
          <ComposedChart data={model.roll.map(r => ({ name: mShort(r.m), expense: Math.round(r.expense), avg: Math.round(r.avg) }))} margin={{ left: 4, right: 8, top: 6 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
            <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
            <Tooltip {...tip} formatter={(v, n) => [R(v), n]} />
            <Bar dataKey="expense" fill="rgba(86,182,247,0.35)" radius={[4, 4, 0, 0]} name="Monthly" barSize={16} />
            <Line type="monotone" dataKey="avg" stroke={C.amber} strokeWidth={2.5} dot={false} name="3-mo avg" />
          </ComposedChart>
        </H>
      </Glass>
    </div>
  );
}

function Categories({ model, goTx }) {
  const narrow = useIsNarrow();
  const tree = model.catList.slice(0, 14).map(d => ({ name: d.k, size: Math.round(d.v), fill: catColor(d.k) })).filter(d => isFinite(d.size) && d.size > 0);
  const radarData = ["Food Delivery", "Ride-hailing", "Groceries", "Dining", "Fuel", "Subscriptions", "Shopping", "Airtime"]
    .map(c => ({ cat: c, v: Math.round(model.catTot[c] || 0) }));
  return (
    <div className="flex flex-col" style={{ gap: 16 }}>
      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr" : "1.4fr 1fr", gap: 16 }}>
        <Glass>
          <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>size = total spend</span>}>Category Treemap</CardTitle>
          {tree.length ? (
            <H h={300}>
              <Treemap data={tree} dataKey="size" stroke={C.ink} content={<TreeCell />} />
            </H>
          ) : <div style={{ color: C.faint, fontSize: 12, padding: "40px 0", textAlign: "center" }}>No categorised spending yet.</div>}
        </Glass>
        <Glass>
          <CardTitle>Lifestyle profile</CardTitle>
          <H h={300}>
            <RadarChart data={radarData} outerRadius={110}>
              <PolarGrid stroke={C.line} />
              <PolarAngleAxis dataKey="cat" tick={{ fill: C.sub, fontSize: 10 }} />
              <PolarRadiusAxis tick={false} axisLine={false} />
              <Radar dataKey="v" stroke={C.violet} fill={C.violet} fillOpacity={0.35} />
              <Tooltip {...tip} formatter={(v) => [R(v), "Spent"]} />
            </RadarChart>
          </H>
        </Glass>
      </div>
      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>click a row to see transactions</span>}>All categories</CardTitle>
        <div className="flex flex-col" style={{ gap: 6 }}>
          {model.catList.map((d, i) => {
            const pct = (d.v / model.totals.expense) * 100;
            return (
              <button key={i} onClick={() => goTx(d.k)} className="flex items-center" style={{ gap: 12, background: "transparent", border: "none", cursor: "pointer", padding: "6px 4px" }}>
                <span style={{ width: 10, height: 10, borderRadius: 3, background: catColor(d.k) }} />
                <span style={{ color: C.text, fontSize: 13, width: 150, textAlign: "left" }}>{d.k}</span>
                <div style={{ flex: 1, height: 8, background: "rgba(255,255,255,0.05)", borderRadius: 99, overflow: "hidden" }}>
                  <div style={{ width: pct + "%", height: "100%", background: catColor(d.k), borderRadius: 99 }} />
                </div>
                <span style={{ color: C.sub, fontSize: 12, width: 84, textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{R(d.v)}</span>
                <span style={{ color: C.faint, fontSize: 11, width: 44, textAlign: "right" }}>{pct.toFixed(1)}%</span>
              </button>
            );
          })}
        </div>
      </Glass>
    </div>
  );
}
const TreeCell = (p) => {
  const { x, y, width, height, name, size, fill } = p;
  if (width < 0 || height < 0) return null;
  return (
    <g>
      <rect x={x} y={y} width={width} height={height} fill={fill} stroke={C.ink} strokeWidth={2} rx={6} />
      {width > 60 && height > 30 && (
        <>
          <text x={x + 8} y={y + 20} fill="#0A0B10" fontSize={12} fontWeight={700}>{name}</text>
          <text x={x + 8} y={y + 36} fill="rgba(10,11,16,0.75)" fontSize={11}>{Rk(size)}</text>
        </>
      )}
    </g>
  );
};

function CashFlow({ model }) {
  const narrow = useIsNarrow();
  // waterfall of monthly net
  let run = 0;
  const wf = model.monthly.map(r => {
    const base = run; run += r.savings;
    return { name: mShort(r.m), base: Math.round(Math.min(base, run)), delta: Math.round(Math.abs(r.savings)), pos: r.savings >= 0, end: Math.round(run) };
  });
  return (
    <div className="flex flex-col" style={{ gap: 16 }}>
      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr" : "1fr 1fr", gap: 16 }}>
        <Glass>
          <CardTitle>Income vs Expenses</CardTitle>
          <H h={240}>
            <BarChart data={model.monthly.map(r => ({ name: mShort(r.m), income: Math.round(r.income), expense: Math.round(r.expense) }))} margin={{ left: 4, right: 8, top: 6 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
              <Tooltip {...tip} formatter={(v, n) => [R(v), n]} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
              <Bar dataKey="income" fill={C.green} radius={[4, 4, 0, 0]} barSize={12} name="Income" />
              <Bar dataKey="expense" fill={C.rose} radius={[4, 4, 0, 0]} barSize={12} name="Expenses" />
            </BarChart>
          </H>
        </Glass>
        <Glass>
          <CardTitle>Savings rate</CardTitle>
          <H h={240}>
            <LineChart data={model.monthly.map(r => ({ name: mShort(r.m), rate: r.income ? Math.round((r.savings / r.income) * 100) : 0 }))} margin={{ left: 4, right: 8, top: 6 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={(v) => v + "%"} width={40} />
              <ReferenceLine y={0} stroke={C.faint} />
              <Tooltip {...tip} formatter={(v) => [v + "%", "Savings rate"]} />
              <Line type="monotone" dataKey="rate" stroke={C.violet} strokeWidth={2.5} dot={{ r: 3, fill: C.violet }} />
            </LineChart>
          </H>
        </Glass>
      </div>
      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>running balance of net savings</span>}>Savings Waterfall</CardTitle>
        <H h={260}>
          <ComposedChart data={wf} margin={{ left: 4, right: 8, top: 6 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
            <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
            <Tooltip {...tip} formatter={(v, n, p) => n === "delta" ? [R(p.payload.pos ? p.payload.delta : -p.payload.delta), "Net month"] : null} />
            <Bar dataKey="base" stackId="a" fill="transparent" />
            <Bar dataKey="delta" stackId="a" radius={[4, 4, 0, 0]}>
              {wf.map((d, i) => <Cell key={i} fill={d.pos ? C.green : C.rose} />)}
            </Bar>
          </ComposedChart>
        </H>
      </Glass>
      <Glass>
        <CardTitle right={model.salaryHist.length ? <span style={{ color: C.faint, fontSize: 12 }}>{R(model.totals.salary / model.salaryHist.length)}/mo avg</span> : null}>Salary history</CardTitle>
        {model.salaryHist.length ? (
          <H h={200}>
            <BarChart data={model.salaryHist.map(r => ({ name: mShort(r.m), v: Math.round(r.v) }))} margin={{ left: 4, right: 8, top: 6 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
              <XAxis dataKey="name" {...axis} /><YAxis {...axis} tickFormatter={Rk} width={44} />
              <Tooltip {...tip} formatter={(v) => [R(v), "Salary"]} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
              <Bar dataKey="v" fill={C.teal} radius={[5, 5, 0, 0]} />
            </BarChart>
          </H>
        ) : (
          <div style={{ color: C.faint, fontSize: 12.5, lineHeight: 1.5, padding: "28px 4px" }}>
            No salary detected in this data. Tag a recurring income transaction as <b style={{ color: C.sub }}>Salary</b> on the Transactions tab (use "apply to all matching") and it will chart here.
          </div>
        )}
      </Glass>
    </div>
  );
}

function Insights({ model }) {
  const items = buildInsights(model);
  return (
    <div className="grid" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: 16 }}>
      {items.map((it, i) => {
        const Icon = it.icon;
        return (
          <Glass key={i}>
            <div className="flex items-start" style={{ gap: 12 }}>
              <div style={{ width: 38, height: 38, borderRadius: 12, display: "grid", placeItems: "center", background: it.color + "22", flexShrink: 0 }}>
                <Icon size={18} color={it.color} />
              </div>
              <div>
                <div style={{ color: C.text, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{it.title}</div>
                <div style={{ color: C.sub, fontSize: 13, lineHeight: 1.5 }}>{it.body}</div>
              </div>
            </div>
          </Glass>
        );
      })}
    </div>
  );
}

const ALL_CATS = [...new Set(Object.keys(CAT_COLOR))].sort();

function ClassifyModal({ txn, currentCat, matchCount, cats, onSave, onClose }) {
  const list = cats && cats.length ? cats : ALL_CATS;
  const [cat, setCat] = useState(currentCat);
  const [custom, setCustom] = useState("");
  const [all, setAll] = useState(matchCount > 1);
  const apply = () => { onSave(custom.trim() || cat, all ? "desc" : "single"); onClose(); };
  return (
    <div onClick={onClose} style={{
      position:"fixed",inset:0,background:"rgba(0,0,0,.75)",zIndex:300,
      display:"flex",alignItems:"center",justifyContent:"center",padding:16
    }}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:C.panel2,border:`1px solid ${C.line}`,borderRadius:18,
        padding:22,width:"100%",maxWidth:430
      }}>
        <div style={{fontWeight:700,fontSize:15,color:C.text,marginBottom:14}}>Reclassify Transaction</div>
        <div style={{background:"rgba(255,255,255,0.04)",borderRadius:12,padding:"10px 14px",marginBottom:14}}>
          <div style={{color:C.text,fontSize:13,fontWeight:600,marginBottom:3}}>{txn.n}</div>
          <div style={{color:C.faint,fontSize:11.5}}>{txn.d} · {txn.o>0?R(txn.o,2)+" out":"+"+R(txn.i,2)+" in"}</div>
        </div>
        <div style={{color:C.sub,fontSize:12,marginBottom:8}}>
          Current: <b style={{color:catColor(currentCat)}}>{currentCat}</b>
        </div>
        <select value={cat} onChange={e=>setCat(e.target.value)} style={{
          width:"100%",background:C.panel,border:`1px solid ${C.line}`,color:C.text,
          fontSize:13,padding:"9px 12px",borderRadius:10,outline:"none",marginBottom:10,boxSizing:"border-box"
        }}>
          {list.map(c=><option key={c} value={c} style={{background:C.panel2}}>{c}</option>)}
        </select>
        <input value={custom} onChange={e=>setCustom(e.target.value)} placeholder="Or type a custom category name…"
          style={{width:"100%",background:C.panel,border:`1px solid ${C.line}`,color:C.text,
            fontSize:13,padding:"9px 12px",borderRadius:10,outline:"none",marginBottom:12,boxSizing:"border-box"}} />
        {matchCount > 1 && (
          <label className="flex items-center" style={{gap:9,marginBottom:12,cursor:"pointer",userSelect:"none"}}>
            <input type="checkbox" checked={all} onChange={e=>setAll(e.target.checked)}
              style={{width:16,height:16,accentColor:C.violet,cursor:"pointer"}} />
            <span style={{color:C.text,fontSize:12.5}}>
              Apply to all <b>{matchCount}</b> matching entries (past &amp; future)
            </span>
          </label>
        )}
        <div style={{color:C.faint,fontSize:11,lineHeight:1.5,marginBottom:16}}>
          Saved on this device and reapplied to matching transactions when you import new statements.
        </div>
        <div className="flex" style={{gap:10}}>
          <button onClick={apply} style={{flex:1,padding:"10px",borderRadius:10,border:"none",
            background:C.violet,color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>Apply</button>
          <button onClick={onClose} style={{flex:1,padding:"10px",borderRadius:10,
            border:`1px solid ${C.line}`,background:"transparent",color:C.sub,fontSize:13,cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

function Transactions({ txns, overrides={}, onReclassify, presetCat, presetClear }) {
  const [q, setQ] = useState("");
  const [acc, setAcc] = useState("All");
  const [cat, setCat] = useState(presetCat || "All");
  const [sort, setSort] = useState({ k: "d", dir: -1 });
  const [limit, setLimit] = useState(60);
  const [classifyTxn, setClassifyTxn] = useState(null);

  const effCat = t => effCategory(t, overrides);
  const matchCount = t => txns.filter(x => x.n === t.n).length;

  useEffect(() => { if (presetCat) { setCat(presetCat); } }, [presetCat]);
  const cats = useMemo(() => ["All", ...[...new Set(txns.map(t => effCat(t)))].sort()], [txns, overrides]);
  const allCats = useMemo(() => [...new Set([...ALL_CATS, ...txns.map(t => effCat(t))])].sort(), [txns, overrides]);
  const rows = useMemo(() => {
    let r = txns.filter(t =>
      (acc === "All" || t.a === acc) && (cat === "All" || effCat(t) === cat) &&
      (!q || (t.n + " " + effCat(t)).toLowerCase().includes(q.toLowerCase())));
    r = [...r].sort((a, b) => {
      let x = a[sort.k], y = b[sort.k];
      if (sort.k === "amt") { x = a.o || a.i; y = b.o || b.i; }
      return (x < y ? -1 : x > y ? 1 : 0) * sort.dir;
    });
    return r;
  }, [txns, q, acc, cat, sort, overrides]);
  const sums = useMemo(() => rows.reduce((a, t) => ({ o: a.o + t.o, i: a.i + t.i }), { o: 0, i: 0 }), [rows]);

  const downloadCSV = () => {
    const esc = v => `"${String(v).replace(/"/g, '""')}"`;
    const head = "Date,Account,Description,Category,Money Out (R),Money In (R)\n";
    const body = rows.map(t => [t.d, t.a, esc(t.n), esc(effCat(t)), t.o || "", t.i || ""].join(",")).join("\n");
    const blob = new Blob([head + body], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "transactions.csv";
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  };

  const Th = ({ k, children, w }) => (
    <th onClick={() => k && setSort(s => ({ k, dir: s.k === k ? -s.dir : -1 }))}
      style={{ textAlign:k==="amt"?"right":"left", padding:"8px 10px", color:C.sub, fontSize:11, fontWeight:700, textTransform:"uppercase", letterSpacing:0.5, cursor:k?"pointer":"default", width:w, userSelect:"none" }}>
      {children}{sort.k===k?(sort.dir<0?" ↓":" ↑"):""}
    </th>
  );

  const reclassCount = Object.keys(overrides).length;

  return (
    <>
      {classifyTxn && (
        <ClassifyModal txn={classifyTxn} currentCat={effCat(classifyTxn)} matchCount={matchCount(classifyTxn)} cats={allCats}
          onSave={(newCat,scope)=>onReclassify&&onReclassify(classifyTxn,newCat,scope)}
          onClose={()=>setClassifyTxn(null)} />
      )}
      <Glass pad={false}>
        <div className="flex items-center flex-wrap" style={{ gap:10, padding:14, borderBottom:`1px solid ${C.line}` }}>
          <div className="flex items-center" style={{ gap:8, background:"rgba(255,255,255,0.05)", borderRadius:10, padding:"8px 12px", flex:1, minWidth:180 }}>
            <Search size={15} color={C.faint} />
            <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search merchant or category…"
              style={{ background:"transparent", border:"none", outline:"none", color:C.text, fontSize:13, width:"100%" }} />
          </div>
          <Select value={acc} set={setAcc} opts={["All","CC","CHQ"]} labels={{CC:"Credit Card",CHQ:"Cheque"}} />
          <Select value={cat} set={setCat} opts={cats} />
          {(cat!=="All"&&presetCat)&&<button onClick={()=>{setCat("All");presetClear&&presetClear();}} style={btnGhost}><X size={13}/> Clear</button>}
          <button onClick={downloadCSV} style={btnGhost}><Download size={13}/> CSV</button>
        </div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>
              <Th k="d" w={92}>Date</Th><Th k="a" w={56}>Acct</Th><Th k="n">Description</Th>
              <Th k="c" w={155}>Category</Th><Th k="amt" w={110}>Amount</Th>
            </tr></thead>
            <tbody>
              {rows.slice(0,limit).map((t,i) => (
                <tr key={i} style={{ borderTop:`1px solid ${C.line}` }}>
                  <td style={td}>{t.d.slice(5)}</td>
                  <td style={td}><span style={{fontSize:10,color:t.a==="CC"?C.pink:C.blue,fontWeight:700}}>{t.a}</span></td>
                  <td style={{...td,color:C.text}}>{t.n}</td>
                  <td style={td}>
                    <span className="flex items-center" style={{gap:5}}>
                      <span style={{width:8,height:8,borderRadius:2,background:catColor(effCat(t)),flexShrink:0}}/>
                      <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis"}}>{effCat(t)}</span>
                      {(overrides["id:"+txnId(t)]||overrides["desc:"+t.n])&&<span title="Reclassified" style={{color:C.violet,fontSize:9,flexShrink:0}}>✎</span>}
                      <button onClick={()=>setClassifyTxn(t)} title="Reclassify"
                        style={{background:"transparent",border:"none",cursor:"pointer",color:C.faint,padding:"2px 3px",display:"flex",flexShrink:0,borderRadius:4}}>
                        <Tag size={11}/>
                      </button>
                    </span>
                  </td>
                  <td style={{...td,textAlign:"right",color:t.o?C.text:C.green,fontVariantNumeric:"tabular-nums",fontWeight:600}}>{t.o?R(t.o,2):"+"+R(t.i,2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="flex items-center justify-between flex-wrap" style={{padding:12,gap:8,borderTop:`1px solid ${C.line}`}}>
          <span style={{color:C.faint,fontSize:12}}>
            {rows.length} transactions{reclassCount>0?` · ${reclassCount} reclassified`:""}
          </span>
          <span style={{fontSize:12.5,fontWeight:700,fontVariantNumeric:"tabular-nums"}}>
            <span style={{color:C.faint,fontWeight:500}}>Total out </span><span style={{color:C.rose}}>{R(sums.o,2)}</span>
            <span style={{color:C.faint,fontWeight:500}}>  ·  in </span><span style={{color:C.green}}>{R(sums.i,2)}</span>
          </span>
          {limit<rows.length&&<button onClick={()=>setLimit(l=>l+60)} style={btnGhost}>Load more</button>}
        </div>
      </Glass>
    </>
  );
}
const td = { padding: "9px 10px", color: C.sub, fontSize: 12.5, whiteSpace: "nowrap" };
const btnGhost = { display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.05)", border: `1px solid ${C.line}`, color: C.text, fontSize: 12, padding: "8px 12px", borderRadius: 10, cursor: "pointer" };
function Select({ value, set, opts, labels = {} }) {
  return (
    <select value={value} onChange={e => set(e.target.value)}
      style={{ background: "rgba(255,255,255,0.05)", border: `1px solid ${C.line}`, color: C.text, fontSize: 12, padding: "8px 12px", borderRadius: 10, outline: "none", cursor: "pointer", maxWidth: 170 }}>
      {opts.map(o => <option key={o} value={o} style={{ background: C.panel2 }}>{labels[o] || o}</option>)}
    </select>
  );
}

function SettingsPage({ setTxns, status, setStatus, onExportAll, onClearTags, tagCount = 0 }) {
  const [pass, setPass] = useState("");
  const [needsPass, setNeedsPass] = useState(false);
  const [loading, setLoading] = useState(false);
  const queue = useRef([]);
  const accum = useRef([]);
  const pendingBytes = useRef(null);
  const doneCount = useRef(0);

  const loadPdfjs = () => new Promise((resolve, reject) => {
    if (window.pdfjsLib) { resolve(window.pdfjsLib); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    s.onerror = () => reject(new Error("Could not load PDF reader — check connection."));
    document.head.appendChild(s);
  });

  // ---- South-African amount parser: "1 040,65" | "2 000.00" | "142.38" -> number
  const zar = (str) => {
    let s = String(str).replace(/[^\d.,]/g, "");
    if (!s) return NaN;
    const dec = Math.max(s.lastIndexOf(","), s.lastIndexOf("."));
    if (dec >= 0 && s.length - dec - 1 === 2) {
      return parseFloat(s.slice(0, dec).replace(/[.,]/g, "") + "." + s.slice(dec + 1));
    }
    return parseFloat(s.replace(/[.,]/g, ""));
  };

  // ---- keyword categoriser (mirrors the ledger's own categories)
  const jsCat = (account, typ, desc) => {
    const d = (desc || "").toLowerCase();
    if (account === "CC") {
      if (typ === "CHG") return "Bank Charges";
      if (typ === "INT") return "Interest";
      if (["PY", "CV", "PUR"].includes(typ)) return "Card Payment";
      if (typ === "CA") return "Cash";
      if (d.includes("uber eats") || d.includes("mr d") || d.includes("mrd food")) return "Food Delivery";
      if (d.includes("uber") || d.includes("bolt")) return "Ride-hailing";
      if (["pick n pay", "pnp", "woolworth", "checker", "spar", "shoprite", "usave", "food lover", "ok mini", "superspar"].some(k => d.includes(k))) return "Groceries";
      if (["netflix", "apple.com", "google", "spotify", "showmax", "dstv", "youtube", "tradingview", "openai", "temu", "crunchyroll"].some(k => d.includes(k))) return "Subscriptions";
      if (["engen", "sasol", "bp ", "shell", "petroport", "caltex", "total", "toll", "sanral"].some(k => d.includes(k))) return "Fuel";
      if (["booking.com", "hotel", "airbnb", "flysafair", "lift airline", "airport", "car hire", "avis", "europcar"].some(k => d.includes(k))) return "Travel";
      if (["takealot", "pep ", "toys r us", "ctm", "truworths", "dischem", "clicks", "specsavers", "netflorist", "computicket", "hardware", "incredible connection"].some(k => d.includes(k))) return "Shopping";
      if (["kfc", "mcd", "mcdonald", "debonair", "nando", "steers", "wimpy", "spur", "rocomamas", "mugg and bean", "burger", "pizza", "chicken licken", "roman"].some(k => d.includes(k))) return "Dining";
      return "Other Spend";
    }
    if (d.includes("salary") || d.includes("payroll") || d.includes("wage")) return "Salary";
    if (d.includes("transfer") && d.includes("card")) return "Transfer to Card";
    if (/prepaid mobile|prepaid purchase|airtime|vodacom|\bvoda\b|\bmtn\b|cell c|telkom/.test(d)) return "Airtime";
    if (/monthly account|management fee|admin fee|service fee|transaction fee|notification fee|decline fee|debit order fee|insufficient funds fee|prepaid mobile purchase fee|external payment fee|withdrawal fee|deposit fee|\bfee\b/.test(d)) return "Bank Charges";
    if (d.includes("uber eats")) return "Food Delivery";
    if (/\buber\b|bolt/.test(d)) return "Ride-hailing";
    if (/mcdonald|\bkfc\b|bossies|debonair|nando|steers|burger|takeaway|spur|wimpy/.test(d)) return "Dining";
    if (/sasol|shell|engen|\bbp\b|caltex|petroport|\bfuel\b/.test(d)) return "Fuel";
    if (/sanral|\btoll\b|\btolls\b/.test(d)) return "Tolls";
    if (/netflix|tv licen|spotify|showmax|\bdstv\b|canva|subscription/.test(d)) return "Subscriptions";
    if (/home loan|sbsa hl|\bbond\b/.test(d)) return "Home Loan";
    if (/funeral|capfuneral|life insurance|legalwise|insurance|food for life/.test(d)) return "Insurance";
    if (/itransact|investment|ee-?915|tfsa|easy equities/.test(d)) return TFSA_CAT;
    if (d.includes("transfer")) return "Transfer";
    if (/cashsend|\batm\b|cash withdrawal/.test(d)) return "Cash";
    if (/woolworth|pick n pay|\bpnp\b|checkers|\bspar\b|shoprite|usave|food lover/.test(d)) return "Groceries";
    if (/debicheck|debit order/.test(d)) return "Debit Order";
    if (/rates|levy|municipal/.test(d)) return "Rates";
    if (/parking|servest/.test(d)) return "Parking";
    if (d.includes("mom")) return "Mom";
    if (/payment received|cash deposit|payshap.*received|interest received|\brefund\b|reversal/.test(d)) return "Income";
    if (/payshap|immediate payment|external payment/.test(d)) return "People";
    return "Other";
  };

  const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const AMT = "(\\d{1,3}(?:[ ]\\d{3})*[.,]\\d{2})"; // one SA-formatted money token

  // Group positioned text items into visual lines (per page, stacked top->bottom).
  const itemsToLines = (items) => {
    const its = items.map(i => ({ x: i.x, y: i.y, s: i.str })).filter(i => i.s !== "");
    its.sort((a, b) => b.y - a.y || a.x - b.x);
    const out = []; let cur = null;
    for (const it of its) {
      if (!cur || Math.abs(it.y - cur.y) > 3) { cur = { y: it.y, parts: [it] }; out.push(cur); }
      else cur.parts.push(it);
    }
    return out.map(l => l.parts.sort((a, b) => a.x - b.x).map(p => p.s).join(" ").replace(/\s+/g, " ").trim()).filter(Boolean);
  };

  // ---- Absa statement parser (credit-card & cheque layouts)
  const parseAbsaStatement = (lines) => {
    const full = lines.join("\n");
    const isCard = /credit card statement/i.test(full);
    const isCheque = /cheque account statement/i.test(full);
    if (!isCard && !isCheque) return null;

    const per = full.match(/(\d{1,2})\s+([A-Za-z]{3,})\w*\s+(20\d{2})\s+to\s+(\d{1,2})\s+([A-Za-z]{3,})\w*\s+(20\d{2})/);
    let startY = 2026, endY = 2026, startM = 1, endM = 12, endDay = 28;
    if (per) {
      startM = MON[per[2].slice(0, 3).toLowerCase()]; startY = +per[3];
      endM = MON[per[5].slice(0, 3).toLowerCase()]; endY = +per[6]; endDay = +per[4];
    }
    const inferYear = (mon) => startY === endY ? startY : (mon >= startM ? startY : endY);
    const iso = (day, mon) => `${inferYear(mon)}-${String(mon).padStart(2, "0")}-${String(day).padStart(2, "0")}`;

    const txns = [];
    if (isCard) {
      let section = "";
      const reTx = new RegExp("^(\\d{1,2})\\s+([A-Za-z]{3})\\s+(\\d{1,2})\\s+([A-Za-z]{3})\\s+([A-Z]{2,4})\\s+(.+?)\\s+" + AMT + "\\s*$");
      const rePay = new RegExp("^(\\d{1,2})\\s+([A-Za-z]{3})\\s+\\d{1,2}\\s+[A-Za-z]{3}\\s+([A-Z]{2,4})\\s+(.+?)\\s+" + AMT + "\\s*Cr?\\s*$", "i");
      const reChg = new RegExp("^(\\d{1,2})\\s+([A-Za-z]{3})\\s+(.+?)\\s+" + AMT + "\\s*$");
      for (const line of lines) {
        const lo = line.toLowerCase();
        if (/^transactions\b/.test(lo) && !/continued/.test(lo)) { section = "tx"; continue; }
        if (/payments\/credits/.test(lo)) { section = "pay"; continue; }
        if (/bank charges/.test(lo)) { section = "chg"; continue; }
        if (/interest calculated/.test(lo)) { section = "int"; continue; }
        if (/total balance on/.test(lo)) { section = ""; continue; }
        if (/^\(subtotal\)/i.test(lo) || /transactions for/i.test(lo) || /balance from/i.test(lo)) continue;

        if (section === "tx") {
          const m = line.match(reTx);
          if (m) {
            const mon = MON[m[2].toLowerCase()], amt = zar(m[7]);
            if (mon && amt > 0) txns.push({ d: iso(+m[1], mon), a: "CC", t: m[5], n: m[6].trim(), o: amt, i: 0, c: jsCat("CC", m[5], m[6]) });
          }
        } else if (section === "pay") {
          const m = line.match(rePay);
          if (m) {
            const mon = MON[m[2].toLowerCase()], amt = zar(m[5]);
            const desc = m[4].replace(/^\d{12,}\s*/, "").trim();
            if (mon && amt > 0) txns.push({ d: iso(+m[1], mon), a: "CC", t: m[3], n: desc, o: 0, i: amt, c: "Card Payment" });
          }
        } else if (section === "chg") {
          const m = line.match(reChg);
          if (m) {
            const mon = MON[m[2].toLowerCase()], amt = zar(m[4]);
            const desc = m[3].replace(new RegExp("\\s+" + AMT + "$"), "").trim() || "Bank charge";
            if (mon && amt > 0) txns.push({ d: iso(+m[1], mon), a: "CC", t: "CHG", n: desc, o: amt, i: 0, c: "Bank Charges" });
          }
        } else if (section === "int") {
          const m = line.match(new RegExp("interest amount.*?" + AMT + "\\s*$", "i"));
          if (m) { const amt = zar(m[1]); if (amt > 0) txns.push({ d: iso(endDay, endM), a: "CC", t: "INT", n: "Interest charged", o: amt, i: 0, c: "Interest" }); }
        }
      }
    } else {
      // cheque: dd/mm/yyyy desc <amount> <balance>; direction from balance delta
      let prevBal = null;
      const bf = full.match(new RegExp("balance brought forward[^0-9]*" + AMT, "i"));
      if (bf) prevBal = zar(bf[1]);
      const reRow = new RegExp("^(\\d{1,2})[\\/\\-](\\d{1,2})[\\/\\-](\\d{2,4})\\s+(.+?)\\s+" + AMT + "\\s+" + AMT + "\\s*$");
      for (const line of lines) {
        const m = line.match(reRow);
        if (!m) continue;
        let yr = +m[3]; if (yr < 100) yr += 2000;
        const amt = zar(m[5]), bal = zar(m[6]);
        const desc = m[4].replace(/\bsettlement\b/i, "").replace(/\s+/g, " ").trim();
        const dir = prevBal == null ? "out" : (bal >= prevBal ? "in" : "out");
        prevBal = bal;
        txns.push({ d: `${yr}-${String(+m[2]).padStart(2, "0")}-${String(+m[1]).padStart(2, "0")}`, a: "CHQ", t: "", n: desc, o: dir === "out" ? amt : 0, i: dir === "in" ? amt : 0, c: jsCat("CHQ", "", desc) });
      }
    }
    return txns.length ? txns : null;
  };

  // ---- FNB statement parser (rows like "20 Jun [desc] 83.00 3,609.67Cr")
  const parseFnbStatement = (lines) => {
    const full = lines.join("\n");
    if (!/first national bank|fnb\.co\.za|firstrand|gold business account/i.test(full)) return null;
    const pe = full.match(/to\s+\d{1,2}\s+([A-Za-z]+)\s+(20\d{2})/i);
    const ps = full.match(/(?:statement period[^0-9]*)?(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})\s+to/i);
    const endY = pe ? +pe[2] : (full.match(/(20\d{2})/) ? +full.match(/(20\d{2})/)[1] : 2023);
    const startY = ps ? +ps[3] : endY;
    const startM = ps ? MON[ps[2].slice(0, 3).toLowerCase()] : 1;
    const inferYear = (mon) => startY === endY ? startY : (mon >= startM ? startY : endY);
    const signed = (numStr, suf) => zar(numStr) * (/dr/i.test(suf || "") ? -1 : 1);
    const ob = full.match(/opening balance\s+([\d,]+\.\d{2})\s*(Cr|Dr)/i);
    let prev = ob ? signed(ob[1], ob[2]) : null;
    const txns = [];
    for (const line of lines) {
      const dm = line.match(/^(\d{1,2})\s+([A-Za-z]{3})\b/);
      if (!dm) continue;
      const mon = MON[dm[2].toLowerCase()]; if (!mon) continue;
      // money tokens (with optional Cr/Dr) on the line
      const toks = [...line.matchAll(/([\d,]+\.\d{2})\s*(Cr|Dr)?/g)];
      if (toks.length < 2) continue;
      // balance = last token carrying Cr/Dr; else the last token
      let balIdx = -1;
      for (let z = toks.length - 1; z >= 0; z--) { if (toks[z][2]) { balIdx = z; break; } }
      if (balIdx < 1) balIdx = toks.length - 1;
      const balTok = toks[balIdx], amtTok = toks[balIdx - 1];
      if (!amtTok) continue;
      const bal = signed(balTok[1], balTok[2]);
      const amt = zar(amtTok[1]);
      const dir = prev == null ? (bal < 0 ? "out" : "in") : (bal >= prev ? "in" : "out");
      prev = bal;
      if (!(amt > 0)) continue;
      const desc = line.slice(dm[0].length, amtTok.index).replace(/[#*]/g, "").replace(/\s+/g, " ").trim();
      const c = desc ? jsCat("CHQ", "", desc) : "Bank Charges";
      txns.push({ d: `${inferYear(mon)}-${String(mon).padStart(2, "0")}-${String(+dm[1]).padStart(2, "0")}`, a: "CHQ", t: "", n: desc || "FNB transaction", o: dir === "out" ? amt : 0, i: dir === "in" ? amt : 0, c });
    }
    return txns.length ? txns : null;
  };

  // ---- Standard Bank (MyMo) statement parser
  const parseStandardBank = (lines) => {
    const full = lines.join("\n");
    if (!/standardbank\.co\.za|standard bank of south africa|mymoacc/i.test(full)) return null;
    const toY = full.match(/to:\s*(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2})\b/i);
    const frY = full.match(/from:\s*(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2})\b/i);
    const endY = toY ? 2000 + +toY[3] : 2026;
    const startY = frY ? 2000 + +frY[3] : endY;
    const startM = frY ? MON[frY[2].toLowerCase()] : 1;
    const inferY = (mon) => startY === endY ? startY : (mon >= startM ? startY : endY);
    const txns = [];
    let inTx = false, curDate = null, buf = "";
    const dateRe = /^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2})\b\s*(.*)$/;
    const pairRe = /^(.*?)(-?\d[\d ]*[.,]\d{2})\s+(\d[\d ]*[.,]\d{2})\s*$/;
    const tryEmit = () => {
      const m = buf.match(pairRe);
      if (!m || !curDate) return;
      const dtext = m[1].replace(/\s+/g, " ").trim();
      const neg = m[2].trim().startsWith("-");
      const amt = zar(m[2]);
      buf = "";
      if (!(amt > 0)) return;
      txns.push({ d: curDate, a: "CHQ", t: "", n: dtext || "Standard Bank", o: neg ? amt : 0, i: neg ? 0 : amt, c: jsCat("CHQ", "", dtext) });
    };
    for (const raw of lines) {
      const line = raw.trim();
      if (/date\s+description\s+payments|transaction details/i.test(line)) { inTx = true; continue; }
      if (/statement summary|please verify/i.test(line)) { inTx = false; continue; }
      if (!inTx) continue;
      const dm = line.match(dateRe);
      if (dm) {
        const mon = MON[dm[2].toLowerCase()];
        if (mon) curDate = `${inferY(mon)}-${String(mon).padStart(2, "0")}-${String(+dm[1]).padStart(2, "0")}`;
        buf = dm[4] || "";
        tryEmit();
      } else { buf += " " + line; tryEmit(); }
    }
    return txns.length ? txns : null;
  };

  // ---- Capitec statement parser (explicit category column + Money In/Out/Fee)
  const CAP_CATS = ["Home Loan Payments", "Digital Subscriptions", "Credit Card Payments", "Other Income", "Cash Deposit", "Funeral Cover", "Life Insurance", "Cash Withdrawal", "Digital Payments", "Uncategorised", "Investments", "Cellphone", "Takeaways", "Transfer", "Parking", "Interest", "Tolls", "Fuel", "Fees"];
  const CAP_MAP = { "Fees": "Bank Charges", "Cellphone": "Airtime", "Takeaways": "Dining", "Digital Subscriptions": "Subscriptions", "Cash Withdrawal": "Cash", "Other Income": "Income", "Cash Deposit": "Income", "Interest": "Income", "Home Loan Payments": "Home Loan", "Life Insurance": "Insurance", "Funeral Cover": "Insurance", "Credit Card Payments": "Credit Card Payment", "Uncategorised": "Other", "Investments": TFSA_CAT };
  const parseCapitec = (lines) => {
    const full = lines.join("\n");
    if (!/capitecbank\.co\.za|capitec bank limited|fsp46669/i.test(full)) return null;
    const txns = [];
    let curDate = null, buf = "";
    const dateRe = /^(\d{2})\/(\d{2})\/(\d{4})\b\s*(.*)$/;
    const moneyG = /-?\d[\d ]*[.,]\d{2}/g;
    const tryEmit = () => {
      if (!curDate) return;
      const m = buf.match(/^(.*?)((?:\s*-?\d[\d ]*[.,]\d{2}){2,})\s*$/);
      if (!m) return;
      const head = m[1].replace(/\s+/g, " ").trim();
      let cat = null, desc = head;
      for (const k of CAP_CATS) { if (head.toLowerCase().endsWith(k.toLowerCase())) { cat = k; desc = head.slice(0, head.length - k.length).trim(); break; } }
      if (!cat) return;
      const nums = (m[2].match(moneyG) || []).map(s => ({ neg: s.trim().startsWith("-"), v: zar(s) }));
      const pre = nums.slice(0, -1); // drop balance
      let moneyIn = 0; const outs = [];
      for (const p of pre) { if (p.neg) outs.push(p.v); else moneyIn += p.v; }
      let fee = 0, mainOut = 0;
      if (outs.length) {
        if (moneyIn > 0) fee = outs.reduce((a, b) => a + b, 0);
        else if (outs.length >= 2) { fee = outs[outs.length - 1]; mainOut = outs.slice(0, -1).reduce((a, b) => a + b, 0); }
        else mainOut = outs[0];
      }
      const mapped = CAP_MAP[cat] || cat;
      if (moneyIn > 0) txns.push({ d: curDate, a: "CHQ", t: "", n: desc || cat, o: 0, i: moneyIn, c: mapped });
      if (mainOut > 0) txns.push({ d: curDate, a: "CHQ", t: "", n: desc || cat, o: mainOut, i: 0, c: cat === "Fees" ? "Bank Charges" : mapped });
      if (fee > 0) txns.push({ d: curDate, a: "CHQ", t: "", n: "Bank fee", o: fee, i: 0, c: "Bank Charges" });
      buf = "";
    };
    for (const raw of lines) {
      const line = raw.trim();
      if (/includes vat|page \d+ of|client care|capitec bank is an|unique document/i.test(line)) continue;
      const dm = line.match(dateRe);
      if (dm) { curDate = `${dm[3]}-${dm[2]}-${dm[1]}`; buf = dm[4] || ""; tryEmit(); }
      else { buf += " " + line; tryEmit(); }
    }
    return txns.length ? txns : null;
  };

  // ---- fallback: the app's own ISO-dated PDF export
  const parseIsoLedger = (lines) => {
    const out = [];
    for (const line of lines) {
      const dm = line.match(/\b(\d{4}-\d{2}-\d{2})\b/); if (!dm) continue;
      const d = dm[1];
      const rest = line.slice(line.indexOf(d) + d.length).trim();
      const ams = [...rest.matchAll(/\b(\d{1,3}(?:[,\s]\d{3})*(?:[.,]\d{2})?)\b/g)]
        .map(m => zar(m[1])).filter(n => n > 0 && n < 2000000);
      if (!ams.length) continue;
      const a = /\bCC\b|credit/i.test(rest) ? "CC" : "CHQ";
      const c = Object.keys(CAT_COLOR).find(k => rest.toLowerCase().includes(k.toLowerCase())) || "Other";
      const desc = rest.replace(/\b\d[\d,. ]*\d\b/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
      out.push({ d, a, t: "IMP", n: desc, o: ams[0] || 0, i: ams[1] || 0, c });
    }
    return out.length ? out : null;
  };

  const finish = () => {
    setLoading(false); setNeedsPass(false); setPass(""); pendingBytes.current = null;
    const merged = accum.current;
    if (!merged.length) { setStatus({ ok: false, msg: "No transactions could be read from the selected file(s)." }); return; }
    const seen = new Set(); const uniq = [];
    for (const t of merged) { const k = txnId(t); if (!seen.has(k)) { seen.add(k); uniq.push(t); } }
    uniq.sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : 0);
    setTxns(uniq);
    setStatus({ ok: true, msg: `Imported ${uniq.length} transactions from ${doneCount.current} file(s).` });
  };

  const processNext = async () => {
    if (!queue.current.length) { finish(); return; }
    const f = queue.current[0];
    if (f.name.toLowerCase().endsWith(".pdf")) {
      const ab = await f.arrayBuffer();
      pendingBytes.current = new Uint8Array(ab);
      await tryPdf(null);
    } else if (f.name.toLowerCase().endsWith(".csv")) {
      try {
        const text = await f.text();
        const mapped = csvToTxns(text);
        if (!mapped.length) throw new Error("no rows");
        accum.current = accum.current.concat(mapped);
        doneCount.current++;
      } catch (err) {
        setStatus({ ok: false, msg: `Skipped "${f.name}" — couldn't read the CSV.` });
      }
      queue.current.shift();
      processNext();
    } else {
      try {
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
        const sh = wb.Sheets["All Transactions"] || wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sh, { defval: "" });
        const mapped = json.map(r => {
          const acctRaw = String(r["Account"] || "");
          const d = r["Date"] instanceof Date ? r["Date"].toISOString().slice(0, 10) : String(r["Date"]).slice(0, 10);
          const ee = isEEAccount(acctRaw);
          const a = ee ? "EE" : (acctRaw.toLowerCase().includes("credit") || acctRaw === "CC" ? "CC" : "CHQ");
          const c = ee ? TFSA_CAT : (r["Category"] || "Other");
          return { d, a, t: r["Type"] || "", n: String(r["Description"] || ""), o: +r["Money Out (R)"] || 0, i: +r["Money In (R)"] || 0, c };
        }).filter(r => r.d && r.d !== "undefined");
        accum.current = accum.current.concat(mapped);
        doneCount.current++;
      } catch (err) {
        setStatus({ ok: false, msg: `Skipped "${f.name}" — couldn't read it.` });
      }
      queue.current.shift();
      processNext();
    }
  };

  const tryPdf = async (password) => {
    setLoading(true);
    try {
      const lib = await loadPdfjs();
      const data = pendingBytes.current.slice();
      const pdf = await lib.getDocument({ data, password: password || undefined }).promise;
      const N = pdf.numPages;
      let items = [];
      for (let p = 1; p <= N; p++) {
        const page = await pdf.getPage(p);
        const ct = await page.getTextContent();
        // offset y per page so pages stack in reading order without colliding
        const off = (N - p) * 100000;
        for (const it of ct.items) {
          if (!it.str) continue;
          items.push({ x: it.transform[4], y: it.transform[5] + off, str: it.str });
        }
      }
      const lines = itemsToLines(items);
      if (!lines.length) {
        setStatus({ ok: false, msg: `"${queue.current[0]?.name}" looks scanned (image-only) — export it as XLSX instead.` });
      } else {
        const parsed = parseAbsaStatement(lines) || parseFnbStatement(lines) || parseStandardBank(lines) || parseCapitec(lines) || parseIsoLedger(lines);
        if (!parsed || !parsed.length) {
          setStatus({ ok: false, msg: `Read "${queue.current[0]?.name}" but couldn't recognise transaction rows. If it's an Absa statement export it as XLSX, or check the layout.` });
        } else {
          accum.current = accum.current.concat(parsed);
        }
      }
      doneCount.current++;
      setNeedsPass(false); setPass(""); pendingBytes.current = null;
      queue.current.shift();
      processNext();
    } catch (err) {
      const isPass = err && (err.name === "PasswordException" || /password/i.test(err.message || ""));
      if (isPass) {
        setNeedsPass(true); setLoading(false);
        setStatus({ ok: false, msg: `"${queue.current[0]?.name}" is password protected — enter its password below.` });
      } else {
        setStatus({ ok: false, msg: "PDF error: " + (err.message || String(err)) });
        setNeedsPass(false); pendingBytes.current = null;
        queue.current.shift();
        processNext();
      }
    }
  };

  const onFiles = (e) => {
    const files = [...(e.target.files || [])];
    e.target.value = "";
    if (!files.length) return;
    queue.current = files; accum.current = []; doneCount.current = 0;
    setStatus({ ok: true, msg: `Reading ${files.length} file(s)…` });
    processNext();
  };

  return (
    <div className="flex flex-col" style={{ gap: 16, maxWidth: 640 }}>
      <Glass>
        <CardTitle>Import Updated File(s)</CardTitle>
        <p style={{ color: C.sub, fontSize: 13, lineHeight: 1.55, marginBottom: 14 }}>
          Import one or more <b style={{ color: C.text }}>XLSX</b> workbooks (each needs an <em>All Transactions</em> sheet)
          and/or text-based <b style={{ color: C.text }}>PDF</b> statements. Select several at once — they're merged into one collective dataset.
          Scanned/image PDFs can't be parsed; password-protected PDFs prompt for a password.
        </p>
        <label style={{ ...btnGhost, display: "inline-flex", padding: "10px 16px", cursor: loading ? "default" : "pointer", opacity: loading ? 0.6 : 1 }}>
          <Upload size={15} /> {loading ? "Reading…" : "Choose files (.xlsx / .pdf / .csv)"}
          <input type="file" accept=".xlsx,.xls,.pdf,.csv" multiple style={{ display: "none" }} onChange={onFiles} disabled={loading} />
        </label>
        {needsPass && (
          <div style={{ marginTop: 16, padding: 16, background: "rgba(255,255,255,0.04)", borderRadius: 12, border: `1px solid ${C.line}` }}>
            <div style={{ color: C.amber, fontSize: 13, fontWeight: 600, marginBottom: 10 }}>🔒 Password Required</div>
            <div className="flex" style={{ gap: 8 }}>
              <input type="password" value={pass} autoFocus
                onChange={e => setPass(e.target.value)}
                onKeyDown={e => e.key === "Enter" && tryPdf(pass)}
                placeholder="Enter PDF password…"
                style={{ flex: 1, background: C.panel, border: `1px solid ${C.line}`, color: C.text, fontSize: 13, padding: "9px 12px", borderRadius: 10, outline: "none" }} />
              <button onClick={() => tryPdf(pass)}
                style={{ ...btnGhost, background: C.violet, color: "#fff", border: "none", fontWeight: 700, padding: "9px 16px" }}>
                Unlock
              </button>
            </div>
          </div>
        )}
        {status && <div style={{ marginTop: 12, color: status.ok ? C.green : C.rose, fontSize: 13, lineHeight: 1.5 }}>{status.msg}</div>}
      </Glass>
      <Glass>
        <CardTitle>Export</CardTitle>
        <div className="flex" style={{ gap: 10, flexWrap: "wrap" }}>
          <button onClick={() => onExportAll && onExportAll()} style={btnGhost}><Download size={14} /> Download CSV</button>
        </div>
        <p style={{ color: C.faint, fontSize: 12, marginTop: 10 }}>
          Exports all transactions to a CSV file you can re-import here, open in Excel/Sheets, or keep as a backup.
        </p>
        <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${C.line}` }}>
          <div className="flex items-center justify-between" style={{ gap: 10, flexWrap: "wrap" }}>
            <div style={{ color: C.sub, fontSize: 12.5 }}>{tagCount} saved category tag{tagCount === 1 ? "" : "s"} — reapplied automatically when you import new statements.</div>
            {tagCount > 0 && <button onClick={() => onClearTags && onClearTags()} style={btnGhost}><X size={13} /> Clear saved tags</button>}
          </div>
        </div>
      </Glass>
      <Glass>
        <CardTitle>About</CardTitle>
        <p style={{ color: C.sub, fontSize: 13, lineHeight: 1.6 }}>
          Both accounts are treated as one collective wallet — transfers between the cheque account and credit card net to zero,
          deposits are positive, withdrawals negative. Tap the <b style={{ color: C.text }}>tag icon</b> on any transaction to reclassify it
          (optionally cascading to every matching entry). Re-import an updated XLSX to make changes permanent.
        </p>
      </Glass>
    </div>
  );
}

/* ============================ APP SHELL ============================ */
function EmptyState({ onImport }) {
  return (
    <div className="flex flex-col items-center justify-center" style={{ minHeight: "60vh", textAlign: "center", padding: 24 }}>
      <div style={{ width: 56, height: 56, borderRadius: 16, background: GRAD, marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <Upload size={26} color="#fff" />
      </div>
      <h2 style={{ fontSize: 20, fontWeight: 800, marginBottom: 8 }}>No transactions yet</h2>
      <p style={{ color: C.sub, fontSize: 14, maxWidth: 420, lineHeight: 1.5, marginBottom: 22 }}>
        Import your bank statements (Absa, FNB, Standard Bank, Capitec PDFs or an XLSX) to populate the dashboard. Everything stays on this device for the session.
      </p>
      <button onClick={onImport} style={{ ...btnGhost, background: GRAD, border: "none", color: "#fff", padding: "11px 18px", fontSize: 14 }}>
        <Upload size={15} /> Import statements
      </button>
    </div>
  );
}

/* ============================ TFSA ============================ */
const TFSA_ANNUAL = 46000, TFSA_LIFETIME = 500000, STXNDQ_RATE = 25.19;
const moneyAxis = (v) => { const a = Math.abs(v); if (a >= 1e6) return "R" + (v / 1e6).toFixed(1) + "M"; if (a >= 1e3) return "R" + Math.round(v / 1e3) + "k"; return "R" + Math.round(v); };

// Matches calculator.net's interest calculator (compound annually, contribute at the beginning
// of each period by default). Each year is 12 monthly steps at the geometric monthly rate
// f = (1+r)^(1/12)-1, so the principal compounds to exactly A = P(1+r)^t over t = maturity−now
// years, while monthly contributions accrue interest through the year (as their schedule does).
function projectTFSA({ balance0, lifetime0, ytd0, monthly, ratePct, startYear, maturityYear, timing = "begin" }) {
  const A = TFSA_ANNUAL, CAP = TFSA_LIFETIME;
  const r = ratePct / 100;
  const f = Math.pow(1 + r, 1 / 12) - 1;
  let balance = balance0, lifetime = lifetime0;
  const points = [{ year: startYear, balance, contributed: lifetime }];
  let capYear = lifetime >= CAP - 1e-6 ? startYear : null;
  const t = Math.max(0, maturityYear - startYear);
  for (let yr = 1; yr <= t; yr++) {
    let roomLeft = (yr === 1) ? Math.max(0, A - ytd0) : A;   // R46k allowance per tax year
    for (let mth = 0; mth < 12; mth++) {
      const room = Math.max(0, Math.min(monthly, roomLeft, CAP - lifetime));
      balance = (timing === "begin") ? (balance + room) * (1 + f) : balance * (1 + f) + room;
      lifetime += room; roomLeft -= room;
      if (capYear === null && lifetime >= CAP - 1e-6) capYear = startYear + yr;
    }
    points.push({ year: startYear + yr, balance, contributed: lifetime });
  }
  return { points, finalBalance: balance, totalContributed: lifetime, capYear };
}

function Progress({ label, value, max, accent, note }) {
  const pct = Math.max(0, Math.min(100, max ? (value / max) * 100 : 0));
  return (
    <div>
      <div className="flex items-center justify-between" style={{ marginBottom: 8 }}>
        <span style={{ color: C.text, fontWeight: 600, fontSize: 14 }}>{label}</span>
        <span style={{ color: C.sub, fontSize: 13, fontVariantNumeric: "tabular-nums" }}>{R(value)} / {R(max)}</span>
      </div>
      <div style={{ height: 14, borderRadius: 99, background: "rgba(255,255,255,0.06)", overflow: "hidden", border: `1px solid ${C.line}` }}>
        <div style={{ width: `${pct}%`, height: "100%", borderRadius: 99, background: `linear-gradient(90deg, ${accent}, ${accent}bb)`, boxShadow: `0 0 16px ${accent}55`, transition: "width .7s ease" }} />
      </div>
      <div className="flex items-center justify-between" style={{ marginTop: 6 }}>
        <span style={{ color: accent, fontSize: 12, fontWeight: 700 }}>{pct.toFixed(1)}% used</span>
        <span style={{ color: C.faint, fontSize: 12 }}>{note}</span>
      </div>
    </div>
  );
}

function NumField({ label, value, onChange, prefix = "R", suffix, hint }) {
  return (
    <label style={{ display: "block" }}>
      <div style={{ color: C.sub, fontSize: 12, fontWeight: 600, marginBottom: 6 }}>{label}</div>
      <div className="flex items-center" style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, padding: "0 10px" }}>
        {prefix && <span style={{ color: C.faint, fontSize: 13 }}>{prefix}</span>}
        <input value={value} inputMode="decimal" onChange={e => onChange(e.target.value)}
          style={{ flex: 1, background: "transparent", border: "none", color: C.text, fontSize: 14, padding: "9px 6px", outline: "none", width: "100%" }} />
        {suffix && <span style={{ color: C.faint, fontSize: 13 }}>{suffix}</span>}
      </div>
      {hint && <div style={{ color: C.faint, fontSize: 11, marginTop: 5 }}>{hint}</div>}
    </label>
  );
}

const MiniStat = ({ label, value, accent = C.violet }) => (
  <Glass>
    <div style={{ color: C.sub, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>{label}</div>
    <div style={{ color: C.text, fontWeight: 700, fontSize: 22, marginTop: 6, fontVariantNumeric: "tabular-nums" }}>{value}</div>
    <div style={{ height: 3, width: 34, borderRadius: 99, background: accent, marginTop: 9 }} />
  </Glass>
);

function TFSAPage({ txns = [], overrides = {} }) {
  const narrow = useIsNarrow();
  const now = new Date();
  const curYear = now.getFullYear();
  const taxStartYear = now.getMonth() + 1 >= 3 ? curYear : curYear - 1;
  const taxStartDate = `${taxStartYear}-03-01`;
  // every transaction in the EE/TFSA category is a contribution movement (inflow to the TFSA or transfer into it)
  const tfsaTxns = txns.filter(t => effCategory(t, overrides) === TFSA_CAT);
  const flow = t => (t.o || 0) + (t.i || 0);
  const byMonth = {};
  tfsaTxns.forEach(t => { const m = (t.d || "").slice(0, 7); byMonth[m] = (byMonth[m] || 0) + flow(t); });
  const detectedLifetime = Math.round(tfsaTxns.reduce((a, t) => a + flow(t), 0));
  const detectedYtd = Math.round(tfsaTxns.filter(t => (t.d || "") >= taxStartDate).reduce((a, t) => a + flow(t), 0));
  const dataMonths = Object.keys(byMonth).length;
  const detectedMonthly = Math.round(dataMonths ? detectedLifetime / dataMonths : 0);

  const num = (v, d = 0) => { const n = parseFloat(String(v).replace(/[^\d.]/g, "")); return isFinite(n) ? n : d; };
  const [lifetime, setLifetime] = useState(String(detectedLifetime));
  const [ytd, setYtd] = useState(String(detectedYtd));
  const [balance, setBalance] = useState(String(detectedLifetime));
  const [mode, setMode] = useState("current");
  const [customM, setCustomM] = useState(String(detectedMonthly || 3000));
  const [rate, setRate] = useState(String(STXNDQ_RATE));
  const [matYear, setMatYear] = useState(String(curYear + 20));
  const [timing, setTiming] = useState("begin");

  const lifetimeN = Math.min(num(lifetime), TFSA_LIFETIME);
  const ytdN = num(ytd), balanceN = num(balance);
  const monthly = mode === "current" ? detectedMonthly : num(customM);
  const rateN = num(rate, STXNDQ_RATE);
  const matN = Math.max(curYear + 1, Math.round(num(matYear, curYear + 20)));

  const proj = useMemo(() => projectTFSA({
    balance0: balanceN || lifetimeN, lifetime0: lifetimeN, ytd0: ytdN,
    monthly, ratePct: rateN, startYear: curYear, maturityYear: matN, timing
  }), [balanceN, lifetimeN, ytdN, monthly, rateN, matN, curYear, timing]);
  const growth = proj.finalBalance - proj.totalContributed;
  const monthlyCapped = monthly > TFSA_ANNUAL / 12;

  const seg = (on) => ({ flex: 1, textAlign: "center", padding: "9px 12px", borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: "pointer", border: `1px solid ${on ? C.violet : C.line}`, background: on ? "rgba(139,124,246,0.16)" : "transparent", color: on ? C.text : C.sub });
  const resetDetected = () => { setLifetime(String(detectedLifetime)); setYtd(String(detectedYtd)); setBalance(String(detectedLifetime)); setCustomM(String(detectedMonthly || 3000)); };

  return (
    <div className="flex flex-col" style={{ gap: 16 }}>
      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>STXNDQ · Nasdaq-100</span>}>Tax-Free Savings Account</CardTitle>
        <p style={{ color: C.sub, fontSize: 13, lineHeight: 1.55 }}>
          A South African TFSA grows free of tax on interest, dividends and capital gains. You may contribute up to <b style={{ color: C.text }}>R46 000</b> per tax year (1 March – end February) and <b style={{ color: C.text }}>R500 000</b> over your lifetime. Once the lifetime cap is reached, no further contributions are allowed — but the balance can keep compounding well beyond it.
        </p>
      </Glass>

      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr" : "1fr 1fr", gap: 16 }}>
        <Glass>
          <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>resets 1 March</span>}>Annual limit · {taxStartYear}/{(taxStartYear + 1) % 100}</CardTitle>
          <Progress label="This tax year" value={ytdN} max={TFSA_ANNUAL} accent={C.violet} note={`${R(Math.max(0, TFSA_ANNUAL - ytdN))} room left`} />
        </Glass>
        <Glass>
          <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>contributions only</span>}>Lifetime limit</CardTitle>
          <Progress label="All-time contributed" value={lifetimeN} max={TFSA_LIFETIME} accent={C.amber} note={`${R(Math.max(0, TFSA_LIFETIME - lifetimeN))} room left`} />
        </Glass>
      </div>

      <Glass>
        <CardTitle right={<button onClick={resetDetected} style={{ background: "transparent", border: "none", color: C.violet, fontSize: 12, cursor: "pointer", fontWeight: 600 }}>↻ Use detected</button>}>Your figures</CardTitle>
        <p style={{ color: C.faint, fontSize: 11.5, marginBottom: 14 }}>
          Detected from your <b style={{ color: C.text }}>TFSA</b> transactions — any EE-915 account movement, plus iTransact/EasyEquities/TFSA entries{dataMonths ? ` (${R(detectedLifetime)} over ${dataMonths} month${dataMonths === 1 ? "" : "s"})` : " — none yet"}. Adjust any value, including prior years not in your statements.
        </p>
        <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr 1fr" : "1fr 1fr 1fr", gap: 12 }}>
          <NumField label="Contributed this tax year" value={ytd} onChange={setYtd} />
          <NumField label="Lifetime contributed" value={lifetime} onChange={setLifetime} />
          <NumField label="Current TFSA value" value={balance} onChange={setBalance} hint="incl. growth so far" />
        </div>
      </Glass>

      <Glass>
        <CardTitle>Growth forecast</CardTitle>
        <div className="flex flex-col" style={{ gap: 14 }}>
          <div>
            <div style={{ color: C.sub, fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Contribution rate</div>
            <div className="flex" style={{ gap: 8 }}>
              <div onClick={() => setMode("current")} style={seg(mode === "current")}>Current ({R(detectedMonthly)}/mo)</div>
              <div onClick={() => setMode("custom")} style={seg(mode === "custom")}>Custom amount</div>
            </div>
            {mode === "custom"
              ? <div style={{ marginTop: 12 }}><NumField label="Monthly contribution" value={customM} onChange={setCustomM} hint="A discretionary amount you'd contribute each month" /></div>
              : <div style={{ color: C.faint, fontSize: 11.5, marginTop: 8 }}>Based on what you currently transfer to your TFSA{detectedMonthly ? "" : " — import statements or switch to a custom amount"}.</div>}
            {monthlyCapped && <div style={{ color: C.amber, fontSize: 11.5, marginTop: 8 }}>Note: {R(monthly)}/mo exceeds the R3 833/mo that fills the R46 000 annual allowance — contributions pause each year once the annual limit is hit.</div>}
          </div>
          <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr 1fr" : "1fr 1fr", gap: 12 }}>
            <NumField label="Growth rate (annual)" value={rate} onChange={setRate} prefix="" suffix="%" hint="STXNDQ since-inception ≈ 25.2% (high — edit to taste)" />
            <NumField label="Year of maturity" value={matYear} onChange={setMatYear} prefix="" hint={`default ${curYear + 20}`} />
          </div>
          <div>
            <div style={{ color: C.sub, fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Contributions applied at</div>
            <div className="flex" style={{ gap: 8 }}>
              <div onClick={() => setTiming("begin")} style={seg(timing === "begin")}>Start of period</div>
              <div onClick={() => setTiming("end")} style={seg(timing === "end")}>End of period</div>
            </div>
            <div style={{ color: C.faint, fontSize: 11.5, marginTop: 8 }}>Compounds annually (matches calculator.net). "Start" = each contribution earns a full period's interest — calculator.net's default.</div>
          </div>
        </div>
      </Glass>

      <div className="grid" style={{ gridTemplateColumns: narrow ? "1fr 1fr" : "repeat(4, 1fr)", gap: 12 }}>
        <MiniStat label={`Value in ${matN}`} value={R(proj.finalBalance)} accent={C.violet} />
        <MiniStat label="Total contributed" value={R(proj.totalContributed)} accent={C.amber} />
        <MiniStat label="Tax-free growth" value={R(growth)} accent={C.green} />
        <MiniStat label="R500k cap reached" value={proj.capYear || "—"} accent={C.rose} />
      </div>

      <Glass>
        <CardTitle right={<span style={{ color: C.faint, fontSize: 12 }}>balance vs contributions</span>}>Projection to {matN}</CardTitle>
        <H h={300}>
          <ComposedChart data={proj.points} margin={{ top: 12, right: 14, bottom: 4, left: 4 }}>
            <defs>
              <linearGradient id="tfsaBal" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={C.violet} stopOpacity={0.5} />
                <stop offset="100%" stopColor={C.violet} stopOpacity={0.02} />
              </linearGradient>
            </defs>
            <CartesianGrid stroke="rgba(255,255,255,0.05)" vertical={false} />
            <XAxis dataKey="year" tick={{ fill: C.faint, fontSize: 11 }} tickLine={false} axisLine={false} />
            <YAxis tickFormatter={moneyAxis} tick={{ fill: C.faint, fontSize: 11 }} tickLine={false} axisLine={false} width={50} />
            <Tooltip formatter={(v, n) => [R(v), n === "balance" ? "Projected value" : "Contributed"]} labelFormatter={(y) => `Year ${y}`} contentStyle={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, color: C.text }} />
            <ReferenceLine y={TFSA_LIFETIME} stroke={C.amber} strokeDasharray="4 4" label={{ value: "R500k cap", fill: C.amber, fontSize: 11, position: "insideTopRight" }} />
            {proj.capYear ? <ReferenceLine x={proj.capYear} stroke={C.rose} strokeDasharray="3 3" label={{ value: "contributions stop", fill: C.rose, fontSize: 10, position: "insideTopLeft" }} /> : null}
            <Area type="monotone" dataKey="balance" stroke={C.violet} strokeWidth={2.4} fill="url(#tfsaBal)" />
            <Line type="monotone" dataKey="contributed" stroke={C.green} strokeWidth={2} strokeDasharray="5 4" dot={false} />
          </ComposedChart>
        </H>
        <p style={{ color: C.faint, fontSize: 11, marginTop: 10, lineHeight: 1.5 }}>
          Illustrative only — not financial advice. Past STXNDQ performance (driven by a tech bull run and rand weakness) is not indicative of future returns; the contribution line plateaus at R500 000 while the balance keeps compounding to {matN}.
        </p>
      </Glass>
    </div>
  );
}

const NAV = [
  { k: "dash", label: "Dashboard", icon: LayoutDashboard },
  { k: "tx", label: "Transactions", icon: ListOrdered },
  { k: "cat", label: "Categories", icon: PieIcon },
  { k: "trends", label: "Monthly Trends", icon: TrendingUp },
  { k: "flow", label: "Cash Flow", icon: Waves },
  { k: "insights", label: "Insights", icon: Sparkles },
  { k: "tfsa", label: "TFSA", icon: PiggyBank },
  { k: "settings", label: "Settings", icon: SettingsIcon },
];

export default function App() {
  const [txns, setTxns] = useState([]);
  const [page, setPage] = useState("dash");
  const [presetCat, setPresetCat] = useState(null);
  const [status, setStatus] = useState(null);
  const [search, setSearch] = useState("");
  const [overrides, setOverrides] = useState({});
  const [hydrated, setHydrated] = useState(false);
  // load saved categorisations once on mount
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        if (typeof window !== "undefined" && window.storage) {
          const r = await window.storage.get("tagOverrides");
          if (alive && r && r.value) { const parsed = JSON.parse(r.value); if (parsed && typeof parsed === "object") setOverrides(parsed); }
        } else if (typeof window !== "undefined" && window.localStorage) {
          const v = window.localStorage.getItem("tagOverrides");
          if (alive && v) { const parsed = JSON.parse(v); if (parsed && typeof parsed === "object") setOverrides(parsed); }
        }
      } catch (e) { /* no saved tags / storage unavailable */ }
      if (alive) setHydrated(true);
    })();
    return () => { alive = false; };
  }, []);
  // persist whenever they change (after initial hydration so we don't clobber saved data with {})
  useEffect(() => {
    if (!hydrated) return;
    (async () => {
      try {
        if (typeof window !== "undefined" && window.storage) await window.storage.set("tagOverrides", JSON.stringify(overrides));
        else if (typeof window !== "undefined" && window.localStorage) window.localStorage.setItem("tagOverrides", JSON.stringify(overrides));
      } catch (e) {}
    })();
  }, [overrides, hydrated]);
  const onReclassify = (t, cat, scope) =>
    setOverrides(o => ({ ...o, [(scope === "desc" ? "desc:" + t.n : "id:" + txnId(t))]: cat }));
  const clearTags = () => {
    setOverrides({});
    (async () => { try { if (typeof window !== "undefined" && window.storage) await window.storage.delete("tagOverrides"); else if (typeof window !== "undefined" && window.localStorage) window.localStorage.removeItem("tagOverrides"); } catch (e) {} })();
  };
  const model = useModel(txns, overrides);
  const narrow = useIsNarrow();
  const onExportAll = () => {
    const esc = v => `"${String(v).replace(/"/g, '""')}"`;
    const head = "Date,Account,Description,Category,Money Out (R),Money In (R)\n";
    const body = txns.map(t => [t.d, t.a, esc(t.n), esc(effCategory(t, overrides)), t.o || "", t.i || ""].join(",")).join("\n");
    const blob = new Blob([head + body], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "all-transactions.csv";
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  };

  const goCat = (k) => { setPresetCat(k); setPage("tx"); };
  const onImport = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true });
        const sh = wb.Sheets["All Transactions"] || wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sh, { defval: "" });
        const mapped = json.map(r => {
          const acctRaw = String(r["Account"] || "");
          const d = r["Date"] instanceof Date ? r["Date"].toISOString().slice(0, 10) : String(r["Date"]).slice(0, 10);
          const ee = isEEAccount(acctRaw);
          const a = ee ? "EE" : (acctRaw.toLowerCase().includes("credit") || acctRaw === "CC" ? "CC" : "CHQ");
          const c = ee ? TFSA_CAT : (r["Category"] || "Other");
          return { d, a, t: r["Type"] || "", n: String(r["Description"] || ""), o: +r["Money Out (R)"] || 0, i: +r["Money In (R)"] || 0, c };
        }).filter(r => r.d && r.d !== "undefined");
        if (!mapped.length) throw new Error("No rows found");
        setTxns(mapped); setStatus({ ok: true, msg: `Imported ${mapped.length} transactions. Dashboard updated.` });
        setPage("dash");
      } catch (err) { setStatus({ ok: false, msg: "Couldn't read that file. Make sure it has an 'All Transactions' sheet." }); }
    };
    rd.readAsArrayBuffer(f);
  };

  const globalResults = useMemo(() => {
    if (!search.trim()) return [];
    const s = search.toLowerCase();
    return txns.filter(t => (t.n + " " + t.c).toLowerCase().includes(s)).slice(0, 6);
  }, [search, txns]);

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: `radial-gradient(1200px 600px at 80% -10%, rgba(139,124,246,0.10), transparent), radial-gradient(900px 500px at -10% 110%, rgba(246,177,74,0.07), transparent), ${C.ink}`, color: C.text, fontFamily: "system-ui,-apple-system,Segoe UI,Roboto,sans-serif" }}>
      {/* Sidebar */}
      <aside className="hidden md:flex" style={{ flexDirection: "column", width: 232, padding: 18, borderRight: `1px solid ${C.line}`, position: "sticky", top: 0, height: "100vh" }}>
        <div className="flex items-center" style={{ gap: 10, marginBottom: 26, padding: "4px 6px" }}>
          <div style={{ width: 30, height: 30, borderRadius: 9, background: GRAD }} />
          <div>
            <div style={{ fontWeight: 800, fontSize: 14, letterSpacing: 0.2, lineHeight: 1.15 }}>Your Money<br/>in Focus</div>
          </div>
        </div>
        {NAV.map(n => {
          const A = n.icon, active = page === n.k;
          return (
            <button key={n.k} onClick={() => { setPage(n.k); if (n.k !== "tx") setPresetCat(null); }}
              className="flex items-center" style={{
                gap: 11, padding: "10px 12px", marginBottom: 4, borderRadius: 11, cursor: "pointer", border: "none", textAlign: "left",
                background: active ? "rgba(139,124,246,0.14)" : "transparent",
                color: active ? C.text : C.sub, fontSize: 13.5, fontWeight: active ? 700 : 500,
              }}>
              <A size={17} color={active ? C.violet : C.faint} /> {n.label}
            </button>
          );
        })}
        <div style={{ marginTop: "auto", padding: 12, borderRadius: 12, background: "rgba(255,255,255,0.03)", border: `1px solid ${C.line}` }}>
          <div style={{ color: C.faint, fontSize: 11, marginBottom: 6 }}>Saved & Invested</div>
          <div style={{ fontWeight: 800, fontSize: 18, background: GRAD, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>{R(model.totals.savings)}</div>
          <div style={{ color: C.faint, fontSize: 10.5, marginTop: 2 }}>{model.totals.rate.toFixed(1)}% savings rate</div>
        </div>
      </aside>

      {/* Main */}
      <main style={{ flex: 1, minWidth: 0, padding: narrow ? "14px 10px 52px" : "18px 22px 60px" }}>
        {/* top bar */}
        <div className="flex items-center" style={{ gap: 12, marginBottom: 18, flexWrap: "wrap" }}>
          <div>
            <h1 style={{ fontSize: 22, fontWeight: 800, letterSpacing: -0.3 }}>{NAV.find(n => n.k === page)?.label}</h1>
            <p style={{ color: C.faint, fontSize: 12.5 }}>{model.months.length ? `${mLabel(model.months[0])} – ${mLabel(model.months[model.months.length - 1])} · ` : ""}{txns.length} transactions</p>
          </div>
          <div style={{ marginLeft: "auto", position: "relative" }}>
            <div className="flex items-center" style={{ gap: 8, background: "rgba(255,255,255,0.05)", border: `1px solid ${C.line}`, borderRadius: 11, padding: "9px 13px", width: 230 }}>
              <Search size={15} color={C.faint} />
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search everything…"
                style={{ background: "transparent", border: "none", outline: "none", color: C.text, fontSize: 13, width: "100%" }} />
            </div>
            {globalResults.length > 0 && (
              <div style={{ position: "absolute", top: 46, right: 0, width: 320, background: C.panel2, border: `1px solid ${C.line}`, borderRadius: 12, padding: 6, zIndex: 30, boxShadow: "0 20px 50px rgba(0,0,0,0.5)" }}>
                {globalResults.map((t, i) => (
                  <button key={i} onClick={() => { goCat(t.c); setSearch(""); }} className="flex items-center justify-between" style={{ width: "100%", gap: 8, padding: "8px 10px", background: "transparent", border: "none", cursor: "pointer", borderRadius: 8 }}>
                    <span className="flex items-center" style={{ gap: 8, minWidth: 0 }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: catColor(t.c), flexShrink: 0 }} />
                      <span style={{ color: C.text, fontSize: 12.5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.n}</span>
                    </span>
                    <span style={{ color: C.sub, fontSize: 12, flexShrink: 0 }}>{R(t.o || t.i, 2)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* mobile nav */}
        <div className="flex md:hidden" style={{ gap: 8, overflowX: "auto", marginBottom: 16, paddingBottom: 4 }}>
          {NAV.map(n => (
            <button key={n.k} onClick={() => { setPage(n.k); if (n.k !== "tx") setPresetCat(null); }}
              style={{ whiteSpace: "nowrap", padding: "8px 13px", borderRadius: 10, border: `1px solid ${C.line}`, background: page === n.k ? "rgba(139,124,246,0.16)" : "transparent", color: page === n.k ? C.text : C.sub, fontSize: 12.5, fontWeight: 600 }}>
              {n.label}
            </button>
          ))}
        </div>

        {txns.length === 0 && page !== "settings" && page !== "tfsa"
          ? <EmptyState onImport={() => setPage("settings")} />
          : <>
            {page === "dash" && <Dashboard model={model} goCat={goCat} />}
            {page === "tx" && <Transactions txns={txns} overrides={overrides} onReclassify={onReclassify} presetCat={presetCat} presetClear={() => setPresetCat(null)} />}
            {page === "cat" && <Categories model={model} goTx={goCat} />}
            {page === "trends" && <Trends model={model} />}
            {page === "flow" && <CashFlow model={model} />}
            {page === "insights" && <Insights model={model} />}
          </>}
        {page === "tfsa" && <TFSAPage txns={txns} overrides={overrides} />}
        {page === "settings" && <SettingsPage setTxns={t=>{setTxns(t);setPage("dash");}} status={status} setStatus={setStatus} onExportAll={onExportAll} onClearTags={clearTags} tagCount={Object.keys(overrides).length} />}
      </main>
    </div>
  );
}
