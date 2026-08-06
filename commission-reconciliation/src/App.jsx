import { useState, useMemo, useCallback, useEffect, Fragment } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

/* =========================================================================
   Commission Reconciliation — shared, login-gated version
   Sources: Cobra (what BT paid) · NetSuite (what we expect/recorded) · Sch5 (BT source feed)
   Join key: BT order number  (Cobra "Job Header" = NetSuite "Order ref" = Sch5 "MAIN ORDER NUM")
   Data is uploaded once and saved to Supabase, so everyone signed in sees the same view.
   ========================================================================= */

// ============ PASTE YOUR SUPABASE DETAILS HERE ============
// Both come from Supabase → Project Settings → API.
// The anon key is SAFE to be public (Supabase designed it that way) — the real
// protection is the login + the email allow-list you set up in the SQL step.
// The URL below is already your existing project; just paste the anon key.
const SUPABASE_URL = "https://xrekebgnubhjqtpllbcz.supabase.co";
const SUPABASE_ANON_KEY = "PASTE-YOUR-ANON-KEY-HERE";
// ==========================================================

const supabase =
  SUPABASE_URL.startsWith("https://") && !SUPABASE_ANON_KEY.startsWith("PASTE")
    ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
    : null;

// a throwaway client for creating other people's logins WITHOUT signing the
// admin out of their own session (signUp normally hijacks the current session)
const makeSignupClient = () =>
  createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: { persistSession: false, autoRefreshToken: false, storageKey: "recon-signup-temp" },
  });

// fallback passcode — ONLY used if Supabase isn't configured yet (local testing)
const PASSCODE = "BTLBDCSDTEST";

// ---------- helpers ----------
const money = (v) => {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isNaN(v) ? null : v;
  const n = parseFloat(String(v).replace(/[£,\s]/g, ""));
  return isNaN(n) ? null : n;
};
const sum = (arr) => arr.reduce((a, b) => a + (b || 0), 0);
const normOrder = (v) => (v == null ? "" : String(v).toUpperCase().replace(/\s+/g, "").trim());
const normLE = (v) => (v == null ? "" : String(v).trim().replace(/^0+/, ""));
const isYes = (v) => v != null && String(v).trim().toLowerCase().startsWith("y");
const gbp = (n) =>
  n == null
    ? "—"
    : new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(n);
const firstDefined = (...xs) => xs.find((x) => x != null && x !== "") ?? null;

// derive a YYYY-MM key from many date shapes (Date, ISO string, UK D/M/Y, YYYYMM)
const monthKey = (v) => {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}`;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})/); // ISO
  if (m) return `${m[1]}-${m[2]}`;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/); // UK day/month/year
  if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; return `${y}-${String(Number(m[2])).padStart(2, "0")}`; }
  m = s.match(/^(\d{4})(\d{2})$/); // YYYYMM (e.g. 202601)
  if (m) return `${m[1]}-${m[2]}`;
  return null;
};
const periodLabel = (key) => {
  if (!key) return "No date";
  const [y, mo] = key.split("-");
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${names[Number(mo) - 1] || mo} ${y}`;
};

// resolve a value from a raw row by trying candidate header names (exact, then fuzzy)
const pick = (row, candidates) => {
  const keys = Object.keys(row);
  const clean = (s) => s.toLowerCase().replace(/\s+/g, " ").trim();
  for (const c of candidates) {
    const hit = keys.find((k) => clean(k) === clean(c));
    if (hit !== undefined) return row[hit];
  }
  const squash = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  for (const c of candidates) {
    const t = squash(c);
    const hit = keys.find((k) => squash(k).includes(t) && t.length > 2);
    if (hit) return row[hit];
  }
  return undefined;
};

// ---------- file parsing (CSV or Excel) ----------
async function parseWorkbook(file) {
  const isCsv = /\.(csv|tsv|txt)$/i.test(file.name || "");
  let wb;
  if (isCsv) {
    const text = await file.text();
    wb = XLSX.read(text, { type: "string", cellDates: true });
  } else {
    const buf = await file.arrayBuffer();
    wb = XLSX.read(buf, { cellDates: true });
  }
  // pick the sheet with the most rows
  let best = wb.SheetNames[0];
  let bestCount = -1;
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const ref = ws["!ref"];
    const rows = ref ? XLSX.utils.decode_range(ref).e.r : 0;
    if (rows > bestCount) {
      bestCount = rows;
      best = sn;
    }
  }
  const ws = wb.Sheets[best];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
  // find header row (first with >=3 non-empty string cells)
  let hi = 0;
  for (let i = 0; i < Math.min(aoa.length, 12); i++) {
    const strs = (aoa[i] || []).filter((c) => typeof c === "string" && c.trim()).length;
    if (strs >= 3) {
      hi = i;
      break;
    }
  }
  const headers = (aoa[hi] || []).map((h, i) =>
    h == null || String(h).trim() === "" ? `col${i}` : String(h).trim()
  );
  const rows = [];
  for (let i = hi + 1; i < aoa.length; i++) {
    const r = aoa[i];
    if (!r || r.every((c) => c == null || c === "")) continue;
    const obj = {};
    headers.forEach((h, ci) => (obj[h] = r[ci] ?? null));
    rows.push(obj);
  }
  return { sheet: best, headers, rows, sheetNames: wb.SheetNames };
}

// ---------- canonical row extractors ----------
const nsRow = (r) => ({
  src: "netsuite",
  oppId: firstDefined(pick(r, ["Opp ID"])),
  orderNum: normOrder(pick(r, ["Order ref"])),
  company: firstDefined(pick(r, ["Company Name"])),
  status: firstDefined(pick(r, ["Order Status"])),
  sov: money(pick(r, ["Contract Value"])),
  product: firstDefined(pick(r, ["Item: Name (Grouped)", "Item Name"])),
  expected: money(pick(r, ["Product GP"])), // what we expect to be paid
  recordedCobra: money(pick(r, ["Cobra Payment"])), // what NS records Cobra paid
  itemPaid: firstDefined(pick(r, ["Item Paid"])),
  overpayment: firstDefined(pick(r, ["Overpayment"])),
  accelerator: firstDefined(pick(r, ["Accelerator?", "Accelerator"])),
  agent: firstDefined(pick(r, ["Admin Agent"])),
  le: normLE(pick(r, ["Customer Le", "Customer Ledger", "Customer Le "])),
  date: pick(r, ["Netsuite Date"]),
  raw: r,
});

const cobraRow = (r) => ({
  src: "cobra",
  orderNum: normOrder(pick(r, ["Job Header"])),
  optyId: firstDefined(pick(r, ["Opty ID"])),
  le: normLE(pick(r, ["LE Code"])),
  company: firstDefined(pick(r, ["Customer Name"])),
  status: firstDefined(pick(r, ["Status"])),
  month: firstDefined(pick(r, ["Month"])),
  sov: money(pick(r, ["Contract Value"])),
  due: money(pick(r, ["Commission Due"])), // what BT says is owed
  paid: money(pick(r, ["Commission Paid"])), // what BT actually paid
  product: firstDefined(pick(r, ["Measure", "Plan Name"])),
  prodCode: firstDefined(pick(r, ["Prod Code"])),
  date: pick(r, ["Closed Date", "Order Date"]),
  raw: r,
});

const sch5Row = (r) => ({
  src: "sch5",
  orderNum: normOrder(pick(r, ["MAIN ORDER NUM", "ADT REF"])),
  optyId: firstDefined(pick(r, ["OPPORTUNITY ID"])),
  le: normLE(pick(r, ["LE CODE"])),
  company: firstDefined(pick(r, ["CUSTOMER NAME"])),
  status: firstDefined(pick(r, ["ORDER STATUS"])),
  subStatus: firstDefined(pick(r, ["ORDER SUB STATUS"])),
  cancelDate: firstDefined(pick(r, ["CANCEL DATE"])),
  sov: money(pick(r, ["SOV"])),
  commissionable: pick(r, ["COMMISSION FLAG"]),
  product: firstDefined(pick(r, ["PRODUCT SUB NAME1"])),
  transactional: firstDefined(pick(r, ["TRANSACTIONAL"])),
  recentCancel: firstDefined(pick(r, ["RECENT CANCELLATION"])),
  date: pick(r, ["CLOSED DATE", "ORDER DATE", "REPORTING MONTH"]),
  raw: r,
});

const groupBy = (rows, keyFn) => {
  const m = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
};

// ---------- the reconciliation engine ----------
function reconcile(files, tol, period = "all") {
  const ns = (files.netsuite?.rows || []).map(nsRow);
  const cb = (files.cobra?.rows || []).map(cobraRow);
  const s5 = (files.sch5?.rows || []).map(sch5Row);

  const nsBy = groupBy(ns.filter((r) => r.orderNum), (r) => r.orderNum);
  const cbBy = groupBy(cb.filter((r) => r.orderNum), (r) => r.orderNum);
  const s5By = groupBy(s5.filter((r) => r.orderNum), (r) => r.orderNum);

  const keys = new Set([...nsBy.keys(), ...cbBy.keys(), ...s5By.keys()]);
  const within = (a, b) => {
    const d = Math.abs((a || 0) - (b || 0));
    const base = Math.max(Math.abs(a || 0), Math.abs(b || 0), 1);
    return d <= tol.abs || d / base <= tol.pct / 100;
  };

  const records = [];
  for (const key of keys) {
    const nsL = nsBy.get(key) || [];
    const cbL = cbBy.get(key) || [];
    const s5L = s5By.get(key) || [];
    const inNS = nsL.length > 0;
    const inCobra = cbL.length > 0;
    const inSch5 = s5L.length > 0;

    const expected = sum(nsL.map((r) => r.expected)); // NS Product GP
    const recordedCobra = sum(nsL.map((r) => r.recordedCobra)); // NS "Cobra Payment"
    const due = sum(cbL.map((r) => r.due)); // Cobra Commission Due
    const paid = sum(cbL.map((r) => r.paid)); // Cobra Commission Paid
    const nsSov = sum(nsL.map((r) => r.sov));
    const s5Sov = sum(s5L.map((r) => r.sov));

    const sch5Cancelled =
      inSch5 &&
      s5L.some(
        (r) =>
          (r.status && /cancel/i.test(String(r.status))) ||
          (r.subStatus && /cancel|cease/i.test(String(r.subStatus))) ||
          r.cancelDate != null ||
          isYes(r.recentCancel)
      );
    const sch5NonComm =
      inSch5 && s5L.some((r) => r.commissionable != null && !isYes(r.commissionable));

    const company =
      firstDefined(nsL[0]?.company, cbL[0]?.company, s5L[0]?.company) || "(no name)";
    const agent = firstDefined(nsL[0]?.agent);
    const anyUnpaid = nsL.some((r) => r.itemPaid != null && !isYes(r.itemPaid));
    const overpaymentFlag = nsL.some((r) => r.overpayment != null && r.overpayment !== "");

    const payDelta = inNS && inCobra ? paid - expected : null; // BT paid vs expected
    const recordDelta = inNS && inCobra ? paid - recordedCobra : null; // NS record vs actual
    const dueVsPaid = inCobra ? paid - due : null; // BT paid vs owed
    const sovDelta = inNS && inSch5 ? nsSov - s5Sov : null;

    // exception detection (ordered by severity)
    const flags = [];
    if (sch5Cancelled && paid > 0)
      flags.push({ code: "CANCELLED_BUT_PAID", sev: 3, msg: "BT cancelled the order (Sch5) but a commission was paid — clawback risk" });
    if (inNS && expected > 0 && !inCobra)
      flags.push({ code: "MISSING_PAYMENT", sev: 3, msg: "Expected in NetSuite but no Cobra payment line — possible unpaid / missing" });
    if (sch5NonComm && expected > 0)
      flags.push({ code: "NON_COMMISSIONABLE", sev: 3, msg: "Sch5 flags order as non-commissionable, but we expect commission" });
    if (inNS && inCobra && payDelta != null && !within(paid, expected))
      flags.push({ code: "AMOUNT_MISMATCH", sev: 2, msg: `BT paid ${gbp(paid)} vs expected ${gbp(expected)} (Δ ${gbp(payDelta)})` });
    if (inNS && inCobra && recordDelta != null && !within(paid, recordedCobra))
      flags.push({ code: "NS_RECORD_OFF", sev: 2, msg: `NetSuite records ${gbp(recordedCobra)} but Cobra shows ${gbp(paid)} — NetSuite figure looks inaccurate` });
    if (inCobra && !inNS)
      flags.push({ code: "UNEXPECTED_PAYMENT", sev: 2, msg: "Cobra paid but no matching NetSuite line" });
    if (overpaymentFlag)
      flags.push({ code: "OVERPAYMENT_FLAG", sev: 2, msg: "NetSuite has an overpayment flag on this order" });
    if (sovDelta != null && !within(nsSov, s5Sov))
      flags.push({ code: "SOV_MISMATCH", sev: 1, msg: `SOV differs — NetSuite ${gbp(nsSov)} vs Sch5 ${gbp(s5Sov)}` });
    if (inNS && !inSch5)
      flags.push({ code: "NOT_IN_SCH5", sev: 1, msg: "Order not found in the Sch5 (BT source) feed" });

    let matchState = "unmatched";
    if (inNS && inCobra) matchState = flags.some((f) => f.sev >= 2) ? "mismatch" : "matched";

    records.push({
      key, orderNum: key, company, agent,
      inNS, inCobra, inSch5,
      expected, recordedCobra, due, paid, nsSov, s5Sov,
      payDelta, recordDelta, dueVsPaid, sovDelta,
      sch5Cancelled, sch5NonComm, anyUnpaid, overpaymentFlag,
      period: monthKey(nsL[0]?.date) || monthKey(cbL[0]?.date) || monthKey(s5L[0]?.date) || null,
      status: firstDefined(nsL[0]?.status, cbL[0]?.status, s5L[0]?.status),
      sch5Status: firstDefined(s5L[0]?.status),
      flags, matchState,
      nsL, cbL, s5L,
    });
  }

  records.sort((a, b) => {
    const sa = Math.max(0, ...a.flags.map((f) => f.sev));
    const sb = Math.max(0, ...b.flags.map((f) => f.sev));
    return sb - sa || (b.expected || 0) - (a.expected || 0);
  });

  // ---- forecast: NetSuite lines not yet paid ----
  const unpaidLines = ns.filter((r) => r.itemPaid != null && !isYes(r.itemPaid) && (r.expected || 0) > 0);
  const forecastAll = unpaidLines.map((r) => {
    const s5L = s5By.get(r.orderNum) || [];
    const cancelled = s5L.some(
      (x) => (x.status && /cancel/i.test(String(x.status))) || x.cancelDate != null
    );
    const nonComm = s5L.some((x) => x.commissionable != null && !isYes(x.commissionable));
    const inS5 = s5L.length > 0;
    let verdict = "expected";
    if (cancelled) verdict = "at_risk_cancelled";
    else if (nonComm) verdict = "at_risk_noncomm";
    else if (!inS5) verdict = "unverified";
    return { ...r, verdict, period: monthKey(r.date) || null };
  });

  // full list of months present, for the dropdown (built before filtering)
  const periods = [...new Set(records.map((r) => r.period).filter(Boolean))].sort().reverse();

  // apply the selected-month filter
  const inPeriod = (p) => period === "all" || p === period;
  const recordsF = records.filter((r) => inPeriod(r.period));
  const forecast = forecastAll.filter((r) => inPeriod(r.period));

  const overlap = recordsF.filter((r) => {
    let c = 0;
    if (r.inNS) c++;
    if (r.inCobra) c++;
    if (r.inSch5) c++;
    return c >= 2;
  }).length;

  return {
    records: recordsF, forecast, periods,
    counts: { ns: ns.length, cb: cb.length, s5: s5.length, overlap },
  };
}

// ---------- CSV export ----------
function downloadCSV(rows, filename) {
  if (!rows.length) return;
  const cols = Object.keys(rows[0]);
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv = [cols.join(","), ...rows.map((r) => cols.map((c) => esc(r[c])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// =========================================================================
//  UI
// =========================================================================
const STYLES = `
* { box-sizing: border-box; }
.recon { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color:#1b1636; background:#f5f5fa; min-height:100%; padding:22px; line-height:1.45; }
.mono { font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace; font-variant-numeric: tabular-nums; }
.num { font-variant-numeric: tabular-nums; text-align:right; white-space:nowrap; }
.wrap { max-width:1180px; margin:0 auto; }
.head { display:flex; align-items:baseline; gap:12px; flex-wrap:wrap; margin-bottom:2px; }
.head h1 { font-size:22px; margin:0; letter-spacing:-.02em; }
.head .accent { color:#5514b4; }
.sub { color:#6b6784; font-size:13px; margin:0 0 18px; }
.panel { background:#fff; border:1px solid #e7e6f0; border-radius:12px; padding:16px; margin-bottom:16px; }
.panel h2 { font-size:13px; text-transform:uppercase; letter-spacing:.06em; color:#8a8aa3; margin:0 0 12px; }
.uploads { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; }
@media (max-width:820px){ .uploads{ grid-template-columns:1fr; } }
.slot { border:1.5px dashed #d3d0e6; border-radius:10px; padding:14px; transition:.15s; }
.slot.loaded { border-style:solid; border-color:#c4bce6; background:#faf9ff; }
.slot .role { font-weight:700; font-size:14px; }
.slot .desc { color:#7a7690; font-size:12px; margin:2px 0 10px; }
.slot .status { font-size:12px; color:#14804a; margin-top:8px; }
.slot .err { color:#b3261e; }
label.file { display:inline-block; cursor:pointer; background:#5514b4; color:#fff; font-size:12.5px;
  font-weight:600; padding:7px 12px; border-radius:7px; }
label.file input { display:none; }
label.file.reload { background:#efeaff; color:#5514b4; }
.tabs { display:flex; gap:4px; flex-wrap:wrap; margin-bottom:14px; }
.tab { border:1px solid #e2e0ee; background:#fff; color:#4b4766; padding:8px 14px; border-radius:8px;
  font-size:13px; font-weight:600; cursor:pointer; }
.tab.active { background:#5514b4; color:#fff; border-color:#5514b4; }
.kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; }
@media (max-width:820px){ .kpis{ grid-template-columns:repeat(2,1fr); } }
.kpi { background:#fff; border:1px solid #e7e6f0; border-radius:12px; padding:14px 16px; }
.kpi .lab { font-size:11.5px; text-transform:uppercase; letter-spacing:.05em; color:#8a8aa3; }
.kpi .val { font-size:24px; font-weight:700; margin-top:4px; letter-spacing:-.02em; }
.kpi .val.mono { font-size:22px; }
.kpi .foot { font-size:12px; color:#7a7690; margin-top:2px; }
table { width:100%; border-collapse:collapse; font-size:13px; }
th { text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#8a8aa3;
  border-bottom:1px solid #eceaf4; padding:8px 10px; position:sticky; top:0; background:#fff; }
th.num { text-align:right; }
td { padding:9px 10px; border-bottom:1px solid #f1f0f8; vertical-align:top; }
tr.click { cursor:pointer; }
tr.click:hover { background:#faf9ff; }
.chip { display:inline-block; font-size:11px; font-weight:600; padding:2px 8px; border-radius:20px; }
.chip.matched { background:#e6f4ec; color:#14804a; }
.chip.mismatch { background:#fbf0d9; color:#8a5a00; }
.chip.unmatched { background:#eef0f4; color:#5b6472; }
.chip.risk { background:#fbe9e7; color:#b3261e; }
.pres { display:inline-flex; gap:3px; }
.dot { width:20px; height:18px; border-radius:5px; font-size:10px; font-weight:700; display:inline-flex;
  align-items:center; justify-content:center; color:#fff; }
.dot.on { background:#5514b4; } .dot.off { background:#e2e0ee; color:#a7a3bf; }
.exc { border-left:4px solid #cfd8e3; background:#fff; border:1px solid #eceaf4; border-left-width:4px;
  border-radius:8px; padding:11px 14px; margin-bottom:8px; }
.exc.s3 { border-left-color:#b3261e; } .exc.s2 { border-left-color:#d98a00; } .exc.s1 { border-left-color:#8a8aa3; }
.exc .top { display:flex; justify-content:space-between; gap:10px; align-items:baseline; flex-wrap:wrap; }
.exc .code { font-size:11px; font-weight:700; letter-spacing:.04em; }
.exc.s3 .code { color:#b3261e; } .exc.s2 .code { color:#8a5a00; } .exc.s1 .code { color:#5b6472; }
.exc .msg { font-size:13px; color:#3a3556; margin-top:3px; }
.exc .meta { font-size:12px; color:#8a8aa3; }
.detail { background:#faf9ff; border-radius:8px; padding:10px; margin-top:6px; }
.detail table { font-size:12px; } .detail th { background:#faf9ff; }
.banner { background:#fff8e6; border:1px solid #f0dfa8; color:#7a5b00; border-radius:10px;
  padding:12px 14px; font-size:13px; margin-bottom:16px; }
.banner.info { background:#f0ecff; border-color:#d6cbf5; color:#4b3a8a; }
.empty { text-align:center; color:#8a8aa3; padding:40px 20px; font-size:14px; }
.pos { color:#14804a; } .neg { color:#b3261e; }
.settings { display:flex; gap:16px; align-items:center; flex-wrap:wrap; font-size:13px; color:#4b4766; }
.settings input { width:70px; padding:5px 8px; border:1px solid #d3d0e6; border-radius:6px; font-size:13px; }
.btn { background:#efeaff; color:#5514b4; border:none; font-weight:600; font-size:12.5px; padding:7px 12px;
  border-radius:7px; cursor:pointer; }
.note { font-size:12px; color:#8a8aa3; margin-top:10px; }
.lock { max-width:360px; margin:80px auto; }
.lock input { width:100%; padding:11px 13px; border:1px solid #d3d0e6; border-radius:9px; font-size:15px; margin:12px 0; }
.lock button { width:100%; background:#5514b4; color:#fff; border:none; padding:11px; border-radius:9px; font-weight:700; cursor:pointer; }
`;

function Presence({ ns, cb, s5 }) {
  return (
    <span className="pres">
      <span className={"dot " + (ns ? "on" : "off")} title="NetSuite">N</span>
      <span className={"dot " + (cb ? "on" : "off")} title="Cobra">C</span>
      <span className={"dot " + (s5 ? "on" : "off")} title="Sch5">S</span>
    </span>
  );
}

function Delta({ v }) {
  if (v == null) return <span className="num">—</span>;
  const cls = Math.abs(v) < 0.005 ? "" : v > 0 ? "pos" : "neg";
  return <span className={"num mono " + cls}>{v > 0 ? "+" : ""}{gbp(v)}</span>;
}

export default function ReconciliationTool() {
  const [unlocked, setUnlocked] = useState(false);
  const [pass, setPass] = useState("");
  const [files, setFiles] = useState({ cobra: null, netsuite: null, sch5: null });
  const [errors, setErrors] = useState({});
  const [tab, setTab] = useState("dashboard");
  const [tol, setTol] = useState({ abs: 1, pct: 1 });
  const [expanded, setExpanded] = useState(new Set());
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authMsg, setAuthMsg] = useState(null);
  const [sharedMeta, setSharedMeta] = useState(null);
  const [saving, setSaving] = useState(false);
  const [period, setPeriod] = useState("all");
  const [isAdmin, setIsAdmin] = useState(false);
  const [users, setUsers] = useState([]);
  const [newUser, setNewUser] = useState({ email: "", password: "" });
  const [userMsg, setUserMsg] = useState(null);

  // track the logged-in session
  useEffect(() => {
    if (!supabase) { setAuthReady(true); return; }
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setAuthReady(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  // once signed in, load the shared dataset so everyone sees the same thing
  useEffect(() => {
    if (!supabase || !session) return;
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("recon_datasets").select("*").eq("id", "current").maybeSingle();
      if (cancelled || error || !data) return;
      const next = {};
      for (const k of ["cobra", "netsuite", "sch5"]) if (data[k]) next[k] = data[k];
      if (Object.keys(next).length) setFiles((f) => ({ ...f, ...next }));
      setSharedMeta({ by: data.uploaded_by, at: data.uploaded_at });
    })();
    return () => { cancelled = true; };
  }, [session]);

  // work out if the signed-in user is an admin, and load the user list
  const loadUsers = useCallback(async () => {
    if (!supabase || !session) return;
    const { data } = await supabase.from("recon_allowlist").select("email, is_admin").order("email");
    if (data) {
      setUsers(data);
      const me = (session.user?.email || "").toLowerCase();
      setIsAdmin(data.some((u) => (u.email || "").toLowerCase() === me && u.is_admin));
    }
  }, [session]);

  useEffect(() => { loadUsers(); }, [loadUsers]);

  // admin: create a login (email + password) and add them to the allow-list
  const addUser = useCallback(async () => {
    setUserMsg(null);
    const em = newUser.email.trim().toLowerCase();
    const pw = newUser.password;
    if (!em || pw.length < 6) { setUserMsg({ err: true, text: "Enter an email and a password of at least 6 characters." }); return; }
    // create the auth account on a throwaway client so it doesn't sign you out
    const tmp = makeSignupClient();
    const { error: sErr } = await tmp.auth.signUp({ email: em, password: pw });
    if (sErr && !/already registered/i.test(sErr.message)) {
      setUserMsg({ err: true, text: sErr.message }); return;
    }
    // add to the allow-list so they can actually see the data
    const { error: aErr } = await supabase.from("recon_allowlist").upsert({ email: em }, { onConflict: "email" });
    if (aErr) { setUserMsg({ err: true, text: "Account made, but couldn't add to the allow-list: " + aErr.message }); return; }
    setNewUser({ email: "", password: "" });
    setUserMsg({ err: false, text: `${em} can now sign in with that password.` });
    loadUsers();
  }, [newUser, loadUsers]);

  const removeUser = useCallback(async (em) => {
    if ((em || "").toLowerCase() === (session?.user?.email || "").toLowerCase()) return; // don't remove yourself
    await supabase.from("recon_allowlist").delete().eq("email", em);
    loadUsers();
  }, [session, loadUsers]);
  const saveShared = useCallback(async (nextFiles) => {
    if (!supabase || !session) return;
    setSaving(true);
    const payload = {
      id: "current",
      cobra: nextFiles.cobra || null,
      netsuite: nextFiles.netsuite || null,
      sch5: nextFiles.sch5 || null,
      uploaded_by: session.user?.email || null,
      uploaded_at: new Date().toISOString(),
    };
    const { error } = await supabase.from("recon_datasets").upsert(payload, { onConflict: "id" });
    setSaving(false);
    if (error) setErrors((e) => ({ ...e, save: "Couldn't save to the shared store — are you on the allow-list?" }));
    else { setSharedMeta({ by: payload.uploaded_by, at: payload.uploaded_at }); setErrors((e) => ({ ...e, save: null })); }
  }, [session]);

  const onFile = useCallback(async (which, file) => {
    if (!file) return;
    try {
      const parsed = await parseWorkbook(file);
      const entry = { sheet: parsed.sheet, headers: parsed.headers, rows: parsed.rows, name: file.name };
      let nextFiles;
      setFiles((f) => (nextFiles = { ...f, [which]: entry }));
      setErrors((e) => ({ ...e, [which]: null }));
      saveShared(nextFiles);
    } catch (err) {
      setErrors((e) => ({ ...e, [which]: "Couldn't read that file — is it a valid CSV or Excel file?" }));
    }
  }, [saveShared]);

  const result = useMemo(() => reconcile(files, tol, period), [files, tol, period]);
  const anyLoaded = files.cobra || files.netsuite || files.sch5;
  const allLoaded = files.cobra && files.netsuite && files.sch5;

  const totals = useMemo(() => {
    const r = result.records;
    return {
      expected: sum(r.map((x) => x.expected)),
      paid: sum(r.map((x) => x.paid)),
      variance: sum(r.map((x) => x.paid)) - sum(r.map((x) => x.expected)),
      matched: r.filter((x) => x.matchState === "matched").length,
      mismatch: r.filter((x) => x.matchState === "mismatch").length,
      risk: r.filter((x) => x.flags.some((f) => f.sev === 3)).length,
      exceptions: r.filter((x) => x.flags.length).length,
    };
  }, [result]);

  const toggle = (k) =>
    setExpanded((s) => {
      const n = new Set(s);
      n.has(k) ? n.delete(k) : n.add(k);
      return n;
    });

  // ----- access gate -----
  if (supabase) {
    if (!authReady) {
      return (
        <div className="recon"><style>{STYLES}</style>
          <div className="lock panel"><p className="sub">Loading…</p></div>
        </div>
      );
    }
    if (!session) {
      const signIn = async () => {
        setAuthMsg(null);
        const { error } = await supabase.auth.signInWithPassword({
          email: email.trim(),
          password,
        });
        if (error) setAuthMsg({ err: true, text: "Email or password not recognised." });
      };
      return (
        <div className="recon"><style>{STYLES}</style>
          <div className="lock panel">
            <h1 style={{ margin: 0, fontSize: 20 }}>
              Commission <span style={{ color: "#5514b4" }}>Reconciliation</span>
            </h1>
            <p className="sub" style={{ margin: "6px 0 0" }}>
              Sign in to view the shared reconciliation.
            </p>
            <input
              type="email"
              value={email}
              placeholder="Email"
              onChange={(e) => setEmail(e.target.value)}
            />
            <input
              type="password"
              value={password}
              placeholder="Password"
              onChange={(e) => setPassword(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && email.trim() && password && signIn()}
            />
            <button onClick={() => email.trim() && password && signIn()}>Sign in</button>
            {authMsg && (
              <p className="sub" style={{ color: authMsg.err ? "#b3261e" : "#14804a", marginTop: 10 }}>{authMsg.text}</p>
            )}
          </div>
        </div>
      );
    }
  }

  if (!supabase && !unlocked) {
    return (
      <div className="recon">
        <style>{STYLES}</style>
        <div className="lock panel">
          <h1 style={{ margin: 0, fontSize: 20 }}>
            Commission <span style={{ color: "#5514b4" }}>Reconciliation</span>
          </h1>
          <p className="sub" style={{ margin: "6px 0 0" }}>Restricted test tool. Enter the access code.</p>
          <p className="sub" style={{ margin: "6px 0 0", color: "#b3261e" }}>
            Not connected to the database — this is local-only mode (no logins, uploads won't be saved or shared).
            To turn on email/password logins + sharing, paste your Supabase anon key into App.jsx.
          </p>
          <input
            type="password"
            value={pass}
            placeholder="Access code"
            onChange={(e) => setPass(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && pass === PASSCODE && setUnlocked(true)}
          />
          <button onClick={() => pass === PASSCODE && setUnlocked(true)}>Unlock</button>
          {pass && pass !== PASSCODE && (
            <p className="sub" style={{ color: "#b3261e", marginTop: 10 }}>That code isn't right.</p>
          )}
        </div>
      </div>
    );
  }

  const TABS = [
    ["dashboard", "Dashboard"],
    ["reconcile", "Reconciliation"],
    ["btpay", "BT Payment Check"],
    ["exceptions", "Exceptions & Risk"],
    ["obi", "OBI Checks"],
    ...(isAdmin ? [["users", "Users"]] : []),
  ];

  return (
    <div className="recon">
      <style>{STYLES}</style>
      <div className="wrap">
        <div className="head">
          <h1>Commission <span className="accent">Reconciliation</span></h1>
          <span className="sub" style={{ margin: 0 }}>Cobra · NetSuite · Sch5 — shared</span>
          {session && (
            <span className="sub" style={{ marginLeft: "auto" }}>
              {session.user?.email}{" · "}
              <a href="#" onClick={async (e) => {
                e.preventDefault();
                const np = window.prompt("Enter a new password (at least 6 characters):");
                if (!np) return;
                if (np.length < 6) { window.alert("Password must be at least 6 characters."); return; }
                const { error } = await supabase.auth.updateUser({ password: np });
                window.alert(error ? ("Couldn't change password: " + error.message) : "Password updated. Use it next time you sign in.");
              }} style={{ color: "#5514b4" }}>Change password</a>
              {" · "}
              <a href="#" onClick={(e) => { e.preventDefault(); supabase.auth.signOut(); }} style={{ color: "#5514b4" }}>Sign out</a>
            </span>
          )}
        </div>
        <p className="sub">
          {sharedMeta?.at
            ? `Shared data last updated by ${sharedMeta.by || "someone"} on ${new Date(sharedMeta.at).toLocaleString("en-GB")}.`
            : "Upload the three exports once — they're saved and shared with everyone signed in."}
          {saving && " · Saving…"}
          {errors.save && <span style={{ color: "#b3261e" }}> · {errors.save}</span>}
        </p>

        {!supabase && (
          <div className="banner">
            <strong>Not connected to the database.</strong> Running in local mode — uploads work but are
            <em> not saved or shared</em>, and there's no login. Paste your Supabase anon key into
            <span className="mono"> App.jsx</span> (the <span className="mono">SUPABASE_ANON_KEY</span> line) and redeploy to switch on logins + sharing.
          </div>
        )}

        {/* uploads */}
        <div className="panel">
          <h2>1 · Upload exports</h2>
          <div className="uploads">
            {[
              ["cobra", "Cobra", "What BT actually paid"],
              ["netsuite", "NetSuite", "What we expect & recorded"],
              ["sch5", "Sch5", "BT source feed — status, cancels, SOV"],
            ].map(([key, role, desc]) => {
              const f = files[key];
              const err = errors[key];
              return (
                <div key={key} className={"slot " + (f ? "loaded" : "")}>
                  <div className="role">{role}</div>
                  <div className="desc">{desc}</div>
                  <label className={"file " + (f ? "reload" : "")}>
                    {f ? "Replace file" : "Choose CSV"}
                    <input type="file" accept=".csv,.tsv,.txt,.xlsx,.xls" onChange={(e) => onFile(key, e.target.files[0])} />
                  </label>
                  {f && (
                    <div className="status">
                      ✓ {f.rows.length} rows · sheet "{f.sheet}"
                      <br />
                      <span style={{ color: "#8a8aa3" }}>{f.name}</span>
                    </div>
                  )}
                  {err && <div className="status err">{err}</div>}
                </div>
              );
            })}
          </div>
          <div className="settings" style={{ marginTop: 14 }}>
            <span>Month to check:</span>
            <select value={period} onChange={(e) => setPeriod(e.target.value)}
              style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
              <option value="all">All months</option>
              {result.periods.map((p) => (
                <option key={p} value={p}>{periodLabel(p)}</option>
              ))}
            </select>
            <span style={{ width: 18 }} />
            <span>Match tolerance:</span>
            <label>£<input type="number" value={tol.abs} min={0} step={0.5}
              onChange={(e) => setTol((t) => ({ ...t, abs: Number(e.target.value) }))} /></label>
            <label>or <input type="number" value={tol.pct} min={0} step={0.5}
              onChange={(e) => setTol((t) => ({ ...t, pct: Number(e.target.value) }))} />%</label>
            <span style={{ color: "#8a8aa3" }}>anything within this counts as a match.</span>
          </div>
        </div>

        {!anyLoaded && (
          <div className="empty panel">Upload at least NetSuite and Cobra to start reconciling.</div>
        )}

        {anyLoaded && allLoaded && result.counts.overlap === 0 && (
          <div className="banner">
            <strong>No shared order numbers across your files.</strong> These look like disjoint sample
            exports — every order appears in only one file, so nothing can cross-match yet. Load Cobra,
            NetSuite and Sch5 exports that cover the <em>same orders / period</em> to see live reconciliation.
          </div>
        )}

        {anyLoaded && (
          <>
            <div className="tabs">
              {TABS.map(([k, l]) => (
                <button key={k} className={"tab " + (tab === k ? "active" : "")} onClick={() => setTab(k)}>{l}</button>
              ))}
            </div>

            {/* DASHBOARD */}
            {tab === "dashboard" && (
              <>
                <div className="kpis">
                  <div className="kpi"><div className="lab">Expected (NetSuite GP)</div><div className="val mono">{gbp(totals.expected)}</div><div className="foot">across {result.counts.ns} lines</div></div>
                  <div className="kpi"><div className="lab">Paid (Cobra)</div><div className="val mono">{gbp(totals.paid)}</div><div className="foot">across {result.counts.cb} lines</div></div>
                  <div className="kpi"><div className="lab">Variance</div><div className={"val mono " + (Math.abs(totals.variance) < 1 ? "" : totals.variance < 0 ? "neg" : "pos")}>{gbp(totals.variance)}</div><div className="foot">paid − expected</div></div>
                  <div className="kpi"><div className="lab">Flags to review</div><div className="val" style={{ color: totals.risk ? "#b3261e" : "#14804a" }}>{totals.exceptions}</div><div className="foot">{totals.risk} high-risk</div></div>
                </div>

                <div className="panel" style={{ marginTop: 16 }}>
                  <h2>Forecast — expected commission not yet paid</h2>
                  <ForecastPanel forecast={result.forecast} hasSch5={!!files.sch5} />
                </div>

                <div className="panel">
                  <h2>Commission by agent (NetSuite)</h2>
                  <BreakdownTable records={result.records} />
                </div>
              </>
            )}

            {/* RECONCILIATION */}
            {tab === "reconcile" && (
              <div className="panel">
                <h2>Order-level reconciliation — Cobra vs NetSuite</h2>
                {result.records.length === 0 ? (
                  <div className="empty">Nothing to show yet.</div>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                    <table>
                      <thead>
                        <tr>
                          <th>Order #</th><th>Customer</th><th>Sources</th>
                          <th className="num">Expected</th><th className="num">Paid (Cobra)</th>
                          <th className="num">Δ pay</th><th>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {result.records.map((r) => (
                          <Fragment key={r.key}>
                            <tr className="click" onClick={() => toggle(r.key)}>
                              <td className="mono">{r.orderNum || "—"}</td>
                              <td>{r.company}</td>
                              <td><Presence ns={r.inNS} cb={r.inCobra} s5={r.inSch5} /></td>
                              <td className="num mono">{gbp(r.expected)}</td>
                              <td className="num mono">{gbp(r.paid)}</td>
                              <td><Delta v={r.payDelta} /></td>
                              <td>
                                <span className={"chip " + r.matchState}>{r.matchState}</span>
                                {r.flags.some((f) => f.sev === 3) && <span className="chip risk" style={{ marginLeft: 4 }}>risk</span>}
                              </td>
                            </tr>
                            {expanded.has(r.key) && (
                              <tr><td colSpan={7}><LineDetail r={r} /></td></tr>
                            )}
                          </Fragment>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                <p className="note">Click any row for the underlying lines. N/C/S = present in NetSuite / Cobra / Sch5.</p>
              </div>
            )}

            {/* BT PAYMENT CHECK */}
            {tab === "btpay" && (
              <div className="panel">
                <h2>Did BT pay correctly — and is NetSuite's record accurate?</h2>
                <p className="note" style={{ marginTop: 0, marginBottom: 12 }}>
                  <strong>Paid vs Due</strong> = did BT pay what Cobra says is owed · <strong>Paid vs Expected</strong> = did BT pay what we expected ·
                  <strong> NS record</strong> = does NetSuite's "Cobra Payment" match what Cobra actually shows.
                </p>
                <div style={{ overflowX: "auto" }}>
                  <table>
                    <thead>
                      <tr>
                        <th>Order #</th><th>Customer</th>
                        <th className="num">Due (Cobra)</th><th className="num">Paid (Cobra)</th><th className="num">Paid−Due</th>
                        <th className="num">Expected (NS)</th><th className="num">Paid−Exp</th>
                        <th className="num">NS record</th><th className="num">vs actual</th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.records.filter((r) => r.inCobra).map((r) => (
                        <tr key={r.key}>
                          <td className="mono">{r.orderNum}</td>
                          <td>{r.company}</td>
                          <td className="num mono">{gbp(r.due)}</td>
                          <td className="num mono">{gbp(r.paid)}</td>
                          <td><Delta v={r.dueVsPaid} /></td>
                          <td className="num mono">{r.inNS ? gbp(r.expected) : "—"}</td>
                          <td><Delta v={r.payDelta} /></td>
                          <td className="num mono">{r.inNS ? gbp(r.recordedCobra) : "—"}</td>
                          <td><Delta v={r.recordDelta} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {result.records.filter((r) => r.inCobra).length === 0 && <div className="empty">Load a Cobra export to run payment checks.</div>}
              </div>
            )}

            {/* EXCEPTIONS */}
            {tab === "exceptions" && (
              <div className="panel">
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <h2 style={{ margin: 0 }}>Mismatches, missing payments & financial risk</h2>
                  <button className="btn" onClick={() => downloadCSV(
                    result.records.filter((r) => r.flags.length).flatMap((r) =>
                      r.flags.map((f) => ({ order: r.orderNum, customer: r.company, severity: f.sev, code: f.code, detail: f.msg, expected: r.expected, paid: r.paid }))
                    ), "reconciliation-exceptions.csv")}>Export CSV</button>
                </div>
                <div style={{ marginTop: 14 }}>
                  {result.records.filter((r) => r.flags.length).length === 0 ? (
                    <div className="empty">No exceptions found — everything within tolerance.</div>
                  ) : (
                    result.records.filter((r) => r.flags.length).map((r) =>
                      r.flags.map((f, i) => (
                        <div key={r.key + i} className={"exc s" + f.sev}>
                          <div className="top">
                            <span className="code">{f.code}</span>
                            <span className="meta mono">{r.orderNum} · {r.company}</span>
                          </div>
                          <div className="msg">{f.msg}</div>
                        </div>
                      ))
                    )
                  )}
                </div>
              </div>
            )}

            {/* OBI CHECKS */}
            {tab === "obi" && (
              <div className="panel">
                <h2>OBI checks (Sch5-driven)</h2>
                <div className="banner info">
                  These are sensible default checks built from the Sch5 feed to support Tommy's process.
                  Send Tommy's actual OBI checklist and I'll wire these to match it exactly.
                </div>
                {!files.sch5 ? (
                  <div className="empty">Load a Sch5 export to run OBI checks.</div>
                ) : (
                  <ObiChecks records={result.records} />
                )}
              </div>
            )}
            {tab === "users" && isAdmin && (
              <div className="panel">
                <h2>Users — who can sign in</h2>
                <div className="banner info">
                  Add a person's email and a password here, then hand them those two things — that's how they sign in.
                  Only people in this list can see the data.
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 14 }}>
                  <input type="email" placeholder="Email" value={newUser.email}
                    onChange={(e) => setNewUser((u) => ({ ...u, email: e.target.value }))}
                    style={{ padding: "8px 10px", border: "1px solid #d3d0e6", borderRadius: 7, fontSize: 13, minWidth: 220 }} />
                  <input type="text" placeholder="Password (min 6 chars)" value={newUser.password}
                    onChange={(e) => setNewUser((u) => ({ ...u, password: e.target.value }))}
                    style={{ padding: "8px 10px", border: "1px solid #d3d0e6", borderRadius: 7, fontSize: 13, minWidth: 200 }} />
                  <button className="btn" style={{ background: "#5514b4", color: "#fff" }} onClick={addUser}>Add login</button>
                </div>
                {userMsg && (
                  <p className="sub" style={{ color: userMsg.err ? "#b3261e" : "#14804a", marginTop: 0 }}>{userMsg.text}</p>
                )}
                <table>
                  <thead><tr><th>Email</th><th>Role</th><th></th></tr></thead>
                  <tbody>
                    {users.map((u) => (
                      <tr key={u.email}>
                        <td className="mono">{u.email}</td>
                        <td>{u.is_admin ? <span className="chip matched">admin</span> : <span className="chip unmatched">viewer</span>}</td>
                        <td className="num">
                          {(u.email || "").toLowerCase() !== (session?.user?.email || "").toLowerCase() && (
                            <button className="btn" onClick={() => removeUser(u.email)}>Remove</button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <p className="note">Removing someone stops them seeing the data. Their password reset (if ever needed) is done from Supabase.</p>
              </div>
            )}
          </>
        )}

        <p className="note" style={{ textAlign: "center", marginTop: 20 }}>
          Access is limited to the emails in the Users list. Real commission data is stored in Supabase, protected by that list.
        </p>
      </div>
    </div>
  );
}

function LineDetail({ r }) {
  const block = (title, rows, cols) =>
    rows.length ? (
      <div style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#5514b4", marginBottom: 4 }}>{title}</div>
        <table>
          <thead><tr>{cols.map((c) => <th key={c[0]} className={c[2] ? "num" : ""}>{c[0]}</th>)}</tr></thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i}>{cols.map((c) => <td key={c[0]} className={c[2] ? "num mono" : ""}>{c[2] ? gbp(c[1](row)) : (c[1](row) ?? "—")}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    ) : null;
  return (
    <div className="detail">
      {block("NetSuite", r.nsL, [["Product", (x) => x.product], ["Status", (x) => x.status], ["Item paid", (x) => x.itemPaid], ["Expected GP", (x) => x.expected, true], ["NS Cobra Pmt", (x) => x.recordedCobra, true]])}
      {block("Cobra", r.cbL, [["Product", (x) => x.product], ["Month", (x) => x.month], ["Due", (x) => x.due, true], ["Paid", (x) => x.paid, true]])}
      {block("Sch5", r.s5L, [["Product", (x) => x.product], ["Status", (x) => x.status], ["Sub-status", (x) => x.subStatus], ["Commissionable", (x) => x.commissionable], ["SOV", (x) => x.sov, true]])}
      {!r.nsL.length && !r.cbL.length && !r.s5L.length && <div style={{ color: "#8a8aa3", fontSize: 12 }}>No lines.</div>}
    </div>
  );
}

function ForecastPanel({ forecast, hasSch5 }) {
  if (!forecast.length) return <div className="empty" style={{ padding: 20 }}>No unpaid expected-commission lines found (NetSuite "Item Paid" = No).</div>;
  const t = (v) => sum(forecast.filter((f) => f.verdict === v).map((r) => r.expected));
  const total = sum(forecast.map((r) => r.expected));
  const atRisk = t("at_risk_cancelled") + t("at_risk_noncomm");
  const unverified = t("unverified");
  const clean = t("expected");
  return (
    <>
      <div className="kpis" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div className="kpi"><div className="lab">Forecast (unpaid)</div><div className="val mono">{gbp(total)}</div><div className="foot">{forecast.length} lines awaiting payment</div></div>
        <div className="kpi"><div className="lab">Expected to pay</div><div className="val mono pos">{gbp(clean)}</div><div className="foot">live in Sch5, commissionable</div></div>
        <div className="kpi"><div className="lab">At risk</div><div className="val mono neg">{gbp(atRisk)}</div><div className="foot">cancelled / non-commissionable in Sch5</div></div>
        <div className="kpi"><div className="lab">Unverified</div><div className="val mono">{gbp(unverified)}</div><div className="foot">{hasSch5 ? "not found in Sch5" : "no Sch5 loaded"}</div></div>
      </div>
      <div style={{ overflowX: "auto", marginTop: 12 }}>
        <table>
          <thead><tr><th>Order #</th><th>Customer</th><th>Product</th><th className="num">Expected</th><th>Forecast verdict</th></tr></thead>
          <tbody>
            {[...forecast].sort((a, b) => (b.expected || 0) - (a.expected || 0)).map((f, i) => (
              <tr key={i}>
                <td className="mono">{f.orderNum || "—"}</td>
                <td>{f.company}</td>
                <td>{f.product}</td>
                <td className="num mono">{gbp(f.expected)}</td>
                <td>
                  {f.verdict === "expected" && <span className="chip matched">expected</span>}
                  {f.verdict === "at_risk_cancelled" && <span className="chip risk">cancelled in Sch5</span>}
                  {f.verdict === "at_risk_noncomm" && <span className="chip risk">non-commissionable</span>}
                  {f.verdict === "unverified" && <span className="chip unmatched">not in Sch5</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function BreakdownTable({ records }) {
  const byAgent = useMemo(() => {
    const m = new Map();
    for (const r of records) {
      const a = r.agent || "(unassigned)";
      if (!m.has(a)) m.set(a, { expected: 0, paid: 0, n: 0 });
      const g = m.get(a);
      g.expected += r.expected || 0;
      g.paid += r.paid || 0;
      g.n += 1;
    }
    return [...m.entries()].sort((a, b) => b[1].expected - a[1].expected);
  }, [records]);
  if (!byAgent.length) return <div className="empty" style={{ padding: 20 }}>No data.</div>;
  return (
    <table>
      <thead><tr><th>Admin agent</th><th className="num">Orders</th><th className="num">Expected</th><th className="num">Paid</th><th className="num">Variance</th></tr></thead>
      <tbody>
        {byAgent.map(([a, g]) => (
          <tr key={a}>
            <td>{a}</td>
            <td className="num mono">{g.n}</td>
            <td className="num mono">{gbp(g.expected)}</td>
            <td className="num mono">{gbp(g.paid)}</td>
            <td><Delta v={g.paid - g.expected} /></td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function ObiChecks({ records }) {
  const checks = [
    { name: "Cancelled in Sch5 but commission paid", test: (r) => r.sch5Cancelled && r.paid > 0, sev: 3 },
    { name: "Non-commissionable in Sch5 but expected in NetSuite", test: (r) => r.sch5NonComm && r.expected > 0, sev: 3 },
    { name: "In NetSuite but missing from Sch5 feed", test: (r) => r.inNS && !r.inSch5, sev: 1 },
    { name: "SOV differs between NetSuite and Sch5", test: (r) => r.flags.some((f) => f.code === "SOV_MISMATCH"), sev: 1 },
    { name: "Expected but unpaid (Item Paid = No)", test: (r) => r.anyUnpaid, sev: 2 },
  ];
  return (
    <div>
      {checks.map((c) => {
        const hits = records.filter(c.test);
        return (
          <div key={c.name} className={"exc s" + c.sev}>
            <div className="top">
              <span className="code">{c.name}</span>
              <span className="meta">{hits.length === 0 ? "✓ clear" : `${hits.length} order${hits.length > 1 ? "s" : ""}`}</span>
            </div>
            {hits.length > 0 && (
              <div className="meta mono" style={{ marginTop: 6 }}>
                {hits.slice(0, 12).map((r) => r.orderNum).join(" · ")}{hits.length > 12 ? " …" : ""}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
