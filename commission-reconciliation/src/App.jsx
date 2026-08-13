import { useState, useMemo, useCallback, useEffect, useRef, useTransition, Fragment } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

/* =========================================================================
   BTLB Cheques and Balances — shared, login-gated commission reconciliation
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
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhyZWtlYmdudWJoanF0cGxsYmN6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU0MDEyNDMsImV4cCI6MjEwMDk3NzI0M30.1MbG2AX63hFNvzZoZB56pjkOlW6Dq7s4U5mGaaJGi80";
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
const APP_VERSION = "2026-08-12-uk-dates";
const APP_NAME = "BTLB Cheques and Balances";
const PAGE_TITLE = "Cheques and Balances";
// BT roundel, transparent background — inlined so nothing extra needs deploying
const FAVICON = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAQjElEQVR42u1bfYxdxXX//c7Mvfe93fUutt8aSCJBIgqJKUpb40SJku5ajRJwDdgxzw4YQVVaIiUBBakVRGny/FRVIVIlKiitoIkiUkzrfTjmw5ikSuV1opISi7YpybZxCTVVFbB313i9X+/eOzOnf9z3vGt7d/3tGNGR7h/79t6Z833OzO8McdaGsg+DZjdWudm/3lL5h1/zcCsA/BYUy5V6GVR7AXSBiItPkQGYADlM5esghgD8i4F9+e9GPvlfs+frwy67G/0eoJ4Nqnk2GK8C0gB9+++Nvd/7iCpvIMLvBA3LjcSdAgNFQICHqodCAegRMgiCNBAYEIIADx+ySaEMKeQfSX1u6/CnftRmvAo1DSCcqSDOSABVDJgGNngAqF6ys9c4c5si3EGaD1qW4DWD0xSKkFMRQFKhlhAhBGwtr1AoAhQhEHRQVSWEkMgygWEMp02o+p8Q8ri3/onGm6uHj6XhvAmghprUsVkB6tp3fWdpnJbvAfnZiOWLPTI4bQao5iCNwFrLBISBwsNrhoA8VXCKqnlhM4wI7RBEiWF85F2nKQKcg6oHGVmWxCBGrtP7ofpolkw/9PQvPz0KKGvYzDrq4ZwLYLbEN1Z23gWYr1gpvceFKXi4nKASJo5YRoCHC80DBF8GdI8KXwlB96nXYRPseBo1cwBI8lLkxS2iYa8IL2fQawCuVOgKK6VlAoNcp6HwmUJpYCMrHXCh+b+A/9OtI6sfO11rOCUBFAFolasuffoqYfkvLZNPODThNc8BpWXJtjQ0DOI5qNlm8+xHT46teet0LO3Wnh2LXRR/BPTrobghYrm3ZWEOoBpGkUUJTtPvB53+QmN07c/bNJ5lAShrAOtguHnpc5ss40eEUU+uk04BNbCRZRl5mN4rlEdSg7/fvv9TB2ZcRmUQg7IMw7ocVQU2o3Ch2W61mcBmDKHBA+hlP/pDHTxi0usu/t6yxOMzQcPnIylf6XS6ZXFgxE4bNB9zmn3+qdEbttSgUgf0ZAIkT4b51qu6ofL8A5Yd9zltIiB3hDBml8m1uZ/Qr01G+d/seOPGqbY5AkAD1TOI1MoqGlLMU5j2mkuf7ejMoz9U8EsRSxdnOuEVQQWRtSzB6dTXB0Z+9/7ZdJ+BAIpJatjM/6h86PFYum9Lw9isBRN4Tb+dAfdvH7n+jXORp+erM9ZVXrg0Bh4wTG4vgmWhkER6TBYOP/GBkR/fMWNl89PCE5k9sBlDlQ8PJLJofRrGcm35umo4rOo/NzC6esus+OBnJfdzNdiHXUcEsWHpzk2k+StSup02HUFNpCdKw/i25SMvbSjcbX53kPkXGTR1MAxVVn47kUXrm+FQDihjdtoQ/N4gUx8bGF29pQ+7LKBsEXSumQcALdZS9mGXHRhdvSXI1MdC8HtjdlpA2QyH8kQWrR+qrPx2HQx9GDTzKVvmifZmN1a5myvPPpBIz6Y0jLWCTZf1mu3J7eTHGwfWvTITcXk+GD9WR7obq1wfdtnGgXWv5Hby416zPRG7LAGmYSxPpGfTzZVnH2i9Z05KAFUMmN1Y5dZXnrklZvd9aRhzisCIndZruqeZv/HJ7fvXH2i/h1/x2I1VrooBs33/+gPN/I1Pek33ROy0isA0jLmY3fetrzxzS/u9BWNAkT6oG3t3vg+B/6ZER0AeLMs2BPeqRuMfbby5YfhMy89zMdo0VS8Z6GW+6EURe4XTaSeIhIopiP7G1uHVr9WgnJ1ej7KAITQIQEPw3zSSdAXkQRCJBj+hmt50oTLfTpNVDJjGmxuGVdObNPgJQSQBeTCSdIXgvwlAWzwe7wJHJFjZcWcs3X2ZTjpCaFkSx/SuxsG1Q33YZS9E5mcLoQ+7bOPg2iHH9C7LkhDCTCddLN191cqOO9uCOsYFinx/R89gz5Sd+k+hWeaRu4SLoqYe3rJt5MbbTrXE/FWONq3rK88+UWL3plTHc4PIBvUHOlzH+x8f6x9rB1IpPhg0AHXKHr43lkUXe+TOIDK5To+ItffWUJPdGAx4m4zdGAw11ESsvTfX6RGDyHjkLpZFF0/Zw/cC1FZqBAvtE9Xu7y5G7PcKZImHcwm7oxSH7nlq+KaHF9Z+8f2JcvfJldynszvXOSu9Ns039z5zd4KLHkr1cG5gbUA4iMxc2Th83VuAQlqSUMb57TG7lgb1uWViszD+2qU90WOF9vv9Qvm4RcVCD6oYMO2i6QzmmeOZuwbZjX5fQ00u7Ykey8L4a5aJDerzmF1LNc5vB6B9GDRsbzi0Uv6JZelqH9IslkVxivF7nhpes6D2a6jJG7jBAC/PK56f40o99vtWug2nMs/8YwUuxQo/e77jrWDH3QkWPZSF8cxIEjtt/owj0x9soBosQDVLdlwLJld7TT1p4kzHD8Y5ngDAubRfnMfR/6x3xR2xTnwl13d5gHNWWsswGap4/jBF9grkhQm7r1F/g1M1qAyhwQY2+J9Wrv29BBN/stA885i/jzhhhvD8X2AED7Xpmm0FABhneCKLxjeTZonX1FuWrsYSXouDfMkCgBOsKzGBC808lrJJw8TTT46teatIjTxOAAcwyJbRLraSvNfrFAizwH6LENjfJGRjZ/7e+z/d+8wf1If5T1X8NAbgqVhszYnmmcv7PSwTZH6iMpuu2W5VxYB5cmzNWzcv3fF0LOXfT8N4biUxTUnXAXhJACWBT3jNAMIEeFDQmN9XjxrOaRaChsxpHpzmwWvuvebZzJNlXlOf66RPdcwJ5f2RJt/bsPi5X1+OhmudCeZO0zxomHaa507z3Gvm2nPOflq/507z1vtpDobsRAGWgkaABwjjNQOBTwBKu7H3hfeFgOUOKQQ2ysP0SKLyIkBtQE+Q+kgCAqgQFEAhsLDsMLP1lOsUFB6EINepPGF3Zyruz+uoX1foSUuxdEYBWTRjAQqnzeNWtCxJu35T+ChmJ3JMluYvjooDmdjvfDHl9IjQVpymgGL5xt7t77Ne/UorpbLTZh6xHOU6tWfLwdWHjw1UJ2GQKojpke/zYfxvCRVAVBkqhNwqsN0BTgmxmU4qiVUbe164fOvY9ftIfT3TiR96TR2UVkEFQ5nAtQCPskSn6cuqOkmQoLoMExZwrxbxZljnyi41qNQP8nB16Y49lsn1mU7mBc+60gJYITCgIpAGAH8MAIMYFACnIAAGw9h4zfc+Nbrmq7P/s37pju8asdtVvRIQRQiWpTi3zcsB7GuM3LgVwNajAu2SgXdDyvsAWiBoYW0CB3vLd0aPRotml8Jz/T7DC39MmOupCFLE2hUWyuXKgAK08KDglTM4qwAVyR2X7Sol2SIBfon/6bQ+GQ4/yOx0KrQlhdc2FiQaona66p+pNAVA2CvSMVfutUw7aqgJ0C9ofdPGKE6oIsErCt/iNQDK5ZbEZQEeCrVeUyh0HwD0oz/sPg2cRYn08ddXHeW81d7nbrAoJ06bgUXcoNM0wCT/Xaw1GNqgRg011FEPG/mdMFf9qJRQRz3UsBl1rGoJrb4gVW1eVHWfRwqF2gAPEpdZqPa2ApQEdSmCDhefbT5V7UsRXfWqauW5rwEUBZSqSxnkM4Gu5cwaYi4yWRj/UWNk9S8KlKl+jvcZLV58GA6iKSmJqgdUey3ALoUWgCR1in56Ysas6qeifQZkENjLLDvun13hFlnAtQ2fXrNDQfjHAHUIA3KuN0dHePHNCTUdUwJJFAEAuyyocbETIaHIkPRmZ4K1evgQdPwY99WoFc2DZdmkOvbNbcNrXyzOF87jFjvpzRAmM5JQqIKIzrr0BSKGcTzzRDFAKlQBGKdTPkLXF25e+uzn5junO59DoMxaBAJEjHQ4Pn39E6p+2of0VR+yV33IXnUh20cIDCIWQHiQgDwxTB6uLnv+mgY2+BpUzgu36XAMIi56E0gocgF0giAUAVTtgCl1zWB1pzS8ZRlKfWn56JqrFo9esvzAaOcHODp9lYf/bYXuF1ht1YbOsiTq9e5ZefqcjSO8mFIXVTsUodWboBMCcpg0UIQgtAmM9J5eFpiJenUwPIoVbjdWuQaqbtvIDT8MyP/asoMAvILikSmBDxc7zlX+vGQBI71CmyhCUfSRw6KK14u2FDrDBCQvPzOtkHfh0egevBDXsMtW0UgKsfDA0TtEJYCuFmChOCvtOnOPNi8kLzdMQNAJDFTxugV1iJDVUFXSQAOuAbDtdKMAgOwxfDafXb5X3zNQRlNu98hUi81TWwgTrcMSng9YTQOuoRgUvApAHbIAXg7wUEIUHoB+6PQqQRWvKQhcXa3seBhQUVAJRGjy44bxB5w2tbVrdAYRPZs/mQ3FnSvGZ3jRD2mL1wAPAC9bQ7PHh3SalHLR0ISVm5bs7K4f5OHWQaWerOkH5BDYd1t2fOHoY4sMTqeLvieoAlRFIMFvzL+LO2t6Zx0Mm5bs7E4RVjpNQUjkQzptJN8jW4evfw3EkGWCAJdHUq5kJnx0dnPCQltgLc4MgkKDgiHA+0zHs5lnIvOa5YQEhTqAKMmSKA0TDw+MrPlBDSoLgC3h6Pn1lEvmggdlZsJHIylXAlxumQDE0Nbhda9JoQ183zAGFF5goAHVk9G8KqOIZRFGsWVJLBMxTIxhEs9+LEtRxA6TSLcVRGnTv/Vn20ZvvKeNRc7JOQzb8xomppi/JNQgp5OZNKAqMIDCG8ZQ4PsA1QKADdjumN5HMsp1GiTW3tqz44+eHFtzaC43mGWyb7ow/e9Bcw84s8A+NAP4hqj556DT2xqja3+OY0DK4yo0F9Ig+b8C3hRuAxKEI6ZOxfwbYLi1Z8finFhb8MbIaQobsP0IMHK6x+JniuSeL4hsoWNx6cNgm5hvmFbrrkeqDPji3VfsTIqDCj3THM02MFJDTc4PwKrsx2C4+4qdCQO+6JEqALR4/EYBpLaAkTODxi5sgPSE0FgBFO4yjcPXHYSGRyw7WRw+TgfR+KvVS3b2FlZQk7cP+zXpx2CoXrKzVzT+qtPpUByndRLqHmkcvu5gUYG20OECQVF2uK4HszC+3yCyHrmPWK4E5x6sox760C9vH+33Sx31EJx7MGK54pF7g8hmYXx/h+t+sGjqKhAvmUFQGvL42KpDoH7ZskwAzHTCJ1y0qd1jU4Cbbw/TX1955paEizZlOuEB0LJMUL/8+NiqQ0VtQMWxG5AjXSJLnxuMpLMv04lWjw2ngjY/3Di4duhCbZE5iv4lTy8Xll5SakdAHmJ22TxM7m6M3tB/LP1HmXXRxwuKmDt9SI/02FBMF5k8U71koPfYFpMLjvlLBnrJ5BmKOdLj5EM6IWLuBMAWj5hTAHUwVDEgW4dX/8Ixu8uyLITA6bSzklwhrvv5m3q+ddGFJoQ28zf1fOsicd3PW0mucDrtCIFlWRyzu7YOr/5FFQPHoV1cMIVUnn2gJIvvS8NYDmirUTLdk5nJNdv3rz9wIaTHNg3rLt62LPadOwyTlblOOLRaZpvhra8/NXLj/fPROl+Bc6Qft1rZ8UQiPZua4dCR1nQXsr1qpm+e6RY9F83RJy502s3T1WXbr6EvP2UlvrLdwl+Si6I0jG1pjKy5baE+5nd8s/T/t8ufzMLtSd4eFyaaDwyMXP+ls3Rh4mh3eIdemTk+4r4DL03NvZd/x12bmzHXd/DFyfms4R11dfb4SP0OvTy9UJ5ujwv9+vz/AfsLWkAuNHarAAAAAElFTkSuQmCC";

// ---------- helpers ----------
const money = (v) => {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isNaN(v) ? null : v;
  let s = String(v).trim();
  if (!s) return null;
  // accounting negatives: (1,234.56)
  const bracketed = /^\(.*\)$/.test(s);
  // strip EVERYTHING that isn't a digit, dot or minus — covers £, $, commas,
  // non-breaking spaces, and mis-encoded symbols that arrive as "?" or "Â£"
  s = s.replace(/[^0-9.\-]/g, "");
  if (s === "" || s === "-" || s === ".") return null;
  // keep only the first decimal point and a leading minus
  const neg = s.startsWith("-") || bracketed;
  s = s.replace(/-/g, "");
  const parts = s.split(".");
  s = parts.length > 1 ? parts[0] + "." + parts.slice(1).join("") : s;
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return neg ? -n : n;
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
// Dates are UK format: day/month/year. Never let JS parse "12/8/2026" itself —
// it reads that as 8 December (US), which is how rows ended up in the wrong month.
const monthKey = (v) => {
  if (v == null || v === "") return null;
  const s = String(v).trim();
  // day/month/year, with optional time — checked BEFORE anything else
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    const day = Number(m[1]), mon = Number(m[2]);
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) return `${y}-${String(mon).padStart(2, "0")}`;
    // if the 2nd part can't be a month, the file is genuinely month/day — fall back
    if (day >= 1 && day <= 12) return `${y}-${String(day).padStart(2, "0")}`;
    return null;
  }
  m = s.match(/^(\d{4})-(\d{2})/); // ISO yyyy-mm
  if (m) return `${m[1]}-${m[2]}`;
  m = s.match(/^(\d{4})(\d{2})$/); // YYYYMM (e.g. 202601)
  if (m) return `${m[1]}-${m[2]}`;
  // NOTE: a real Date object here means something already parsed the text for us,
  // which is exactly how "12/08/2026" became December. Use it only as a last resort.
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}`;
  return null;
};
const periodLabel = (key) => {
  if (!key) return "No date";
  const [y, mo] = key.split("-");
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${names[Number(mo) - 1] || mo} ${y}`;
};

// resolve a value from a raw row by trying candidate header names (exact, then fuzzy)
// exact-name lookup only. Used for dates: a fuzzy match on "Date" would happily
// grab "Expected Order Close Date" and scatter rows into future months.
const pickExact = (row, candidates) => {
  const clean = (s) => String(s).toLowerCase().replace(/\s+/g, " ").trim();
  const keys = Object.keys(row);
  for (const c of candidates) {
    const hit = keys.find((k) => clean(k) === clean(c));
    if (hit !== undefined && row[hit] !== null && row[hit] !== "") return row[hit];
  }
  return undefined;
};

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
    // Excel writes CSVs as Windows-1252, so "£" is a byte UTF-8 can't decode.
    // Try strict UTF-8 first, then fall back — otherwise £ turns into "?" and numbers break.
    const buf = await file.arrayBuffer();
    let text;
    try {
      text = new TextDecoder("utf-8", { fatal: true }).decode(buf);
    } catch {
      text = new TextDecoder("windows-1252").decode(buf);
    }
    wb = XLSX.read(text, { type: "string", raw: true, cellDates: false, cellText: false });
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
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false, raw: true });
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

// ---------- only these columns are kept when saving to the shared store ----------
// (everything the reconciliation actually uses; the rest — e.g. Sch5's ~69 cols — is dropped
//  so the saved dataset stays small enough to store quickly)
const KEEP_COLS = {
  netsuite: ["Opp ID", "Order ref", "Company Name", "Order Status", "Contract Value",
    "Item: Name (Grouped)", "Item Name", "Product GP", "Cobra Payment", "Item Paid",
    "Overpayment", "Accelerator?", "Admin Agent", "Customer Le", "Netsuite Date",
    "Product Group 2", "Product Group2", "Product Group", "Netsuite Ref",
    // new-format NetSuite report
    "Sales Agent GP", "Date", "Item: Product Group 2", "Document Number", "Partner",
    "Partner Role: Name", "Team", "Sales Team Issue Dirty Order?", "Item: Schedule 5",
    "Customer CUG", "Order Status Last Changed Date", "Class (Item): Name", "Description"],
  cobra: ["Job Header", "Opty ID", "LE Code", "Customer Name", "Status", "Month",
    "Contract Value", "Commission Due", "Commission Paid", "Measure", "Plan Name",
    "Prod Code", "Closed Date", "Order Date", "Month", "FY"],
  sch5: ["MAIN ORDER NUM", "ADT REF", "OPPORTUNITY ID", "LE CODE", "CUSTOMER NAME",
    "ORDER STATUS", "ORDER SUB STATUS", "CANCEL DATE", "SOV", "COMMISSION FLAG",
    "PRODUCT SUB NAME1", "TRANSACTIONAL", "RECENT CANCELLATION", "CLOSED DATE",
    "ORDER DATE", "REPORTING MONTH", "Sch5 Order", "ORDER"],
};
const normHead = (s) => String(s).toLowerCase().replace(/\s+/g, " ").trim();
const slimFileForSave = (f, which) => {
  if (!f) return null;
  const keep = new Set((KEEP_COLS[which] || []).map(normHead));
  const rows = (f.rows || []).map((row) => {
    const out = {};
    for (const k of Object.keys(row)) if (keep.has(normHead(k))) out[k] = row[k];
    return out;
  });
  return { sheet: f.sheet, name: f.name, rows };
};

// ---------- canonical row extractors ----------
const nsRow = (r) => ({
  src: "netsuite",
  oppId: firstDefined(pick(r, ["Opp ID"])),
  // the new report has no BT order number, so fall back to Opp ID / Document Number
  // new report carries the BT order reference in the "Description" column
  orderNum: normOrder(firstDefined(
    pick(r, ["Order ref"]), pick(r, ["Description"]), pick(r, ["Opp ID"]), pick(r, ["Document Number"]))),
  docNo: firstDefined(pick(r, ["Document Number"])),
  partner: firstDefined(pick(r, ["Partner"])),
  partnerRole: firstDefined(pick(r, ["Partner Role: Name"])),
  team: firstDefined(pick(r, ["Team"])),
  dirty: firstDefined(pick(r, ["Sales Team Issue Dirty Order?"])),
  schedule5: firstDefined(pick(r, ["Item: Schedule 5"])),
  cug: firstDefined(pick(r, ["Customer CUG"])),
  company: firstDefined(pick(r, ["Company Name"])),
  status: firstDefined(pick(r, ["Order Status"])),
  // SOV is only counted on the Sales Closer's row, so the same order isn't counted per partner
  sov: (() => {
    const role = pick(r, ["Partner Role: Name"]);
    if (role != null && role !== "" && !/sales closer/i.test(String(role))) return null;
    return money(pick(r, ["Contract Value"]));
  })(),
  product: firstDefined(pick(r, ["Item: Name (Grouped)", "Item Name"])),
  expected: money(firstDefined(pick(r, ["Product GP"]), pick(r, ["Sales Agent GP"]))),
  recordedCobra: money(pick(r, ["Cobra Payment"])), // what NS records Cobra paid
  itemPaid: firstDefined(pick(r, ["Item Paid"])),
  // new report has no "Item Paid" flag — Order Status "Paid" carries it instead
  statusPaid: /(^|\b)paid(\b|$)/i.test(String(pick(r, ["Order Status"]) || "")),
  overpayment: firstDefined(pick(r, ["Overpayment"])),
  accelerator: firstDefined(pick(r, ["Accelerator?", "Accelerator"])),
  agent: firstDefined(pick(r, ["Admin Agent"])),
  le: normLE(pick(r, ["Customer Le", "Customer Ledger", "Customer Le "])),
  // month comes from the "Date" column (or "Netsuite Date" on the old report) and nothing else.
  // Values are UK format: 12/08/2026 is 12 AUGUST 2026.
  date: firstDefined(pickExact(r, ["Date"]), pickExact(r, ["Netsuite Date"])),
  dateCol: pickExact(r, ["Date"]) != null ? "Date" : (pickExact(r, ["Netsuite Date"]) != null ? "Netsuite Date" : null),
  productGroup2: firstDefined(pick(r, ["Item: Product Group 2"]), pick(r, ["Product Group 2", "Product Group2"])),
  netsuiteRef: firstDefined(pick(r, ["Netsuite Ref", "NetSuite Ref"]), pick(r, ["Document Number"])),
  raw: r,
});

// "BB - FTTP ACQ Hyperfast 5year" -> "BB";  "Data - BT Net Xsell" -> "Data"
const productGroupOf = (itemName) => {
  if (!itemName) return null;
  const s = String(itemName);
  const i = s.indexOf("-");
  if (i <= 0) return null; // no dash -> no group (keeps the filter list clean)
  const head = s.slice(0, i).trim();
  return head || null;
};

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
  date: firstDefined(pickExact(r, ["Order Date"]), pickExact(r, ["Closed Date"])),
  cobraMonth: firstDefined(pick(r, ["Month"])),
  cobraFy: firstDefined(pick(r, ["FY"])),
  raw: r,
});

// Cobra ships its own commission-run month + financial year (e.g. Month "Apr", FY "FY2026").
// FY2026 runs Apr 2026 -> Mar 2027, so months Jan-Mar belong to the following calendar year.
const MONTH_ABBR = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
// Month comes from Cobra's "Order Date" column, read as UK day/month/year.
// The Month + FY columns are only a fallback when Order Date is missing.
const cobraPeriod = (r) => {
  const fromDate = monthKey(r.date);
  if (fromDate) return fromDate;
  const mi = MONTH_ABBR.indexOf(String(r.cobraMonth || "").trim().slice(0, 3).toLowerCase());
  const fy = String(r.cobraFy || "").match(/(\d{4})/);
  if (mi >= 0 && fy) {
    const startYear = Number(fy[1]);
    const year = mi >= 3 ? startYear : startYear + 1; // Apr(3)..Dec stay, Jan-Mar roll over
    return `${year}-${String(mi + 1).padStart(2, "0")}`;
  }
  return null;
};

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
  // month comes from the Sch5 "Order Date" column, read as UK day/month/year
  date: pickExact(r, ["ORDER DATE"]),
  // "Order" flag: use a literal ORDER column if the export has one, else COMMISSION FLAG
  orderFlag: firstDefined(pickExact(r, ["SCH5 ORDER"]), pickExact(r, ["ORDER"]), pickExact(r, ["COMMISSION FLAG"])),
  orderFlagCol: pickExact(r, ["SCH5 ORDER"]) != null ? "SCH5 ORDER"
    : pickExact(r, ["ORDER"]) != null ? "ORDER"
      : pickExact(r, ["COMMISSION FLAG"]) != null ? "COMMISSION FLAG" : null,
  raw: r,
});

// In the new NetSuite report the same line-item appears once PER PARTNER, repeating
// Contract Value and Cobra Payment on every row. Those must be counted once per item,
// while Sales Agent GP is genuinely per-partner and IS summed across rows.
const sumOncePerDoc = (rows, valFn) => {
  const seen = new Set();
  let total = 0;
  for (const r of rows) {
    const v = valFn(r);
    if (v == null) continue;
    const k = String(r.docNo ?? r.orderNum ?? "");
    if (seen.has(k)) continue;
    seen.add(k);
    total += v;
  }
  return total;
};

// ---------- big-file storage ----------
// Row objects repeat every column name on every row, which blows past the request size
// limit on large exports. Store them column-wise (headers once, values as arrays) and
// write them in chunks so no single request is too big.
const CHUNK_ROWS = 1500;
// fixed match tolerance — the on-screen control was removed; put it back if it is ever needed
const TOLERANCE = { abs: 1, pct: 1 };

const packRows = (rows) => {
  const cols = [];
  const seen = new Set();
  for (const r of rows) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); cols.push(k); }
  return { c: cols, d: rows.map((r) => cols.map((k) => (r[k] === undefined ? null : r[k]))) };
};
const unpackRows = (cols, data) =>
  (data || []).map((vals) => {
    const o = {};
    cols.forEach((c, i) => { o[c] = vals[i]; });
    return o;
  });

// Run jobs N-at-a-time. Everything below used to run strictly one after another,
// so a 20-chunk file paid 20 round trips end to end.
async function pooled(items, limit, worker, onDone) {
  const results = new Array(items.length);
  let next = 0, done = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (true) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await worker(items[i], i);
      if (onDone) onDone(++done, items.length);
    }
  }));
  return results;
}

async function writeChunks(sb, prefix, data, onProgress) {
  await sb.from("recon_datasets").delete().like("id", `${prefix}-%`);
  const total = Math.ceil(data.length / CHUNK_ROWS) || 0;
  const jobs = Array.from({ length: total }, (_, i) => i);
  const errs = await pooled(jobs, 4, async (i) => {
    const slice = data.slice(i * CHUNK_ROWS, (i + 1) * CHUNK_ROWS);
    const { error } = await sb.from("recon_datasets")
      .upsert({ id: `${prefix}-${i}`, netsuite: { d: slice } }, { onConflict: "id" });
    return error ? `chunk ${i + 1}/${total}: ${error.message}` : null;
  }, onProgress);
  const bad = errs.find(Boolean);
  if (bad) throw new Error(bad);
  return total;
}

async function readChunks(sb, prefix, count, onProgress) {
  const ids = Array.from({ length: count }, (_, i) => `${prefix}-${i}`);
  const batches = [];
  for (let i = 0; i < ids.length; i += 10) batches.push(ids.slice(i, i + 10));
  const got = await pooled(batches, 6, async (batch) => {
    const { data } = await sb.from("recon_datasets").select("id,netsuite").in("id", batch);
    return data || [];
  }, onProgress);
  const parts = got.flat();
  parts.sort((a, b) => Number(String(a.id).split("-").pop()) - Number(String(b.id).split("-").pop()));
  return parts.flatMap((p) => p.netsuite?.d || []);
}

// meta -> full { sheet, name, rows }. Understands both the old and the chunked format.
async function hydrate(sb, meta, prefix, onProgress) {
  if (!meta) return null;
  if (Array.isArray(meta.rows)) return meta;                       // old single-row format
  if (!meta.c || !meta.n) return { ...meta, rows: [] };
  const data = await readChunks(sb, prefix, meta.n, onProgress);
  return { sheet: meta.sheet, name: meta.name, rows: unpackRows(meta.c, data) };
}

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
function reconcile(files, tol, period = "all", statusSettings = {}) {
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
    const recordedCobra = sumOncePerDoc(nsL, (r) => r.recordedCobra); // once per Document Number
    const due = sum(cbL.map((r) => r.due)); // Cobra Commission Due
    const paid = sum(cbL.map((r) => r.paid)); // Cobra Commission Paid
    const nsSov = sum(nsL.map((r) => r.sov)); // already restricted to Sales Closer rows
    const s5Sov = sum(s5L.map((r) => r.sov));
    const cobraSov = sum(cbL.map((r) => r.sov));

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
    const anyUnpaid = nsL.some((r) => (r.itemPaid != null ? !isYes(r.itemPaid) : !r.statusPaid));
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
      expected, recordedCobra, due, paid, nsSov, s5Sov, cobraSov,
      payDelta, recordDelta, dueVsPaid, sovDelta,
      sch5Cancelled, sch5NonComm, anyUnpaid, overpaymentFlag,
      period: monthKey(nsL[0]?.date) || (cbL[0] ? cobraPeriod(cbL[0]) : null) || monthKey(s5L[0]?.date) || null,
      product: firstDefined(nsL[0]?.product, cbL[0]?.product, s5L[0]?.product),
      productGroup: firstDefined(...nsL.map((x) => x.productGroup2)) || productGroupOf(firstDefined(...nsL.map((x) => x.product))),
      netsuiteRef: firstDefined(...nsL.map((x) => x.netsuiteRef)),
      nsStatus: firstDefined(...nsL.map((x) => x.status)),
      status: firstDefined(nsL[0]?.status, cbL[0]?.status, s5L[0]?.status),
      sch5Status: firstDefined(s5L[0]?.status),
      // Net figures use the SchThrive rule: drop excluded statuses and non-commissionable lines
      netSov: !statusCounts(statusSettings, firstDefined(...nsL.map((x) => x.status)))
        ? 0 : sum(nsL.filter((x) => !isNonCommissionable(x)).map((x) => x.sov)),
      netGp: !statusCounts(statusSettings, firstDefined(...nsL.map((x) => x.status)))
        ? 0 : sum(nsL.filter((x) => !isNonCommissionable(x)).map((x) => x.expected)),
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
  color:#1b1636; background:#f5f5fa; min-height:100%; padding:16px 14px; line-height:1.45; }
.mono { font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace; font-variant-numeric: tabular-nums; }
.num { font-variant-numeric: tabular-nums; text-align:center; white-space:nowrap; }
.wrap { max-width:100%; margin:0 auto; }
.shell { display:flex; flex-direction:column; min-height:100vh; margin:-16px -14px; }
.sidebar { background:#0b2050; color:#c9d4ee; padding:0 14px; display:flex; flex-direction:row;
  align-items:center; gap:2px; flex-wrap:wrap; }
.sidebar .brand { color:#fff; font-weight:800; font-size:13.5px; padding:12px 16px 12px 4px;
  letter-spacing:-.01em; margin-right:8px; white-space:nowrap; }
.navi { display:inline-flex; align-items:center; gap:7px; padding:14px 13px; cursor:pointer;
  font-size:13px; font-weight:600; color:#c9d4ee; border:none; background:none; white-space:nowrap;
  border-bottom:3px solid transparent; }
.navi:hover { color:#fff; background:#14306e; }
.navi.on { color:#fff; border-bottom-color:#4f9bff; background:#14306e; }
.navi .ic { font-size:13px; }
.main { flex:1 1 auto; min-width:0; padding:16px 18px 40px; background:#f5f5fa; }
.ptitle { font-size:23px; font-weight:800; letter-spacing:-.02em; margin:0 0 14px; color:#0b2050; }
.kpis5 { display:grid; grid-template-columns:repeat(5,1fr); gap:12px; }
@media (max-width:1100px){ .kpis5 { grid-template-columns:repeat(2,1fr); } }
.kpic { background:#fff; border:1px solid #e7e6f0; border-top-width:3px; border-radius:12px; padding:14px; }
.kpic .lab { font-size:12px; color:#5b5676; font-weight:600; }
.kpic .val { font-size:21px; font-weight:800; margin-top:3px; letter-spacing:-.02em; }
.kpic .sub2 { font-size:12px; font-weight:700; margin-top:1px; }
.two { display:grid; grid-template-columns:1.55fr 1fr; gap:14px; align-items:start; }
@media (max-width:1100px){ .two { grid-template-columns:1fr; } }
.legend { display:flex; gap:14px; font-size:12px; color:#5b5676; margin-bottom:6px; flex-wrap:wrap; }
.legend i { width:10px; height:10px; border-radius:3px; display:inline-block; margin-right:5px; }
.step { display:flex; align-items:center; gap:8px; }
.spin { width:14px; height:14px; border-radius:50%; border:2px solid #c9c6da; border-top-color:#5514b4;
  display:inline-block; animation:sp .7s linear infinite; }
@keyframes sp { to { transform:rotate(360deg); } }
.fcards { display:flex; gap:8px; flex-wrap:wrap; margin:10px 0 12px; }
.fcard { background:#fff; border:1px solid #e7e6f0; border-top-width:3px; border-radius:10px;
  padding:8px 12px; cursor:pointer; text-align:left; min-width:104px; }
.fcard:hover { background:#faf9ff; }
.fcard.on { box-shadow:0 0 0 2px #5514b4 inset; }
.fcard .n { display:block; font-size:17px; font-weight:800; letter-spacing:-.02em; }
.fcard .l { display:block; font-size:11px; color:#5b5676; margin-top:1px; }
.step .n { width:24px; height:24px; border-radius:50%; background:#1e64d6; color:#fff; font-size:12px;
  font-weight:700; display:flex; align-items:center; justify-content:center; }
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
th { text-align:center; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#8a8aa3;
  border-bottom:1px solid #eceaf4; padding:8px 10px; position:sticky; top:0; background:#fff; }
th.num { text-align:center; }
td { padding:9px 10px; border-bottom:1px solid #f1f0f8; vertical-align:middle; text-align:center; }
td.left, th.left { text-align:left; }
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
.wip table { font-size:12.5px; }
.wip th { background:#f7f6fc; border-bottom:1px solid #e4e1f0; padding:7px 8px; color:#5b5676; font-size:10.5px; }
.wip td { padding:6px 8px; border-bottom:1px solid #f4f3fa; }
.wip tbody tr:hover { background:#faf9ff; }
.wip .lbl { position:sticky; left:0; background:#fff; z-index:2; box-shadow:1px 0 0 #eceaf4; min-width:250px; }
.wip tbody tr:hover .lbl { background:#faf9ff; }
.wip .grp td { border-top:2px solid #e4e1f0; }
.wip .tot { background:#faf9ff; font-weight:700; }
.pay table { font-size:12px; }
.pay th { padding:6px 7px; font-size:10px; background:#f7f6fc; }
.pay td { padding:4px 7px; line-height:1.25; }
.pay .agent { position:sticky; left:0; background:#fff; z-index:2; box-shadow:1px 0 0 #eceaf4; min-width:150px; }
.pay tbody tr:hover .agent { background:#faf9ff; }
.pay tbody tr:hover { background:#faf9ff; }
.pay .moves td { background:#f4f2fd; font-weight:600; }
.pay .moves .agent { background:#f4f2fd; }
.wip input { width:78px; padding:3px 5px; border:1px solid #e2e0ee; border-radius:5px; font-size:12px;
  text-align:center; background:#fbfaff; font-variant-numeric:tabular-nums; }
.wip input:focus { outline:2px solid #c4bce6; background:#fff; }
.side { width:50%; flex:0 0 50%; max-height:82vh; overflow:auto; position:sticky; top:14px; }
@media (max-width:1100px){ .side { width:100%; flex:1 1 auto; position:static; } }
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
  const [tab, setTab] = useState("overview");
  const [tabPending, startTabChange] = useTransition();

  useEffect(() => {
    document.title = PAGE_TITLE;
    let link = document.querySelector("link[rel~='icon']");
    if (!link) { link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
    link.type = "image/png";
    link.href = FAVICON;
  }, []);
  const goTab = useCallback((k) => startTabChange(() => setTab(k)), []);
  const [expanded, setExpanded] = useState(new Set());
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authMsg, setAuthMsg] = useState(null);
  const [sharedMeta, setSharedMeta] = useState(null);
  const [saving, setSaving] = useState(false);
  const [saveProgress, setSaveProgress] = useState(null);
  const [loadProgress, setLoadProgress] = useState(null);
  const [period, setPeriod] = useState("all");
  const [isAdmin, setIsAdmin] = useState(false);
  const [users, setUsers] = useState([]);
  const [newUser, setNewUser] = useState({ email: "", password: "" });
  const [userMsg, setUserMsg] = useState(null);
  // shared settings: risk levels per order status, agent payplans, paid marks, WIP manual rows.
  // Stored as a second row (id='settings') in the SAME table — no database change needed.
  const [settings, setSettings] = useState({ risk: {}, payplans: {}, paidMarks: {}, wip: {} });
  const [settingsSaving, setSettingsSaving] = useState(false);
  // agent list from the existing SchThrive `staff` table (read-only; no schema change)
  const [staff, setStaff] = useState({ rows: [], status: "idle", planCol: null });
  // tracks which slots the user has replaced in this session, so a late DB read can't overwrite them
  const uploadedHere = useRef({});
  const settingsRef = useRef({});

  const EXCLUDED_MANAGER = "tracy webber";

  // We can't list tables through the API, so try likely names and report what exists.
  const PLAN_TABLE_CANDIDATES = ["payplans", "payplan", "pay_plans", "pay_plan", "agent_payplans",
    "staff_payplans", "targets", "agent_targets", "commission_plans", "staff"];
  const [probe, setProbe] = useState({ status: "idle", found: [] });

  const probeTables = useCallback(async () => {
    if (!supabase || !session) return;
    setProbe({ status: "loading", found: [] });
    const found = [];
    for (const t of PLAN_TABLE_CANDIDATES) {
      const { data, error } = await supabase.from(t).select("*").limit(1);
      if (!error && data) found.push({ table: t, cols: data.length ? Object.keys(data[0]) : [], rows: data.length });
    }
    setProbe({ status: "done", found });
  }, [session]);

  const loadStaff = useCallback(async () => {
    if (!supabase || !session) return;
    setStaff((s) => ({ ...s, status: "loading" }));
    const { data, error } = await supabase.from("staff").select("*");
    if (error || !data) { setStaff({ rows: [], status: "error", planCol: null, msg: error?.message }); return; }
    const cols = data.length ? Object.keys(data[0]) : [];
    const planCol = cols.find((c) => /pay ?_?plan|target|quota/i.test(c)) || null;
    const mgrCols = cols.filter((c) => /manager|team/i.test(c));
    // the Order Delivery team (Tracy Webber) never owns sales, so drop them
    const rows = data.filter((r) => !mgrCols.some((c) => String(r[c] || "").toLowerCase().includes(EXCLUDED_MANAGER)));
    // payplans may live in their own table; fall back to a column on staff
    let plans = {};
    const cfg = (settingsRef.current || {}).payplanSource;
    let plansByMonth = {}, planNames = {};
    if (cfg && cfg.table && cfg.nameCol && cfg.valueCol) {
      const { data: cd } = await supabase.from(cfg.table).select("*");
      if (cd) {
        for (const r of cd) {
          const nm = String(r[cfg.nameCol] || "").trim();
          if (!nm) continue;
          const amt = Number(r[cfg.valueCol]) || 0;
          const mk = cfg.monthCol ? planMonthKey(r[cfg.monthCol]) : null;
          if (mk) {
            plansByMonth[nm] = plansByMonth[nm] || {};
            plansByMonth[nm][mk] = amt;
          } else {
            plans[nm] = amt;              // no month column -> a standing plan
          }
          if (cfg.planNameCol && r[cfg.planNameCol]) {
            planNames[nm] = planNames[nm] || {};
            planNames[nm][mk || "all"] = String(r[cfg.planNameCol]);
          }
        }
        setStaff((st) => ({ ...st, plans, plansByMonth, planNames, planSource: `${cfg.table}.${cfg.valueCol}` }));
      }
    }
    const nameOf = (r) => {
      const nc = cols.find((c) => /^(name|full_?name|agent)$/i.test(c)) || cols.find((c) => /name/i.test(c));
      return nc ? String(r[nc] || "").trim() : "";
    };
    const { data: pp } = await supabase.from("payplans").select("*");
    if (pp && pp.length) {
      const pcols = Object.keys(pp[0]);
      const pName = pcols.find((c) => /name|agent|staff/i.test(c));
      const pVal = pcols.find((c) => /pay ?_?plan|amount|target|value|quota/i.test(c));
      if (pName && pVal) for (const r of pp) plans[String(r[pName] || "").trim()] = Number(r[pVal]) || 0;
    } else if (planCol) {
      for (const r of rows) plans[nameOf(r)] = Number(r[planCol]) || 0;
    }
    setStaff({ rows, status: "ok", planCol, plans, excluded: data.length - rows.length });
  }, [session]);

  useEffect(() => { loadStaff(); }, [loadStaff]);

  // status_config drives how each order status is treated (its "tone" column)
  const [statusConfig, setStatusConfig] = useState({ rows: [], status: "idle", cols: {} });
  useEffect(() => {
    if (!supabase || !session) return;
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase.from("status_config").select("*");
      if (cancelled) return;
      if (error || !data || !data.length) { setStatusConfig({ rows: [], status: "error", cols: {} }); return; }
      const cols = Object.keys(data[0]);
      const nameCol = cols.find((c) => /^status$/i.test(c)) || cols.find((c) => /status|name|label/i.test(c));
      const toneCol = cols.find((c) => /^tone$/i.test(c)) || cols.find((c) => /tone|type|category/i.test(c));
      setStatusConfig({ rows: data, status: "ok", cols: { nameCol, toneCol } });
    })();
    return () => { cancelled = true; };
  }, [session]);

  // order statuses as used by SchThrive WebOS (same database, read-only)
  const [webosStatuses, setWebosStatuses] = useState({ list: [], status: "idle" });
  useEffect(() => {
    if (!supabase || !session) return;
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase.from("orders").select("status").limit(5000);
      if (cancelled) return;
      if (error || !data) { setWebosStatuses({ list: [], status: "error" }); return; }
      const counts = new Map();
      for (const r of data) {
        const v = String(r.status || "").trim();
        counts.set(v, (counts.get(v) || 0) + 1);
      }
      setWebosStatuses({
        list: [...counts.entries()].map(([name, n]) => ({ name, n })).sort((a, b) => b.n - a.n),
        status: "ok",
      });
    })();
    return () => { cancelled = true; };
  }, [session]);

  const staffManagers = useMemo(() => {
    if (!staff.rows.length) return {};
    const cols = Object.keys(staff.rows[0]);
    const nameCol = cols.find((c) => /^(name|full_?name|agent)$/i.test(c)) || cols.find((c) => /name/i.test(c) && !/manager/i.test(c));
    const mgrCol = cols.find((c) => /^manager_?name$/i.test(c)) || cols.find((c) => /manager/i.test(c) && /name/i.test(c))
      || cols.find((c) => /^team$/i.test(c));
    if (!nameCol || !mgrCol) return {};
    const out = {};
    for (const r of staff.rows) {
      const n = String(r[nameCol] || "").trim();
      if (n) out[n] = String(r[mgrCol] || "").trim();
    }
    return out;
  }, [staff]);

  const staffNames = useMemo(() => {
    const nameCol = staff.rows.length
      ? Object.keys(staff.rows[0]).find((c) => /^(name|full_?name|agent)$/i.test(c)) ||
        Object.keys(staff.rows[0]).find((c) => /name/i.test(c))
      : null;
    if (!nameCol) return [];
    return staff.rows.map((r) => String(r[nameCol] || "").trim()).filter(Boolean).sort();
  }, [staff]);

  useEffect(() => {
    if (!supabase || !session) return;
    let cancelled = false;
    (async () => {
      const { data } = await supabase.from("recon_datasets").select("cobra").eq("id", "settings").maybeSingle();
      if (!cancelled && data && data.cobra) {
        setSettings((s) => ({ ...s, ...data.cobra }));
      }
    })();
    return () => { cancelled = true; };
  }, [session]);

  useEffect(() => { settingsRef.current = settings; }, [settings]);

  const settingsTimer = useRef(null);
  const saveSettings = useCallback((next) => {
    setSettings(next);
    settingsRef.current = next;
    if (!supabase || !session) return;
    setSettingsSaving(true);
    if (settingsTimer.current) clearTimeout(settingsTimer.current);
    settingsTimer.current = setTimeout(() => {
      supabase.from("recon_datasets").upsert(
        { id: "settings", cobra: settingsRef.current, uploaded_by: session.user?.email || null,
          uploaded_at: new Date().toISOString() },
        { onConflict: "id" }
      ).then(() => setSettingsSaving(false), () => setSettingsSaving(false));
    }, 1200);
  }, [session]);

  const snapshot = settings.snapshot || null;
  const saveSnapshot = useCallback((snap) => saveSettings({ ...settings, snapshot: snap }), [settings, saveSettings]);

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
      const wanted = ["cobra", "netsuite", "sch5"].filter((k) => data[k]);
      const totalParts = wanted.reduce((n, k) => n + Math.ceil((data[k].n || 0) / 10), 0) || 1;
      let doneParts = 0;
      setLoadProgress({ done: 0, total: totalParts });
      const loaded = await Promise.all(wanted.map((k) =>
        hydrate(supabase, data[k], `chunk-current-${k}`, () => {
          doneParts += 1;
          setLoadProgress({ done: doneParts, total: totalParts });
        })
      ));
      setLoadProgress(null);
      if (cancelled) return;
      const next = {};
      wanted.forEach((k, i) => { if (loaded[i]) next[k] = loaded[i]; });
      // never clobber a file the user has just uploaded while this was loading
      if (Object.keys(next).length) setFiles((f) => {
        const merged = { ...f };
        for (const k of Object.keys(next)) if (!uploadedHere.current[k]) merged[k] = next[k];
        return merged;
      });
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
    setErrors((e) => ({ ...e, save: null }));
    setSaveProgress(null);
    try {
      const payload = {
        id: "current",
        uploaded_by: session.user?.email || null,
        uploaded_at: new Date().toISOString(),
      };
      for (const which of ["cobra", "netsuite", "sch5"]) {
        const slim = slimFileForSave(nextFiles[which], which);
        if (!slim) { payload[which] = null; continue; }
        const packed = packRows(slim.rows);
        const n = await writeChunks(supabase, `chunk-current-${which}`, packed.d,
          (done, total) => setSaveProgress(`${which} — part ${done} of ${total}`));
        payload[which] = { sheet: slim.sheet, name: slim.name, c: packed.c, n, count: slim.rows.length };
      }
      setSaveProgress("finishing");
      const { error } = await supabase.from("recon_datasets").upsert(payload, { onConflict: "id" });
      if (error) throw new Error(error.message);
      setSharedMeta({ by: payload.uploaded_by, at: payload.uploaded_at });
      setErrors((e) => ({ ...e, save: null }));
    } catch (err) {
      setErrors((e) => ({ ...e, save: "Couldn't save: " + (err.message || err) }));
    }
    setSaveProgress(null);
    setSaving(false);
  }, [session]);

  const onFile = useCallback(async (which, file) => {
    if (!file) return;
    try {
      const parsed = await parseWorkbook(file);
      const entry = { sheet: parsed.sheet, headers: parsed.headers, rows: parsed.rows, name: file.name };
      uploadedHere.current[which] = true;
      setFiles((f) => ({ ...f, [which]: entry }));
      setErrors((e) => ({ ...e, [which]: null }));
    } catch (err) {
      setErrors((e) => ({ ...e, [which]: "Couldn't read that file — is it a valid CSV or Excel file?" }));
    }
  }, [saveShared]);

  // status_config (tone column) is the base; anything set by hand in Settings overrides it
  const effectiveStatusRules = useMemo(() => {
    const out = {};
    const { nameCol, toneCol } = statusConfig.cols || {};
    if (nameCol && toneCol) {
      for (const r of statusConfig.rows) {
        const nm = String(r[nameCol] || "").trim();
        const t = toneToTreatment(r[toneCol]);
        if (nm && t) out[nm] = t;
      }
    }
    return { ...out, ...(settings.risk || {}) };
  }, [statusConfig, settings.risk]);

  const result = useMemo(() => reconcile(files, TOLERANCE, period, effectiveStatusRules),
    [files, period, effectiveStatusRules]);

  // 30k+ rows mapped through the extractors is the single most expensive thing here,
  // so do it once and share it rather than once per tab.
  const sources = useMemo(() => ({
    ns: (files.netsuite?.rows || []).map(nsRow).map((r) => ({ ...r, period: monthKey(r.date), rawDate: r.date })),
    cb: (files.cobra?.rows || []).map(cobraRow).map((r) => ({ ...r, period: cobraPeriod(r) })),
    s5: (files.sch5?.rows || []).map(sch5Row).map((r) => ({ ...r, period: monthKey(r.date) })),
  }), [files]);
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
            <h1 style={{ margin: 0, fontSize: 20, display: "flex", alignItems: "center", gap: 9 }}>
              <img src={FAVICON} alt="" width="26" height="26" />
              BTLB <span style={{ color: "#5514b4" }}>Cheques and Balances</span>
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
            BTLB <span style={{ color: "#5514b4" }}>Cheques and Balances</span>
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

  const NAV = [
    ["overview", "Overview", "◴"],
    ["cross", "Cross-Reference", "⇄"],
    ["wip", "Cobra Dashboard", "▦"],
    ["agents", "Payroll", "◫"],
    ["risk", "Top Risk Orders", "▲"],
    ["settings", "Settings", "⚙"],
  ];


  return (
    <div className="recon">
      <style>{STYLES}</style>
      <div className="shell">
        <nav className="sidebar">
          <div className="brand">
            <img src={FAVICON} alt="" width="20" height="20" style={{ verticalAlign: "-4px", marginRight: 8 }} />
            {APP_NAME}
          </div>
          {NAV.map(([k, l, ic]) => (
            <button key={k} className={"navi " + (tab === k ? "on" : "")} onClick={() => goTab(k)}>
              <span className="ic">{ic}</span><span className="lbl2">{l}</span>
            </button>
          ))}
          <div style={{ marginLeft: "auto", fontSize: 11, color: "#7f92c4", paddingRight: 4 }}>
            {snapshot?.at ? `Snapshot ${new Date(snapshot.at).toLocaleDateString("en-GB")}` : "No snapshot"}
          </div>
        </nav>
        <div className="main">
        <div className="head">
          <h1 className="ptitle" style={{ margin: 0 }}>
            {(NAV.find(([kk]) => kk === tab) || [, "Overview"])[1]}
            {tabPending && <span className="sub" style={{ marginLeft: 10, fontWeight: 400 }}>loading…</span>}
          </h1>
          <span className="sub" style={{ margin: 0 }}><span className="mono" style={{ opacity: .6 }}>({APP_VERSION})</span></span>
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
        {loadProgress && (
          <div className="banner info" style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span className="spin" />
            <span>
              Loading the shared data — {loadProgress.done} of {loadProgress.total} parts.
              Everything will fill in as it arrives.
            </span>
          </div>
        )}
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

        {/* controls: month + tolerance (only once there's data) */}
        {anyLoaded && (
          <div className="panel">
            <div className="settings">
              <span>Month to check:</span>
              <select value={period} onChange={(e) => setPeriod(e.target.value)}
                style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
                <option value="all">All months</option>
                {result.periods.map((p) => (
                  <option key={p} value={p}>{periodLabel(p)}</option>
                ))}
              </select>
            </div>
          </div>
        )}

        {anyLoaded && allLoaded && result.counts.overlap === 0 && (
          <div className="banner">
            <strong>No shared order numbers across your files.</strong> These look like disjoint sample
            exports — every order appears in only one file, so nothing can cross-match yet. Load Cobra,
            NetSuite and Sch5 exports that cover the <em>same orders / period</em> to see live reconciliation.
          </div>
        )}


        {tab === "settings" && (
          <Settings
            files={files} errors={errors} onFile={onFile} supabase={supabase} session={session}
            saving={saving} sharedMeta={sharedMeta} saveShared={saveShared} anyLoaded={anyLoaded} saveProgress={saveProgress}
            isAdmin={isAdmin} users={users} newUser={newUser} setNewUser={setNewUser}
            addUser={addUser} removeUser={removeUser} userMsg={userMsg}
            settings={settings} saveSettings={saveSettings} settingsSaving={settingsSaving}
            records={result.records} staffNames={staffNames} staff={staff} loadStaff={loadStaff} managers={staffManagers}
            webosStatuses={webosStatuses} probe={probe} probeTables={probeTables} statusConfig={statusConfig} statusRules={effectiveStatusRules}
          />
        )}

        {tab !== "settings" && tab !== "overview" && !anyLoaded && (
          <div className="empty panel">No data loaded yet. Open <strong>Settings → Raw Data</strong> to upload the three exports.</div>
        )}

        {tab !== "settings" && (anyLoaded || tab === "overview") && (
          <>
            {tab === "overview" && (
              <Overview sources={sources} records={result.records} files={files} settings={settings}
                snapshot={snapshot} saveSnapshot={saveSnapshot} setTab={goTab}
                supabase={supabase} session={session} statusRules={effectiveStatusRules} />
            )}

            {tab === "risk" && (
              <TopRiskOrders records={result.records} settings={settings}
                statusRules={effectiveStatusRules} saveSettings={saveSettings} />
            )}

            {tab === "cross" && <CrossReference records={result.records} settings={{ ...settings, risk: effectiveStatusRules }} />}

            {tab === "wip" && <WipTracker sources={sources} files={files} settings={{ ...settings, risk: effectiveStatusRules }} saveSettings={saveSettings}
              settingsSaving={settingsSaving} supabase={supabase} session={session} />}

            {tab === "agents" && <AgentPayments sources={sources} files={files} settings={settings} saveSettings={saveSettings} staffNames={staffNames}
              dbPlans={staff.plans || {}} dbPlansByMonth={staff.plansByMonth || {}} dbPlanNames={staff.planNames || {}}
              managers={staffManagers} statusRules={effectiveStatusRules} supabase={supabase} session={session} />}

            {/* RECONCILIATION */}
          </>
        )}

        <p className="note" style={{ textAlign: "center", marginTop: 20 }}>
          Access is limited to the emails in the Users list. Real commission data is stored in Supabase, protected by that list.
        </p>
        </div>
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

// ---------- Raw Data tab: uploads + collapsible source previews ----------
function fmtCell(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return v.toLocaleDateString("en-GB");
  const s = String(v);
  return s.length > 60 ? s.slice(0, 60) + "…" : s;
}

function RawTable({ f }) {
  const rows = f.rows || [];
  const cols = f.headers && f.headers.length ? f.headers : (rows[0] ? Object.keys(rows[0]) : []);
  const shown = rows.slice(0, 200);
  if (!rows.length) return <p className="note" style={{ marginTop: 6 }}>No rows.</p>;
  return (
    <div style={{ overflowX: "auto", marginTop: 8 }}>
      <table>
        <thead><tr>{cols.map((c) => <th key={c}>{c}</th>)}</tr></thead>
        <tbody>
          {shown.map((r, i) => (
            <tr key={i}>{cols.map((c) => <td key={c} className="mono" style={{ fontSize: 12 }}>{fmtCell(r[c])}</td>)}</tr>
          ))}
        </tbody>
      </table>
      {rows.length > 200 && <p className="note">Showing the first 200 of {rows.length} rows.</p>}
    </div>
  );
}

function RawData({ files, errors, onFile, supabase, session, saving, sharedMeta, saveShared, anyLoaded, saveProgress }) {
  const [open, setOpen] = useState({});
  const sources = [
    ["cobra", "Cobra", "What BT actually paid"],
    ["netsuite", "NetSuite", "What we expect & recorded"],
    ["sch5", "Sch5", "BT source feed — status, cancels, SOV"],
  ];
  return (
    <>
      <div className="panel">
        <h2>Upload exports</h2>
        <div className="uploads">
          {sources.map(([key, role, desc]) => {
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
                  <div className="status">✓ {f.rows.length} rows<br /><span style={{ color: "#8a8aa3" }}>{f.name}</span></div>
                )}
                {err && <div className="status err">{err}</div>}
              </div>
            );
          })}
        </div>
        {supabase && session && anyLoaded && (
          <div style={{ marginTop: 14, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
            <button onClick={() => saveShared(files)} disabled={saving}
              style={{ background: saving ? "#b6a9e0" : "#5514b4", color: "#fff", border: "none", padding: "9px 16px", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: saving ? "default" : "pointer" }}>
              {saving ? "Saving…" : "Save & share"}
            </button>
            <span className="sub" style={{ margin: 0 }}>
              {saving ? (saveProgress ? `Storing the data — ${saveProgress}…` : "Storing the data…")
                : sharedMeta?.at ? `Saved. Shared with everyone signed in (last: ${new Date(sharedMeta.at).toLocaleString("en-GB")}).`
                  : "Load your files, then click Save & share once."}
            </span>
            {errors.save && <span className="sub" style={{ margin: 0, color: "#b3261e" }}>{errors.save}</span>}
          </div>
        )}
      </div>

      <div className="panel">
        <h2>What's held in here</h2>
        <p className="note" style={{ marginTop: 0 }}>The data behind each source. Click a row to open it (collapsed by default).</p>
        {sources.map(([key, label]) => {
          const f = files[key];
          const isOpen = !!open[key];
          return (
            <div key={key} style={{ borderTop: "1px solid #f1f0f8", padding: "10px 0" }}>
              <div style={{ cursor: f ? "pointer" : "default", fontWeight: 600, display: "flex", gap: 8, color: f ? "#1b1636" : "#a7a3bf" }}
                onClick={() => f && setOpen((o) => ({ ...o, [key]: !o[key] }))}>
                <span style={{ color: "#5514b4" }}>{f ? (isOpen ? "▾" : "▸") : "•"}</span>
                {label} {f ? <span className="sub" style={{ margin: 0 }}>— {f.rows.length} rows</span> : <span className="sub" style={{ margin: 0 }}>— not loaded</span>}
              </div>
              {isOpen && f && <RawTable f={f} />}
            </div>
          );
        })}
      </div>
    </>
  );
}

// ---------- Cross-Reference tab: order ref across all three sheets ----------
function CrossReference({ records, settings }) {
  const riskOf = (r) => (settings?.risk || {})[String(r.nsStatus || "").trim()] || "";
  const [group, setGroup] = useState("all");
  const [sheets, setSheets] = useState("all");
  const [paidView, setPaidView] = useState("show");
  const [sort, setSort] = useState({ col: "diff", dir: "desc" });
  const [showCount, setShowCount] = useState(200);

  const spread = (vals) => { const v = vals.filter((x) => x != null); if (!v.length) return 0; return Math.max(...v) - Math.min(...v); };
  const isPaid = (r) => /paid/i.test(String(r.nsStatus || ""));

  // GP delta: Cobra paid MINUS what we expected.
  // positive = BT paid us more than expected (good, green); negative = underpaid (bad, red)
  const enriched = useMemo(() => records.map((r) => ({
    ...r,
    sovSpread: spread([r.inNS ? r.nsSov : null, r.inSch5 ? r.s5Sov : null, r.inCobra ? r.cobraSov : null]),
    gpDiff: (r.inCobra ? r.paid : 0) - (r.inNS ? r.expected : 0),
    sheetCode: `${r.inNS ? "N" : "-"}${r.inCobra ? "C" : "-"}${r.inSch5 ? "S" : "-"}`,
    onAll: r.inNS && r.inCobra && r.inSch5,
    paid: r.paid,
    paidFlag: isPaid(r),
    risk: riskOf(r),
  })), [records, settings]);

  const groups = useMemo(
    () => [...new Set(enriched.map((r) => r.productGroup).filter(Boolean))].sort(), [enriched]);

  const bySheets = (r) => {
    if (sheets === "all") return true;
    if (sheets === "all3") return r.onAll;
    if (sheets === "missing") return !r.onAll;
    if (sheets === "nsOnly") return r.inNS && !r.inCobra && !r.inSch5;
    if (sheets === "cbOnly") return r.inCobra && !r.inNS && !r.inSch5;
    if (sheets === "s5Only") return r.inSch5 && !r.inNS && !r.inCobra;
    if (sheets === "noNS") return !r.inNS;
    if (sheets === "noCobra") return !r.inCobra;
    if (sheets === "noSch5") return !r.inSch5;
    return true;
  };
  const byPaid = (r) => {
    if (paidView === "only") return r.paidFlag;
    if (paidView === "hide") return !r.paidFlag;
    return true;
  };

  const filtered = useMemo(() => enriched.filter(
    (r) => (group === "all" || (r.productGroup || "(blank)") === group) && bySheets(r) && byPaid(r)
  ), [enriched, group, sheets, paidView]);

  const getVal = (r, col) => {
    switch (col) {
      case "order": return r.orderNum || "";
      case "nsref": return r.netsuiteRef || "";
      case "company": return (r.company || "").toLowerCase();
      case "item": return (r.product || "").toLowerCase();
      case "group": return (r.productGroup || "").toLowerCase();
      case "status": return (r.nsStatus || "").toLowerCase();
      case "risk": return ["high","medium","low","none",""].indexOf(r.risk || "");
      case "sheets": return r.sheetCode;
      case "nsSov": return r.inNS ? r.nsSov : null;
      case "s5Sov": return r.inSch5 ? r.s5Sov : null;
      case "cbSov": return r.inCobra ? r.cobraSov : null;
      case "sovSpread": return r.sovSpread;
      case "nsGp": return r.inNS ? r.expected : null;
      case "cbGp": return r.inCobra ? r.paid : null;
      case "gpDiff": return r.gpDiff;
      default: return Math.abs(r.sovSpread) + Math.abs(r.gpDiff);
    }
  };

  const rows = useMemo(() => [...filtered].sort((a, b) => {
    const av = getVal(a, sort.col), bv = getVal(b, sort.col);
    const an = av == null, bn = bv == null;
    if (an && bn) return 0;
    if (an) return 1;
    if (bn) return -1;
    let c;
    if (typeof av === "number" && typeof bv === "number") c = av - bv;
    else c = String(av).localeCompare(String(bv));
    return sort.dir === "asc" ? c : -c;
  }), [filtered, sort]);

  const T = useMemo(() => {
    const t = { nsSov: 0, nsGp: 0, s5Sov: 0, cbSov: 0, cbGp: 0, netSov: 0, netGp: 0 };
    for (const r of filtered) {
      t.nsSov += r.nsSov || 0; t.nsGp += r.expected || 0; t.s5Sov += r.s5Sov || 0;
      t.cbSov += r.cobraSov || 0; t.cbGp += r.paid || 0;
      t.netSov += r.netSov || 0; t.netGp += r.netGp || 0;
    }
    return t;
  }, [filtered]);

  const shown = rows.slice(0, showCount);
  const cell = (v) => <span className="num mono">{gbp(v)}</span>;
  const diffCell = (v) => <span className={"num mono " + (Math.abs(v) < 0.005 ? "" : v > 0 ? "pos" : "neg")}>{gbp(v)}</span>;
  const selStyle = { padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 };

  const click = (col) =>
    setSort((s) => ({ col, dir: s.col === col ? (s.dir === "asc" ? "desc" : "asc") : (["order", "nsref", "company", "item", "group", "status", "sheets"].includes(col) ? "asc" : "desc") }));
  const H = ({ col, label, num, left }) => (
    <th className={left ? "left" : "num"} style={{ cursor: "pointer", userSelect: "none" }} onClick={() => click(col)}>
      {label}{sort.col === col ? (sort.dir === "asc" ? " ▲" : " ▼") : ""}
    </th>
  );

  return (
    <>
      <div className="panel">
        <h2>Totals — SOV &amp; GP across the three sheets</h2>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>Metric</th><th className="num">NetSuite</th><th className="num">Sch5</th><th className="num">Cobra</th>
                <th className="num">Cobra − NS</th><th className="num">Sch5 − NS</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><strong>SOV</strong> <span className="sub">(Contract Value / SOV)</span></td>
                <td className="num mono">{gbp(T.nsSov)}</td>
                <td className="num mono">{gbp(T.s5Sov)}</td>
                <td className="num mono">{gbp(T.cbSov)}</td>
                <td>{diffCell(T.cbSov - T.nsSov)}</td>
                <td>{diffCell(T.s5Sov - T.nsSov)}</td>
              </tr>
              <tr>
                <td><strong>GP</strong> <span className="sub">(Product GP / Commission Paid)</span></td>
                <td className="num mono">{gbp(T.nsGp)}</td>
                <td className="num mono">-</td>
                <td className="num mono">{gbp(T.cbGp)}</td>
                <td>{diffCell(T.cbGp - T.nsGp)}</td>
                <td className="num mono">-</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="note">Green = Cobra paid more than we expected. Red = paid less. Totals follow the filters below.</p>
      </div>

      <div className="panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <h2 style={{ margin: 0, marginBottom: 10 }}>By order reference</h2>
          <button className="btn" onClick={() => downloadCSV(rows.map((r) => ({
            "Order ref": r.orderNum, "NetSuite ref": r.netsuiteRef || "", Company: r.company,
            "Item name": r.product || "", Product: r.productGroup || "", "Order status": r.nsStatus || "",
            Risk: r.risk || "", "On NetSuite": r.inNS ? "Y" : "", "On Cobra": r.inCobra ? "Y" : "", "On Sch5": r.inSch5 ? "Y" : "",
            "NS SOV": r.inNS ? r.nsSov : "", "Sch5 SOV": r.inSch5 ? r.s5Sov : "", "Cobra SOV": r.inCobra ? r.cobraSov : "",
            "SOV diff": r.sovSpread, "NS GP": r.inNS ? r.expected : "", "Cobra GP": r.inCobra ? r.paid : "", "GP diff": r.gpDiff,
          })), "cross-reference.csv")}>Export CSV</button>
        </div>
        <div className="settings" style={{ marginBottom: 12 }}>
          <span>Product:</span>
          <select value={group} onChange={(e) => setGroup(e.target.value)} style={selStyle}>
            <option value="all">All products</option>
            {groups.map((g) => <option key={g} value={g}>{g}</option>)}
            <option value="(blank)">(blank)</option>
          </select>
          <span style={{ width: 12 }} />
          <span>Appears on:</span>
          <select value={sheets} onChange={(e) => setSheets(e.target.value)} style={selStyle}>
            <option value="all">Any sheet</option>
            <option value="all3">All three sheets</option>
            <option value="missing">Missing from at least one</option>
            <option value="noNS">Not on NetSuite</option>
            <option value="noCobra">Not on Cobra</option>
            <option value="noSch5">Not on Sch5</option>
            <option value="nsOnly">NetSuite only</option>
            <option value="cbOnly">Cobra only</option>
            <option value="s5Only">Sch5 only</option>
          </select>
          <span style={{ width: 12 }} />
          <span>Paid orders:</span>
          <select value={paidView} onChange={(e) => setPaidView(e.target.value)} style={selStyle}>
            <option value="show">Show paid</option>
            <option value="only">Show only paid</option>
            <option value="hide">Do not show paid</option>
          </select>
          <span style={{ color: "#8a8aa3" }}>{filtered.length} of {enriched.length} references</span>
        </div>
        <p className="note" style={{ marginTop: 0 }}>
          Joined on order reference — <span className="mono">Order ref</span> (NetSuite) = <span className="mono">MAIN ORDER NUM</span> (Sch5) = <span className="mono">Job Header</span> (Cobra).
          Click any column heading to sort.
        </p>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <H col="order" label="Order ref" left />
                <H col="nsref" label="NetSuite ref" left />
                <H col="company" label="Company" left />
                <H col="item" label="Item name" left />
                <H col="group" label="Product" left />
                <H col="status" label="Order status" />
                <H col="risk" label="Risk" num />
                <H col="sheets" label="Sheets" num />
                <H col="nsSov" label="NS SOV" num />
                <H col="s5Sov" label="Sch5 SOV" num />
                <H col="cbSov" label="Cobra SOV" num />
                <H col="sovSpread" label="SOV Δ" num />
                <H col="nsGp" label="NS GP" num />
                <H col="cbGp" label="Cobra GP" num />
                <H col="gpDiff" label="GP Δ" num />
              </tr>
            </thead>
            <tbody>
              {shown.map((r) => (
                <tr key={r.key}>
                  <td className="left mono">{r.orderNum}</td>
                  <td className="left mono">{r.netsuiteRef || "-"}</td>
                  <td className="left">{r.company}</td>
                  <td className="left">{r.product || "-"}</td>
                  <td className="left">{r.productGroup || "-"}</td>
                  <td>{r.nsStatus || "-"}</td>
                  <td className="num">{riskChip(r.risk) || "-"}</td>
                  <td className="num"><Presence ns={r.inNS} cb={r.inCobra} s5={r.inSch5} /></td>
                  <td>{r.inNS ? cell(r.nsSov) : <span className="num sub">-</span>}</td>
                  <td>{r.inSch5 ? cell(r.s5Sov) : <span className="num sub">-</span>}</td>
                  <td>{r.inCobra ? cell(r.cobraSov) : <span className="num sub">-</span>}</td>
                  <td>{diffCell(r.sovSpread)}</td>
                  <td>{r.inNS ? cell(r.expected) : <span className="num sub">-</span>}</td>
                  <td>{r.inCobra ? cell(r.paid) : <span className="num sub">-</span>}</td>
                  <td>{diffCell(r.gpDiff)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {rows.length > shown.length && (
          <p className="note">
            Showing {shown.length.toLocaleString()} of {rows.length.toLocaleString()} references.{" "}
            <button className="btn" onClick={() => setShowCount((n) => n + 500)}>Show 500 more</button>
          </p>
        )}
        {rows.length === 0 && <div className="empty">Nothing matches these filters.</div>}
      </div>
    </>
  );
}

// =========================================================================
//  Shared helpers for the month-based tabs
// =========================================================================
const MONTHS_FY = [
  ["04", "Apr"], ["05", "May"], ["06", "June"], ["07", "July"], ["08", "Aug"], ["09", "Sept"],
  ["10", "Oct"], ["11", "Nov"], ["12", "Dec"], ["01", "Jan"], ["02", "Feb"], ["03", "March"],
];
// financial year starting April: monthKey "2025-06" -> 2025 ; "2026-02" -> 2025
const fyStartOf = (key) => {
  if (!key) return null;
  const [y, m] = key.split("-").map(Number);
  return m >= 4 ? y : y - 1;
};
const fyLabel = (start) => `${start}>${start + 1}`;
// the 12 monthKeys of a financial year, in Apr..Mar order
const fyMonthKeys = (start) =>
  MONTHS_FY.map(([mm]) => `${Number(mm) >= 4 ? start : start + 1}-${mm}`);

// a payplan month column could hold "2026-04", "202604", "01/04/2026", "Apr 2026", "April 2026"…
const MONTH_NAMES = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const planMonthKey = (v) => {
  if (v == null || v === "") return null;
  const direct = monthKey(v);
  if (direct) return direct;
  const s = String(v).trim().toLowerCase();
  const yr = s.match(/(20\d{2})/);
  const nm = MONTH_NAMES.findIndex((m) => s.includes(m));
  if (yr && nm >= 0) return `${yr[1]}-${String(nm + 1).padStart(2, "0")}`;
  return null;
};

const pct = (n, d) => (!d ? null : (n / d) * 100);
const fmtPct = (v) => (v == null ? "-" : v.toFixed(2) + "%");
const gbp0 = (n) =>
  n == null ? "-" : new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP", maximumFractionDigits: 0 }).format(n);

// pull NetSuite + Cobra lines out of the loaded files (unfiltered by the month selector)
function useSourceLines(files, shared) {
  return useMemo(() => {
    if (shared) return shared;
    const ns = (files.netsuite?.rows || []).map(nsRow).map((r) => ({ ...r, period: monthKey(r.date), rawDate: r.date }));
    const cb = (files.cobra?.rows || []).map(cobraRow).map((r) => ({ ...r, period: cobraPeriod(r) }));
    const s5 = (files.sch5?.rows || []).map(sch5Row).map((r) => ({ ...r, period: monthKey(r.date) }));
    return { ns, cb, s5 };
  }, [files, shared]);
}

// =========================================================================
//  WIP Tracker — replica of the monthly GP/WIP sheet
// =========================================================================
function WipTracker({ files, settings, saveSettings, settingsSaving, supabase, session , sources }) {
  const statusSettings = settings.risk || {};
  const { ns } = useSourceLines(files, sources);

  const years = useMemo(() => {
    const s = new Set();
    ns.forEach((r) => { const f = fyStartOf(r.period); if (f) s.add(f); });
    return [...s].sort((a, b) => b - a);
  }, [ns]);

  const [fy, setFy] = useState(null);
  const [drill, setDrill] = useState(null);   // { label, monthIndex }
  const [drillFilter, setDrillFilter] = useState("all");
  const year = fy ?? years[0] ?? null;
  const keys = year != null ? fyMonthKeys(year) : [];

  // ---- figures that come straight from the reports ----
  const rep = useMemo(() => {
    const z = () => keys.map(() => 0);
    const latestStats = z(), paidCobra = z(), unpaidWip = z(), accOwed = z(), accPaid = z(), overage = z(), redGp = z();
    const idx = new Map(keys.map((k, i) => [k, i]));
    const paySeen = new Set();
    for (const r of ns) {
      const i = idx.get(r.period);
      if (i == null) continue;
      const gp = r.expected || 0;
      if (!statusCounts(statusSettings, r.status) || isNonCommissionable(r)) { redGp[i] += gp; continue; }
      latestStats[i] += gp;
      const payKey = String(r.docNo ?? r.orderNum ?? "");
      if (r.recordedCobra != null && !paySeen.has(payKey)) { paySeen.add(payKey); paidCobra[i] += r.recordedCobra; }
      const unpaid = r.itemPaid != null ? !isYes(r.itemPaid) : !r.statusPaid;
      if (unpaid) unpaidWip[i] += gp;
      if (isYes(r.accelerator)) { accOwed[i] += gp; accPaid[i] += r.recordedCobra || 0; }
      overage[i] += money(r.overpayment) || 0;
    }
    return { latestStats, paidCobra, unpaidWip, accOwed, accPaid, overage, redGp };
  }, [ns, keys.join(), statusSettings]);

  // ---- "Original Payroll stats GP" comes from the static month-end snapshots ----
  const [statics, setStatics] = useState({});
  useEffect(() => {
    if (!supabase || !session || year == null) return;
    let cancelled = false;
    (async () => {
      const { data } = await supabase.from("recon_datasets").select("*").in("id", keys.map((mk) => `static-${mk}`));
      if (cancelled || !data) return;
      const out = {};
      for (const r of data) {
        if (cancelled) return;
        const mk = String(r.id).replace("static-", "");
        const full = await hydrate(supabase, r.netsuite, `chunk-${r.id}`);
        if (full?.rows) out[mk] = full.rows.map(nsRow);
      }
      setStatics(out);
    })();
    return () => { cancelled = true; };
  }, [session, year, keys.join()]);

  const origPayrollFromStatic = useMemo(() => keys.map((mk) => {
    const rows = statics[mk];
    if (!rows) return null;                       // no snapshot for that month yet
    let t = 0;
    for (const r of rows) {
      if (!statusCounts(statusSettings, r.status) || isNonCommissionable(r)) continue;
      t += r.expected || 0;
    }
    return t;
  }), [statics, keys.join(), statusSettings]);

  // ---- manual rows, stored in shared settings ----
  const wipAll = settings.wip || {};
  const manual = (wipAll[year] || {});
  const getM = (rowKey, i) => {
    const v = manual[rowKey]?.[i];
    return v == null || v === "" ? null : Number(v);
  };
  const setM = (rowKey, i, val) => {
    const next = JSON.parse(JSON.stringify(settings.wip || {}));
    next[year] = next[year] || {};
    next[year][rowKey] = next[year][rowKey] || [];
    next[year][rowKey][i] = val === "" ? null : Number(val);
    saveSettings({ ...settings, wip: next });
  };

  const rowVals = (rowKey) => keys.map((_, i) => getM(rowKey, i));
  const netOrig = rowVals("netOriginal");
  const origPayroll = origPayrollFromStatic;
  // "Latest Payroll GP" is no longer typed in: it now follows payable GP (statuses that count).
  const latestPayroll = rep.latestStats;
  const pillar = rowVals("pillar");
  const accUplift = rowVals("accUplift");

  // ---- formulas ----
  const n0 = (v) => (v == null ? 0 : v);
  const netVsLatest = keys.map((_, i) => n0(netOrig[i]) - rep.latestStats[i]);
  const changeToPayroll = keys.map((_, i) => n0(origPayroll[i]) - rep.latestStats[i]);
  const cancelled = keys.map((_, i) => pct(netVsLatest[i], n0(netOrig[i])));
  const paidPct = keys.map((_, i) => pct(rep.paidCobra[i], rep.latestStats[i]));


  const tot = (arr) => sum(arr.filter((v) => v != null));
  const cellNum = (v) => (v == null ? "-" : gbp0(v));

  const totalPct = {
    "How Much Cancelled?": pct(tot(netVsLatest), tot(netOrig)),
    "How much has been paid?": pct(tot(rep.paidCobra), tot(rep.latestStats)),
  };

  const Row = (props) => (
    <WipRow {...props} manual={manual} setM={setM} drill={drill} setDrill={setDrill}
      totalPct={totalPct} tot={tot} />
  );

  // unpaid WIP split by product group, for the summary strip
  const diag = useMemo(() => {
    let nsIn = 0, nsNoDate = 0, nsOther = 0;
    for (const r of ns) {
      if (!r.period) nsNoDate++;
      else if (fyStartOf(r.period) === year) nsIn++;
      else nsOther++;
    }
    return { nsIn, nsNoDate, nsOther };
  }, [ns, year]);

  const wipByProduct = useMemo(() => {
    const m = new Map();
    for (const r of ns) {
      if (fyStartOf(r.period) !== year) continue;
      if (r.itemPaid != null ? isYes(r.itemPaid) : r.statusPaid) continue;
      if (!statusCounts(statusSettings, r.status) || isNonCommissionable(r)) continue;
      const g = r.productGroup2 || productGroupOf(r.product) || "Other";
      m.set(g, (m.get(g) || 0) + (r.expected || 0));
    }
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  }, [ns, year]);

  if (year == null) return <div className="empty panel">No dated rows found in the loaded files.</div>;

  return (
    <>
      <div className="panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <h2 style={{ margin: 0 }}>Outstanding unpaid WIP — {fyLabel(year)}</h2>
          <div className="settings">
            <span>Financial year:</span>
            <select value={year} onChange={(e) => setFy(Number(e.target.value))}
              style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
              {years.map((y) => <option key={y} value={y}>{fyLabel(y)}</option>)}
            </select>
            {settingsSaving && <span className="sub" style={{ margin: 0 }}>Saving…</span>}
          </div>
        </div>
        <div style={{ overflowX: "auto", marginTop: 12 }}>
          <table>
            <thead><tr>
              {wipByProduct.map(([g]) => <th key={g} className="num">{g}</th>)}
              <th className="num">Total outstanding</th>
            </tr></thead>
            <tbody><tr>
              {wipByProduct.map(([g, v]) => <td key={g} className="num mono">{gbp(v)}</td>)}
              <td className="num mono"><strong>{gbp(sum(wipByProduct.map((x) => x[1])))}</strong></td>
            </tr></tbody>
          </table>
        </div>
        <p className="note">Unpaid WIP = NetSuite GP where Item Paid is not "Yes", grouped by product. This page reads the NetSuite report only.</p>
      </div>

      <div style={{ display: "flex", gap: 14, alignItems: "flex-start" }}>
      <div className="panel wip" style={{ flex: "1 1 auto", minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <h2 style={{ margin: 0 }}>{fyLabel(year)} — monthly GP tracker</h2>
          <div className="settings">
            <span>Financial year:</span>
            <select value={year} onChange={(e) => setFy(Number(e.target.value))}
              style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
              {years.map((y) => <option key={y} value={y}>{fyLabel(y)}</option>)}
            </select>
          </div>
        </div>
        <p className="note" style={{ marginTop: 0 }}>
          <span className="chip unmatched" style={{ fontSize: 10 }}>Manual</span> you type ·
          <span className="chip matched" style={{ fontSize: 10, marginLeft: 6 }}>Report</span> from NetSuite/Cobra ·
          <span className="chip mismatch" style={{ fontSize: 10, marginLeft: 6 }}>Formula</span> worked out.
          Manual figures are shared with everyone signed in.
        </p>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th className="left lbl">{fyLabel(year)}</th>
                {MONTHS_FY.map(([, l]) => <th key={l} className="num">{l}</th>)}
                <th className="num">Totals</th>
              </tr>
            </thead>
            <tbody>
              <Row label="Net Original stats GP" tag="Manual" kind="input" rowKey="netOriginal" vals={netOrig} />
              <Row label="Original Payroll stats GP" tag="Report" vals={origPayroll} />
              <Row group label="Net Vs Latest Diff" tag="Formula" vals={netVsLatest} danger={(v) => v > 0} />
              <Row group label="Latest stats" tag="Report" vals={rep.latestStats} />
              <Row label="Latest Paid on Cobra" tag="Report" vals={rep.paidCobra} />
              <Row label="O/S unpaid WIP" tag="Report" vals={rep.unpaidWip} />
              <Row group label="Pillar Bonus & Incentives" tag="Manual" kind="input" rowKey="pillar" vals={pillar} />
              <Row label="Accelerator Uplift" tag="Manual" kind="input" rowKey="accUplift" vals={accUplift} />
              <Row group label="Change From Original to Payroll" tag="Formula" vals={changeToPayroll} />
              <Row group label="How Much Cancelled?" tag="Formula" kind="pct" vals={cancelled} danger={(v) => v >= 20} />
              <Row label="How much has been paid?" tag="Formula" kind="pct" vals={paidPct} />
              <Row group label="Accelerator Owed" tag="Report" vals={rep.accOwed} />
              <Row label="Accelerator Paid" tag="Report" vals={rep.accPaid} />
              <Row label="Overage" tag="Report" vals={rep.overage} />
            </tbody>
          </table>
        </div>
        <p className="note">
          <strong>Source:</strong> NetSuite report only, and the month comes from the
          {" "}<span className="mono">Netsuite Date</span> column alone (day/month/year) — no other date column is used.
          Click any <span className="chip matched" style={{ fontSize: 10 }}>Report</span> figure to see the orders behind it.
          {" "}{diag.nsIn} rows fall in {fyLabel(year)}; {diag.nsNoDate} have no readable date
          {diag.nsOther > 0 ? ` and ${diag.nsOther} sit in other years` : ""}.
        </p>
        <p className="note">
          Net Vs Latest Diff = Net Original − Latest stats · Change From Original to Payroll = Original Payroll stats GP − Latest stats ·
          How Much Cancelled = Net Vs Latest Diff ÷ Net Original · How much has been paid = Paid on Cobra ÷ Latest stats ·
          Overage = Paid on Cobra − Latest stats.
          <br /><strong>Original Payroll stats GP</strong> comes from the static month-end snapshots uploaded under
          Settings → Monthly Reports (payable GP only); months with no snapshot show "-".
          <br /><strong>Note:</strong> "Latest Payroll GP" is no longer typed in — it now follows payable GP
          (everything except Red GP statuses), which is what feeds "Change From Original to Payroll" and "Change in Value".
        </p>
      </div>

      {drill && (() => {
          const mk = keys[drill.monthIndex];
          const inMonth = ns.filter((r) => r.period === mk);
          const seenDoc = new Set();
          let list = [];
          const counted = (r) => statusCounts(statusSettings, r.status) && !isNonCommissionable(r);
          if (drill.label === "Original Payroll stats GP") {
            list = (statics[mk] || []).filter((r) => statusCounts(statusSettings, r.status) && !isNonCommissionable(r));
          }
          else if (drill.label === "Latest stats") list = inMonth.filter(counted);
          else if (drill.label === "O/S unpaid WIP") list = inMonth.filter((r) => counted(r) && (r.itemPaid != null ? !isYes(r.itemPaid) : !r.statusPaid));
          else if (drill.label === "Latest Paid on Cobra") list = inMonth.filter((r) => {
            if (r.recordedCobra == null) return false;
            const k = String(r.docNo ?? r.orderNum ?? "");
            if (seenDoc.has(k)) return false;
            seenDoc.add(k); return true;
          });
          else if (drill.label === "Accelerator Owed" || drill.label === "Accelerator Paid") list = inMonth.filter((r) => isYes(r.accelerator));
          else if (drill.label === "Overage") list = inMonth.filter((r) => money(r.overpayment) != null);
          const val = (r) =>
            drill.label === "Latest Paid on Cobra" || drill.label === "Accelerator Paid" ? r.recordedCobra
              : drill.label === "Overage" ? money(r.overpayment) : r.expected;

          // compare this month against its snapshot so the same change filters apply
          const snapRows = statics[mk];
          let changeBy = null;
          if (snapRows) {
            const catOf = (r) => (isNonCommissionable(r) ? "red" : statusCategory(statusSettings, r.status));
            const bucket = (rows) => {
              const m = new Map();
              for (const r of rows) {
                if (!isRealLine(r)) continue;
                const k = matchKey(r);
                if (!k) continue;
                if (!m.has(k)) m.set(k, { ...r, gp: 0 });
                m.get(k).gp += r.expected || 0;
              }
              return m;
            };
            const diffRows = buildDiff(bucket(snapRows), bucket(inMonth), catOf);
            changeBy = new Map(diffRows.map((r) => [matchKey(r), r]));
          }
          const withChange = list.map((r) => ({ ...r, ...(changeBy?.get(matchKey(r)) || {}) }));
          const filteredList = changeBy ? withChange.filter((r) => matchesChangeFilter(r, drillFilter)) : list;
          return (
            <div className="panel side" style={{ background: "#faf9ff", marginTop: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <h2 style={{ margin: 0, fontSize: 12 }}>{drill.label}<br />{periodLabel(mk)} · {filteredList.length} rows · {gbp0(sum(filteredList.map(val)))}</h2>
                <div>
                  <button className="btn" style={{ marginRight: 6 }} onClick={() => downloadCSV(filteredList.map((r) => ({
                    "Order ref": r.orderNum || "", "NetSuite ref": r.netsuiteRef || "", Company: r.company || "",
                    "Netsuite Date": r.rawDate == null ? "" : String(r.rawDate), Partner: r.partner || "", Role: r.partnerRole || "",
                    Item: r.product || "", Status: r.status || "", Value: val(r) ?? "",
                  })), `${drill.label.replace(/[^a-z0-9]+/gi, "-")}-${mk}.csv`)}>Export</button>
                  <button className="btn" onClick={() => setDrill(null)}>Close</button>
                </div>
              </div>
              {changeBy && <ChangeFilters rows={withChange} value={drillFilter} onChange={setDrillFilter} />}
              {filteredList.length === 0 ? <div className="empty">No rows behind this figure.</div> : (
                <div style={{ overflowX: "auto", marginTop: 10, maxHeight: 420, overflowY: "auto" }}>
                  <table>
                    <thead><tr>
                      <th className="left">Order ref</th><th className="left">NetSuite ref</th><th className="left">Company</th>
                      <th className="left">Netsuite Date</th><th className="left">Partner</th><th className="left">Item</th>
                      <th className="left">Status</th><th className="left">Change</th><th className="num">Value</th>
                    </tr></thead>
                    <tbody>
                      {filteredList.slice(0, 200).map((r, i) => (
                        <tr key={i}>
                          <td className="left mono">{r.orderNum || "-"}</td>
                          <td className="left mono">{r.netsuiteRef || "-"}</td>
                          <td className="left">{r.company || "-"}</td>
                          <td className="left mono">{r.rawDate == null || r.rawDate === "" ? "(blank)" : String(r.rawDate)}</td>
                          <td className="left">{r.partner || r.agent || "-"}</td>
                          <td className="left">{r.product || "-"}</td>
                          <td className="left">{r.status || "-"}</td>
                          <td className="left">
                            {r.change === "new" ? <span className="chip matched">added</span>
                              : r.change === "removed" ? <span className="chip risk">removed</span>
                                : r.becamePayable ? <span className="chip matched">now payable</span>
                                  : r.lostPayable ? <span className="chip risk">lost payable</span>
                                    : r.valueChanged ? <span className="chip mismatch">value</span>
                                      : <span className="sub">-</span>}
                          </td>
                          <td className="num mono">{gbp(val(r))}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })()}

      </div>
    </>
  );
}

// =========================================================================
//  Payments Per Agent
// =========================================================================
// Total = everything except red. Payable = payable tone. Waiting = claimed + issued.
// export files often carry a totals/summary row with no identifiers — never treat it as an order
// Match a snapshot line to a live line on order reference + company name.
// The old key used Document Number and item text, which change between exports and made
// almost everything look "no longer on the report".
const matchKey = (r) => {
  const ord = normOrder(r.orderNum || r.netsuiteRef || "");
  const co = String(r.company || "").trim().toLowerCase().replace(/\s+/g, " ");
  return ord ? `o:${ord}` : co ? `c:${co}` : null;
};

// Compare a snapshot bucket against a live bucket and describe what moved.
function buildDiff(snap, live, catOf) {
  const rows = [];
  for (const [k, b] of live) {
    const a = snap.get(k);
    const nowCat = catOf(b);
    if (!a) {
      rows.push({ ...b, cat: nowCat, wasCat: null, wasStatus: "(not in payroll)", change: "new",
        was: null, now: b.gp, becamePayable: false, lostPayable: false });
      continue;
    }
    const wasCat = catOf(a);
    const valueMoved = Math.abs((a.gp || 0) - (b.gp || 0)) > 0.005;
    const catMoved = wasCat !== nowCat;
    const statusMoved = String(a.status || "") !== String(b.status || "");
    rows.push({
      ...b, cat: nowCat, wasCat, wasStatus: a.status, was: a.gp, now: b.gp,
      becamePayable: wasCat !== "payable" && nowCat === "payable",
      lostPayable: wasCat === "payable" && nowCat !== "payable",
      valueChanged: valueMoved,
      change: valueMoved && (catMoved || statusMoved) ? "value+status"
        : valueMoved ? "value"
          : catMoved ? "status"
            : statusMoved ? "status-only" : null,
    });
  }
  for (const [k, a] of snap) {
    if (live.has(k)) continue;
    rows.push({ ...a, wasCat: catOf(a), wasStatus: a.status, cat: null, status: "(not in live report)",
      change: "removed", was: a.gp, now: 0, becamePayable: false, lostPayable: false });
  }
  return rows;
}

const isRealLine = (r) =>
  !!(String(r.netsuiteRef || "").trim() || String(r.orderNum || "").trim() || String(r.company || "").trim());

const summariseGp = (rows, statusRules) => {
  let total = 0, payable = 0, waiting = 0;
  for (const r of rows || []) {
    if (!isRealLine(r)) continue;
    const gp = r.expected || 0;
    const cat = isNonCommissionable(r) ? "red" : statusCategory(statusRules, r.status);
    if (cat === "red") continue;
    total += gp;
    if (cat === "payable") payable += gp;
    else waiting += gp;
  }
  return { total, payable, waiting };
};

// One place for the change-type filters so all three lists behave the same.
const CHANGE_FILTERS = [
  ["all", "All changes", "#5514b4"],
  ["added", "Added since PR", "#14804a"],
  ["removed", "Removed since PR", "#b3261e"],
  ["value", "Value change", "#d98a00"],
  ["gained", "Now payable", "#14804a"],
  ["lost", "No longer payable", "#b3261e"],
];

const matchesChangeFilter = (r, f) =>
  f === "all" ? true
    : f === "added" ? r.change === "new"
      : f === "removed" ? r.change === "removed"
        : f === "value" ? !!r.valueChanged
          : f === "gained" ? !!r.becamePayable
            : f === "lost" ? !!r.lostPayable : true;

function ChangeFilters({ rows, value, onChange }) {
  return (
    <div className="fcards">
      {CHANGE_FILTERS.map(([v, label, colour]) => {
        const n = rows.filter((r) => matchesChangeFilter(r, v) && (v === "all" ? r.change : true)).length;
        return (
          <button key={v} className={"fcard " + (value === v ? "on" : "")}
            style={{ borderTopColor: colour }} onClick={() => onChange(v)}>
            <span className="n" style={{ color: colour }}>{n.toLocaleString()}</span>
            <span className="l">{label}</span>
          </button>
        );
      })}
    </div>
  );
}

function NumInput({ value, placeholder, onCommit, style }) {
  const [v, setV] = useState(value ?? "");
  useEffect(() => { setV(value ?? ""); }, [value]);
  return (
    <input
      type="number" value={v} placeholder={placeholder}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => { if (String(v) !== String(value ?? "")) onCommit(v); }}
      onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
      onPaste={(e) => {
        const t = (e.clipboardData || window.clipboardData).getData("text");
        const n = t.replace(/[^0-9.\-]/g, "");
        if (n !== t) { e.preventDefault(); setV(n); }
      }}
      style={style}
    />
  );
}

const blankAgent = (n) => ({
  earned: Array(n).fill(0), payable: Array(n).fill(0), claimed: Array(n).fill(0),
  issued: Array(n).fill(0), red: Array(n).fill(0), claimedIssued: Array(n).fill(0), paid: Array(n).fill(0),
});

function AgentPayments({ files, settings, saveSettings, staffNames = [], dbPlans = {}, dbPlansByMonth = {}, dbPlanNames = {},
  managers = {}, statusRules = {}, supabase, session , sources }) {
  const monthlyPlans = settings.payplanMonthly || {};
  const { ns, cb } = useSourceLines(files, sources);
  const [basis, setBasis] = useState("earned");

  const years = useMemo(() => {
    const s = new Set();
    ns.forEach((r) => { const f = fyStartOf(r.period); if (f) s.add(f); });
    return [...s].sort((a, b) => b - a);
  }, [ns]);
  const [fy, setFy] = useState(null);
  const year = fy ?? years[0] ?? null;
  const keys = year != null ? fyMonthKeys(year) : [];
  const idx = useMemo(() => new Map(keys.map((k, i) => [k, i])), [keys.join()]);

  // Cobra paid, keyed by order ref, so we can attribute payments to the NetSuite agent
  const paidByOrder = useMemo(() => {
    const m = new Map();
    for (const r of cb) m.set(r.orderNum, (m.get(r.orderNum) || 0) + (r.paid || 0));
    return m;
  }, [cb]);

  const agents = useMemo(() => {
    const m = new Map();
    for (const r of ns) {
      const i = idx.get(r.period);
      if (i == null) continue;
      const name = r.partner || r.agent || "(unassigned)";   // GP is credited to the Partner
      if (!m.has(name)) m.set(name, blankAgent(keys.length));
      const g = m.get(name);
      const gp = r.expected || 0;
      const cat = isNonCommissionable(r) ? "red" : statusCategory(statusRules, r.status);
      if (!g[cat]) g[cat] = Array(keys.length).fill(0);   // guard against an unexpected category
      g[cat][i] += gp;
      if (cat !== "red") g.earned[i] += gp;                 // Total GP ignores red
      if (cat === "claimed" || cat === "issued") g.claimedIssued[i] += gp;
      g.paid[i] += paidByOrder.get(r.orderNum) || 0;
    }
    for (const n of staffNames) if (!m.has(n)) m.set(n, blankAgent(keys.length));
    // if we have a staff list, only show those people (drops Office Doublecount, ex-staff, Tracy's team)
    const allow = new Set(staffNames.map((n) => n.toLowerCase()));
    const out = [...m.entries()].filter(([n]) => !staffNames.length || allow.has(n.toLowerCase()));
    return out.sort((a, b) => a[0].localeCompare(b[0]));   // alphabetical
  }, [ns, idx, paidByOrder, keys.join(), staffNames.join(), statusRules]);

  // static month-end snapshots for this financial year
  const [statics, setStatics] = useState({});
  useEffect(() => {
    if (!supabase || !session || year == null) return;
    let cancelled = false;
    (async () => {
      const { data } = await supabase.from("recon_datasets").select("*").in("id", keys.map((mk) => `static-${mk}`));
      if (cancelled || !data) return;
      const out = {};
      for (const r of data) {
        if (cancelled) return;
        const mk = String(r.id).replace("static-", "");
        const full = await hydrate(supabase, r.netsuite, `chunk-${r.id}`);
        if (full?.rows) out[mk] = full.rows.map(nsRow);
      }
      setStatics(out);
    })();
    return () => { cancelled = true; };
  }, [session, year, keys.join()]);


  // what payroll saw, per agent per month
  const staticByAgent = useMemo(() => {
    const out = {};
    keys.forEach((mk, i) => {
      for (const r of statics[mk] || []) {
        if (!isRealLine(r)) continue;
        const n = r.partner || r.agent || "(unassigned)";
        out[n] = out[n] || keys.map(() => ({ total: 0, payable: 0, waiting: 0 }));
        const gp = r.expected || 0;
        const cat = isNonCommissionable(r) ? "red" : statusCategory(statusRules, r.status);
        if (cat === "red") continue;
        out[n][i].total += gp;
        if (cat === "payable") out[n][i].payable += gp;
        else out[n][i].waiting += gp;
      }
    });
    return out;
  }, [statics, keys.join(), statusRules]);

  // Every agent+month compared in a single pass — rebuilding this per cell was the
  // main reason the page crawled.
  const allChanges = useMemo(() => {
    const out = new Map();
    if (!Object.keys(statics).length) return out;
    const catOf = (r) => (isNonCommissionable(r) ? "red" : statusCategory(statusRules, r.status));
    const agentOf = (r) => r.partner || r.agent || "(unassigned)";

    // bucket the live report once: agent|monthIndex -> Map(matchKey -> aggregated line)
    const liveBy = new Map();
    for (const r of ns) {
      const i = idx.get(r.period);
      if (i == null || !isRealLine(r)) continue;
      const k = matchKey(r);
      if (!k) continue;
      const bk = `${agentOf(r)}|${i}`;
      if (!liveBy.has(bk)) liveBy.set(bk, new Map());
      const m = liveBy.get(bk);
      if (!m.has(k)) m.set(k, { ...r, gp: 0 });
      m.get(k).gp += r.expected || 0;
    }

    const snapBy = new Map();
    keys.forEach((mk, i) => {
      for (const r of statics[mk] || []) {
        if (!isRealLine(r)) continue;
        const k = matchKey(r);
        if (!k) continue;
        const bk = `${agentOf(r)}|${i}`;
        if (!snapBy.has(bk)) snapBy.set(bk, new Map());
        const m = snapBy.get(bk);
        if (!m.has(k)) m.set(k, { ...r, gp: 0 });
        m.get(k).gp += r.expected || 0;
      }
    });

    const bucketKeys = new Set([...liveBy.keys(), ...snapBy.keys()]);
    for (const bk of bucketKeys) {
      const i = Number(bk.split("|").pop());
      if (!statics[keys[i]]) continue;               // no snapshot for that month
      const snap = snapBy.get(bk) || new Map();
      const live = liveBy.get(bk) || new Map();
      const rows = buildDiff(snap, live, catOf);
      out.set(bk, rows);
    }
    return out;
  }, [statics, ns, keys.join(), statusRules, idx]);

  const changesFor = useCallback((agent, i) =>
    (statics[keys[i]] ? (allChanges.get(`${agent}|${i}`) || []) : null), [allChanges, statics, keys.join()]);

  const [drill, setDrill] = useState(null);   // { agent, i }
  const [drillFilter, setDrillFilter] = useState("all");   // all | added | removed | value | gained | lost

  // manager filter
  const [mgr, setMgr] = useState("all");
  const managerList = useMemo(
    () => [...new Set(Object.values(managers).filter(Boolean))].sort(), [managers]);
  const shownAgents = useMemo(
    () => agents.filter(([n]) => mgr === "all" || (managers[n] || "") === mgr), [agents, mgr, managers]);

  const payplans = { ...dbPlans, ...(settings.payplans || {}) };   // manual overrides win
  const paidMarks = settings.paidMarks || {};
  const markKey = (agent, i) => `${year}|${agent}|${i}`;
  const isMarkedPaid = (agent, i) => !!paidMarks[markKey(agent, i)];
  const togglePaid = (agent, i) => {
    const next = { ...paidMarks };
    const k = markKey(agent, i);
    if (next[k]) delete next[k]; else next[k] = true;
    saveSettings({ ...settings, paidMarks: next });
  };

  const planFor = (agent, i) => {
    const m = monthlyPlans[year]?.[agent]?.[i];
    if (m !== "" && m != null) return Number(m);
    const fromDb = dbPlansByMonth[agent]?.[keys[i]];       // what the database says they were on that month
    if (fromDb != null) return Number(fromDb);
    const std = payplans[agent];
    return std === "" || std == null ? 0 : Number(std);
  };
  const planNameFor = (agent, i) => dbPlanNames[agent]?.[keys[i]] || dbPlanNames[agent]?.all || null;
  const below = (agent, v, i) => {
    const plan = planFor(agent, i);
    return plan > 0 && v < plan;
  };

  // paid in a month where the agent didn't reach their target
  const breaches = useMemo(() => {
    const out = [];
    for (const [name, g] of shownAgents) {
      keys.forEach((mk, i) => {
        const target = planFor(name, i);
        if (!(target > 0)) return;
        const earned = g.payable[i] || 0;
        const paidOut = g.paid[i] || 0;
        const marked = isMarkedPaid(name, i);
        if (earned < target && (paidOut > 0 || marked)) {
          out.push({ name, i, mk, target, earned, paidOut, marked, shortfall: target - earned, plan: planNameFor(name, i) });
        }
      });
    }
    return out.sort((a, b) => b.shortfall - a.shortfall);
  }, [shownAgents, keys.join(), monthlyPlans, payplans, dbPlansByMonth, paidMarks]);

  if (year == null) return <div className="empty panel">No dated NetSuite rows found.</div>;

  // "Now payable" / "No longer payable" / added / removed per month, across the agents shown
  const monthlyMoves = useMemo(() => {
    const blank = keys.map(() => ({ gained: 0, lost: 0, nGained: 0, nLost: 0, added: 0, removed: 0, nAdded: 0, nRemoved: 0 }));
    if (!Object.keys(statics).length) return null;
    const allow = new Set(shownAgents.map(([n]) => n));
    for (const [bk, rows] of allChanges) {
      const parts = bk.split("|");
      const i = Number(parts.pop());
      if (!allow.has(parts.join("|"))) continue;
      for (const r of rows) {
        if (r.becamePayable) { blank[i].gained += r.now || 0; blank[i].nGained++; }
        if (r.lostPayable) { blank[i].lost += r.was || 0; blank[i].nLost++; }
        if (r.change === "new") { blank[i].added += r.now || 0; blank[i].nAdded++; }
        if (r.change === "removed") { blank[i].removed += r.was || 0; blank[i].nRemoved++; }
      }
    }
    return blank;
  }, [shownAgents, statics, keys.join(), allChanges]);

  // payable totals per month: what payroll saw vs what the live report says now
  const payableRows = useMemo(() => {
    const stat = keys.map(() => 0), now = keys.map(() => 0);
    for (const [name, g] of shownAgents) {
      keys.forEach((_, i) => {
        stat[i] += staticByAgent[name]?.[i]?.payable || 0;
        now[i] += g.payable[i] || 0;
      });
    }
    return { stat, now, hasStatic: Object.keys(statics).length > 0 };
  }, [shownAgents, staticByAgent, keys.join(), statics]);

  const drillRows = drill ? changesFor(drill.agent, drill.i) : null;

  return (
    <div style={{ display: "flex", gap: 14, alignItems: "flex-start" }}>
    <div className="panel" style={{ flex: "1 1 auto", minWidth: 0 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <h2 style={{ margin: 0 }}>Payroll — {fyLabel(year)}</h2>
        <div className="settings">
          <span>Financial year:</span>
          <select value={year} onChange={(e) => setFy(Number(e.target.value))}
            style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
            {years.map((y) => <option key={y} value={y}>{fyLabel(y)}</option>)}
          </select>
          <span style={{ width: 12 }} />
          <span>Manager:</span>
          <select value={mgr} onChange={(e) => setMgr(e.target.value)}
            style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
            <option value="all">All managers</option>
            {managerList.map((mm) => <option key={mm} value={mm}>{mm}</option>)}
          </select>
          <span style={{ width: 12 }} />
          <span>Show:</span>
          <select value={basis} onChange={(e) => setBasis(e.target.value)}
            style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
            <option value="earned">Total GP</option>
            <option value="payable">Payable GP only</option>
            <option value="claimedIssued">Waiting only</option>
            <option value="paid">Paid on Cobra</option>
          </select>
        </div>
      </div>
      <p className="note" style={{ marginTop: 6 }}>
        Each cell shows <strong>Total GP</strong> (red statuses ignored), with{" "}
        <span style={{ color: "#14804a" }}>payable</span> and{" "}
        <span style={{ color: "#8a5a00" }}>waiting</span> beneath it — categories come from the
        status_config tone. Click a cell to see that month's orders on the right. Red background = below target. Targets come from the database where available, and can be
        overridden in Settings → Payplans. Tick a cell to mark the agent paid for that month.
      </p>

      {breaches.length > 0 && (
        <div className="banner" style={{ marginBottom: 12 }}>
          <strong>{breaches.length} payment{breaches.length > 1 ? "s" : ""} made in a month below target.</strong>
          <div style={{ overflowX: "auto", marginTop: 8 }}>
            <table>
              <thead><tr>
                <th className="left">Agent</th><th className="left">Month</th><th className="left">Plan</th>
                <th className="num">Target</th><th className="num">GP earned</th><th className="num">Shortfall</th>
                <th className="num">Paid</th>
              </tr></thead>
              <tbody>
                {breaches.slice(0, 40).map((b, n) => (
                  <tr key={n}>
                    <td className="left">{b.name}</td>
                    <td className="left">{periodLabel(b.mk)}</td>
                    <td className="left">{b.plan || "-"}</td>
                    <td className="num mono">{gbp0(b.target)}</td>
                    <td className="num mono">{gbp0(b.earned)}</td>
                    <td className="num mono neg">{gbp0(b.shortfall)}</td>
                    <td className="num mono">{b.paidOut > 0 ? gbp0(b.paidOut) : (b.marked ? "marked paid" : "-")}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {breaches.length > 40 && <p className="note">Showing 40 of {breaches.length}.</p>}
        </div>
      )}
      <div style={{ overflowX: "auto" }} className="pay">
        <table>
          <thead>
            <tr>
              <th className="left agent">Agent</th>
              <th className="num">Payplan</th>
              {MONTHS_FY.map(([, l]) => <th key={l} className="num">{l}</th>)}
              <th className="num">Total</th>
            </tr>
          </thead>
          <tbody>
            {monthlyMoves && (
              <>
                <tr className="moves">
                  <td className="left agent">Now payable</td>
                  <td className="num">-</td>
                  {monthlyMoves.map((m, i) => (
                    <td key={i} className="num mono" style={{ color: m.gained ? "#14804a" : undefined }}>
                      {m.gained ? gbp0(m.gained) : "-"}
                      {m.nGained > 0 && <div style={{ fontSize: 9.5, fontWeight: 400 }}>{m.nGained} order{m.nGained === 1 ? "" : "s"}</div>}
                    </td>
                  ))}
                  <td className="num mono">{gbp0(sum(monthlyMoves.map((m) => m.gained)))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent">No longer payable</td>
                  <td className="num">-</td>
                  {monthlyMoves.map((m, i) => (
                    <td key={i} className="num mono" style={{ color: m.lost ? "#b3261e" : undefined }}>
                      {m.lost ? gbp0(m.lost) : "-"}
                      {m.nLost > 0 && <div style={{ fontSize: 9.5, fontWeight: 400 }}>{m.nLost} order{m.nLost === 1 ? "" : "s"}</div>}
                    </td>
                  ))}
                  <td className="num mono">{gbp0(sum(monthlyMoves.map((m) => m.lost)))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent">Added since PR</td>
                  <td className="num">-</td>
                  {monthlyMoves.map((m, i) => (
                    <td key={i} className="num mono" style={{ color: m.added ? "#14804a" : undefined }}>
                      {m.added ? gbp0(m.added) : "-"}
                      {m.nAdded > 0 && <div style={{ fontSize: 9.5, fontWeight: 400 }}>{m.nAdded} order{m.nAdded === 1 ? "" : "s"}</div>}
                    </td>
                  ))}
                  <td className="num mono">{gbp0(sum(monthlyMoves.map((m) => m.added)))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent">Removed since PR</td>
                  <td className="num">-</td>
                  {monthlyMoves.map((m, i) => (
                    <td key={i} className="num mono" style={{ color: m.removed ? "#b3261e" : undefined }}>
                      {m.removed ? gbp0(m.removed) : "-"}
                      {m.nRemoved > 0 && <div style={{ fontSize: 9.5, fontWeight: 400 }}>{m.nRemoved} order{m.nRemoved === 1 ? "" : "s"}</div>}
                    </td>
                  ))}
                  <td className="num mono">{gbp0(sum(monthlyMoves.map((m) => m.removed)))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent">Payroll total payable</td>
                  <td className="num">-</td>
                  {payableRows.stat.map((v, i) => (
                    <td key={i} className="num mono">{v ? gbp0(v) : "-"}</td>
                  ))}
                  <td className="num mono">{gbp0(sum(payableRows.stat))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent">Current total payable</td>
                  <td className="num">-</td>
                  {payableRows.now.map((v, i) => (
                    <td key={i} className="num mono">{v ? gbp0(v) : "-"}</td>
                  ))}
                  <td className="num mono">{gbp0(sum(payableRows.now))}</td>
                </tr>
                <tr className="moves">
                  <td className="left agent"><strong>Difference</strong></td>
                  <td className="num">-</td>
                  {payableRows.now.map((v, i) => {
                    const d = v - payableRows.stat[i];
                    return (
                      <td key={i} className={"num mono " + (d === 0 ? "" : d > 0 ? "pos" : "neg")}>
                        {payableRows.stat[i] || v ? `${d > 0 ? "+" : ""}${gbp0(d)}` : "-"}
                      </td>
                    );
                  })}
                  <td className={"num mono " + (sum(payableRows.now) - sum(payableRows.stat) >= 0 ? "pos" : "neg")}>
                    {sum(payableRows.now) - sum(payableRows.stat) > 0 ? "+" : ""}
                    {gbp0(sum(payableRows.now) - sum(payableRows.stat))}
                  </td>
                </tr>
              </>
            )}
            {shownAgents.map(([name, g]) => {
              const vals = g[basis] || g.earned;
              return (
                <tr key={name}>
                  <td className="left agent">
                    {name}
                    {managers[name] && <div className="sub" style={{ margin: 0, fontSize: 10 }}>{managers[name]}</div>}
                  </td>
                  <td className="num mono">{payplans[name] ?? dbPlans[name] ? gbp(Number(payplans[name] ?? dbPlans[name])) : "-"}</td>
                  {vals.map((v, i) => {
                    const bad = below(name, g.payable[i] || 0, i);
                    const plan = planFor(name, i);
                    const marked = isMarkedPaid(name, i);
                    return (
                      <td key={i} className="num mono"
                        onClick={() => setDrill(drill && drill.agent === name && drill.i === i ? null : { agent: name, i })}
                        title="Click to see the orders behind this month"
                        style={{ cursor: "pointer",
                          boxShadow: drill && drill.agent === name && drill.i === i ? "inset 0 0 0 2px #5514b4" : undefined,
                          background: bad && (marked || (g.paid[i] || 0) > 0) ? "#f9d9d5"
                            : bad ? "#fbe9e7" : marked ? "#e6f4ec" : undefined,
                          color: bad ? "#b3261e" : undefined,
                          outline: bad && (marked || (g.paid[i] || 0) > 0) ? "2px solid #b3261e" : undefined }}>
                        {basis === "earned" ? (
                          <>
                            {staticByAgent[name] && (
                              <div style={{ fontSize: 10.5, color: "#5b5676" }}>
                                Payroll: {gbp0(staticByAgent[name][i].total)}
                              </div>
                            )}
                            <div><strong>Now: {v ? gbp0(v) : "-"}</strong></div>
                            <div style={{ fontSize: 10, color: "#14804a" }}>payable {gbp0(g.payable[i] || 0)}</div>
                            <div style={{ fontSize: 10, color: "#8a5a00" }}>waiting {gbp0(g.claimedIssued[i] || 0)}</div>
                          </>
                        ) : <div>{v ? gbp0(v) : "-"}</div>}
                        {plan > 0 && <div style={{ fontSize: 10, color: bad ? "#b3261e" : "#8a8aa3" }}>plan {gbp0(plan)}</div>}
                        <label style={{ fontSize: 10, color: "#8a8aa3", cursor: "pointer", display: "inline-flex", gap: 3, alignItems: "center" }}>
                          <input type="checkbox" checked={marked} onChange={() => togglePaid(name, i)} style={{ margin: 0 }} />
                          paid
                        </label>
                      </td>
                    );
                  })}
                  <td className="num mono"><strong>{gbp0(sum(vals))}</strong></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {shownAgents.length === 0 && <div className="empty">No agents match this filter.</div>}
    </div>

    {drill && (
      <div className="panel side">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <h2 style={{ margin: 0 }}>{drill.agent} — {periodLabel(keys[drill.i])}</h2>
          <button className="btn" onClick={() => setDrill(null)}>Close</button>
        </div>
        {!drillRows ? (
          <div className="empty">
            No static snapshot saved for {periodLabel(keys[drill.i])}.
            Upload one under <strong>Settings → Monthly Reports</strong> to compare.
          </div>
        ) : (() => {
          const base = drillRows.filter((r) => matchesChangeFilter(r, drillFilter));
          const moved = base.filter((r) => r.change);
          const rest = drillFilter === "all" ? base.filter((r) => !r.change) : [];
          const hue = (r) => r.change === "removed" ? "#f7c8c2"
            : r.change === "value+status" ? "#fbd5cf"
              : r.change === "status" ? "#fde2dd"
                : r.change === "status-only" ? "#fef4f2"
                  : r.change === "value" ? "#fdecea"
                    : r.change === "new" ? "#e6f4ec" : undefined;
          const Line = ({ r, i }) => (
            <tr key={i} style={{ background: hue(r) }}>
              <td className="left">
                <div className="mono">{r.orderNum || r.netsuiteRef || "-"}</div>
                <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.company || "-"}</div>
                <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.product || "-"}</div>
              </td>
              <td className="left">
                {(() => {
                  const moved = (r.wasStatus != null && String(r.wasStatus) !== String(r.status)) || (r.wasCat && r.wasCat !== r.cat);
                  if (!moved) return (
                    <>
                      <span className="chip" style={CAT_STYLE[r.cat] || {}}>{r.cat}</span>
                      <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.status || "-"}</div>
                    </>
                  );
                  return (
                    <>
                      <div style={{ fontSize: 10.5, color: "#8a8aa3" }}>was</div>
                      <div>
                        <span className="chip" style={CAT_STYLE[r.wasCat] || {}}>{r.wasCat || "-"}</span>{" "}
                        <span style={{ fontSize: 11 }}>{r.wasStatus || "-"}</span>
                      </div>
                      <div style={{ fontSize: 10.5, color: "#8a8aa3", marginTop: 3 }}>now</div>
                      <div>
                        <span className="chip" style={CAT_STYLE[r.cat] || {}}>{r.cat || "-"}</span>{" "}
                        <span style={{ fontSize: 11 }}>{r.status || "-"}</span>
                      </div>
                    </>
                  );
                })()}
              </td>
              <td className="num mono">
                <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>was {r.was == null ? "-" : gbp0(r.was)}</div>
                <div>now {gbp0(r.now)}</div>
                {r.becamePayable && <div className="chip matched" style={{ fontSize: 9.5 }}>now payable</div>}
                {r.lostPayable && <div className="chip risk" style={{ fontSize: 9.5 }}>lost payable</div>}
              </td>
              <td className={"num mono " + (((r.now || 0) - (r.was || 0)) === 0 ? "" : ((r.now || 0) - (r.was || 0)) > 0 ? "pos" : "neg")}>
                {r.was == null ? "new" : (((r.now || 0) - (r.was || 0)) > 0 ? "+" : "") + gbp0((r.now || 0) - (r.was || 0))}
              </td>
            </tr>
          );
          return (
            <>
              {(() => {
                const stat = staticByAgent[drill.agent]?.[drill.i] || { total: 0, payable: 0, waiting: 0 };
                const g = (shownAgents.find(([n]) => n === drill.agent) || [, blankAgent(keys.length)])[1];
                const now = { total: g.earned[drill.i] || 0, payable: g.payable[drill.i] || 0, waiting: g.claimedIssued[drill.i] || 0 };
                const row = (lab, f, cls) => (
                  <tr>
                    <td className="left">{lab}</td>
                    <td className="num mono">{gbp0(f.total)}</td>
                    <td className="num mono">{gbp0(f.payable)}</td>
                    <td className="num mono">{gbp0(f.waiting)}</td>
                  </tr>
                );
                return (
                  <table style={{ marginTop: 8, marginBottom: 10 }}>
                    <thead><tr>
                      <th className="left"></th><th className="num">Total</th>
                      <th className="num">Payable</th><th className="num">Waiting</th>
                    </tr></thead>
                    <tbody>
                      {row("Payroll", stat)}
                      {row("Now", now)}
                      <tr style={{ background: "#faf9ff" }}>
                        <td className="left"><strong>Difference</strong></td>
                        {["total", "payable", "waiting"].map((f) => {
                          const d = now[f] - stat[f];
                          return <td key={f} className={"num mono " + (d === 0 ? "" : d > 0 ? "pos" : "neg")}>
                            {d > 0 ? "+" : ""}{gbp0(d)}
                          </td>;
                        })}
                      </tr>
                    </tbody>
                  </table>
                );
              })()}
              <ChangeFilters rows={drillRows} value={drillFilter} onChange={setDrillFilter} />
              <p className="note" style={{ marginTop: 0 }}>
                {moved.length} line{moved.length === 1 ? "" : "s"} shown, compared with the {periodLabel(keys[drill.i])} snapshot.
              </p>
              <div style={{ overflowX: "auto" }}>
                <table>
                  <thead><tr>
                    <th className="left">Order</th><th className="left">Status was → now</th>
                    <th className="num">GP</th><th className="num">Difference</th>
                  </tr></thead>
                  <tbody>
                    {moved.map((r, i) => <Line key={"m" + i} r={r} i={i} />)}
                    {moved.length > 0 && rest.length > 0 && (
                      <tr><td colSpan={4} style={{ background: "#f7f6fc", fontSize: 11, color: "#8a8aa3" }}>Unchanged</td></tr>
                    )}
                    {rest.slice(0, 100).map((r, i) => <Line key={"r" + i} r={r} i={i} />)}
                  </tbody>
                </table>
              </div>
              {rest.length > 100 && <p className="note">Showing 100 unchanged lines of {rest.length}.</p>}
            </>
          );
        })()}
      </div>
    )}
    </div>
  );
}

// =========================================================================
//  Settings — Risk Levels · Payplans · Raw Data · Users
// =========================================================================
// Same rules as the SchThrive / Chris P dashboards: these statuses never count
// towards Net GP / Net SOV. Blank status is excluded too.
const RED_GP_STATUS = /rejected|sent to customer|cancelled then reissued|cancelled/i;
// status_config "tone" -> how the GP is summarised.
//   payable  = counts as Payable GP
//   claimed  = claimed but not yet payable
//   issued   = issued, awaiting the rest
//   red      = ignored entirely
const TONE_MAP = {
  payable: "payable", payble: "payable", green: "payable", success: "payable", positive: "payable",
  good: "payable", paid: "payable", complete: "payable",
  claimed: "claimed", claim: "claimed", amber: "claimed", orange: "claimed", warning: "claimed", pending: "claimed",
  issued: "issued", issue: "issued", blue: "issued", info: "issued", grey: "issued", gray: "issued", neutral: "issued",
  red: "red", danger: "red", negative: "red", error: "red", bad: "red", cancelled: "red", rejected: "red",
};
const toneCategory = (tone) => TONE_MAP[String(tone || "").trim().toLowerCase()] || null;
const toneToTreatment = toneCategory;
// everything except red contributes to Total GP
const CATEGORY_COUNTS = { payable: true, claimed: true, issued: true, red: false };
const isExcludedStatus = (st) => {
  const v = String(st || "").trim();
  if (!v) return true;
  return RED_GP_STATUS.test(v);
};
// user overrides from Settings -> Order Statuses win over the built-in rule
const statusCounts = (statusSettings, st) => {
  const key = String(st || "").trim();
  const cfg = statusSettings || {};
  const v = normaliseCategory(cfg[key] ?? cfg[key.toLowerCase()]);
  if (v) return CATEGORY_COUNTS[v] !== false;   // only "red" is excluded
  return !isExcludedStatus(st);
};
// always returns one of: payable | claimed | issued | red.
// Older saved settings used "counts"/"review", so those are mapped rather than trusted.
const LEGACY_CATEGORY = { counts: "payable", review: "claimed", none: "payable", low: "payable",
  medium: "claimed", high: "red", "": "payable" };
const normaliseCategory = (v) => {
  const s = String(v ?? "").trim().toLowerCase();
  if (s === "payable" || s === "claimed" || s === "issued" || s === "red") return s;
  return LEGACY_CATEGORY[s] || null;
};
const statusCategory = (statusSettings, st) => {
  const key = String(st || "").trim();
  const cfg = statusSettings || {};
  const v = normaliseCategory(cfg[key] ?? cfg[key.toLowerCase()]);
  if (v) return v;
  return isExcludedStatus(st) ? "red" : "payable";
};

const statusTreatment = (statusSettings, st) => {
  const key = String(st || "").trim();
  const cfg = statusSettings || {};
  return cfg[key] ?? cfg[key.toLowerCase()] ?? (isExcludedStatus(st) ? "red" : "counts");
};

const isNonCommissionable = (r) =>
  /non-?commissionable/i.test(String(r.productGroup2 || "")) ||
  /non-?commissionable/i.test(String(r.product || ""));

// how likely the GP is to fall away, by category — overridable per status in Settings
const DEFAULT_RISK_PCT = { payable: 0, claimed: 40, issued: 25, red: 100 };
const riskPctFor = (settings, statusRules, status) => {
  const key = String(status || "").trim();
  const v = (settings?.riskPct || {})[key];
  if (v != null && v !== "") return Number(v);
  return DEFAULT_RISK_PCT[statusCategory(statusRules, status)] ?? 0;
};

const CAT_STYLE = {
  payable: { background: "#e6f4ec", color: "#14804a" },
  claimed: { background: "#fbf0d9", color: "#8a5a00" },
  issued: { background: "#e8eefc", color: "#1e4fa3" },
  red: { background: "#fbe9e7", color: "#b3261e" },
};

const RISK_LEVELS = [
  ["", "Not set"],
  ["none", "No risk"],
  ["low", "Low"],
  ["medium", "Medium"],
  ["high", "High"],
];
const riskChip = (lvl) => {
  if (!lvl || lvl === "none") return null;
  const cls = lvl === "high" ? "risk" : lvl === "medium" ? "mismatch" : "unmatched";
  return <span className={"chip " + cls}>{lvl}</span>;
};

function Settings(props) {
  const { files, errors, onFile, supabase, session, saving, sharedMeta, saveShared, anyLoaded,
    isAdmin, users, newUser, setNewUser, addUser, removeUser, userMsg, saveProgress,
    settings, saveSettings, settingsSaving, records, staffNames = [], staff = {}, loadStaff,
    webosStatuses = { list: [], status: 'idle' }, probe = { status: 'idle', found: [] }, probeTables,
    statusConfig = { rows: [], status: 'idle', cols: {} }, managers = {}, statusRules = {} } = props;
  const [sub, setSub] = useState("risk");
  const thisFy = fyStartOf(`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}`);
  const planYears = [thisFy + 1, thisFy, thisFy - 1, thisFy - 2];
  const [planYear, setPlanYear] = useState(thisFy);

  const SUBS = [
    ["risk", "Order Statuses"],
    ["payplans", "Payplans"],
    ["rawdata", "Raw Data"],
    ["monthly", "Monthly Reports"],
    ...(isAdmin ? [["users", "Users"]] : []),
  ];

  // every order status seen in the loaded data
  const { nameCol: scName, toneCol: scTone } = statusConfig.cols || {};
  // offer the tone words this business actually uses, falling back to the four categories
  const toneOptions = useMemo(() => {
    const seen = new Map();
    if (scTone) for (const r of statusConfig.rows) {
      const t = String(r[scTone] || "").trim();
      const cat = toneCategory(t);
      if (t && cat && !seen.has(cat)) seen.set(cat, t);
    }
    for (const c of ["payable", "claimed", "issued", "red"]) if (!seen.has(c)) seen.set(c, c);
    return [...seen.entries()];   // [category, label]
  }, [statusConfig, scTone]);
  const allStatuses = useMemo(() => {
    const m = new Map();
    const get = (name) => {
      if (!m.has(name)) m.set(name, { name, nsCount: 0, webosCount: 0, gp: 0, tone: null });
      return m.get(name);
    };
    if (scName && scTone) {
      for (const r of statusConfig.rows) {
        const nm = String(r[scName] || "").trim();
        if (nm) get(nm).tone = r[scTone];
      }
    }
    records.forEach((r) => {
      const st = String(r.nsStatus || "").trim();
      const e = get(st);
      e.nsCount++;
      e.gp += r.expected || 0;
    });
    (webosStatuses.list || []).forEach((w) => { get(w.name).webosCount = w.n; });
    return [...m.values()].sort((a, b) => (b.nsCount + b.webosCount) - (a.nsCount + a.webosCount));
  }, [records, webosStatuses, statusConfig]);

  const agentNames = useMemo(() => {
    const s = new Set(staffNames);
    records.forEach((r) => { if (r.agent) s.add(r.agent); });
    return [...s].sort((a, b) => {
      const ma = managers[a] || "zzz", mb = managers[b] || "zzz";
      return ma.localeCompare(mb) || a.localeCompare(b);
    });
  }, [records, staffNames.join(), managers]);

  const risk = settings.risk || {};
  const suggestedRisk = (st) => (isExcludedStatus(st) ? "high" : "");
  const payplans = settings.payplans || {};
  const dbPlans = staff.plans || {};
  const cfg = settings.payplanSource || {};
  const cfgCols = (probe.found.find((f) => f.table === cfg.table) || {}).cols || [];
  const selStyle = { padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 };

  return (
    <>
      <div className="tabs" style={{ marginBottom: 14 }}>
        {SUBS.map(([k, l]) => (
          <button key={k} className={"tab " + (sub === k ? "active" : "")} onClick={() => setSub(k)}>{l}</button>
        ))}
        {settingsSaving && <span className="sub" style={{ alignSelf: "center", marginLeft: 8 }}>Saving…</span>}
      </div>

      {sub === "risk" && (
        <div className="panel">
          <h2>Order statuses</h2>
          <p className="note" style={{ marginTop: 0 }}>
            <strong>Risk %</strong> is how likely that GP is to fall away — it drives the Top Risk Orders tab.
            Leave it blank to use the default for the category.
            Statuses marked <span className="chip risk">Red GP</span> are <strong>not counted</strong> towards GP or SOV —
            the same rule SchThrive WebOS uses. This drives the Cobra Dashboard and the Net figures on Cross-Reference.
            {statusConfig.status === "ok"
              ? ` Treatment comes from the status_config table's "${scTone}" column; change one here to override it.`
              : statusConfig.status === "error" ? " (Couldn't read status_config — using the built-in rules.)" : ""}
          </p>
          <div style={{ marginBottom: 12 }}>
            <button className="btn" onClick={() => {
              saveSettings({ ...settings, risk: {} });
            }}>Reset to status_config</button>
            <span className="sub" style={{ marginLeft: 10 }}>
              Clears any manual overrides so every status follows the database tone again.
            </span>
          </div>
          {allStatuses.length === 0 ? (
            <div className="empty">No order statuses found yet — load the NetSuite export first.</div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead><tr>
                  <th className="left">Order status</th>
                  <th className="num">In NetSuite</th>
                  <th className="num">In SchThrive</th>
                  <th className="num">Tone</th>
                  <th className="num">GP in NetSuite</th>
                  <th className="num">Risk %</th>
                  <th className="num">Payable?</th>
                </tr></thead>
                <tbody>
                  {allStatuses.map((st) => {
                    const setting = risk[st.name] || toneToTreatment(st.tone) || (isExcludedStatus(st.name) ? "red" : "counts");
                    return (
                      <tr key={st.name || "(blank)"} style={{ background: setting === "red" ? "#fdf3f2" : undefined }}>
                        <td className="left">{st.name || <em className="sub">(blank)</em>}</td>
                        <td className="num mono">{st.nsCount || "-"}</td>
                        <td className="num mono">{st.webosCount || "-"}</td>
                        <td className="num">{st.tone ? <span className="chip" style={CAT_STYLE[toneCategory(st.tone)] || {}}>{String(st.tone)}</span> : "-"}</td>
                        <td className="num mono">{st.gp ? gbp(st.gp) : "-"}</td>
                        <td className="num">
                          <NumInput
                            value={(settings.riskPct || {})[st.name] ?? ""}
                            placeholder={String(DEFAULT_RISK_PCT[setting] ?? 0)}
                            onCommit={(nv) => saveSettings({ ...settings,
                              riskPct: { ...(settings.riskPct || {}), [st.name]: nv === "" ? "" : Number(nv) } })}
                            style={{ width: 62, padding: "4px 6px", border: "1px solid #d3d0e6",
                              borderRadius: 6, fontSize: 12, textAlign: "center" }} />
                        </td>
                        <td className="num">
                          <select value={setting}
                            onChange={(e) => saveSettings({ ...settings, risk: { ...risk, [st.name]: e.target.value } })}
                            style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13,
                              color: setting === "red" ? "#b3261e" : "#14804a", fontWeight: 600 }}>
                            {toneOptions.map(([cat, label]) => (
                              <option key={cat} value={cat}>
                                {label}{cat === "red" ? " (ignored)" : ""}
                              </option>
                            ))}
                          </select>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {sub === "payplans" && (
        <div className="panel">
          <h2>Payplans — who is on what, month by month</h2>
          <p className="note" style={{ marginTop: 0 }}>
            The commission each agent had to hit that month. Any month below this is flagged red on Payments Per Agent.
            Leave a month blank to fall back to the agent's standard payplan.
          </p>
          <div className="panel" style={{ background: "#faf9ff", marginBottom: 14 }}>
            <h2 style={{ marginBottom: 8 }}>Where payplans come from</h2>
            {!cfg.table ? (
              <p className="note" style={{ marginTop: 0 }}>
                Not linked to the database yet — payplans typed below are used instead.
                Click <strong>Find payplan tables</strong> and I'll check the database for one.
              </p>
            ) : (
              <p className="note" style={{ marginTop: 0 }}>
                Reading from <span className="mono">{cfg.table}</span> — name <span className="mono">{cfg.nameCol}</span>,
                amount <span className="mono">{cfg.valueCol}</span>
                {cfg.monthCol ? <> , month <span className="mono">{cfg.monthCol}</span></> : " (no month column — one standing plan per agent)"}.
                {" "}{Object.keys(staff.plansByMonth || {}).length
                  ? `${Object.keys(staff.plansByMonth).length} agents have month-by-month targets.`
                  : Object.keys(dbPlans).length ? `${Object.keys(dbPlans).length} standing payplans loaded.` : "No rows matched yet."}
              </p>
            )}
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <button className="btn" onClick={probeTables}>
                {probe.status === "loading" ? "Checking…" : "Find payplan tables"}
              </button>
              {probe.status === "done" && probe.found.length === 0 && (
                <span className="sub" style={{ margin: 0, color: "#b3261e" }}>
                  No matching tables found. Tell me the table name and I'll add it.
                </span>
              )}
              {probe.status === "done" && probe.found.length > 0 && (
                <>
                  <select value={cfg.table || ""} style={selStyle}
                    onChange={(e) => saveSettings({ ...settings, payplanSource: { table: e.target.value, nameCol: "", valueCol: "" } })}>
                    <option value="">Choose a table…</option>
                    {probe.found.map((f) => <option key={f.table} value={f.table}>{f.table} ({f.cols.length} columns)</option>)}
                  </select>
                  {cfgCols.length > 0 && (
                    <>
                      <span>Agent name:</span>
                      <select value={cfg.nameCol || ""} style={selStyle}
                        onChange={(e) => saveSettings({ ...settings, payplanSource: { ...cfg, nameCol: e.target.value } })}>
                        <option value="">…</option>
                        {cfgCols.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <span>Payplan amount:</span>
                      <select value={cfg.valueCol || ""} style={selStyle}
                        onChange={(e) => saveSettings({ ...settings, payplanSource: { ...cfg, valueCol: e.target.value } })}>
                        <option value="">…</option>
                        {cfgCols.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <span>Month (optional):</span>
                      <select value={cfg.monthCol || ""} style={selStyle}
                        onChange={(e) => saveSettings({ ...settings, payplanSource: { ...cfg, monthCol: e.target.value } })}>
                        <option value="">(no month — one standing plan)</option>
                        {cfgCols.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <span>Plan name (optional):</span>
                      <select value={cfg.planNameCol || ""} style={selStyle}
                        onChange={(e) => saveSettings({ ...settings, payplanSource: { ...cfg, planNameCol: e.target.value } })}>
                        <option value="">(none)</option>
                        {cfgCols.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <button className="btn" onClick={loadStaff}>Load payplans</button>
                    </>
                  )}
                </>
              )}
            </div>
            {probe.status === "done" && probe.found.length > 0 && (
              <p className="note">
                Found: {probe.found.map((f) => f.table).join(", ")}. Columns in the chosen table:{" "}
                <span className="mono">{cfgCols.join(", ") || "(pick a table)"}</span>
              </p>
            )}
          </div>

          <div className="settings" style={{ marginBottom: 12 }}>
            <span>Financial year:</span>
            <select value={planYear} onChange={(e) => setPlanYear(Number(e.target.value))}
              style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
              {planYears.map((y) => <option key={y} value={y}>{fyLabel(y)}</option>)}
            </select>
            <button className="btn" onClick={loadStaff}>Reload agents</button>
            <span className="sub" style={{ margin: 0 }}>
              {staff.status === "ok"
                ? `${staffNames.length} agents${staff.excluded ? `, ${staff.excluded} excluded (Tracy Webber's team)` : ""}${Object.keys(staff.plans || {}).length ? " · standard payplans loaded from the database" : " · no payplans in the database, so set them here"}`
                : staff.status === "error" ? "Couldn't read the staff table — showing agents found in NetSuite."
                  : "Loading the staff list…"}
            </span>
          </div>
          {agentNames.length === 0 ? (
            <div className="empty">No agents found yet.</div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th className="left">Agent</th>
                    <th className="left">Manager</th>
                    <th className="num">Standard</th>
                    {MONTHS_FY.map(([, l]) => <th key={l} className="num">{l}</th>)}
                    <th className="num">Year total</th>
                  </tr>
                </thead>
                <tbody>
                  {agentNames.map((a2) => {
                    const std = payplans[a2] ?? dbPlans[a2] ?? "";
                    const months = (settings.payplanMonthly || {})[planYear]?.[a2] || [];
                    const effective = MONTHS_FY.map((_, i) => {
                      const v = months[i];
                      if (v !== "" && v != null) return Number(v);
                      const fromDb = (staff.plansByMonth || {})[a2]?.[fyMonthKeys(planYear)[i]];
                      if (fromDb != null) return Number(fromDb);
                      return std === "" ? null : Number(std);
                    });
                    return (
                      <tr key={a2}>
                        <td className="left">{a2}</td>
                        <td className="left sub">{managers[a2] || "-"}</td>
                        <td className="num">
                          <NumInput value={payplans[a2] ?? ""} placeholder={dbPlans[a2] ? String(dbPlans[a2]) : "-"}
                            onCommit={(nv) => saveSettings({ ...settings, payplans: { ...payplans, [a2]: nv === "" ? "" : Number(nv) } })}
                            style={{ width: 90, padding: "4px 6px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 12, textAlign: "center" }} />
                        </td>
                        {MONTHS_FY.map((_, i) => {
                          const mk = fyMonthKeys(planYear)[i];
                          const fromDb = (staff.plansByMonth || {})[a2]?.[mk];
                          const planNm = (staff.planNames || {})[a2]?.[mk];
                          return (
                          <td key={i} className="num" title={planNm ? `Plan: ${planNm}` : undefined}>
                            <NumInput value={months[i] ?? ""}
                              placeholder={fromDb != null ? String(fromDb) : (std === "" ? "-" : String(std))}
                              onCommit={(nv) => {
                                const next = JSON.parse(JSON.stringify(settings.payplanMonthly || {}));
                                next[planYear] = next[planYear] || {};
                                next[planYear][a2] = next[planYear][a2] || [];
                                next[planYear][a2][i] = nv === "" ? "" : Number(nv);
                                saveSettings({ ...settings, payplanMonthly: next });
                              }}
                              style={{ width: 74, padding: "4px 5px", border: "1px solid #e2e0ee", borderRadius: 5,
                                fontSize: 12, textAlign: "center", background: months[i] == null || months[i] === "" ? "#faf9ff" : "#fff" }} />
                            {planNm && <div className="sub" style={{ margin: 0, fontSize: 9.5 }}>{planNm}</div>}
                          </td>
                          );
                        })}
                        <td className="num mono"><strong>{gbp(sum(effective.filter((v) => v != null)))}</strong></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {sub === "monthly" && (
        <MonthlyReports supabase={supabase} session={session} files={files} statusRules={statusRules} />
      )}

      {sub === "rawdata" && (
        <RawData files={files} errors={errors} onFile={onFile} supabase={supabase}
          session={session} saving={saving} sharedMeta={sharedMeta} saveShared={saveShared} anyLoaded={anyLoaded} saveProgress={saveProgress} />
      )}

      {sub === "users" && isAdmin && (
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
          {userMsg && <p className="sub" style={{ color: userMsg.err ? "#b3261e" : "#14804a", marginTop: 0 }}>{userMsg.text}</p>}
          <table>
            <thead><tr><th className="left">Email</th><th className="num">Role</th><th className="num"></th></tr></thead>
            <tbody>
              {users.map((u) => (
                <tr key={u.email}>
                  <td className="left mono">{u.email}</td>
                  <td className="num">{u.is_admin ? <span className="chip matched">admin</span> : <span className="chip unmatched">viewer</span>}</td>
                  <td className="num">
                    {(u.email || "").toLowerCase() !== (session?.user?.email || "").toLowerCase() && (
                      <button className="btn" onClick={() => removeUser(u.email)}>Remove</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="note">Removing someone stops them seeing the data. Password resets are done from Supabase.</p>
        </div>
      )}
    </>
  );
}

// =========================================================================
//  Overview — the landing page
// =========================================================================
function Overview({ records, files, settings, snapshot, saveSnapshot, setTab, supabase, session, statusRules = {} , sources }) {
  const { ns, cb, s5 } = useSourceLines(files, sources);

  // financial years present across all three sources
  const chartYears = useMemo(() => {
    const set = new Set();
    for (const r of [...ns, ...cb, ...s5]) { const f = fyStartOf(r.period); if (f) set.add(f); }
    for (const r of records) { const f = fyStartOf(r.period); if (f) set.add(f); }
    return [...set].sort((a, b) => b - a);
  }, [ns, cb, s5, records]);
  const [chartFy, setChartFy] = useState(null);
  const cYear = chartFy ?? chartYears[0] ?? null;

  // every figure on this page is for the selected financial year only
  const recs = useMemo(
    () => (cYear == null ? records : records.filter((r) => fyStartOf(r.period) === cYear)),
    [records, cYear]
  );

  // ---- headline position ----
  const k = useMemo(() => {
    let nsGp = 0, cobraUnmatched = 0, sch5Unmatched = 0;
    let nMatched = 0, nReview = 0, nException = 0;
    for (const r of recs) {
      nsGp += r.expected || 0;
      if (r.inCobra && !r.inNS) cobraUnmatched += r.paid || 0;   // C but not N
      if (r.inSch5 && !r.inNS) sch5Unmatched += r.s5Sov || 0;    // S but not N
      // NetSuite vs Cobra: does what NetSuite expects agree with what Cobra paid?
      const hasNs = r.inNS, hasCb = r.inCobra;
      if (!hasNs || !hasCb) { nException++; continue; }      // only on one side
      const d = (r.paid || 0) - (r.expected || 0);
      if (Math.abs(d) < 0.005) nMatched++;
      else nException++;
    }
    const nTotal = Math.max(1, nMatched + nReview + nException);
    const cobraPaid = sum(cb.filter((r) => cYear == null || fyStartOf(r.period) === cYear).map((r) => r.paid));
    return {
      nsGp, cobraPaid, underpaid: nsGp - cobraPaid, cobraUnmatched, sch5Unmatched,
      pctMatched: (nMatched / nTotal) * 100,
      pctReview: (nReview / nTotal) * 100,
      pctException: (nException / nTotal) * 100,
      nMatched, nReview, nException,
    };
  }, [recs, cb, cYear]);

  // ---- monthly series, for both GP and SOV ----
  const [mode, setMode] = useState("gp");

  // always 12 buckets: April through March
  const monthly = useMemo(() => {
    if (cYear == null) return [];
    const keys = fyMonthKeys(cYear);
    const idx = new Map(keys.map((kk, i) => [kk, i]));
    const rows = keys.map((p) => ({ p, nsGp: 0, cbPaid: 0, nsSov: 0, s5Sov: 0, cbSov: 0 }));
    for (const r of ns) {
      const i = idx.get(r.period); if (i == null) continue;
      rows[i].nsGp += r.expected || 0;
      rows[i].nsSov += r.sov || 0;                 // Sales Closer rows only
    }
    for (const r of cb) {
      const i = idx.get(r.period); if (i == null) continue;
      rows[i].cbPaid += r.paid || 0;               // Commission Paid
      rows[i].cbSov += r.sov || 0;                 // Contract Value
    }
    for (const r of s5) {
      const i = idx.get(r.period); if (i == null) continue;
      if (!isYes(r.orderFlag)) continue;           // only rows flagged Y in "Sch5 Order"
      rows[i].s5Sov += r.sov || 0;
    }
    return rows;
  }, [ns, cb, s5, cYear]);

  const flagCol = useMemo(() => s5.find((r) => r.orderFlagCol)?.orderFlagCol || null, [s5]);

  const SERIES = mode === "gp"
    ? [["nsGp", "#1e64d6", "Total GP (NetSuite)"], ["cbPaid", "#8b5cf6", "Paid on Cobra"]]
    : [["nsSov", "#1e64d6", "SOV (NetSuite)"], ["s5Sov", "#12a594", "SOV (Sch5)"], ["cbSov", "#8b5cf6", "Contract Value (Cobra)"]];
  const maxY = Math.max(1, ...monthly.flatMap((d) => SERIES.map(([kk]) => d[kk])));
  const niceMax = Math.ceil(maxY / 50000) * 50000 || maxY;

  // ---- orders that have moved since the month-end (payroll) snapshots ----
  const [statics, setStatics] = useState({});
  useEffect(() => {
    if (!supabase || !session || cYear == null) return;
    let cancelled = false;
    (async () => {
      const mks = fyMonthKeys(cYear);
      const { data } = await supabase.from("recon_datasets").select("*").in("id", mks.map((mk) => `static-${mk}`));
      if (cancelled || !data) return;
      const out = {};
      for (const r of data) {
        if (cancelled) return;
        const mk = String(r.id).replace("static-", "");
        const full = await hydrate(supabase, r.netsuite, `chunk-${r.id}`);
        if (full?.rows) out[mk] = full.rows.map(nsRow);
      }
      setStatics(out);
    })();
    return () => { cancelled = true; };
  }, [session, cYear]);

  const [flagFilter, setFlagFilter] = useState("all");
  const flagged = useMemo(() => {
    if (cYear == null) return [];
  
    const catOf = (r) => (isNonCommissionable(r) ? "red" : statusCategory(statusRules, r.status));
    const out = [];
    for (const mk of fyMonthKeys(cYear)) {
      const snapRows = statics[mk];
      if (!snapRows) continue;
      const snap = new Map();
      for (const r of snapRows) {
        if (!isRealLine(r)) continue;
        const k = matchKey(r);
        if (!k) continue;
        if (!snap.has(k)) snap.set(k, { ...r, gp: 0 });
        snap.get(k).gp += r.expected || 0;
      }
      const live = new Map();
      for (const r of ns) {
        if (r.period !== mk || !isRealLine(r)) continue;
        const k = matchKey(r);
        if (!k) continue;
        if (!live.has(k)) live.set(k, { ...r, gp: 0 });
        live.get(k).gp += r.expected || 0;
      }
      for (const [k, b] of live) {
        const a = snap.get(k);
        if (!a) {
          out.push({ ...b, mk, cat: catOf(b), wasCat: null, wasStatus: "(not in payroll)", was: null, now: b.gp,
            reason: "New since payroll", change: "new", valueChanged: false, becamePayable: false, lostPayable: false });
          continue;
        }
        const wasCat = catOf(a), nowCat = catOf(b);
        const moved = Math.abs((a.gp || 0) - (b.gp || 0)) > 0.005;
        const statusMoved = String(a.status || "") !== String(b.status || "");
        if (!moved && wasCat === nowCat && !statusMoved) continue;
        const becamePayable = wasCat !== "payable" && nowCat === "payable";
        const lostPayable = wasCat === "payable" && nowCat !== "payable";
        const reason = moved && (wasCat !== nowCat || statusMoved) ? "Value and status changed"
          : moved ? ((b.gp || 0) > (a.gp || 0) ? "Value increased" : "Value reduced")
            : nowCat === "red" ? "Moved into a red status"
              : nowCat === "payable" && wasCat !== nowCat ? "Became payable"
                : wasCat !== nowCat ? `Moved to ${nowCat}` : "Status changed";
        out.push({ ...b, mk, cat: nowCat, wasCat, wasStatus: a.status, was: a.gp, now: b.gp, reason,
          change: moved ? "value" : "status", valueChanged: moved, becamePayable, lostPayable });
      }
      for (const [k, a] of snap) {
        if (!live.has(k)) {
          out.push({ ...a, mk, wasCat: catOf(a), wasStatus: a.status, cat: null, status: "(not in live report)",
            was: a.gp, now: 0, reason: "Removed since payroll", change: "removed", valueChanged: false,
            becamePayable: false, lostPayable: false });
        }
      }
    }
    const rank = (r) => (r.reason.includes("Removed") ? 0 : r.reason.includes("red") ? 1 : 2);
    return out.sort((a, b) => rank(a) - rank(b) || Math.abs((b.now - (b.was || 0))) - Math.abs((a.now - (a.was || 0))));
  }, [statics, ns, cYear, statusRules]);

  const shownFlagged = useMemo(
    () => flagged.filter((r) => matchesChangeFilter(r, flagFilter)), [flagged, flagFilter]);

  // ---- exceptions needing action ----
  // GP split by tone category, for the snapshot comparison
  const catTotals = useMemo(() => {
    let total = 0, payable = 0, issued = 0;
    for (const r of recs) {
      const gp = r.expected || 0;
      const cat = statusCategory(statusRules, r.nsStatus);
      if (cat === "red") continue;
      total += gp;
      if (cat === "payable") payable += gp;
      if (cat === "issued") issued += gp;
    }
    return { total, payable, issued };
  }, [recs, statusRules]);

  // ---- commission pay control ----
  const pay = useMemo(() => {
    const btPaidUs = sum(recs.map((r) => r.paid));
    return { btPaidUs };
  }, [recs]);

  const changes = useMemo(() => {
    if (!snapshot) return null;
    return {
      total: catTotals.total - (snapshot.total ?? snapshot.received ?? 0),
      payable: catTotals.payable - (snapshot.payable ?? 0),
      issued: catTotals.issued - (snapshot.issued ?? 0),
    };
  }, [snapshot, catTotals]);

  const W = 640, H = 230, padL = 46, padB = 26, padT = 8;
  const bw = monthly.length ? (W - padL - 8) / monthly.length : 0;
  const y = (v) => padT + (H - padT - padB) * (1 - v / niceMax);

  const Card = ({ colour, label, value, sub2, subColour }) => (
    <div className="kpic" style={{ borderTopColor: colour }}>
      <div className="lab">{label}</div>
      <div className="val" style={{ color: colour }}>{value}</div>
      {sub2 && <div className="sub2" style={{ color: subColour || colour }}>{sub2}</div>}
    </div>
  );

  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 10, flexWrap: "wrap" }}>
        <div className="settings" style={{ margin: 0 }}>
          <span style={{ fontWeight: 600 }}>Financial year:</span>
          <select value={cYear ?? ""} onChange={(e) => setChartFy(Number(e.target.value))}
            style={{ padding: "7px 10px", border: "1px solid #d3d0e6", borderRadius: 7, fontSize: 13.5, fontWeight: 600 }}>
            {chartYears.map((yy) => <option key={yy} value={yy}>{fyLabel(yy)}</option>)}
          </select>
          <span className="sub" style={{ margin: 0 }}>
            {cYear != null ? `Every figure on this page covers ${fyLabel(cYear)} only.` : "No dated rows yet."}
          </span>
        </div>
        <button className="btn" style={{ background: "#1e64d6", color: "#fff", padding: "9px 14px" }}
          onClick={() => saveSnapshot({
            at: new Date().toISOString(), total: catTotals.total, payable: catTotals.payable,
            issued: catTotals.issued, nException: k.nException,
          })}>
          {snapshot?.at ? `Re-lock snapshot (last ${new Date(snapshot.at).toLocaleDateString("en-GB")})` : "Lock snapshot"}
        </button>
      </div>

      <div className="kpis5">
        <Card colour="#1e64d6" label="Total GP from NetSuite" value={gbp0(k.nsGp)}
          sub2="Sales Agent GP, all rows" subColour="#8a8aa3" />
        <Card colour="#14804a" label="Matched" value={gbp0(k.cobraPaid)}
          sub2={`Paid on Cobra · ${k.nsGp ? ((k.cobraPaid / k.nsGp) * 100).toFixed(1) : "0.0"}%`} subColour="#8a8aa3" />
        <Card colour="#d98a00" label="Underpaid" value={gbp0(k.underpaid)}
          sub2="GP not yet paid by Cobra" subColour="#8a8aa3" />
        <Card colour="#b3261e" label="Overpaid" value="—"
          sub2="needs a database change" subColour="#8a8aa3" />
        <div className="kpic" style={{ borderTopColor: "#5b6472" }}>
          <div className="lab">Unallocated</div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginTop: 4 }}>
            <div>
              <div className="val" style={{ color: "#5b6472", fontSize: 17 }}>{gbp0(k.cobraUnmatched)}</div>
              <div className="sub2" style={{ color: "#8a8aa3", fontWeight: 500 }}>Cobra GP, no NetSuite</div>
            </div>
            <div>
              <div className="val" style={{ color: "#5b6472", fontSize: 17 }}>{gbp0(k.sch5Unmatched)}</div>
              <div className="sub2" style={{ color: "#8a8aa3", fontWeight: 500 }}>Sch5 SOV, no NetSuite</div>
            </div>
          </div>
        </div>
      </div>

      <div className="two" style={{ marginTop: 14 }}>
        <div className="panel">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <h2 style={{ margin: 0 }}>
              Monthly {mode === "gp" ? "GP" : "SOV"} position{cYear != null ? ` — ${fyLabel(cYear)}` : ""}
            </h2>
            <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
              {[["gp", "GP"], ["sov", "SOV"]].map(([v, l]) => (
                <button key={v} className={"tab " + (mode === v ? "active" : "")}
                  style={{ padding: "5px 14px", fontSize: 12.5 }} onClick={() => setMode(v)}>{l}</button>
              ))}
            </div>
          </div>
          <div className="legend" style={{ marginTop: 10 }}>
            {SERIES.map(([kk, c, l]) => <span key={kk}><i style={{ background: c }} />{l}</span>)}
          </div>
          {monthly.length === 0 ? <div className="empty">No dated rows yet.</div> : (
            <div style={{ overflowX: "auto" }}>
              <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", minWidth: 520, height: 240 }}>
                {[0, 0.25, 0.5, 0.75, 1].map((f) => (
                  <g key={f}>
                    <line x1={padL} x2={W - 4} y1={y(niceMax * f)} y2={y(niceMax * f)} stroke="#eceaf4" />
                    <text x={padL - 6} y={y(niceMax * f) + 3} textAnchor="end" fontSize="9" fill="#8a8aa3">
                      {Math.round((niceMax * f) / 1000)}K
                    </text>
                  </g>
                ))}
                {monthly.map((d, i) => (
                  <g key={d.p}>
                    {SERIES.map(([kk, c], j) => {
                      const n = SERIES.length, slot = (bw - 8) / n;
                      const h = Math.max(0, y(0) - y(d[kk]));
                      return <rect key={kk} x={padL + i * bw + 4 + j * slot} y={y(d[kk])}
                        width={Math.max(2, slot - 2)} height={h} fill={c} rx="1.5" />;
                    })}
                    <text x={padL + i * bw + bw / 2} y={H - 8} textAnchor="middle" fontSize="9" fill="#8a8aa3">
                      {MONTHS_FY[i][1]}
                    </text>
                  </g>
                ))}
              </svg>
            </div>
          )}
          {mode === "sov" && (
            <p className="note">
              Sch5 SOV counts only rows flagged Y in{" "}
              <span className="mono">{flagCol || "(no flag column found)"}</span>
              {flagCol && flagCol !== "SCH5 ORDER" ? " — no \"Sch5 Order\" column in this export, so this is the closest match." : ""}
            </p>
          )}
        </div>

        <div className="panel">
          <h2>NetSuite vs Cobra</h2>
          <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
            <Donut matched={k.pctMatched} review={0} exception={k.pctException} />
            <div style={{ fontSize: 13 }}>
              {[["Matched", k.pctMatched, "#14804a", k.nMatched],
                ["Mismatch", k.pctException, "#b3261e", k.nException]].map(([l, v, c, n]) => (
                <div key={l} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <i style={{ width: 10, height: 10, borderRadius: "50%", background: c, display: "inline-block" }} />
                  <span style={{ minWidth: 78 }}>{l}</span>
                  <strong className="mono">{v.toFixed(1)}%</strong>
                  <span className="sub" style={{ margin: 0 }}>({n})</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      <div className="two" style={{ marginTop: 14 }}>
        <div className="panel">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <h2 style={{ margin: 0 }}>Flagged orders from payroll updates <span className="sub">({flagged.length})</span></h2>
            {flagged.length > 0 && (
              <button className="btn" onClick={() => downloadCSV(flagged.map((r) => ({
                Month: periodLabel(r.mk), "Order ref": r.orderNum || "", "NetSuite ref": r.netsuiteRef || "",
                Company: r.company || "", Partner: r.partner || "", Item: r.product || "",
                "Status was": r.wasStatus || "", "Status now": r.status || "",
                "Category was": r.wasCat || "", "Category now": r.cat || "",
                "GP was": r.was ?? "", "GP now": r.now ?? "", Reason: r.reason,
              })), "flagged-orders.csv")}>Export</button>
            )}
          </div>
          {flagged.length > 0 && <ChangeFilters rows={flagged} value={flagFilter} onChange={setFlagFilter} />}
          {flagged.length > 0 && (() => {
            const t = { was: 0, now: 0 };
            for (const r of shownFlagged) { t.was += r.was || 0; t.now += r.now || 0; }
            const counts = shownFlagged.reduce((m, r) => { m[r.reason] = (m[r.reason] || 0) + 1; return m; }, {});
            return (
              <div style={{ overflowX: "auto", marginTop: 10, marginBottom: 4 }}>
                <table>
                  <thead><tr>
                    <th className="left">Flagged total</th>
                    <th className="num">GP at payroll</th><th className="num">GP now</th><th className="num">Difference</th>
                  </tr></thead>
                  <tbody>
                    <tr>
                      <td className="left">{shownFlagged.length} order{shownFlagged.length === 1 ? "" : "s"}</td>
                      <td className="num mono">{gbp0(t.was)}</td>
                      <td className="num mono">{gbp0(t.now)}</td>
                      <td className={"num mono " + (t.now - t.was >= 0 ? "pos" : "neg")}>
                        {t.now - t.was > 0 ? "+" : ""}{gbp0(t.now - t.was)}
                      </td>
                    </tr>
                  </tbody>
                </table>
                <p className="note" style={{ marginTop: 6 }}>
                  {Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([r, n]) => `${r}: ${n}`).join(" · ")}
                </p>
              </div>
            );
          })()}
          {flagged.length === 0 ? (
            <div className="empty">
              {Object.keys(statics).length === 0
                ? <>No month-end snapshots uploaded for this year yet — add them under <strong>Settings → Monthly Reports</strong> to see what has moved since payroll.</>
                : "Nothing has changed since the payroll snapshots."}
            </div>
          ) : (
            <div style={{ overflowX: "auto", marginTop: 10, maxHeight: 420, overflowY: "auto" }}>
              <table>
                <thead><tr>
                  <th className="left">Order</th><th className="left">Month</th>
                  <th className="left">Status change</th><th className="num">GP was</th>
                  <th className="num">GP now</th><th className="num">Change</th><th className="left">Reason</th>
                </tr></thead>
                <tbody>
                  {shownFlagged.slice(0, 100).map((r, i) => {
                    const delta = (r.now || 0) - (r.was || 0);
                    return (
                      <tr key={i} style={{ background: r.reason.includes("Removed") ? "#f7c8c2"
                        : r.reason.includes("red") ? "#fbd5cf"
                          : r.reason.includes("status") ? "#fde2dd"
                            : r.reason.includes("Value") ? "#fdecea" : undefined }}>
                        <td className="left">
                          <div className="mono">{r.orderNum || r.netsuiteRef || "-"}</div>
                          <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.company || "-"}</div>
                          <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.partner || "-"}</div>
                        </td>
                        <td className="left">{periodLabel(r.mk)}</td>
                        <td className="left">
                          {(r.wasStatus != null && String(r.wasStatus) !== String(r.status)) || (r.wasCat && r.wasCat !== r.cat) ? (
                            <>
                              <div style={{ fontSize: 10.5, color: "#8a8aa3" }}>was</div>
                              <div>
                                <span className="chip" style={CAT_STYLE[r.wasCat] || {}}>{r.wasCat || "-"}</span>{" "}
                                <span style={{ fontSize: 11 }}>{r.wasStatus || "-"}</span>
                              </div>
                              <div style={{ fontSize: 10.5, color: "#8a8aa3", marginTop: 3 }}>now</div>
                              <div>
                                <span className="chip" style={CAT_STYLE[r.cat] || {}}>{r.cat || "-"}</span>{" "}
                                <span style={{ fontSize: 11 }}>{r.status || "-"}</span>
                              </div>
                            </>
                          ) : (
                            <>
                              <span className="chip" style={CAT_STYLE[r.cat] || {}}>{r.cat}</span>
                              <div className="sub" style={{ margin: 0, fontSize: 10.5 }}>{r.status || "-"}</div>
                            </>
                          )}
                        </td>
                        <td className="num mono">{r.was == null ? "-" : gbp0(r.was)}</td>
                        <td className="num mono">{gbp0(r.now)}</td>
                        <td className={"num mono " + (delta >= 0 ? "pos" : "neg")}>{delta > 0 ? "+" : ""}{gbp0(delta)}</td>
                        <td className="left">{r.reason}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {shownFlagged.length > 100 && <p className="note">Showing 100 of {shownFlagged.length} — use Export for the full list.</p>}
            </div>
          )}
        </div>

        <div>
          <div className="panel">
            <h2>Commission pay control</h2>
            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginTop: 8 }}>
              <div className="step"><span className="n">1</span>
                <div><div className="sub" style={{ margin: 0 }}>Total payable GP</div>
                  <strong className="mono" style={{ fontSize: 16 }}>{gbp0(catTotals.payable)}</strong></div>
              </div>
              <span style={{ color: "#c9c6da" }}>›</span>
              <div className="step"><span className="n">2</span>
                <div><div className="sub" style={{ margin: 0 }}>Total paid from BT</div>
                  <strong className="mono" style={{ fontSize: 16 }}>{gbp0(pay.btPaidUs)}</strong></div>
              </div>
              <span style={{ color: "#c9c6da" }}>›</span>
              <div className="step"><span className="n">3</span>
                <div><div className="sub" style={{ margin: 0 }}>Remaining to be paid</div>
                  <strong className="mono" style={{ fontSize: 16, color: catTotals.payable - pay.btPaidUs > 0 ? "#b3261e" : "#14804a" }}>
                    {gbp0(catTotals.payable - pay.btPaidUs)}
                  </strong></div>
              </div>
            </div>
            <p className="note">
              Payable GP uses the tone categories; paid comes from Cobra's Commission Paid.
            </p>
          </div>

          <div className="panel">
            <h2>Changes since snapshot</h2>
            {!snapshot ? (
              <div className="empty" style={{ padding: 20 }}>No snapshot locked yet. Lock one to start tracking movement.</div>
            ) : (
              <table>
                <tbody>
                  {[["Total GP Change", changes.total], ["Total Payable Change", changes.payable],
                    ["Total Issued Change", changes.issued]].map(([l, v]) => (
                    <tr key={l}>
                      <td className="left">{l}</td>
                      <td className={"num mono " + (v === 0 ? "" : v > 0 ? "pos" : "neg")}>
                        {v > 0 ? "+" : ""}{gbp0(v)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
            {snapshot && <p className="note">Locked {new Date(snapshot.at).toLocaleString("en-GB")}.</p>}
          </div>
        </div>
      </div>
    </>
  );
}

function Donut({ matched, review, exception }) {
  const R = 54, r = 34, C = 2 * Math.PI * ((R + r) / 2), sw = R - r;
  const segs = [[matched, "#14804a"], [review, "#d98a00"], [exception, "#b3261e"]];
  let offset = 0;
  return (
    <svg viewBox="0 0 120 120" style={{ width: 130, height: 130 }}>
      <circle cx="60" cy="60" r={(R + r) / 2} fill="none" stroke="#eceaf4" strokeWidth={sw} />
      {segs.map(([v, c], i) => {
        const len = (Math.max(0, v) / 100) * C;
        const el = <circle key={i} cx="60" cy="60" r={(R + r) / 2} fill="none" stroke={c} strokeWidth={sw}
          strokeDasharray={`${len} ${C - len}`} strokeDashoffset={-offset} transform="rotate(-90 60 60)" />;
        offset += len;
        return el;
      })}
      <text x="60" y="58" textAnchor="middle" fontSize="17" fontWeight="800" fill="#1b1636">
        {matched.toFixed(0)}%
      </text>
      <text x="60" y="72" textAnchor="middle" fontSize="8.5" fill="#8a8aa3">matched</text>
    </svg>
  );
}

// =========================================================================
//  Monthly Reports — a static NetSuite snapshot per month, compared to live
// =========================================================================
function MonthlyReports({ supabase, session, files, statusRules }) {
  const thisFy = fyStartOf(`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}`);
  const [year, setYear] = useState(thisFy);
  const [saved, setSaved] = useState({});          // monthKey -> { rows, name, at, by }
  const [busy, setBusy] = useState(null);
  const [msg, setMsg] = useState(null);
  const [compare, setCompare] = useState(null);    // monthKey being compared

  const keys = fyMonthKeys(year);
  const rowId = (mk) => `static-${mk}`;

  // load whatever snapshots already exist for this year
  useEffect(() => {
    if (!supabase || !session) return;
    let cancelled = false;
    (async () => {
      const { data } = await supabase.from("recon_datasets").select("*").in("id", keys.map(rowId));
      if (cancelled || !data) return;
      const out = {};
      for (const r of data) {
        const mk = String(r.id).replace("static-", "");
        if (r.netsuite) out[mk] = { ...r.netsuite, rows: null, at: r.uploaded_at, by: r.uploaded_by, count: r.netsuite.count ?? (r.netsuite.rows || []).length };
      }
      setSaved(out);
    })();
    return () => { cancelled = true; };
  }, [session, year]);

  const upload = useCallback(async (mk, file) => {
    if (!file || !supabase) return;
    setBusy(mk); setMsg(null);
    try {
      const parsed = await parseWorkbook(file);
      const slim = slimFileForSave({ sheet: parsed.sheet, name: file.name, rows: parsed.rows }, "netsuite");
      const packed = packRows(slim.rows);
      const n = await writeChunks(supabase, `chunk-${rowId(mk)}`, packed.d,
        (done, total) => setMsg({ err: false, text: `${periodLabel(mk)} — saving part ${done} of ${total}…` }));
      const payload = {
        id: rowId(mk),
        netsuite: { sheet: slim.sheet, name: slim.name, c: packed.c, n, count: slim.rows.length },
        uploaded_by: session?.user?.email || null, uploaded_at: new Date().toISOString(),
      };
      const { error } = await supabase.from("recon_datasets").upsert(payload, { onConflict: "id" });
      if (error) setMsg({ err: true, text: "Couldn't save: " + error.message });
      else {
        setSaved((sv) => ({ ...sv, [mk]: { ...slim, at: payload.uploaded_at, by: payload.uploaded_by } }));
        setMsg({ err: false, text: `${periodLabel(mk)} snapshot saved (${parsed.rows.length} rows).` });
      }
    } catch (e) {
      setMsg({ err: true, text: "Couldn't read that file — is it a valid CSV or Excel export?" });
    }
    setBusy(null);
  }, [session, keys.join()]);

  const removeSnapshot = useCallback(async (mk) => {
    if (!supabase) return;
    await supabase.from("recon_datasets").delete().like("id", `chunk-${rowId(mk)}-%`);
    await supabase.from("recon_datasets").delete().eq("id", rowId(mk));
    setSaved((sv) => { const n = { ...sv }; delete n[mk]; return n; });
  }, []);

  // ---- compare a saved month against the live report ----
  const diff = useMemo(() => {
    if (!compare || !saved[compare]) return null;
    const key = (r) => `${r.netsuiteRef || ""}|${r.orderNum || ""}|${r.product || ""}|${r.partner || ""}`;
    const build = (rows) => {
      const m = new Map();
      for (const raw of rows || []) {
        const r = nsRow(raw);
        const k = key(r);
        if (!m.has(k)) m.set(k, { ...r, gp: 0, pay: 0 });
        const e = m.get(k);
        e.gp += r.expected || 0;
        if (r.recordedCobra != null) e.pay = r.recordedCobra;
      }
      return m;
    };
    const snap = build(saved[compare].rows);
    const liveAll = build(files.netsuite?.rows || []);
    // only compare live rows from the same month
    const live = new Map([...liveAll.entries()].filter(([, v]) => monthKey(v.date) === compare));

    const removed = [], changed = [], unpaid = [], added = [];
    for (const [k, a] of snap) {
      const b = live.get(k);
      if (!b) { removed.push({ ...a, was: a.gp }); continue; }
      if (Math.abs((a.gp || 0) - (b.gp || 0)) > 0.005) changed.push({ ...b, was: a.gp, now: b.gp });
      const paidNow = b.itemPaid != null ? isYes(b.itemPaid) : b.statusPaid;
      if (!paidNow && statusCounts(statusRules, b.status) && !isNonCommissionable(b) && (b.gp || 0) > 0) {
        unpaid.push({ ...b, was: a.gp });
      }
    }
    for (const [k, b] of live) if (!snap.has(k)) added.push(b);
    return { removed, changed, unpaid, added, snapCount: snap.size, liveCount: live.size };
  }, [compare, saved, files, statusRules]);

  const Section = ({ title, rows, colour, cols }) => (
    <div className="panel" style={{ borderLeft: `4px solid ${colour}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h2 style={{ margin: 0 }}>{title} <span className="sub">({rows.length})</span></h2>
        {rows.length > 0 && (
          <button className="btn" onClick={() => downloadCSV(rows.map((r) => ({
            "Order ref": r.orderNum || "", "NetSuite ref": r.netsuiteRef || "", Company: r.company || "",
            Partner: r.partner || "", Item: r.product || "", Status: r.status || "",
            "GP was": r.was ?? "", "GP now": r.now ?? r.gp ?? "",
          })), `${title.replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-${compare}.csv`)}>Export</button>
        )}
      </div>
      {rows.length === 0 ? <p className="note" style={{ marginTop: 8 }}>None.</p> : (
        <div style={{ overflowX: "auto", marginTop: 8, maxHeight: 320, overflowY: "auto" }}>
          <table>
            <thead><tr>
              <th className="left">Order</th><th className="left">Company</th><th className="left">Partner</th>
              <th className="left">Status</th>{cols.map((c) => <th key={c} className="num">{c}</th>)}
            </tr></thead>
            <tbody>
              {rows.slice(0, 150).map((r, i) => (
                <tr key={i}>
                  <td className="left mono">{r.orderNum || r.netsuiteRef || "-"}</td>
                  <td className="left">{r.company || "-"}</td>
                  <td className="left">{r.partner || "-"}</td>
                  <td className="left">{r.status || "-"}</td>
                  {cols.includes("Was") && <td className="num mono">{gbp(r.was)}</td>}
                  {cols.includes("Now") && <td className="num mono">{gbp(r.now ?? r.gp)}</td>}
                  {cols.includes("Change") && (
                    <td className={"num mono " + ((r.now ?? r.gp) - r.was >= 0 ? "pos" : "neg")}>
                      {gbp((r.now ?? r.gp) - r.was)}
                    </td>
                  )}
                  {cols.includes("GP") && <td className="num mono">{gbp(r.gp)}</td>}
                </tr>
              ))}
            </tbody>
          </table>
          {rows.length > 150 && <p className="note">Showing 150 of {rows.length} — use Export for the full list.</p>}
        </div>
      )}
    </div>
  );

  return (
    <>
      <div className="panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <h2 style={{ margin: 0 }}>Monthly NetSuite snapshots — {fyLabel(year)}</h2>
          <div className="settings">
            <span>Financial year:</span>
            <select value={year} onChange={(e) => { setYear(Number(e.target.value)); setCompare(null); }}
              style={{ padding: "5px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13 }}>
              {[thisFy + 1, thisFy, thisFy - 1, thisFy - 2].map((y) => <option key={y} value={y}>{fyLabel(y)}</option>)}
            </select>
          </div>
        </div>
        <p className="note" style={{ marginTop: 6 }}>
          Upload the NetSuite report as it stood at each month end. Comparing a month against the live report shows
          what has since been removed, re-valued, or is still sitting unpaid — using the same payable rules as the rest of the app.
        </p>
        {msg && <p className="sub" style={{ color: msg.err ? "#b3261e" : "#14804a" }}>{msg.text}</p>}
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>
              <th className="left">Month</th><th className="num">Rows</th><th className="left">Saved</th>
              <th className="num">File</th><th className="num"></th>
            </tr></thead>
            <tbody>
              {keys.map((mk, i) => {
                const sv = saved[mk];
                return (
                  <tr key={mk} style={{ background: compare === mk ? "#efeaff" : undefined }}>
                    <td className="left"><strong>{MONTHS_FY[i][1]}</strong> <span className="sub">{mk}</span></td>
                    <td className="num mono">{sv ? (sv.count ?? sv.rows?.length ?? 0) : "-"}</td>
                    <td className="left sub">{sv?.at ? `${new Date(sv.at).toLocaleDateString("en-GB")} · ${sv.by || ""}` : "not uploaded"}</td>
                    <td className="num">
                      <label className="file" style={{ fontSize: 11, padding: "5px 9px" }}>
                        {busy === mk ? "Saving…" : sv ? "Replace" : "Upload"}
                        <input type="file" accept=".csv,.tsv,.txt,.xlsx,.xls"
                          onChange={(e) => upload(mk, e.target.files[0])} />
                      </label>
                    </td>
                    <td className="num">
                      {sv && (
                        <>
                          <button className="btn" style={{ marginRight: 6 }}
                            onClick={() => setCompare(compare === mk ? null : mk)}>
                            {compare === mk ? "Hide" : "Compare"}
                          </button>
                          <button className="btn" onClick={() => removeSnapshot(mk)}>Delete</button>
                        </>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {compare && !diff && <div className="empty panel">Loading that snapshot…</div>}
      {compare && diff && (
        <>
          <div className="banner info">
            <strong>{periodLabel(compare)}:</strong> snapshot holds {diff.snapCount} lines, the live report holds {diff.liveCount} for that month.
          </div>
          <Section title="Removed since the snapshot" rows={diff.removed} colour="#b3261e" cols={["Was"]} />
          <Section title="Value changed" rows={diff.changed} colour="#d98a00" cols={["Was", "Now", "Change"]} />
          <Section title="Still unpaid" rows={diff.unpaid} colour="#1e64d6" cols={["GP"]} />
          <Section title="New since the snapshot" rows={diff.added} colour="#14804a" cols={["GP"]} />
        </>
      )}
    </>
  );
}

// Row of the Cobra Dashboard grid. Kept outside WipTracker so React reuses it between
// renders — otherwise every keystroke remounts the row and the input loses focus.
function WipRow({ label, tag, vals, kind, rowKey, danger, group, manual, setM, drill, setDrill, totalPct, tot }) {
  const cellNum = (v) => (v == null ? "-" : gbp0(v));
  return (
    <tr className={group ? "grp" : ""}>
      <td className="left lbl" style={{ whiteSpace: "nowrap" }}>
        <span className={"chip " + (tag === "Formula" ? "mismatch" : tag === "Report" ? "matched" : "unmatched")}
          style={{ marginRight: 8, fontSize: 10 }}>{tag}</span>
        <strong>{label}</strong>
      </td>
      {vals.map((v, i) => (
        <td key={i} className="num mono"
          onClick={tag === "Report" ? () => setDrill({ label, monthIndex: i }) : undefined}
          title={tag === "Report" ? "Click to see the orders behind this figure" : undefined}
          style={{
            cursor: tag === "Report" ? "pointer" : undefined,
            textDecoration: tag === "Report" && v ? "underline dotted" : undefined,
            background: drill && drill.label === label && drill.monthIndex === i ? "#efeaff"
              : danger && v != null && danger(v) ? "#fbe9e7" : undefined,
            color: danger && v != null && danger(v) ? "#b3261e" : undefined,
          }}>
          {kind === "input"
            ? <NumInput value={manual[rowKey]?.[i] ?? ""} placeholder="-" onCommit={(nv) => setM(rowKey, i, nv)} />
            : kind === "pct" ? fmtPct(v) : cellNum(v)}
        </td>
      ))}
      <td className="num mono tot">{kind === "pct" ? fmtPct(totalPct[label]) : gbp0(tot(vals))}</td>
    </tr>
  );
}

// =========================================================================
//  Top Risk Orders — orders whose status carries a high chance of falling away
// =========================================================================
function TopRiskOrders({ records, settings, statusRules, saveSettings }) {
  const threshold = settings.riskThreshold ?? 50;

  const rows = useMemo(() => {
    const out = [];
    for (const r of records) {
      if (!isRealLine(r)) continue;
      const pctRisk = riskPctFor(settings, statusRules, r.nsStatus);
      if (pctRisk < threshold) continue;
      const gp = r.expected || 0;
      if (!gp) continue;
      out.push({ ...r, pctRisk, gp, atRisk: gp * (pctRisk / 100) });
    }
    return out.sort((a, b) => b.atRisk - a.atRisk);
  }, [records, settings, statusRules, threshold]);

  const totals = useMemo(() => ({
    gp: sum(rows.map((r) => r.gp)),
    atRisk: sum(rows.map((r) => r.atRisk)),
  }), [rows]);

  const byStatus = useMemo(() => {
    const m = new Map();
    for (const r of rows) {
      const k = String(r.nsStatus || "(blank)").trim() || "(blank)";
      if (!m.has(k)) m.set(k, { status: k, n: 0, gp: 0, atRisk: 0, pctRisk: r.pctRisk });
      const e = m.get(k);
      e.n++; e.gp += r.gp; e.atRisk += r.atRisk;
    }
    return [...m.values()].sort((a, b) => b.atRisk - a.atRisk);
  }, [rows]);

  const [limit, setLimit] = useState(150);

  return (
    <>
      <div className="panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <h2 style={{ margin: 0 }}>Orders at {threshold}% risk or higher</h2>
          <div className="settings">
            <span>Risk threshold:</span>
            <NumInput value={settings.riskThreshold ?? ""} placeholder="50"
              onCommit={(nv) => saveSettings({ ...settings, riskThreshold: nv === "" ? "" : Number(nv) })}
              style={{ width: 70, padding: "6px 8px", border: "1px solid #d3d0e6", borderRadius: 6, fontSize: 13, textAlign: "center" }} />
            <span style={{ color: "#8a8aa3" }}>% — set each status's risk in Settings → Order Statuses.</span>
          </div>
        </div>
        <div className="kpis5" style={{ gridTemplateColumns: "repeat(3,1fr)", marginTop: 12 }}>
          <div className="kpic" style={{ borderTopColor: "#b3261e" }}>
            <div className="lab">Orders at risk</div>
            <div className="val" style={{ color: "#b3261e" }}>{rows.length.toLocaleString()}</div>
          </div>
          <div className="kpic" style={{ borderTopColor: "#d98a00" }}>
            <div className="lab">GP on those orders</div>
            <div className="val" style={{ color: "#d98a00" }}>{gbp0(totals.gp)}</div>
          </div>
          <div className="kpic" style={{ borderTopColor: "#b3261e" }}>
            <div className="lab">Weighted GP at risk</div>
            <div className="val" style={{ color: "#b3261e" }}>{gbp0(totals.atRisk)}</div>
            <div className="sub2" style={{ color: "#8a8aa3", fontWeight: 500 }}>GP × risk %</div>
          </div>
        </div>
      </div>

      {byStatus.length > 0 && (
        <div className="panel">
          <h2>By status</h2>
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead><tr>
                <th className="left">Status</th><th className="num">Risk %</th><th className="num">Orders</th>
                <th className="num">GP</th><th className="num">Weighted at risk</th>
              </tr></thead>
              <tbody>
                {byStatus.map((b) => (
                  <tr key={b.status}>
                    <td className="left">{b.status}</td>
                    <td className="num mono">{b.pctRisk}%</td>
                    <td className="num mono">{b.n.toLocaleString()}</td>
                    <td className="num mono">{gbp0(b.gp)}</td>
                    <td className="num mono neg">{gbp0(b.atRisk)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2 style={{ margin: 0 }}>Orders</h2>
          {rows.length > 0 && (
            <button className="btn" onClick={() => downloadCSV(rows.map((r) => ({
              "Order ref": r.orderNum || "", "NetSuite ref": r.netsuiteRef || "", Company: r.company || "",
              Partner: r.agent || "", Status: r.nsStatus || "", "Risk %": r.pctRisk,
              GP: r.gp, "Weighted at risk": Math.round(r.atRisk * 100) / 100,
            })), "top-risk-orders.csv")}>Export</button>
          )}
        </div>
        {rows.length === 0 ? (
          <div className="empty">No orders at or above {threshold}% risk.</div>
        ) : (
          <div style={{ overflowX: "auto", marginTop: 10 }}>
            <table>
              <thead><tr>
                <th className="left">Order</th><th className="left">Company</th><th className="left">Status</th>
                <th className="num">Risk %</th><th className="num">GP</th><th className="num">At risk</th>
              </tr></thead>
              <tbody>
                {rows.slice(0, limit).map((r) => (
                  <tr key={r.key} style={{ background: r.pctRisk >= 90 ? "#fbe9e7" : r.pctRisk >= 70 ? "#fdecea" : undefined }}>
                    <td className="left mono">{r.orderNum || r.netsuiteRef || "-"}</td>
                    <td className="left">{r.company || "-"}</td>
                    <td className="left">{r.nsStatus || "-"}</td>
                    <td className="num mono" style={{ color: r.pctRisk >= 70 ? "#b3261e" : "#8a5a00" }}>{r.pctRisk}%</td>
                    <td className="num mono">{gbp0(r.gp)}</td>
                    <td className="num mono neg">{gbp0(r.atRisk)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {rows.length > limit && (
              <p className="note">
                Showing {limit.toLocaleString()} of {rows.length.toLocaleString()}.{" "}
                <button className="btn" onClick={() => setLimit((n) => n + 500)}>Show 500 more</button>
              </p>
            )}
          </div>
        )}
      </div>
    </>
  );
}
