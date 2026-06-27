// Pure data-cleaning utilities (no React, no DOM) — easy to unit-test and reuse.

export type Row = Record<string, unknown>;

export type ColumnType = "number" | "date" | "string" | "empty";

export interface ColumnStats {
  name: string;
  type: ColumnType;
  missing: number;
  filled: number;
  unique: number;
  min?: number;
  max?: number;
  mean?: number;
}

export interface CleanReport {
  originalRows: number;
  cleanedRows: number;
  duplicatesRemoved: number;
  missingFilled: number;
  whitespaceTrimmed: number;
  caseNormalized: number;
  datesNormalized: number;
  numbersCoerced: number;
  columns: ColumnStats[];
  notes: string[];
}

export interface CleanOptions {
  dropDuplicates: boolean;
  trimWhitespace: boolean;
  normalizeCase: boolean; // titlecase for string columns that look like categories
  fillNumericWith: "mean" | "median" | "zero" | "none";
  fillStringWith: "mode" | "unknown" | "none";
  normalizeDates: boolean;
  dropEmptyColumns: boolean;
}

export const defaultOptions: CleanOptions = {
  dropDuplicates: true,
  trimWhitespace: true,
  normalizeCase: true,
  fillNumericWith: "median",
  fillStringWith: "mode",
  normalizeDates: true,
  dropEmptyColumns: true,
};

const MISSING_TOKENS = new Set(["", "na", "n/a", "null", "none", "nan", "-", "--", "?"]);

const isMissing = (v: unknown): boolean => {
  if (v === null || v === undefined) return true;
  if (typeof v === "number") return Number.isNaN(v);
  if (typeof v === "string") return MISSING_TOKENS.has(v.trim().toLowerCase());
  return false;
};

const tryNumber = (v: unknown): number | null => {
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  if (typeof v !== "string") return null;
  const cleaned = v.replace(/[, $]/g, "").trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
};

const tryDate = (v: unknown): Date | null => {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v !== "string" && typeof v !== "number") return null;
  const s = String(v).trim();
  if (!s) return null;
  // Common formats: ISO, dd/mm/yyyy, mm/dd/yyyy, dd-mm-yyyy
  const iso = new Date(s);
  if (!isNaN(iso.getTime()) && /\d{4}/.test(s)) return iso;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let [, a, b, y] = m;
    let year = parseInt(y, 10);
    if (year < 100) year += 2000;
    // prefer dd/mm/yyyy
    const day = parseInt(a, 10);
    const month = parseInt(b, 10);
    if (day > 12 && month <= 12) return new Date(year, month - 1, day);
    if (month > 12 && day <= 12) return new Date(year, day - 1, month);
    return new Date(year, month - 1, day);
  }
  return null;
};

const detectType = (values: unknown[]): ColumnType => {
  const non = values.filter((v) => !isMissing(v));
  if (non.length === 0) return "empty";
  let nums = 0;
  let dates = 0;
  for (const v of non) {
    if (tryNumber(v) !== null) nums++;
    else if (tryDate(v) !== null) dates++;
  }
  if (nums / non.length >= 0.8) return "number";
  if (dates / non.length >= 0.8) return "date";
  return "string";
};

const median = (xs: number[]): number => {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
};

const mean = (xs: number[]): number =>
  xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;

const mode = <T,>(xs: T[]): T | undefined => {
  const counts = new Map<T, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best: T | undefined;
  let bestCount = 0;
  for (const [k, c] of counts) if (c > bestCount) { best = k; bestCount = c; }
  return best;
};

const titleCase = (s: string): string =>
  s.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());

export function cleanData(
  rows: Row[],
  options: CleanOptions = defaultOptions,
): { cleaned: Row[]; report: CleanReport } {
  const report: CleanReport = {
    originalRows: rows.length,
    cleanedRows: 0,
    duplicatesRemoved: 0,
    missingFilled: 0,
    whitespaceTrimmed: 0,
    caseNormalized: 0,
    datesNormalized: 0,
    numbersCoerced: 0,
    columns: [],
    notes: [],
  };

  if (!rows.length) return { cleaned: [], report };

  // Collect all columns preserving order
  const columns: string[] = [];
  const seenCol = new Set<string>();
  for (const r of rows) {
    for (const k of Object.keys(r)) {
      if (!seenCol.has(k)) { seenCol.add(k); columns.push(k); }
    }
  }

  // Per-column raw values
  const colValues: Record<string, unknown[]> = {};
  for (const c of columns) colValues[c] = rows.map((r) => r[c]);

  // Detect type
  const types: Record<string, ColumnType> = {};
  for (const c of columns) types[c] = detectType(colValues[c]);

  // Drop fully empty cols
  let activeCols = columns;
  if (options.dropEmptyColumns) {
    activeCols = columns.filter((c) => types[c] !== "empty");
    const removed = columns.length - activeCols.length;
    if (removed) report.notes.push(`Dropped ${removed} empty column(s).`);
  }

  // Compute fillers per column
  const fillers: Record<string, unknown> = {};
  for (const c of activeCols) {
    const t = types[c];
    const present = colValues[c].filter((v) => !isMissing(v));
    if (t === "number") {
      const nums = present.map(tryNumber).filter((n): n is number => n !== null);
      if (options.fillNumericWith === "mean") fillers[c] = mean(nums);
      else if (options.fillNumericWith === "median") fillers[c] = median(nums);
      else if (options.fillNumericWith === "zero") fillers[c] = 0;
    } else if (t === "string") {
      const strs = present.map((v) => String(v).trim()).filter(Boolean);
      if (options.fillStringWith === "mode") fillers[c] = mode(strs) ?? "Unknown";
      else if (options.fillStringWith === "unknown") fillers[c] = "Unknown";
    }
  }

  // Transform each row
  const cleaned: Row[] = [];
  const seen = new Set<string>();
  for (const r of rows) {
    const out: Row = {};
    for (const c of activeCols) {
      let v = r[c];
      const t = types[c];

      if (typeof v === "string" && options.trimWhitespace) {
        const trimmed = v.replace(/\s+/g, " ").trim();
        if (trimmed !== v) report.whitespaceTrimmed++;
        v = trimmed;
      }

      if (isMissing(v)) {
        if (fillers[c] !== undefined && options.fillNumericWith !== "none" && options.fillStringWith !== "none") {
          v = fillers[c];
          report.missingFilled++;
        } else if (t === "number" && options.fillNumericWith !== "none") {
          v = fillers[c]; report.missingFilled++;
        } else if (t === "string" && options.fillStringWith !== "none") {
          v = fillers[c]; report.missingFilled++;
        } else {
          v = null;
        }
      } else {
        if (t === "number") {
          const n = tryNumber(v);
          if (n !== null) {
            if (n !== v) report.numbersCoerced++;
            v = n;
          }
        } else if (t === "date" && options.normalizeDates) {
          const d = tryDate(v);
          if (d) {
            const iso = d.toISOString().slice(0, 10);
            if (iso !== v) report.datesNormalized++;
            v = iso;
          }
        } else if (t === "string" && options.normalizeCase && typeof v === "string") {
          // Only normalize when value looks like a short category
          if (v.length <= 40 && /^[a-zA-Z\s\-'/]+$/.test(v)) {
            const tc = titleCase(v);
            if (tc !== v) report.caseNormalized++;
            v = tc;
          }
        }
      }

      out[c] = v;
    }

    if (options.dropDuplicates) {
      const key = JSON.stringify(out);
      if (seen.has(key)) { report.duplicatesRemoved++; continue; }
      seen.add(key);
    }
    cleaned.push(out);
  }

  // Column stats
  for (const c of activeCols) {
    const vals = cleaned.map((r) => r[c]);
    const missing = vals.filter((v) => v === null || v === undefined || v === "").length;
    const filled = vals.length - missing;
    const unique = new Set(vals.map((v) => JSON.stringify(v))).size;
    const stat: ColumnStats = { name: c, type: types[c], missing, filled, unique };
    if (types[c] === "number") {
      const nums = vals.map(tryNumber).filter((n): n is number => n !== null);
      if (nums.length) {
        stat.min = Math.min(...nums);
        stat.max = Math.max(...nums);
        stat.mean = mean(nums);
      }
    }
    report.columns.push(stat);
  }

  report.cleanedRows = cleaned.length;
  return { cleaned, report };
}

// Built-in messy sample dataset
export function generateMessySample(): Row[] {
  const regions = ["north", "South", "EAST", " west ", "North", "south", "East", "West"];
  const products = ["Widget A", "widget a", "Widget B", "WIDGET B", "Gadget", "gadget ", "Sprocket", "Sprocket"];
  const channels = ["Online", "online", "Retail", "RETAIL", "Partner", "partner"];
  const rows: Row[] = [];
  const start = new Date(2024, 0, 1).getTime();
  for (let i = 0; i < 220; i++) {
    const date = new Date(start + Math.floor(Math.random() * 330) * 86400000);
    const fmt = i % 3 === 0
      ? date.toISOString().slice(0, 10)
      : i % 3 === 1
        ? `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`
        : `${date.getMonth() + 1}-${date.getDate()}-${date.getFullYear()}`;
    const qty = Math.random() < 0.08 ? null : Math.floor(Math.random() * 50) + 1;
    const price = Math.random() < 0.1 ? "" : (Math.random() * 200 + 10).toFixed(2);
    const revenueRaw = qty && price ? `$${(qty * Number(price)).toFixed(2)}` : "N/A";
    rows.push({
      OrderID: `ORD-${1000 + i}`,
      Date: fmt,
      Region: Math.random() < 0.05 ? "" : regions[i % regions.length],
      Product: products[i % products.length],
      Channel: channels[i % channels.length],
      Quantity: qty,
      UnitPrice: price,
      Revenue: revenueRaw,
      Customer: Math.random() < 0.07 ? "  " : `Customer ${Math.floor(Math.random() * 60)}`,
    });
    // inject duplicates occasionally
    if (i % 25 === 0 && i > 0) rows.push({ ...rows[rows.length - 1] });
  }
  return rows;
}
