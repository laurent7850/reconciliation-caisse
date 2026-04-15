import * as XLSX from "xlsx";

// ---------- Types ----------
export interface ZData {
  zNumber: number;
  openDate: Date;
  closeDate: Date;
  caTTC: number;
  ttc21: number;
  ttc12: number;
  ttc6: number;
  tickets: number;
  sourceFile: string;
}
export interface CAData {
  zNumber: number;
  date: Date;
  cash: number;
  carteBanque: number;
  virementBancaire: number;
  sourceFile: string;
}
export interface ZEntry {
  zNumber: number;
  day: number;
  monthIndex: number; // 0-11
  monthLabel: string;
  dateLabel: string;
  report: ZData;
  ca: CAData | null;
}
export interface ReconciliationRow {
  zNumber: number;
  day: number;
  monthLabel: string;
  dateLabel: string;
  sheetName: string;
  rowIndex: number; // row index in sheet (0-based)
  values: {
    zNumber: number;
    totalTVAC: number;
    total21: number;
    total12: number;
    total6: number;
    cartes: number;
    virement: number;
    cash: number;
  };
  existing: Partial<ReconciliationRow["values"]>;
  conflicts: string[]; // field keys that already contain a value differing from proposed
  hasData: boolean; // the target row already has any non-zero/non-null data
  applied: boolean;
}

// ---------- Month/sheet helpers ----------
const MONTHS_FR = [
  "JANVIER",
  "FEVRIER",
  "MARS",
  "AVRIL",
  "MAI",
  "JUIN",
  "JUILLET",
  "AOUT",
  "SEPTEMBRE",
  "OCTOBRE",
  "NOVEMBRE",
  "DECEMBRE",
];

export function monthLabelFromIndex(i: number): string {
  return MONTHS_FR[i];
}

// Normalize "FEVRIER"/"Février"/"FÉVRIER" → "FEVRIER"
function normalizeMonth(s: string): string {
  return s
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();
}

function findSheetForMonth(wb: XLSX.WorkBook, monthIndex: number): string | null {
  const target = MONTHS_FR[monthIndex];
  for (const name of wb.SheetNames) {
    if (normalizeMonth(name) === target) return name;
  }
  return null;
}

// ---------- Date parsing ----------
function parseDateFR(s: string): Date {
  // "04/04/2026 19:03" or "2026/04/04"
  const str = String(s).trim();
  let m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    return new Date(
      Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1]),
      m[4] ? Number(m[4]) : 0,
      m[5] ? Number(m[5]) : 0
    );
  }
  m = str.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  throw new Error(`Date illisible: ${str}`);
}

function num(v: unknown): number {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const cleaned = String(v).replace(/\s/g, "").replace(",", ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

// ---------- Parsers ----------
export function parseReportZ(file: File, rows: unknown[][]): ZData {
  const header = rows[0] as string[];
  const data = rows[1] as unknown[];
  const idx = (k: string) => header.findIndex((h) => String(h).trim().toLowerCase() === k.toLowerCase());
  const z = Number(data[idx("Rapport Z")]);
  const openDate = parseDateFR(String(data[idx("Date d'ouverture")]));
  const closeDate = parseDateFR(String(data[idx("Date de fermeture")]));
  return {
    zNumber: z,
    openDate,
    closeDate,
    caTTC: num(data[idx("CA TTC")]),
    ttc21: num(data[idx("TTC 21%")]),
    ttc12: num(data[idx("TTC 12%")]),
    ttc6: num(data[idx("TTC 6%")]),
    tickets: num(data[idx("Tickets")]),
    sourceFile: file.name,
  };
}

export function parseCA(file: File, rows: unknown[][]): CAData {
  const header = rows[0] as string[];
  const data = rows[1] as unknown[];
  const idx = (k: string) => header.findIndex((h) => String(h).trim().toLowerCase() === k.toLowerCase());
  return {
    // CA file doesn't contain Z number — we'll match by date/filename later
    zNumber: NaN,
    date: parseDateFR(String(data[idx("Date")])),
    cash: num(data[idx("Cash")]),
    carteBanque: num(data[idx("Carte banque")]),
    virementBancaire: num(data[idx("Virement bancaire")]),
    sourceFile: file.name,
  };
}

// Filename pattern: ReportZStats_1_442.xlsx / CA_1_442.xlsx
export function zFromFilename(name: string): number | null {
  const m = name.match(/_(\d+)\.xlsx?$/i);
  return m ? Number(m[1]) : null;
}

// ---------- Recap structure ----------
/**
 * Each month sheet has:
 * row 0: title row ("JANVIER", "SK", ...)
 * row 1: headers ("Jour", "Z N°", "TOTAL TVAC", "TOTAL 21% TVAC", "TOTAL 12% TVAC",
 *                 "TOTAL 6% TVAC", "Paiements cartes", ["BON CADEAU"?],
 *                 "Virement client resto sur le compte", "Dépôt Cash - 58000",
 *                 "FOURNISSEURS", "Montant", [TOTAL CAISSE header moved around], "CASH")
 * rows 2..N : 1 row per day (col 0 = day number)
 *
 * We'll detect columns by header matching, tolerating shifts.
 */
export interface SheetLayout {
  sheetName: string;
  headerRow: number; // usually 1
  colJour: number;
  colZ: number;
  colTVAC: number;
  col21: number;
  col12: number;
  col6: number;
  colCartes: number;
  colVirement: number;
  colCash: number; // "CASH" column
}

function headerMatch(h: unknown, keys: string[]): boolean {
  if (h == null) return false;
  const raw = String(h).toLowerCase().replace(/\s+/g, " ").trim();
  return keys.some((k) => raw.includes(k.toLowerCase()));
}

export function detectLayout(ws: XLSX.WorkSheet, sheetName: string): SheetLayout | null {
  const data = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null });
  const headerRowIdx = 1;
  const header = data[headerRowIdx];
  if (!header) return null;

  const find = (keys: string[]): number =>
    header.findIndex((h) => headerMatch(h, keys));

  const layout: SheetLayout = {
    sheetName,
    headerRow: headerRowIdx,
    colJour: find(["jour"]),
    colZ: find(["z n"]),
    colTVAC: find(["total tvac"]),
    col21: find(["total 21"]),
    col12: find(["total 12"]),
    col6: find(["total 6"]),
    colCartes: find(["paiements cartes", "paiement cartes"]),
    colVirement: find(["virement client"]),
    colCash: -1,
  };

  // CASH column is tricky: sometimes a header cell like "CASH" is in the TITLE row (row 0)
  // or in the header row (row 1). Scan both.
  const scanCash = (row: unknown[] | undefined): number =>
    row ? row.findIndex((h) => h != null && String(h).toUpperCase().trim() === "CASH") : -1;
  let cashCol = scanCash(data[headerRowIdx]);
  if (cashCol < 0) cashCol = scanCash(data[0]);
  layout.colCash = cashCol;

  // Minimal sanity: we need jour, z, tvac at least
  if (layout.colJour < 0 || layout.colZ < 0 || layout.colTVAC < 0) return null;
  return layout;
}

// ---------- Reconciliation logic ----------
export interface ReconcileInput {
  recapFile: File;
  recapWB: XLSX.WorkBook;
  entries: ZEntry[]; // merged Z+CA entries
}

export function buildEntries(zFiles: ZData[], caFiles: CAData[]): ZEntry[] {
  const caByZ = new Map<number, CAData>();
  const caByDate = new Map<string, CAData>();
  for (const ca of caFiles) {
    const z = zFromFilename(ca.sourceFile);
    if (z != null) caByZ.set(z, ca);
    caByDate.set(ca.date.toISOString().slice(0, 10), ca);
  }
  return zFiles.map((z) => {
    const ca =
      caByZ.get(z.zNumber) ??
      caByDate.get(z.openDate.toISOString().slice(0, 10)) ??
      null;
    return {
      zNumber: z.zNumber,
      day: z.openDate.getDate(),
      monthIndex: z.openDate.getMonth(),
      monthLabel: MONTHS_FR[z.openDate.getMonth()],
      dateLabel: z.openDate.toLocaleDateString("fr-BE"),
      report: z,
      ca,
    };
  });
}

const FIELD_TOLERANCE = 0.005; // 0.5 cent

function isEmpty(v: unknown): boolean {
  if (v == null || v === "") return true;
  if (typeof v === "number") return v === 0;
  const n = num(v);
  return n === 0 && String(v).trim() === "0";
}

export function computeReconciliation(
  entries: ZEntry[],
  wb: XLSX.WorkBook
): ReconciliationRow[] {
  const out: ReconciliationRow[] = [];
  for (const e of entries) {
    const sheetName = findSheetForMonth(wb, e.monthIndex);
    if (!sheetName) continue;
    const ws = wb.Sheets[sheetName];
    const layout = detectLayout(ws, sheetName);
    if (!layout) continue;
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null });
    // Find row where colJour == e.day
    const rowIdx = rows.findIndex(
      (r, i) => i > layout.headerRow && Number(r?.[layout.colJour]) === e.day
    );
    if (rowIdx < 0) continue;
    const row = rows[rowIdx];
    const proposed = {
      zNumber: e.zNumber,
      totalTVAC: e.report.caTTC,
      total21: e.report.ttc21,
      total12: e.report.ttc12,
      total6: e.report.ttc6,
      cartes: e.ca?.carteBanque ?? 0,
      virement: e.ca?.virementBancaire ?? 0,
      cash: e.ca?.cash ?? 0,
    };
    const existing: Partial<typeof proposed> = {};
    const conflicts: string[] = [];
    const check = (key: keyof typeof proposed, colIdx: number) => {
      if (colIdx < 0) return;
      const cur = row[colIdx];
      if (!isEmpty(cur)) {
        existing[key] = num(cur);
        if (Math.abs(num(cur) - proposed[key]) > FIELD_TOLERANCE) {
          conflicts.push(key);
        }
      }
    };
    check("zNumber", layout.colZ);
    check("totalTVAC", layout.colTVAC);
    check("total21", layout.col21);
    check("total12", layout.col12);
    check("total6", layout.col6);
    check("cartes", layout.colCartes);
    check("virement", layout.colVirement);
    check("cash", layout.colCash);
    const hasData = Object.keys(existing).length > 0;
    out.push({
      zNumber: e.zNumber,
      day: e.day,
      monthLabel: e.monthLabel,
      dateLabel: e.dateLabel,
      sheetName,
      rowIndex: rowIdx,
      values: proposed,
      existing,
      conflicts,
      hasData,
      applied: false,
    });
  }
  return out;
}

/**
 * Apply proposals to workbook.
 * Policy: never overwrite a non-empty cell. Return the list of rows with final "applied" flag.
 */
export function applyReconciliation(
  wb: XLSX.WorkBook,
  rows: ReconciliationRow[]
): ReconciliationRow[] {
  for (const r of rows) {
    const ws = wb.Sheets[r.sheetName];
    const layout = detectLayout(ws, r.sheetName);
    if (!layout) continue;
    const data = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null });
    const row = data[r.rowIndex] ?? [];
    const writeIfEmpty = (colIdx: number, value: number) => {
      if (colIdx < 0) return;
      if (isEmpty(row[colIdx])) {
        row[colIdx] = value;
      }
    };
    writeIfEmpty(layout.colZ, r.values.zNumber);
    writeIfEmpty(layout.colTVAC, r.values.totalTVAC);
    writeIfEmpty(layout.col21, r.values.total21);
    writeIfEmpty(layout.col12, r.values.total12);
    writeIfEmpty(layout.col6, r.values.total6);
    writeIfEmpty(layout.colCartes, r.values.cartes);
    writeIfEmpty(layout.colVirement, r.values.virement);
    writeIfEmpty(layout.colCash, r.values.cash);
    data[r.rowIndex] = row;
    // Rebuild sheet while preserving other rows as-is
    const newWs = XLSX.utils.aoa_to_sheet(data);
    // Keep original !ref and column widths if present
    if (ws["!cols"]) newWs["!cols"] = ws["!cols"];
    if (ws["!merges"]) newWs["!merges"] = ws["!merges"];
    wb.Sheets[r.sheetName] = newWs;
    r.applied = true;
  }
  return rows;
}

export async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { cellDates: false });
}

export async function readSheetAOA(file: File): Promise<unknown[][]> {
  const wb = await readWorkbook(file);
  const first = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<unknown[]>(first, { header: 1, defval: null });
}

export function downloadWorkbook(wb: XLSX.WorkBook, filename: string) {
  XLSX.writeFile(wb, filename, { bookType: "xlsx" });
}
