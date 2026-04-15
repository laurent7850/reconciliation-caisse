import ExcelJS from "exceljs";

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
  monthIndex: number;
  monthLabel: string;
  dateLabel: string;
  report: ZData;
  ca: CAData | null;
}
export interface ProposedValues {
  zNumber: number;
  totalTVAC: number;
  total21: number;
  total12: number;
  total6: number;
  cartes: number;
  virement: number;
  cash: number;
}
export interface ReconciliationRow {
  zNumber: number;
  day: number;
  monthLabel: string;
  dateLabel: string;
  sheetName: string;
  excelRow: number; // 1-indexed row in sheet
  values: ProposedValues;
  existing: Partial<ProposedValues>;
  conflicts: string[];
  hasData: boolean;
  applied: boolean;
}

const MONTHS_FR = [
  "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
  "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE",
];

function normalizeMonth(s: string): string {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
}

function findSheetForMonth(wb: ExcelJS.Workbook, monthIndex: number): ExcelJS.Worksheet | null {
  const target = MONTHS_FR[monthIndex];
  for (const ws of wb.worksheets) {
    if (normalizeMonth(ws.name) === target) return ws;
  }
  return null;
}

function parseDateFR(s: string | Date | number): Date {
  if (s instanceof Date) return s;
  if (typeof s === "number") {
    // Excel serial date
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + s * 86400000);
  }
  const str = String(s).trim();
  let m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    return new Date(
      Number(m[3]), Number(m[2]) - 1, Number(m[1]),
      m[4] ? Number(m[4]) : 0, m[5] ? Number(m[5]) : 0
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
  if (typeof v === "object" && v != null && "result" in v) return num((v as { result: unknown }).result);
  const cleaned = String(v).replace(/\s/g, "").replace(",", ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

// Convert ExcelJS cell value (may be object with .result, .text, .richText) to simple value
function cellVal(c: ExcelJS.Cell): unknown {
  const v = c.value;
  if (v == null) return null;
  if (typeof v === "object") {
    if ("result" in v) return (v as { result: unknown }).result;
    if ("text" in v) return (v as { text: unknown }).text;
    if ("richText" in v) return (v as { richText: { text: string }[] }).richText.map(r => r.text).join("");
  }
  return v;
}

function headerText(c: ExcelJS.Cell): string {
  const v = cellVal(c);
  return v == null ? "" : String(v).toLowerCase().replace(/\s+/g, " ").trim();
}

// ---------- Source file parsing ----------
export async function loadWorkbook(file: File): Promise<ExcelJS.Workbook> {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

function findHeaderIndex(ws: ExcelJS.Worksheet, headerRow: number, key: string): number {
  const row = ws.getRow(headerRow);
  for (let c = 1; c <= row.cellCount; c++) {
    if (headerText(row.getCell(c)) === key.toLowerCase()) return c;
  }
  return -1;
}

export async function parseReportZ(file: File): Promise<ZData> {
  const wb = await loadWorkbook(file);
  const ws = wb.worksheets[0];
  const find = (k: string) => findHeaderIndex(ws, 1, k);
  const row = ws.getRow(2);
  const get = (k: string) => cellVal(row.getCell(find(k)));
  const zNumber = Number(get("Rapport Z"));
  const openDate = parseDateFR(get("Date d'ouverture") as string);
  const closeDate = parseDateFR(get("Date de fermeture") as string);
  return {
    zNumber,
    openDate,
    closeDate,
    caTTC: num(get("CA TTC")),
    ttc21: num(get("TTC 21%")),
    ttc12: num(get("TTC 12%")),
    ttc6: num(get("TTC 6%")),
    tickets: num(get("Tickets")),
    sourceFile: file.name,
  };
}

export async function parseCA(file: File): Promise<CAData> {
  const wb = await loadWorkbook(file);
  const ws = wb.worksheets[0];
  const find = (k: string) => findHeaderIndex(ws, 1, k);
  const row = ws.getRow(2);
  const get = (k: string) => cellVal(row.getCell(find(k)));
  return {
    zNumber: NaN,
    date: parseDateFR(get("Date") as string),
    cash: num(get("Cash")),
    carteBanque: num(get("Carte banque")),
    virementBancaire: num(get("Virement bancaire")),
    sourceFile: file.name,
  };
}

export function zFromFilename(name: string): number | null {
  const m = name.match(/_(\d+)\.xlsx?$/i);
  return m ? Number(m[1]) : null;
}

// ---------- Layout detection on a recap sheet ----------
export interface SheetLayout {
  headerRow: number;
  colJour: number;
  colZ: number;
  colTVAC: number;
  col21: number;
  col12: number;
  col6: number;
  colCartes: number;
  colVirement: number;
  colCash: number;
}

function headerMatch(ws: ExcelJS.Worksheet, row: number, keys: string[]): number {
  const r = ws.getRow(row);
  for (let c = 1; c <= Math.max(r.cellCount, 30); c++) {
    const t = headerText(r.getCell(c));
    if (t && keys.some(k => t.includes(k.toLowerCase()))) return c;
  }
  return -1;
}

export function detectLayout(ws: ExcelJS.Worksheet): SheetLayout | null {
  const headerRow = 2; // row 1 = title, row 2 = headers
  const layout: SheetLayout = {
    headerRow,
    colJour: headerMatch(ws, headerRow, ["jour"]),
    colZ: headerMatch(ws, headerRow, ["z n"]),
    colTVAC: headerMatch(ws, headerRow, ["total tvac"]),
    col21: headerMatch(ws, headerRow, ["total 21"]),
    col12: headerMatch(ws, headerRow, ["total 12"]),
    col6: headerMatch(ws, headerRow, ["total 6"]),
    colCartes: headerMatch(ws, headerRow, ["paiements cartes", "paiement cartes"]),
    colVirement: headerMatch(ws, headerRow, ["virement client"]),
    colCash: -1,
  };

  // Cash column: cell whose value equals exactly "CASH" (row 1 or row 2)
  const scanCash = (rowNum: number): number => {
    const r = ws.getRow(rowNum);
    for (let c = 1; c <= Math.max(r.cellCount, 30); c++) {
      const v = cellVal(r.getCell(c));
      if (v != null && String(v).toUpperCase().trim() === "CASH") return c;
    }
    return -1;
  };
  layout.colCash = scanCash(headerRow);
  if (layout.colCash < 0) layout.colCash = scanCash(1);

  if (layout.colJour < 0 || layout.colZ < 0 || layout.colTVAC < 0) return null;
  return layout;
}

// ---------- Build entries & reconcile ----------
export function buildEntries(zFiles: ZData[], caFiles: CAData[]): ZEntry[] {
  const caByZ = new Map<number, CAData>();
  const caByDate = new Map<string, CAData>();
  for (const ca of caFiles) {
    const z = zFromFilename(ca.sourceFile);
    if (z != null) caByZ.set(z, ca);
    caByDate.set(ca.date.toISOString().slice(0, 10), ca);
  }
  return zFiles.map(z => {
    const ca = caByZ.get(z.zNumber) ?? caByDate.get(z.openDate.toISOString().slice(0, 10)) ?? null;
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

const FIELD_TOLERANCE = 0.005;

function isEmpty(v: unknown): boolean {
  if (v == null || v === "") return true;
  if (typeof v === "number") return v === 0;
  const n = num(v);
  return n === 0 && String(v).trim() === "0";
}

export function computeReconciliation(
  entries: ZEntry[],
  wb: ExcelJS.Workbook
): ReconciliationRow[] {
  const out: ReconciliationRow[] = [];
  for (const e of entries) {
    const ws = findSheetForMonth(wb, e.monthIndex);
    if (!ws) continue;
    const layout = detectLayout(ws);
    if (!layout) continue;
    // Find row where col Jour value equals e.day
    let targetRow = -1;
    const maxRow = Math.min(ws.actualRowCount || ws.rowCount, 100);
    for (let r = layout.headerRow + 1; r <= maxRow; r++) {
      const dv = cellVal(ws.getRow(r).getCell(layout.colJour));
      if (dv != null && Number(dv) === e.day) {
        targetRow = r;
        break;
      }
    }
    if (targetRow < 0) continue;

    const proposed: ProposedValues = {
      zNumber: e.zNumber,
      totalTVAC: e.report.caTTC,
      total21: e.report.ttc21,
      total12: e.report.ttc12,
      total6: e.report.ttc6,
      cartes: e.ca?.carteBanque ?? 0,
      virement: e.ca?.virementBancaire ?? 0,
      cash: e.ca?.cash ?? 0,
    };
    const existing: Partial<ProposedValues> = {};
    const conflicts: string[] = [];
    const check = (key: keyof ProposedValues, col: number) => {
      if (col < 0) return;
      const cur = cellVal(ws.getRow(targetRow).getCell(col));
      if (!isEmpty(cur)) {
        existing[key] = num(cur);
        if (Math.abs(num(cur) - proposed[key]) > FIELD_TOLERANCE) conflicts.push(key);
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

    out.push({
      zNumber: e.zNumber,
      day: e.day,
      monthLabel: e.monthLabel,
      dateLabel: e.dateLabel,
      sheetName: ws.name,
      excelRow: targetRow,
      values: proposed,
      existing,
      conflicts,
      hasData: Object.keys(existing).length > 0,
      applied: false,
    });
  }
  return out;
}

/**
 * Apply reconciliation in-place. Only writes to cells that are empty.
 * Preserves all formatting, colors, merges, column widths — ExcelJS keeps existing cell styles.
 */
export function applyReconciliation(wb: ExcelJS.Workbook, rows: ReconciliationRow[]): ReconciliationRow[] {
  for (const r of rows) {
    const ws = wb.getWorksheet(r.sheetName);
    if (!ws) continue;
    const layout = detectLayout(ws);
    if (!layout) continue;
    const row = ws.getRow(r.excelRow);
    const writeIfEmpty = (col: number, value: number) => {
      if (col < 0) return;
      const cell = row.getCell(col);
      if (isEmpty(cellVal(cell))) {
        cell.value = value;
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
    row.commit();
    r.applied = true;
  }
  return rows;
}

export async function downloadWorkbook(wb: ExcelJS.Workbook, filename: string): Promise<void> {
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function cloneWorkbook(src: ExcelJS.Workbook): Promise<ExcelJS.Workbook> {
  const buf = await src.xlsx.writeBuffer();
  const dst = new ExcelJS.Workbook();
  await dst.xlsx.load(buf as ArrayBuffer);
  return dst;
}
