import JSZip from "jszip";

/**
 * Surgical XLSX editor: modifies cell values in specific rows without
 * touching anything else (styles, merges, formulas, shared strings, drawings, etc).
 *
 * An XLSX is a ZIP. Each worksheet is an XML file at xl/worksheets/sheetN.xml.
 * Cells look like: <c r="B12" s="71"/>              (empty with style)
 *                  <c r="B12" s="71"><v>442</v></c> (numeric value)
 *                  <c r="B12" s="5" t="s"><v>0</v></c>  (shared-string idx)
 * We only write numbers, and we do NOT alter the style attribute.
 */

export interface Patch {
  sheetName: string;
  row: number; // 1-indexed
  col: number; // 1-indexed (1=A, 2=B, ...)
  value: number;
}

function colLetter(col: number): string {
  let s = "";
  let n = col;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

interface SheetRef {
  name: string;
  path: string; // e.g. "xl/worksheets/sheet5.xml"
}

async function listSheets(zip: JSZip): Promise<SheetRef[]> {
  const wbXml = await zip.file("xl/workbook.xml")!.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");

  // Build rId → target map
  const rMap = new Map<string, string>();
  for (const m of relsXml.matchAll(/<Relationship\s+([^>]*)\/>/g)) {
    const attrs = m[1];
    const id = /Id="([^"]+)"/.exec(attrs)?.[1];
    const type = /Type="([^"]+)"/.exec(attrs)?.[1];
    const target = /Target="([^"]+)"/.exec(attrs)?.[1];
    if (id && target && type && type.endsWith("/worksheet")) {
      // Target is relative to xl/, resolve
      const path = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
      rMap.set(id, path);
    }
  }

  const sheets: SheetRef[] = [];
  for (const m of wbXml.matchAll(/<sheet\s+([^/]+)\/>/g)) {
    const attrs = m[1];
    const name = /name="([^"]+)"/.exec(attrs)?.[1];
    const rid = /r:id="([^"]+)"/.exec(attrs)?.[1];
    if (name && rid && rMap.has(rid)) {
      sheets.push({ name, path: rMap.get(rid)! });
    }
  }
  return sheets;
}

function patchSheetXml(xml: string, patches: Patch[]): string {
  // Group patches by row
  const byRow = new Map<number, Patch[]>();
  for (const p of patches) {
    if (!byRow.has(p.row)) byRow.set(p.row, []);
    byRow.get(p.row)!.push(p);
  }

  let out = xml;
  for (const [rowNum, rowPatches] of byRow) {
    // Match the full <row r="N" ...>...</row>
    const rowRegex = new RegExp(
      `(<row\\s[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`,
      "m"
    );
    const match = rowRegex.exec(out);
    if (!match) {
      // Row doesn't exist — skip (shouldn't happen, the recap has all day rows)
      continue;
    }
    const rowOpen = match[1];
    let rowBody = match[2];
    const rowClose = match[3];

    for (const p of rowPatches) {
      const cellRef = colLetter(p.col) + p.row;
      // Try to match the existing <c> for this cell
      // Variants: <c r="X" ... /> | <c r="X" ...></c> | <c r="X" ...><v>...</v></c>
      const cellRegex = new RegExp(
        `<c\\s+r="${cellRef}"([^/>]*)(/>|>([\\s\\S]*?)</c>)`,
        "m"
      );
      const cm = cellRegex.exec(rowBody);
      const valueXml = `<v>${p.value}</v>`;
      if (cm) {
        // Preserve attributes (style s="..."), but remove t="s" (shared string) if present
        let attrs = cm[1];
        attrs = attrs.replace(/\s+t="[^"]*"/, ""); // force numeric type
        const replacement = `<c r="${cellRef}"${attrs}>${valueXml}</c>`;
        rowBody = rowBody.slice(0, cm.index) + replacement + rowBody.slice(cm.index + cm[0].length);
      } else {
        // Cell doesn't exist — insert at the right position
        // Cells must be ordered by column letter
        const newCell = `<c r="${cellRef}">${valueXml}</c>`;
        // Find first cell whose column letter > ours and insert before it
        const insRegex = /<c\s+r="([A-Z]+)(\d+)"/g;
        let insertAt = rowBody.length;
        let m2: RegExpExecArray | null;
        while ((m2 = insRegex.exec(rowBody))) {
          if (compareColLetters(m2[1], colLetter(p.col)) > 0) {
            insertAt = m2.index;
            break;
          }
        }
        rowBody = rowBody.slice(0, insertAt) + newCell + rowBody.slice(insertAt);
      }
    }
    out = out.slice(0, match.index) + rowOpen + rowBody + rowClose + out.slice(match.index + match[0].length);
  }

  return out;
}

function compareColLetters(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  return a.localeCompare(b);
}

/**
 * Force Excel to recalculate all formulas when the file is opened.
 * Sets fullCalcOnLoad="1" on <calcPr> in xl/workbook.xml. Without this,
 * formula cells keep their stale cached <v> after we patch a referenced cell.
 */
function forceRecalcOnLoad(wbXml: string): string {
  if (/<calcPr\b[^>]*\bfullCalcOnLoad="1"/.test(wbXml)) return wbXml;
  if (/<calcPr\b/.test(wbXml)) {
    return wbXml.replace(/<calcPr\b([^/>]*)(\/?>)/, (_m, attrs, end) => {
      const cleaned = attrs.replace(/\s+fullCalcOnLoad="[^"]*"/, "");
      return `<calcPr${cleaned} fullCalcOnLoad="1"${end}`;
    });
  }
  // No <calcPr> — insert one before </workbook>
  return wbXml.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>');
}

/**
 * Apply a list of cell patches to an XLSX ArrayBuffer. Returns new ArrayBuffer.
 * Only the targeted cell values are modified. Everything else (styles, merges,
 * column widths, colors, formulas, drawings, pivot caches...) is left untouched.
 */
export async function patchXlsx(
  original: ArrayBuffer,
  patches: Patch[]
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(original);
  const sheets = await listSheets(zip);
  const byName = new Map<string, string>();
  for (const s of sheets) byName.set(s.name, s.path);

  // Group patches by sheet name
  const bySheet = new Map<string, Patch[]>();
  for (const p of patches) {
    if (!bySheet.has(p.sheetName)) bySheet.set(p.sheetName, []);
    bySheet.get(p.sheetName)!.push(p);
  }

  for (const [sheetName, ps] of bySheet) {
    const path = byName.get(sheetName);
    if (!path) continue;
    const xml = await zip.file(path)!.async("string");
    const patched = patchSheetXml(xml, ps);
    zip.file(path, patched);
  }

  // Force formula recalculation on next open so RECAP ANNUEL etc. refresh.
  const wbFile = zip.file("xl/workbook.xml");
  if (wbFile) {
    const wbXml = await wbFile.async("string");
    zip.file("xl/workbook.xml", forceRecalcOnLoad(wbXml));
  }

  // Write ZIP back
  const out = await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
    // preserve ZIP layout as much as possible
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  return out;
}
