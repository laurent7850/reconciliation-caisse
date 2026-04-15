import { useCallback, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  applyReconciliation,
  buildEntries,
  computeReconciliation,
  downloadWorkbook,
  parseCA,
  parseReportZ,
  readSheetAOA,
  readWorkbook,
  type CAData,
  type ReconciliationRow,
  type ZData,
} from "./reconcile";

type Dropped = { file: File; kind: "recap" | "z" | "ca" | "other" };

function classifyFile(f: File): Dropped["kind"] {
  const n = f.name.toLowerCase();
  if (n.startsWith("reportzstats")) return "z";
  if (n.startsWith("ca_") && !n.startsWith("caby")) return "ca";
  if (n.includes("caisses") || n.includes("saint kilda") || n.includes("recap")) return "recap";
  return "other";
}

export default function App() {
  const [recapFile, setRecapFile] = useState<File | null>(null);
  const [recapWB, setRecapWB] = useState<XLSX.WorkBook | null>(null);
  const [zFiles, setZFiles] = useState<ZData[]>([]);
  const [caFiles, setCaFiles] = useState<CAData[]>([]);
  const [rows, setRows] = useState<ReconciliationRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [drag, setDrag] = useState<"recap" | "sources" | null>(null);

  const reset = () => {
    setRecapFile(null);
    setRecapWB(null);
    setZFiles([]);
    setCaFiles([]);
    setRows([]);
    setError(null);
  };

  const handleFiles = useCallback(async (files: FileList | File[], where: "recap" | "sources") => {
    setError(null);
    const arr = Array.from(files);
    try {
      if (where === "recap") {
        const f = arr[0];
        if (!f) return;
        const wb = await readWorkbook(f);
        setRecapFile(f);
        setRecapWB(wb);
        return;
      }
      const newZ: ZData[] = [...zFiles];
      const newCA: CAData[] = [...caFiles];
      for (const f of arr) {
        const kind = classifyFile(f);
        const rows0 = await readSheetAOA(f);
        if (kind === "z") newZ.push(parseReportZ(f, rows0));
        else if (kind === "ca") newCA.push(parseCA(f, rows0));
      }
      setZFiles(newZ);
      setCaFiles(newCA);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [zFiles, caFiles]);

  const onDrop = (where: "recap" | "sources") => (e: React.DragEvent) => {
    e.preventDefault();
    setDrag(null);
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files, where);
  };

  const entries = useMemo(() => buildEntries(zFiles, caFiles), [zFiles, caFiles]);

  const compute = () => {
    if (!recapWB) {
      setError("Charge d'abord le fichier récap.");
      return;
    }
    try {
      const r = computeReconciliation(entries, recapWB);
      setRows(r);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  const download = () => {
    if (!recapWB) return;
    // Clone workbook via write/read trick to avoid mutating preview
    const buf = XLSX.write(recapWB, { bookType: "xlsx", type: "array" });
    const clone = XLSX.read(buf);
    const applied = applyReconciliation(clone, rows);
    const stamp = new Date().toISOString().slice(0, 10);
    const fname = recapFile
      ? recapFile.name.replace(/\.xlsx?$/i, "") + `-reconcilie-${stamp}.xlsx`
      : `recap-reconcilie-${stamp}.xlsx`;
    downloadWorkbook(clone, fname);
    // Update visual state
    setRows(applied.map((r) => ({ ...r })));
  };

  const stats = useMemo(() => {
    const applicable = rows.filter(
      (r) => !r.hasData || (r.conflicts.length === 0 && r.hasData)
    ).length;
    const conflict = rows.filter((r) => r.conflicts.length > 0).length;
    const already = rows.filter((r) => r.hasData && r.conflicts.length === 0).length;
    return { total: rows.length, applicable, conflict, already };
  }, [rows]);

  return (
    <div className="container">
      <h1>Réconciliation Caisses — Saint Kilda</h1>
      <p className="subtitle">
        Importe le fichier récap annuel et les rapports Z du caissier. L'app détecte
        automatiquement le mois/jour et propose le remplissage. Les cellules déjà
        remplies ne sont <strong>jamais écrasées</strong> — les conflits sont signalés.
      </p>

      <div className="grid">
        <div
          className={`dropzone ${drag === "recap" ? "drag" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDrag("recap"); }}
          onDragLeave={() => setDrag(null)}
          onDrop={onDrop("recap")}
          onClick={() => document.getElementById("recap-input")?.click()}
        >
          <h3>1. Fichier récap annuel</h3>
          <p>
            Glisse <code>SAINT KILDA - CAISSES 2026.xlsx</code> ici (ou clique).
          </p>
          {recapFile && (
            <ul className="file-list"><li>✓ {recapFile.name}</li></ul>
          )}
          <input
            id="recap-input"
            type="file"
            accept=".xlsx,.xls"
            hidden
            onChange={(e) => e.target.files && handleFiles(e.target.files, "recap")}
          />
        </div>

        <div
          className={`dropzone ${drag === "sources" ? "drag" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDrag("sources"); }}
          onDragLeave={() => setDrag(null)}
          onDrop={onDrop("sources")}
          onClick={() => document.getElementById("src-input")?.click()}
        >
          <h3>2. Rapports Z + CA (un ou plusieurs jours)</h3>
          <p>
            Glisse <code>ReportZStats_1_N.xlsx</code> et <code>CA_1_N.xlsx</code>. Les autres fichiers sont ignorés.
          </p>
          <ul className="file-list">
            {zFiles.map((z) => (
              <li key={"z-" + z.zNumber}>✓ Z {z.zNumber} — {z.openDate.toLocaleDateString("fr-BE")} — {z.caTTC.toFixed(2)} €</li>
            ))}
            {caFiles.map((c) => (
              <li key={"ca-" + c.sourceFile}>✓ CA — {c.date.toLocaleDateString("fr-BE")} — cash {c.cash.toFixed(2)} / cartes {c.carteBanque.toFixed(2)}</li>
            ))}
          </ul>
          <input
            id="src-input"
            type="file"
            accept=".xlsx,.xls"
            multiple
            hidden
            onChange={(e) => e.target.files && handleFiles(e.target.files, "sources")}
          />
        </div>
      </div>

      <div className="toolbar">
        <button className="btn" onClick={compute} disabled={!recapWB || entries.length === 0}>
          Calculer la réconciliation
        </button>
        <button className="btn" onClick={download} disabled={rows.length === 0}>
          Télécharger le récap mis à jour
        </button>
        <button className="btn secondary" onClick={reset}>Tout effacer</button>
      </div>

      {error && <p style={{ color: "var(--err)" }}>⚠ {error}</p>}

      {rows.length > 0 && (
        <>
          <div className="summary">
            <div className="kpi"><div className="label">Entrées</div><div className="value">{stats.total}</div></div>
            <div className="kpi"><div className="label">À appliquer</div><div className="value" style={{ color: "var(--ok)" }}>{stats.applicable - stats.already}</div></div>
            <div className="kpi"><div className="label">Déjà OK</div><div className="value" style={{ color: "var(--muted)" }}>{stats.already}</div></div>
            <div className="kpi"><div className="label">Conflits</div><div className="value" style={{ color: "var(--err)" }}>{stats.conflict}</div></div>
          </div>

          <div className="section table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Statut</th>
                  <th>Mois</th>
                  <th>Jour</th>
                  <th>Z N°</th>
                  <th className="num">TOTAL TVAC</th>
                  <th className="num">21%</th>
                  <th className="num">12%</th>
                  <th className="num">6%</th>
                  <th className="num">Cartes</th>
                  <th className="num">Virement</th>
                  <th className="num">Cash</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => {
                  let badge = <span className="badge ok">Nouveau</span>;
                  if (r.conflicts.length > 0) badge = <span className="badge err">Conflit: {r.conflicts.join(", ")}</span>;
                  else if (r.hasData) badge = <span className="badge warn">Déjà rempli (cohérent)</span>;
                  if (r.applied) badge = <span className="badge ok">✓ Appliqué</span>;
                  return (
                    <tr key={r.zNumber + "-" + r.sheetName} className={r.conflicts.length ? "conflict" : ""}>
                      <td>{badge}</td>
                      <td>{r.monthLabel}</td>
                      <td>{r.day}</td>
                      <td>{r.values.zNumber}</td>
                      <td className="num">{r.values.totalTVAC.toFixed(2)}</td>
                      <td className="num">{r.values.total21.toFixed(2)}</td>
                      <td className="num">{r.values.total12.toFixed(2)}</td>
                      <td className="num">{r.values.total6.toFixed(2)}</td>
                      <td className="num">{r.values.cartes.toFixed(2)}</td>
                      <td className="num">{r.values.virement.toFixed(2)}</td>
                      <td className="num">{r.values.cash.toFixed(2)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {rows.some((r) => r.conflicts.length > 0) && (
            <details open>
              <summary>Détail des conflits</summary>
              <ul>
                {rows.filter((r) => r.conflicts.length > 0).map((r) => (
                  <li key={"c-" + r.zNumber}>
                    <strong>{r.monthLabel} jour {r.day} (Z {r.zNumber})</strong> —
                    {r.conflicts.map((k) => (
                      <span key={k}> <code>{k}</code>: récap <strong>{r.existing[k as keyof typeof r.existing] ?? "-"}</strong> vs Z <strong>{(r.values as any)[k]}</strong></span>
                    ))}
                  </li>
                ))}
              </ul>
            </details>
          )}
        </>
      )}

      <div className="section" style={{ color: "var(--muted)", fontSize: ".85rem" }}>
        Mapping: <code>Rapport Z → Z N°</code>, <code>CA TTC → TOTAL TVAC</code>,
        <code> TTC 21/12/6% → TOTAL 21/12/6% TVAC</code>,
        <code> Carte banque → Paiements cartes</code>,
        <code> Virement bancaire → Virement client</code>,
        <code> Cash → CASH</code>.
      </div>
    </div>
  );
}
