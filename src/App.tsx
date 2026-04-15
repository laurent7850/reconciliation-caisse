import { useCallback, useMemo, useState } from "react";
import ExcelJS from "exceljs";
import {
  buildEntries,
  buildPatches,
  computeReconciliation,
  downloadPatched,
  FIELD_LABELS,
  loadWorkbook,
  parseCA,
  parseReportZ,
  type CAData,
  type ProposedValues,
  type ReconciliationRow,
  type ZData,
} from "./reconcile";

function fmt(k: keyof ProposedValues, v: number): string {
  if (k === "zNumber") return String(v);
  return v.toLocaleString("fr-BE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
}

function conflictExplanation(k: keyof ProposedValues): string {
  if (k === "zNumber") return "le numéro de Z saisi dans le récap ne correspond pas au rapport importé";
  return "la valeur saisie dans le récap ne correspond pas au rapport Z";
}

type Kind = "recap" | "z" | "ca" | "other";

function classifyFile(f: File): Kind {
  const n = f.name.toLowerCase();
  if (n.startsWith("reportzstats")) return "z";
  if (n.startsWith("ca_") && !n.startsWith("caby")) return "ca";
  if (n.includes("caisses") || n.includes("saint kilda") || n.includes("recap")) return "recap";
  return "other";
}

export default function App() {
  const [recapFile, setRecapFile] = useState<File | null>(null);
  const [recapWB, setRecapWB] = useState<ExcelJS.Workbook | null>(null);
  const [recapBytes, setRecapBytes] = useState<ArrayBuffer | null>(null);
  const [zFiles, setZFiles] = useState<ZData[]>([]);
  const [caFiles, setCaFiles] = useState<CAData[]>([]);
  const [rows, setRows] = useState<ReconciliationRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [drag, setDrag] = useState<"recap" | "sources" | null>(null);
  const [busy, setBusy] = useState(false);

  const reset = () => {
    setRecapFile(null);
    setRecapWB(null);
    setRecapBytes(null);
    setZFiles([]);
    setCaFiles([]);
    setRows([]);
    setError(null);
  };

  const handleFiles = useCallback(
    async (files: FileList | File[], where: "recap" | "sources") => {
      setError(null);
      setBusy(true);
      try {
        const arr = Array.from(files);
        if (where === "recap") {
          const f = arr[0];
          if (!f) return;
          const bytes = await f.arrayBuffer();
          // load a copy so the original ArrayBuffer is preserved for surgical patching
          const wb = await loadWorkbook(new File([bytes.slice(0)], f.name));
          setRecapFile(f);
          setRecapBytes(bytes);
          setRecapWB(wb);
          return;
        }
        const newZ: ZData[] = [...zFiles];
        const newCA: CAData[] = [...caFiles];
        for (const f of arr) {
          const kind = classifyFile(f);
          if (kind === "z") newZ.push(await parseReportZ(f));
          else if (kind === "ca") newCA.push(await parseCA(f));
        }
        setZFiles(newZ);
        setCaFiles(newCA);
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setBusy(false);
      }
    },
    [zFiles, caFiles]
  );

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

  const download = async () => {
    if (!recapWB || !recapBytes) return;
    setBusy(true);
    try {
      const patches = buildPatches(recapWB, rows);
      const stamp = new Date().toISOString().slice(0, 10);
      const fname = recapFile
        ? recapFile.name.replace(/\.xlsx?$/i, "") + `-reconcilie-${stamp}.xlsx`
        : `recap-reconcilie-${stamp}.xlsx`;
      // Pass a fresh copy of the bytes (JSZip may consume them)
      await downloadPatched(recapBytes.slice(0), patches, fname);
      setRows(rows.map((r) => ({ ...r })));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const stats = useMemo(() => {
    const conflict = rows.filter((r) => r.conflicts.length > 0).length;
    const already = rows.filter((r) => r.hasData && r.conflicts.length === 0).length;
    const toFill = rows.length - conflict - already;
    return { total: rows.length, toFill, already, conflict };
  }, [rows]);

  return (
    <div className="container">
      <h1>Réconciliation Caisses — Saint Kilda</h1>
      <p className="subtitle">
        L'app remplit <strong>uniquement les cellules vides</strong> du fichier récap existant.
        Toute la mise en page (couleurs, fusions, formules, largeurs, bordures) est préservée.
        Les cellules déjà remplies ne sont jamais écrasées, les conflits sont signalés.
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
          <p>Glisse <code>SAINT KILDA - CAISSES 2026.xlsx</code> ici (ou clique).</p>
          {recapFile && <ul className="file-list"><li>✓ {recapFile.name}</li></ul>}
          <input
            id="recap-input" type="file" accept=".xlsx,.xls" hidden
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
          <h3>2. Rapports Z + CA</h3>
          <p>Glisse <code>ReportZStats_1_N.xlsx</code> et <code>CA_1_N.xlsx</code> (les autres exports sont ignorés).</p>
          <ul className="file-list">
            {zFiles.map((z) => (
              <li key={"z-" + z.zNumber}>✓ Z {z.zNumber} — {z.openDate.toLocaleDateString("fr-BE")} — {z.caTTC.toFixed(2)} €</li>
            ))}
            {caFiles.map((c) => (
              <li key={"ca-" + c.sourceFile}>✓ CA — {c.date.toLocaleDateString("fr-BE")} — cash {c.cash.toFixed(2)} / cartes {c.carteBanque.toFixed(2)}</li>
            ))}
          </ul>
          <input
            id="src-input" type="file" accept=".xlsx,.xls" multiple hidden
            onChange={(e) => e.target.files && handleFiles(e.target.files, "sources")}
          />
        </div>
      </div>

      <div className="toolbar">
        <button className="btn" onClick={compute} disabled={!recapWB || entries.length === 0 || busy}>
          Calculer la réconciliation
        </button>
        <button className="btn" onClick={download} disabled={rows.length === 0 || busy}>
          Télécharger le récap mis à jour
        </button>
        <button className="btn secondary" onClick={reset} disabled={busy}>Tout effacer</button>
        {busy && <span style={{ color: "var(--muted)" }}>…</span>}
      </div>

      {error && <p style={{ color: "var(--err)" }}>⚠ {error}</p>}

      {rows.length > 0 && (
        <>
          <div className="summary">
            <div className="kpi"><div className="label">Entrées</div><div className="value">{stats.total}</div></div>
            <div className="kpi"><div className="label">À remplir</div><div className="value" style={{ color: "var(--ok)" }}>{stats.toFill}</div></div>
            <div className="kpi"><div className="label">Déjà OK</div><div className="value" style={{ color: "var(--muted)" }}>{stats.already}</div></div>
            <div className="kpi"><div className="label">Conflits</div><div className="value" style={{ color: "var(--err)" }}>{stats.conflict}</div></div>
          </div>

          <div className="section table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Statut</th><th>Mois</th><th>Jour</th><th>Z N°</th>
                  <th className="num">TOTAL TVAC</th><th className="num">21%</th>
                  <th className="num">12%</th><th className="num">6%</th>
                  <th className="num">Cartes</th><th className="num">Virement</th><th className="num">Cash</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => {
                  let badge = <span className="badge ok">À remplir</span>;
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
            <details open className="conflicts-panel">
              <summary>
                ⚠ {rows.filter((r) => r.conflicts.length > 0).length} conflit(s) détecté(s) — à vérifier manuellement
              </summary>
              <p className="conflicts-intro">
                Les cellules ci-dessous contiennent déjà une valeur dans ton récap, différente de celle du rapport Z.
                <strong> Elles n'ont PAS été modifiées.</strong> Pour chaque ligne, vérifie qui a raison :
                si le récap est faux, corrige-le manuellement (ou vide la cellule puis relance l'import).
              </p>
              <div className="conflicts-list">
                {rows.filter((r) => r.conflicts.length > 0).map((r) => (
                  <div key={"c-" + r.zNumber} className="conflict-card">
                    <div className="conflict-header">
                      📅 <strong>{r.monthLabel}</strong> — jour <strong>{r.day}</strong>{" "}
                      <span style={{ color: "var(--muted)" }}>
                        (Z {r.zNumber} du {r.dateLabel})
                      </span>
                    </div>
                    <table className="conflict-table">
                      <thead>
                        <tr>
                          <th>Champ</th>
                          <th>Dans ton récap</th>
                          <th>Dans le rapport Z</th>
                          <th>Remarque</th>
                        </tr>
                      </thead>
                      <tbody>
                        {r.conflicts.map((k) => (
                          <tr key={k}>
                            <td><strong>{FIELD_LABELS[k]}</strong></td>
                            <td className="val-recap">{fmt(k, r.existing[k] as number)}</td>
                            <td className="val-z">{fmt(k, r.values[k])}</td>
                            <td style={{ color: "var(--muted)", fontSize: ".85em" }}>
                              {conflictExplanation(k)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
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
