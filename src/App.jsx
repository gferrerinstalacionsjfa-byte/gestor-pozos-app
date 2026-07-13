import React, { useState, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { Upload, Droplets, AlertTriangle, Download, ChevronDown, ChevronRight, Search, Waves, CheckCircle2, RefreshCw } from "lucide-react";

const BAR_TO_MH2O = 10.19716;
const ASSOC_STORAGE_KEY = "pozos-assoc-v1";

function normEUI(s) {
  return String(s || "").replace(/[^0-9a-fA-F]/g, "").toLowerCase();
}

function parseCableFinsCota(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  const m = s.match(/^(\d+(?:[.,]\d+)?)\s*m\.?$/i);
  if (!m) return null;
  return parseFloat(m[1].replace(",", "."));
}

function findCol(headers, patterns) {
  for (const pat of patterns) {
    const idx = headers.findIndex((h) => pat.test(String(h || "").trim()));
    if (idx !== -1) return idx;
  }
  return -1;
}

function decodePayload(hex) {
  if (!hex || typeof hex !== "string") return null;
  const clean = hex.replace(/\s/g, "");
  if (!/^[0-9a-fA-F]+$/.test(clean) || clean.length < 16) return null;
  const bytes = [];
  for (let i = 0; i < clean.length; i += 2) bytes.push(parseInt(clean.substr(i, 2), 16));

  const out = {};
  for (let i = 0; i + 8 <= bytes.length; i++) {
    if (bytes[i] === 0xff && bytes[i + 1] === 0x0e) {
      const channel = bytes[i + 2];
      const buf = new ArrayBuffer(4);
      const view = new DataView(buf);
      for (let k = 0; k < 4; k++) view.setUint8(k, bytes[i + 4 + k]);
      const val = view.getFloat32(0, true); // little-endian
      if (Number.isFinite(val)) {
        if (channel === 7) out.pressureBar = val;
        else if (channel === 8) out.tempC = val;
        else if (channel === 9) out.condMScm = val;
      }
      i += 7; // skip consumed bytes (loop will +1 more)
    }
  }
  return Object.keys(out).length ? out : null;
}

function findSheetWithHeader(workbook, patterns) {
  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false });
    if (!rows.length) continue;
    const headers = rows[0].map((h) => String(h || "").trim());
    const hasAll = patterns.every((pat) => headers.some((h) => pat.test(h)));
    if (hasAll) return { rows, headers };
  }
  return null;
}

async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: "array", cellDates: false });
}

function fmt(n, d = 2) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  return n.toLocaleString("es-ES", { minimumFractionDigits: d, maximumFractionDigits: d });
}

async function parseAssociationWorkbook(file) {
  const assocWb = await readWorkbook(file);
  const assocSheet = findSheetWithHeader(assocWb, [/dev\s*eui/i, /cable\s*fins\s*cota/i]);
  if (!assocSheet) {
    throw new Error('No encuentro una hoja con columnas "DEV EUI" y "Cable fins cota" en el archivo de asociación.');
  }
  const aHeaders = assocSheet.headers;
  const euiIdx = findCol(aHeaders, [/dev\s*eui/i, /device_eui/i]);
  const pozoIdx = findCol(aHeaders, [/c[oó]digo\s*pozo/i, /^codi$/i]);
  const cotaIdx = findCol(aHeaders, [/^cota$/i]);
  const cableIdx = findCol(aHeaders, [/cable\s*fins\s*cota/i]);

  const assocMap = new Map();
  const excluded = [];
  for (let r = 1; r < assocSheet.rows.length; r++) {
    const row = assocSheet.rows[r];
    if (!row || !row.length) continue;
    const euiRaw = row[euiIdx];
    const eui = normEUI(euiRaw);
    if (!eui) continue;
    const pozo = (pozoIdx !== -1 ? row[pozoIdx] : null) || eui;
    const cota = cotaIdx !== -1 ? parseFloat(String(row[cotaIdx]).replace(",", ".")) : null;
    const cableRaw = cableIdx !== -1 ? row[cableIdx] : null;
    const cable = parseCableFinsCota(cableRaw);
    if (cable === null) {
      excluded.push({ pozo, eui, motivo: cableRaw ? `valor no reconocido: "${cableRaw}"` : "sin valor" });
      continue;
    }
    assocMap.set(eui, { pozo: String(pozo), cota, cable });
  }
  return { map: assocMap, excluded };
}

function loadAssocFromStorage() {
  try {
    const raw = localStorage.getItem(ASSOC_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return {
      map: new Map(parsed.map),
      excluded: parsed.excluded || [],
      fileName: parsed.fileName,
      savedAt: parsed.savedAt,
    };
  } catch (e) {
    return null;
  }
}

function saveAssocToStorage(data) {
  try {
    localStorage.setItem(
      ASSOC_STORAGE_KEY,
      JSON.stringify({
        map: Array.from(data.map.entries()),
        excluded: data.excluded,
        fileName: data.fileName,
        savedAt: data.savedAt,
      })
    );
    return true;
  } catch (e) {
    return false; // p.ej. no disponible en la vista previa de Claude.ai
  }
}

export default function App() {
  const [assocData, setAssocData] = useState(null); // { map, excluded, fileName, savedAt }
  const [assocLoading, setAssocLoading] = useState(false);
  const [assocError, setAssocError] = useState(null);
  const [assocPersisted, setAssocPersisted] = useState(true);
  const [readingsFile, setReadingsFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [results, setResults] = useState(null); // { groups, excluded, stats }
  const [query, setQuery] = useState("");
  const [collapsed, setCollapsed] = useState({});
  const [showExcluded, setShowExcluded] = useState(false);

  useEffect(() => {
    const saved = loadAssocFromStorage();
    if (saved) setAssocData(saved);
  }, []);

  const handleAssocFile = useCallback(async (file) => {
    setAssocLoading(true);
    setAssocError(null);
    try {
      const { map, excluded } = await parseAssociationWorkbook(file);
      const data = { map, excluded, fileName: file.name, savedAt: new Date().toISOString() };
      const persisted = saveAssocToStorage(data);
      setAssocPersisted(persisted);
      setAssocData(data);
      setResults(null);
    } catch (e) {
      setAssocError(e.message || String(e));
    } finally {
      setAssocLoading(false);
    }
  }, []);

  const clearAssocFile = () => {
    try {
      localStorage.removeItem(ASSOC_STORAGE_KEY);
    } catch (e) {}
    setAssocData(null);
    setResults(null);
  };

  const process = useCallback(async (assocMap, assocExcluded, readings) => {
    setLoading(true);
    setError(null);
    try {
      const readWb = await readWorkbook(readings);
      const readSheet = findSheetWithHeader(readWb, [/dev\s*eui/i, /payload/i]);
      if (!readSheet) {
        throw new Error('No encuentro una hoja con columnas "DevEUI" y "Payload (HEX)" en el archivo de lecturas.');
      }
      const rHeaders = readSheet.headers;
      const rEuiIdx = findCol(rHeaders, [/dev\s*eui/i]);
      const rTsIdx = findCol(rHeaders, [/marca\s*de\s*temps/i, /timestamp/i]);
      const rPayloadIdx = findCol(rHeaders, [/payload/i]);

      const groups = new Map(); // pozo -> date -> {pressures,temps,conds,nivels,count}
      let totalReadings = 0;
      let matchedReadings = 0;
      let decodedReadings = 0;

      for (let r = 1; r < readSheet.rows.length; r++) {
        const row = readSheet.rows[r];
        if (!row || !row.length) continue;
        const eui = normEUI(row[rEuiIdx]);
        if (!eui) continue;
        totalReadings++;
        const info = assocMap.get(eui);
        if (!info) continue;
        matchedReadings++;
        const decoded = decodePayload(row[rPayloadIdx]);
        if (!decoded || decoded.pressureBar === undefined) continue;
        decodedReadings++;

        const ts = row[rTsIdx];
        const date = String(ts || "").slice(0, 10) || "sin-fecha";
        const mH2O = decoded.pressureBar * BAR_TO_MH2O;
        const nivel = info.cable - mH2O;

        if (!groups.has(info.pozo)) groups.set(info.pozo, new Map());
        const byDate = groups.get(info.pozo);
        if (!byDate.has(date)) {
          byDate.set(date, { pressures: [], temps: [], conds: [], nivels: [], cota: info.cota, cable: info.cable });
        }
        const bucket = byDate.get(date);
        bucket.pressures.push(decoded.pressureBar);
        if (decoded.tempC !== undefined) bucket.temps.push(decoded.tempC);
        if (decoded.condMScm !== undefined) bucket.conds.push(decoded.condMScm);
        bucket.nivels.push(nivel);
      }

      const avg = (arr) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null);

      const groupList = Array.from(groups.entries())
        .map(([pozo, byDate]) => {
          const dates = Array.from(byDate.entries())
            .map(([date, b]) => ({
              date,
              n: b.pressures.length,
              pressureBar: avg(b.pressures),
              pressureMH2O: avg(b.pressures) * BAR_TO_MH2O,
              tempC: avg(b.temps),
              condMScm: avg(b.conds),
              nivel: avg(b.nivels),
              cota: b.cota,
              cable: b.cable,
            }))
            .sort((a, b) => (a.date < b.date ? 1 : -1));
          return { pozo, dates };
        })
        .sort((a, b) => a.pozo.localeCompare(b.pozo));

      setResults({
        groups: groupList,
        excluded: [...assocExcluded].sort((a, b) => a.pozo.localeCompare(b.pozo)),
        stats: { totalReadings, matchedReadings, decodedReadings, wellsOk: assocMap.size, wellsExcluded: assocExcluded.length },
      });
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  const handleFile = (setter) => (e) => {
    const f = e.target.files?.[0];
    if (f) setter(f);
  };

  const canProcess = assocData && readingsFile && !loading;

  const filteredGroups = useMemo(() => {
    if (!results) return [];
    if (!query.trim()) return results.groups;
    const q = query.trim().toLowerCase();
    return results.groups.filter((g) => g.pozo.toLowerCase().includes(q));
  }, [results, query]);

  const exportExcel = () => {
    if (!results) return;
    const rows = [];
    results.groups.forEach((g) => {
      g.dates.forEach((d) => {
        rows.push({
          Pozo: g.pozo,
          Fecha: d.date,
          "Lecturas": d.n,
          "Presión media (bar)": +d.pressureBar.toFixed(4),
          "Presión media (mH2O)": +d.pressureMH2O.toFixed(4),
          "Temperatura media (°C)": d.tempC !== null ? +d.tempC.toFixed(2) : "",
          "Conductividad media (mS/cm)": d.condMScm !== null ? +d.condMScm.toFixed(3) : "",
          "Cota (m)": d.cota,
          "Cable fins cota (m)": d.cable,
          "Nivel freático (m)": +d.nivel.toFixed(3),
        });
      });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Nivel freático");

    if (results.excluded.length) {
      const wsExc = XLSX.utils.json_to_sheet(
        results.excluded.map((x) => ({ Pozo: x.pozo, DevEUI: x.eui, Motivo: x.motivo }))
      );
      XLSX.utils.book_append_sheet(wb, wsExc, "Pozos excluidos");
    }
    XLSX.writeFile(wb, `nivel_freatico_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const toggle = (pozo) => setCollapsed((c) => ({ ...c, [pozo]: !c[pozo] }));

  return (
    <div style={styles.page}>
      <div style={styles.headerBar}>
        <div style={styles.brand}>
          <Waves size={22} color="#e8f3f2" />
          <span style={styles.brandText}>Nivell Freàtic</span>
        </div>
        <span style={styles.brandSub}>Gestor de sondes LoRaWAN en pous</span>
      </div>

      <div style={styles.container}>
        <div style={styles.uploadRow}>
          {assocData ? (
            <div style={styles.assocLoadedCard}>
              <CheckCircle2 size={20} color="#2a8f6c" />
              <div style={{ flex: 1 }}>
                <div style={styles.uploadLabel}>Fitxer d'associació carregat</div>
                <div style={styles.uploadHint}>
                  {assocData.fileName} · desat el {new Date(assocData.savedAt).toLocaleString("es-ES")}
                  {!assocPersisted && " · no s'ha pogut desar per a la propera visita"}
                </div>
              </div>
              <label htmlFor="assoc-input" style={styles.smallLinkBtn}>
                <RefreshCw size={13} style={{ marginRight: 4, verticalAlign: "-2px" }} />
                Canviar
              </label>
              <input
                id="assoc-input"
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleAssocFile(f);
                }}
              />
            </div>
          ) : (
            <UploadCard
              label="1. Fitxer d'associació (pou ↔ DevEUI)"
              hint={assocLoading ? "Carregant…" : "Excel amb columnes DEV EUI, Código pozo, COTA, Cable fins cota"}
              file={null}
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleAssocFile(f);
              }}
              inputId="assoc-input"
            />
          )}
          <UploadCard
            label="2. Fitxer de lectures"
            hint="Excel export amb columnes DevEUI, Marca de temps, Payload (HEX)"
            file={readingsFile}
            onChange={handleFile(setReadingsFile)}
            inputId="readings-input"
          />
        </div>

        {assocData && (
          <button style={styles.tinyClearBtn} onClick={clearAssocFile}>
            Eliminar fitxer d'associació desat
          </button>
        )}

        <div style={styles.actionRow}>
          <button
            style={{ ...styles.primaryBtn, ...(canProcess ? {} : styles.btnDisabled) }}
            disabled={!canProcess}
            onClick={() => process(assocData.map, assocData.excluded, readingsFile)}
          >
            {loading ? "Processant…" : "Processar dades"}
          </button>
          {results && (
            <button style={styles.secondaryBtn} onClick={exportExcel}>
              <Download size={16} style={{ marginRight: 6, verticalAlign: "-3px" }} />
              Exportar a Excel
            </button>
          )}
        </div>

        {(error || assocError) && (
          <div style={styles.errorBox}>
            <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>{error || assocError}</span>
          </div>
        )}

        {results && (
          <>
            <div style={styles.statsRow}>
              <Stat label="Pous amb cota vàlida" value={results.stats.wellsOk} />
              <Stat label="Pous exclosos" value={results.stats.wellsExcluded} warn />
              <Stat label="Lectures totals" value={results.stats.totalReadings} />
              <Stat label="Lectures amb pou associat" value={results.stats.matchedReadings} />
              <Stat label="Lectures decodificades" value={results.stats.decodedReadings} />
            </div>

            <div style={styles.searchBar}>
              <Search size={16} color="#7c9490" />
              <input
                style={styles.searchInput}
                placeholder="Filtrar per codi de pou…"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
            </div>

            <div style={styles.groupsWrap}>
              {filteredGroups.length === 0 && (
                <div style={styles.emptyMsg}>Cap pou coincideix amb aquest filtre.</div>
              )}
              {filteredGroups.map((g) => (
                <div key={g.pozo} style={styles.groupCard}>
                  <button style={styles.groupHeader} onClick={() => toggle(g.pozo)}>
                    {collapsed[g.pozo] ? <ChevronRight size={16} /> : <ChevronDown size={16} />}
                    <Droplets size={16} color="#3e8e86" style={{ marginRight: 6 }} />
                    <span style={styles.groupTitle}>{g.pozo}</span>
                    <span style={styles.groupMeta}>
                      {g.dates.length} {g.dates.length === 1 ? "dia" : "dies"} · cota {fmt(g.dates[0]?.cota, 1)} m ·
                      cable {fmt(g.dates[0]?.cable, 2)} m
                    </span>
                  </button>
                  {!collapsed[g.pozo] && (
                    <div style={styles.tableWrap}>
                      <table style={styles.table}>
                        <thead>
                          <tr>
                            <th style={styles.th}>Data</th>
                            <th style={styles.th}>Lectures</th>
                            <th style={styles.th}>Pressió (bar)</th>
                            <th style={styles.th}>Pressió (mH2O)</th>
                            <th style={styles.th}>Temp. (°C)</th>
                            <th style={styles.th}>Conductivitat (mS/cm)</th>
                            <th style={{ ...styles.th, color: "#2a6f68" }}>Nivell freàtic (m)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {g.dates.map((d) => (
                            <tr key={d.date}>
                              <td style={styles.td}>{d.date}</td>
                              <td style={styles.td}>{d.n}</td>
                              <td style={styles.td}>{fmt(d.pressureBar, 4)}</td>
                              <td style={styles.td}>{fmt(d.pressureMH2O, 3)}</td>
                              <td style={styles.td}>{fmt(d.tempC, 2)}</td>
                              <td style={styles.td}>{d.condMScm !== null ? fmt(d.condMScm, 3) : "—"}</td>
                              <td style={{ ...styles.td, fontWeight: 600, color: "#1f4b46" }}>{fmt(d.nivel, 3)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              ))}
            </div>

            {results.excluded.length > 0 && (
              <div style={styles.excludedBox}>
                <button style={styles.excludedHeader} onClick={() => setShowExcluded((s) => !s)}>
                  {showExcluded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  <AlertTriangle size={15} color="#a86a2d" style={{ margin: "0 6px" }} />
                  {results.excluded.length} pous exclosos per falta de "Cable fins cota" vàlid
                </button>
                {showExcluded && (
                  <div style={styles.tableWrap}>
                    <table style={styles.table}>
                      <thead>
                        <tr>
                          <th style={styles.th}>Pou</th>
                          <th style={styles.th}>DevEUI</th>
                          <th style={styles.th}>Motiu</th>
                        </tr>
                      </thead>
                      <tbody>
                        {results.excluded.map((x, i) => (
                          <tr key={i}>
                            <td style={styles.td}>{x.pozo}</td>
                            <td style={{ ...styles.td, fontFamily: "monospace", fontSize: 12 }}>{x.eui}</td>
                            <td style={styles.td}>{x.motivo}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function UploadCard({ label, hint, file, onChange, inputId }) {
  return (
    <label htmlFor={inputId} style={{ ...styles.uploadCard, ...(file ? styles.uploadCardFilled : {}) }}>
      <Upload size={20} color={file ? "#3e8e86" : "#7c9490"} />
      <div style={{ flex: 1 }}>
        <div style={styles.uploadLabel}>{label}</div>
        <div style={styles.uploadHint}>{file ? file.name : hint}</div>
      </div>
      <input id={inputId} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={onChange} />
    </label>
  );
}

function Stat({ label, value, warn }) {
  return (
    <div style={styles.statCard}>
      <div style={{ ...styles.statValue, color: warn && value > 0 ? "#a86a2d" : "#1f4b46" }}>{value}</div>
      <div style={styles.statLabel}>{label}</div>
    </div>
  );
}

const styles = {
  page: {
    minHeight: "100vh",
    background: "#f4f8f7",
    fontFamily: "'Segoe UI', Roboto, -apple-system, sans-serif",
    color: "#1c2b29",
  },
  headerBar: {
    background: "linear-gradient(120deg, #123c3a, #1f5e59)",
    padding: "22px 28px",
    display: "flex",
    alignItems: "baseline",
    gap: 14,
  },
  brand: { display: "flex", alignItems: "center", gap: 8 },
  brandText: { color: "#e8f3f2", fontSize: 20, fontWeight: 700, letterSpacing: 0.3 },
  brandSub: { color: "#9fc4c0", fontSize: 13 },
  container: { maxWidth: 1100, margin: "0 auto", padding: "24px 20px 60px" },
  uploadRow: { display: "flex", gap: 16, flexWrap: "wrap" },
  uploadCard: {
    flex: "1 1 320px",
    display: "flex",
    alignItems: "center",
    gap: 12,
    background: "#fff",
    border: "1.5px dashed #cfe0dd",
    borderRadius: 10,
    padding: "14px 16px",
    cursor: "pointer",
    transition: "border-color .15s",
  },
  uploadCardFilled: { borderColor: "#3e8e86", borderStyle: "solid" },
  assocLoadedCard: {
    flex: "1 1 320px",
    display: "flex",
    alignItems: "center",
    gap: 12,
    background: "#f2faf6",
    border: "1.5px solid #bfe3d0",
    borderRadius: 10,
    padding: "14px 16px",
  },
  smallLinkBtn: {
    fontSize: 12.5,
    fontWeight: 600,
    color: "#1f5e59",
    cursor: "pointer",
    whiteSpace: "nowrap",
    padding: "6px 10px",
    border: "1px solid #cfe0dd",
    borderRadius: 6,
    background: "#fff",
  },
  tinyClearBtn: {
    marginTop: 8,
    background: "none",
    border: "none",
    color: "#a86a2d",
    fontSize: 12,
    cursor: "pointer",
    padding: 0,
    textDecoration: "underline",
  },
  uploadLabel: { fontSize: 13.5, fontWeight: 600, color: "#1c2b29" },
  uploadHint: { fontSize: 12.5, color: "#7c9490", marginTop: 2, wordBreak: "break-all" },
  actionRow: { display: "flex", gap: 12, marginTop: 18, alignItems: "center" },
  primaryBtn: {
    background: "#1f5e59",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    padding: "10px 22px",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
  },
  btnDisabled: { background: "#b7cbc8", cursor: "not-allowed" },
  secondaryBtn: {
    background: "#fff",
    color: "#1f5e59",
    border: "1.5px solid #1f5e59",
    borderRadius: 8,
    padding: "10px 18px",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
  },
  errorBox: {
    marginTop: 18,
    background: "#fdf1e8",
    border: "1px solid #e8b98a",
    color: "#8a4a13",
    borderRadius: 8,
    padding: "12px 14px",
    display: "flex",
    gap: 10,
    fontSize: 13.5,
  },
  statsRow: { display: "flex", gap: 12, marginTop: 22, flexWrap: "wrap" },
  statCard: {
    background: "#fff",
    border: "1px solid #e1ecea",
    borderRadius: 10,
    padding: "12px 18px",
    minWidth: 130,
  },
  statValue: { fontSize: 22, fontWeight: 700 },
  statLabel: { fontSize: 11.5, color: "#7c9490", marginTop: 2 },
  searchBar: {
    marginTop: 22,
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: "#fff",
    border: "1px solid #e1ecea",
    borderRadius: 8,
    padding: "8px 12px",
    maxWidth: 320,
  },
  searchInput: { border: "none", outline: "none", fontSize: 13.5, flex: 1, background: "transparent" },
  groupsWrap: { marginTop: 16, display: "flex", flexDirection: "column", gap: 10 },
  emptyMsg: { color: "#7c9490", fontSize: 13.5, padding: "20px 0" },
  groupCard: { background: "#fff", border: "1px solid #e1ecea", borderRadius: 10, overflow: "hidden" },
  groupHeader: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    gap: 4,
    background: "transparent",
    border: "none",
    padding: "12px 16px",
    cursor: "pointer",
    textAlign: "left",
  },
  groupTitle: { fontSize: 14.5, fontWeight: 700, color: "#1c2b29", marginRight: 10 },
  groupMeta: { fontSize: 12, color: "#7c9490" },
  tableWrap: { overflowX: "auto", borderTop: "1px solid #eef4f3" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: {
    textAlign: "left",
    padding: "9px 14px",
    background: "#f4f8f7",
    color: "#4a615d",
    fontWeight: 600,
    fontSize: 11.5,
    textTransform: "uppercase",
    letterSpacing: 0.3,
    borderBottom: "1px solid #e1ecea",
  },
  td: { padding: "9px 14px", borderBottom: "1px solid #f0f5f4" },
  excludedBox: { marginTop: 22, background: "#fffaf3", border: "1px solid #f0d9b5", borderRadius: 10, overflow: "hidden" },
  excludedHeader: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    background: "transparent",
    border: "none",
    padding: "12px 16px",
    cursor: "pointer",
    fontSize: 13.5,
    fontWeight: 600,
    color: "#8a4a13",
    textAlign: "left",
  },
};
