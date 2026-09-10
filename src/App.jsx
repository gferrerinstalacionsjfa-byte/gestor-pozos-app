import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Upload, Droplets, AlertTriangle, Download, ChevronDown, ChevronRight, ChevronUp, Search, Waves, CheckCircle2, RefreshCw, Minimize2, Maximize2, FileSpreadsheet, RadioTower } from "lucide-react";

const BAR_TO_MH2O = 10.19716;
const ASSOC_STORAGE_KEY = "pozos-assoc-v1";
const ASSOC_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAi-_8-SD1mogQMDKb5j2kfl0xzJub5kXE1F0YTnkVV_qiBBjYFfTOTRsE-_ylRNcTxu9vKbswww2W/pub?gid=1661818251&single=true&output=csv";
// DevEUI que no corresponen a pous reals i s'han d'ignorar sempre
const IGNORED_EUIS = new Set(["24e124847f420841"]); // mesurador de cobertura

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

async function fetchWorkbookFromUrl(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`No s'ha pogut descarregar el full de Google Sheets (${res.status}).`);
  const text = await res.text();
  return XLSX.read(text, { type: "string" });
}

function fmt(n, d = 2) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  return n.toLocaleString("es-ES", { minimumFractionDigits: d, maximumFractionDigits: d });
}

function parseAssociationFromWorkbook(assocWb) {
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

async function parseAssociationFile(file) {
  return parseAssociationFromWorkbook(await readWorkbook(file));
}

async function parseAssociationUrl(url) {
  return parseAssociationFromWorkbook(await fetchWorkbookFromUrl(url));
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

async function parseMuntatgesFromWorkbook(assocWb) {
  const sheet = findSheetWithHeader(assocWb, [/c[oó]digo\s*pozo/i, /data\s*de\s*muntatge/i]);
  if (!sheet) {
    throw new Error('No encuentro una hoja con columnas "Código pozo" y "Data de muntatge".');
  }
  const headers = sheet.headers;
  const pozoIdx = findCol(headers, [/c[oó]digo\s*pozo/i, /^codi$/i]);
  const euiIdx = findCol(headers, [/dev\s*eui/i]);
  const cotaIdx = findCol(headers, [/^cota$/i]);
  const cableIdx = findCol(headers, [/cable\s*fins\s*cota/i]);
  const installerIdx = findCol(headers, [/instal·?lador/i, /instalador/i]);
  const dateIdx = findCol(headers, [/data\s*de\s*muntatge/i]);
  const acabatIdx = findCol(headers, [/^acabat$/i]);
  const remesaIdx = findCol(headers, [/^remesa$/i]);

  const rows = [];
  for (let r = 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (!row || !row.length) continue;
    const pozo = pozoIdx !== -1 ? row[pozoIdx] : null;
    if (!pozo) continue;
    const cable = cableIdx !== -1 ? parseCableFinsCota(row[cableIdx]) : null;
    if (cable === null) continue; // solo pozos con medida de cable válida
    rows.push({
      pozo: String(pozo),
      eui: euiIdx !== -1 ? normEUI(row[euiIdx]) : "",
      cota: cotaIdx !== -1 ? parseFloat(String(row[cotaIdx] || "").replace(",", ".")) : null,
      cable,
      installer: installerIdx !== -1 ? String(row[installerIdx] || "").trim() : "",
      date: dateIdx !== -1 ? normalizeMuntatgeDateString(row[dateIdx]) : "",
      acabat: acabatIdx !== -1 ? String(row[acabatIdx] || "").trim() : "",
      remesa: remesaIdx !== -1 ? String(row[remesaIdx] || "").trim() : "",
    });
  }
  return rows;
}

async function parseLastReadingByEui(file) {
  const wb = await readWorkbook(file);
  const sheet = findSheetWithHeader(wb, [/dev\s*eui/i, /payload/i]);
  if (!sheet) {
    throw new Error('No encuentro columnas "DevEUI" y "Payload" en este archivo de lecturas.');
  }
  const euiIdx = findCol(sheet.headers, [/dev\s*eui/i]);
  const tsIdx = findCol(sheet.headers, [/marca\s*de\s*temps/i, /timestamp/i]);
  const payloadIdx = findCol(sheet.headers, [/payload/i]);

  const lastByEui = new Map(); // eui -> { lastAny, lastReal }
  for (let r = 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (!row || !row.length) continue;
    const eui = normEUI(row[euiIdx]);
    if (!eui || IGNORED_EUIS.has(eui)) continue;
    const ts = String(row[tsIdx] || "");
    const decoded = decodePayload(row[payloadIdx]);
    const isReal = !!(decoded && decoded.pressureBar !== undefined);
    const entry = lastByEui.get(eui) || { lastAny: "", lastReal: "" };
    if (ts > entry.lastAny) entry.lastAny = ts;
    if (isReal && ts > entry.lastReal) entry.lastReal = ts;
    lastByEui.set(eui, entry);
  }
  return lastByEui;
}

function excelSerialToDate(serial) {
  const base = Date.UTC(1899, 11, 30);
  return new Date(base + serial * 86400000);
}

function formatMuntatgeDate(d) {
  const pad = (n) => String(n).padStart(2, "0");
  const datePart = `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  const hasTime = d.getUTCHours() || d.getUTCMinutes() || d.getUTCSeconds();
  if (!hasTime) return datePart;
  return `${datePart} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}

function normalizeMuntatgeDateString(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) return s; // ja ve en format dd/mm/aaaa
  // número de sèrie de data del full de càlcul (dies des del 30/12/1899), a vegades amb decimals d'hora
  const num = Number(s);
  if (!Number.isNaN(num) && num > 1000 && num < 100000) {
    return formatMuntatgeDate(excelSerialToDate(num));
  }
  return s; // valor no reconegut: es deixa tal qual, no es descarta
}

function parseMuntatgeDate(s) {
  // Formats vistos: "31/07/2026" o "25/08/2026 7:51:00"
  const m = String(s || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function MuntatgesView() {
  const [rows, setRows] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [readingsFile, setReadingsFile] = useState(null);
  const [lastByEui, setLastByEui] = useState(null);
  const [readingsError, setReadingsError] = useState(null);
  const [readingsLoading, setReadingsLoading] = useState(false);
  const [groupSortDir, setGroupSortDir] = useState({}); // installer -> "desc" | "asc"
  const [collapsed, setCollapsed] = useState({});

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const wb = await fetchWorkbookFromUrl(ASSOC_CSV_URL);
      const data = await parseMuntatgesFromWorkbook(wb);
      setRows(data);
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const handleReadingsFile = useCallback(async (file) => {
    setReadingsFile(file);
    setReadingsLoading(true);
    setReadingsError(null);
    setLastByEui(null);
    try {
      const map = await parseLastReadingByEui(file);
      setLastByEui(map);
    } catch (e) {
      setReadingsError(e.message || String(e));
    } finally {
      setReadingsLoading(false);
    }
  }, []);

  const sortWells = (wells, dir) =>
    [...wells].sort((a, b) => {
      const da = parseMuntatgeDate(a.date);
      const db = parseMuntatgeDate(b.date);
      if (da && db) return dir === "desc" ? db - da : da - db;
      if (da) return -1;
      if (db) return 1;
      return a.pozo.localeCompare(b.pozo);
    });

  const toggleGroupSort = (installer) =>
    setGroupSortDir((d) => ({ ...d, [installer]: (d[installer] || "desc") === "desc" ? "asc" : "desc" }));

  const groups = useMemo(() => {
    if (!rows) return [];
    const q = query.trim().toLowerCase();
    const list = q ? rows.filter((r) => r.pozo.toLowerCase().includes(q)) : rows;
    const byInstaller = new Map();
    for (const r of list) {
      const key = r.installer || "Sense instal·lador assignat";
      if (!byInstaller.has(key)) byInstaller.set(key, []);
      byInstaller.get(key).push(r);
    }
    return Array.from(byInstaller.entries())
      .map(([installer, wells]) => ({ installer, wells }))
      .sort((a, b) => a.installer.localeCompare(b.installer));
  }, [rows, query]);

  const pendents = useMemo(() => (rows || []).filter((r) => !parseMuntatgeDate(r.date)), [rows]);
  const toggle = (installer) => setCollapsed((c) => ({ ...c, [installer]: !c[installer] }));

  const exportPdf = () => {
    const doc = new jsPDF();
    doc.setFontSize(14);
    doc.text("Muntatges — Instal·lacions JFA", 14, 15);
    doc.setFontSize(9);
    doc.setTextColor(120);
    doc.text(`Generat el ${new Date().toLocaleString("es-ES")}`, 14, 21);

    const head = [["Pou", "Remesa", "Instal·lador", "Data de muntatge", ...(lastByEui ? ["Última lectura real"] : [])]];
    const body = [];
    groups.forEach((g) => {
      const dir = groupSortDir[g.installer] || "desc";
      sortWells(g.wells, dir).forEach((r) => {
        const activity = lastByEui && r.eui ? lastByEui.get(r.eui) : null;
        const row = [r.pozo, r.remesa || "—", g.installer, r.date || "pendent"];
        if (lastByEui) {
          row.push(!r.eui ? "—" : activity && activity.lastReal ? activity.lastReal.slice(0, 16).replace("T", " ") : "sense lectures reals");
        }
        body.push(row);
      });
    });

    autoTable(doc, { head, body, startY: 26, styles: { fontSize: 8 }, headStyles: { fillColor: [31, 94, 89] } });
    doc.save(`muntatges_${new Date().toISOString().slice(0, 10)}.pdf`);
  };

  return (
    <div style={styles.page}>
      <div style={styles.headerBar}>
        <div style={styles.brand}>
          <Waves size={22} color="#e8f3f2" />
          <span style={styles.brandText}>Muntatges</span>
        </div>
        <span style={styles.brandSub}>Data de muntatge i instal·lador, pous amb cable fins cota mesurat</span>
      </div>

      <div style={styles.container}>
        <UploadCard
          label="Fitxer de lectures (opcional)"
          hint="Puja l'export de lectures per veure l'última comunicació real de cada pou — arrossega'l aquí o fes clic"
          file={readingsFile}
          onFile={handleReadingsFile}
          inputId="muntatges-readings-input"
          extra={
            readingsError ? (
              <span style={{ color: "#a86a2d" }}>{readingsError}</span>
            ) : readingsLoading ? (
              <span>Analitzant…</span>
            ) : lastByEui ? (
              <span style={{ color: "#2a8f6c", fontWeight: 600 }}>{lastByEui.size} DevEUI amb activitat al fitxer</span>
            ) : null
          }
        />

        <div style={styles.searchRow}>
          <div style={styles.searchBar}>
            <Search size={16} color="#7c9490" />
            <input
              style={styles.searchInput}
              placeholder="Filtrar per codi de pou…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <button style={styles.secondaryBtn} onClick={exportPdf} disabled={!rows || !rows.length}>
            <Download size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
            Descarregar PDF
          </button>
          <button style={styles.secondaryBtn} onClick={load} disabled={loading}>
            <RefreshCw size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
            {loading ? "Actualitzant…" : "Actualitzar ara"}
          </button>
        </div>

        {error && (
          <div style={styles.errorBox}>
            <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>{error}</span>
          </div>
        )}

        {rows && (
          <>
            <div style={styles.statsRow}>
              <Stat label="Pous amb cable vàlid" value={rows.length} />
              <Stat label="Pendents de muntar" value={pendents.length} warn />
              <Stat label="Instal·ladors" value={groups.length} />
            </div>

            <div style={styles.groupsWrap}>
              {groups.map((g) => (
                <div key={g.installer} style={styles.groupCard}>
                  <button style={styles.groupHeader} onClick={() => toggle(g.installer)}>
                    {collapsed[g.installer] ? <ChevronRight size={16} /> : <ChevronDown size={16} />}
                    <CheckCircle2 size={16} color="#3e8e86" style={{ marginRight: 6 }} />
                    <span style={styles.groupTitle}>{g.installer}</span>
                    <span style={styles.groupMeta}>{g.wells.length} pous</span>
                  </button>
                  {!collapsed[g.installer] && (
                    <div style={styles.tableWrap}>
                      <table style={styles.table}>
                        <thead>
                          <tr>
                            <th style={styles.th}>Pou</th>
                            <th style={styles.th}>Remesa</th>
                            <th
                              style={{ ...styles.th, cursor: "pointer", userSelect: "none" }}
                              onClick={() => toggleGroupSort(g.installer)}
                            >
                              Data de muntatge
                              {(groupSortDir[g.installer] || "desc") === "desc" ? (
                                <ChevronDown size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                              ) : (
                                <ChevronUp size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                              )}
                            </th>
                            {lastByEui && <th style={styles.th}>Última lectura real</th>}
                          </tr>
                        </thead>
                        <tbody>
                          {sortWells(g.wells, groupSortDir[g.installer] || "desc").map((r, i) => {
                            const activity = lastByEui && r.eui ? lastByEui.get(r.eui) : null;
                            return (
                              <tr key={i}>
                                <td style={{ ...styles.td, fontWeight: 600 }}>{r.pozo}</td>
                                <td style={styles.td}>{r.remesa || "—"}</td>
                                <td style={styles.td}>
                                  {r.date || <span style={{ color: "#a86a2d" }}>pendent</span>}
                                </td>
                                {lastByEui && (
                                  <td style={styles.td}>
                                    {!r.eui ? (
                                      "—"
                                    ) : activity && activity.lastReal ? (
                                      <span style={{ color: "#2a8f6c" }}>{activity.lastReal.slice(0, 16).replace("T", " ")}</span>
                                    ) : (
                                      <span style={{ color: "#b03a3a" }}>sense lectures reals</span>
                                    )}
                                  </td>
                                )}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </>
        )}
      </div>

      <footer style={styles.footer}>
        © {new Date().getFullYear()} Instal·lacions JFA. Tots els drets reservats.
      </footer>
    </div>
  );
}

export default function App() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("vista") === "muntatges") return <MuntatgesView />;
  return <NivellApp />;
}

function NivellApp() {
  const [assocData, setAssocData] = useState(null); // { map, excluded, fileName, savedAt, source }
  const [assocLoading, setAssocLoading] = useState(true);
  const [assocError, setAssocError] = useState(null);
  const [assocPersisted, setAssocPersisted] = useState(true);
  const [readingsFile, setReadingsFile] = useState(null);
  const [readingsPreview, setReadingsPreview] = useState(null); // { rows, uniqueWells }
  const [readingsPreviewError, setReadingsPreviewError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [results, setResults] = useState(null); // { groups, excluded, stats }
  const [query, setQuery] = useState("");
  const [collapsed, setCollapsed] = useState({});
  const [showExcluded, setShowExcluded] = useState(false);
  const [showNoSignal, setShowNoSignal] = useState(false);
  const noSignalRef = useRef(null);

  const jumpToNoSignal = () => {
    setShowNoSignal(true);
    setTimeout(() => noSignalRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  };

  const syncFromSheet = useCallback(async ({ silent } = {}) => {
    if (!silent) setAssocLoading(true);
    setAssocError(null);
    try {
      const { map, excluded } = await parseAssociationUrl(ASSOC_CSV_URL);
      const data = { map, excluded, fileName: "Google Sheets (automàtic)", savedAt: new Date().toISOString(), source: "url" };
      setAssocPersisted(saveAssocToStorage(data));
      setAssocData(data);
      setResults(null);
    } catch (e) {
      // Si falla la conexión (p. ex. sense internet), usem la darrera còpia guardada.
      const saved = loadAssocFromStorage();
      if (saved) {
        setAssocData(saved);
        setAssocError(`No s'ha pogut connectar amb Google Sheets ara mateix — mostrant l'última còpia guardada (${new Date(saved.savedAt).toLocaleString("es-ES")}).`);
      } else {
        setAssocError(e.message || String(e));
      }
    } finally {
      setAssocLoading(false);
    }
  }, []);

  useEffect(() => {
    syncFromSheet();
  }, [syncFromSheet]);

  const handleAssocFile = useCallback(async (file) => {
    setAssocLoading(true);
    setAssocError(null);
    try {
      const { map, excluded } = await parseAssociationFile(file);
      const data = { map, excluded, fileName: file.name, savedAt: new Date().toISOString(), source: "file" };
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

  const handleReadingsFile = useCallback(async (file) => {
    setReadingsFile(file);
    setResults(null);
    setReadingsPreview(null);
    setReadingsPreviewError(null);
    try {
      const wb = await readWorkbook(file);
      const sheet = findSheetWithHeader(wb, [/dev\s*eui/i, /payload/i]);
      if (!sheet) {
        setReadingsPreviewError('No trobo columnes "DevEUI" i "Payload" en aquest fitxer.');
        return;
      }
      const euiIdx = findCol(sheet.headers, [/dev\s*eui/i]);
      const wells = new Set();
      let rows = 0;
      for (let r = 1; r < sheet.rows.length; r++) {
        const row = sheet.rows[r];
        if (!row || !row.length) continue;
        rows++;
        const eui = normEUI(row[euiIdx]);
        if (eui && !IGNORED_EUIS.has(eui)) wells.add(eui);
      }
      setReadingsPreview({ rows, uniqueWells: wells.size });
    } catch (e) {
      setReadingsPreviewError(e.message || String(e));
    }
  }, []);


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
      const euisInFile = new Set();
      let totalReadings = 0;
      let matchedReadings = 0;
      let decodedReadings = 0;

      for (let r = 1; r < readSheet.rows.length; r++) {
        const row = readSheet.rows[r];
        if (!row || !row.length) continue;
        const eui = normEUI(row[rEuiIdx]);
        if (!eui || IGNORED_EUIS.has(eui)) continue;
        totalReadings++;
        euisInFile.add(eui);
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

      const noSignal = [];
      for (const [eui, info] of assocMap.entries()) {
        if (!euisInFile.has(eui)) {
          noSignal.push({ pozo: info.pozo, eui, cota: info.cota, cable: info.cable });
        }
      }
      noSignal.sort((a, b) => a.pozo.localeCompare(b.pozo));

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
        noSignal,
        stats: {
          totalReadings,
          matchedReadings,
          decodedReadings,
          wellsOk: assocMap.size,
          wellsExcluded: assocExcluded.length,
          wellsNoSignal: noSignal.length,
        },
      });
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  }, []);

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

    if (results.noSignal.length) {
      const wsNoSignal = XLSX.utils.json_to_sheet(
        results.noSignal.map((x) => ({ Pozo: x.pozo, DevEUI: x.eui, "Cota (m)": x.cota, "Cable fins cota (m)": x.cable }))
      );
      XLSX.utils.book_append_sheet(wb, wsNoSignal, "Sin señal");
    }
    if (results.excluded.length) {
      const wsExc = XLSX.utils.json_to_sheet(
        results.excluded.map((x) => ({ Pozo: x.pozo, DevEUI: x.eui, Motivo: x.motivo }))
      );
      XLSX.utils.book_append_sheet(wb, wsExc, "Pozos excluidos");
    }
    XLSX.writeFile(wb, `nivel_freatico_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const toggle = (pozo) => setCollapsed((c) => ({ ...c, [pozo]: !c[pozo] }));

  const allExpandedVisible = filteredGroups.length > 0 && filteredGroups.every((g) => !collapsed[g.pozo]);
  const toggleAll = () => {
    setCollapsed((c) => {
      const next = { ...c };
      filteredGroups.forEach((g) => {
        next[g.pozo] = allExpandedVisible; // si tot estava desplegat, ara es plega tot
      });
      return next;
    });
  };

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
                <div style={styles.uploadLabel}>
                  {assocData.source === "url" ? "Pous sincronitzats amb Google Sheets" : "Fitxer d'associació carregat manualment"}
                </div>
                <div style={styles.uploadHint}>
                  {assocData.fileName} · actualitzat el {new Date(assocData.savedAt).toLocaleString("es-ES")}
                  {!assocPersisted && " · no s'ha pogut desar còpia local per quan no hi hagi connexió"}
                </div>
                <div style={styles.uploadExtra}>
                  <strong>{assocData.map.size}</strong> pous amb dades vàlides
                  {assocData.excluded.length > 0 && ` · ${assocData.excluded.length} exclosos`}
                </div>
              </div>
              <button style={styles.smallLinkBtn} onClick={() => syncFromSheet()} disabled={assocLoading}>
                <RefreshCw size={13} style={{ marginRight: 4, verticalAlign: "-2px" }} />
                {assocLoading ? "Actualitzant…" : "Actualitzar ara"}
              </button>
            </div>
          ) : (
            <div style={styles.assocLoadedCard}>
              <RefreshCw size={20} color="#7c9490" />
              <div style={{ flex: 1 }}>
                <div style={styles.uploadLabel}>{assocLoading ? "Connectant amb Google Sheets…" : "Sense dades d'associació"}</div>
                <div style={styles.uploadHint}>
                  {assocLoading ? "Descarregant el full de pous" : "No s'ha pogut carregar cap font de dades"}
                </div>
              </div>
            </div>
          )}
          <UploadCard
            label="Fitxer de lectures"
            hint="Excel export amb columnes DevEUI, Marca de temps, Payload (HEX) — arrossega'l aquí o fes clic"
            file={readingsFile}
            onFile={handleReadingsFile}
            inputId="readings-input"
            extra={
              readingsPreviewError ? (
                <span style={{ color: "#a86a2d" }}>{readingsPreviewError}</span>
              ) : readingsPreview ? (
                <span style={{ color: "#2a8f6c", fontWeight: 600 }}>
                  {readingsPreview.rows} lectures · {readingsPreview.uniqueWells} pous diferents al fitxer
                </span>
              ) : null
            }
          />
        </div>

        <div style={styles.manualUploadRow}>
          <label htmlFor="assoc-input" style={styles.tinyDropZone}>
            <FileSpreadsheet size={13} style={{ marginRight: 4, verticalAlign: "-2px" }} />
            Pujar un fitxer d'associació manualment (en lloc de Google Sheets)
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
              <Stat label="Pous instal·lats sense senyal" value={results.stats.wellsNoSignal} warn onClick={results.stats.wellsNoSignal > 0 ? jumpToNoSignal : undefined} />
              <Stat label="Lectures totals" value={results.stats.totalReadings} />
              <Stat label="Lectures amb pou associat" value={results.stats.matchedReadings} />
              <Stat label="Lectures decodificades" value={results.stats.decodedReadings} />
            </div>

            <div style={styles.searchRow}>
              <div style={styles.searchBar}>
                <Search size={16} color="#7c9490" />
                <input
                  style={styles.searchInput}
                  placeholder="Filtrar per codi de pou…"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                />
              </div>
              {filteredGroups.length > 0 && (
                <button style={styles.secondaryBtn} onClick={toggleAll}>
                  {allExpandedVisible ? (
                    <>
                      <Minimize2 size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
                      Plegar tot
                    </>
                  ) : (
                    <>
                      <Maximize2 size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
                      Desplegar tot
                    </>
                  )}
                </button>
              )}
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

            {results.noSignal.length > 0 && (
              <div style={styles.noSignalBox} ref={noSignalRef}>
                <button style={styles.noSignalHeader} onClick={() => setShowNoSignal((s) => !s)}>
                  {showNoSignal ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  <RadioTower size={15} color="#b03a3a" style={{ margin: "0 6px" }} />
                  {results.noSignal.length} pous instal·lats (amb cota vàlida) que no han emès cap lectura en aquest fitxer
                </button>
                {showNoSignal && (
                  <div style={styles.tableWrap}>
                    <table style={styles.table}>
                      <thead>
                        <tr>
                          <th style={styles.th}>Pou</th>
                          <th style={styles.th}>DevEUI</th>
                          <th style={styles.th}>Cota</th>
                          <th style={styles.th}>Cable fins cota</th>
                        </tr>
                      </thead>
                      <tbody>
                        {results.noSignal.map((x, i) => (
                          <tr key={i}>
                            <td style={styles.td}>{x.pozo}</td>
                            <td style={{ ...styles.td, fontFamily: "monospace", fontSize: 12 }}>{x.eui}</td>
                            <td style={styles.td}>{fmt(x.cota, 1)} m</td>
                            <td style={styles.td}>{fmt(x.cable, 2)} m</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

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

      <footer style={styles.footer}>
        © {new Date().getFullYear()} Instal·lacions JFA. Tots els drets reservats.
      </footer>
    </div>
  );
}

function UploadCard({ label, hint, file, onFile, inputId, extra }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <label
      htmlFor={inputId}
      style={{
        ...styles.uploadCard,
        ...(file ? styles.uploadCardFilled : {}),
        ...(dragOver ? styles.uploadCardDragOver : {}),
      }}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const f = e.dataTransfer.files?.[0];
        if (f) onFile(f);
      }}
    >
      <Upload size={20} color={file ? "#3e8e86" : "#7c9490"} />
      <div style={{ flex: 1 }}>
        <div style={styles.uploadLabel}>{label}</div>
        <div style={styles.uploadHint}>{file ? file.name : hint}</div>
        {extra && <div style={styles.uploadExtra}>{extra}</div>}
      </div>
      <input
        id={inputId}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onFile(f);
        }}
      />
    </label>
  );
}

function Stat({ label, value, warn, onClick }) {
  const clickable = typeof onClick === "function";
  return (
    <div
      style={{ ...styles.statCard, ...(clickable ? styles.statCardClickable : {}) }}
      onClick={onClick}
      role={clickable ? "button" : undefined}
      tabIndex={clickable ? 0 : undefined}
    >
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
    display: "flex",
    flexDirection: "column",
  },
  footer: {
    marginTop: "auto",
    textAlign: "center",
    padding: "18px 20px",
    fontSize: 12,
    color: "#7c9490",
    borderTop: "1px solid #e1ecea",
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
  uploadCardDragOver: { borderColor: "#1f5e59", background: "#eef8f6", borderStyle: "solid" },
  uploadExtra: { fontSize: 12, color: "#2a8f6c", marginTop: 4 },
  tinyDropZone: {
    display: "inline-flex",
    alignItems: "center",
    fontSize: 12,
    color: "#4a615d",
    cursor: "pointer",
    padding: "6px 10px",
    border: "1px dashed #cfe0dd",
    borderRadius: 6,
    background: "#fff",
  },
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
  manualUploadRow: { marginTop: 4 },
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
  statCardClickable: { cursor: "pointer", transition: "box-shadow .15s, border-color .15s" },
  statValue: { fontSize: 22, fontWeight: 700 },
  statLabel: { fontSize: 11.5, color: "#7c9490", marginTop: 2 },
  searchBar: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: "#fff",
    border: "1px solid #e1ecea",
    borderRadius: 8,
    padding: "8px 12px",
    maxWidth: 320,
    flex: "1 1 260px",
  },
  searchRow: { display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginTop: 22 },
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
  noSignalBox: { marginTop: 22, background: "#fdf2f2", border: "1px solid #f0c4c4", borderRadius: 10, overflow: "hidden" },
  noSignalHeader: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    background: "transparent",
    border: "none",
    padding: "12px 16px",
    cursor: "pointer",
    fontSize: 13.5,
    fontWeight: 600,
    color: "#b03a3a",
    textAlign: "left",
  },
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
