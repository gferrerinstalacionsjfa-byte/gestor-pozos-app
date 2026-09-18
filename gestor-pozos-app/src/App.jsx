import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import { Upload, Droplets, AlertTriangle, Download, ChevronDown, ChevronRight, ChevronUp, Search, Waves, CheckCircle2, RefreshCw, Minimize2, Maximize2, FileSpreadsheet, RadioTower, Map as MapIcon } from "lucide-react";

const BAR_TO_MH2O = 10.19716;
const ASSOC_STORAGE_KEY = "pozos-assoc-v1";
const ASSOC_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAi-_8-SD1mogQMDKb5j2kfl0xzJub5kXE1F0YTnkVV_qiBBjYFfTOTRsE-_ylRNcTxu9vKbswww2W/pub?gid=1661818251&single=true&output=csv";
const COORDS_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQAi-_8-SD1mogQMDKb5j2kfl0xzJub5kXE1F0YTnkVV_qiBBjYFfTOTRsE-_ylRNcTxu9vKbswww2W/pub?gid=392569350&single=true&output=csv";
// DevEUI que no corresponen a pous reals i s'han d'ignorar sempre
const IGNORED_EUIS = new Set(["24e124847f420841"]); // mesurador de cobertura

function fixMangledCoord(raw, intDigits) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  if (!s) return null;
  const dotCount = (s.match(/\./g) || []).length;
  const commaCount = (s.match(/,/g) || []).length;
  if (dotCount <= 1 && commaCount === 0) {
    const v = parseFloat(s);
    return Number.isNaN(v) ? null : v;
  }
  if (commaCount === 1 && dotCount === 0) {
    const v = parseFloat(s.replace(",", "."));
    return Number.isNaN(v) ? null : v;
  }
  // Google Sheets ha "maquillat" el número amb punts de milers (ex: "3.956.469.418")
  const digits = s.replace(/[^\d]/g, "");
  if (digits.length <= intDigits) return null;
  const fixed = `${digits.slice(0, intDigits)}.${digits.slice(intDigits)}`;
  const v = parseFloat(fixed);
  return Number.isNaN(v) ? null : v;
}

async function parseCoordsFromWorkbook(wb) {
  const sheet = findSheetWithHeader(wb, [/c[oó]digo\s*pozo/i, /latitud/i]);
  if (!sheet) throw new Error('No encuentro columnas "Codigo pozo", "Latitud" i "Longitud" al full de coordenades.');
  const pozoIdx = findCol(sheet.headers, [/c[oó]digo\s*pozo/i]);
  const latIdx = findCol(sheet.headers, [/latitud/i]);
  const lonIdx = findCol(sheet.headers, [/longitud/i]);
  const coords = new Map();
  for (let r = 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (!row || !row.length) continue;
    const pozo = pozoIdx !== -1 ? String(row[pozoIdx] || "").trim() : "";
    if (!pozo) continue;
    const lat = fixMangledCoord(row[latIdx], 2);
    const lon = fixMangledCoord(row[lonIdx], 1);
    if (lat === null || lon === null) continue;
    coords.set(pozo, { lat, lon });
  }
  return coords;
}

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
  const clean = hex.replace(/\s/g, "").toUpperCase();
  if (!/^[0-9A-F]+$/.test(clean) || clean.length < 4) return null;
  const bytes = [];
  for (let i = 0; i < clean.length; i += 2) bytes.push(parseInt(clean.substr(i, 2), 16));
  const len = bytes.length;

  const readFloat32LE = (offset) => {
    if (offset + 4 > len) return null;
    const buf = new ArrayBuffer(4);
    const view = new DataView(buf);
    for (let k = 0; k < 4; k++) view.setUint8(k, bytes[offset + k]);
    const val = view.getFloat32(0, true);
    return Number.isFinite(val) ? val : null;
  };
  const readUint32LE = (offset) => {
    if (offset + 4 > len) return null;
    return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
  };
  const readUint16LE = (offset) => {
    if (offset + 2 > len) return null;
    return bytes[offset] | (bytes[offset + 1] << 8);
  };

  const out = { errors: [], deviceInfo: null, ack: null, historical: [] };

  // --- Comandos / ACK (comença per FE) ---
  if (bytes[0] === 0xfe) {
    const code = bytes[1];
    if (code === 0x02) out.ack = `ACK collecting interval: ${readUint16LE(2)} s`;
    else if (code === 0x03) out.ack = `ACK reporting interval: ${readUint16LE(2)} s`;
    else if (code === 0x68) out.ack = `ACK Data Storage: ${bytes[2] === 0x01 ? "activat" : "desactivat"}`;
    else if (code === 0x69) out.ack = `ACK Data Retransmission: ${bytes[2] === 0x01 ? "activada" : "desactivada"}`;
    else if (code === 0x6a) out.ack = "ACK comando FF6A";
    else if (code === 0x10) out.ack = "ACK reinici UC502";
    else if (code === 0x28) out.ack = "ACK petició de dada actual";
    else out.ack = `Resposta/ACK UC502: ${clean}`;
  }

  // --- GPIO ---
  if (bytes[0] === 0x03 && bytes[1] === 0x00) out.gpio1 = bytes[2];
  if (bytes[3] === 0x04 && bytes[4] === 0x00) out.gpio2 = bytes[5];

  // --- Bateria (últims 3 bytes: 01 75 XX) ---
  if (len >= 3 && bytes[len - 3] === 0x01 && bytes[len - 2] === 0x75) {
    out.battery = bytes[len - 1];
  }

  // --- Errors RS485/Modbus i lectures en viu (FF0E = dada, FF15 = error) ---
  for (let i = 0; i + 3 <= len; i++) {
    if (bytes[i] === 0xff && (bytes[i + 1] === 0x0e || bytes[i + 1] === 0x15)) {
      const channel = bytes[i + 2];
      if (bytes[i + 1] === 0x15) {
        if (channel === 7) out.errors.push("presió");
        else if (channel === 8) out.errors.push("temperatura");
        else if (channel === 9) out.errors.push("conductivitat");
        i += 2;
        continue;
      }
      if (bytes[i + 3] === 0x25) {
        const val = readFloat32LE(i + 4);
        if (val !== null) {
          if (channel === 7) out.pressureBar = val;
          else if (channel === 8) out.tempC = val;
          else if (channel === 9) out.condMScm = val;
        }
        i += 7;
      }
    }
  }

  // --- Informació del dispositiu ---
  for (let i = 0; i + 2 <= len; i++) {
    if (bytes[i] !== 0xff) continue;
    const code = bytes[i + 1];
    if (code === 0x0b) {
      out.deviceInfo = out.deviceInfo || {};
      out.deviceInfo.powerOn = true;
    } else if (code === 0x16 && i + 10 <= len) {
      out.deviceInfo = out.deviceInfo || {};
      out.deviceInfo.serial = bytes
        .slice(i + 2, i + 10)
        .map((b) => b.toString(16).padStart(2, "0").toUpperCase())
        .join("");
    } else if (code === 0x09 && i + 4 <= len) {
      out.deviceInfo = out.deviceInfo || {};
      out.deviceInfo.firmware = `${bytes[i + 2]}.${bytes[i + 3]}`;
    } else if (code === 0x0a && i + 4 <= len) {
      out.deviceInfo = out.deviceInfo || {};
      out.deviceInfo.hardware = `${String(bytes[i + 2]).padStart(2, "0")}.${String(bytes[i + 3]).padStart(2, "0")}`;
    } else if (code === 0x0f && i + 3 <= len) {
      out.deviceInfo = out.deviceInfo || {};
      out.deviceInfo.classType = bytes[i + 2] === 0x00 ? "Class A" : "Desconeguda";
    }
  }

  // --- Històrics 20DC / 20DD ---
  for (let i = 0; i + 2 <= len; i++) {
    if (bytes[i] === 0x20 && bytes[i + 1] === 0xdc) {
      const ts = readUint32LE(i + 2);
      if (ts !== null) out.historical.push({ timestamp: new Date(ts * 1000), type: "20DC" });
    } else if (bytes[i] === 0x20 && bytes[i + 1] === 0xdd) {
      const ts = readUint32LE(i + 2);
      const mask = readUint16LE(i + 6);
      if (ts === null || mask === null) continue;
      const rec = { timestamp: new Date(ts * 1000), type: "20DD" };
      let offset = i + 8;
      if (mask & 1) {
        if (bytes[offset] === 0x25) {
          const v = readFloat32LE(offset + 1);
          if (v !== null) rec.pressureBar = v;
        }
        offset += 5;
      }
      if (mask & 2) {
        if (bytes[offset] === 0x25) {
          const v = readFloat32LE(offset + 1);
          if (v !== null) rec.tempC = v;
        }
        offset += 5;
      }
      if (mask & 4) {
        if (bytes[offset] === 0x25) {
          const v = readFloat32LE(offset + 1);
          if (v !== null) rec.condMScm = v;
        }
        offset += 5;
      }
      out.historical.push(rec);
    }
  }

  const hasAnything =
    out.pressureBar !== undefined ||
    out.tempC !== undefined ||
    out.condMScm !== undefined ||
    out.battery !== undefined ||
    out.gpio1 !== undefined ||
    out.gpio2 !== undefined ||
    out.errors.length ||
    out.deviceInfo ||
    out.ack ||
    out.historical.length;

  return hasAnything ? out : null;
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
  const notMounted = [];
  for (let r = 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (!row || !row.length) continue;
    const pozo = pozoIdx !== -1 ? row[pozoIdx] : null;
    if (!pozo) continue;
    const cable = cableIdx !== -1 ? parseCableFinsCota(row[cableIdx]) : null;
    const eui = euiIdx !== -1 ? normEUI(row[euiIdx]) : "";
    const cota = cotaIdx !== -1 ? parseFloat(String(row[cotaIdx] || "").replace(",", ".")) : null;
    const installer = installerIdx !== -1 ? String(row[installerIdx] || "").trim() : "";
    const date = dateIdx !== -1 ? normalizeMuntatgeDateString(row[dateIdx]) : "";
    const acabat = acabatIdx !== -1 ? String(row[acabatIdx] || "").trim() : "";
    const remesa = remesaIdx !== -1 ? String(row[remesaIdx] || "").trim() : "";
    if (cable === null) {
      // sense "Cable fins cota" vàlid -> no consta com a muntat, però es conserva per poder
      // comprovar si comunica igualment quan es puja el fitxer de lectures
      notMounted.push({ pozo: String(pozo), eui, cota, installer, date, acabat, remesa });
      continue;
    }
    rows.push({ pozo: String(pozo), eui, cota, cable, installer, date, acabat, remesa });
  }
  return { rows, notMounted };
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

  const lastByEui = new Map(); // eui -> { lastAny, lastReal, battery, errors, gpio1, gpio2, deviceInfo }
  const historyByEui = new Map(); // eui -> [{ ts, status, errors }]
  const deviceInfoSummary = (d) => {
    if (!d) return null;
    const parts = [];
    if (d.serial) parts.push(`Sèrie ${d.serial}`);
    if (d.firmware) parts.push(`FW ${d.firmware}`);
    if (d.hardware) parts.push(`HW ${d.hardware}`);
    if (d.classType) parts.push(d.classType);
    if (d.powerOn) parts.push("power_on");
    return parts.join(" · ") || null;
  };
  for (let r = 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    if (!row || !row.length) continue;
    const eui = normEUI(row[euiIdx]);
    if (!eui || IGNORED_EUIS.has(eui)) continue;
    const ts = String(row[tsIdx] || "");
    const decoded = decodePayload(row[payloadIdx]);
    const entry = lastByEui.get(eui) || {
      lastAny: "",
      lastReal: "",
      battery: null,
      errors: new Set(),
      gpio1: undefined,
      gpio2: undefined,
      deviceInfoRaw: null,
    };
    if (ts > entry.lastAny) entry.lastAny = ts;

    // --- registre per a l'historial detallat (una entrada per lectura) ---
    const hasOk =
      !!decoded &&
      (decoded.pressureBar !== undefined ||
        decoded.tempC !== undefined ||
        decoded.condMScm !== undefined ||
        (decoded.historical && decoded.historical.some((h) => h.pressureBar !== undefined || h.tempC !== undefined || h.condMScm !== undefined)));
    const hasErr = !!decoded && decoded.errors && decoded.errors.length > 0;
    let status = "sense_dades";
    if (hasOk && !hasErr) status = "ok";
    else if (hasOk && hasErr) status = "parcial";
    else if (!hasOk && hasErr) status = "error";
    else if (decoded) status = "info"; // bateria/GPIO/info dispositiu/ACK sense mesura ni error
    if (!historyByEui.has(eui)) historyByEui.set(eui, []);
    historyByEui.get(eui).push({ ts, status, errors: decoded ? decoded.errors : [] });

    if (decoded) {
      const isReal = decoded.pressureBar !== undefined || decoded.tempC !== undefined || decoded.condMScm !== undefined;
      if (isReal && ts > entry.lastReal) entry.lastReal = ts;
      if (decoded.historical) {
        for (const h of decoded.historical) {
          if (h.pressureBar === undefined && h.tempC === undefined && h.condMScm === undefined) continue;
          const hIso = h.timestamp.toISOString();
          if (hIso > entry.lastReal) entry.lastReal = hIso;
        }
      }
      if (decoded.battery !== undefined) entry.battery = decoded.battery;
      decoded.errors.forEach((e) => entry.errors.add(e));
      if (decoded.gpio1 !== undefined) entry.gpio1 = decoded.gpio1;
      if (decoded.gpio2 !== undefined) entry.gpio2 = decoded.gpio2;
      if (decoded.deviceInfo) entry.deviceInfoRaw = { ...entry.deviceInfoRaw, ...decoded.deviceInfo };
    }
    lastByEui.set(eui, entry);
  }
  for (const entry of lastByEui.values()) {
    entry.errors = Array.from(entry.errors);
    entry.deviceInfo = deviceInfoSummary(entry.deviceInfoRaw);
    delete entry.deviceInfoRaw;
  }
  for (const hist of historyByEui.values()) {
    hist.sort((a, b) => (a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0));
  }
  return { lastByEui, historyByEui };
}

const READING_STATUS_COLORS = {
  ok: "#2a8f6c",
  parcial: "#c98a2d",
  error: "#b03a3a",
  info: "#c7d3d1",
  sense_dades: "#e6ecea",
};
const READING_STATUS_LABELS = {
  ok: "lectura correcta",
  parcial: "parcial (algun paràmetre amb error)",
  error: "error de lectura",
  info: "missatge sense mesura (bateria/estat/info)",
  sense_dades: "sense dades reconegudes",
};

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
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(.*)$/);
  if (m) {
    let day = Number(m[1]);
    let month = Number(m[2]);
    const rest = m[4] || "";
    // si el "mes" no pot ser-ho (>12) però el "dia" sí, és que ve escrit mes/dia -> es capgira
    if (month > 12 && day <= 12) {
      [day, month] = [month, day];
    }
    // si els dos números són <=12 és ambigu: es manté l'ordre tal com ve (dia/mes, conveni majoritari al full)
    const pad = (n) => String(n).padStart(2, "0");
    return `${pad(day)}/${pad(month)}/${m[3]}${rest}`;
  }
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

function getIlla(pozoCode) {
  const p = String(pozoCode || "").toUpperCase();
  if (p.startsWith("MA")) return "Mallorca";
  if (p.startsWith("ME")) return "Menorca";
  if (p.startsWith("EI")) return "Eivissa";
  if (p.startsWith("FO")) return "Formentera";
  return "Altres";
}

function MuntatgesView() {
  const [rows, setRows] = useState(null);
  const [notMounted, setNotMounted] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [readingsFile, setReadingsFile] = useState(null);
  const [lastByEui, setLastByEui] = useState(null);
  const [historyByEui, setHistoryByEui] = useState(null);
  const [expandedHistory, setExpandedHistory] = useState({});
  const [readingsError, setReadingsError] = useState(null);
  const [readingsLoading, setReadingsLoading] = useState(false);
  const [groupBy, setGroupBy] = useState("installer"); // "installer" | "illa" | "none"
  const [sortField, setSortField] = useState(null); // null | "muntatge" | "lectura"
  const [sortDir, setSortDir] = useState("desc"); // "desc" | "asc"
  const [collapsed, setCollapsed] = useState({});
  const [showNotMounted, setShowNotMounted] = useState(false);
  const [showMap, setShowMap] = useState(false);
  const [coordsByPozo, setCoordsByPozo] = useState(null);
  const [coordsError, setCoordsError] = useState(null);
  const [visibleCols, setVisibleCols] = useState({ battery: true, errors: true, gpio: true, deviceInfo: true });
  const toggleCol = (key) => setVisibleCols((c) => ({ ...c, [key]: !c[key] }));
  const toggleHistory = (eui) => setExpandedHistory((c) => ({ ...c, [eui]: !c[eui] }));

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const wb = await fetchWorkbookFromUrl(ASSOC_CSV_URL);
      const { rows: data, notMounted: nm } = await parseMuntatgesFromWorkbook(wb);
      setRows(data);
      setNotMounted(nm);
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchWorkbookFromUrl(COORDS_CSV_URL)
      .then(parseCoordsFromWorkbook)
      .then(setCoordsByPozo)
      .catch((e) => setCoordsError(e.message || String(e)));
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const handleReadingsFile = useCallback(async (file) => {
    setReadingsFile(file);
    setReadingsLoading(true);
    setReadingsError(null);
    setLastByEui(null);
    setHistoryByEui(null);
    setExpandedHistory({});
    try {
      const { lastByEui: lm, historyByEui: hm } = await parseLastReadingByEui(file);
      setLastByEui(lm);
      setHistoryByEui(hm);
    } catch (e) {
      setReadingsError(e.message || String(e));
    } finally {
      setReadingsLoading(false);
    }
  }, []);

  // si es queda sense fitxer de lectures mentre s'ordenava per última lectura, es treu l'ordre
  useEffect(() => {
    if (sortField === "lectura" && !lastByEui) setSortField(null);
  }, [lastByEui, sortField]);

  // clic a una capçalera ordenable: 1r clic -> desc (▾), 2n clic -> asc (▴), 3r clic -> treu l'ordre
  const toggleSort = (field) => {
    setSortField((current) => {
      if (current !== field) {
        setSortDir("desc");
        return field;
      }
      if (sortDir === "desc") {
        setSortDir("asc");
        return field;
      }
      return null; // 3r clic: treu el filtre d'ordre
    });
  };

  const sortValue = useCallback(
    (r) => {
      if (sortField === "lectura") {
        const activity = lastByEui && r.eui ? lastByEui.get(r.eui) : null;
        return activity && activity.lastReal ? new Date(activity.lastReal) : null;
      }
      if (sortField === "muntatge") return parseMuntatgeDate(r.date);
      return null;
    },
    [sortField, lastByEui]
  );

  const sortWells = useCallback(
    (wells) => {
      if (!sortField) return [...wells].sort((a, b) => a.pozo.localeCompare(b.pozo));
      return [...wells].sort((a, b) => {
        const va = sortValue(a);
        const vb = sortValue(b);
        if (va && vb) return sortDir === "desc" ? vb - va : va - vb;
        if (va) return -1;
        if (vb) return 1;
        return a.pozo.localeCompare(b.pozo);
      });
    },
    [sortValue, sortDir, sortField]
  );

  const filteredRows = useMemo(() => {
    if (!rows) return [];
    const q = query.trim().toLowerCase();
    return q ? rows.filter((r) => r.pozo.toLowerCase().includes(q)) : rows;
  }, [rows, query]);

  const groups = useMemo(() => {
    if (groupBy === "none") {
      return [{ key: "__all__", label: "Tots els pous", wells: filteredRows }];
    }
    const byKey = new Map();
    for (const r of filteredRows) {
      const key = groupBy === "illa" ? getIlla(r.pozo) : r.installer || "Sense instal·lador assignat";
      if (!byKey.has(key)) byKey.set(key, []);
      byKey.get(key).push(r);
    }
    return Array.from(byKey.entries())
      .map(([key, wells]) => ({ key, label: key, wells }))
      .sort((a, b) => a.label.localeCompare(b.label));
  }, [filteredRows, groupBy]);

  const pendents = useMemo(() => (rows || []).filter((r) => !parseMuntatgeDate(r.date)), [rows]);

  const notMountedCommunicating = useMemo(() => {
    if (!notMounted || !lastByEui) return [];
    return notMounted.filter((r) => r.eui && lastByEui.has(r.eui));
  }, [notMounted, lastByEui]);
  const toggle = (key) => setCollapsed((c) => ({ ...c, [key]: !c[key] }));

  const exportPdf = () => {
    const doc = new jsPDF();
    doc.setFontSize(14);
    doc.text("Muntatges — Instal·lacions JFA", 14, 15);
    doc.setFontSize(9);
    doc.setTextColor(120);
    doc.text(`Generat el ${new Date().toLocaleString("es-ES")}`, 14, 21);

    const extraCols = lastByEui
      ? [
          "Última lectura real",
          ...(visibleCols.battery ? ["Bateria (%)"] : []),
          ...(visibleCols.errors ? ["Errors Modbus"] : []),
          ...(visibleCols.gpio ? ["GPIO 1/2"] : []),
          ...(visibleCols.deviceInfo ? ["Info dispositiu"] : []),
        ]
      : [];
    const head = [["Pou", "DevEUI", "Illa", "Remesa", "Instal·lador", "Data de muntatge", ...extraCols]];
    const body = [];
    groups.forEach((g) => {
      sortWells(g.wells).forEach((r) => {
        const activity = lastByEui && r.eui ? lastByEui.get(r.eui) : null;
        const row = [r.pozo, r.eui ? r.eui.toUpperCase() : "—", getIlla(r.pozo), r.remesa || "—", r.installer || "—", r.date || "pendent"];
        if (lastByEui) {
          row.push(!r.eui ? "—" : activity && activity.lastReal ? activity.lastReal.slice(0, 16).replace("T", " ") : "sense lectures reals");
          if (visibleCols.battery) row.push(activity && activity.battery !== null && activity.battery !== undefined ? activity.battery : "—");
          if (visibleCols.errors) row.push(activity && activity.errors && activity.errors.length ? activity.errors.join(", ") : "—");
          if (visibleCols.gpio)
            row.push(activity && (activity.gpio1 !== undefined || activity.gpio2 !== undefined) ? `${activity.gpio1 ?? "—"} / ${activity.gpio2 ?? "—"}` : "—");
          if (visibleCols.deviceInfo) row.push((activity && activity.deviceInfo) || "—");
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

        <div style={styles.controlsRow}>
          <div style={styles.segmentGroup}>
            <span style={styles.segmentLabel}>Agrupar per:</span>
            {[
              ["installer", "Instal·lador"],
              ["illa", "Illa"],
              ["none", "Tot junt"],
            ].map(([val, label]) => (
              <button
                key={val}
                style={{ ...styles.segmentBtn, ...(groupBy === val ? styles.segmentBtnActive : {}) }}
                onClick={() => setGroupBy(val)}
              >
                {label}
              </button>
            ))}
          </div>
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
          <button
            style={{ ...styles.secondaryBtn, ...(showMap ? styles.segmentBtnActive : {}) }}
            onClick={() => setShowMap((s) => !s)}
            disabled={!coordsByPozo}
          >
            <MapIcon size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
            {coordsByPozo ? (showMap ? "Amagar mapa" : "Veure mapa") : "Carregant mapa…"}
          </button>
          <button style={styles.secondaryBtn} onClick={exportPdf} disabled={!rows || !rows.length}>
            <Download size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
            Descarregar PDF
          </button>
          <button style={styles.secondaryBtn} onClick={load} disabled={loading}>
            <RefreshCw size={14} style={{ marginRight: 6, verticalAlign: "-2px" }} />
            {loading ? "Actualitzant…" : "Actualitzar ara"}
          </button>
        </div>

        {coordsError && (
          <div style={styles.errorBox}>
            <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>No s'ha pogut carregar el mapa de coordenades: {coordsError}</span>
          </div>
        )}

        {showMap && coordsByPozo && rows && (
          <PousMap rows={rows} coordsByPozo={coordsByPozo} lastByEui={lastByEui} query={query} />
        )}

        {lastByEui && (
          <div style={styles.segmentGroup}>
            <span style={styles.segmentLabel}>Columnes:</span>
            {[
              ["battery", "Bateria"],
              ["errors", "Errors Modbus"],
              ["gpio", "GPIO"],
              ["deviceInfo", "Info dispositiu"],
            ].map(([key, label]) => (
              <button
                key={key}
                style={{ ...styles.segmentBtn, ...(visibleCols[key] ? styles.segmentBtnActive : {}) }}
                onClick={() => toggleCol(key)}
              >
                {label}
              </button>
            ))}
          </div>
        )}

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
              {groupBy !== "none" && <Stat label={groupBy === "illa" ? "Illes" : "Instal·ladors"} value={groups.length} />}
            </div>

            <div style={styles.groupsWrap}>
              {groups.map((g) => (
                <div key={g.key} style={styles.groupCard}>
                  {groupBy !== "none" && (
                    <button style={styles.groupHeader} onClick={() => toggle(g.key)}>
                      {collapsed[g.key] ? <ChevronRight size={16} /> : <ChevronDown size={16} />}
                      <CheckCircle2 size={16} color="#3e8e86" style={{ marginRight: 6 }} />
                      <span style={styles.groupTitle}>{g.label}</span>
                      <span style={styles.groupMeta}>{g.wells.length} pous</span>
                    </button>
                  )}
                  {(groupBy === "none" || !collapsed[g.key]) && (
                    <div style={styles.tableWrap}>
                      <table style={styles.table}>
                        <thead>
                          <tr>
                            <th style={styles.th}>Pou</th>
                            <th style={styles.th}>DevEUI</th>
                            {groupBy !== "illa" && <th style={styles.th}>Illa</th>}
                            <th style={styles.th}>Remesa</th>
                            {groupBy !== "installer" && <th style={styles.th}>Instal·lador</th>}
                            <th
                              style={{ ...styles.th, cursor: "pointer", userSelect: "none" }}
                              onClick={() => toggleSort("muntatge")}
                            >
                              Data de muntatge
                              {sortField === "muntatge" &&
                                (sortDir === "desc" ? (
                                  <ChevronDown size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                                ) : (
                                  <ChevronUp size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                                ))}
                            </th>
                            {lastByEui && (
                              <th
                                style={{ ...styles.th, cursor: "pointer", userSelect: "none" }}
                                onClick={() => toggleSort("lectura")}
                              >
                                Última lectura real
                                {sortField === "lectura" &&
                                  (sortDir === "desc" ? (
                                    <ChevronDown size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                                  ) : (
                                    <ChevronUp size={12} style={{ marginLeft: 4, verticalAlign: "-1px" }} />
                                  ))}
                              </th>
                            )}
                            {lastByEui && visibleCols.battery && <th style={styles.th}>Bateria (%)</th>}
                            {lastByEui && visibleCols.errors && <th style={styles.th}>Errors Modbus</th>}
                            {lastByEui && visibleCols.gpio && <th style={styles.th}>GPIO 1/2</th>}
                            {lastByEui && visibleCols.deviceInfo && <th style={styles.th}>Info dispositiu</th>}
                            {historyByEui && <th style={styles.th}>Historial</th>}
                          </tr>
                        </thead>
                        <tbody>
                          {sortWells(g.wells).map((r, i) => {
                            const activity = lastByEui && r.eui ? lastByEui.get(r.eui) : null;
                            const hist = historyByEui && r.eui ? historyByEui.get(r.eui) : null;
                            const isExpanded = r.eui && expandedHistory[r.eui];
                            return (
                              <React.Fragment key={i}>
                                <tr>
                                  <td style={{ ...styles.td, fontWeight: 600 }}>{r.pozo}</td>
                                  <td style={{ ...styles.td, fontFamily: "monospace", fontSize: 12 }}>
                                    {r.eui ? r.eui.toUpperCase() : "—"}
                                  </td>
                                  {groupBy !== "illa" && <td style={styles.td}>{getIlla(r.pozo)}</td>}
                                  <td style={styles.td}>{r.remesa || "—"}</td>
                                  {groupBy !== "installer" && <td style={styles.td}>{r.installer || "—"}</td>}
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
                                  {lastByEui && visibleCols.battery && (
                                    <td style={styles.td}>
                                      {activity && activity.battery !== null && activity.battery !== undefined
                                        ? activity.battery
                                        : "—"}
                                    </td>
                                  )}
                                  {lastByEui && visibleCols.errors && (
                                    <td style={styles.td}>
                                      {activity && activity.errors && activity.errors.length ? (
                                        <span style={{ color: "#b03a3a" }}>{activity.errors.join(", ")}</span>
                                      ) : (
                                        "—"
                                      )}
                                    </td>
                                  )}
                                  {lastByEui && visibleCols.gpio && (
                                    <td style={styles.td}>
                                      {activity && (activity.gpio1 !== undefined || activity.gpio2 !== undefined)
                                        ? `${activity.gpio1 ?? "—"} / ${activity.gpio2 ?? "—"}`
                                        : "—"}
                                    </td>
                                  )}
                                  {lastByEui && visibleCols.deviceInfo && (
                                    <td style={styles.td}>{(activity && activity.deviceInfo) || "—"}</td>
                                  )}
                                  {historyByEui && (
                                    <td style={styles.td}>
                                      {hist && hist.length ? (
                                        <button style={styles.smallLinkBtn} onClick={() => toggleHistory(r.eui)}>
                                          {isExpanded ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
                                          {" "}
                                          {hist.length} lectures
                                        </button>
                                      ) : (
                                        "—"
                                      )}
                                    </td>
                                  )}
                                </tr>
                                {isExpanded && hist && (
                                  <tr>
                                    <td colSpan={20} style={{ ...styles.td, background: "#f7faf9" }}>
                                      <ReadingHistorySummary history={hist} />
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
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

        {notMountedCommunicating.length > 0 && (
          <div style={styles.noSignalBox}>
            <button style={styles.noSignalHeader} onClick={() => setShowNotMounted((s) => !s)}>
              {showNotMounted ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
              <RadioTower size={15} color="#b03a3a" style={{ margin: "0 6px" }} />
              {notMountedCommunicating.length} pous que comuniquen però NO consten com a muntats (sense "Cable fins cota" vàlid)
            </button>
            {showNotMounted && (
              <div style={styles.tableWrap}>
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={styles.th}>Pou</th>
                      <th style={styles.th}>DevEUI</th>
                      <th style={styles.th}>Illa</th>
                      <th style={styles.th}>Instal·lador</th>
                      <th style={styles.th}>Cota</th>
                      <th style={styles.th}>Última comunicació</th>
                    </tr>
                  </thead>
                  <tbody>
                    {notMountedCommunicating.map((r, i) => {
                      const activity = lastByEui.get(r.eui);
                      return (
                        <tr key={i}>
                          <td style={{ ...styles.td, fontWeight: 600 }}>{r.pozo}</td>
                          <td style={{ ...styles.td, fontFamily: "monospace", fontSize: 12 }}>{r.eui.toUpperCase()}</td>
                          <td style={styles.td}>{getIlla(r.pozo)}</td>
                          <td style={styles.td}>{r.installer || "—"}</td>
                          <td style={styles.td}>{r.cota !== null && !Number.isNaN(r.cota) ? fmt(r.cota, 1) + " m" : "—"}</td>
                          <td style={styles.td}>{activity.lastAny ? activity.lastAny.slice(0, 16).replace("T", " ") : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
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
  const [expandedExcluded, setExpandedExcluded] = useState({});
  const toggleExpandedExcluded = (pozo) => setExpandedExcluded((c) => ({ ...c, [pozo]: !c[pozo] }));
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
      const excludedGroups = new Map(); // pozo -> date -> {pressures,temps,conds}
      const excludedByEui = new Map(assocExcluded.map((x) => [x.eui, x]));
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
        if (!info) {
          const exclInfo = excludedByEui.get(eui);
          if (exclInfo) {
            const decodedExcl = decodePayload(row[rPayloadIdx]);
            if (decodedExcl && decodedExcl.pressureBar !== undefined) {
              const ts = row[rTsIdx];
              const date = String(ts || "").slice(0, 10) || "sin-fecha";
              if (!excludedGroups.has(exclInfo.pozo)) excludedGroups.set(exclInfo.pozo, new Map());
              const byDate = excludedGroups.get(exclInfo.pozo);
              if (!byDate.has(date)) byDate.set(date, { pressures: [], temps: [], conds: [] });
              const bucket = byDate.get(date);
              bucket.pressures.push(decodedExcl.pressureBar);
              if (decodedExcl.tempC !== undefined) bucket.temps.push(decodedExcl.tempC);
              if (decodedExcl.condMScm !== undefined) bucket.conds.push(decodedExcl.condMScm);
            }
          }
          continue;
        }
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

      const excludedReadingsByPozo = new Map();
      for (const [pozo, byDate] of excludedGroups.entries()) {
        const dates = Array.from(byDate.entries())
          .map(([date, b]) => ({
            date,
            n: b.pressures.length,
            pressureBar: avg(b.pressures),
            pressureMH2O: avg(b.pressures) * BAR_TO_MH2O,
            tempC: avg(b.temps),
            condMScm: avg(b.conds),
          }))
          .sort((a, b) => (a.date < b.date ? 1 : -1));
        excludedReadingsByPozo.set(pozo, dates);
      }

      setResults({
        groups: groupList,
        excluded: [...assocExcluded].sort((a, b) => a.pozo.localeCompare(b.pozo)),
        excludedReadingsByPozo,
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
          "Presión media (bar)": d.pressureBar !== null ? +d.pressureBar.toFixed(4) : "",
          "Presión media (mH2O)": d.pressureMH2O !== null ? +d.pressureMH2O.toFixed(4) : "",
          "Temperatura media (°C)": d.tempC !== null ? +d.tempC.toFixed(2) : "",
          "Conductividad media (mS/cm)": d.condMScm !== null ? +d.condMScm.toFixed(3) : "",
          "Cota (m)": d.cota,
          "Cable fins cota (m)": d.cable,
          "Nivel freático (m)": d.nivel !== null ? +d.nivel.toFixed(3) : "",
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
                    <div>
                      <NivellChart dates={g.dates} />
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
                          <th style={styles.th}>Lectures</th>
                        </tr>
                      </thead>
                      <tbody>
                        {results.excluded.map((x, i) => {
                          const exclDates = results.excludedReadingsByPozo && results.excludedReadingsByPozo.get(x.pozo);
                          const isOpen = expandedExcluded[x.pozo];
                          return (
                            <React.Fragment key={i}>
                              <tr>
                                <td style={styles.td}>{x.pozo}</td>
                                <td style={{ ...styles.td, fontFamily: "monospace", fontSize: 12 }}>{x.eui}</td>
                                <td style={styles.td}>{x.motivo}</td>
                                <td style={styles.td}>
                                  {exclDates && exclDates.length ? (
                                    <button style={styles.smallLinkBtn} onClick={() => toggleExpandedExcluded(x.pozo)}>
                                      {isOpen ? <ChevronUp size={12} /> : <ChevronDown size={12} />} {exclDates.length} dies
                                    </button>
                                  ) : (
                                    "—"
                                  )}
                                </td>
                              </tr>
                              {isOpen && exclDates && (
                                <tr>
                                  <td colSpan={4} style={{ ...styles.td, background: "#fffaf3" }}>
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
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {exclDates.map((d) => (
                                            <tr key={d.date}>
                                              <td style={styles.td}>{d.date}</td>
                                              <td style={styles.td}>{d.n}</td>
                                              <td style={styles.td}>{fmt(d.pressureBar, 4)}</td>
                                              <td style={styles.td}>{fmt(d.pressureMH2O, 3)}</td>
                                              <td style={styles.td}>{fmt(d.tempC, 2)}</td>
                                              <td style={styles.td}>{d.condMScm !== null ? fmt(d.condMScm, 3) : "—"}</td>
                                            </tr>
                                          ))}
                                        </tbody>
                                      </table>
                                    </div>
                                    <div style={{ fontSize: 11.5, color: "#a86a2d", marginTop: 6 }}>
                                      Sense "Cable fins cota" vàlid no es pot calcular el nivell freàtic — només es mostra la
                                      presió, temperatura i conductivitat.
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        })}
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

function NivellChart({ dates }) {
  const points = [...dates].filter((d) => d.nivel !== null && !Number.isNaN(d.nivel)).sort((a, b) => (a.date < b.date ? -1 : 1));
  if (points.length < 2) return null;

  const W = 720;
  const H = 160;
  const padL = 52;
  const padR = 16;
  const padT = 14;
  const padB = 28;
  const values = points.map((p) => p.nivel);
  let min = Math.min(...values);
  let max = Math.max(...values);
  if (min === max) {
    min -= 0.5;
    max += 0.5;
  }
  const pad = (max - min) * 0.08;
  min -= pad;
  max += pad;

  const x = (i) => padL + (i / (points.length - 1)) * (W - padL - padR);
  const y = (v) => padT + (1 - (v - min) / (max - min)) * (H - padT - padB);

  const path = points.map((p, i) => `${i === 0 ? "M" : "L"}${x(i).toFixed(1)},${y(p.nivel).toFixed(1)}`).join(" ");
  const ticksY = 4;
  const yTickVals = Array.from({ length: ticksY + 1 }, (_, i) => min + ((max - min) * i) / ticksY);
  const xLabelEvery = Math.max(1, Math.ceil(points.length / 6));

  return (
    <div style={{ padding: "14px 18px 4px", background: "#fff" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}>
        {yTickVals.map((v, i) => (
          <g key={i}>
            <line x1={padL} x2={W - padR} y1={y(v)} y2={y(v)} stroke="#eef4f3" strokeWidth="1" />
            <text x={padL - 8} y={y(v) + 4} textAnchor="end" fontSize="10" fill="#7c9490">
              {v.toFixed(2)}
            </text>
          </g>
        ))}
        {points.map(
          (p, i) =>
            i % xLabelEvery === 0 && (
              <text key={i} x={x(i)} y={H - 8} textAnchor="middle" fontSize="9.5" fill="#7c9490">
                {p.date.slice(5)}
              </text>
            )
        )}
        <path d={path} fill="none" stroke="#1f5e59" strokeWidth="2" />
        {points.map((p, i) => (
          <circle key={i} cx={x(i)} cy={y(p.nivel)} r="3" fill="#1f5e59">
            <title>
              {p.date}: {p.nivel.toFixed(3)} m
            </title>
          </circle>
        ))}
      </svg>
      <div style={{ fontSize: 11, color: "#9fb3ae", marginBottom: 4 }}>Nivell freàtic (m) per dia</div>
    </div>
  );
}

function PousMap({ rows, coordsByPozo, lastByEui, query }) {
  const mapRef = useRef(null);
  const mapInstance = useRef(null);
  const markersLayer = useRef(null);

  const q = (query || "").trim().toLowerCase();
  const filtered = q ? rows.filter((r) => r.pozo.toLowerCase().includes(q)) : rows;

  const points = filtered
    .map((r) => {
      const coord = coordsByPozo.get(r.pozo);
      if (!coord) return null;
      let status = "unknown"; // sense fitxer de lectures carregat
      if (lastByEui) {
        const activity = r.eui ? lastByEui.get(r.eui) : null;
        status = activity && activity.lastReal ? "ok" : "sense";
      }
      return { pozo: r.pozo, lat: coord.lat, lon: coord.lon, status, installer: r.installer, date: r.date };
    })
    .filter(Boolean);

  useEffect(() => {
    if (!mapRef.current || mapInstance.current) return;
    mapInstance.current = L.map(mapRef.current).setView([39.6, 2.9], 9);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
      maxZoom: 19,
    }).addTo(mapInstance.current);
    markersLayer.current = L.layerGroup().addTo(mapInstance.current);
    return () => {
      mapInstance.current.remove();
      mapInstance.current = null;
    };
  }, []);

  useEffect(() => {
    if (!mapInstance.current || !markersLayer.current) return;
    markersLayer.current.clearLayers();
    const colors = { ok: "#2a8f6c", sense: "#b03a3a", unknown: "#7c9490" };
    const bounds = [];
    points.forEach((p) => {
      bounds.push([p.lat, p.lon]);
      const marker = L.circleMarker([p.lat, p.lon], {
        radius: 7,
        color: "#fff",
        weight: 1.5,
        fillColor: colors[p.status],
        fillOpacity: 0.9,
      });
      const statusLabel =
        p.status === "ok" ? "Comunica" : p.status === "sense" ? "Sense lectures reals" : "Estat desconegut (puja el fitxer de lectures)";
      marker.bindPopup(
        `<strong>${p.pozo}</strong><br/>${statusLabel}${p.installer ? "<br/>Instal·lador: " + p.installer : ""}${
          p.date ? "<br/>Muntatge: " + p.date : ""
        }`
      );
      marker.addTo(markersLayer.current);
    });
    if (bounds.length) mapInstance.current.fitBounds(bounds, { padding: [30, 30] });
  }, [points]);

  return (
    <div style={{ marginTop: 16 }}>
      <div ref={mapRef} style={{ height: 480, borderRadius: 10, border: "1px solid #e1ecea" }} />
      <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 12.5, color: "#4a615d", flexWrap: "wrap" }}>
        <span>
          <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: "50%", background: "#2a8f6c", marginRight: 5 }} />
          Comunica
        </span>
        <span>
          <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: "50%", background: "#b03a3a", marginRight: 5 }} />
          Sense lectures reals
        </span>
        <span>
          <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: "50%", background: "#7c9490", marginRight: 5 }} />
          {lastByEui ? "Sense coordenades / no aplica" : "Puja el fitxer de lectures per veure l'estat"}
        </span>
        <span style={{ marginLeft: "auto" }}>{points.length} pous al mapa</span>
      </div>
    </div>
  );
}

function ReadingHistorySummary({ history }) {
  const counts = { ok: 0, parcial: 0, error: 0, info: 0, sense_dades: 0 };
  history.forEach((h) => counts[h.status]++);
  const total = history.length;
  const pct = (n) => (total ? Math.round((n / total) * 100) : 0);

  return (
    <div style={{ padding: "8px 4px" }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 12.5, marginBottom: 10 }}>
        <span>
          <strong>{total}</strong> lectures totals
        </span>
        {counts.ok > 0 && (
          <span style={{ color: READING_STATUS_COLORS.ok }}>
            ● {counts.ok} correctes ({pct(counts.ok)}%)
          </span>
        )}
        {counts.parcial > 0 && (
          <span style={{ color: READING_STATUS_COLORS.parcial }}>
            ● {counts.parcial} parcials ({pct(counts.parcial)}%)
          </span>
        )}
        {counts.error > 0 && (
          <span style={{ color: READING_STATUS_COLORS.error }}>
            ● {counts.error} amb error ({pct(counts.error)}%)
          </span>
        )}
        {counts.info > 0 && (
          <span style={{ color: "#7c9490" }}>
            ● {counts.info} sense mesura ({pct(counts.info)}%)
          </span>
        )}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 2, maxWidth: "100%" }}>
        {history.map((h, i) => (
          <div
            key={i}
            title={`${h.ts.slice(0, 16).replace("T", " ")} — ${READING_STATUS_LABELS[h.status]}${
              h.errors && h.errors.length ? " (" + h.errors.join(", ") + ")" : ""
            }`}
            style={{
              width: 10,
              height: 10,
              borderRadius: 2,
              background: READING_STATUS_COLORS[h.status],
              cursor: "default",
            }}
          />
        ))}
      </div>
      <div style={{ fontSize: 11, color: "#9fb3ae", marginTop: 6 }}>
        Ordre cronològic (esquerra = més antic). Passa el cursor per sobre de cada quadrat per veure la data i el detall.
      </div>
    </div>
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
  controlsRow: { display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap", marginTop: 22 },
  segmentGroup: { display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  segmentLabel: { fontSize: 12.5, color: "#7c9490", marginRight: 2 },
  segmentBtn: {
    background: "#fff",
    border: "1px solid #e1ecea",
    borderRadius: 7,
    padding: "6px 12px",
    fontSize: 12.5,
    color: "#4a615d",
    cursor: "pointer",
  },
  segmentBtnActive: { background: "#1f5e59", borderColor: "#1f5e59", color: "#fff", fontWeight: 600 },
  segmentBtnDisabled: { opacity: 0.4, cursor: "not-allowed" },
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
