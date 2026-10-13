"use client";

import { useCallback, useDeferredValue, useEffect, useMemo, useRef, useState } from "react";

/*

Evon
 * Excel Workspace — single-file Next.js App Router page.
 * Sections are deliberately isolated for later extraction:
 * file format/parser, IndexedDB, selection/search, combining, export, UI.
 * Requires: xlsx (SheetJS CE 0.20.3). Uses no server routes.
 */

const DB_NAME = "excel-workspace-v1";
const DB_VERSION = 1;
const MAX_FILE_BYTES = 60 * 1024 * 1024;
const MAX_CELLS_PER_WORKBOOK = 650_000;
const MAX_ROWS_PER_SHEET = 200_000;
const PAGE_SIZE = 75;
const PREVIEW_LIMIT = 100_000;
const OPERATORS = ["contains", "equals", "not equals", "starts with", "ends with", "greater than", "less than", "empty", "not empty"];
const uid = () => (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `id-${Date.now()}-${Math.random().toString(36).slice(2)}`);
const normal = (value) => String(value ?? "").toLocaleLowerCase().trim();
const alphabet = (index) => { let n = index + 1, result = ""; while (n > 0) { n--; result = String.fromCharCode(65 + (n % 26)) + result; n = Math.floor(n / 26); } return result; };
const textValue = (value) => value == null ? "" : value instanceof Date ? value.toISOString() : String(value);
const displayValue = (cell) => cell?.w != null ? String(cell.w) : textValue(cell?.v);
const safeSheetName = (name) => String(name).replace(/[\\/?*\[\]:]/g, "_").slice(0, 31) || "Sheet";
const cleanExportValue = (value) => { const str = textValue(value); return /^[\s\x00-\x1f]*[=+@\-\t\r]/.test(str) ? `'${str}` : str; };

// ─── IndexedDB: large structured session data (not localStorage) ───
function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("state")) db.createObjectStore("state");
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
async function readSession() {
  const db = await openDatabase();
  try { return await new Promise((resolve, reject) => {
    const tx = db.transaction("state", "readonly");
    const request = tx.objectStore("state").get("workspace");
    request.onsuccess = () => resolve(request.result ?? null);
    request.onerror = () => reject(request.error);
  }); } finally { db.close(); }
}
async function writeSession(data) {
  const db = await openDatabase();
  try { await new Promise((resolve, reject) => {
    const tx = db.transaction("state", "readwrite");
    tx.objectStore("state").put(data, "workspace");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || Error("Storage transaction aborted"));
  }); } finally { db.close(); }
}

// ─── Workbook detection and import ───
function detectFormat(bytes, name) {
  const ext = (name.split(".").pop() || "").toLowerCase();
  const ole = bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0;
  const zip = bytes[0] === 0x50 && bytes[1] === 0x4B;
  if (ole) return { family: "BIFF / OLE", label: "Legacy Excel (.xls)", kind: "binary" };
  if (zip) return { family: "ZIP / OpenXML", label: ext === "xlsb" ? "Excel binary (.xlsb)" : ext === "xlsm" ? "Macro-enabled (.xlsm)" : ext === "ods" ? "OpenDocument (.ods)" : "Excel OpenXML (.xlsx)", kind: "binary" };
  if (["csv", "tsv", "txt"].includes(ext)) return { family: "Delimited text", label: ext.toUpperCase(), kind: "text" };
  if (["xlsx", "xlsm", "xlsb", "xls", "ods"].includes(ext)) throw Error("File signature does not match a supported workbook container.");
  throw Error("Unsupported file type. Use XLS, XLSX, XLSM, XLSB, ODS, CSV, or TSV.");
}
async function parseFile(file, XLSX) {
  if (file.size > MAX_FILE_BYTES) throw Error(`${file.name}: over the 60 MB per-file limit.`);
  const bytes = new Uint8Array(await file.arrayBuffer());
  const format = detectFormat(bytes, file.name);
  const wb = XLSX.read(bytes, { type: "array", cellFormula: true, cellText: true, cellDates: true, bookVBA: false, WTF: false, dense: false });
  let totalCells = 0;
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const raw = ws?.["!ref"];
    const bounds = raw ? XLSX.utils.decode_range(raw) : null;
    if (bounds && bounds.e.r - bounds.s.r + 1 > MAX_ROWS_PER_SHEET) throw Error(`${name}: exceeds 200,000 rows. Split this sheet before importing.`);
    const rowMap = new Map();
    let formulaCount = 0;
    let uncachedCount = 0;
    let maxCol = 0;
    for (const [address, cell] of Object.entries(ws || {})) {
      if (address.startsWith("!")) continue;
      if (!/^[A-Z]+[1-9]\d*$/i.test(address)) continue;
      if (++totalCells > MAX_CELLS_PER_WORKBOOK) throw Error(`${file.name}: exceeds ${MAX_CELLS_PER_WORKBOOK.toLocaleString()} populated cells. No data was partially imported.`);
      const { r, c } = XLSX.utils.decode_cell(address);
      if (r >= MAX_ROWS_PER_SHEET) throw Error(`${name}: row limit exceeded.`);
      if (c > 300) throw Error(`${name}: more than 301 columns are populated. This initial version limits sheets to column KO.`);
      maxCol = Math.max(maxCol, c);
      if (cell?.f) { formulaCount++; if (cell.v == null) uncachedCount++; }
      if (!rowMap.has(r)) rowMap.set(r, {});
      rowMap.get(r)[c] = {
        v: cell?.v instanceof Date ? cell.v.toISOString() : (cell?.v ?? null),
        w: cell?.w ?? null,
        f: cell?.f ?? null,
        t: cell?.t || null,
      };
    }
    const rows = [...rowMap.entries()].sort((a, b) => a[0] - b[0]).map(([r, cells]) => ({ r, cells }));
    const headerRow = rows[0]?.r ?? 0;
    return { id: uid(), name, rows, headerRow, maxCol, formulaCount, uncachedCount };
  });
  return { id: uid(), name: file.name, size: file.size, format: format.label, family: format.family, imported: Date.now(), sheets };
}

// ─── Selection/filter engine: data only, never eval / SQL / formulas ───
function columnsFor(sheet) {
  const h = sheet?.rows.find((row) => row.r === sheet.headerRow);
  const seen = new Map();
  const columns = [];
  for (let i = 0; i <= Math.min(sheet?.maxCol ?? 0, 300); i++) {
    const raw = displayValue(h?.cells[i]).trim();
    const base = raw || `Column ${alphabet(i)}`;
    const next = (seen.get(base) || 0) + 1;
    seen.set(base, next);
    columns.push({ index: i, key: `${alphabet(i)} · ${base}${next > 1 ? ` (${next})` : ""}`, name: base, letter: alphabet(i) });
  }
  return columns;
}
function levenshteinWithin(a, b, threshold = 2) {
  if (Math.abs(a.length - b.length) > threshold) return false;
  let previous = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const current = [i]; let min = current[0];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      min = Math.min(min, current[j]);
    }
    if (min > threshold) return false;
    previous = current;
  }
  return previous[b.length] <= threshold;
}
function fuzzyIncludes(value, query) {
  const a = normal(value), q = normal(query);
  if (!q) return true;
  if (a.includes(q)) return true;
  if (q.length < 3) return false;
  const words = a.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
  return words.some((word) => levenshteinWithin(word, q, q.length >= 7 ? 2 : 1));
}
function passCondition(cell, condition) {
  const value = displayValue(cell);
  const left = normal(value), right = normal(condition.value);
  switch (condition.op) {
    case "contains": return left.includes(right);
    case "equals": return left === right;
    case "not equals": return left !== right;
    case "starts with": return left.startsWith(right);
    case "ends with": return left.endsWith(right);
    case "empty": return left === "";
    case "not empty": return left !== "";
    case "greater than": return value !== "" && right !== "" && !Number.isNaN(Number(value)) && !Number.isNaN(Number(condition.value)) && Number(value) > Number(condition.value);
    case "less than": return value !== "" && right !== "" && !Number.isNaN(Number(value)) && !Number.isNaN(Number(condition.value)) && Number(value) < Number(condition.value);
    default: return false;
  }
}
function parseRange(input, XLSX) {
  const trimmed = input.trim().toUpperCase();
  if (!trimmed) return null;
  if (!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(trimmed)) throw Error("Use a cell or range such as B2 or A2:D40.");
  const range = XLSX.utils.decode_range(trimmed);
  if (range.s.r > range.e.r || range.s.c > range.e.c || range.e.r >= MAX_ROWS_PER_SHEET || range.e.c > 300) throw Error("Invalid or oversized range.");
  return range;
}
function buildSelection({ sheet, columns, selected, firstRow, lastRow, search, searchMode, searchColumn, conditions, range }) {
  if (!sheet) return { rows: [], columns: [] };
  const available = columns.filter((col) => selected.includes(col.index));
  const bounded = available.filter((col) => !range || (col.index >= range.s.c && col.index <= range.e.c));
  const matchRow = (row) => {
    const excelRow = row.r + 1;
    if (excelRow < firstRow || excelRow > lastRow || (range && (row.r < range.s.r || row.r > range.e.r))) return false;
    if (search.trim()) {
      const candidates = searchColumn === "all" ? columns.map((col) => row.cells[col.index]) : [row.cells[Number(searchColumn)]];
      const found = candidates.some((cell) => searchMode === "exact" ? normal(displayValue(cell)) === normal(search) : fuzzyIncludes(displayValue(cell), search));
      if (!found) return false;
    }
    if (!conditions.every((c) => passCondition(row.cells[Number(c.column)], c))) return false;
    return true;
  };
  const rows = sheet.rows.filter((row) => row.r !== sheet.headerRow && matchRow(row)).map((row) => ({
    sourceRow: row.r + 1,
    values: bounded.map((col) => displayValue(row.cells[col.index])),
  }));
  return { rows, columns: bounded.map((col) => col.key) };
}
function recordsFromSelection(selection) {
  return selection.rows.map((row) => Object.fromEntries(selection.columns.map((col, i) => [col, row.values[i] ?? ""])));
}
function appendTables(tables) {
  const headers = [...new Set(tables.flatMap((t) => t.columns))];
  return { columns: headers, records: tables.flatMap((t) => t.records.map((r) => Object.fromEntries(headers.map((h) => [h, r[h] ?? ""])))) };
}
function joinTables(left, right, leftKey, rightKey, how = "left") {
  const rightNames = right.columns.map((col) => left.columns.includes(col) ? `Right · ${col}` : col);
  const columns = [...left.columns, ...rightNames];
  const lookup = new Map();
  for (const r of right.records) {
    const key = normal(r[rightKey]);
    if (!key) continue;
    if (!lookup.has(key)) lookup.set(key, []);
    lookup.get(key).push(r);
  }
  const records = [];
  for (const l of left.records) {
    const key = normal(l[leftKey]);
    const matches = key ? lookup.get(key) || [] : [];
    if (!matches.length && how === "inner") continue;
    for (const r of matches.length ? matches : [null]) {
      const item = { ...l };
      right.columns.forEach((col, i) => { item[rightNames[i]] = r?.[col] ?? ""; });
      records.push(item);
      if (records.length > PREVIEW_LIMIT) throw Error(`Join exceeds ${PREVIEW_LIMIT.toLocaleString()} output rows. Narrow the inputs first.`);
    }
  }
  return { columns, records };
}

// ─── Export: values only, neutralize spreadsheet formula injection ───
async function exportTable(table, kind, filename) {
  if (!table.columns.length) throw Error("Nothing to export.");
  const XLSX = await import("xlsx");
  const matrix = [table.columns.map(cleanExportValue), ...table.records.map((record) => table.columns.map((column) => cleanExportValue(record[column])))];
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const safeName = (filename || "excel-workspace").replace(/[^a-z\d _.-]/gi, "_").slice(0, 70);
  if (kind === "csv") {
    const csv = "\uFEFF" + XLSX.utils.sheet_to_csv(ws);
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const a = document.createElement("a"); a.href = url; a.download = `${safeName}.csv`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } else {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, ws, safeSheetName("Results"));
    XLSX.writeFile(book, `${safeName}.xlsx`, { bookType: "xlsx", compression: true });
  }
}

// ─── Reusable UI pieces ───
const field = "w-full min-w-0 rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm text-slate-800 outline-none transition focus:border-slate-400 focus:ring-2 focus:ring-slate-100";
const button = "inline-flex items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium transition disabled:cursor-not-allowed disabled:opacity-40";
const ghost = `${button} border border-slate-200 bg-white text-slate-700 hover:bg-slate-50`;
const primary = `${button} bg-slate-900 text-white hover:bg-slate-800`;
function Panel({ title, action, children, className = "" }) { return <section className={`rounded-2xl border border-slate-200 bg-white shadow-sm ${className}`}><div className="flex items-center justify-between gap-3 border-b border-slate-100 px-5 py-4"><h2 className="text-sm font-semibold text-slate-900">{title}</h2>{action}</div><div className="p-5">{children}</div></section>; }
function Empty({ children }) { return <div className="rounded-xl border border-dashed border-slate-200 p-8 text-center text-sm text-slate-400">{children}</div>; }
function DataTable({ columns, records, sourceRows, page, onPage, empty = "No matching rows", limitNote }) {
  const pages = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
  const safePage = Math.min(page, pages - 1);
  const chunk = records.slice(safePage * PAGE_SIZE, (safePage + 1) * PAGE_SIZE);
  if (!columns.length) return <Empty>{empty}</Empty>;
  return <div className="min-w-0 overflow-hidden rounded-xl border border-slate-200"><div className="max-h-[470px] overflow-auto"><table className="w-full min-w-max border-collapse text-left text-xs"><thead className="sticky top-0 z-10 bg-slate-50"><tr>{sourceRows && <th className="border-b border-r border-slate-200 px-3 py-3 text-slate-500">Row</th>}{columns.map((c, i) => <th key={`${c}-${i}`} className="max-w-64 border-b border-r border-slate-200 px-4 py-3 font-semibold text-slate-700">{c}</th>)}</tr></thead><tbody>{chunk.map((r, index) => <tr key={index + safePage * PAGE_SIZE} className="hover:bg-slate-50">{sourceRows && <td className="border-b border-r border-slate-100 px-3 py-2 text-slate-400">{sourceRows[safePage * PAGE_SIZE + index]}</td>}{columns.map((c, i) => <td title={textValue(r[c])} key={`${c}-${i}`} className="max-w-64 truncate border-b border-r border-slate-100 px-4 py-2.5 text-slate-700">{textValue(r[c])}</td>)}</tr>)}</tbody></table></div><div className="flex flex-wrap items-center justify-between gap-3 bg-white px-4 py-3 text-xs text-slate-500"><span>{records.length.toLocaleString()} rows{limitNote ? ` · ${limitNote}` : ""}</span><div className="flex items-center gap-2"><button className={ghost} disabled={!safePage} onClick={() => onPage(safePage - 1)}>Previous</button><span>{safePage + 1} / {pages}</span><button className={ghost} disabled={safePage >= pages - 1} onClick={() => onPage(safePage + 1)}>Next</button></div></div></div>;
}

export default function Page() {
  const [workbooks, setWorkbooks] = useState([]);
  const [bins, setBins] = useState([]);
  const [queryBySheet, setQueryBySheet] = useState({});
  const [bookId, setBookId] = useState("");
  const [sheetId, setSheetId] = useState("");
  const [selected, setSelected] = useState([]);
  const [firstRow, setFirstRow] = useState("1");
  const [lastRow, setLastRow] = useState(String(MAX_ROWS_PER_SHEET));
  const [rangeText, setRangeText] = useState("");
  const [search, setSearch] = useState("");
  const [searchMode, setSearchMode] = useState("fuzzy");
  const [searchColumn, setSearchColumn] = useState("all");
  const [conditions, setConditions] = useState([]);
  const [binName, setBinName] = useState("");
  const [activeBinId, setActiveBinId] = useState("");
  const [combineMode, setCombineMode] = useState("append");
  const [leftId, setLeftId] = useState("");
  const [rightId, setRightId] = useState("");
  const [leftKey, setLeftKey] = useState("");
  const [rightKey, setRightKey] = useState("");
  const [joinHow, setJoinHow] = useState("left");
  const [combined, setCombined] = useState(null);
  const [view, setView] = useState("explore");
  const [page, setPage] = useState(0);
  const [binPage, setBinPage] = useState(0);
  const [busy, setBusy] = useState(false);
  const [hydrated, setHydrated] = useState(false);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [storageState, setStorageState] = useState("Loading saved workspace…");
  const fileRef = useRef(null);
  const saveTimer = useRef(null);
  const ready = useRef(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const saved = await readSession();
        if (cancelled) return;
        if (saved) { setWorkbooks(saved.workbooks || []); setBins(saved.bins || []); setQueryBySheet(saved.queryBySheet || {}); setBookId(saved.bookId || ""); setSheetId(saved.sheetId || ""); setActiveBinId(saved.activeBinId || ""); }
        try { const prefs = JSON.parse(localStorage.getItem("excel-workspace-prefs") || "{}"); setSearchMode(prefs.searchMode === "exact" ? "exact" : "fuzzy"); setView(prefs.view === "bins" ? "bins" : "explore"); } catch {}
        setStorageState("Saved locally");
      } catch (err) { setStorageState("Local storage unavailable"); setError(`Could not restore saved workspace: ${err.message}`); }
      finally { if (!cancelled) { ready.current = true; setHydrated(true); } }
    })();
    return () => { cancelled = true; };
  }, []);
  useEffect(() => {
    if (!ready.current || !hydrated) return;
    clearTimeout(saveTimer.current);
    setStorageState("Saving…");
    saveTimer.current = setTimeout(() => {
      writeSession({ workbooks, bins, queryBySheet, bookId, sheetId, activeBinId }).then(() => setStorageState("Saved locally")).catch((err) => { setStorageState("Not saved"); setError(`Cannot save locally: ${err.message}. Browser storage may be full.`); });
    }, 650);
    return () => clearTimeout(saveTimer.current);
  }, [workbooks, bins, queryBySheet, bookId, sheetId, activeBinId, hydrated]);
  useEffect(() => { if (hydrated) try { localStorage.setItem("excel-workspace-prefs", JSON.stringify({ searchMode, view })); } catch {} }, [searchMode, view, hydrated]);

  const book = workbooks.find((b) => b.id === bookId) || workbooks[0] || null;
  const sheet = book?.sheets.find((s) => s.id === sheetId) || book?.sheets[0] || null;
  const columns = useMemo(() => columnsFor(sheet), [sheet]);
  // Restore each sheet's last query without applying it to another sheet.
  useEffect(() => {
    const q = queryBySheet[sheet?.id] || {};
    setSelected(q.selected || columns.map((c) => c.index));
    setConditions(q.conditions || []);
    setSearchColumn(q.searchColumn || "all");
    setSearch(q.search || "");
    setRangeText(q.rangeText || "");
    setFirstRow(q.firstRow || "1");
    setLastRow(q.lastRow || String(MAX_ROWS_PER_SHEET));
    setPage(0);
  }, [sheet?.id]);
  const deferredSearch = useDeferredValue(search);
  const chosen = selected.length ? selected : [];
  const rangeState = useMemo(() => {
    if (!rangeText.trim()) return { range: null, error: "" };
    try { const match = /^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/i.exec(rangeText.trim()); if (!match) throw Error("Use A2:D40 or B3.");
      const colNum = (s) => s.toUpperCase().split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const r1 = Number(match[2]) - 1, r2 = Number(match[4] || match[2]) - 1, c1 = colNum(match[1]), c2 = colNum(match[3] || match[1]);
      if (r1 > r2 || c1 > c2 || r2 >= MAX_ROWS_PER_SHEET || c2 > 300) throw Error("Range is reversed or too large.");
      return { range: { s: { r: r1, c: c1 }, e: { r: r2, c: c2 } }, error: "" };
    } catch (err) { return { range: null, error: err.message }; }
  }, [rangeText]);
  const selectionResult = useMemo(() => {
    const min = Math.max(1, Number(firstRow) || 1), max = Math.min(MAX_ROWS_PER_SHEET, Number(lastRow) || MAX_ROWS_PER_SHEET);
    if (min > max || rangeState.error || !selected.length) return { rows: [], columns: [] };
    return buildSelection({ sheet, columns, selected: chosen, firstRow: min, lastRow: max, search: deferredSearch, searchMode, searchColumn, conditions, range: rangeState.range });
  }, [sheet, columns, chosen, firstRow, lastRow, deferredSearch, searchMode, searchColumn, conditions, rangeState]);
  // Persist filters explicitly on actions / input changes (debounced by the session writer).
  const captureQuery = (patch) => {
    const id = sheet?.id;
    if (!id) return;
    setQueryBySheet((prev) => ({ ...prev, [id]: {
      selected, firstRow, lastRow, rangeText, search, searchColumn, conditions,
      ...(prev[id] || {}), ...patch,
    } }));
  };
  const updateConditions = (value) => { setConditions(value); captureQuery({ conditions: value }); };
  const selectedRecords = useMemo(() => recordsFromSelection(selectionResult), [selectionResult]);
  const activeBin = bins.find((b) => b.id === activeBinId) || bins[0] || null;
  const left = bins.find((b) => b.id === leftId) || bins[0] || null;
  const right = bins.find((b) => b.id === rightId) || bins[1] || null;
  const formulaCount = sheet?.formulaCount || 0;
  const alert = (message) => { setError(""); setNotice(message); };
  const fail = (err) => { setNotice(""); setError(err.message || String(err)); };

  const importFiles = useCallback(async (files) => {
    const items = Array.from(files || []);
    if (!items.length) return;
    setBusy(true); setError(""); setNotice("");
    try {
      const XLSX = await import("xlsx");
      const incoming = [];
      const failures = [];
      for (const file of items) {
        try { incoming.push(await parseFile(file, XLSX)); }
        catch (err) { failures.push(`${file.name}: ${err.message}`); }
      }
      if (incoming.length) { setWorkbooks((old) => [...old, ...incoming]); setBookId(incoming[incoming.length - 1].id); setSheetId(incoming[incoming.length - 1].sheets[0]?.id || ""); setView("explore"); }
      if (failures.length) setError(failures.join("\n"));
      if (incoming.length) setNotice(`${incoming.length} workbook${incoming.length === 1 ? "" : "s"} imported. Original files stay on your device.`);
    } catch (err) { fail(err); } finally { setBusy(false); if (fileRef.current) fileRef.current.value = ""; }
  }, []);
  const removeBook = (id) => { setWorkbooks((items) => items.filter((b) => b.id !== id)); if (bookId === id) { setBookId(""); setSheetId(""); } setCombined(null); };
  const addBin = () => {
    if (!selectionResult.columns.length || !selectedRecords.length) return fail(Error("Select at least one column and one matching data row."));
    const name = binName.trim() || `${book.name} · ${sheet.name}`;
    const item = { id: uid(), name: name.slice(0, 90), columns: selectionResult.columns, records: selectedRecords, origin: `${book.name} / ${sheet.name}`, created: Date.now() };
    setBins((items) => [...items, item]); setActiveBinId(item.id); setBinName(""); alert(`${selectedRecords.length.toLocaleString()} rows saved to ${name}.`);
  };
  const createCombined = () => {
    try {
      if (bins.length < 2) throw Error("Create at least two bins first.");
      if (!left || !right || left.id === right.id) throw Error("Choose two different bins.");
      let output;
      if (combineMode === "append") output = appendTables([left, right]);
      else { if (!leftKey || !rightKey) throw Error("Choose a matching column for each bin."); output = joinTables(left, right, leftKey, rightKey, joinHow); }
      setCombined({ ...output, name: `${left.name} + ${right.name}` }); setBinPage(0); alert(`${output.records.length.toLocaleString()} combined rows ready.`);
    } catch (err) { fail(err); }
  };
  const saveCombined = () => { if (!combined) return; const next = { ...combined, id: uid(), origin: "Combined bins", created: Date.now() }; setBins((items) => [...items, next]); setActiveBinId(next.id); setCombined(null); alert("Combined result saved as a new bin."); };
  const doExport = async (table, kind, name) => { try { setBusy(true); await exportTable(table, kind, name); alert(`Exported ${kind.toUpperCase()}.`); } catch (err) { fail(err); } finally { setBusy(false); } };
  const currentTable = { columns: selectionResult.columns, records: selectedRecords };
  const shownTable = combined || (activeBin ? { columns: activeBin.columns, records: activeBin.records } : null);

  return <main className="min-h-screen bg-[#f7f8fa] text-slate-900">
    <div className="mx-auto max-w-[1550px] px-4 pb-20 pt-6 sm:px-6 lg:px-9">
      <header className="mb-6 flex flex-wrap items-center justify-between gap-4"><div><div className="flex items-center gap-3"><span className="grid h-10 w-10 place-items-center rounded-2xl bg-emerald-600 text-xl font-bold text-white">X</span><h1 className="text-2xl font-semibold tracking-tight">Excel Workspace</h1></div><p className="ml-[52px] mt-1 text-xs text-slate-500">{storageState}</p></div><div className="flex gap-2"><button className={ghost} onClick={() => setView("bins")}>Bins <span className="rounded-md bg-slate-100 px-2 py-0.5">{bins.length}</span></button><button className={primary} disabled={busy} onClick={() => fileRef.current?.click()}>{busy ? "Processing…" : "+ Add files"}</button><input ref={fileRef} type="file" hidden multiple accept=".xls,.xlsx,.xlsm,.xlsb,.csv,.tsv,.ods" onChange={(e) => importFiles(e.target.files)} /></div></header>
      {error && <div role="alert" className="mb-4 whitespace-pre-wrap rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">{error} <button className="float-right font-bold" onClick={() => setError("")}>×</button></div>}
      {notice && <div role="status" className="mb-4 rounded-xl border border-emerald-100 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{notice} <button className="float-right font-bold" onClick={() => setNotice("")}>×</button></div>}
      <nav className="mb-5 inline-flex gap-1 rounded-xl border border-slate-200 bg-white p-1"><button onClick={() => setView("explore")} className={`rounded-lg px-5 py-2 text-sm font-medium ${view === "explore" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}>Explore</button><button onClick={() => setView("bins")} className={`rounded-lg px-5 py-2 text-sm font-medium ${view === "bins" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}>Bins & Combine</button></nav>
      {view === "explore" ? <div className="grid items-start gap-5 lg:grid-cols-[300px_minmax(0,1fr)]">
        <aside className="space-y-5"><Panel title="Workbooks" action={<span className="text-xs text-slate-400">{workbooks.length}</span>}>{!workbooks.length ? <Empty>Import an Excel or CSV file to begin.</Empty> : <div className="space-y-2">{workbooks.map((b) => <div key={b.id} className={`rounded-xl border p-3 ${book?.id === b.id ? "border-emerald-300 bg-emerald-50/50" : "border-slate-200"}`}><div className="flex items-start justify-between gap-2"><button onClick={() => { setBookId(b.id); setSheetId(b.sheets[0]?.id || ""); }} className="min-w-0 flex-1 text-left"><div className="truncate text-sm font-medium" title={b.name}>{b.name}</div><div className="mt-1 text-xs text-slate-500">{b.format} · {b.sheets.length} sheets</div></button><button aria-label={`Remove ${b.name}`} onClick={() => removeBook(b.id)} className="rounded p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-600">×</button></div>{book?.id === b.id && <div className="mt-3 space-y-1 border-t border-emerald-100 pt-2">{b.sheets.map((s) => <button key={s.id} onClick={() => setSheetId(s.id)} className={`block w-full truncate rounded-lg px-3 py-2 text-left text-sm ${sheet?.id === s.id ? "bg-emerald-600 font-medium text-white" : "hover:bg-white"}`}>{s.name}</button>)}</div>}</div>)}</div>}</Panel>
          {sheet && <Panel title="Select data"><div className="space-y-4"><div><label className="mb-1.5 block text-xs font-medium text-slate-600">Cell range (optional)</label><input className={field} placeholder="e.g. A2:D150" value={rangeText} onChange={(e) => { setRangeText(e.target.value); captureQuery({ rangeText: e.target.value }); }} />{rangeState.error && <p className="mt-1 text-xs text-rose-600">{rangeState.error}</p>}</div><div className="grid grid-cols-2 gap-2"><label className="text-xs font-medium text-slate-600">From row<input type="number" min="1" max={MAX_ROWS_PER_SHEET} className={`${field} mt-1`} value={firstRow} onChange={(e) => { setFirstRow(e.target.value); captureQuery({ firstRow: e.target.value }); }} /></label><label className="text-xs font-medium text-slate-600">To row<input type="number" min="1" max={MAX_ROWS_PER_SHEET} className={`${field} mt-1`} value={lastRow} onChange={(e) => { setLastRow(e.target.value); captureQuery({ lastRow: e.target.value }); }} /></label></div><div className="flex items-center justify-between"><label className="text-xs font-medium text-slate-600">Columns ({selected.length})</label><div className="flex gap-3 text-xs"><button className="text-emerald-700" onClick={() => { const value = columns.map((c) => c.index); setSelected(value); captureQuery({ selected: value }); }}>All</button><button className="text-slate-500" onClick={() => { setSelected([]); captureQuery({ selected: [] }); }}>None</button></div></div><div className="max-h-52 space-y-1 overflow-auto rounded-xl border border-slate-200 p-2">{columns.map((col) => <label key={col.index} className="flex cursor-pointer items-center gap-2 rounded-lg px-2 py-1.5 text-xs hover:bg-slate-50"><input type="checkbox" className="accent-emerald-600" checked={selected.includes(col.index)} onChange={(e) => { const next = e.target.checked ? [...selected, col.index].sort((a, b) => a - b) : selected.filter((n) => n !== col.index); setSelected(next); captureQuery({ selected: next }); }} /><span className="w-7 shrink-0 font-medium text-slate-400">{col.letter}</span><span className="truncate" title={col.name}>{col.name}</span></label>)}</div></div></Panel>}
        </aside>
        <div className="min-w-0 space-y-5">{!sheet ? <Panel title="Get started"><Empty>Add a workbook. Choose a sheet, filter its data, and save selections into bins.</Empty></Panel> : <>
          <Panel title="Find & Filter" action={<button className="text-xs text-slate-500 hover:text-slate-900" onClick={() => { setSearch(""); setConditions([]); setRangeText(""); captureQuery({ search: "", conditions: [], rangeText: "" }); }}>Clear filters</button>}><div className="grid gap-3 sm:grid-cols-[minmax(0,1fr)_160px_120px]"><input value={search} onChange={(e) => { setSearch(e.target.value); captureQuery({ search: e.target.value }); setPage(0); }} className={field} placeholder="Search values…" /><select className={field} value={searchColumn} onChange={(e) => { setSearchColumn(e.target.value); captureQuery({ searchColumn: e.target.value }); }}><option value="all">All columns</option>{columns.map((c) => <option key={c.index} value={c.index}>{c.letter} · {c.name}</option>)}</select><select className={field} value={searchMode} onChange={(e) => setSearchMode(e.target.value)}><option value="fuzzy">Fuzzy</option><option value="exact">Exact</option></select></div><div className="mt-4 space-y-2">{conditions.map((c) => <div key={c.id} className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)_minmax(0,1fr)_32px] gap-2"><select className={field} value={c.column} onChange={(e) => updateConditions(conditions.map((x) => x.id === c.id ? { ...x, column: e.target.value } : x))}>{columns.map((col) => <option key={col.index} value={col.index}>{col.letter} · {col.name}</option>)}</select><select className={field} value={c.op} onChange={(e) => updateConditions(conditions.map((x) => x.id === c.id ? { ...x, op: e.target.value } : x))}>{OPERATORS.map((op) => <option key={op}>{op}</option>)}</select><input className={field} disabled={["empty", "not empty"].includes(c.op)} value={c.value} placeholder="Value" onChange={(e) => updateConditions(conditions.map((x) => x.id === c.id ? { ...x, value: e.target.value } : x))} /><button aria-label="Remove filter" className="text-slate-400 hover:text-rose-600" onClick={() => updateConditions(conditions.filter((x) => x.id !== c.id))}>×</button></div>)}<button className="text-sm font-medium text-emerald-700 hover:text-emerald-800" onClick={() => updateConditions([...conditions, { id: uid(), column: String(columns[0]?.index ?? 0), op: "contains", value: "" }])}>+ Add condition (AND)</button></div></Panel>
          {formulaCount > 0 && <div className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-xs text-amber-900">{formulaCount.toLocaleString()} formula cells in this sheet. Displaying stored results, not recalculating formulas.{sheet.uncachedCount > 0 && ` ${sheet.uncachedCount.toLocaleString()} formulas have no cached result and may appear blank.`} Cached results can be outdated.</div>}
          <Panel title="Results" action={<span className="text-xs text-slate-500">{selectionResult.rows.length.toLocaleString()} matched</span>}><div className="mb-4 flex flex-wrap items-center gap-2"><input className={`${field} max-w-[260px]`} placeholder="Bin name (optional)" value={binName} onChange={(e) => setBinName(e.target.value)} /><button className={primary} disabled={!selectedRecords.length || !selectionResult.columns.length} onClick={addBin}>+ Save to bin</button><div className="flex-1"/><button className={ghost} disabled={busy || !selectedRecords.length} onClick={() => doExport(currentTable, "csv", "selection")}>CSV</button><button className={ghost} disabled={busy || !selectedRecords.length} onClick={() => doExport(currentTable, "xlsx", "selection")}>Excel</button></div><DataTable columns={selectionResult.columns} records={selectedRecords} sourceRows={selectionResult.rows.map((r) => r.sourceRow)} page={page} onPage={setPage} empty="No rows match your selection." /></Panel>
        </>}</div>
      </div> : <div className="grid items-start gap-5 lg:grid-cols-[300px_minmax(0,1fr)]"><aside className="space-y-5"><Panel title="Saved bins" action={<span className="text-xs text-slate-400">{bins.length}</span>}>{!bins.length ? <Empty>Save a selection from Explore to create a bin.</Empty> : <div className="space-y-2">{bins.map((b) => <div key={b.id} className={`flex items-start gap-2 rounded-xl border p-3 ${activeBin?.id === b.id ? "border-emerald-300 bg-emerald-50/50" : "border-slate-200"}`}><button onClick={() => { setActiveBinId(b.id); setCombined(null); setBinPage(0); }} className="min-w-0 flex-1 text-left"><div className="truncate text-sm font-medium">{b.name}</div><div className="mt-1 text-xs text-slate-500">{b.records.length.toLocaleString()} rows · {b.columns.length} columns</div></button><button title="Delete bin" aria-label={`Delete ${b.name}`} className="text-slate-400 hover:text-rose-600" onClick={() => { setBins((prev) => prev.filter((x) => x.id !== b.id)); if (activeBinId === b.id) setActiveBinId(""); setCombined(null); }}>×</button></div>)}</div>}</Panel></aside><div className="min-w-0 space-y-5">
        <Panel title="Combine bins"><div className="grid gap-3 sm:grid-cols-2"><label className="text-xs font-medium text-slate-600">First bin<select className={`${field} mt-1.5`} value={left?.id || ""} onChange={(e) => { setLeftId(e.target.value); setLeftKey(""); }}>{bins.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></label><label className="text-xs font-medium text-slate-600">Second bin<select className={`${field} mt-1.5`} value={right?.id || ""} onChange={(e) => { setRightId(e.target.value); setRightKey(""); }}>{bins.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></label></div><div className="mt-4 flex gap-2"><button onClick={() => setCombineMode("append")} className={combineMode === "append" ? primary : ghost}>Append rows</button><button onClick={() => setCombineMode("join")} className={combineMode === "join" ? primary : ghost}>Match by column</button></div>{combineMode === "join" && <div className="mt-4 grid gap-3 sm:grid-cols-3"><label className="text-xs font-medium text-slate-600">First key<select className={`${field} mt-1.5`} value={leftKey} onChange={(e) => setLeftKey(e.target.value)}><option value="">Select column</option>{left?.columns.map((c) => <option key={c} value={c}>{c}</option>)}</select></label><label className="text-xs font-medium text-slate-600">Second key<select className={`${field} mt-1.5`} value={rightKey} onChange={(e) => setRightKey(e.target.value)}><option value="">Select column</option>{right?.columns.map((c) => <option key={c} value={c}>{c}</option>)}</select></label><label className="text-xs font-medium text-slate-600">Join type<select className={`${field} mt-1.5`} value={joinHow} onChange={(e) => setJoinHow(e.target.value)}><option value="left">Keep all first-bin rows</option><option value="inner">Only matched rows</option></select></label></div>}<div className="mt-4 flex justify-end"><button className={primary} disabled={bins.length < 2} onClick={createCombined}>Preview combined result</button></div></Panel>
        <Panel title={combined ? "Combined preview" : activeBin ? activeBin.name : "Bin preview"} action={shownTable && <span className="text-xs text-slate-500">{shownTable.records.length.toLocaleString()} rows</span>}>{!shownTable ? <Empty>No bin selected.</Empty> : <><div className="mb-4 flex flex-wrap gap-2">{combined && <button className={primary} onClick={saveCombined}>+ Save combined as bin</button>}<div className="flex-1"/><button className={ghost} disabled={busy} onClick={() => doExport(shownTable, "csv", combined?.name || activeBin?.name)}>CSV</button><button className={ghost} disabled={busy} onClick={() => doExport(shownTable, "xlsx", combined?.name || activeBin?.name)}>Excel</button></div><DataTable columns={shownTable.columns} records={shownTable.records} page={binPage} onPage={setBinPage} /></>}</Panel>
      </div></div>}
      <footer className="mt-8 text-center text-xs text-slate-400">Processed in your browser · No macros executed · Formula results are not recalculated</footer>
    </div>
  </main>;
}
