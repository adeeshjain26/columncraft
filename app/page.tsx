'use client';

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowDownToLine, ArrowRight, Check, CheckCheck, ChevronDown, CircleHelp,
  Columns3, FileSpreadsheet, FileUp, Fingerprint, FolderOpen, LoaderCircle,
  LockKeyhole, Search, ShieldCheck, Sparkles, Table2, Upload, X,
} from 'lucide-react';

type Column = { index: number; label: string; letter: string };
type ParsedSheet = { name: string; rows: unknown[][]; columns: Column[]; headerRow: number };
type OutputFormat = 'xlsx' | 'xls' | 'csv' | 'tsv';
const SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.tsv', '.ods'];
const extensionOf = (name: string) => name.toLowerCase().slice(name.lastIndexOf('.'));
const formatLabel = (extension: string) => ({ '.xlsx': 'Excel workbook', '.xls': 'Legacy Excel workbook', '.xlsm': 'Macro-enabled Excel workbook', '.xlsb': 'Binary Excel workbook', '.csv': 'CSV file', '.tsv': 'TSV file', '.ods': 'OpenDocument spreadsheet' }[extension] ?? 'spreadsheet');

const columnLetter = (n: number) => {
  let result = '';
  let value = n + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
};

const displayCell = (value: unknown) => {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
};

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [activeSheetName, setActiveSheetName] = useState('');
  const [headerRow, setHeaderRow] = useState(1);
  const [selected, setSelected] = useState<number[]>([]);
  const [search, setSearch] = useState('');
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState('');
  const [exporting, setExporting] = useState(false);
  const [exported, setExported] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(true);
  const [outputFormat, setOutputFormat] = useState<OutputFormat>('xlsx');

  const sheetNames = workbook?.SheetNames ?? [];
  const activeSheet = useMemo<ParsedSheet | null>(() => {
    if (!workbook || !activeSheetName || !workbook.Sheets[activeSheetName]) return null;
    const sheet = workbook.Sheets[activeSheetName];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1, defval: '', raw: true, blankrows: true,
    });
    const maxColumns = Math.max(0, ...rows.map((row) => row.length));
    const actualHeaderRow = Math.max(1, Math.min(headerRow, Math.max(1, rows.length)));
    const headers = rows[actualHeaderRow - 1] ?? [];
    const columns = Array.from({ length: maxColumns }, (_, index) => {
      const value = displayCell(headers[index]).trim();
      return { index, label: value || `Column ${columnLetter(index)}`, letter: columnLetter(index) };
    });
    return { name: activeSheetName, rows, columns, headerRow: actualHeaderRow };
  }, [workbook, activeSheetName, headerRow]);

  const visibleColumns = useMemo(() => {
    if (!activeSheet) return [];
    const needle = search.trim().toLowerCase();
    return activeSheet.columns.filter((column) =>
      !needle || column.label.toLowerCase().includes(needle) || column.letter.toLowerCase().includes(needle),
    );
  }, [activeSheet, search]);

  const dataRows = activeSheet ? activeSheet.rows.slice(activeSheet.headerRow).filter((row) =>
    row.some((cell) => displayCell(cell).trim() !== ''),
  ) : [];
  const allVisibleSelected = visibleColumns.length > 0 && visibleColumns.every((c) => selected.includes(c.index));
  const selectedColumns = activeSheet?.columns.filter((c) => selected.includes(c.index)) ?? [];
  const previewRows = dataRows.slice(0, 5);

  async function loadFile(file?: File) {
    setError(''); setExported(false);
    if (!file) return;
    const extension = extensionOf(file.name);
    if (!SUPPORTED_EXTENSIONS.includes(extension)) {
      setError('Unsupported file type. Choose XLSX, XLS, XLSM, XLSB, CSV, TSV, or ODS.');
      return;
    }
    if (file.size > 100 * 1024 * 1024) {
      setError('This file is over 100 MB. Try a smaller workbook for a smoother experience.');
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      // SheetJS parses each supported format locally in the browser; no file is uploaded.
      const parsed = XLSX.read(buffer, { type: 'array', cellDates: true, dense: false, raw: false });
      if (!parsed.SheetNames.length) throw new Error('No worksheets were found in this workbook.');
      const firstSheet = parsed.SheetNames[0];
      const firstRows = XLSX.utils.sheet_to_json<unknown[]>(parsed.Sheets[firstSheet], { header: 1, defval: '', blankrows: true });
      if (!firstRows.length) throw new Error('This workbook appears to be empty.');
      setWorkbook(parsed); setFileName(file.name); setFileSize(file.size);
      setActiveSheetName(firstSheet); setHeaderRow(1); setSearch(''); setExported(false);
      const width = Math.max(0, ...firstRows.map((row) => row.length));
      setSelected(Array.from({ length: width }, (_, index) => index));
    } catch (e) {
      setWorkbook(null);
      setError(e instanceof Error ? `Couldn't read this workbook: ${e.message}` : 'Could not read this workbook. Please check the file and try again.');
    }
  }

  function onInputChange(event: ChangeEvent<HTMLInputElement>) {
    void loadFile(event.target.files?.[0]);
    event.target.value = '';
  }
  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault(); setDragging(false); void loadFile(event.dataTransfer.files?.[0]);
  }
  function switchSheet(name: string) {
    if (!workbook) return;
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], { header: 1, defval: '', blankrows: true });
    setActiveSheetName(name); setHeaderRow(1); setSearch(''); setExported(false);
    setSelected(Array.from({ length: Math.max(0, ...rows.map((row) => row.length)) }, (_, index) => index));
  }
  function toggleColumn(index: number) {
    setSelected((current) => current.includes(index) ? current.filter((item) => item !== index) : [...current, index].sort((a, b) => a - b));
    setExported(false);
  }
  function toggleVisible() {
    const visibleIndexes = visibleColumns.map((column) => column.index);
    setSelected((current) => allVisibleSelected
      ? current.filter((index) => !visibleIndexes.includes(index))
      : Array.from(new Set([...current, ...visibleIndexes])).sort((a, b) => a - b));
    setExported(false);
  }

  function exportWorkbook() {
    if (!activeSheet || !workbook || !selectedColumns.length) return;
    setExporting(true); setError('');
    try {
      const outputRows = activeSheet.rows.slice(0, activeSheet.headerRow).map((row) => selectedColumns.map((c) => row[c.index] ?? ''));
      const header = selectedColumns.map((c) => {
        const original = displayCell(activeSheet.rows[activeSheet.headerRow - 1]?.[c.index]).trim();
        return original || c.label;
      });
      outputRows[activeSheet.headerRow - 1] = header;
      for (const row of dataRows) outputRows.push(selectedColumns.map((c) => row[c.index] ?? ''));
      const outputSheet = XLSX.utils.aoa_to_sheet(outputRows);
      const outputBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outputBook, outputSheet, activeSheet.name.slice(0, 31) || 'Extracted');
      const base = fileName.replace(/\.[^.]+$/, '');
      const safeSheet = activeSheet.name.replace(/[\\/:*?"<>|]/g, '_');
      const outputName = `${base}_${safeSheet}_extracted.${outputFormat}`;
      if (outputFormat === 'tsv') {
        const tsv = XLSX.utils.sheet_to_csv(outputSheet, { FS: '\t' });
        const blob = new Blob(['\uFEFF', tsv], { type: 'text/tab-separated-values;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a'); link.href = url; link.download = outputName; link.click();
        URL.revokeObjectURL(url);
      } else {
        XLSX.writeFile(outputBook, outputName, { bookType: outputFormat === 'xls' ? 'biff8' : outputFormat, compression: true });
      }
      setExported(true);
    } catch {
      setError('Export failed. Try a smaller workbook or a different worksheet.');
    } finally {
      setExporting(false);
    }
  }

  const formatSize = (bytes: number) => bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="ColumnCraft home">
          <span className="brand-mark"><Columns3 size={21} strokeWidth={2.4} /></span>
          <span>column<span className="brand-accent">craft</span><small>EXCEL TOOLKIT</small></span>
        </a>
        <div className="topbar-right"><span className="privacy-pill"><LockKeyhole size={13} /> 100% local processing</span><a href="#how-it-works" className="help-link"><CircleHelp size={16} /> How it works</a></div>
      </header>

      <section className="hero" id="top">
        <div className="hero-copy">
          <div className="eyebrow"><Sparkles size={14} /> LESS SPREADSHEET CLUTTER</div>
          <h1>Keep the columns.<br /><span>Lose the noise.</span></h1>
          <p className="hero-description">Extract exactly the data you need from Excel, CSV, TSV, and OpenDocument spreadsheets. Pick a sheet, choose your columns, and download a clean file in seconds.</p>
          <div className="trust-row"><span><ShieldCheck size={16} /> Your files stay on your device</span><span className="trust-divider" /><span>No sign-up. No uploads.</span></div>
        </div>
        <div className="hero-art" aria-hidden="true">
          <div className="art-orbit orbit-one" /><div className="art-orbit orbit-two" />
          <div className="floating-sheet sheet-back"><div /><div /><div /></div>
          <div className="floating-sheet sheet-front"><div className="sheet-title"><FileSpreadsheet size={20} /><span>sales_data.xlsx</span><span className="sheet-dot" /></div><div className="sheet-grid"><span /><span /><span /><span /><span /><span /><span /><span /><span /><span /><span /><span /></div><div className="sheet-selected"><Check size={14} /> 3 columns selected</div></div>
          <div className="art-check"><Check size={19} /></div><div className="art-spark spark-a">✳</div><div className="art-spark spark-b">✦</div>
        </div>
      </section>

      <section className="workspace" aria-label="Excel column extractor">
        <div className="section-heading"><div><span className="section-kicker">YOUR WORKSPACE</span><h2>{workbook ? 'Customize your extract' : 'Start with a workbook'}</h2></div>{workbook && <button className="text-button" onClick={() => { setWorkbook(null); setFileName(''); setSelected([]); setError(''); setExported(false); }}><X size={15} /> Start over</button>}</div>

        {!workbook ? (
          <div className={`dropzone ${dragging ? 'is-dragging' : ''}`} onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop}>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv,.tsv,.ods" onChange={onInputChange} hidden />
            <div className="upload-icon"><FileUp size={27} /></div>
            <h3>Drop a spreadsheet file here</h3><p>or choose a file from your device</p>
            <button className="primary-button upload-button" onClick={() => inputRef.current?.click()}><FolderOpen size={17} /> Choose spreadsheet <ArrowRight size={16} /></button>
            <div className="upload-foot"><span><Check size={13} /> XLSX, XLS, XLSM</span><span><Check size={13} /> CSV, TSV, ODS</span><span><Check size={13} /> Up to 100 MB</span><span><Check size={13} /> Private by design</span></div>
          </div>
        ) : (
          <div className="editor-layout">
            <div className="left-panel">
              <div className="file-card"><div className="file-icon"><FileSpreadsheet size={23} /></div><div className="file-info"><strong title={fileName}>{fileName}</strong><span>{formatSize(fileSize)} <i /> {formatLabel(extensionOf(fileName))} <i /> {sheetNames.length} {sheetNames.length === 1 ? 'worksheet' : 'worksheets'}</span></div><span className="ready-badge"><Check size={13} /> Ready</span><button className="icon-button" aria-label="Choose another file" onClick={() => inputRef.current?.click()}><Upload size={16} /></button><input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv,.tsv,.ods" onChange={onInputChange} hidden /></div>
              <div className="panel-card sheet-panel"><div className="panel-title"><span className="step-number">1</span><div><h3>Choose worksheet</h3><p>Select which sheet to work with</p></div></div><label className="field-label" htmlFor="sheet-select">WORKSHEET</label><div className="select-wrap"><Table2 size={16} /><select id="sheet-select" value={activeSheetName} onChange={(e) => switchSheet(e.target.value)}>{sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}</select><ChevronDown size={16} className="select-chevron" /></div>
                <div className="header-row-control"><label htmlFor="header-row">Header row</label><div className="number-control"><button aria-label="Decrease header row" onClick={() => setHeaderRow((n) => Math.max(1, n - 1))} disabled={headerRow <= 1}>−</button><input id="header-row" type="number" min={1} max={Math.max(1, activeSheet?.rows.length ?? 1)} value={headerRow} onChange={(e) => { const n = Number(e.target.value); if (n >= 1) { setHeaderRow(Math.min(n, Math.max(1, activeSheet?.rows.length ?? 1))); setSelected([]); setExported(false); } }} /><button aria-label="Increase header row" onClick={() => setHeaderRow((n) => Math.min(n + 1, Math.max(1, activeSheet?.rows.length ?? 1)))} disabled={headerRow >= (activeSheet?.rows.length ?? 1)}>+</button></div></div>
              </div>
              <div className="panel-card columns-panel"><div className="panel-title"><span className="step-number">2</span><div><h3>Choose your columns</h3><p>Keep only the fields you need</p></div></div><div className="column-stats"><span><b>{selected.length}</b> of {activeSheet?.columns.length ?? 0} selected</span><span>{dataRows.length.toLocaleString()} data rows</span></div><div className="search-box"><Search size={16} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search column names..." aria-label="Search column names" />{search && <button aria-label="Clear search" onClick={() => setSearch('')}><X size={14} /></button>}</div><button className="select-all-row" onClick={toggleVisible}><span className={`fake-checkbox ${allVisibleSelected ? 'checked' : ''}`}>{allVisibleSelected && <Check size={13} />}</span><span>{allVisibleSelected ? 'Deselect visible' : 'Select visible'}</span><span className="select-count">{visibleColumns.length} fields</span></button><div className="column-list">{visibleColumns.map((column) => { const checked = selected.includes(column.index); return <label className={`column-option ${checked ? 'column-checked' : ''}`} key={column.index}><input type="checkbox" checked={checked} onChange={() => toggleColumn(column.index)} /><span className="fake-checkbox">{checked && <Check size={13} />}</span><span className="column-letter">{column.letter}</span><span className="column-name" title={column.label}>{column.label}</span>{checked && <CheckCheck size={15} className="column-tick" />}</label>; })}{visibleColumns.length === 0 && <div className="no-results">No columns match “{search}”</div>}</div><div className="column-footer"><span><Fingerprint size={14} /> Original file stays untouched</span><button className="text-button compact" onClick={() => { setSelected([]); setExported(false); }}>Clear selection</button></div></div>
            </div>

            <div className="right-panel"><div className="preview-card"><div className="preview-head"><div><span className="step-number">3</span><div><h3>Preview your extract</h3><p>See what your new file will contain</p></div></div><button className="preview-toggle" aria-label={previewOpen ? 'Collapse preview' : 'Expand preview'} onClick={() => setPreviewOpen((v) => !v)}><ChevronDown size={18} className={previewOpen ? '' : 'rotated'} /></button></div>{previewOpen && <><div className="preview-summary"><div><span className="summary-label">COLUMNS</span><strong>{selectedColumns.length}</strong></div><div><span className="summary-label">DATA ROWS</span><strong>{dataRows.length.toLocaleString()}</strong></div><div><span className="summary-label">WORKSHEET</span><strong className="summary-sheet" title={activeSheetName}>{activeSheetName}</strong></div></div><div className="table-wrap">{selectedColumns.length > 0 ? <table><thead><tr><th className="row-index">#</th>{selectedColumns.map((c) => <th key={c.index} title={c.label}>{c.label}</th>)}</tr></thead><tbody>{previewRows.map((row, rowIndex) => <tr key={rowIndex}><td className="row-index">{rowIndex + 1}</td>{selectedColumns.map((c) => <td key={c.index} title={displayCell(row[c.index])}>{displayCell(row[c.index]) || <span className="empty-cell">—</span>}</td>)}</tr>)}</tbody></table> : <div className="empty-preview"><Columns3 size={23} /><span>Select at least one column to preview your extract.</span></div>}</div><div className="preview-foot"><span>Showing {Math.min(previewRows.length, 5)} of {dataRows.length.toLocaleString()} rows</span><span className="preview-safe"><LockKeyhole size={12} /> Private preview</span></div></>}</div>
              <div className="export-card"><div className="export-icon"><ArrowDownToLine size={20} /></div><div className="export-copy"><h3>Ready to export?</h3><p>Your new file will include only the selected columns.</p><label className="field-label output-format-label" htmlFor="output-format">OUTPUT FORMAT</label><select id="output-format" className="output-format-select" value={outputFormat} onChange={(e) => setOutputFormat(e.target.value as OutputFormat)}><option value="xlsx">Excel (.xlsx)</option><option value="xls">Excel 97–2003 (.xls)</option><option value="csv">CSV (.csv)</option><option value="tsv">Tab-separated (.tsv)</option></select></div><button className="primary-button export-button" disabled={!selectedColumns.length || exporting} onClick={exportWorkbook}>{exporting ? <LoaderCircle size={17} className="spin" /> : <ArrowDownToLine size={17} />}{exporting ? 'Preparing...' : 'Download Excel'} </button>{exported && <div className="success-message"><Check size={15} /> Export started. Check your downloads.</div>}</div>
              <div className="privacy-note"><ShieldCheck size={18} /><p><strong>Your data stays yours.</strong> This tool runs in your browser. The workbook is never uploaded to a server or stored by us.</p></div>
            </div>
          </div>
        )}
        {error && <div className="error-message" role="alert"><X size={16} /> {error}</div>}
      </section>

      <section className="how-section" id="how-it-works"><div className="how-intro"><span className="section-kicker">SIMPLE BY DESIGN</span><h2>From messy to ready<br />in three steps.</h2><p>No complicated spreadsheet tools. Just the columns you need, in a clean new file.</p></div><div className="how-steps"><div className="how-step"><div className="how-icon"><Upload size={20} /></div><span className="how-number">01 / IMPORT</span><h3>Bring your file</h3><p>Choose an Excel, CSV, TSV, or ODS file from your device. Nothing gets uploaded.</p></div><div className="how-step"><div className="how-icon"><Columns3 size={20} /></div><span className="how-number">02 / CUSTOMIZE</span><h3>Pick your columns</h3><p>Select a worksheet and check only the fields you want to keep.</p></div><div className="how-step"><div className="how-icon"><ArrowDownToLine size={20} /></div><span className="how-number">03 / EXPORT</span><h3>Download and go</h3><p>Download a new spreadsheet with your chosen columns and data rows.</p></div></div></section>
      <footer className="footer"><a className="brand footer-brand" href="#top"><span className="brand-mark"><Columns3 size={17} /></span><span>column<span className="brand-accent">craft</span></span></a><span>Made for cleaner spreadsheets.</span><span className="footer-private"><LockKeyhole size={13} /> Local processing by design</span></footer>
    </main>
  );
}
