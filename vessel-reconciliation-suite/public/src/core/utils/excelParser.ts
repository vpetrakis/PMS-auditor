// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Omni-Parser
//
// ARCHITECTURE — identical to Poseidon Titan's proven approach:
//   1. Read the full sheet as a raw grid (no assumed headers)
//   2. Scan for the header row by looking for sentinel keywords
//   3. Forward-fill the top header row (handles merged cells)
//   4. Combine top + bottom header into a semantic column map
//   5. Extract data rows below the header pair
//
// Two public functions:
//   parsePMSFile(file)  → ComponentData[]
//   parseLogFile(file)  → DailyLog[]
// ═══════════════════════════════════════════════════════════════════════════

import * as XLSX from 'xlsx';
import type { DailyLog, ComponentData } from '../schemas/tec001.schema';
import type { SystemLabel } from '../../types/vessel';

// ─── Raw cell value type ──────────────────────────────────────────────────────
type CV  = string | number | boolean | Date | null | undefined;
type Row = CV[];

// ─────────────────────────────────────────────────────────────────────────────
// SAFE NUMBER — mirrors Poseidon Titan's _sn()
// Strips any non-numeric characters, returns NaN on garbage.
// ─────────────────────────────────────────────────────────────────────────────
function sn(val: CV): number {
  if (val === null || val === undefined) return NaN;
  if (typeof val === 'number') return isNaN(val) ? NaN : val;
  if (typeof val === 'boolean') return NaN;
  if (val instanceof Date) return NaN;
  const s = String(val).trim().toUpperCase();
  if (['NIL','N/A','NA','XXX','NONE','UNKNOWN','-','X','','NULL','BLANK'].includes(s)) return NaN;
  const cleaned = s.replace(/[^\d.\-]/g, '');
  if (!cleaned || cleaned === '.' || cleaned === '-' || cleaned === '-.') return NaN;
  const n = parseFloat(cleaned);
  return isNaN(n) ? NaN : n;
}

function sn0(val: CV): number {
  const v = sn(val);
  return isNaN(v) ? 0 : v;
}

// ─────────────────────────────────────────────────────────────────────────────
// SAFE DATE — handles: Date objects, XLSX serial numbers,
//             ISO strings, DD/MM/YYYY, DD MMM YYYY, etc.
// ─────────────────────────────────────────────────────────────────────────────
function safeDate(val: CV): Date | null {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  // XLSX serial number
  if (typeof val === 'number') {
    try {
      const p = XLSX.SSF.parse_date_code(val);
      if (p) return new Date(p.y, p.m - 1, p.d);
    } catch { /* fall through */ }
    return null;
  }

  const raw = String(val).trim();
  if (!raw || ['-','n/a','nil','none','null'].includes(raw.toLowerCase())) return null;

  // Typo correction (mirrors Poseidon: re.sub r'20224' → '2024')
  const fixed = raw
    .replace(/20224/g, '2024')
    .replace(/20023/g, '2023')
    .replace(/20225/g, '2025');

  // "15 Jan 2024" or "15 Jan. 2024"
  const dmyText = fixed.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\.?\s+(\d{4})$/);
  if (dmyText) {
    const d = new Date(`${dmyText[2]} ${dmyText[1]}, ${dmyText[3]}`);
    if (!isNaN(d.getTime())) return d;
  }

  // Native parse (handles ISO, "Jan 15 2024", etc.)
  const native = new Date(fixed);
  if (!isNaN(native.getTime())) return native;

  // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
  const dmy = fixed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (dmy) {
    const yr = dmy[3].length === 2 ? (parseInt(dmy[3]) > 50 ? 1900 : 2000) + parseInt(dmy[3]) : parseInt(dmy[3]);
    const d  = new Date(yr, parseInt(dmy[2]) - 1, parseInt(dmy[1]));
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// NORMALISE — uppercase + collapse whitespace
// ─────────────────────────────────────────────────────────────────────────────
function norm(v: CV): string {
  if (v === null || v === undefined) return '';
  return String(v).trim().toUpperCase().replace(/\s+/g, ' ');
}

// ─────────────────────────────────────────────────────────────────────────────
// FORWARD-FILL — mirrors Poseidon's pd.Series.ffill()
// Propagates the last non-empty cell to the right.
// This is essential for merged-cell header rows.
// ─────────────────────────────────────────────────────────────────────────────
function fwdFill(row: CV[]): string[] {
  let last = '';
  return row.map(v => {
    const s = norm(v);
    if (s) last = s;
    return last;
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SHEET → RAW GRID
// Reads every cell; Date-formatted cells come back as JS Date objects.
// ─────────────────────────────────────────────────────────────────────────────
function sheetToGrid(ws: XLSX.WorkSheet): Row[] {
  if (!ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const grid: Row[] = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: Row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) { row.push(null); continue; }

      if (cell.t === 'd' && cell.v instanceof Date) { row.push(cell.v); continue; }
      // XLSX serial with date format
      if (cell.t === 'n' && typeof cell.z === 'string' && /[\/\-mdyMDY]/.test(cell.z)) {
        const d = safeDate(cell.v);
        row.push(d ?? cell.v);
        continue;
      }
      row.push(cell.v ?? null);
    }
    grid.push(row);
  }
  return grid;
}

async function readWB(file: File): Promise<XLSX.WorkBook> {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload  = e => {
      try { res(XLSX.read(e.target!.result as ArrayBuffer, { type: 'array', cellDates: true, cellNF: true, cellText: false })); }
      catch (err) { rej(new Error(`Cannot read "${file.name}": ${err}`)); }
    };
    reader.onerror = () => rej(new Error(`File read error: ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN MAPPER — LOG FILES
// Mirrors _map_columns() from Poseidon Titan.
// c1 = forward-filled top header, c2 = bottom sub-header, combined = c1 + c2
// ─────────────────────────────────────────────────────────────────────────────
interface LogColMap {
  date?: number; me?: number; dg1?: number; dg2?: number; dg3?: number;
}

function mapLogCols(topFilled: string[], bottom: string[], n: number): LogColMap {
  const m: LogColMap = {};

  for (let j = 0; j < n; j++) {
    const c1 = topFilled[j] ?? '';
    const c2 = bottom[j] ?? '';
    const cb = `${c1} ${c2}`.trim();

    if (m.date === undefined && (cb.includes('DATE') || cb.includes('DAY')))
      m.date = j;

    if (m.me === undefined && (
      cb.includes('MAIN ENGINE') || cb.includes('MAIN ENG') ||
      cb.includes('M/E') || cb.includes('M.E.') ||
      c1.includes('MAIN') || c2 === 'ME' || c2 === 'M/E' || c2 === 'M.E.'
    )) m.me = j;

    if (m.dg1 === undefined && (
      cb.includes('DG1') || cb.includes('D/G 1') || cb.includes('D.G.1') ||
      cb.includes('DG NO.1') || cb.includes('GEN 1') || cb.includes('G/E 1') ||
      cb.includes('AUX 1') || cb.includes('GEN.1') ||
      // Pattern: "D/G" in top + "No.1" or "1" in bottom
      (c1.includes('D/G') && (c2.includes('1') || c2.includes('NO.1')))
    )) m.dg1 = j;

    if (m.dg2 === undefined && (
      cb.includes('DG2') || cb.includes('D/G 2') || cb.includes('D.G.2') ||
      cb.includes('DG NO.2') || cb.includes('GEN 2') || cb.includes('G/E 2') ||
      cb.includes('AUX 2') || cb.includes('GEN.2') ||
      (c1.includes('D/G') && (c2.includes('2') || c2.includes('NO.2')))
    )) m.dg2 = j;

    if (m.dg3 === undefined && (
      cb.includes('DG3') || cb.includes('D/G 3') || cb.includes('D.G.3') ||
      cb.includes('DG NO.3') || cb.includes('GEN 3') || cb.includes('G/E 3') ||
      cb.includes('AUX 3') || cb.includes('GEN.3') ||
      (c1.includes('D/G') && (c2.includes('3') || c2.includes('NO.3')))
    )) m.dg3 = j;
  }

  return m;
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN MAPPER — PMS FILES
// ─────────────────────────────────────────────────────────────────────────────
interface PmsColMap {
  component?: number; ohDate?: number; hours?: number; system?: number; interval?: number;
}

function mapPmsCols(topFilled: string[], bottom: string[], n: number): PmsColMap {
  const m: PmsColMap = {};

  for (let j = 0; j < n; j++) {
    const c1 = topFilled[j] ?? '';
    const c2 = bottom[j] ?? '';
    const cb = `${c1} ${c2}`.trim();

    if (m.component === undefined && (
      cb.includes('COMPONENT') || cb.includes('EQUIPMENT') || cb.includes('DESCRIPTION') ||
      cb.includes('ITEM') || cb.includes('JOB DESCRIPTION') || cb.includes('TASK') ||
      cb.includes('NAME') || c2.includes('DESCRIPTION') || c2.includes('COMPONENT')
    )) m.component = j;

    if (m.ohDate === undefined && (
      cb.includes('OVERHAUL') || cb.includes('LAST OH') || cb.includes('O/H DATE') ||
      cb.includes('LAST O/H') || cb.includes('INSP DATE') || cb.includes('LAST INSP') ||
      cb.includes('LAST INSPECTION') || cb.includes('COMPLETED') ||
      (cb.includes('DATE') && m.component !== undefined && m.ohDate === undefined)
    )) m.ohDate = j;

    if (m.hours === undefined && (
      cb.includes('RUNNING HOURS') || cb.includes('R/H') || cb.includes('HRS SINCE') ||
      cb.includes('CURRENT HRS') || cb.includes('CLAIMED') || cb.includes('HOURS SINCE') ||
      cb.includes('RUN HRS') || cb.includes('CURRENT RUNNING') ||
      (cb.includes('HOURS') && j > 0 && m.ohDate !== undefined)
    )) m.hours = j;

    if (m.system === undefined && (
      cb.includes('SYSTEM') || cb.includes('MACHINERY') || cb.includes('DEPT') ||
      cb.includes('CATEGORY') || cb.includes('PARENT')
    )) m.system = j;

    if (m.interval === undefined && (
      cb.includes('INTERVAL') || cb.includes('DUE HOURS') || cb.includes('NEXT OH')
    )) m.interval = j;
  }

  return m;
}

// ─────────────────────────────────────────────────────────────────────────────
// SYSTEM INFERENCE
// ─────────────────────────────────────────────────────────────────────────────
function inferSystem(name: string, sysCell?: CV): SystemLabel {
  const s = `${norm(name)} ${norm(sysCell)}`;
  if (s.includes('DG3') || s.includes('D/G 3') || s.includes('GEN 3') || s.includes('G/E 3') || s.includes('AUX 3')) return 'DG3';
  if (s.includes('DG2') || s.includes('D/G 2') || s.includes('GEN 2') || s.includes('G/E 2') || s.includes('AUX 2')) return 'DG2';
  if (s.includes('DG1') || s.includes('D/G 1') || s.includes('GEN 1') || s.includes('G/E 1') || s.includes('AUX 1')) return 'DG1';
  if (s.includes('DG') || s.includes('D/G') || s.includes('DIESEL GEN') || s.includes('GENERATOR') || s.includes('GENSET')) return 'DG1';
  return 'MAIN_ENGINE';
}

// ─────────────────────────────────────────────────────────────────────────────
// HEADER SCANNER — finds the row index containing sentinel keywords
// Mirrors Poseidon's "for i in range(min(150, len(df_raw)))" loop
// ─────────────────────────────────────────────────────────────────────────────
function findHeaderRow(
  grid: Row[],
  sentinelGroups: string[][]  // every group must match at least one column
): number {
  for (let i = 0; i < Math.min(80, grid.length); i++) {
    const cells = grid[i].map(norm);
    const allMatch = sentinelGroups.every(group =>
      group.some(kw => cells.some(c => c.includes(kw)))
    );
    if (allMatch) return i;
  }
  return -1;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC: PARSE LOG FILE
// ─────────────────────────────────────────────────────────────────────────────
export async function parseLogFile(file: File): Promise<DailyLog[]> {
  const wb   = await readWB(file);
  const logs: DailyLog[] = [];
  const errs: string[]   = [];

  for (const sheetName of wb.SheetNames) {
    const ws   = wb.Sheets[sheetName];
    if (!ws['!ref']) continue;
    const grid = sheetToGrid(ws);

    // Must have DATE (or DAY) AND a MAIN ENGINE variant
    const headerIdx = findHeaderRow(grid, [
      ['DATE', 'DAY'],
      ['MAIN', 'M/E', 'M.E.', 'MAIN ENGINE', 'ENGINE'],
    ]);
    if (headerIdx < 0) continue;

    // Poseidon approach: top row forward-filled + bottom sub-header
    const topFilled = fwdFill(grid[headerIdx]);
    const bottom    = (grid[headerIdx + 1] ?? []).map(norm);

    const m = mapLogCols(topFilled, bottom, Math.max(topFilled.length, bottom.length));
    if (m.date === undefined || m.me === undefined) continue;

    // Data starts after the header pair
    const dataStart = headerIdx + (bottom.some(c => c && !['DATE','DAY','MAIN','M/E','ENGINE'].includes(c)) ? 2 : 1);

    for (let r = dataStart; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row.every(v => v === null || v === undefined || v === '')) continue;

      const date = safeDate(row[m.date]);
      if (!date || date.getFullYear() < 1980 || date.getFullYear() > 2100) continue;

      const log: DailyLog = {
        date,
        mainEngineHours: sn0(row[m.me]),
        dg1Hours:  m.dg1 !== undefined ? sn0(row[m.dg1]) : 0,
        dg2Hours:  m.dg2 !== undefined ? sn0(row[m.dg2]) : 0,
        dg3Hours:  m.dg3 !== undefined ? sn0(row[m.dg3]) : 0,
      };
      logs.push(log);
    }

    if (logs.length > 0) break; // Found data — stop scanning sheets
  }

  if (logs.length === 0) {
    const detail = errs.slice(0, 3).join(' | ');
    throw new Error(
      `Log parse failed for "${file.name}". No sheet contained a recognisable DATE + MAIN ENGINE header. ` +
      (detail ? `Details: ${detail}` : 'Verify the file has a row with "Date"/"Day" and "Main Engine"/"M/E" columns.')
    );
  }

  console.info(`[LOG] ${file.name}: ${logs.length} daily entries parsed.`);
  return logs;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC: PARSE PMS FILE
// ─────────────────────────────────────────────────────────────────────────────
export async function parsePMSFile(file: File): Promise<ComponentData[]> {
  const wb    = await readWB(file);
  const comps: ComponentData[] = [];
  const errs:  string[]        = [];

  for (const sheetName of wb.SheetNames) {
    const ws   = wb.Sheets[sheetName];
    if (!ws['!ref']) continue;
    const grid = sheetToGrid(ws);

    // PMS header must have some kind of component/item column + a date column
    const headerIdx = findHeaderRow(grid, [
      ['COMPONENT', 'EQUIPMENT', 'DESCRIPTION', 'ITEM', 'NAME', 'JOB'],
      ['DATE', 'OVERHAUL', 'OH', 'INSP', 'O/H', 'LAST'],
    ]);

    let hIdx = headerIdx;

    // Hard fallback for TEC-001 standard layout (no semantic headers in row)
    if (hIdx < 0 && grid.length > 10 && (grid[0]?.length ?? 0) >= 8) {
      hIdx = 7; // TEC-001 data typically starts at row 8
    }
    if (hIdx < 0) continue;

    const topFilled = fwdFill(grid[hIdx]);
    const bottom    = (grid[hIdx + 1] ?? []).map(norm);

    let m = mapPmsCols(topFilled, bottom, Math.max(topFilled.length, bottom.length));

    // TEC-001 column fallback (col 1 = name, col 5 = date, col 7 = hours)
    if (m.component === undefined) m.component = 1;
    if (m.ohDate    === undefined) m.ohDate    = 5;
    if (m.hours     === undefined) m.hours     = 7;

    const dataStart = hIdx + (bottom.some(c => c.length > 0) ? 2 : 1);

    for (let r = dataStart; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row.every(v => v === null || v === undefined || v === '')) continue;

      const rawName = row[m.component!];
      if (!rawName) continue;
      const compName = String(rawName).trim();
      if (!compName || ['NAN','NULL','NONE','N/A','-'].includes(compName.toUpperCase())) continue;

      const ohDate = safeDate(row[m.ohDate!]);
      if (!ohDate) { errs.push(`Row ${r}: no valid overhaul date for "${compName}"`); continue; }
      if (ohDate.getFullYear() < 1980 || ohDate > new Date(Date.now() + 365*86_400_000)) continue;

      const legacyH     = sn0(m.hours !== undefined ? row[m.hours] : null);
      const parentSystem = inferSystem(compName, m.system !== undefined ? row[m.system] : undefined);

      comps.push({ componentName: compName, lastOverhaulDate: ohDate, legacyClaimedHours: legacyH, parentSystem });
    }

    if (comps.length > 0) break;
  }

  if (comps.length === 0) {
    throw new Error(
      `PMS parse failed for "${file.name}". ` +
      (errs.length ? `Errors: ${errs.slice(0,5).join(' | ')}` : 'No rows contained a valid component name + overhaul date pair.')
    );
  }

  console.info(`[PMS] ${file.name}: ${comps.length} components parsed. Skipped ${errs.length} rows.`);
  return comps;
}
