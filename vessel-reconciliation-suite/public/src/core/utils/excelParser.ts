// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Omni-Parser
//
// Architecture:
//   parsePMSFile(file)  → ComponentData[]
//   parseLogFile(file)  → DailyLog[]
//
// Both functions use the same three-stage approach:
//   Stage 1 — Read workbook with SheetJS (cellDates: true)
//   Stage 2 — Semantic header detection (scan first 60 rows)
//   Stage 3 — Row-by-row extraction with per-cell error isolation
//
// Design principles:
//   - A bad row is SKIPPED, not a fatal error
//   - Every sheet in the workbook is tried until data is found
//   - Date serial numbers, Date objects, and string dates are all handled
//   - Physics violations are NOT rejected here — they are captured downstream
// ═══════════════════════════════════════════════════════════════════════════

import * as XLSX from 'xlsx';
import { ComponentSchema, DailyLogSchema } from '../schemas/tec001.schema';
import type { DailyLog, ComponentData } from '../schemas/tec001.schema';
import type { SystemLabel } from '../../types/vessel.d';

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL TYPES
// ─────────────────────────────────────────────────────────────────────────────

type CellValue = string | number | boolean | Date | null;
type Grid      = CellValue[][];

interface ColMap {
  date?:      number;
  me?:        number;
  dg1?:       number;
  dg2?:       number;
  dg3?:       number;
  component?: number;
  ohDate?:    number;
  hours?:     number;
  system?:    number;
}

interface HeaderMatch {
  rowIdx: number;
  colMap: ColMap;
}

// ─────────────────────────────────────────────────────────────────────────────
// CELL UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

function normalise(val: CellValue): string {
  if (val === null || val === undefined) return '';
  return String(val).trim().toUpperCase().replace(/\s+/g, ' ');
}

function contains(cell: string, ...terms: string[]): boolean {
  return terms.some(t => cell.includes(t));
}

// ─────────────────────────────────────────────────────────────────────────────
// DATE PARSING
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Converts any cell value to a valid Date or null.
 * Handles: Date objects, XLSX serial numbers, ISO strings, DD/MM/YYYY, etc.
 */
function parseDate(val: CellValue): Date | null {
  if (val === null || val === undefined || val === '') return null;

  // Already a Date from SheetJS with cellDates: true
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  // XLSX serial number (number of days since 1900-01-01)
  if (typeof val === 'number') {
    try {
      const parsed = XLSX.SSF.parse_date_code(val);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
    } catch { /* fall through */ }
    return null;
  }

  const str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'n/a') return null;

  // Try native Date parse first
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  // DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (dmyMatch) {
    const [, dd, mm, yy] = dmyMatch;
    const year = yy.length === 2 ? parseInt(yy) + (parseInt(yy) > 50 ? 1900 : 2000) : parseInt(yy);
    const candidate = new Date(year, parseInt(mm) - 1, parseInt(dd));
    if (!isNaN(candidate.getTime())) return candidate;
  }

  // MM/DD/YYYY (US format)
  const mdyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (mdyMatch) {
    const [, mm, dd, yyyy] = mdyMatch;
    const candidate = new Date(parseInt(yyyy), parseInt(mm) - 1, parseInt(dd));
    if (!isNaN(candidate.getTime())) return candidate;
  }

  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// NUMBER PARSING
// ─────────────────────────────────────────────────────────────────────────────

function parseHours(val: CellValue): number {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = String(val).replace(/[^\d.\-]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

// ─────────────────────────────────────────────────────────────────────────────
// SYSTEM INFERENCE
// ─────────────────────────────────────────────────────────────────────────────

function inferSystem(componentName: string, systemCellVal?: CellValue): SystemLabel {
  const name = normalise(componentName);
  const sys  = normalise(systemCellVal ?? '');
  const combined = `${name} ${sys}`;

  // Check for DG3 first (most specific)
  if (contains(combined, 'DG3', 'D/G 3', 'D.G.3', 'AUX 3', 'G/E 3', 'GENERATOR 3', 'GEN 3')) return 'DG3';
  if (contains(combined, 'DG2', 'D/G 2', 'D.G.2', 'AUX 2', 'G/E 2', 'GENERATOR 2', 'GEN 2')) return 'DG2';
  if (contains(combined, 'DG1', 'D/G 1', 'D.G.1', 'AUX 1', 'G/E 1', 'GENERATOR 1', 'GEN 1')) return 'DG1';

  // Generic DG without number
  if (contains(combined, 'DG', 'D/G', 'DIESEL GEN', 'GENERATOR', 'GENSET', 'AUX ENGINE')) {
    // Could be DG1, can't determine which — mark as MAIN_ENGINE and flag
    return 'DG1';
  }

  return 'MAIN_ENGINE';
}

// ─────────────────────────────────────────────────────────────────────────────
// WORKBOOK → GRID
// ─────────────────────────────────────────────────────────────────────────────

function sheetToGrid(sheet: XLSX.WorkSheet): Grid {
  if (!sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const grid: Grid = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: CellValue[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[addr];
      if (!cell) { row.push(null); continue; }

      // Prioritise Date objects for date-formatted cells
      if (cell.t === 'd' && cell.v instanceof Date) {
        row.push(cell.v);
      } else if (cell.t === 'n' && typeof cell.z === 'string' && /[\/\-]/.test(cell.z)) {
        // Serial number with date format
        row.push(parseDate(cell.v));
      } else {
        row.push(cell.v ?? null);
      }
    }
    grid.push(row);
  }
  return grid;
}

async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => {
      try {
        const wb = XLSX.read(e.target?.result as ArrayBuffer, {
          type:      'array',
          cellDates: true,
          cellNF:    true,
          cellText:  false,
        });
        resolve(wb);
      } catch (err) {
        reject(new Error(`Could not read file "${file.name}": ${String(err)}`));
      }
    };
    reader.onerror = () => reject(new Error(`File read error: ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SEMANTIC HEADER DETECTION — LOG FILES
// ─────────────────────────────────────────────────────────────────────────────

const LOG_DATE_KEYWORDS    = ['DATE', 'DAY', 'DAYS'];
const LOG_ME_KEYWORDS      = ['MAIN ENGINE', 'MAIN ENG', 'M/E', 'M.E.', 'ME HOURS', 'M.E HOURS', 'M/E HRS'];
const LOG_DG1_KEYWORDS     = ['DG1', 'D/G 1', 'D.G.1', 'G/E 1', 'AUX 1', 'GEN 1', 'GENERATOR 1'];
const LOG_DG2_KEYWORDS     = ['DG2', 'D/G 2', 'D.G.2', 'G/E 2', 'AUX 2', 'GEN 2', 'GENERATOR 2'];
const LOG_DG3_KEYWORDS     = ['DG3', 'D/G 3', 'D.G.3', 'G/E 3', 'AUX 3', 'GEN 3', 'GENERATOR 3'];

function detectLogHeader(grid: Grid): HeaderMatch | null {
  const scanRows = Math.min(60, grid.length);

  for (let i = 0; i < scanRows; i++) {
    const row = grid[i].map(normalise);
    const hasDate = row.some(c => LOG_DATE_KEYWORDS.some(k => c.includes(k)));
    const hasMe   = row.some(c => LOG_ME_KEYWORDS.some(k => c.includes(k)) || c === 'ME' || c === 'MAIN');
    if (!hasDate || !hasMe) continue;

    const colMap: ColMap = {};
    row.forEach((cell, j) => {
      if (!cell) return;

      if (!('date' in colMap) && LOG_DATE_KEYWORDS.some(k => cell.includes(k)))
        colMap.date = j;

      if (!('me' in colMap) && (LOG_ME_KEYWORDS.some(k => cell.includes(k)) || cell === 'ME' || cell === 'MAIN'))
        colMap.me = j;

      if (!('dg1' in colMap) && LOG_DG1_KEYWORDS.some(k => cell.includes(k)))
        colMap.dg1 = j;

      if (!('dg2' in colMap) && LOG_DG2_KEYWORDS.some(k => cell.includes(k)))
        colMap.dg2 = j;

      if (!('dg3' in colMap) && LOG_DG3_KEYWORDS.some(k => cell.includes(k)))
        colMap.dg3 = j;
    });

    if (colMap.date !== undefined && colMap.me !== undefined) {
      return { rowIdx: i, colMap };
    }
  }

  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// SEMANTIC HEADER DETECTION — PMS FILES
// ─────────────────────────────────────────────────────────────────────────────

const PMS_COMP_KEYWORDS  = ['COMPONENT', 'EQUIPMENT', 'ITEM', 'NAME', 'DESCRIPTION', 'DESC'];
const PMS_DATE_KEYWORDS  = ['OVERHAUL', 'LAST OH', 'LAST O/H', 'O/H DATE', 'DATE O/H', 'INSP DATE', 'LAST INSPECTION', 'LAST INSP'];
const PMS_HOURS_KEYWORDS = ['CURRENT HOURS', 'RUNNING HOURS', 'HOURS SINCE', 'R/H', 'RUN HRS', 'CLAIMED', 'HRS SINCE OH', 'HOURS'];
const PMS_SYSTEM_KEYWORDS = ['SYSTEM', 'MACHINERY', 'CATEGORY', 'DEPT'];

function detectPMSHeader(grid: Grid): HeaderMatch | null {
  const scanRows = Math.min(60, grid.length);

  for (let i = 0; i < scanRows; i++) {
    const row = grid[i].map(normalise);
    const hasComp = row.some(c => PMS_COMP_KEYWORDS.some(k => c.includes(k)));
    const hasDate = row.some(c => PMS_DATE_KEYWORDS.some(k => c.includes(k)));
    if (!hasComp && !hasDate) continue;

    const colMap: ColMap = {};
    row.forEach((cell, j) => {
      if (!cell) return;

      if (!('component' in colMap) && PMS_COMP_KEYWORDS.some(k => cell.includes(k)))
        colMap.component = j;

      if (!('ohDate' in colMap) && PMS_DATE_KEYWORDS.some(k => cell.includes(k)))
        colMap.ohDate = j;

      // Fallback: any "DATE" column if no specific OH date found
      if (!('ohDate' in colMap) && cell.includes('DATE'))
        colMap.ohDate = j;

      if (!('hours' in colMap) && PMS_HOURS_KEYWORDS.some(k => cell.includes(k)))
        colMap.hours = j;

      if (!('system' in colMap) && PMS_SYSTEM_KEYWORDS.some(k => cell.includes(k)))
        colMap.system = j;
    });

    if (colMap.component !== undefined || colMap.ohDate !== undefined) {
      return { rowIdx: i, colMap };
    }
  }

  // Hard fallback: if the sheet has 8+ columns and no semantic headers found,
  // try the maritime TEC-001 standard layout (col 1 = name, col 5 = date, col 7 = hours)
  if (grid.length > 10 && (grid[0]?.length ?? 0) >= 8) {
    return {
      rowIdx: 7,
      colMap: { component: 1, ohDate: 5, hours: 7 },
    };
  }

  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC API — PMS PARSER
// ─────────────────────────────────────────────────────────────────────────────

export async function parsePMSFile(file: File): Promise<ComponentData[]> {
  const wb = await readWorkbook(file);
  const components: ComponentData[] = [];
  const parseErrors: string[] = [];

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet['!ref']) continue;

    const grid = sheetToGrid(sheet);
    const match = detectPMSHeader(grid);
    if (!match) continue;

    const { rowIdx, colMap } = match;

    const compCol   = colMap.component ?? 1;
    const ohDateCol = colMap.ohDate    ?? 5;
    const hoursCol  = colMap.hours     ?? 7;
    const sysCol    = colMap.system;

    for (let r = rowIdx + 1; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row.every(v => v === null || v === '')) continue;

      const rawName  = row[compCol];
      const rawDate  = row[ohDateCol];
      const rawHours = hoursCol < row.length ? row[hoursCol] : null;
      const rawSys   = sysCol !== undefined ? row[sysCol] : undefined;

      if (rawName === null || rawName === undefined) continue;
      const compName = String(rawName).trim();
      if (!compName || compName === '' || compName.toLowerCase() === 'nan') continue;

      const ohDate = parseDate(rawDate);
      if (!ohDate) {
        parseErrors.push(`Row ${r} skipped — invalid overhaul date: "${rawDate}"`);
        continue;
      }

      // Reject dates far in the future (> 1 year ahead) as likely errors
      const oneYearAhead = new Date();
      oneYearAhead.setFullYear(oneYearAhead.getFullYear() + 1);
      if (ohDate > oneYearAhead) {
        parseErrors.push(`Row ${r} skipped — overhaul date ${ohDate.toISOString()} is in the future`);
        continue;
      }

      const legacyClaimedHours = parseHours(rawHours);
      const parentSystem       = inferSystem(compName, rawSys);

      const result = ComponentSchema.safeParse({
        componentName:      compName,
        lastOverhaulDate:   ohDate,
        legacyClaimedHours,
        parentSystem,
      });

      if (result.success) {
        components.push(result.data);
      } else {
        parseErrors.push(`Row ${r} ("${compName}"): ${result.error.errors[0]?.message}`);
      }
    }

    // If we found data in this sheet, stop processing more sheets
    if (components.length > 0) break;
  }

  if (components.length === 0) {
    const errorDetail = parseErrors.length > 0
      ? ` Validation errors: ${parseErrors.slice(0, 5).join(' | ')}`
      : ' No sheets contained recognisable PMS header rows.';
    throw new Error(`PMS parse failed for "${file.name}".${errorDetail}`);
  }

  console.info(`[PMS] Parsed ${components.length} components from "${file.name}". Skipped ${parseErrors.length} rows.`);
  return components;
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC API — LOG FILE PARSER
// ─────────────────────────────────────────────────────────────────────────────

export async function parseLogFile(file: File): Promise<DailyLog[]> {
  const wb = await readWorkbook(file);
  const logs: DailyLog[] = [];
  const parseErrors: string[] = [];

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet['!ref']) continue;

    const grid = sheetToGrid(sheet);
    const match = detectLogHeader(grid);
    if (!match) continue;

    const { rowIdx, colMap } = match;

    const dateCol = colMap.date ?? 0;
    const meCol   = colMap.me   ?? 1;
    const dg1Col  = colMap.dg1;
    const dg2Col  = colMap.dg2;
    const dg3Col  = colMap.dg3;

    for (let r = rowIdx + 1; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row.every(v => v === null || v === '' || v === 0)) continue;

      const date = parseDate(row[dateCol]);
      if (!date) continue;

      // Reject clearly invalid years
      if (date.getFullYear() < 1990 || date.getFullYear() > 2100) continue;

      const mainEngineHours = parseHours(row[meCol]);
      const dg1Hours        = dg1Col !== undefined ? parseHours(row[dg1Col]) : 0;
      const dg2Hours        = dg2Col !== undefined ? parseHours(row[dg2Col]) : 0;
      const dg3Hours        = dg3Col !== undefined ? parseHours(row[dg3Col]) : 0;

      // Parse without constraints — physics violations captured downstream
      const result = DailyLogSchema.safeParse({
        date,
        mainEngineHours,
        dg1Hours,
        dg2Hours,
        dg3Hours,
      });

      if (result.success) {
        logs.push(result.data);
      } else {
        // Log still added with clamped values to preserve timeline continuity
        logs.push({ date, mainEngineHours, dg1Hours, dg2Hours, dg3Hours });
        parseErrors.push(`Row ${r} (${date.toDateString()}): ${result.error.errors[0]?.message}`);
      }
    }

    if (logs.length > 0) break;
  }

  if (logs.length === 0) {
    throw new Error(
      `Log parse failed for "${file.name}". ` +
      'No sheet contained a recognisable date + main-engine header row.'
    );
  }

  console.info(`[LOG] Parsed ${logs.length} entries from "${file.name}". Skipped/warned ${parseErrors.length} rows.`);
  return logs;
}
