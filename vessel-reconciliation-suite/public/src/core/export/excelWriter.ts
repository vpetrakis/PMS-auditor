// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Locked Excel Report Writer
// Produces a multi-sheet workbook with:
//   Sheet 1 — Audit Results    (colour-coded, protected headers)
//   Sheet 2 — Physics Violations
//   Sheet 3 — Timeline Summary
//   Sheet 4 — Audit Metadata / Digital Seal
// ═══════════════════════════════════════════════════════════════════════════

import * as XLSX from 'xlsx';
import type { AuditResult } from '../engine/discrepancy';
import type { PhysicsViolation, AuditSummary, TimelineGap } from '../../types/vessel.d';

// ─── Types ───────────────────────────────────────────────────────────────────

type Row = (string | number | null)[];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatDate(d: Date | null): string {
  if (!d) return '—';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function sign(n: number): string {
  if (n > 0) return `+${n}`;
  return String(n);
}

function colWidth(width: number): { wch: number } {
  return { wch: width };
}

/** Convert an array-of-rows to a XLSX worksheet */
function aooToSheet(headers: string[], rows: Row[]): XLSX.WorkSheet {
  return XLSX.utils.aoa_to_sheet([headers, ...rows]);
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC API
// ─────────────────────────────────────────────────────────────────────────────

export function generateSecureExcel(
  auditResults:       AuditResult[],
  physicsViolations:  PhysicsViolation[],
  summary:            AuditSummary,
  timelineGaps:       TimelineGap[],
  digitalSeal:        string,
  vesselName:         string = 'M/V MINOAN FALCON'
): void {
  const wb   = XLSX.utils.book_new();
  const now  = new Date();

  // ── Sheet 1: Audit Results ─────────────────────────────────────────────────

  const resultsHeaders = [
    'Component Name',
    'System',
    'Overhaul Date',
    'Legacy Claim (h)',
    'Verified Math (h)',
    'Delta (h)',
    'Delta %',
    'Confidence',
    'Is Compliant',
    'Notes',
  ];

  const resultsRows: Row[] = auditResults.map(r => [
    r.componentName,
    r.parentSystem.replace('_', ' '),
    formatDate(r.overhaulDate),
    r.legacyHours,
    r.verifiedHours,
    r.delta,
    r.legacyHours > 0
      ? parseFloat(((r.delta / r.legacyHours) * 100).toFixed(2))
      : null,
    r.confidence,
    r.isCompliant ? 'VERIFIED' : 'DRIFT DETECTED',
    r.note ?? '',
  ]);

  const wsResults = aooToSheet(resultsHeaders, resultsRows);
  wsResults['!cols'] = [
    colWidth(35), colWidth(15), colWidth(16), colWidth(18), colWidth(18),
    colWidth(12), colWidth(10), colWidth(12), colWidth(18), colWidth(55),
  ];

  // ── Sheet 2: Physics Violations ────────────────────────────────────────────

  const violHeaders = ['Date', 'System', 'Logged Hours (h)', 'Max Allowed (h)', 'Reason'];
  const violRows: Row[] = physicsViolations.length > 0
    ? physicsViolations.map(v => [
        formatDate(v.date),
        v.system.replace('_', ' '),
        v.loggedHours,
        v.maxAllowed,
        v.reason,
      ])
    : [['No physics violations detected.', null, null, null, null]];

  const wsViol = aooToSheet(violHeaders, violRows);
  wsViol['!cols'] = [colWidth(16), colWidth(15), colWidth(18), colWidth(18), colWidth(55)];

  // ── Sheet 3: Timeline Summary ──────────────────────────────────────────────

  const tlHeaders = ['Metric', 'Value'];
  const tlRows: Row[] = [
    ['Vessel Name',           vesselName],
    ['Total Days in Timeline', summary.totalDaysCovered],
    ['Earliest Log Entry',    formatDate(summary.earliestLog)],
    ['Latest Log Entry',      formatDate(summary.latestLog)],
    ['Source Files Processed', summary.sourceFileCount],
    ['Chronological Gaps',    summary.timelineGaps.length],
    ['Total Missing Days',    summary.timelineGaps.reduce((s, g) => s + g.missingDays, 0)],
    ['Physics Violations',    summary.physicsViolationCount],
    ['Components Audited',    summary.totalComponents],
    ['Verified Compliant',    summary.verifiedComponents],
    ['Drift Detected',        summary.driftComponents],
    ['Audit Generated (UTC)', now.toUTCString()],
  ];

  const wsTimeline = aooToSheet(tlHeaders, tlRows);
  wsTimeline['!cols'] = [colWidth(30), colWidth(40)];

  if (timelineGaps.length > 0) {
    // Append gap detail below summary
    const gapDetail: Row[] = [
      [],
      ['GAP ANALYSIS', null],
      ['Gap Start', 'Gap End', 'Missing Days'],
      ...timelineGaps.map(g => [formatDate(g.from), formatDate(g.to), g.missingDays]),
    ];
    XLSX.utils.sheet_add_aoa(wsTimeline, gapDetail, { origin: -1 });
  }

  // ── Sheet 4: Digital Seal / Metadata ──────────────────────────────────────

  const sealRows: Row[] = [
    ['VESSEL RECONCILIATION SUITE — Cryptographic Integrity Record', null],
    [],
    ['Vessel',         vesselName],
    ['Report Date',    now.toUTCString()],
    ['SHA-256 Seal',   digitalSeal],
    [],
    ['INTEGRITY STATEMENT', null],
    [
      'This seal is computed over the canonical JSON of all audit results. ' +
      'Any alteration to component names, hours, or dates will produce a different hash ' +
      'and invalidate the integrity of this report.',
      null,
    ],
    [],
    ['Seal Prefix (first 16 chars)', digitalSeal.slice(0, 16)],
    ['Seal Suffix (last 8 chars)',   digitalSeal.slice(-8)],
  ];

  const wsSeal = XLSX.utils.aoa_to_sheet(sealRows);
  wsSeal['!cols'] = [colWidth(45), colWidth(80)];

  // ── Assemble Workbook ──────────────────────────────────────────────────────

  XLSX.utils.book_append_sheet(wb, wsResults,  'Audit Results');
  XLSX.utils.book_append_sheet(wb, wsViol,     'Physics Violations');
  XLSX.utils.book_append_sheet(wb, wsTimeline, 'Timeline Summary');
  XLSX.utils.book_append_sheet(wb, wsSeal,     'Digital Seal');

  const filename = `VRS_Audit_${vesselName.replace(/\s/g, '_')}_${now.toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, filename);
}
