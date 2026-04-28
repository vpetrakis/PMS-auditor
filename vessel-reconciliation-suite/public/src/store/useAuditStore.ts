// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Zustand Audit Store
// Single source of truth for the entire audit pipeline state machine.
// ═══════════════════════════════════════════════════════════════════════════

import { create } from 'zustand';
import type { DailyLog, ComponentData } from '../core/schemas/tec001.schema';
import type { AuditResult }             from '../core/engine/discrepancy';
import type { MasterTimeline }          from '../core/engine/timeline';
import type { PhysicsViolation, AuditSummary } from '../types/vessel.d';

import { stitchTimeline }                      from '../core/engine/timeline';
import { detectPhysicsViolations, calculateDiscrepancies } from '../core/engine/discrepancy';
import { generateDigitalSeal }                 from '../core/crypto/hash';
import { generateSecurePDF }                   from '../core/export/pdfBuilder';
import { generateSecureExcel }                 from '../core/export/excelWriter';

// ─── State Shape ─────────────────────────────────────────────────────────────

interface AuditState {
  // ── Raw Ingested Data ──
  components:     ComponentData[];
  allLogSets:     DailyLog[][];   // one DailyLog[] per uploaded file
  logFileNames:   string[];

  // ── Computed Results ──
  timeline:          MasterTimeline | null;
  auditResults:      AuditResult[]  | null;
  physicsViolations: PhysicsViolation[];
  digitalSeal:       string | null;
  summary:           AuditSummary   | null;

  // ── UI State ──
  isProcessing:  boolean;
  error:         string | null;
  vesselName:    string;

  // ── Actions ──
  setComponents:   (comps: ComponentData[]) => void;
  addLogSet:       (logs: DailyLog[], filename: string) => void;
  removeLogFile:   (filename: string) => void;
  setVesselName:   (name: string) => void;
  runAuditPipeline: () => Promise<void>;
  exportPDF:       () => Promise<void>;
  exportExcel:     () => void;
  resetStore:      () => void;
}

// ─── Initial State ────────────────────────────────────────────────────────────

const initial = {
  components:        [] as ComponentData[],
  allLogSets:        [] as DailyLog[][],
  logFileNames:      [] as string[],
  timeline:          null,
  auditResults:      null,
  physicsViolations: [] as PhysicsViolation[],
  digitalSeal:       null,
  summary:           null,
  isProcessing:      false,
  error:             null,
  vesselName:        'M/V MINOAN FALCON',
};

// ─── Store ────────────────────────────────────────────────────────────────────

export const useAuditStore = create<AuditState>((set, get) => ({
  ...initial,

  // ── Setters ─────────────────────────────────────────────────────────────────

  setComponents: (comps) => set({ components: comps, error: null }),

  addLogSet: (logs, filename) =>
    set(s => ({
      allLogSets:    [...s.allLogSets,    logs    ],
      logFileNames:  [...s.logFileNames,  filename],
      error: null,
    })),

  removeLogFile: (filename) =>
    set(s => {
      const idx = s.logFileNames.indexOf(filename);
      if (idx === -1) return s;
      const logSets  = [...s.allLogSets];
      const names    = [...s.logFileNames];
      logSets.splice(idx, 1);
      names.splice(idx, 1);
      return { allLogSets: logSets, logFileNames: names };
    }),

  setVesselName: (name) => set({ vesselName: name }),

  // ── Core Pipeline ────────────────────────────────────────────────────────────

  runAuditPipeline: async () => {
    const { components, allLogSets } = get();

    if (components.length === 0)
      return set({ error: 'No PMS data loaded. Upload a TEC-001 file first.' });
    if (allLogSets.length === 0)
      return set({ error: 'No log files loaded. Upload at least one operating-hours file.' });

    set({ isProcessing: true, error: null });

    try {
      // Step 1: Build master timeline
      const timeline = stitchTimeline(allLogSets, allLogSets.length);

      if (!timeline.startDate)
        throw new Error('Timeline is empty after stitching. Check that log files contain valid date rows.');

      // Step 2: Detect physics violations (does not modify timeline data)
      const allLogs         = timeline.logs;
      const physicsViolations = detectPhysicsViolations(allLogs);

      // Step 3: Calculate discrepancies
      const auditResults = calculateDiscrepancies(timeline, components);

      // Step 4: Generate cryptographic seal
      const digitalSeal = await generateDigitalSeal(auditResults);

      // Step 5: Build summary
      const summary: AuditSummary = {
        totalComponents:       auditResults.length,
        verifiedComponents:    auditResults.filter(r => r.isCompliant).length,
        driftComponents:       auditResults.filter(r => !r.isCompliant).length,
        totalDaysCovered:      timeline.totalDays,
        earliestLog:           timeline.startDate,
        latestLog:             timeline.endDate,
        timelineGaps:          timeline.gaps,
        physicsViolationCount: physicsViolations.length,
        sourceFileCount:       timeline.sourceFiles,
      };

      set({
        timeline,
        auditResults,
        physicsViolations,
        digitalSeal,
        summary,
        isProcessing: false,
      });

    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      set({ error: message, isProcessing: false });
    }
  },

  // ── Exports ──────────────────────────────────────────────────────────────────

  exportPDF: async () => {
    const { auditResults, physicsViolations, summary, digitalSeal, vesselName } = get();
    if (!auditResults || !digitalSeal || !summary) return;
    await generateSecurePDF(auditResults, physicsViolations, summary, digitalSeal, vesselName);
  },

  exportExcel: () => {
    const { auditResults, physicsViolations, summary, timeline, digitalSeal, vesselName } = get();
    if (!auditResults || !digitalSeal || !summary || !timeline) return;
    generateSecureExcel(
      auditResults,
      physicsViolations,
      summary,
      timeline.gaps,
      digitalSeal,
      vesselName
    );
  },

  // ── Reset ────────────────────────────────────────────────────────────────────

  resetStore: () => set({ ...initial }),
}));
