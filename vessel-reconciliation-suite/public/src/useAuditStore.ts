import { create } from 'zustand';
import type { DailyLog, ComponentData } from '../core/schemas/tec001.schema';
import type { AuditResult }             from '../core/engine/discrepancy';
import type { MasterTimeline }          from '../core/engine/timeline';
import type { PhysicsViolation, AuditSummary } from '../types/vessel';

import { stitchTimeline }                               from '../core/engine/timeline';
import { detectPhysicsViolations, calculateDiscrepancies } from '../core/engine/discrepancy';
import { generateDigitalSeal }                          from '../core/crypto/hash';
import { generateSecurePDF }                            from '../core/export/pdfBuilder';
import { generateSecureExcel }                          from '../core/export/excelWriter';

// ─── State ────────────────────────────────────────────────────────────────────

interface AuditState {
  // Ingested data
  components:    ComponentData[];
  allLogSets:    DailyLog[][];
  logFileNames:  string[];

  // Results
  timeline:          MasterTimeline | null;
  auditResults:      AuditResult[]  | null;
  physicsViolations: PhysicsViolation[];
  digitalSeal:       string | null;
  summary:           AuditSummary   | null;

  // UI
  isProcessing: boolean;
  error:        string | null;
  vesselName:   string;

  // Actions
  setComponents:    (c: ComponentData[]) => void;
  addLogSet:        (logs: DailyLog[], name: string) => void;
  removeLogFile:    (name: string) => void;
  setVesselName:    (n: string) => void;
  runAuditPipeline: () => Promise<void>;
  exportPDF:        () => Promise<void>;
  exportExcel:      () => void;
  resetStore:       () => void;
}

const INIT = {
  components: [] as ComponentData[], allLogSets: [] as DailyLog[][], logFileNames: [] as string[],
  timeline: null, auditResults: null, physicsViolations: [] as PhysicsViolation[],
  digitalSeal: null, summary: null, isProcessing: false, error: null, vesselName: 'M/V MINOAN FALCON',
};

export const useAuditStore = create<AuditState>((set, get) => ({
  ...INIT,

  setComponents:  c    => set({ components: c, error: null }),
  addLogSet:      (l, n) => set(s => ({ allLogSets: [...s.allLogSets, l], logFileNames: [...s.logFileNames, n], error: null })),
  removeLogFile:  name => set(s => {
    const i = s.logFileNames.indexOf(name);
    if (i < 0) return s;
    const ls = [...s.allLogSets]; ls.splice(i, 1);
    const ns = [...s.logFileNames]; ns.splice(i, 1);
    return { allLogSets: ls, logFileNames: ns };
  }),
  setVesselName:  n => set({ vesselName: n }),

  runAuditPipeline: async () => {
    const { components, allLogSets } = get();
    if (!components.length) return set({ error: 'No PMS data loaded. Upload a TEC-001 file first.' });
    if (!allLogSets.length)  return set({ error: 'No log files loaded. Upload at least one operating-hours log.' });

    set({ isProcessing: true, error: null });
    try {
      const timeline         = stitchTimeline(allLogSets, allLogSets.length);
      if (!timeline.startDate) throw new Error('Timeline is empty after stitching. Check that log files contain valid date rows.');

      const physicsViolations = detectPhysicsViolations(timeline.logs);
      const auditResults      = calculateDiscrepancies(timeline, components);
      const digitalSeal       = await generateDigitalSeal(auditResults);

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

      set({ timeline, auditResults, physicsViolations, digitalSeal, summary, isProcessing: false });
    } catch (err: unknown) {
      set({ error: err instanceof Error ? err.message : String(err), isProcessing: false });
    }
  },

  exportPDF: async () => {
    const { auditResults, physicsViolations, summary, digitalSeal, vesselName } = get();
    if (!auditResults || !digitalSeal || !summary) return;
    await generateSecurePDF(auditResults, physicsViolations, summary, digitalSeal, vesselName);
  },

  exportExcel: () => {
    const { auditResults, physicsViolations, summary, timeline, digitalSeal, vesselName } = get();
    if (!auditResults || !digitalSeal || !summary || !timeline) return;
    generateSecureExcel(auditResults, physicsViolations, summary, timeline.gaps, digitalSeal, vesselName);
  },

  resetStore: () => set({ ...INIT }),
}));
