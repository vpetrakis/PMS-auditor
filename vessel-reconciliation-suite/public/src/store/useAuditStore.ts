import { create } from 'zustand';
import { DailyLog, ComponentData } from '../core/schemas/tec001.schema';
import { AuditResult, calculateDiscrepancies } from '../core/engine/discrepancy';
import { generateDigitalSeal } from '../core/crypto/hash';

interface AuditState {
  // State
  dailyLogs: DailyLog[];
  components: ComponentData[];
  auditResults: AuditResult[] | null;
  digitalSeal: string | null;
  isProcessing: boolean;
  error: string | null;

  // Actions
  ingestData: (logs: DailyLog[], comps: ComponentData[]) => void;
  runAuditPipeline: () => Promise<void>;
  resetStore: () => void;
}

export const useAuditStore = create<AuditState>((set, get) => ({
  dailyLogs: [],
  components: [],
  auditResults: null,
  digitalSeal: null,
  isProcessing: false,
  error: null,

  ingestData: (logs, comps) => set({ dailyLogs: logs, components: comps, error: null }),

  runAuditPipeline: async () => {
    set({ isProcessing: true, error: null });
    try {
      const { dailyLogs, components } = get();
      
      // Step 1: Run the Math Engine
      const results = calculateDiscrepancies(dailyLogs, components);
      
      // Step 2: Generate Cryptographic Hash
      const seal = await generateDigitalSeal(results);

      // Step 3: Lock state
      set({ auditResults: results, digitalSeal: seal, isProcessing: false });
    } catch (err: any) {
      set({ error: err.message, isProcessing: false });
    }
  },

  resetStore: () => set({
    dailyLogs: [],
    components: [],
    auditResults: null,
    digitalSeal: null,
    error: null
  })
}));
