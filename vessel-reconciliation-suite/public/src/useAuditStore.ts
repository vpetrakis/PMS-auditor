import { create } from 'zustand';
import { DailyLog, ComponentData } from '../core/schemas/tec001.schema';
import { AuditResult, calculateDiscrepancies } from '../core/engine/discrepancy';
import { validatePhysicsConstraints } from '../core/engine/timeline';
import { generateDigitalSeal } from '../core/crypto/hash';
import { PhysicsViolation } from '../types/vessel';

interface AuditState {
  dailyLogs: DailyLog[];
  components: ComponentData[];
  auditResults: AuditResult[] | null;
  physicsViolations: PhysicsViolation[];
  digitalSeal: string | null;
  isProcessing: boolean;
  error: string | null;

  processFileAndAudit: (components: ComponentData[], logs: DailyLog[]) => Promise<void>;
  resetStore: () => void;
}

export const useAuditStore = create<AuditState>((set, get) => ({
  dailyLogs: [],
  components: [],
  auditResults: null,
  physicsViolations: [],
  digitalSeal: null,
  isProcessing: false,
  error: null,

  processFileAndAudit: async (components, logs) => {
    set({ isProcessing: true, error: null, components, dailyLogs: logs });
    try {
      // Phase 1: Physics Audit (Check Daily Logs for impossibilities)
      const violations = validatePhysicsConstraints(logs);
      
      // Phase 2: Date-to-Date Math Audit (Audit PMS through Logs)
      const results = calculateDiscrepancies(logs, components);
      
      // Phase 3: Cryptographic Seal
      const seal = await generateDigitalSeal(results);

      set({ 
        physicsViolations: violations, 
        auditResults: results, 
        digitalSeal: seal, 
        isProcessing: false 
      });
    } catch (err: any) {
      set({ error: err.message, isProcessing: false });
    }
  },

  resetStore: () => set({
    dailyLogs: [], components: [], auditResults: null, physicsViolations: [], digitalSeal: null, error: null
  })
}));
