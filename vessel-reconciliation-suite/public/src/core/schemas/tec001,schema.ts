// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — TEC-001 Zod Schemas
// NOTE: Physics violation detection (>24h) runs as a separate audit step
// so that bad data is captured and reported rather than silently dropped.
// ═══════════════════════════════════════════════════════════════════════════

import { z } from 'zod';

// ─── Daily Running Hours ────────────────────────────────────────────────────

export const DailyLogSchema = z.object({
  date: z.date({
    required_error: 'A valid date is required for every log entry.',
    invalid_type_error: 'Date must be a valid Date object.',
  }),
  mainEngineHours: z.number().min(0, 'ME hours cannot be negative.'),
  dg1Hours: z.number().min(0, 'DG1 hours cannot be negative.'),
  dg2Hours: z.number().min(0, 'DG2 hours cannot be negative.'),
  dg3Hours: z.number().min(0, 'DG3 hours cannot be negative.'),
});

// ─── PMS Component Entry ────────────────────────────────────────────────────

export const ComponentSchema = z.object({
  componentName: z.string().min(1, 'Component name cannot be empty.'),
  lastOverhaulDate: z.date({
    required_error: 'Last overhaul date is required.',
    invalid_type_error: 'Overhaul date must be a valid Date.',
  }),
  legacyClaimedHours: z.number().min(0, 'Claimed hours cannot be negative.'),
  parentSystem: z.enum(['MAIN_ENGINE', 'DG1', 'DG2', 'DG3'], {
    errorMap: () => ({ message: 'System must be MAIN_ENGINE, DG1, DG2, or DG3.' }),
  }),
});

// ─── Inferred Types ─────────────────────────────────────────────────────────

export type DailyLog = z.infer<typeof DailyLogSchema>;
export type ComponentData = z.infer<typeof ComponentSchema>;
