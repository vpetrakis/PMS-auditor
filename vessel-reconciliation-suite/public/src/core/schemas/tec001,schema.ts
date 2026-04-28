import { z } from 'zod';

// Validates a single day of running hours
export const DailyLogSchema = z.object({
  date: z.date({
    required_error: "A valid Date is required.",
    invalid_type_error: "Date must be a valid Date object.",
  }),
  mainEngineHours: z.number().min(0).max(24, "Critical: Main Engine hours cannot exceed 24 in a single day."),
  dg1Hours: z.number().min(0).max(24, "Critical: DG1 hours cannot exceed 24."),
  dg2Hours: z.number().min(0).max(24, "Critical: DG2 hours cannot exceed 24."),
  dg3Hours: z.number().min(0).max(24, "Critical: DG3 hours cannot exceed 24."),
});

// Validates a component's legacy PMS entry
export const ComponentSchema = z.object({
  componentName: z.string().min(1, "Component name cannot be empty."),
  lastOverhaulDate: z.date(),
  legacyClaimedHours: z.number().min(0),
  parentSystem: z.enum(["MAIN_ENGINE", "DG1", "DG2", "DG3"]),
});

export type DailyLog = z.infer<typeof DailyLogSchema>;
export type ComponentData = z.infer<typeof ComponentSchema>;
