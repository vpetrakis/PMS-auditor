// ═══════════════════════════════════════════════════════════════════════════
// Global Types — import from here as: import type { ... } from '../types/vessel'
// ═══════════════════════════════════════════════════════════════════════════

export type SystemLabel = 'MAIN_ENGINE' | 'DG1' | 'DG2' | 'DG3';

export interface PhysicsViolation {
  date: Date;
  system: SystemLabel;
  loggedHours: number;
  maxAllowed: number;
  reason: string;
}

export interface TimelineGap {
  from: Date;
  to: Date;
  missingDays: number;
}

export interface AuditSummary {
  totalComponents: number;
  verifiedComponents: number;
  driftComponents: number;
  totalDaysCovered: number;
  earliestLog: Date | null;
  latestLog: Date | null;
  timelineGaps: TimelineGap[];
  physicsViolationCount: number;
  sourceFileCount: number;
}
