// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Discrepancy & Physics Audit Engine
//
// Phase 1 — Physics Violations: any day where a system logs > 24h or < 0h
// Phase 2 — Baseline Drift:     PMS claimed hours vs. mathematically verified
//                                hours summed from the master timeline
// ═══════════════════════════════════════════════════════════════════════════

import type { DailyLog, ComponentData } from '../schemas/tec001.schema';
import type { PhysicsViolation, SystemLabel } from '../../types/vessel.d';
import {
  type MasterTimeline,
  sumSystemHoursAfterDate,
  hasGapsAfterDate,
  missingDaysInPeriod,
} from './timeline';

// ─── Audit Result ─────────────────────────────────────────────────────────────

export interface AuditResult {
  componentName:    string;
  parentSystem:     SystemLabel;
  overhaulDate:     Date;
  legacyHours:      number;
  verifiedHours:    number;
  delta:            number;            // verifiedHours - legacyHours
  isCompliant:      boolean;           // |delta| ≤ TOLERANCE
  confidence:       'HIGH' | 'MEDIUM' | 'LOW';
  note:             string | null;
  missingDaysInWindow: number;
}

// ─── Constants ───────────────────────────────────────────────────────────────

/** Hours tolerance for rounding / minor data-entry differences (6 min = 0.1 h) */
const TOLERANCE_HOURS = 0.5;

const SYSTEMS: ReadonlyArray<{ key: keyof DailyLog; label: SystemLabel }> = [
  { key: 'mainEngineHours', label: 'MAIN_ENGINE' },
  { key: 'dg1Hours',        label: 'DG1'         },
  { key: 'dg2Hours',        label: 'DG2'         },
  { key: 'dg3Hours',        label: 'DG3'         },
];

// ─── Phase 1: Physics Violation Detector ──────────────────────────────────────

/**
 * Scans every daily log entry across all systems for impossible values.
 * Does NOT mutate or discard the offending records — they are reported
 * and still included in the arithmetic so the forensic trail is complete.
 */
export function detectPhysicsViolations(logs: DailyLog[]): PhysicsViolation[] {
  const violations: PhysicsViolation[] = [];

  for (const log of logs) {
    for (const { key, label } of SYSTEMS) {
      const hrs = log[key] as number;
      if (hrs > 24) {
        violations.push({
          date:        log.date,
          system:      label,
          loggedHours: hrs,
          maxAllowed:  24,
          reason:      `Exceeds 24-hour physical maximum (${hrs.toFixed(2)} h logged)`,
        });
      } else if (hrs < 0) {
        violations.push({
          date:        log.date,
          system:      label,
          loggedHours: hrs,
          maxAllowed:  24,
          reason:      `Negative hours are physically impossible (${hrs.toFixed(2)} h logged)`,
        });
      }
    }
  }

  return violations;
}

// ─── Phase 2: Baseline Drift Calculator ───────────────────────────────────────

/**
 * For each PMS component:
 *   1. Sums system hours from the master timeline since the last overhaul date.
 *   2. Computes the delta vs. the legacy claimed hours.
 *   3. Assigns a confidence level based on timeline completeness.
 */
export function calculateDiscrepancies(
  timeline: MasterTimeline,
  components: ComponentData[]
): AuditResult[] {
  return components.map(comp => {
    const { componentName, lastOverhaulDate, legacyClaimedHours, parentSystem } = comp;

    const verifiedRaw   = sumSystemHoursAfterDate(timeline, lastOverhaulDate, parentSystem);
    const verifiedHours = Math.round(verifiedRaw * 10) / 10;
    const delta         = Math.round((verifiedHours - legacyClaimedHours) * 10) / 10;
    const isCompliant   = Math.abs(delta) <= TOLERANCE_HOURS;

    // ── Confidence Assessment ────────────────────────────────────────────────
    let confidence: AuditResult['confidence'] = 'HIGH';
    let note: string | null = null;
    const missingDaysInWindow = missingDaysInPeriod(timeline, lastOverhaulDate);

    if (!timeline.startDate || timeline.startDate > lastOverhaulDate) {
      // Timeline doesn't reach back to the overhaul date
      const gapDays = timeline.startDate
        ? Math.round((timeline.startDate.getTime() - lastOverhaulDate.getTime()) / 86_400_000)
        : 0;
      confidence = gapDays > 30 ? 'LOW' : 'MEDIUM';
      note = `Log data starts ${gapDays}d after overhaul. Hours likely underreported.`;
    } else if (hasGapsAfterDate(timeline, lastOverhaulDate)) {
      confidence = 'MEDIUM';
      note = `Timeline has ${missingDaysInWindow} missing day(s) after overhaul. Hours may be incomplete.`;
    }

    return {
      componentName,
      parentSystem,
      overhaulDate:       lastOverhaulDate,
      legacyHours:        legacyClaimedHours,
      verifiedHours,
      delta,
      isCompliant,
      confidence,
      note,
      missingDaysInWindow,
    };
  });
}
