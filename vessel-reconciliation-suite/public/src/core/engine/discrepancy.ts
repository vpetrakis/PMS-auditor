import type { DailyLog, ComponentData } from '../schemas/tec001.schema';
import type { PhysicsViolation, SystemLabel } from '../../types/vessel';
import { type MasterTimeline, sumHoursAfterDate, hasGapsAfter, missingDaysAfter } from './timeline';

export interface AuditResult {
  componentName:       string;
  parentSystem:        SystemLabel;
  overhaulDate:        Date;
  legacyHours:         number;
  verifiedHours:       number;
  delta:               number;
  isCompliant:         boolean;
  confidence:          'HIGH' | 'MEDIUM' | 'LOW';
  note:                string | null;
  missingDaysInWindow: number;
}

const SYSTEMS: ReadonlyArray<{ key: keyof DailyLog; label: SystemLabel }> = [
  { key: 'mainEngineHours', label: 'MAIN_ENGINE' },
  { key: 'dg1Hours',        label: 'DG1'         },
  { key: 'dg2Hours',        label: 'DG2'         },
  { key: 'dg3Hours',        label: 'DG3'         },
];

const TOLERANCE = 0.5; // hours

export function detectPhysicsViolations(logs: DailyLog[]): PhysicsViolation[] {
  const out: PhysicsViolation[] = [];
  for (const log of logs) {
    for (const { key, label } of SYSTEMS) {
      const h = log[key] as number;
      if (h > 24 || h < 0) out.push({ date: log.date, system: label, loggedHours: h, maxAllowed: 24, reason: h > 24 ? `Exceeds 24h max (${h.toFixed(2)}h)` : `Negative hours (${h.toFixed(2)}h)` });
    }
  }
  return out;
}

export function calculateDiscrepancies(tl: MasterTimeline, comps: ComponentData[]): AuditResult[] {
  return comps.map(c => {
    const rawV  = sumHoursAfterDate(tl, c.lastOverhaulDate, c.parentSystem);
    const vh    = Math.round(rawV * 10) / 10;
    const delta = Math.round((vh - c.legacyClaimedHours) * 10) / 10;
    const missing = missingDaysAfter(tl, c.lastOverhaulDate);

    let confidence: AuditResult['confidence'] = 'HIGH';
    let note: string | null = null;

    if (!tl.startDate || tl.startDate > c.lastOverhaulDate) {
      const gap = tl.startDate ? Math.round((tl.startDate.getTime() - c.lastOverhaulDate.getTime()) / 86_400_000) : 0;
      confidence = gap > 30 ? 'LOW' : 'MEDIUM';
      note = `Log data starts ${gap}d after overhaul — hours likely underreported.`;
    } else if (hasGapsAfter(tl, c.lastOverhaulDate)) {
      confidence = 'MEDIUM';
      note = `${missing} missing day(s) in timeline after overhaul date.`;
    }

    return { componentName: c.componentName, parentSystem: c.parentSystem, overhaulDate: c.lastOverhaulDate, legacyHours: c.legacyClaimedHours, verifiedHours: vh, delta, isCompliant: Math.abs(delta) <= TOLERANCE, confidence, note, missingDaysInWindow: missing };
  });
}
