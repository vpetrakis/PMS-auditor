import { DailyLog, ComponentData } from '../schemas/tec001.schema';

export interface AuditResult {
  componentName: string;
  legacyHours: number;
  verifiedHours: number;
  delta: number;
  isCompliant: boolean;
}

export function calculateDiscrepancies(
  dailyLogs: DailyLog[],
  components: ComponentData[]
): AuditResult[] {
  // 1. Sort daily logs chronologically to ensure accurate summation
  const sortedLogs = [...dailyLogs].sort((a, b) => a.date.getTime() - b.date.getTime());

  // 2. Map over each component to calculate its true running hours
  return components.map((component) => {
    // Find all logs that occurred ON or AFTER the last overhaul date
    const applicableLogs = sortedLogs.filter(
      (log) => log.date.getTime() >= component.lastOverhaulDate.getTime()
    );

    // Sum the correct hours based on the parent system
    const verifiedHours = applicableLogs.reduce((sum, log) => {
      switch (component.parentSystem) {
        case "MAIN_ENGINE": return sum + log.mainEngineHours;
        case "DG1": return sum + log.dg1Hours;
        case "DG2": return sum + log.dg2Hours;
        case "DG3": return sum + log.dg3Hours;
        default: return sum;
      }
    }, 0);

    const delta = verifiedHours - component.legacyClaimedHours;

    return {
      componentName: component.componentName,
      legacyHours: component.legacyClaimedHours,
      verifiedHours: verifiedHours,
      delta: delta,
      isCompliant: delta === 0, // True ONLY if mathematically perfect
    };
  });
}
