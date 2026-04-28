import { DailyLog } from '../schemas/tec001.schema';
import { PhysicsViolation } from '../../types/vessel';

export function validatePhysicsConstraints(logs: DailyLog[]): PhysicsViolation[] {
  const violations: PhysicsViolation[] = [];

  logs.forEach(log => {
    // 1. Layer 1 Audit: The 24-Hour Rule
    const checks = [
      { sys: "MAIN_ENGINE", hrs: log.mainEngineHours },
      { sys: "DG1", hrs: log.dg1Hours },
      { sys: "DG2", hrs: log.dg2Hours },
      { sys: "DG3", hrs: log.dg3Hours }
    ] as const;

    checks.forEach(({ sys, hrs }) => {
      if (hrs > 24) {
        violations.push({
          date: log.date,
          system: sys,
          loggedHours: hrs,
          maxAllowed: 24,
          reason: "Critical: Daily logged hours exceed 24 physical hours."
        });
      }
      if (hrs < 0) {
        violations.push({
          date: log.date,
          system: sys,
          loggedHours: hrs,
          maxAllowed: 24,
          reason: "Critical: Negative hours are mathematically impossible."
        });
      }
    });
  });

  return violations;
}
