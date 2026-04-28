// Global interface definitions for the entire application
export interface PhysicsViolation {
  date: Date;
  system: "MAIN_ENGINE" | "DG1" | "DG2" | "DG3";
  loggedHours: number;
  maxAllowed: number;
  reason: string;
}

export interface AuditTrailEntry {
  timestamp: string;
  component: string;
  action: string;
  oldValue: number | string;
  newValue: number | string;
  delta: number;
}
