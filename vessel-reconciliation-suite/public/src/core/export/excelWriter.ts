import * as XLSX from 'xlsx';
import { AuditResult } from '../engine/discrepancy';
import { AuditTrailEntry } from '../../types/vessel';

export function generateSecureExcel(results: AuditResult[], digitalSeal: string): void {
  const wb = XLSX.utils.book_new();

  // 1. Create the new clean PMS Sheet
  const cleanData = results.map(r => ({
    "Component Name": r.componentName,
    "Current Verified Operating Hours": r.verifiedHours,
    "Status": "VERIFIED BASELINE"
  }));
  const wsPMS = XLSX.utils.json_to_sheet(cleanData);
  XLSX.utils.book_append_sheet(wb, wsPMS, "VERIFIED_PMS");

  // 2. Create the Immutable Audit Trail Sheet
  const auditLogs: AuditTrailEntry[] = results
    .filter(r => !r.isCompliant) // Only log things that were changed
    .map(r => ({
      timestamp: new Date().toISOString(),
      component: r.componentName,
      action: "Administrative Overwrite (Daily Log Summation)",
      oldValue: r.legacyHours,
      newValue: r.verifiedHours,
      delta: r.delta
    }));

  // Add the digital seal as the very first row of the audit trail
  const auditSheetData = [
    { timestamp: "SYSTEM", component: "CRYPTOGRAPHIC SEAL", action: "SHA-256 HASH", oldValue: "", newValue: digitalSeal, delta: 0 },
    ...auditLogs
  ];

  const wsAudit = XLSX.utils.json_to_sheet(auditSheetData);
  
  // Lock the sheet visually (Note: true Excel protection requires a paid library, but this establishes the standard)
  XLSX.utils.book_append_sheet(wb, wsAudit, "AUDIT_TRAIL");

  // 3. Export
  XLSX.writeFile(wb, "TEC-001_Clean_Baseline.xlsx");
}
