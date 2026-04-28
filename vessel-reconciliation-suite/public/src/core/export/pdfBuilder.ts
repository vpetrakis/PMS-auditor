import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import QRCode from 'qrcode';
import { AuditResult } from '../engine/discrepancy';

export async function generateSecurePDF(results: AuditResult[], digitalSeal: string, vesselName: string): Promise<void> {
  const doc = new jsPDF('p', 'pt', 'a4');
  const dateStr = new Date().toLocaleDateString('en-GB');

  // 1. Header
  doc.setFontSize(18);
  doc.setTextColor(40, 40, 40);
  doc.text("OFFICIAL RUNNING HOURS BASELINE RECONCILIATION", 40, 50);
  
  doc.setFontSize(12);
  doc.text(`Vessel: M/V ${vesselName}`, 40, 75);
  doc.text(`Date of Audit: ${dateStr}`, 40, 95);

  // 2. Generate QR Code containing the Hash
  try {
    const qrDataUrl = await QRCode.toDataURL(`VERIFIED HASH: ${digitalSeal}`, { width: 100, margin: 0 });
    doc.addImage(qrDataUrl, 'PNG', 450, 40, 80, 80);
  } catch (err) {
    console.error("QR Code Generation Failed", err);
  }

  // 3. Build Table Data
  const tableBody = results.map(r => [
    r.componentName,
    r.legacyHours.toString(),
    r.verifiedHours.toString(),
    r.delta > 0 ? `+${r.delta}` : r.delta.toString(),
    r.isCompliant ? 'VERIFIED' : 'CORRECTED'
  ]);

  // 4. Draw Table
  autoTable(doc, {
    startY: 140,
    head: [['Component', 'Legacy Hours', 'Verified Hours', 'Delta', 'Status']],
    body: tableBody,
    theme: 'grid',
    headStyles: { fillColor: [41, 128, 185], textColor: 255 },
    alternateRowStyles: { fillColor: [245, 245, 245] },
  });

  // 5. Cryptographic Footer & Signatures
  const finalY = (doc as any).lastAutoTable.finalY || 140;
  
  doc.setFontSize(9);
  doc.setTextColor(100, 100, 100);
  doc.text(`CRYPTOGRAPHIC SHA-256 SEAL:`, 40, finalY + 40);
  doc.setFont("courier", "bold");
  doc.text(digitalSeal, 40, finalY + 55);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(0, 0, 0);
  doc.text("___________________________", 40, finalY + 120);
  doc.text("Chief Engineer Signature", 40, finalY + 135);

  doc.text("___________________________", 350, finalY + 120);
  doc.text("Class Surveyor / Tech Supt.", 350, finalY + 135);

  // 6. Save locally
  doc.save(`TEC-004_Verified_Baseline_${dateStr.replace(/\//g, '')}.pdf`);
}
