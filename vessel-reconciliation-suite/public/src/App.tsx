import React from 'react';
import { useAuditStore } from './store/useAuditStore';
import { generateSecurePDF } from './core/export/pdfBuilder';
import { generateSecureExcel } from './core/export/excelWriter';
import { ShieldCheck, FileWarning, Download } from 'lucide-react';

export default function App() {
  const { auditResults, digitalSeal, isProcessing, error } = useAuditStore();

  const handleExport = async () => {
    if (!auditResults || !digitalSeal) return;
    generateSecureExcel(auditResults, digitalSeal);
    await generateSecurePDF(auditResults, digitalSeal, "MINOAN FALCON");
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans p-8">
      <header className="max-w-6xl mx-auto mb-8 border-b border-gray-200 pb-6">
        <h1 className="text-3xl font-bold tracking-tight text-gray-900 flex items-center gap-3">
          <ShieldCheck className="w-8 h-8 text-blue-600" />
          Vessel Reconciliation Suite
        </h1>
        <p className="text-sm text-gray-500 mt-2">Enterprise Client-Side Zero-Trust Auditor</p>
      </header>

      <main className="max-w-6xl mx-auto">
        {/* Placeholder for the Ingestion Drag & Drop Component */}
        {!auditResults && !isProcessing && (
          <div className="border-2 border-dashed border-gray-300 rounded-xl p-16 text-center bg-white shadow-sm">
             <h2 className="text-xl font-medium mb-2">Drop TEC-001.xlsx Here</h2>
             <p className="text-gray-500">Local processing only. Data never leaves this device.</p>
             {/* In a full build, your DragDropZone.tsx component goes here and calls useAuditStore().ingestData() */}
          </div>
        )}

        {/* The Triage Dashboard (Rendered when results exist) */}
        {auditResults && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 flex justify-between items-center">
               <div>
                 <h2 className="text-lg font-semibold text-green-700 flex items-center gap-2">
                   <ShieldCheck className="w-5 h-5" />
                   Cryptographic Seal Verified
                 </h2>
                 <p className="text-xs font-mono text-gray-500 mt-1">{digitalSeal}</p>
               </div>
               <button 
                 onClick={handleExport}
                 className="bg-blue-600 text-white px-6 py-2 rounded-lg font-medium hover:bg-blue-700 transition flex items-center gap-2"
               >
                 <Download className="w-4 h-4" />
                 Export Locked Baseline
               </button>
            </div>

            <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
              <table className="w-full text-left text-sm">
                <thead className="bg-gray-50 text-gray-600 border-b">
                  <tr>
                    <th className="px-6 py-4 font-medium">Component</th>
                    <th className="px-6 py-4 font-medium">Legacy Claim</th>
                    <th className="px-6 py-4 font-medium">Verified Math</th>
                    <th className="px-6 py-4 font-medium">Delta</th>
                    <th className="px-6 py-4 font-medium">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {auditResults.map((res, i) => (
                    <tr key={i} className={res.isCompliant ? '' : 'bg-red-50/50'}>
                      <td className="px-6 py-4 font-medium">{res.componentName}</td>
                      <td className="px-6 py-4 text-gray-500">{res.legacyHours}</td>
                      <td className="px-6 py-4 font-bold">{res.verifiedHours}</td>
                      <td className={`px-6 py-4 font-mono ${res.delta > 0 ? 'text-red-600' : 'text-gray-500'}`}>
                        {res.delta > 0 ? `+${res.delta}` : res.delta}
                      </td>
                      <td className="px-6 py-4">
                        {res.isCompliant 
                          ? <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">Verified</span>
                          : <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800">Overwrite</span>
                        }
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
