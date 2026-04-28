import React from 'react';
import { useAuditStore } from './store/useAuditStore';
import IngestionPipeline from './ui/features/Ingestion/IngestionPipeline';
import { generateSecurePDF } from './core/export/pdfBuilder';
import { generateSecureExcel } from './core/export/excelWriter';
import { ShieldCheck, Download, RefreshCcw, AlertTriangle, FileText, Table } from 'lucide-react';

export default function App() {
  // Pull state and actions from the Zustand global store
  const { auditResults, digitalSeal, error, resetStore } = useAuditStore();

  const handleExportPDF = async () => {
    if (!auditResults || !digitalSeal) return;
    await generateSecurePDF(auditResults, digitalSeal, "MINOAN FALCON");
  };

  const handleExportExcel = () => {
    if (!auditResults || !digitalSeal) return;
    generateSecureExcel(auditResults, digitalSeal);
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans p-4 md:p-8">
      {/* --- Global Header --- */}
      <header className="max-w-6xl mx-auto mb-10 border-b border-gray-200 pb-6 flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900 flex items-center gap-3">
            <ShieldCheck className="w-8 h-8 text-blue-700" />
            Vessel Reconciliation Suite
          </h1>
          <p className="text-sm text-gray-500 mt-1 font-medium tracking-wide uppercase">
            M/V Minoan Falcon | Zero-Trust Auditor
          </p>
        </div>
        
        {/* Only show the Reset button if results exist or an error occurred */}
        {(auditResults || error) && (
          <button 
            onClick={resetStore}
            className="flex items-center gap-2 text-sm font-medium text-gray-500 hover:text-gray-900 transition-colors bg-white border border-gray-200 px-4 py-2 rounded-lg shadow-sm"
          >
            <RefreshCcw className="w-4 h-4" /> Start New Audit
          </button>
        )}
      </header>

      <main className="max-w-6xl mx-auto">
        
        {/* --- Global Error Boundary --- */}
        {error && (
          <div className="mb-8 p-6 bg-red-50 border-l-4 border-red-500 rounded-r-lg flex items-start gap-4 shadow-sm">
            <AlertTriangle className="w-6 h-6 text-red-600 flex-shrink-0 mt-0.5" />
            <div>
              <h3 className="text-red-800 font-semibold text-lg">System Error</h3>
              <p className="text-red-700 mt-1">{error}</p>
            </div>
          </div>
        )}

        {/* --- STATE 1: Data Ingestion (No Results Yet) --- */}
        {!auditResults && !error && (
          <IngestionPipeline />
        )}

        {/* --- STATE 2: The Triage Dashboard (Results Verified) --- */}
        {auditResults && digitalSeal && (
          <div className="space-y-8 animate-in fade-in duration-500">
            
            {/* Top Metrics Row */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                <h3 className="text-gray-500 text-sm font-medium">Components Audited</h3>
                <p className="text-3xl font-bold mt-2">{auditResults.length}</p>
              </div>
              <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                <h3 className="text-gray-500 text-sm font-medium">Discrepancies Corrected</h3>
                <p className="text-3xl font-bold mt-2 text-blue-700">
                  {auditResults.filter(r => !r.isCompliant).length}
                </p>
              </div>
              <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                <h3 className="text-gray-500 text-sm font-medium flex items-center gap-2">
                  <ShieldCheck className="w-4 h-4 text-green-600" /> Cryptographic Seal (SHA-256)
                </h3>
                <p className="text-sm font-mono mt-3 text-gray-700 bg-gray-50 p-2 rounded border border-gray-100 overflow-hidden text-ellipsis">
                  {digitalSeal}
                </p>
              </div>
            </div>

            {/* Export Actions Panel */}
            <div className="bg-blue-50/50 p-6 rounded-xl border border-blue-100 flex flex-col sm:flex-row items-center justify-between gap-4">
              <div>
                <h3 className="font-semibold text-blue-900">Baseline Locked & Verified</h3>
                <p className="text-sm text-blue-700 mt-1">Export the immutable Class Report or the cleaned Excel file.</p>
              </div>
              <div className="flex gap-3">
                <button 
                  onClick={handleExportExcel}
                  className="bg-white border border-gray-300 text-gray-700 px-5 py-2.5 rounded-lg font-medium hover:bg-gray-50 transition flex items-center gap-2 shadow-sm"
                >
                  <Table className="w-4 h-4" /> Export Excel
                </button>
                <button 
                  onClick={handleExportPDF}
                  className="bg-blue-600 text-white px-5 py-2.5 rounded-lg font-medium hover:bg-blue-700 transition flex items-center gap-2 shadow-sm"
                >
                  <FileText className="w-4 h-4" /> Download PDF Report
                </button>
              </div>
            </div>

            {/* Forensic Data Table */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm whitespace-nowrap">
                  <thead className="bg-gray-50 text-gray-600 border-b border-gray-200">
                    <tr>
                      <th className="px-6 py-4 font-semibold tracking-wide">Component Name</th>
                      <th className="px-6 py-4 font-semibold tracking-wide">Legacy Claim</th>
                      <th className="px-6 py-4 font-semibold tracking-wide text-blue-700">Verified Math</th>
                      <th className="px-6 py-4 font-semibold tracking-wide">Delta</th>
                      <th className="px-6 py-4 font-semibold tracking-wide">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {auditResults.map((res, i) => (
                      <tr key={i} className={res.isCompliant ? 'hover:bg-gray-50' : 'bg-red-50/30 hover:bg-red-50/60'}>
                        <td className="px-6 py-4 font-medium text-gray-900">{res.componentName}</td>
                        <td className="px-6 py-4 font-mono text-gray-500">{res.legacyHours.toLocaleString()}</td>
                        <td className="px-6 py-4 font-mono font-bold text-gray-900">{res.verifiedHours.toLocaleString()}</td>
                        <td className={`px-6 py-4 font-mono font-medium ${res.delta !== 0 ? 'text-red-600' : 'text-gray-400'}`}>
                          {res.delta > 0 ? `+${res.delta.toLocaleString()}` : res.delta.toLocaleString()}
                        </td>
                        <td className="px-6 py-4">
                          {res.isCompliant ? (
                            <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-green-100 text-green-800 border border-green-200">
                              VERIFIED
                            </span>
                          ) : (
                            <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-red-100 text-red-800 border border-red-200">
                              OVERWRITTEN
                            </span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}
      </main>
    </div>
  );
}
