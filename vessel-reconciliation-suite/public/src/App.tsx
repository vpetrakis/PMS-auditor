import React from 'react';
import { useAuditStore } from './store/useAuditStore';
import IngestionPipeline from './ui/features/Ingestion/IngestionPipeline';
import { generateSecurePDF } from './core/export/pdfBuilder';
import { generateSecureExcel } from './core/export/excelWriter';
import { ShieldCheck, Download, RefreshCcw, AlertTriangle, FileText, Table } from 'lucide-react';

export default function App() {
  const { auditResults, physicsViolations, digitalSeal, error, resetStore } = useAuditStore();

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
      {/* Header */}
      <header className="max-w-5xl mx-auto mb-10 border-b border-gray-200 pb-6 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900 flex items-center gap-3">
            <ShieldCheck className="w-8 h-8 text-blue-700" />
            Vessel Reconciliation Suite
          </h1>
          <p className="text-sm text-gray-500 mt-1 font-medium tracking-wide uppercase">
            M/V Minoan Falcon | Zero-Trust Auditor
          </p>
        </div>
        {(auditResults || error) && (
          <button onClick={resetStore} className="flex items-center gap-2 text-sm font-medium text-gray-600 hover:text-gray-900 bg-white border border-gray-300 px-4 py-2 rounded-lg shadow-sm">
            <RefreshCcw className="w-4 h-4" /> Start New Audit
          </button>
        )}
      </header>

      <main className="max-w-5xl mx-auto">
        {/* Global Error Boundary */}
        {error && (
          <div className="mb-8 p-6 bg-red-50 border-l-4 border-red-500 rounded-lg flex items-start gap-4 shadow-sm">
            <AlertTriangle className="w-6 h-6 text-red-600 flex-shrink-0" />
            <div>
              <h3 className="text-red-800 font-semibold text-lg">System Error</h3>
              <p className="text-red-700 mt-1">{error}</p>
            </div>
          </div>
        )}

        {/* State 1: Ingestion */}
        {!auditResults && !error && <IngestionPipeline />}

        {/* State 2: Results Dashboard */}
        {auditResults && digitalSeal && (
          <div className="space-y-8 animate-in fade-in duration-500">
            
            {/* Phase 1: Daily Logs Physics Check */}
            {physicsViolations.length > 0 && (
              <div className="bg-orange-50 border border-orange-200 rounded-xl overflow-hidden shadow-sm">
                <div className="p-4 bg-orange-100/50 border-b border-orange-200 flex items-center gap-3">
                  <AlertTriangle className="w-5 h-5 text-orange-600" />
                  <h3 className="font-semibold text-orange-900">Phase 1 Warning: Physics Violations Detected in Daily Logs</h3>
                </div>
                <div className="p-0 overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-orange-100/30 text-orange-800">
                      <tr><th className="px-6 py-3">Date</th><th className="px-6 py-3">System</th><th className="px-6 py-3">Logged Hours</th><th className="px-6 py-3">Error Reason</th></tr>
                    </thead>
                    <tbody className="divide-y divide-orange-100">
                      {physicsViolations.map((v, i) => (
                        <tr key={i} className="text-orange-900">
                          <td className="px-6 py-3">{v.date.toLocaleDateString()}</td>
                          <td className="px-6 py-3 font-medium">{v.system}</td>
                          <td className="px-6 py-3 font-bold text-red-600">{v.loggedHours}</td>
                          <td className="px-6 py-3">{v.reason}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Phase 2: PMS Math Check */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
              <div className="p-6 border-b border-gray-200 flex flex-col sm:flex-row items-center justify-between gap-4 bg-gray-50/50">
                <div>
                  <h3 className="font-semibold text-gray-900">Phase 2: PMS Baseline Drift Reconciliation</h3>
                  <p className="text-sm text-gray-500 mt-1">Cryptographic Hash: <span className="font-mono text-xs bg-gray-200 px-2 py-1 rounded text-gray-700">{digitalSeal}</span></p>
                </div>
                <div className="flex gap-3">
                  <button onClick={handleExportExcel} className="bg-white border border-gray-300 text-gray-700 px-4 py-2 rounded-lg font-medium hover:bg-gray-50 flex items-center gap-2 text-sm shadow-sm">
                    <Table className="w-4 h-4" /> Export Locked Excel
                  </button>
                  <button onClick={handleExportPDF} className="bg-blue-600 text-white px-4 py-2 rounded-lg font-medium hover:bg-blue-700 flex items-center gap-2 text-sm shadow-sm">
                    <FileText className="w-4 h-4" /> Class PDF Report
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm whitespace-nowrap">
                  <thead className="bg-white text-gray-500 border-b border-gray-200">
                    <tr>
                      <th className="px-6 py-4 font-semibold">Component Name</th>
                      <th className="px-6 py-4 font-semibold">Legacy Claim</th>
                      <th className="px-6 py-4 font-semibold text-blue-700">Verified Math</th>
                      <th className="px-6 py-4 font-semibold">Delta</th>
                      <th className="px-6 py-4 font-semibold">Status</th>
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
                            <span className="px-2.5 py-1 rounded-full text-xs font-medium bg-green-100 text-green-800 border border-green-200">VERIFIED</span>
                          ) : (
                            <span className="px-2.5 py-1 rounded-full text-xs font-medium bg-red-100 text-red-800 border border-red-200">OVERWRITTEN</span>
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
