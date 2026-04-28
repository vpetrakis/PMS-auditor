import React, { useRef } from 'react';
import { UploadCloud, ShieldAlert, Cpu } from 'lucide-react';
import { processMasterWorkbook } from '../../../core/utils/excelParser';
import { useAuditStore } from '../../../store/useAuditStore';

export default function IngestionPipeline() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { processFileAndAudit, isProcessing, error } = useAuditStore();

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      // Extract both tabs simultaneously
      const { components, dailyLogs } = await processMasterWorkbook(file);
      // Run the bi-directional audit immediately
      await processFileAndAudit(components, dailyLogs);
    } catch (err: any) {
      // Error handled by store/App
    }
  };

  return (
    <div className="max-w-2xl mx-auto mt-12 animate-in fade-in slide-in-from-bottom-4 duration-700">
      <div className="text-center space-y-2 mb-8">
        <h2 className="text-2xl font-bold tracking-tight text-gray-900">Master Workbook Ingestion</h2>
        <p className="text-sm text-gray-500 font-medium">
          Upload TEC-001. System will automatically extract PMS and Daily Logs.
        </p>
      </div>

      <div 
        onClick={() => !isProcessing && fileInputRef.current?.click()}
        className={`relative rounded-xl border-2 border-dashed p-12 transition-all duration-300 text-center shadow-sm
          ${isProcessing ? 'border-blue-300 bg-blue-50 cursor-wait' : 'border-gray-300 hover:border-blue-500 bg-white hover:bg-gray-50 cursor-pointer'}`}
      >
        <input 
          type="file" 
          accept=".xlsx, .xls" 
          className="hidden" 
          ref={fileInputRef} 
          onChange={handleFileUpload} 
          disabled={isProcessing}
        />
        
        {isProcessing ? (
          <div className="flex flex-col items-center justify-center space-y-4 text-blue-700">
            <Cpu className="w-12 h-12 animate-pulse" />
            <h3 className="text-xl font-semibold">Executing Bi-Directional Audit...</h3>
            <p className="text-sm">Cross-referencing timeline constraints against overhaul dates.</p>
          </div>
        ) : (
          <div className="flex flex-col items-center justify-center space-y-4 text-gray-600">
            <UploadCloud className="w-12 h-12 text-gray-400 group-hover:text-blue-500 transition-colors" />
            <h3 className="text-xl font-semibold text-gray-900">Drop TEC-001.xlsx Here</h3>
            <p className="text-sm text-gray-500 max-w-sm">
              Zero-Trust Local Processing. Data is calculated securely in your browser and never uploaded.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
