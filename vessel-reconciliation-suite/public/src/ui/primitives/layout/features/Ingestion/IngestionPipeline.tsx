import React, { useState, useRef } from 'react';
import { UploadCloud, CheckCircle2, AlertCircle, Cpu } from 'lucide-react';
import { parsePMSFile, parseDailyLogsFile } from '../../../core/utils/excelParser';
import { useAuditStore } from '../../../store/useAuditStore';

export default function IngestionPipeline() {
  const [pmsStatus, setPmsStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [pmsMsg, setPmsMsg] = useState('Awaiting TEC-001 Upload');
  
  const [logsStatus, setLogsStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [logsMsg, setLogsMsg] = useState('Awaiting Daily Logs Upload');

  const [isProcessing, setIsProcessing] = useState(false);
  const { ingestData, runAuditPipeline } = useAuditStore();

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>, type: 'PMS' | 'LOGS') => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (type === 'PMS') {
      try {
        const components = await parsePMSFile(file);
        // Temporarily store in window/ref until both are ready
        (window as any).tempComponents = components; 
        setPmsStatus('success');
        setPmsMsg(`File Accepted: ${components.length} components verified.`);
      } catch (err: any) {
        setPmsStatus('error');
        setPmsMsg(err.message);
      }
    } else {
      try {
        const logs = await parseDailyLogsFile(file);
        (window as any).tempLogs = logs;
        setLogsStatus('success');
        setLogsMsg(`File Accepted: Timeline of ${logs.length} days extracted.`);
      } catch (err: any) {
        setLogsStatus('error');
        setLogsMsg(err.message);
      }
    }
  };

  const executePipeline = async () => {
    setIsProcessing(true);
    
    // 1. Move validated data from temp memory to Zustand Global Store
    ingestData((window as any).tempLogs, (window as any).tempComponents);
    
    // 2. Add artificial delay for UX (shows the user heavy processing is occurring)
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    // 3. Trigger the Math & Physics Engine
    await runAuditPipeline();
  };

  const isReady = pmsStatus === 'success' && logsStatus === 'success';

  return (
    <div className="max-w-3xl mx-auto space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
      
      {/* Header */}
      <div className="text-center space-y-2 mb-10">
        <h2 className="text-2xl font-bold tracking-tight text-gray-900">Data Ingestion Pipeline</h2>
        <p className="text-sm text-gray-500 font-medium flex items-center justify-center gap-2">
          <ShieldIcon /> Zero-Trust Local Processing: No data leaves this device.
        </p>
      </div>

      {/* Upload Zone A: PMS */}
      <DropZone 
        title="1. Planned Maintenance System (PMS)"
        description="Drop the TEC-001.xlsx file containing the component Overhaul Dates."
        status={pmsStatus}
        message={pmsMsg}
        onUpload={(e) => handleFileUpload(e, 'PMS')}
      />

      {/* Upload Zone B: Daily Logs */}
      <DropZone 
        title="2. Daily Running Hours Timeline"
        description="Drop the Daily Logs file containing the continuous 24h operational history."
        status={logsStatus}
        message={logsMsg}
        onUpload={(e) => handleFileUpload(e, 'LOGS')}
      />

      {/* Execution Action */}
      <div className="pt-6 border-t border-gray-200 flex justify-end">
        <button
          disabled={!isReady || isProcessing}
          onClick={executePipeline}
          className={`px-8 py-3 rounded-lg font-semibold flex items-center gap-3 transition-all ${
            isReady && !isProcessing
              ? 'bg-blue-600 text-white hover:bg-blue-700 shadow-md hover:shadow-lg'
              : 'bg-gray-100 text-gray-400 cursor-not-allowed'
          }`}
        >
          {isProcessing ? (
            <>
              <Cpu className="w-5 h-5 animate-pulse" />
              Executing Forensic Audit...
            </>
          ) : (
            <>
              Initialize Forensic Engine
            </>
          )}
        </button>
      </div>
    </div>
  );
}

// --- Sub-Component: Reusable Drop Zone ---
function DropZone({ title, description, status, message, onUpload }: any) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const getStatusStyles = () => {
    if (status === 'success') return 'border-green-500 bg-green-50';
    if (status === 'error') return 'border-red-400 bg-red-50';
    return 'border-gray-300 hover:border-blue-400 bg-white hover:bg-gray-50 border-dashed';
  };

  return (
    <div 
      onClick={() => fileInputRef.current?.click()}
      className={`relative rounded-xl border-2 p-8 transition-all duration-300 cursor-pointer shadow-sm ${getStatusStyles()}`}
    >
      <input 
        type="file" 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
        ref={fileInputRef} 
        onChange={onUpload} 
      />
      
      <div className="flex items-start gap-5">
        <div className="mt-1">
          {status === 'idle' && <UploadCloud className="w-8 h-8 text-gray-400" />}
          {status === 'success' && <CheckCircle2 className="w-8 h-8 text-green-600" />}
          {status === 'error' && <AlertCircle className="w-8 h-8 text-red-500" />}
        </div>
        <div>
          <h3 className={`font-semibold text-lg ${status === 'success' ? 'text-green-900' : 'text-gray-900'}`}>
            {title}
          </h3>
          <p className="text-gray-500 text-sm mt-1">{description}</p>
          
          <div className="mt-4 inline-flex items-center px-3 py-1 rounded-full text-xs font-mono font-medium bg-white/60 border border-black/5">
             <span className={`${status === 'success' ? 'text-green-700' : status === 'error' ? 'text-red-600' : 'text-gray-500'}`}>
               {message}
             </span>
          </div>
        </div>
      </div>
    </div>
  );
}

function ShieldIcon() {
  return (
    <svg className="w-4 h-4 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
    </svg>
  );
}
