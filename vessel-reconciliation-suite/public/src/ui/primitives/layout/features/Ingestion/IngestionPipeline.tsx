import { useRef, useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, FileCheck2, X, AlertCircle, ChevronRight, Loader2 } from 'lucide-react';
import { useAuditStore }  from '../../../store/useAuditStore';
import { parsePMSFile, parseLogFile } from '../../../core/utils/excelParser';

function DropZone({ label, sub, accept, multi, onFiles, loading, loadingMsg, names, onRemove, color }: {
  label: string; sub: string; accept: string; multi: boolean;
  onFiles: (f: File[]) => Promise<void>; loading: boolean; loadingMsg: string;
  names: string[]; onRemove?: (n: string) => void; color: 'blue' | 'violet';
}) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  const has = names.length > 0;

  const handle = useCallback(async (files: File[]) => { if (files.length) await onFiles(files); }, [onFiles]);

  const accent = color === 'blue'
    ? { border: has ? 'border-blue-400' : drag ? 'border-blue-400' : 'border-slate-300', bg: has ? 'bg-blue-50/30' : drag ? 'bg-blue-50' : 'bg-slate-50', icon: 'bg-blue-100', iconColor: 'text-blue-600', chip: 'text-blue-700 bg-blue-50 border-blue-200', num: 'text-blue-700 bg-blue-50 border-blue-200' }
    : { border: has ? 'border-violet-400' : drag ? 'border-violet-400' : 'border-slate-300', bg: has ? 'bg-violet-50/30' : drag ? 'bg-violet-50' : 'bg-slate-50', icon: 'bg-violet-100', iconColor: 'text-violet-600', chip: 'text-violet-700 bg-violet-50 border-violet-200', num: 'text-violet-700 bg-violet-50 border-violet-200' };

  return (
    <div className="flex flex-col gap-2.5">
      <div
        className={`border-2 border-dashed rounded-xl p-6 cursor-pointer transition-all duration-200 ${accent.border} ${accent.bg}`}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={async e => { e.preventDefault(); setDrag(false); await handle(Array.from(e.dataTransfer.files)); }}
        onClick={() => ref.current?.click()}
      >
        <input ref={ref} type="file" accept={accept} multiple={multi} className="hidden"
          onChange={async e => { await handle(Array.from(e.target.files ?? [])); if (ref.current) ref.current.value = ''; }} />
        <div className="flex flex-col items-center gap-3 text-center pointer-events-none">
          {loading ? (
            <><Loader2 className="w-8 h-8 text-blue-500 animate-spin" /><p className="text-sm font-medium text-blue-700">{loadingMsg}</p></>
          ) : has ? (
            <><FileCheck2 className={`w-8 h-8 ${accent.iconColor}`} />
            <div><p className="text-sm font-semibold text-slate-700">{label}</p>
            <p className={`text-xs mt-0.5 ${accent.iconColor}`}>{names.length} file{names.length > 1 ? 's' : ''} loaded · drop to add more</p></div></>
          ) : (
            <><div className={`w-12 h-12 rounded-xl flex items-center justify-center ${accent.icon}`}>
              <Upload className={`w-5 h-5 ${accent.iconColor}`} />
            </div>
            <div><p className="text-sm font-semibold text-slate-700">{label}</p><p className="text-xs text-slate-500 mt-0.5">{sub}</p></div></>
          )}
        </div>
      </div>
      {has && (
        <div className="flex flex-col gap-1.5">
          {names.map(n => (
            <div key={n} className="flex items-center gap-2 bg-white border border-slate-200 rounded-lg px-3 py-2 shadow-sm">
              <FileSpreadsheet className={`w-4 h-4 flex-shrink-0 ${accent.iconColor}`} />
              <span className="text-xs font-medium text-slate-700 truncate flex-1">{n}</span>
              {onRemove && <button onClick={e => { e.stopPropagation(); onRemove(n); }} className="text-slate-400 hover:text-red-500 transition-colors flex-shrink-0"><X className="w-3.5 h-3.5" /></button>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function IngestionPipeline() {
  const { components, logFileNames, setComponents, addLogSet, removeLogFile, runAuditPipeline, isProcessing, error, vesselName, setVesselName } = useAuditStore();
  const [pmsName, setPmsName]     = useState<string[]>([]);
  const [pmsErr, setPmsErr]       = useState<string | null>(null);
  const [logErr, setLogErr]       = useState<string | null>(null);
  const [pmsLoad, setPmsLoad]     = useState(false);
  const [logLoad, setLogLoad]     = useState(false);

  const handlePMS = useCallback(async (files: File[]) => {
    setPmsErr(null); setPmsLoad(true);
    try {
      const comps = await parsePMSFile(files[0]);
      setComponents(comps); setPmsName([files[0].name]);
    } catch (e: unknown) { setPmsErr(e instanceof Error ? e.message : String(e)); setPmsName([]); }
    setPmsLoad(false);
  }, [setComponents]);

  const handleLogs = useCallback(async (files: File[]) => {
    setLogErr(null); setLogLoad(true);
    const errs: string[] = [];
    for (const f of files) {
      try { addLogSet(await parseLogFile(f), f.name); }
      catch (e: unknown) { errs.push(`${f.name}: ${e instanceof Error ? e.message : String(e)}`); }
    }
    if (errs.length) setLogErr(errs.join('\n'));
    setLogLoad(false);
  }, [addLogSet]);

  const ready = components.length > 0 && logFileNames.length > 0 && !isProcessing;

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
        <label className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">Vessel Name</label>
        <input value={vesselName} onChange={e => setVesselName(e.target.value)} placeholder="e.g. M/V Minoan Falcon"
          className="w-full sm:w-80 text-sm font-medium text-slate-900 border border-slate-300 rounded-lg px-4 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition" />
      </div>

      <div className="grid md:grid-cols-2 gap-6">
        <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-4">
          <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
            <div className="w-7 h-7 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">1</div>
            <div><h3 className="text-sm font-bold text-slate-900">PMS Baseline (TEC-001)</h3><p className="text-xs text-slate-500">Master planned maintenance sheet</p></div>
          </div>
          <DropZone label="Upload PMS / TEC-001 File" sub=".xlsx or .xls · Single file" accept=".xlsx,.xls" multi={false} onFiles={handlePMS} loading={pmsLoad} loadingMsg="Parsing PMS…" names={pmsName} color="blue" />
          {pmsErr && <div className="flex gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-xs text-red-700"><AlertCircle className="w-4 h-4 text-red-500 flex-shrink-0 mt-0.5" /><p>{pmsErr}</p></div>}
          {components.length > 0 && <p className="text-xs font-medium text-blue-700 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2">✓ {components.length} components extracted</p>}
        </div>

        <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-4">
          <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
            <div className="w-7 h-7 bg-violet-600 text-white rounded-full flex items-center justify-center text-xs font-bold">2</div>
            <div><h3 className="text-sm font-bold text-slate-900">Operating Hours Logs</h3><p className="text-xs text-slate-500">Monthly running-hours reports — multi-select</p></div>
          </div>
          <DropZone label="Upload Log File(s)" sub=".xlsx or .xls · Multiple files supported" accept=".xlsx,.xls" multi={true} onFiles={handleLogs} loading={logLoad} loadingMsg="Parsing log files…" names={logFileNames} onRemove={removeLogFile} color="violet" />
          {logErr && <div className="flex gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-xs text-red-700"><AlertCircle className="w-4 h-4 text-red-500 flex-shrink-0 mt-0.5" /><pre className="whitespace-pre-wrap font-sans">{logErr}</pre></div>}
          {logFileNames.length > 0 && <p className="text-xs font-medium text-violet-700 bg-violet-50 border border-violet-200 rounded-lg px-3 py-2">✓ {logFileNames.length} file{logFileNames.length > 1 ? 's' : ''} stitched into master timeline</p>}
        </div>
      </div>

      {error && <div className="flex gap-3 p-4 bg-red-50 border border-red-200 rounded-xl text-sm text-red-700"><AlertCircle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" /><div><p className="font-semibold">Pipeline Error</p><p className="mt-1">{error}</p></div></div>}

      <div className="flex items-center justify-between border-t border-slate-100 pt-5">
        <p className="text-xs text-slate-400">
          {ready ? <span className="text-emerald-600 font-medium">✓ Ready — {components.length} components × {logFileNames.length} log file{logFileNames.length > 1 ? 's' : ''}</span> : 'Upload both a PMS file and at least one log file to proceed.'}
        </p>
        <button onClick={runAuditPipeline} disabled={!ready}
          className={`flex items-center gap-2 px-6 py-3 rounded-xl font-semibold text-sm transition-all duration-200 shadow-sm ${ready ? 'bg-blue-600 text-white hover:bg-blue-700 hover:shadow-md active:scale-95' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}>
          {isProcessing ? <><Loader2 className="w-4 h-4 animate-spin" />Running Audit…</> : <>Run Forensic Audit<ChevronRight className="w-4 h-4" /></>}
        </button>
      </div>
    </div>
  );
}
