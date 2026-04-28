// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Root Application Component
// ═══════════════════════════════════════════════════════════════════════════

import { useAuditStore }     from './store/useAuditStore';
import IngestionPipeline     from './ui/features/Ingestion/IngestionPipeline';
import { truncateSeal }      from './core/crypto/hash';
import {
  ShieldCheck, AlertTriangle, RefreshCcw,
  FileText, Table, Info, TrendingUp, CheckCircle2, XCircle, Clock,
} from 'lucide-react';

// ─── Badge ────────────────────────────────────────────────────────────────────

function Badge({ label, variant }: { label: string; variant: 'ok' | 'drift' | 'low' | 'medium' | 'high' }) {
  const styles: Record<string, string> = {
    ok:     'bg-emerald-100 text-emerald-800 border-emerald-200',
    drift:  'bg-red-100    text-red-800    border-red-200',
    low:    'bg-red-100    text-red-800    border-red-200',
    medium: 'bg-amber-100  text-amber-800  border-amber-200',
    high:   'bg-emerald-100 text-emerald-800 border-emerald-200',
  };
  return (
    <span className={`px-2.5 py-1 rounded-full text-xs font-semibold border ${styles[variant]}`}>
      {label}
    </span>
  );
}

// ─── Stat Card ────────────────────────────────────────────────────────────────

function StatCard({
  label, value, icon: Icon, accent,
}: { label: string; value: string | number; icon: React.ElementType; accent: string }) {
  return (
    <div className={`bg-white rounded-xl border shadow-sm p-5 border-t-4 ${accent}`}>
      <div className="flex items-start justify-between">
        <div>
          <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{label}</p>
          <p className="text-3xl font-bold text-slate-900 mt-2 font-mono">{value}</p>
        </div>
        <Icon className="w-5 h-5 text-slate-400 mt-1" />
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const {
    auditResults, physicsViolations, digitalSeal, summary,
    error, resetStore, exportPDF, exportExcel, vesselName,
  } = useAuditStore();

  const hasResults = auditResults !== null && digitalSeal !== null && summary !== null;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">

      {/* ── Navigation Header ──────────────────────────────────────────────── */}
      <header className="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 bg-blue-600 rounded-xl flex items-center justify-center shadow-sm">
              <ShieldCheck className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-base font-bold text-slate-900 leading-tight">
                Vessel Reconciliation Suite
              </h1>
              <p className="text-xs text-slate-500 leading-tight">
                {vesselName} · Zero-Trust Forensic Auditor
              </p>
            </div>
          </div>

          {(hasResults || error) && (
            <button
              onClick={resetStore}
              className="flex items-center gap-2 text-sm font-medium text-slate-600 hover:text-slate-900 bg-white border border-slate-300 px-4 py-2 rounded-lg shadow-sm hover:shadow transition-all"
            >
              <RefreshCcw className="w-4 h-4" />
              New Audit
            </button>
          )}
        </div>
      </header>

      {/* ── Main Content ───────────────────────────────────────────────────── */}
      <main className="max-w-6xl mx-auto px-4 sm:px-6 py-8 space-y-8">

        {/* Error state */}
        {error && !hasResults && (
          <div className="p-5 bg-red-50 border border-red-200 rounded-xl flex items-start gap-4">
            <AlertTriangle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
            <div>
              <p className="font-semibold text-red-800">Pipeline Error</p>
              <p className="text-sm text-red-700 mt-1">{error}</p>
            </div>
          </div>
        )}

        {/* Ingestion state */}
        {!hasResults && (
          <section>
            <div className="mb-6">
              <h2 className="text-xl font-bold text-slate-900">Audit Setup</h2>
              <p className="text-sm text-slate-500 mt-1">
                Upload your PMS baseline and operating hours logs to begin the forensic reconciliation.
              </p>
            </div>
            <IngestionPipeline />
          </section>
        )}

        {/* Results state */}
        {hasResults && summary && (
          <div className="space-y-8">

            {/* ── KPI Row ─────────────────────────────────────────────────── */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <StatCard
                label="Components Audited"
                value={summary.totalComponents}
                icon={TrendingUp}
                accent={summary.totalComponents > 0 ? 'border-t-blue-500' : 'border-t-slate-300'}
              />
              <StatCard
                label="Drift Anomalies"
                value={summary.driftComponents}
                icon={summary.driftComponents > 0 ? XCircle : CheckCircle2}
                accent={summary.driftComponents > 0 ? 'border-t-red-500' : 'border-t-emerald-500'}
              />
              <StatCard
                label="Days in Timeline"
                value={summary.totalDaysCovered}
                icon={Clock}
                accent="border-t-violet-500"
              />
              <StatCard
                label="Physics Violations"
                value={summary.physicsViolationCount}
                icon={AlertTriangle}
                accent={summary.physicsViolationCount > 0 ? 'border-t-amber-500' : 'border-t-emerald-500'}
              />
            </div>

            {/* ── Timeline Info Bar ────────────────────────────────────────── */}
            <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-wrap gap-4 text-xs shadow-sm">
              <span className="text-slate-500">
                <span className="font-semibold text-slate-700">Timeline:</span>{' '}
                {summary.earliestLog?.toLocaleDateString('en-GB')} → {summary.latestLog?.toLocaleDateString('en-GB')}
              </span>
              <span className="text-slate-400">·</span>
              <span className="text-slate-500">
                <span className="font-semibold text-slate-700">Source files:</span> {summary.sourceFileCount}
              </span>
              <span className="text-slate-400">·</span>
              <span className={summary.timelineGaps.length > 0 ? 'text-amber-600 font-semibold' : 'text-slate-500'}>
                {summary.timelineGaps.length} chronological gap{summary.timelineGaps.length !== 1 ? 's' : ''}{' '}
                {summary.timelineGaps.length > 0 && `(${summary.timelineGaps.reduce((s, g) => s + g.missingDays, 0)} missing days)`}
              </span>
              <span className="text-slate-400">·</span>
              <span className="text-slate-500 font-mono">
                SHA-256: {truncateSeal(digitalSeal!, 16, 8)}
              </span>
            </div>

            {/* ── Phase 1: Physics Violations ───────────────────────────────── */}
            {physicsViolations.length > 0 && (
              <section className="bg-amber-50 border border-amber-200 rounded-xl overflow-hidden shadow-sm">
                <div className="px-6 py-4 border-b border-amber-200 flex items-center gap-3 bg-amber-100/50">
                  <AlertTriangle className="w-5 h-5 text-amber-600 flex-shrink-0" />
                  <div>
                    <h3 className="font-bold text-amber-900">Phase 1 — Kinematic Violations</h3>
                    <p className="text-xs text-amber-700 mt-0.5">
                      {physicsViolations.length} impossible value{physicsViolations.length > 1 ? 's' : ''} detected in daily logs.
                      These entries are included in the audit arithmetic with their raw values.
                    </p>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm text-left">
                    <thead>
                      <tr className="text-xs text-amber-800 bg-amber-100/40 border-b border-amber-200">
                        <th className="px-6 py-3 font-semibold">Date</th>
                        <th className="px-6 py-3 font-semibold">System</th>
                        <th className="px-6 py-3 font-semibold">Logged Hours</th>
                        <th className="px-6 py-3 font-semibold">Max Allowed</th>
                        <th className="px-6 py-3 font-semibold">Reason</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-amber-100">
                      {physicsViolations.map((v, i) => (
                        <tr key={i} className="text-amber-900 hover:bg-amber-50">
                          <td className="px-6 py-3 font-mono text-xs">{v.date.toLocaleDateString('en-GB')}</td>
                          <td className="px-6 py-3 font-medium">{v.system.replace('_', ' ')}</td>
                          <td className="px-6 py-3 font-bold text-red-700 font-mono">{v.loggedHours.toFixed(2)} h</td>
                          <td className="px-6 py-3 font-mono">{v.maxAllowed} h</td>
                          <td className="px-6 py-3 text-xs">{v.reason}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            )}

            {/* ── Phase 2: Audit Results ────────────────────────────────────── */}
            <section className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
              <div className="px-6 py-5 border-b border-slate-100 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                <div>
                  <h3 className="font-bold text-slate-900">Phase 2 — Baseline Drift Reconciliation</h3>
                  <p className="text-xs text-slate-500 mt-1">
                    Verified hours summed from master timeline since each component's last overhaul date.
                  </p>
                </div>
                <div className="flex gap-3 flex-shrink-0">
                  <button
                    onClick={exportExcel}
                    className="flex items-center gap-2 text-sm font-semibold text-slate-700 bg-white border border-slate-300 px-4 py-2 rounded-lg hover:bg-slate-50 hover:shadow transition-all"
                  >
                    <Table className="w-4 h-4" />
                    Export Excel
                  </button>
                  <button
                    onClick={exportPDF}
                    className="flex items-center gap-2 text-sm font-semibold text-white bg-blue-600 px-4 py-2 rounded-lg hover:bg-blue-700 hover:shadow transition-all"
                  >
                    <FileText className="w-4 h-4" />
                    Class PDF Report
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-500 bg-slate-50 border-b border-slate-200">
                    <tr>
                      <th className="px-6 py-4 font-semibold">Component Name</th>
                      <th className="px-6 py-4 font-semibold">System</th>
                      <th className="px-6 py-4 font-semibold">Overhaul Date</th>
                      <th className="px-6 py-4 font-semibold">Legacy Claim</th>
                      <th className="px-6 py-4 font-semibold text-blue-700">Verified Math</th>
                      <th className="px-6 py-4 font-semibold">Delta</th>
                      <th className="px-6 py-4 font-semibold">Confidence</th>
                      <th className="px-6 py-4 font-semibold">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {auditResults!.map((res, i) => (
                      <tr
                        key={i}
                        className={`transition-colors ${
                          res.isCompliant ? 'hover:bg-slate-50' : 'bg-red-50/40 hover:bg-red-50/70'
                        }`}
                      >
                        <td className="px-6 py-4 font-medium text-slate-900">
                          <div>{res.componentName}</div>
                          {res.note && (
                            <div className="flex items-center gap-1 mt-0.5">
                              <Info className="w-3 h-3 text-amber-500 flex-shrink-0" />
                              <span className="text-xs text-amber-700 font-normal whitespace-normal leading-snug max-w-xs">
                                {res.note}
                              </span>
                            </div>
                          )}
                        </td>
                        <td className="px-6 py-4 text-slate-500 text-xs font-medium">
                          {res.parentSystem.replace('_', ' ')}
                        </td>
                        <td className="px-6 py-4 font-mono text-xs text-slate-600">
                          {res.overhaulDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}
                        </td>
                        <td className="px-6 py-4 font-mono text-slate-500">
                          {res.legacyHours.toLocaleString()} h
                        </td>
                        <td className="px-6 py-4 font-mono font-bold text-slate-900">
                          {res.verifiedHours.toLocaleString()} h
                        </td>
                        <td className={`px-6 py-4 font-mono font-bold ${
                          res.delta > 0 ? 'text-red-600' :
                          res.delta < 0 ? 'text-amber-700' :
                          'text-slate-400'
                        }`}>
                          {res.delta > 0 ? `+${res.delta}` : res.delta === 0 ? '—' : res.delta} h
                        </td>
                        <td className="px-6 py-4">
                          <Badge
                            label={res.confidence}
                            variant={res.confidence.toLowerCase() as 'high' | 'medium' | 'low'}
                          />
                        </td>
                        <td className="px-6 py-4">
                          <Badge
                            label={res.isCompliant ? 'VERIFIED' : 'DRIFT DETECTED'}
                            variant={res.isCompliant ? 'ok' : 'drift'}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Table footer legend */}
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex flex-wrap gap-4 text-xs text-slate-500">
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block" />
                  HIGH confidence: Timeline fully covers overhaul period
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-400 inline-block" />
                  MEDIUM: Gaps exist in timeline after overhaul
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-red-400 inline-block" />
                  LOW: Timeline doesn't reach back to overhaul date
                </span>
                <span className="ml-auto">
                  Tolerance: ±0.5 h (rounding)
                </span>
              </div>
            </section>

          </div>
        )}
      </main>
    </div>
  );
}
