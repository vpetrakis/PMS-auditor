import type { DailyLog }     from '../schemas/tec001.schema';
import type { TimelineGap, SystemLabel } from '../../types/vessel';

export interface MasterTimeline {
  logs:        DailyLog[];
  gaps:        TimelineGap[];
  totalDays:   number;
  startDate:   Date | null;
  endDate:     Date | null;
  sourceFiles: number;
}

function dateKey(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

export function stitchTimeline(sets: DailyLog[][], sourceCount: number): MasterTimeline {
  const map = new Map<string, DailyLog>();
  for (const set of sets) for (const log of set) map.set(dateKey(log.date), log);

  const sorted = [...map.values()].sort((a, b) => a.date.getTime() - b.date.getTime());
  if (!sorted.length) return { logs: [], gaps: [], totalDays: 0, startDate: null, endDate: null, sourceFiles: sourceCount };

  const gaps: TimelineGap[] = [];
  for (let i = 1; i < sorted.length; i++) {
    const diff = Math.round((sorted[i].date.getTime() - sorted[i-1].date.getTime()) / 86_400_000);
    if (diff > 1) gaps.push({ from: sorted[i-1].date, to: sorted[i].date, missingDays: diff - 1 });
  }

  return { logs: sorted, gaps, totalDays: sorted.length, startDate: sorted[0].date, endDate: sorted[sorted.length-1].date, sourceFiles: sourceCount };
}

export function sumHoursAfterDate(tl: MasterTimeline, after: Date, sys: SystemLabel): number {
  const ms = after.getTime();
  return tl.logs.reduce((s, l) => {
    if (l.date.getTime() < ms) return s;
    return s + (sys === 'MAIN_ENGINE' ? l.mainEngineHours : sys === 'DG1' ? l.dg1Hours : sys === 'DG2' ? l.dg2Hours : l.dg3Hours);
  }, 0);
}

export function hasGapsAfter(tl: MasterTimeline, after: Date): boolean {
  return tl.gaps.some(g => g.from.getTime() >= after.getTime());
}

export function missingDaysAfter(tl: MasterTimeline, after: Date): number {
  return tl.gaps.filter(g => g.from.getTime() >= after.getTime()).reduce((s, g) => s + g.missingDays, 0);
}
