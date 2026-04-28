import * as XLSX from 'xlsx';
import { DailyLog, ComponentData } from '../schemas/tec001.schema';

const readFileAsync = (file: File): Promise<ArrayBuffer> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target?.result as ArrayBuffer);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
};

export const processMasterWorkbook = async (file: File): Promise<{ components: ComponentData[], dailyLogs: DailyLog[] }> => {
  const data = await readFileAsync(file);
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });

  // --- 1. Find Sheets ---
  const pmsSheetName = workbook.SheetNames.find(n => n.includes('PMS'));
  const logsSheetName = workbook.SheetNames.find(n => n.includes('DAILY') || n.includes('OPERATING'));

  if (!pmsSheetName || !logsSheetName) {
    throw new Error("Invalid Workbook: Must contain both 'PMS' and 'DAILY OPERATING HOURS' tabs.");
  }

  // --- 2. Extract PMS Data ---
  const pmsRaw = XLSX.utils.sheet_to_json<any>(workbook.Sheets[pmsSheetName], { range: 7, defval: null });
  const components: ComponentData[] = [];
  
  for (const row of pmsRaw) {
    const keys = Object.keys(row);
    if (keys.length < 8) continue; 
    const name = row[keys[1]];
    const lastOhDate = row[keys[5]];
    const legacyHrs = row[keys[7]];

    if (name && lastOhDate instanceof Date) {
      components.push({
        componentName: String(name).trim(),
        lastOverhaulDate: lastOhDate,
        legacyClaimedHours: Number(legacyHrs) || 0,
        parentSystem: "MAIN_ENGINE" 
      });
    }
  }

  // --- 3. Extract Daily Logs Data ---
  const logsRaw = XLSX.utils.sheet_to_json<any>(workbook.Sheets[logsSheetName], { range: 10, defval: null });
  const dailyLogs: DailyLog[] = [];

  for (const row of logsRaw) {
    const keys = Object.keys(row);
    const dateVal = row[keys[0]]; 
    
    if (dateVal instanceof Date) {
      dailyLogs.push({
        date: dateVal,
        mainEngineHours: Number(row[keys[1]]) || 0,
        dg1Hours: Number(row[keys[3]]) || 0,
        dg2Hours: Number(row[keys[5]]) || 0,
        dg3Hours: Number(row[keys[7]]) || 0,
      });
    }
  }

  if (components.length === 0) throw new Error("No valid overhaul dates found in PMS tab.");
  if (dailyLogs.length === 0) throw new Error("No chronological dates found in Daily Logs tab.");

  return { components, dailyLogs };
};
