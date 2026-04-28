import * as XLSX from 'xlsx';
import { DailyLog, ComponentData } from '../schemas/tec001.schema';

// Helper to read file as ArrayBuffer locally
const readFileAsync = (file: File): Promise<ArrayBuffer> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target?.result as ArrayBuffer);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
};

export const parsePMSFile = async (file: File): Promise<ComponentData[]> => {
  const data = await readFileAsync(file);
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  
  // Strict Schema Check 1: Does the PMS sheet exist?
  const sheetName = workbook.SheetNames.find(n => n.includes('PMS'));
  if (!sheetName) throw new Error("Invalid File: Could not locate 'PMS' tab.");

  const worksheet = workbook.Sheets[sheetName];
  // Convert to JSON, assuming headers are on row 8 (skiprows=7 in pandas)
  const rawData = XLSX.utils.sheet_to_json<any>(worksheet, { range: 7, defval: null });

  const components: ComponentData[] = [];
  
  for (const row of rawData) {
    // Dynamically map based on your specific TEC-001 columns
    // Assuming Column 1 is Item Name, Column 5 is Date, Column 7 is Legacy Hours
    const keys = Object.keys(row);
    if (keys.length < 8) continue; 

    const name = row[keys[1]];
    const lastOhDate = row[keys[5]];
    const legacyHrs = row[keys[7]];

    // Only ingest rows that look like actual components (have a valid date)
    if (name && lastOhDate instanceof Date) {
      components.push({
        componentName: String(name).trim(),
        lastOverhaulDate: lastOhDate,
        legacyClaimedHours: Number(legacyHrs) || 0,
        // For simplicity in this build, assigning all to Main Engine
        // In production, you would parse the parent system from the header rows
        parentSystem: "MAIN_ENGINE" 
      });
    }
  }

  if (components.length === 0) throw new Error("Schema Error: No valid component dates found in PMS tab.");
  return components;
};

export const parseDailyLogsFile = async (file: File): Promise<DailyLog[]> => {
  const data = await readFileAsync(file);
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  
  // Check for Daily Hours sheet
  const sheetName = workbook.SheetNames.find(n => n.includes('DAILY OPERATING HOURS') || n.includes('DAILY'));
  if (!sheetName) throw new Error("Invalid File: Could not locate 'DAILY OPERATING HOURS' tab.");

  const worksheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json<any>(worksheet, { range: 10, defval: null });

  const dailyLogs: DailyLog[] = [];

  for (const row of rawData) {
    const keys = Object.keys(row);
    const dateVal = row[keys[0]]; // Assuming DATE is strictly Column A
    
    if (dateVal instanceof Date) {
      dailyLogs.push({
        date: dateVal,
        mainEngineHours: Number(row[keys[1]]) || 0, // MAIN ENGINE Column B
        dg1Hours: Number(row[keys[3]]) || 0,        // Mapping to your specific DG columns
        dg2Hours: Number(row[keys[5]]) || 0,
        dg3Hours: Number(row[keys[7]]) || 0,
      });
    }
  }

  if (dailyLogs.length === 0) throw new Error("Schema Error: No valid chronological dates found in Daily Logs.");
  return dailyLogs;
};
