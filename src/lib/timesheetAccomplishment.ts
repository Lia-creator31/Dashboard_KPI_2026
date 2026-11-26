import * as XLSX from 'xlsx';

// Pre-load CSV bawaan dari src/timesheet/<Bulan>/*.csv
const builtinTimesheetCsv = import.meta.glob('../timesheet/**/*.csv', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

export const MONTH_ORDER = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];

const MONTH_ALIASES: Record<string, string> = {
  januari: 'Januari', january: 'Januari', jan: 'Januari',
  februari: 'Februari', february: 'Februari', feb: 'Februari',
  maret: 'Maret', march: 'Maret', mar: 'Maret',
  april: 'April', apr: 'April',
  mei: 'Mei', may: 'Mei',
  juni: 'Juni', june: 'Juni', jun: 'Juni',
  juli: 'Juli', july: 'Juli', jul: 'Juli',
  agustus: 'Agustus', august: 'Agustus', agu: 'Agustus', aug: 'Agustus', agt: 'Agustus',
  september: 'September', sep: 'September', sept: 'September',
  oktober: 'Oktober', october: 'Oktober', okt: 'Oktober', oct: 'Oktober',
  november: 'November', nov: 'November',
  desember: 'Desember', december: 'Desember', des: 'Desember', dec: 'Desember',
};

export const TIMESHEET_STORAGE_KEY = 'timesheet_accomplishment_uploaded_v1';

export interface RegularEntry {
  kind: 'regular';
  division: string;
  department: string;
  namaPersonil: string;
  expectedTotalTimesheet: number;
  totalOfTimesheet: number;
  accomplishment: number;
}

export interface OvertimeEntry {
  kind: 'overtime';
  division: string;
  department: string;
  namaPersonil: string;
  totalOvertimeHours: number;
  totalTimesheetHours: number;
  accomplishment: number;
}

export interface MonthTimesheetData {
  month: string;
  fileCount: number;
  regular: Record<string, RegularEntry>; // key = cleanText(namaPersonil)
  overtime: Record<string, OvertimeEntry>; // key = cleanText(namaPersonil)
}

export function cleanPersonKey(s: string): string {
  return (s || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

function parseNum(v: unknown): number {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const cleaned = String(v ?? '')
    .trim()
    .replace(/%/g, '')
    .replace(/,/g, '.');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

export function detectMonthName(pathParts: string[]): string | null {
  const folders = pathParts.length > 1 ? pathParts.slice(0, -1) : pathParts;
  for (let i = folders.length - 1; i >= 0; i--) {
    const raw = folders[i].trim().toLowerCase();
    if (MONTH_ALIASES[raw]) return MONTH_ALIASES[raw];
    for (const [alias, canonical] of Object.entries(MONTH_ALIASES)) {
      if (raw.includes(alias)) return canonical;
    }
  }
  return null;
}

export function sortMonths(months: string[]): string[] {
  return [...months].sort((a, b) => {
    const ia = MONTH_ORDER.indexOf(a);
    const ib = MONTH_ORDER.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.localeCompare(b);
  });
}

export function parseSheetRows(
  rows: unknown[][],
  fileName: string,
  target: MonthTimesheetData
): boolean {
  if (!rows || rows.length < 2) return false;

  let headerIdx = -1;
  let headers: string[] = [];
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    const row = (rows[r] || []).map(c => String(c ?? '').trim().toLowerCase());
    if (row.some(c => c.includes('nama personil') || c.includes('accomplishment'))) {
      headerIdx = r;
      headers = row;
      break;
    }
  }
  if (headerIdx === -1) return false;

  const idxDiv = headers.findIndex(h => h === 'division' || h.includes('divisi'));
  const idxDept = headers.findIndex(h => h === 'department' || h.includes('departemen'));
  const idxNama = headers.findIndex(h => h.includes('nama personil') || h === 'nama' || h.includes('personil'));
  const idxAcc = headers.findIndex(h => h.includes('accomplishment'));

  const idxOtHours = headers.findIndex(h => h.includes('total overtime hours'));
  const idxTsHours = headers.findIndex(h => h.includes('total timesheet hours'));

  const idxExpected = headers.findIndex(h => h.includes('expected total of timesheet'));
  const idxTotalTs = headers.findIndex(h => h === 'total of timesheet' || (h.includes('total of timesheet') && !h.includes('expected')));

  if (idxNama === -1 || idxAcc === -1) return false;

  const isOvertime =
    idxOtHours !== -1 ||
    idxTsHours !== -1 ||
    (fileName.toLowerCase().includes('overtime') && idxExpected === -1);

  let parsedAny = false;
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const nama = String(row[idxNama] ?? '').trim();
    if (!nama) continue;

    const division = idxDiv >= 0 ? String(row[idxDiv] ?? '').trim() : '';
    const department = idxDept >= 0 ? String(row[idxDept] ?? '').trim() : '';
    const accomplishment = parseNum(row[idxAcc]);
    const key = cleanPersonKey(nama);
    if (!key) continue;

    parsedAny = true;
    if (isOvertime) {
      const otHours = idxOtHours >= 0 ? parseNum(row[idxOtHours]) : 0;
      const tsHours = idxTsHours >= 0 ? parseNum(row[idxTsHours]) : 0;
      const existing = target.overtime[key];
      if (existing) {
        const sumOt = existing.totalOvertimeHours + otHours;
        const sumTs = existing.totalTimesheetHours + tsHours;
        existing.totalOvertimeHours = sumOt;
        existing.totalTimesheetHours = sumTs;
        existing.accomplishment =
          sumOt > 0 ? Math.min(100, Math.round((sumTs / sumOt) * 1000) / 10) : accomplishment;
      } else {
        target.overtime[key] = {
          kind: 'overtime',
          division,
          department,
          namaPersonil: nama,
          totalOvertimeHours: otHours,
          totalTimesheetHours: tsHours,
          accomplishment,
        };
      }
    } else {
      const expected = idxExpected >= 0 ? parseNum(row[idxExpected]) : 0;
      const totalTs = idxTotalTs >= 0 ? parseNum(row[idxTotalTs]) : 0;
      const existing = target.regular[key];
      if (existing) {
        const sumExp = existing.expectedTotalTimesheet + expected;
        const sumTot = existing.totalOfTimesheet + totalTs;
        existing.expectedTotalTimesheet = sumExp;
        existing.totalOfTimesheet = sumTot;
        existing.accomplishment =
          sumExp > 0 ? Math.min(100, Math.round((sumTot / sumExp) * 1000) / 10) : accomplishment;
      } else {
        target.regular[key] = {
          kind: 'regular',
          division,
          department,
          namaPersonil: nama,
          expectedTotalTimesheet: expected,
          totalOfTimesheet: totalTs,
          accomplishment,
        };
      }
    }
  }

  if (parsedAny) {
    target.fileCount += 1;
  }
  return parsedAny;
}

let cachedBuiltin: Record<string, MonthTimesheetData> | null = null;

export function getBuiltinTimesheetData(): Record<string, MonthTimesheetData> {
  if (cachedBuiltin) return cachedBuiltin;
  const result: Record<string, MonthTimesheetData> = {};
  Object.entries(builtinTimesheetCsv).forEach(([path, raw]) => {
    const parts = path.split('/').filter(Boolean);
    const month = detectMonthName(parts);
    if (!month) return;

    if (!result[month]) {
      result[month] = { month, fileCount: 0, regular: {}, overtime: {} };
    }
    const wb = XLSX.read(raw, { type: 'string' });
    const firstSheet = wb.SheetNames[0];
    if (!firstSheet) return;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[firstSheet], { header: 1, defval: '', raw: false }) as unknown[][];
    parseSheetRows(rows, parts[parts.length - 1] || '', result[month]);
  });
  cachedBuiltin = result;
  return result;
}

export function getStoredUploadedTimesheets(): Record<string, MonthTimesheetData> {
  try {
    const raw = localStorage.getItem(TIMESHEET_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as Record<string, MonthTimesheetData>) : {};
  } catch {
    return {};
  }
}

export function getAllTimesheetData(): Record<string, MonthTimesheetData> {
  const builtin = getBuiltinTimesheetData();
  const uploaded = getStoredUploadedTimesheets();
  return { ...builtin, ...uploaded };
}

export async function parseUploadedTimesheetFolder(
  files: FileList | File[]
): Promise<{ addedMonths: string[]; processedFiles: number; nextUploaded: Record<string, MonthTimesheetData> }> {
  const parsed: Record<string, MonthTimesheetData> = {};
  let processedFiles = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.csv') && !lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
      continue;
    }

    const relPath = (file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name;
    const parts = relPath.split(/[\/\\]/).filter(Boolean);
    const month = detectMonthName(parts);
    if (!month) continue;

    if (!parsed[month]) {
      parsed[month] = { month, fileCount: 0, regular: {}, overtime: {} };
    }

    let wb: XLSX.WorkBook;
    if (lowerName.endsWith('.csv')) {
      const text = await file.text();
      wb = XLSX.read(text, { type: 'string' });
    } else {
      const buf = await file.arrayBuffer();
      wb = XLSX.read(buf, { type: 'array' });
    }

    let fileParsed = false;
    wb.SheetNames.forEach(sheetName => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false }) as unknown[][];
      if (parseSheetRows(rows, file.name, parsed[month])) {
        fileParsed = true;
      }
    });
    if (fileParsed) processedFiles += 1;
  }

  const existingUploaded = getStoredUploadedTimesheets();
  const nextUploaded = { ...existingUploaded, ...parsed };
  try {
    localStorage.setItem(TIMESHEET_STORAGE_KEY, JSON.stringify(nextUploaded));
  } catch (e) {
    console.warn('Failed to persist timesheet to localStorage:', e);
  }

  return {
    addedMonths: sortMonths(Object.keys(parsed)),
    processedFiles,
    nextUploaded,
  };
}

export interface PersonMonthAccomplishment {
  month: string;
  regular?: RegularEntry;
  overtime?: OvertimeEntry;
}

export function getPersonMonthlyTimesheet(
  personName: string,
  allData: Record<string, MonthTimesheetData>
): PersonMonthAccomplishment[] {
  const key = cleanPersonKey(personName);
  if (!key) return [];

  const months = sortMonths(Object.keys(allData));
  const result: PersonMonthAccomplishment[] = [];

  months.forEach(m => {
    const mData = allData[m];
    if (!mData) return;

    // Cari dengan key atau kecocokan substring
    let reg = mData.regular[key];
    if (!reg) {
      const foundKey = Object.keys(mData.regular).find(k => k.includes(key) || key.includes(k));
      if (foundKey) reg = mData.regular[foundKey];
    }

    let ot = mData.overtime[key];
    if (!ot) {
      const foundKey = Object.keys(mData.overtime).find(k => k.includes(key) || key.includes(k));
      if (foundKey) ot = mData.overtime[foundKey];
    }

    if (reg || ot) {
      result.push({
        month: m,
        regular: reg,
        overtime: ot,
      });
    }
  });

  return result;
}
