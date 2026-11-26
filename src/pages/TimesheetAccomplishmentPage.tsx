import React, { useMemo, useState, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { ArrowLeft, FolderUp, Search, X, Clock, FileSpreadsheet, RotateCcw, Calendar } from 'lucide-react';

// Pre-load CSV files dari src/timesheet/<Bulan>/*.csv
const builtinTimesheetCsv = import.meta.glob('../timesheet/**/*.csv', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

const MONTH_ORDER = [
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

const STORAGE_KEY = 'timesheet_accomplishment_uploaded_v1';

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
  regular: Record<string, RegularEntry>; // key = normalized namaPersonil
  overtime: Record<string, OvertimeEntry>; // key = normalized namaPersonil
}

type DetailModalState =
  | {
      scope: 'person';
      month: string;
      namaPersonil: string;
      department: string;
      kind: 'overtime';
      accomplishment: number;
      totalOvertimeHours: number;
      totalTimesheetHours: number;
    }
  | {
      scope: 'person';
      month: string;
      namaPersonil: string;
      department: string;
      kind: 'regular';
      accomplishment: number;
      expectedTotalTimesheet: number;
      totalOfTimesheet: number;
    }
  | {
      scope: 'month';
      month: string;
      kind: 'overtime';
      count: number;
      avgAccomplishment: number;
      totalOvertimeHours: number;
      totalTimesheetHours: number;
    }
  | {
      scope: 'month';
      month: string;
      kind: 'regular';
      count: number;
      avgAccomplishment: number;
      expectedTotalTimesheet: number;
      totalOfTimesheet: number;
    };

function normKey(s: string): string {
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

function formatNum(n: number): string {
  return (Math.round(n * 100) / 100).toLocaleString('id-ID', {
    maximumFractionDigits: 2,
  });
}

function formatPct(n: number): string {
  const rounded = Math.round(n * 100) / 100;
  return `${rounded}%`;
}

function detectMonthName(pathParts: string[]): string | null {
  // Cari dari subfolder (bukan nama file paling akhir jika ada folder)
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

function sortMonths(months: string[]): string[] {
  return [...months].sort((a, b) => {
    const ia = MONTH_ORDER.indexOf(a);
    const ib = MONTH_ORDER.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.localeCompare(b);
  });
}

function parseSheetRows(
  rows: unknown[][],
  fileName: string,
  target: MonthTimesheetData
): boolean {
  if (!rows || rows.length < 2) return false;

  // Cari baris header (cek 5 baris pertama)
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
    const key = normKey(nama);
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

function buildBuiltinData(): Record<string, MonthTimesheetData> {
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
  return result;
}

interface Props {
  onBack: () => void;
}

export default function TimesheetAccomplishmentPage({ onBack }: Props) {
  const builtinData = useMemo(() => buildBuiltinData(), []);

  const [uploadedData, setUploadedData] = useState<Record<string, MonthTimesheetData>>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? (JSON.parse(saved) as Record<string, MonthTimesheetData>) : {};
    } catch {
      return {};
    }
  });

  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDept, setSelectedDept] = useState<string>('ALL');
  const [selectedMonthFilter, setSelectedMonthFilter] = useState<string>('ALL');
  const [detailModal, setDetailModal] = useState<DetailModalState | null>(null);

  // Gabungkan data bawaan + data folder yang diupload (folder upload menimpa bulan yang sama jika ada)
  const monthMap = useMemo(() => {
    return { ...builtinData, ...uploadedData };
  }, [builtinData, uploadedData]);

  const months = useMemo(() => sortMonths(Object.keys(monthMap)), [monthMap]);

  const visibleMonths = useMemo(() => {
    if (selectedMonthFilter === 'ALL') return months;
    return months.filter(m => m === selectedMonthFilter);
  }, [months, selectedMonthFilter]);

  // Daftar seluruh personil unik lintas bulan
  const { personRows, departments } = useMemo(() => {
    const people = new Map<
      string,
      {
        key: string;
        namaPersonil: string;
        department: string;
        byMonth: Record<
          string,
          {
            regular?: RegularEntry;
            overtime?: OvertimeEntry;
          }
        >;
      }
    >();
    const deptSet = new Set<string>();

    months.forEach(m => {
      const data = monthMap[m];
      if (!data) return;

      Object.entries(data.regular).forEach(([key, entry]) => {
        if (entry.department) deptSet.add(entry.department);
        const p = people.get(key) || {
          key,
          namaPersonil: entry.namaPersonil,
          department: entry.department,
          byMonth: {},
        };
        if (!p.department && entry.department) p.department = entry.department;
        p.byMonth[m] = { ...p.byMonth[m], regular: entry };
        people.set(key, p);
      });

      Object.entries(data.overtime).forEach(([key, entry]) => {
        if (entry.department) deptSet.add(entry.department);
        const p = people.get(key) || {
          key,
          namaPersonil: entry.namaPersonil,
          department: entry.department,
          byMonth: {},
        };
        if (!p.department && entry.department) p.department = entry.department;
        p.byMonth[m] = { ...p.byMonth[m], overtime: entry };
        people.set(key, p);
      });
    });

    const list = Array.from(people.values()).sort((a, b) => {
      const deptCmp = (a.department || '').localeCompare(b.department || '');
      if (deptCmp !== 0) return deptCmp;
      return a.namaPersonil.localeCompare(b.namaPersonil);
    });

    return {
      personRows: list,
      departments: Array.from(deptSet).sort((a, b) => a.localeCompare(b)),
    };
  }, [months, monthMap]);

  const filteredPersons = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return personRows.filter(p => {
      if (selectedDept !== 'ALL' && p.department !== selectedDept) return false;
      if (q && !p.namaPersonil.toLowerCase().includes(q) && !p.department.toLowerCase().includes(q)) {
        return false;
      }
      // Jika filter bulan dipilih, hanya tampilkan personil yang punya data di bulan tersebut
      if (selectedMonthFilter !== 'ALL') {
        const mData = p.byMonth[selectedMonthFilter];
        if (!mData?.regular && !mData?.overtime) return false;
      }
      return true;
    });
  }, [personRows, searchQuery, selectedDept, selectedMonthFilter]);

  // Ringkasan per bulan (untuk kartu Accomplishment per bulan di atas tabel)
  const monthSummaries = useMemo(() => {
    return visibleMonths.map(m => {
      const data = monthMap[m];
      const regList = Object.values(data?.regular || {});
      const otList = Object.values(data?.overtime || {});

      const regExp = regList.reduce((acc, r) => acc + r.expectedTotalTimesheet, 0);
      const regTot = regList.reduce((acc, r) => acc + r.totalOfTimesheet, 0);
      const regAvgAcc =
        regList.length > 0
          ? regList.reduce((acc, r) => acc + r.accomplishment, 0) / regList.length
          : 0;

      const otHours = otList.reduce((acc, r) => acc + r.totalOvertimeHours, 0);
      const otTsHours = otList.reduce((acc, r) => acc + r.totalTimesheetHours, 0);
      const otAvgAcc =
        otList.length > 0
          ? otList.reduce((acc, r) => acc + r.accomplishment, 0) / otList.length
          : 0;

      return {
        month: m,
        fileCount: data?.fileCount || 0,
        regCount: regList.length,
        regAvgAcc,
        regExp,
        regTot,
        otCount: otList.length,
        otAvgAcc,
        otHours,
        otTsHours,
      };
    });
  }, [visibleMonths, monthMap]);

  const handleFolderUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsUploading(true);
    setUploadStatus(null);
    try {
      const parsed: Record<string, MonthTimesheetData> = {};
      let processedFiles = 0;

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const lowerName = file.name.toLowerCase();
        if (!lowerName.endsWith('.csv') && !lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
          continue;
        }

        const relPath = (file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name;
        const parts = relPath.split('/').filter(Boolean);
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

      const addedMonths = sortMonths(Object.keys(parsed));
      if (addedMonths.length === 0) {
        setUploadStatus(
          'Tidak ditemukan file Excel/CSV yang valid di dalam folder bulan. Pastikan nama subfolder adalah nama bulan (mis. Agustus, Juli, dll).'
        );
      } else {
        const nextUploaded = { ...uploadedData, ...parsed };
        setUploadedData(nextUploaded);
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(nextUploaded));
        } catch {
          // Abaikan jika melebihi kuota localStorage
        }
        setUploadStatus(
          `Berhasil memuat ${processedFiles} file dari ${addedMonths.length} folder bulan (${addedMonths.join(', ')}).`
        );
      }
    } catch (err) {
      console.error('Error reading timesheet folder:', err);
      setUploadStatus('Gagal membaca folder timesheet. Pastikan format file CSV/Excel valid.');
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  const handleResetUploaded = () => {
    setUploadedData({});
    setUploadStatus('Data upload dibersihkan. Kembali menampilkan data bawaan.');
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {}
  };

  const accBadgeClass = (val: number) => {
    if (val >= 95) return 'bg-emerald-500/15 text-emerald-300 border-emerald-500/30 hover:bg-emerald-500/25';
    if (val >= 75) return 'bg-amber-500/15 text-amber-300 border-amber-500/30 hover:bg-amber-500/25';
    return 'bg-rose-500/15 text-rose-300 border-rose-500/30 hover:bg-rose-500/25';
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/95 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <button
              onClick={onBack}
              className="px-3 py-1.5 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 flex items-center gap-1.5 cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali
            </button>
            <div className="flex items-center gap-2">
              <Clock className="w-4 h-4 text-emerald-400" />
              <span className="font-bold text-sm text-white">ACCOMPLISHMENT TIMESHEET PER BULAN</span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {Object.keys(uploadedData).length > 0 && (
              <button
                onClick={handleResetUploaded}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1.5 cursor-pointer"
                title="Hapus data folder yang diupload dan kembali ke data bawaan"
              >
                <RotateCcw className="w-3.5 h-3.5 text-slate-400" /> Reset Upload
              </button>
            )}
            <label
              className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer transition shadow-sm"
              title="Pilih folder utama yang berisi subfolder bulan-bulan (Januari, Agustus, dll.) beserta file Excel/CSV di dalamnya"
            >
              <FolderUp className="w-4 h-4" />
              <span>{isUploading ? 'Membaca Folder...' : 'Input Folder Bulan (Excel/CSV)'}</span>
              <input
                type="file"
                {...({ webkitdirectory: '', directory: '' } as React.InputHTMLAttributes<HTMLInputElement>)}
                multiple
                onChange={handleFolderUpload}
                disabled={isUploading}
                className="hidden"
              />
            </label>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6 space-y-6">
        {uploadStatus && (
          <div className="bg-emerald-500/10 border border-emerald-500/30 text-emerald-300 text-xs rounded-xl px-4 py-3 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 shrink-0" />
              <span>{uploadStatus}</span>
            </div>
            <button onClick={() => setUploadStatus(null)} className="text-emerald-300/70 hover:text-white cursor-pointer">
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Ringkasan Accomplishment (%) per Bulan */}
        <section className="space-y-3">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div>
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <Calendar className="w-4 h-4 text-blue-400" />
                Ringkasan Accomplishment (%) Tiap Bulan
              </h2>
              <p className="text-[11px] text-slate-400">
                Klik nilai <span className="text-slate-200 font-semibold">Accomplishment (%)</span> di kartu bulan atau pada tabel di bawah untuk melihat rincian jam/timesheet.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {monthSummaries.map(ms => (
              <div
                key={ms.month}
                className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-3"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-sm text-white">{ms.month}</span>
                  <span className="text-[10px] font-mono text-slate-400 bg-slate-800 px-2 py-0.5 rounded">
                    {ms.fileCount} file
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  {/* Tidak Overtime / Reguler */}
                  <button
                    type="button"
                    disabled={ms.regCount === 0}
                    onClick={() =>
                      setDetailModal({
                        scope: 'month',
                        month: ms.month,
                        kind: 'regular',
                        count: ms.regCount,
                        avgAccomplishment: ms.regAvgAcc,
                        expectedTotalTimesheet: ms.regExp,
                        totalOfTimesheet: ms.regTot,
                      })
                    }
                    className={`p-2.5 rounded-lg border text-left transition ${
                      ms.regCount > 0
                        ? `${accBadgeClass(ms.regAvgAcc)} cursor-pointer`
                        : 'bg-slate-950/50 border-slate-800 text-slate-600 cursor-not-allowed'
                    }`}
                  >
                    <div className="text-[10px] uppercase tracking-wide opacity-80">Tidak Overtime</div>
                    <div className="text-lg font-bold font-mono mt-0.5">
                      {ms.regCount > 0 ? formatPct(ms.regAvgAcc) : '-'}
                    </div>
                    <div className="text-[10px] opacity-75 mt-0.5">{ms.regCount} personil</div>
                  </button>

                  {/* Overtime */}
                  <button
                    type="button"
                    disabled={ms.otCount === 0}
                    onClick={() =>
                      setDetailModal({
                        scope: 'month',
                        month: ms.month,
                        kind: 'overtime',
                        count: ms.otCount,
                        avgAccomplishment: ms.otAvgAcc,
                        totalOvertimeHours: ms.otHours,
                        totalTimesheetHours: ms.otTsHours,
                      })
                    }
                    className={`p-2.5 rounded-lg border text-left transition ${
                      ms.otCount > 0
                        ? `${accBadgeClass(ms.otAvgAcc)} cursor-pointer`
                        : 'bg-slate-950/50 border-slate-800 text-slate-600 cursor-not-allowed'
                    }`}
                  >
                    <div className="text-[10px] uppercase tracking-wide opacity-80">Overtime</div>
                    <div className="text-lg font-bold font-mono mt-0.5">
                      {ms.otCount > 0 ? formatPct(ms.otAvgAcc) : '-'}
                    </div>
                    <div className="text-[10px] opacity-75 mt-0.5">{ms.otCount} personil</div>
                  </button>
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* Filter & Pencarian */}
        <section className="flex items-center justify-between gap-3 flex-wrap bg-slate-900 border border-slate-800 rounded-xl p-3">
          <div className="flex items-center gap-2 flex-wrap">
            <div className="relative w-64">
              <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
              <input
                value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
                placeholder="Cari Nama Personil atau Departemen..."
                className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-emerald-500"
              />
            </div>

            <select
              value={selectedDept}
              onChange={e => setSelectedDept(e.target.value)}
              className="px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-emerald-500"
            >
              <option value="ALL">Semua Departemen</option>
              {departments.map(d => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>

            <select
              value={selectedMonthFilter}
              onChange={e => setSelectedMonthFilter(e.target.value)}
              className="px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-emerald-500"
            >
              <option value="ALL">Semua Bulan ({months.length})</option>
              {months.map(m => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          </div>

          <span className="text-[11px] text-slate-400">
            Menampilkan <b className="text-white">{filteredPersons.length}</b> personil
          </span>
        </section>

        {/* Tabel Accomplishment (%) per Personil per Bulan */}
        <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-auto max-h-[calc(100vh-20rem)]">
          <table className="w-full min-w-max text-left text-xs border-collapse">
            <thead>
              <tr className="text-slate-300 uppercase text-[11px]">
                <th
                  rowSpan={2}
                  className="sticky top-0 left-0 z-30 bg-slate-900 border-b border-r border-slate-800 py-2.5 px-3 w-10 text-center"
                >
                  No
                </th>
                <th
                  rowSpan={2}
                  className="sticky top-0 left-10 z-30 bg-slate-900 border-b border-r border-slate-800 py-2.5 px-3 min-w-[220px]"
                >
                  Nama Personil
                </th>
                <th
                  rowSpan={2}
                  className="sticky top-0 z-20 bg-slate-900 border-b border-r border-slate-800 py-2.5 px-3 min-w-[200px]"
                >
                  Departemen
                </th>
                {visibleMonths.map(m => (
                  <th
                    key={m}
                    colSpan={2}
                    className="sticky top-0 z-20 bg-slate-900 border-b border-r border-slate-800 py-2 px-3 text-center text-blue-300 font-bold"
                  >
                    {m}
                  </th>
                ))}
              </tr>
              <tr className="text-slate-400 text-[10px] uppercase">
                {visibleMonths.map(m => (
                  <React.Fragment key={`${m}-sub`}>
                    <th className="sticky top-[33px] z-20 bg-slate-900 border-b border-r border-slate-800 py-1.5 px-2.5 text-center whitespace-nowrap">
                      Tidak Overtime (%)
                    </th>
                    <th className="sticky top-[33px] z-20 bg-slate-900 border-b border-r border-slate-800 py-1.5 px-2.5 text-center whitespace-nowrap">
                      Overtime (%)
                    </th>
                  </React.Fragment>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800">
              {filteredPersons.length === 0 ? (
                <tr>
                  <td colSpan={3 + visibleMonths.length * 2} className="py-10 text-center text-slate-500">
                    Tidak ada data yang cocok.
                  </td>
                </tr>
              ) : (
                filteredPersons.map((p, idx) => (
                  <tr key={p.key} className="group hover:bg-slate-800/40">
                    <td className="sticky left-0 z-10 bg-slate-900 group-hover:bg-slate-800 border-r border-slate-800 py-2.5 px-3 text-center font-mono text-slate-500">
                      {idx + 1}
                    </td>
                    <td className="sticky left-10 z-10 bg-slate-900 group-hover:bg-slate-800 border-r border-slate-800 py-2.5 px-3 font-semibold text-white">
                      {p.namaPersonil}
                    </td>
                    <td className="border-r border-slate-800 py-2.5 px-3 text-slate-400">
                      {p.department || '-'}
                    </td>
                    {visibleMonths.map(m => {
                      const md = p.byMonth[m];
                      const reg = md?.regular;
                      const ot = md?.overtime;
                      return (
                        <React.Fragment key={`${p.key}-${m}`}>
                          <td className="border-r border-slate-800 py-2 px-2.5 text-center">
                            {reg ? (
                              <button
                                type="button"
                                onClick={() =>
                                  setDetailModal({
                                    scope: 'person',
                                    month: m,
                                    namaPersonil: p.namaPersonil,
                                    department: p.department,
                                    kind: 'regular',
                                    accomplishment: reg.accomplishment,
                                    expectedTotalTimesheet: reg.expectedTotalTimesheet,
                                    totalOfTimesheet: reg.totalOfTimesheet,
                                  })
                                }
                                title="Klik untuk melihat Expected Total of Timesheet & Total of Timesheet"
                                className={`px-2 py-0.5 rounded border font-mono font-semibold text-xs cursor-pointer transition ${accBadgeClass(reg.accomplishment)}`}
                              >
                                {formatPct(reg.accomplishment)}
                              </button>
                            ) : (
                              <span className="text-slate-600">-</span>
                            )}
                          </td>
                          <td className="border-r border-slate-800 py-2 px-2.5 text-center">
                            {ot ? (
                              <button
                                type="button"
                                onClick={() =>
                                  setDetailModal({
                                    scope: 'person',
                                    month: m,
                                    namaPersonil: p.namaPersonil,
                                    department: p.department,
                                    kind: 'overtime',
                                    accomplishment: ot.accomplishment,
                                    totalOvertimeHours: ot.totalOvertimeHours,
                                    totalTimesheetHours: ot.totalTimesheetHours,
                                  })
                                }
                                title="Klik untuk melihat Total Overtime Hours & Total Timesheet Hours"
                                className={`px-2 py-0.5 rounded border font-mono font-semibold text-xs cursor-pointer transition ${accBadgeClass(ot.accomplishment)}`}
                              >
                                {formatPct(ot.accomplishment)}
                              </button>
                            ) : (
                              <span className="text-slate-600">-</span>
                            )}
                          </td>
                        </React.Fragment>
                      );
                    })}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </main>

      {/* Modal Detail Saat Nilai Accomplishment (%) Dikelik */}
      {detailModal && (
        <div
          className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4"
          onClick={() => setDetailModal(null)}
        >
          <div
            className="bg-slate-900 border border-slate-700 rounded-2xl max-w-md w-full p-5 space-y-4 shadow-2xl"
            onClick={e => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span
                    className={`px-2 py-0.5 text-[10px] font-bold uppercase rounded border ${
                      detailModal.kind === 'overtime'
                        ? 'bg-purple-500/15 text-purple-300 border-purple-500/30'
                        : 'bg-sky-500/15 text-sky-300 border-sky-500/30'
                    }`}
                  >
                    {detailModal.kind === 'overtime' ? 'Overtime' : 'Tidak Overtime'}
                  </span>
                  <span className="text-xs font-semibold text-blue-400">Bulan {detailModal.month}</span>
                </div>
                <h3 className="text-base font-bold text-white mt-1.5">
                  {detailModal.scope === 'person'
                    ? detailModal.namaPersonil
                    : `Total Keseluruhan Bulan ${detailModal.month}`}
                </h3>
                {detailModal.scope === 'person' ? (
                  <p className="text-xs text-slate-400">{detailModal.department}</p>
                ) : (
                  <p className="text-xs text-slate-400">{detailModal.count} personil tercatat</p>
                )}
              </div>
              <button
                onClick={() => setDetailModal(null)}
                className="p-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 space-y-3">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2.5">
                <span className="text-xs text-slate-400">
                  {detailModal.scope === 'month' ? 'Rata-rata Accomplishment (%)' : 'Accomplishment (%)'}
                </span>
                <span className="text-lg font-bold font-mono text-emerald-400">
                  {formatPct(
                    detailModal.scope === 'person'
                      ? detailModal.accomplishment
                      : detailModal.avgAccomplishment
                  )}
                </span>
              </div>

              {detailModal.kind === 'overtime' ? (
                <>
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-slate-300">Total Overtime Hours</span>
                    <span className="text-sm font-bold font-mono text-white">
                      {formatNum(detailModal.totalOvertimeHours)} jam
                    </span>
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-slate-300">Total Timesheet Hours</span>
                    <span className="text-sm font-bold font-mono text-purple-300">
                      {formatNum(detailModal.totalTimesheetHours)} jam
                    </span>
                  </div>
                </>
              ) : (
                <>
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-slate-300">Expected Total of Timesheet</span>
                    <span className="text-sm font-bold font-mono text-white">
                      {formatNum(detailModal.expectedTotalTimesheet)} hari
                    </span>
                  </div>
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-slate-300">Total of Timesheet</span>
                    <span className="text-sm font-bold font-mono text-sky-300">
                      {formatNum(detailModal.totalOfTimesheet)} hari
                    </span>
                  </div>
                </>
              )}
            </div>

            <div className="flex justify-end">
              <button
                onClick={() => setDetailModal(null)}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg cursor-pointer"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
