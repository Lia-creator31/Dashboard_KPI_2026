import { useState, useEffect } from 'react';
import { departmentsData, monthList, Department } from './data';
import * as XLSX from 'xlsx';
import excelFileUrl from './data_kpi.xlsx?url';

// Import 6 File CSV Absensi
import csvJan from './absensi_januari.csv?raw';
import csvFeb from './absensi_februari.csv?raw';
import csvMar from './absensi_maret.csv?raw';
import csvApr from './absensi_april.csv?raw';
import csvMei from './absensi_mei.csv?raw';
import csvJun from './absensi_juni.csv?raw';

import { 
  Wrench, 
  CircleDollarSign, 
  Users, 
  Truck, 
  ShieldCheck, 
  Laptop, 
  ArrowLeft, 
  Search, 
  ChevronRight, 
  Building2, 
  Calendar,
  FileSpreadsheet,
  Loader2,
  LucideIcon
} from 'lucide-react';

const iconMap: Record<string, LucideIcon> = {
  Wrench,
  CircleDollarSign,
  Users,
  Truck,
  ShieldCheck,
  Laptop,
};

// Peta CSV Absensi per bulan
const csvMonthMap: Record<string, string> = {
  'januari': csvJan,
  'februari': csvFeb,
  'maret': csvMar,
  'april': csvApr,
  'mei': csvMei,
  'juni': csvJun,
};

// Pembacaan otomatis seluruh file di dalam folder src/timesheet/*/*
const timesheetGlobFiles = import.meta.glob('./timesheet/*/*.{csv,txt}', { 
  query: '?raw', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

interface ExcelRow {
  nip: string;
  nama: string;
  plannedHour: number;
  effectiveHour: number;
  overtimeHour: number;
  idleHour: number;
  timesheetHour: number;
  terlambat: number;
  sakit: number;
}

interface SelectedBiroPage {
  biroName: string;
  month: string;
  data: ExcelRow[];
}

export default function App() {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroPage, setSelectedBiroPage] = useState<SelectedBiroPage | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [tableSearch, setTableSearch] = useState('');
  
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [isLoadingExcel, setIsLoadingExcel] = useState<boolean>(true);

  // 1. Baca file Excel KPI saat web pertama dibuka
  useEffect(() => {
    async function loadExcel() {
      try {
        setIsLoadingExcel(true);
        const response = await fetch(excelFileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        setWorkbook(wb);
      } catch (error) {
        console.error("Gagal membaca file Excel:", error);
      } finally {
        setIsLoadingExcel(false);
      }
    }
    loadExcel();
  }, []);

  // Helper konversi angka aman
  const parseValToNumber = (val: any): number => {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    const str = String(val).trim().replace('%', '').replace(',', '.');
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  // Helper Parser CSV Absensi
  const parseAbsensiCSV = (csvContent: string): Map<string, { terlambat: number; sakit: number }> => {
    const absensiMap = new Map<string, { terlambat: number; sakit: number }>();
    if (!csvContent) return absensiMap;

    const lines = csvContent.split(/\r?\n/);
    lines.forEach((line, index) => {
      if (index === 0 || !line.trim()) return;

      const parts = line.split(',');
      if (parts.length >= 6) {
        const rawNama = parts[1] ? parts[1].trim() : '';
        const cleanNameKey = rawNama.toLowerCase().replace(/\s+/g, ' ').trim();

        const terlambatVal = parseValToNumber(parts[parts.length - 3] ?? parts[4]);
        const sakitVal = parseValToNumber(parts[parts.length - 2] ?? parts[5]);

        if (cleanNameKey) {
          absensiMap.set(cleanNameKey, {
            terlambat: terlambatVal,
            sakit: sakitVal
          });
        }
      }
    });

    return absensiMap;
  };

  // Helper Parser 6 Folder Timesheet (Nama: setelah koma ke-2, Nilai: setelah koma terakhir)
  const parseTimesheetFolder = (targetMonth: string): Map<string, number> => {
    const timesheetMap = new Map<string, number>();
    const cleanMonth = targetMonth.toLowerCase().trim();

    Object.entries(timesheetGlobFiles).forEach(([filePath, content]) => {
      const pathLower = filePath.toLowerCase();
      // Pastikan file berasal dari sub-folder bulan yang sedang dipilih
      if (pathLower.includes(`/timesheet/${cleanMonth}/`) || pathLower.includes(`/${cleanMonth}/`)) {
        const lines = content.split(/\r?\n/);
        lines.forEach((line, lineIdx) => {
          if (lineIdx === 0 || !line.trim()) return; // Lewati header

          const parts = line.split(',');
          // Minimal ada 3 kolom (0: Divisi, 1: Departemen, 2: Nama, ... terakhir: Nilai)
          if (parts.length >= 3) {
            // Nama berada setelah koma ke-2 dan sebelum koma ke-3 (Index 2)
            const rawNama = parts[2] ? parts[2].replace(/"/g, '').trim() : '';
            const cleanNameKey = rawNama.toLowerCase().replace(/\s+/g, ' ').trim();

            // Nilai berada setelah koma terakhir (Index terakhir)
            const rawLastVal = parts[parts.length - 1] ? parts[parts.length - 1].replace(/"/g, '').trim() : '0';
            const numVal = parseValToNumber(rawLastVal);

            if (cleanNameKey) {
              const currentTotal = timesheetMap.get(cleanNameKey) || 0;
              timesheetMap.set(cleanNameKey, currentTotal + numVal);
            }
          }
        });
      }
    });

    return timesheetMap;
  };

  // Filter daftar departemen di menu utama
  const filteredDepartments = departmentsData.filter((dept) => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    (dept.description || '').toLowerCase().includes(searchQuery.toLowerCase())
  );

  // 2. Fungsi Hitung Seluruh Metrik
  const handleMonthClick = (biroName: string, month: string) => {
    if (!workbook) {
      alert("File Excel sedang dimuat atau belum terbaca.");
      return;
    }

    const monthPrefix = month.toLowerCase().slice(0, 3);
    const targetSheetName = workbook.SheetNames.find(sheet => {
      const sLower = sheet.toLowerCase().trim();
      return sLower.includes(month.toLowerCase()) || sLower.includes(monthPrefix);
    });

    if (!targetSheetName) {
      alert(`Sheet untuk bulan "${month}" tidak ditemukan di file Excel.`);
      return;
    }

    // 1. Data Absensi
    const rawCsvData = csvMonthMap[month.toLowerCase()] || '';
    const absensiDataMap = parseAbsensiCSV(rawCsvData);

    // 2. Data Timesheet dari Folder Bulan Terkait
    const timesheetDataMap = parseTimesheetFolder(month);

    const worksheet = workbook.Sheets[targetSheetName];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const personMap = new Map<string, {
      nip: string;
      nama: string;
      groupsC: Map<string, number>;
      effectiveSum: number;
      overtimeSum: number;
      idleSum: number;
    }>();

    const cleanTargetBiro = biroName
      .toLowerCase()
      .replace(/biro/g, '')
      .replace(/[^a-z0-9]/g, '')
      .trim();

    rawRows.forEach((row, rowIndex) => {
      if (!row || row.length === 0 || rowIndex === 0) return;

      const colA_NIP   = String(row[0] || '').trim();  // Kolom A
      const colB_Nama  = String(row[1] || '').trim();  // Kolom B
      const colC_Val   = String(row[2] || 'DEFAULT_C').trim(); // Kolom C
      const colG_Val   = parseValToNumber(row[6]);     // Kolom G (Planned)
      const colH_Biro  = String(row[7] || '').trim();  // Kolom H (Biro)
      const colK_Eff   = parseValToNumber(row[10]);    // Kolom K (Effective)
      const colL_Ot    = parseValToNumber(row[11]);    // Kolom L (Overtime)
      const colM_Idle  = parseValToNumber(row[12]);    // Kolom M (Idle)

      const cleanRowBiro = colH_Biro
        .toLowerCase()
        .replace(/biro/g, '')
        .replace(/[^a-z0-9]/g, '')
        .trim();

      const isMatchingBiro = cleanRowBiro && (
        cleanRowBiro.includes(cleanTargetBiro) ||
        cleanTargetBiro.includes(cleanRowBiro)
      );

      if (isMatchingBiro && (colB_Nama || colA_NIP)) {
        const personKey = (colB_Nama || colA_NIP).toLowerCase().replace(/\s+/g, ' ').trim();

        if (!personMap.has(personKey)) {
          personMap.set(personKey, {
            nip: colA_NIP || '-',
            nama: colB_Nama || '-',
            groupsC: new Map<string, number>(),
            effectiveSum: 0,
            overtimeSum: 0,
            idleSum: 0,
          });
        }

        const person = personMap.get(personKey)!;

        // Planned Hour (Max Kolom G per Kolom C)
        const currentMaxG = person.groupsC.get(colC_Val) ?? -Infinity;
        if (colG_Val > currentMaxG) {
          person.groupsC.set(colC_Val, colG_Val);
        }

        person.effectiveSum += colK_Eff;
        person.overtimeSum += colL_Ot;
        person.idleSum += colM_Idle;
      }
    });

    // Susun data akhir tabel
    const formattedData: ExcelRow[] = [];
    personMap.forEach((person) => {
      let totalPlanned = 0;
      person.groupsC.forEach((maxValG) => {
        totalPlanned += maxValG;
      });

      const cleanName = person.nama.toLowerCase().replace(/\s+/g, ' ').trim();
      const absensi = absensiDataMap.get(cleanName);
      const timesheetTotal = timesheetDataMap.get(cleanName) || 0;

      formattedData.push({
        nip: person.nip,
        nama: person.nama,
        plannedHour: Math.round(totalPlanned * 10) / 10,
        effectiveHour: Math.round(person.effectiveSum * 10) / 10,
        overtimeHour: Math.round(person.overtimeSum * 10) / 10,
        idleHour: Math.round(person.idleSum * 10) / 10,
        timesheetHour: Math.round(timesheetTotal * 10) / 10,
        terlambat: absensi ? absensi.terlambat : 0,
        sakit: absensi ? absensi.sakit : 0,
      });
    });

    setTableSearch('');
    setSelectedBiroPage({
      biroName,
      month,
      data: formattedData
    });
  };

  const filteredTableData = (selectedBiroPage?.data || []).filter(item => 
    item.nip.toLowerCase().includes(tableSearch.toLowerCase()) ||
    item.nama.toLowerCase().includes(tableSearch.toLowerCase())
  );

  return (
    <div className="bg-slate-900 text-slate-100 min-h-screen font-sans antialiased">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 bg-slate-900/80 backdrop-blur-md border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div 
            className="flex items-center gap-3 cursor-pointer select-none" 
            onClick={() => {
              setSelectedBiroPage(null);
              setSelectedDept(null);
            }}
          >
            <div className="p-2 bg-blue-600 rounded-lg text-white">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-bold text-base tracking-wide leading-none text-white">PORTAL KPI DESAIN</h1>
              <span className="text-xs text-slate-400">Tahun Anggaran 2026</span>
            </div>
          </div>

          {selectedBiroPage ? (
            <button
              onClick={() => setSelectedBiroPage(null)}
              className="flex items-center gap-2 px-3 py-1.5 text-sm bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4" />
              Kembali ke Daftar Biro
            </button>
          ) : selectedDept ? (
            <button
              onClick={() => setSelectedDept(null)}
              className="flex items-center gap-2 px-3 py-1.5 text-sm bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4" />
              Kembali ke Menu Utama
            </button>
          ) : null}
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {/* LEVEL 1: 6 DEPARTEMEN */}
        {!selectedDept && !selectedBiroPage && (
          <div className="space-y-8 animate-fadeIn">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 border-b border-slate-800 pb-6">
              <div>
                <h2 className="text-2xl sm:text-3xl font-extrabold text-white">Dashboard KPI Desain 2026</h2>
                <p className="text-slate-400 text-sm mt-1">Pilih salah satu departemen untuk mengakses tombol bulan setiap biro.</p>
              </div>
              <div className="relative w-full md:w-72">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Cari departemen..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 bg-slate-800 border border-slate-700 rounded-lg text-sm text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {filteredDepartments.map((dept) => {
                const IconComponent = iconMap[dept.icon] || Building2;
                return (
                  <div
                    key={dept.id}
                    onClick={() => setSelectedDept(dept)}
                    className="group bg-slate-800/60 border border-slate-700/60 hover:border-blue-500/50 hover:bg-slate-800 rounded-xl p-6 cursor-pointer transition-all duration-200 hover:-translate-y-1 hover:shadow-xl hover:shadow-blue-500/10 flex flex-col justify-between"
                  >
                    <div>
                      <div className="flex items-center justify-between mb-4">
                        <div className="p-3 bg-blue-500/10 text-blue-400 border border-blue-500/20 rounded-xl group-hover:bg-blue-600 group-hover:text-white transition-colors">
                          <IconComponent className="w-6 h-6" />
                        </div>
                        {dept.code && (
                          <span className="text-xs font-mono font-semibold px-2.5 py-1 bg-slate-700/60 text-slate-300 rounded border border-slate-600">
                            {dept.code}
                          </span>
                        )}
                      </div>
                      <h3 className="text-lg font-bold text-white group-hover:text-blue-400 transition-colors">
                        {dept.name}
                      </h3>
                      {dept.description && (
                        <p className="text-slate-400 text-sm mt-2 line-clamp-2 leading-relaxed">
                          {dept.description}
                        </p>
                      )}
                    </div>

                    <div className="mt-6 pt-4 border-t border-slate-700/50 flex items-center justify-between text-xs text-slate-400">
                      <span>{dept.biros.length} Biro Terdaftar</span>
                      <span className="flex items-center gap-1 font-medium text-blue-400 group-hover:translate-x-1 transition-transform">
                        Buka Biro <ChevronRight className="w-4 h-4" />
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* LEVEL 2: DAFTAR BIRO */}
        {selectedDept && !selectedBiroPage && (
          <div className="space-y-8 animate-fadeIn">
            <div className="bg-gradient-to-r from-blue-900/40 to-slate-800 border border-blue-500/20 rounded-2xl p-6 sm:p-8">
              <button
                onClick={() => setSelectedDept(null)}
                className="inline-flex items-center gap-2 text-sm text-blue-400 hover:text-blue-300 font-medium mb-4 transition cursor-pointer"
              >
                <ArrowLeft className="w-4 h-4" /> Kembali ke Semua Departemen
              </button>
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  {selectedDept.code && (
                    <div className="inline-block px-2.5 py-0.5 mb-2 bg-blue-500/20 border border-blue-400/30 text-blue-300 rounded text-xs font-mono font-bold">
                      KODE: {selectedDept.code}
                    </div>
                  )}
                  <h2 className="text-2xl sm:text-3xl font-extrabold text-white">{selectedDept.name}</h2>
                  {selectedDept.description && (
                    <p className="text-slate-300 text-sm mt-1 max-w-2xl">{selectedDept.description}</p>
                  )}
                </div>
                <div className="px-4 py-2 bg-slate-800/80 border border-slate-700 rounded-xl text-center">
                  <div className="text-2xl font-black text-blue-400">{selectedDept.biros.length}</div>
                  <div className="text-xs text-slate-400 font-medium uppercase tracking-wider">Total Biro</div>
                </div>
              </div>
            </div>

            <div>
              <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
                <Building2 className="w-5 h-5 text-blue-400" />
                Daftar Biro & Pilihan Periode Bulan
              </h3>
              
              <div className="space-y-4">
                {selectedDept.biros.map((biro, index) => (
                  <div
                    key={biro.id}
                    className="bg-slate-800/70 border border-slate-700/80 rounded-xl p-5 sm:p-6 hover:border-slate-600 transition shadow-sm"
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-4 border-b border-slate-700/60 mb-4">
                      <div className="flex items-center gap-3">
                        <span className="text-xs font-mono font-bold px-2.5 py-1 bg-blue-500/10 text-blue-400 border border-blue-500/20 rounded">
                          Biro 0{index + 1}
                        </span>
                        <h4 className="text-base font-bold text-white">{biro.name}</h4>
                      </div>
                      <span className="text-xs text-slate-400 flex items-center gap-1.5">
                        <Calendar className="w-3.5 h-3.5 text-blue-400" /> Pilih Bulan (Jan - Jun 2026)
                      </span>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2.5">
                      {monthList.map((month) => {
                        return (
                          <button
                            key={month}
                            onClick={() => handleMonthClick(biro.name, month)}
                            disabled={isLoadingExcel}
                            className="px-4 py-2.5 rounded-lg text-xs font-semibold tracking-wide transition-all duration-150 flex items-center justify-center cursor-pointer border bg-slate-800/90 hover:bg-blue-600 hover:text-white border-slate-700 text-slate-300 hover:border-blue-500 disabled:opacity-50"
                          >
                            {isLoadingExcel ? (
                              <Loader2 className="w-3.5 h-3.5 animate-spin" />
                            ) : (
                              month
                            )}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* LEVEL 3: TABEL LENGKAP DENGAN KOLOM PENGISIAN TIMESHEET */}
        {selectedBiroPage && (
          <div className="space-y-6 animate-fadeIn">
            <div className="bg-slate-800/80 border border-slate-700 rounded-2xl p-6 sm:p-8">
              <button
                onClick={() => setSelectedBiroPage(null)}
                className="inline-flex items-center gap-2 text-sm text-blue-400 hover:text-blue-300 font-medium mb-3 transition cursor-pointer"
              >
                <ArrowLeft className="w-4 h-4" /> Kembali ke Pilihan Bulan
              </button>

              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2 text-xs font-semibold text-blue-400 uppercase tracking-wider mb-1">
                    <FileSpreadsheet className="w-4 h-4" />
                    <span>{selectedDept?.name}</span>
                  </div>
                  <h2 className="text-2xl sm:text-3xl font-extrabold text-white">
                    {selectedBiroPage.biroName}
                  </h2>
                  <p className="text-slate-400 text-sm mt-1">
                    Rekapitulasi Jam Kerja, Timesheet & Absensi — <span className="text-blue-300 font-medium">Bulan {selectedBiroPage.month} 2026</span>
                  </p>
                </div>

                <div className="px-5 py-3 bg-slate-900/80 border border-slate-700 rounded-xl text-center md:text-right">
                  <div className="text-2xl font-black text-blue-400">{filteredTableData.length}</div>
                  <div className="text-xs text-slate-400 font-medium uppercase tracking-wider">Total Pegawai</div>
                </div>
              </div>
            </div>

            {/* Pencarian */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div className="relative w-full sm:w-80">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Cari NIP atau Nama Pegawai..."
                  value={tableSearch}
                  onChange={(e) => setTableSearch(e.target.value)}
                  className="w-full pl-9 pr-4 py-2.5 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-sm"
                />
              </div>
              <span className="text-xs text-slate-400">
                Menampilkan <b>{filteredTableData.length}</b> data pegawai
              </span>
            </div>

            {/* Tabel KPI + Pengisian Timesheet + Terlambat & Sakit */}
            <div className="bg-slate-800/60 border border-slate-700/80 rounded-2xl overflow-hidden shadow-xl">
              {filteredTableData.length > 0 ? (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-sm">
                    <thead>
                      <tr className="bg-slate-800/90 text-slate-300 text-xs font-semibold uppercase tracking-wider border-b border-slate-700">
                        <th className="py-4 px-3 w-12 text-center">No</th>
                        <th className="py-4 px-4 w-36">NIP</th>
                        <th className="py-4 px-4">Nama Pegawai</th>
                        <th className="py-4 px-3 text-right">Planned</th>
                        <th className="py-4 px-3 text-right">Effective</th>
                        <th className="py-4 px-3 text-right">Overtime</th>
                        <th className="py-4 px-3 text-right">Idle</th>
                        {/* KOLOM BARU: PENGISIAN TIMESHEET */}
                        <th className="py-4 px-3 text-right text-indigo-400 bg-indigo-950/20">
                          Pengisian Timesheet
                        </th>
                        <th className="py-4 px-3 text-center text-rose-400 bg-rose-950/20">Terlambat</th>
                        <th className="py-4 px-3 text-center text-amber-400 bg-amber-950/20">Sakit</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-700/60 text-slate-200">
                      {filteredTableData.map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-700/40 transition">
                          <td className="py-3.5 px-3 text-center font-mono text-xs text-slate-400">{idx + 1}</td>
                          <td className="py-3.5 px-4 font-mono font-medium text-blue-400">{row.nip}</td>
                          <td className="py-3.5 px-4 font-medium text-white">{row.nama}</td>
                          <td className="py-3.5 px-3 text-right font-mono font-bold text-emerald-400">
                            {row.plannedHour}
                          </td>
                          <td className="py-3.5 px-3 text-right font-mono font-bold text-cyan-400">
                            {row.effectiveHour}
                          </td>
                          <td className="py-3.5 px-3 text-right font-mono font-bold text-amber-400">
                            {row.overtimeHour}
                          </td>
                          <td className="py-3.5 px-3 text-right font-mono font-bold text-slate-300">
                            {row.idleHour}
                          </td>
                          {/* NILAI PENGISIAN TIMESHEET */}
                          <td className="py-3.5 px-3 text-right font-mono font-bold text-indigo-300 bg-indigo-950/10">
                            {row.timesheetHour}%
                          </td>
                          <td className="py-3.5 px-3 text-center font-mono font-bold text-rose-400 bg-rose-950/10">
                            {row.terlambat}
                          </td>
                          <td className="py-3.5 px-3 text-center font-mono font-bold text-amber-400 bg-amber-950/10">
                            {row.sakit}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="py-16 text-center text-slate-400 space-y-2">
                  <FileSpreadsheet className="w-10 h-10 mx-auto text-slate-500 opacity-60" />
                  <p className="text-base font-semibold text-slate-300">Data Pegawai Tidak Ditemukan</p>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    Tidak ditemukan data pegawai untuk <b>{selectedBiroPage.biroName}</b> pada sheet <b>{selectedBiroPage.month}</b>.
                  </p>
                </div>
              )}
            </div>
          </div>
        )}

      </main>
    </div>
  );
}