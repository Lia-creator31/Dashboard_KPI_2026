import { useState, useEffect } from 'react';
import { departmentsData, monthList, Department } from './data';
import * as XLSX from 'xlsx';
import excelFileUrl from './data_kpi.xlsx?url';
import jobcardFileUrl from './JOBCARD DESAIN.xlsx?url';

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
  FileText,
  User,
  Briefcase,
  ChevronDown,
  ChevronUp,
  FolderKanban,
  CheckCircle2,
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

// Peta CSV Absensi
const csvMonthMap: Record<string, string> = {
  januari: csvJan,
  februari: csvFeb,
  maret: csvMar,
  april: csvApr,
  mei: csvMei,
  juni: csvJun,
};

// Membaca semua file CSV/TXT di sub-folder src/ secara otomatis
const allCsvFiles = import.meta.glob('./**/*.{csv,CSV,txt,TXT}', { 
  query: '?raw', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

interface ExcelRow {
  nip: string;
  nama: string;
  effectiveHour: number;
  overtimeHour: number;
  idleHour: number;
  timesheetReguler: number;   // Kolom 1: Timesheet Reguler
  timesheetOvertime: number;  // Kolom 2: Timesheet Overtime
  terlambat: number;
  sakit: number;
  ipm: number;
}

interface JobcardTask {
  project: string;
  taskName: string;
  startDate: string;
  endDate: string;
  totalPersonil: string;
}

interface PersonilJobcardGroup {
  picName: string;
  tasks: JobcardTask[];
}

interface SelectedBiroPage {
  biroName: string;
  month: string;
  data: ExcelRow[];
}

interface SelectedFormPage {
  biroName: string;
  deptName: string;
}

export default function App() {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroPage, setSelectedBiroPage] = useState<SelectedBiroPage | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [tableSearch, setTableSearch] = useState('');
  const [formSearch, setFormSearch] = useState('');
  
  // State Accordion (Buka/Tutup Card Personil)
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  // State Workbooks
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [isLoadingExcel, setIsLoadingExcel] = useState<boolean>(true);

  // 1. Membaca file data_kpi.xlsx dan JOBCARD DESAIN.xlsx saat web pertama dibuka
  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        setIsLoadingExcel(true);
        // Baca data_kpi.xlsx
        const resKpi = await fetch(excelFileUrl);
        const bufferKpi = await resKpi.arrayBuffer();
        const wbKpi = XLSX.read(bufferKpi, { type: 'array' });
        setWorkbook(wbKpi);

        // Baca JOBCARD DESAIN.xlsx
        const resJc = await fetch(jobcardFileUrl);
        const bufferJc = await resJc.arrayBuffer();
        const wbJc = XLSX.read(bufferJc, { type: 'array' });
        setJobcardWorkbook(wbJc);
      } catch (error) {
        console.error("Gagal membaca file Excel:", error);
      } finally {
        setIsLoadingExcel(false);
      }
    }
    loadAllExcelFiles();
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
  const parseAbsensiCSV = (csvContent: string): Map<string, { terlambat: number; sakit: number; ipm: number }> => {
    const absensiMap = new Map<string, { terlambat: number; sakit: number; ipm: number }>();
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
        const ipmVal = parseValToNumber(parts[parts.length - 1] ?? parts[6]);

        if (cleanNameKey) {
          absensiMap.set(cleanNameKey, {
            terlambat: terlambatVal,
            sakit: sakitVal,
            ipm: ipmVal
          });
        }
      }
    });

    return absensiMap;
  };

  // Helper Parser Folder Timesheet
  const parseTimesheetFolder = (targetMonth: string): Map<string, { reguler: number; overtime: number }> => {
    const timesheetMap = new Map<string, { reguler: number; overtime: number }>();
    const monthLower = targetMonth.toLowerCase().trim();
    const monthPrefix = monthLower.slice(0, 3);

    Object.entries(allCsvFiles).forEach(([filePath, content]) => {
      const pathLower = filePath.toLowerCase();

      if (pathLower.includes('absensi_')) return;

      const isTargetMonthFile = pathLower.includes(`/${monthLower}/`) || 
                                pathLower.includes(`/${monthPrefix}/`) ||
                                pathLower.includes(`\\${monthLower}\\`) ||
                                pathLower.includes(`\\${monthPrefix}\\`);

      if (isTargetMonthFile && content) {
        const isOvertimeFile = pathLower.includes('overtime') || 
                               pathLower.includes('lembur') || 
                               content.toLowerCase().includes('total overtime hours');

        const lines = content.split(/\r?\n/);
        lines.forEach((line, lineIdx) => {
          if (lineIdx === 0 || !line.trim()) return;

          const parts = line.split(',');
          if (parts.length >= 3) {
            const rawNama = parts[2] ? parts[2].replace(/"/g, '').trim() : '';
            const cleanNameKey = rawNama.toLowerCase().replace(/\s+/g, ' ').trim();

            const rawLastVal = parts[parts.length - 1] ? parts[parts.length - 1].replace(/"/g, '').trim() : '0';
            const numVal = parseValToNumber(rawLastVal);

            if (cleanNameKey) {
              const current = timesheetMap.get(cleanNameKey) || { reguler: 0, overtime: 0 };
              
              if (isOvertimeFile) {
                current.overtime += numVal;
              } else {
                current.reguler += numVal;
              }

              timesheetMap.set(cleanNameKey, current);
            }
          }
        });
      }
    });

    return timesheetMap;
  };

  // Filter departemen utama
  const filteredDepartments = departmentsData.filter((dept) => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    (dept.description || '').toLowerCase().includes(searchQuery.toLowerCase())
  );

  // 2. Fungsi Hitung Metrik KPI Bulanan
  const handleMonthClick = (biroName: string, month: string) => {
    if (!workbook) {
      alert("File Excel data_kpi.xlsx sedang dimuat atau belum terbaca.");
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

    const rawCsvData = csvMonthMap[month.toLowerCase()] || '';
    const absensiDataMap = parseAbsensiCSV(rawCsvData);
    const timesheetDataMap = parseTimesheetFolder(month);

    const worksheet = workbook.Sheets[targetSheetName];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const personMap = new Map<string, {
      nip: string;
      nama: string;
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
            effectiveSum: 0,
            overtimeSum: 0,
            idleSum: 0,
          });
        }

        const person = personMap.get(personKey)!;
        person.effectiveSum += colK_Eff;
        person.overtimeSum += colL_Ot;
        person.idleSum += colM_Idle;
      }
    });

    const formattedData: ExcelRow[] = [];
    personMap.forEach((person) => {
      const cleanName = person.nama.toLowerCase().replace(/\s+/g, ' ').trim();
      const absensi = absensiDataMap.get(cleanName);
      const ts = timesheetDataMap.get(cleanName) || { reguler: 0, overtime: 0 };

      formattedData.push({
        nip: person.nip,
        nama: person.nama,
        effectiveHour: Math.round(person.effectiveSum * 10) / 10,
        overtimeHour: Math.round(person.overtimeSum * 10) / 10,
        idleHour: Math.round(person.idleSum * 10) / 10,
        timesheetReguler: Math.round(ts.reguler * 10) / 10,
        timesheetOvertime: Math.round(ts.overtime * 10) / 10,
        terlambat: absensi ? absensi.terlambat : 0,
        sakit: absensi ? absensi.sakit : 0,
        ipm: absensi ? absensi.ipm : 0,
      });
    });

    setTableSearch('');
    setSelectedBiroPage({
      biroName,
      month,
      data: formattedData
    });
  };

  // 3. Helper Pengolahan Data Job Card (Rekomendasi 1: Card per Pegawai)
  const getJobcardGroupsForBiro = (biroName: string): PersonilJobcardGroup[] => {
    if (!jobcardWorkbook) return [];

    // Cari sheet 'BASIC'
    const basicSheetName = jobcardWorkbook.SheetNames.find(s => s.trim().toUpperCase() === 'BASIC');
    if (!basicSheetName) return [];

    const worksheet = jobcardWorkbook.Sheets[basicSheetName];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // Peta Personil -> Daftar Tugas
    const picTaskMap = new Map<string, JobcardTask[]>();

    const bNorm = biroName.toLowerCase();

    // Fungsi pencocokan biro yang cerdas
    const isMatchingWorkCenter = (wcRaw: string): boolean => {
      const wNorm = wcRaw.toLowerCase();
      if (!wNorm) return false;

      if (bNorm.includes('kapal selam') && (wNorm.includes('kapal selam') || wNorm.includes('scorpne') || wNorm.includes('submarine'))) return true;
      if (bNorm.includes('kapal permukaan') && (wNorm.includes('kapal permukaan') || wNorm.includes('surface'))) return true;
      if (bNorm.includes('non kapal') && wNorm.includes('non kapal')) return true;
      if (bNorm.includes('pengembangan') && wNorm.includes('pengembangan')) return true;

      const cleanB = bNorm.replace(/biro/g, '').replace(/[^a-z0-9]/g, '');
      const cleanW = wNorm.replace(/biro/g, '').replace(/[^a-z0-9]/g, '');
      return cleanW.includes(cleanB) || cleanB.includes(cleanW);
    };

    rawRows.forEach((row, idx) => {
      if (!row || row.length === 0 || idx < 3) return; // Lewati baris header awal

      const projectCol    = String(row[2] || '').trim(); // Kolom C: Project
      const taskNameCol   = String(row[3] || '').trim(); // Kolom D: Task Name
      const workCenterCol = String(row[4] || '').trim(); // Kolom E: Work Center (Biro)
      const startDateCol  = String(row[5] || '-').trim(); // Kolom F: Start Date
      const endDateCol    = String(row[6] || '-').trim(); // Kolom G: End Date
      const personilCount = String(row[7] || '1').trim(); // Kolom H: Total Personil
      const rawPicCol     = String(row[8] || '').trim(); // Kolom I: PIC (Bisa berisi banyak nama dipisah koma)

      // Hanya proses baris yang work center-nya cocok dan memiliki PIC
      if (isMatchingWorkCenter(workCenterCol) && rawPicCol && !rawPicCol.toLowerCase().includes('(nama)')) {
        // Pecah nama PIC jika ada tanda koma (contoh: baris 313 berisi 10 nama)
        const individualPics = rawPicCol.split(',').map(p => p.trim()).filter(p => p.length > 1);

        individualPics.forEach((cleanPic) => {
          if (!picTaskMap.has(cleanPic)) {
            picTaskMap.set(cleanPic, []);
          }

          picTaskMap.get(cleanPic)!.push({
            project: projectCol || '-',
            taskName: taskNameCol || '-',
            startDate: startDateCol || '-',
            endDate: endDateCol || '-',
            totalPersonil: personilCount || '1'
          });
        });
      }
    });

    // Konversi Map ke Array dan urutkan abjad
    const result: PersonilJobcardGroup[] = [];
    picTaskMap.forEach((tasks, picName) => {
      result.push({ picName, tasks });
    });

    return result.sort((a, b) => a.picName.localeCompare(b.picName));
  };

  // Toggle Accordion Card
  const toggleAccordion = (picName: string) => {
    setExpandedCards(prev => ({
      ...prev,
      [picName]: !prev[picName]
    }));
  };

  const handleExpandAll = (groups: PersonilJobcardGroup[]) => {
    const allExpanded: Record<string, boolean> = {};
    groups.forEach(g => { allExpanded[g.picName] = true; });
    setExpandedCards(allExpanded);
  };

  const handleCollapseAll = () => {
    setExpandedCards({});
  };

  const filteredTableData = (selectedBiroPage?.data || []).filter(item => 
    item.nip.toLowerCase().includes(tableSearch.toLowerCase()) ||
    item.nama.toLowerCase().includes(tableSearch.toLowerCase())
  );

  // Ambil dan filter data Job Card untuk halaman FORM
  const currentJobcardGroups = selectedFormBiro ? getJobcardGroupsForBiro(selectedFormBiro.biroName) : [];
  const filteredJobcardGroups = currentJobcardGroups.filter(group => {
    const searchLower = formSearch.toLowerCase();
    const matchesName = group.picName.toLowerCase().includes(searchLower);
    const matchesTask = group.tasks.some(t => 
      t.project.toLowerCase().includes(searchLower) || 
      t.taskName.toLowerCase().includes(searchLower)
    );
    return matchesName || matchesTask;
  });

  const totalAllTasks = currentJobcardGroups.reduce((acc, g) => acc + g.tasks.length, 0);

  return (
    <div className="bg-slate-900 text-slate-100 min-h-screen font-sans antialiased">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 bg-slate-900/80 backdrop-blur-md border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div 
            className="flex items-center gap-3 cursor-pointer select-none" 
            onClick={() => {
              setSelectedFormBiro(null);
              setSelectedBiroPage(null);
              setSelectedDept(null);
            }}
          >
            <div className="p-2 bg-blue-600 rounded-lg text-white">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-bold text-base tracking-wide leading-none text-white">Rekapitulasi DESAIN 2026</h1>
              <span className="text-xs text-slate-400">Tahun Anggaran 2026</span>
            </div>
          </div>

          {/* Tombol Back Cerdas di Navbar */}
          {selectedFormBiro ? (
            <button
              onClick={() => setSelectedFormBiro(null)}
              className="flex items-center gap-2 px-3 py-1.5 text-sm bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4" />
              Kembali ke Daftar Biro
            </button>
          ) : selectedBiroPage ? (
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
        
        {/* ================= LEVEL 1: 6 DEPARTEMEN ================= */}
        {!selectedDept && !selectedBiroPage && !selectedFormBiro && (
          <div className="space-y-8 animate-fadeIn">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 border-b border-slate-800 pb-6">
              <div>
                <h2 className="text-2xl sm:text-3xl font-extrabold text-white">Dashboard Rekapitulasi Jam Kerja Desain 2026</h2>
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

        {/* ================= LEVEL 2: DAFTAR BIRO ================= */}
        {selectedDept && !selectedBiroPage && !selectedFormBiro && (
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
                {selectedDept.biros.map((biro, index) => {
                  const isDesainDasar = selectedDept.name.toLowerCase().includes('desain dasar');

                  return (
                    <div
                      key={biro.id}
                      className="bg-slate-800/70 border border-slate-700/80 rounded-xl p-5 sm:p-6 hover:border-slate-600 transition shadow-sm"
                    >
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-slate-700/60 mb-4">
                        <div className="flex flex-wrap items-center gap-3">
                          <span className="text-xs font-mono font-bold px-2.5 py-1 bg-blue-500/10 text-blue-400 border border-blue-500/20 rounded">
                            Biro 0{index + 1}
                          </span>
                          <h4 className="text-base font-bold text-white">{biro.name}</h4>

                          {/* TOMBOL "FORM" KHUSUS DEPARTEMEN DESAIN DASAR */}
                          {isDesainDasar && (
                            <button
                              onClick={() => setSelectedFormBiro({
                                biroName: biro.name,
                                deptName: selectedDept.name
                              })}
                              className="px-3 py-1 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-bold rounded-lg border border-emerald-400/40 shadow-md shadow-emerald-900/30 transition-all duration-150 flex items-center gap-1.5 cursor-pointer hover:scale-105 active:scale-95"
                            >
                              <FileText className="w-3.5 h-3.5" />
                              FORM
                            </button>
                          )}
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
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ================= HALAMAN FORMULIR: REKOMENDASI 1 (ACCORDION PER PEGAWAI) ================= */}
        {selectedFormBiro && (
          <div className="space-y-6 animate-fadeIn">
            {/* Header Halaman Form */}
            <div className="bg-gradient-to-r from-teal-900/40 via-slate-800 to-slate-800 border border-emerald-500/20 rounded-2xl p-6 sm:p-8">
              <button
                onClick={() => setSelectedFormBiro(null)}
                className="inline-flex items-center gap-2 text-sm text-emerald-400 hover:text-emerald-300 font-medium mb-3 transition cursor-pointer"
              >
                <ArrowLeft className="w-4 h-4" /> Kembali ke Daftar Biro
              </button>

              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2 text-xs font-semibold text-emerald-400 uppercase tracking-wider mb-1">
                    <FolderKanban className="w-4 h-4" />
                    <span>Monitoring Penugasan Personil — {selectedFormBiro.deptName}</span>
                  </div>
                  <h2 className="text-2xl sm:text-3xl font-extrabold text-white">
                    {selectedFormBiro.biroName}
                  </h2>
                  <p className="text-slate-400 text-sm mt-1">
                    Sumber Data: <b>JOBCARD DESAIN.xlsx</b> (Sheet: <b>BASIC</b>)
                  </p>
                </div>

                {/* Indikator Statistik Ringkas */}
                <div className="flex items-center gap-3">
                  <div className="px-4 py-2.5 bg-slate-900/80 border border-slate-700 rounded-xl text-center">
                    <div className="text-xl font-black text-emerald-400">{currentJobcardGroups.length}</div>
                    <div className="text-[10px] text-slate-400 uppercase font-semibold">Total Personil</div>
                  </div>
                  <div className="px-4 py-2.5 bg-slate-900/80 border border-slate-700 rounded-xl text-center">
                    <div className="text-xl font-black text-cyan-400">{totalAllTasks}</div>
                    <div className="text-[10px] text-slate-400 uppercase font-semibold">Total Tugas</div>
                  </div>
                </div>
              </div>
            </div>

            {/* Filter Search & Tombol Buka/Tutup Semua Card */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div className="relative w-full sm:w-80">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Cari nama personil atau proyek..."
                  value={formSearch}
                  onChange={(e) => setFormSearch(e.target.value)}
                  className="w-full pl-9 pr-4 py-2.5 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500 shadow-sm"
                />
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleExpandAll(filteredJobcardGroups)}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold rounded-lg border border-slate-700 transition cursor-pointer"
                >
                  Buka Semua
                </button>
                <button
                  onClick={handleCollapseAll}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold rounded-lg border border-slate-700 transition cursor-pointer"
                >
                  Tutup Semua
                </button>
              </div>
            </div>

            {/* DAFTAR CARD ACCORDION PER PEGAWAI (REKOMENDASI 1) */}
            {filteredJobcardGroups.length > 0 ? (
              <div className="space-y-4">
                {filteredJobcardGroups.map((person, idx) => {
                  const isExpanded = expandedCards[person.picName] ?? false;

                  return (
                    <div
                      key={idx}
                      className="bg-slate-800/70 border border-slate-700/80 rounded-2xl overflow-hidden shadow-md transition-all duration-200 hover:border-slate-600"
                    >
                      {/* HEADER CARD: KLIK UNTUK BUKA/TUTUP ACCORDION */}
                      <div
                        onClick={() => toggleAccordion(person.picName)}
                        className="p-5 flex items-center justify-between cursor-pointer select-none bg-slate-800/90 hover:bg-slate-800 transition"
                      >
                        <div className="flex items-center gap-3.5">
                          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white font-bold shadow-md shadow-emerald-900/30">
                            <User className="w-5 h-5" />
                          </div>
                          <div>
                            <h4 className="text-base font-bold text-white tracking-wide">
                              {person.picName}
                            </h4>
                            <span className="text-xs text-slate-400">
                              Personil Biro Desain Dasar
                            </span>
                          </div>
                        </div>

                        <div className="flex items-center gap-3">
                          <span className="px-3 py-1 bg-emerald-500/10 text-emerald-300 border border-emerald-500/20 rounded-full text-xs font-bold font-mono">
                            {person.tasks.length} Job Card
                          </span>
                          <div className="p-1.5 rounded-lg bg-slate-700 text-slate-300">
                            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                          </div>
                        </div>
                      </div>

                      {/* BODY ACCORDION: TABEL TUGAS DARI EXCEL */}
                      {isExpanded && (
                        <div className="p-5 border-t border-slate-700/70 bg-slate-900/40 animate-fadeIn">
                          <div className="overflow-x-auto rounded-xl border border-slate-700/60">
                            <table className="w-full text-left border-collapse text-sm">
                              <thead>
                                <tr className="bg-slate-800/90 text-slate-300 text-xs font-bold uppercase tracking-wider border-b border-slate-700">
                                  <th className="py-3 px-4 w-12 text-center">No</th>
                                  <th className="py-3 px-4 w-52">Project</th>
                                  <th className="py-3 px-4">Task Name / Uraian Pekerjaan</th>
                                  <th className="py-3 px-4 w-36 text-center">Start Date</th>
                                  <th className="py-3 px-4 w-36 text-center">End Date</th>
                                  <th className="py-3 px-4 w-24 text-center">Tim</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-700/50 text-slate-200">
                                {person.tasks.map((task, tIdx) => (
                                  <tr key={tIdx} className="hover:bg-slate-800/60 transition">
                                    <td className="py-3.5 px-4 text-center font-mono text-xs text-slate-400">
                                      {tIdx + 1}
                                    </td>
                                    <td className="py-3.5 px-4 font-semibold text-emerald-400 text-xs sm:text-sm">
                                      {task.project}
                                    </td>
                                    <td className="py-3.5 px-4 text-xs sm:text-sm leading-relaxed text-slate-100">
                                      {task.taskName}
                                    </td>
                                    <td className="py-3.5 px-4 text-center font-mono text-xs text-cyan-300">
                                      {task.startDate}
                                    </td>
                                    <td className="py-3.5 px-4 text-center font-mono text-xs text-amber-300">
                                      {task.endDate}
                                    </td>
                                    <td className="py-3.5 px-4 text-center font-mono text-xs text-slate-300">
                                      <span className="px-2 py-0.5 bg-slate-800 border border-slate-700 rounded">
                                        {task.totalPersonil} Orang
                                      </span>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="py-16 text-center text-slate-400 bg-slate-800/40 border border-slate-700/60 rounded-2xl space-y-2">
                <Briefcase className="w-10 h-10 mx-auto text-slate-500 opacity-60" />
                <p className="text-base font-semibold text-slate-300">Belum Ada Data Job Card</p>
                <p className="text-xs text-slate-500 max-w-md mx-auto">
                  Pastikan file <b>src/JOBCARD DESAIN.xlsx</b> memiliki sheet <b>BASIC</b> dan kolom Work Center (Kolom E) sesuai dengan <b>{selectedFormBiro.biroName}</b>.
                </p>
              </div>
            )}
          </div>
        )}

        {/* ================= LEVEL 3: TABEL LENGKAP REKAPITULASI KPI (11 KOLOM PRESISI) ================= */}
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

            {/* TABEL 11 KOLOM PRESISI */}
            <div className="bg-slate-800/60 border border-slate-700/80 rounded-2xl overflow-hidden shadow-xl">
              {filteredTableData.length > 0 ? (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-800/95 text-slate-300 text-xs sm:text-sm font-bold uppercase tracking-wider border-b border-slate-700">
                        <th className="py-4 px-3 w-12 text-center">No</th>
                        <th className="py-4 px-3 w-40">NIP</th>
                        <th className="py-4 px-4 min-w-[200px]">Nama Pegawai</th>
                        <th className="py-4 px-4 text-right">Effective (Hour)</th>
                        <th className="py-4 px-4 text-right">Overtime (Hour)</th>
                        <th className="py-4 px-4 text-right">Idle (Hour)</th>
                        <th className="py-4 px-4 text-right text-indigo-400 bg-indigo-950/30 border-l border-slate-700/70">
                          Timesheet Reguler
                        </th>
                        <th className="py-4 px-4 text-right text-violet-400 bg-violet-950/30 border-r border-slate-700/70">
                          Timesheet Overtime
                        </th>
                        <th className="py-4 px-4 text-center text-rose-400 bg-rose-950/30">Terlambat</th>
                        <th className="py-4 px-4 text-center text-amber-400 bg-amber-950/30">Sakit</th>
                        <th className="py-4 px-4 text-center text-purple-400 bg-purple-950/30">IPM</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-700/60 text-slate-200">
                      {filteredTableData.map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-700/40 transition">
                          <td className="py-4 px-3 text-center font-mono text-sm text-slate-400">{idx + 1}</td>
                          <td className="py-4 px-3 font-mono text-sm font-semibold text-blue-400">{row.nip}</td>
                          <td className="py-4 px-4 font-semibold text-white text-base">{row.nama}</td>
                          <td className="py-4 px-4 text-right font-mono text-base sm:text-lg font-extrabold text-cyan-400">{row.effectiveHour}</td>
                          <td className="py-4 px-4 text-right font-mono text-base sm:text-lg font-extrabold text-amber-400">{row.overtimeHour}</td>
                          <td className="py-4 px-4 text-right font-mono text-base sm:text-lg font-extrabold text-slate-200">{row.idleHour}</td>
                          <td className="py-4 px-4 text-right font-mono text-base sm:text-lg font-extrabold text-indigo-300 bg-indigo-950/15 border-l border-slate-700/50">{row.timesheetReguler}%</td>
                          <td className="py-4 px-4 text-right font-mono text-base sm:text-lg font-extrabold text-violet-300 bg-violet-950/15 border-r border-slate-700/50">{row.timesheetOvertime}%</td>
                          <td className="py-4 px-4 text-center font-mono text-base sm:text-lg font-extrabold text-rose-400 bg-rose-950/15">{row.terlambat}</td>
                          <td className="py-4 px-4 text-center font-mono text-base sm:text-lg font-extrabold text-amber-400 bg-amber-950/15">{row.sakit}</td>
                          <td className="py-4 px-4 text-center font-mono text-base sm:text-lg font-extrabold text-purple-400 bg-purple-950/15">{row.ipm}</td>
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