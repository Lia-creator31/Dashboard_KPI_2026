import { useState, useEffect } from 'react';
import { departmentsData, monthList, Department } from './data';
import * as XLSX from 'xlsx';

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
  Layers,
  User,
  ChevronDown,
  ChevronUp,
  Send,
  Trash2,
  CheckCircle2,
  AlertCircle,
  Hash,
  RotateCcw,
  Lock,
  KeyRound,
  X,
  ShieldAlert,
  Edit3,
  Check,
  Clock,
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

// Membaca file CSV/TXT secara otomatis
const allCsvFiles = import.meta.glob('./**/*.{csv,CSV,txt,TXT}', { 
  query: '?raw', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

// Membaca file Excel ber-underscore secara otomatis
const excelGlobUrls = import.meta.glob('./*.xlsx', { 
  query: '?url', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

interface ExcelRow {
  nip: string;
  nama: string;
  effectiveHour: number;
  overtimeHour: number;
  idleHour: number;
  timesheetReguler: number;
  timesheetOvertime: number;
  terlambat: number;
  sakit: number;
  ipm: number;
}

interface TaskItem {
  id: string;
  biroName: string;
  project: string;
  taskName: string;
  startDate: string;
  endDate: string;
  pic: string;
  jo: string;
  kodeJc: string; // Diisi oleh Mas Hashfi (Planner)
  status: 'pending' | 'approved';
  createdAt: string;
}

interface PersonilCardGroup {
  picName: string;
  tasks: TaskItem[];
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

function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

function isBiroMatch(biro1: string, biro2: string): boolean {
  const b1 = (biro1 || '').toLowerCase();
  const b2 = (biro2 || '').toLowerCase();

  if (b1.includes('pengembangan') && b2.includes('pengembangan')) return true;
  if (
    (b1.includes('kapal selam') || b1.includes('submarine') || b1.includes('scorpne')) &&
    (b2.includes('kapal selam') || b2.includes('submarine') || b2.includes('scorpne'))
  ) return true;
  if (b1.includes('non kapal') && b2.includes('non kapal')) return true;
  if (
    (b1.includes('kapal permukaan') || b1.includes('surface')) &&
    (b2.includes('kapal permukaan') || b2.includes('surface'))
  ) return true;

  const c1 = cleanText(b1.replace(/biro|departemen|desain|dasar/gi, ''));
  const c2 = cleanText(b2.replace(/biro|departemen|desain|dasar/gi, ''));
  if (c1 && c2) return c1.includes(c2) || c2.includes(c1);
  return false;
}

async function fetchSafeWorkbook(paths: (string | undefined)[]): Promise<XLSX.WorkBook | null> {
  for (const p of paths) {
    if (!p) continue;
    try {
      const res = await fetch(p);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const bytes = new Uint8Array(buf.slice(0, 4));
        if (bytes[0] === 80 && bytes[1] === 75 && bytes[2] === 3 && bytes[3] === 4) {
          return XLSX.read(buf, { type: 'array' });
        }
      }
    } catch {
      // Coba path alternatif berikutnya
    }
  }
  return null;
}

export default function App() {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroPage, setSelectedBiroPage] = useState<SelectedBiroPage | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  
  // Tab Mode: 'form' (User Input) vs 'output' (Hasil Rekapitulasi)
  const [formPageMode, setFormPageMode] = useState<'form' | 'output'>('form');

  const [searchQuery, setSearchQuery] = useState('');
  const [tableSearch, setTableSearch] = useState('');
  const [outputSearch, setOutputSearch] = useState('');
  
  // 7 Input Form State (Murni User Biro)
  const [formData, setFormData] = useState({
    nama: '',
    kodeProyek: '',
    taskName: '',
    startDate: '',
    endDate: '',
    pic: '',
    jo: ''
  });

  // Database Job Card Terisi
  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>(() => {
    try {
      const saved = localStorage.getItem('kpi_form_database');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Panel Planner (Mas Hashfi)
  const [isPlannerModalOpen, setIsPlannerModalOpen] = useState(false);
  const [isPlannerUnlocked, setIsPlannerUnlocked] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [pinError, setPinError] = useState(false);

  // Input Kode JC per Task oleh Planner
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});

  const PLANNER_PIN = '2026'; // PIN Rahasia Mas Hashfi

  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  // Workbooks
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [strukturWorkbook, setStrukturWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [isLoadingExcel, setIsLoadingExcel] = useState<boolean>(true);

  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        setIsLoadingExcel(true);

        let kpiUrl = '';
        let jcUrl = '';
        let strukturUrl = '';

        Object.entries(excelGlobUrls).forEach(([path, url]) => {
          const pLower = path.toLowerCase();
          if (pLower.includes('kpi')) kpiUrl = url;
          else if (pLower.includes('jobcard')) jcUrl = url;
          else if (pLower.includes('struktur')) strukturUrl = url;
        });

        const [wbKpi, wbJc, wbStruktur] = await Promise.all([
          fetchSafeWorkbook([kpiUrl, '/data_kpi.xlsx', './data_kpi.xlsx']),
          fetchSafeWorkbook([jcUrl, '/JOBCARD_DESAIN.xlsx', './JOBCARD_DESAIN.xlsx', '/JOBCARD DESAIN.xlsx']),
          fetchSafeWorkbook([
            strukturUrl, 
            '/Struktur_dan_Anggota_Desain_Upd_0826_(1).xlsx',
            './Struktur_dan_Anggota_Desain_Upd_0826_(1).xlsx',
            '/Struktur_dan_Anggota_Desain_Upd_0826.xlsx',
            './Struktur_dan_Anggota_Desain_Upd_0826.xlsx'
          ])
        ]);

        if (wbKpi) setWorkbook(wbKpi);
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbStruktur) setStrukturWorkbook(wbStruktur);
      } catch (error) {
        console.error("Gagal membaca file Excel:", error);
      } finally {
        setIsLoadingExcel(false);
      }
    }
    loadAllExcelFiles();
  }, []);

  const parseValToNumber = (val: any): number => {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    const str = String(val).trim().replace('%', '').replace(',', '.');
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  const parseAbsensiCSV = (csvContent: string): Map<string, { terlambat: number; sakit: number; ipm: number }> => {
    const absensiMap = new Map<string, { terlambat: number; sakit: number; ipm: number }>();
    if (!csvContent) return absensiMap;

    const lines = csvContent.split(/\r?\n/);
    lines.forEach((line, index) => {
      if (index === 0 || !line.trim()) return;
      const parts = line.split(',');
      if (parts.length >= 6) {
        const rawNama = parts[1] ? parts[1].trim() : '';
        const cleanNameKey = cleanText(rawNama);

        const terlambatVal = parseValToNumber(parts[parts.length - 3] ?? parts[4]);
        const sakitVal = parseValToNumber(parts[parts.length - 2] ?? parts[5]);
        const ipmVal = parseValToNumber(parts[parts.length - 1] ?? parts[6]);

        if (cleanNameKey) {
          absensiMap.set(cleanNameKey, { terlambat: terlambatVal, sakit: sakitVal, ipm: ipmVal });
        }
      }
    });

    return absensiMap;
  };

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
            const cleanNameKey = cleanText(rawNama);
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

  // 1. Ambil Nama Personil Murni Sesuai Biro dari Sheet CalonPers
  const getBiroMembers = (biroName: string): string[] => {
    if (!strukturWorkbook) return [];
    const sheet = strukturWorkbook.Sheets['CalonPers'] || strukturWorkbook.Sheets[strukturWorkbook.SheetNames[0]];
    if (!sheet) return [];
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    const colBlocks = [
      { dCol: 2, nCol: 3, maxR: 50 },
      { dCol: 8, nCol: 9, maxR: 55 },
      { dCol: 14, nCol: 15, maxR: 83 },
      { dCol: 22, nCol: 23, maxR: 67 },
      { dCol: 29, nCol: 30, maxR: 66 },
      { dCol: 36, nCol: 37, maxR: 30 },
    ];

    const members: string[] = [];

    colBlocks.forEach(({ dCol, nCol, maxR }) => {
      let currentBiro = '';
      for (let r = 17; r < Math.min(rows.length, maxR); r++) {
        const row = rows[r];
        if (!row) continue;
        const valD = String(row[dCol] || '').trim();
        const valN = String(row[nCol] || '').trim();

        if (valD.toLowerCase().includes('biro')) {
          currentBiro = valD;
        } else if (currentBiro && isBiroMatch(currentBiro, biroName)) {
          if (
            valD && 
            !/^\d+$/.test(valD) && 
            (!valN || valN.toLowerCase() === 'nan') && 
            !valD.toLowerCase().includes('departemen') && 
            !valD.toLowerCase().includes('personil') && 
            !valD.toLowerCase().includes('total')
          ) {
            if (!members.includes(valD) && valD.toLowerCase() !== 'nan') members.push(valD);
          } else if (valN && valN.toLowerCase() !== 'nan' && valN.toUpperCase() !== 'PERSONIL' && !valN.toLowerCase().includes('departemen')) {
            if (!members.includes(valN)) members.push(valN);
          }
        }
      }
    });

    return members;
  };

  // 2. Dropdown Kode Proyek dari Sheet BASIC JOBCARD_DESAIN
  const getJobcardProjects = (): string[] => {
    if (!jobcardWorkbook) return [];
    const basicSheetName = jobcardWorkbook.SheetNames.find(s => s.trim().toUpperCase() === 'BASIC') || jobcardWorkbook.SheetNames[0];
    if (!basicSheetName) return [];
    const worksheet = jobcardWorkbook.Sheets[basicSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const projects = new Set<string>();
    rows.forEach((row, idx) => {
      if (idx < 3 || !row) return;
      const p = String(row[2] || '').trim();
      if (p && p.toLowerCase() !== 'nan' && !p.toLowerCase().includes('kode proyek')) {
        projects.add(p);
      }
    });
    return Array.from(projects).sort();
  };

  // 3. Dropdown Task Name dari Sheet BASIC JOBCARD_DESAIN
  const getJobcardTasks = (): string[] => {
    if (!jobcardWorkbook) return [];
    const basicSheetName = jobcardWorkbook.SheetNames.find(s => s.trim().toUpperCase() === 'BASIC') || jobcardWorkbook.SheetNames[0];
    if (!basicSheetName) return [];
    const worksheet = jobcardWorkbook.Sheets[basicSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const tasks = new Set<string>();
    rows.forEach((row, idx) => {
      if (idx < 3 || !row) return;
      const t = String(row[3] || '').trim();
      if (t && t.toLowerCase() !== 'nan' && !t.toLowerCase().includes('desc pekerjaan') && !t.toLowerCase().includes('revisi ke')) {
        tasks.add(t);
      }
    });
    return Array.from(tasks).sort();
  };

  // 4. Output Accordion: Personil Biro + Tugas Hasil Submit Form
  const getAccordionOutputForBiro = (targetBiroName: string): PersonilCardGroup[] => {
    const biroMembers = getBiroMembers(targetBiroName);
    const personMap = new Map<string, TaskItem[]>();

    biroMembers.forEach(m => {
      personMap.set(m, []);
    });

    const biroKey = cleanText(targetBiroName);
    const tasksForThisBiro = manualTasks[biroKey] || [];

    tasksForThisBiro.forEach(t => {
      let matchedName = biroMembers.find(m => cleanText(m) === cleanText(t.pic)) || t.pic;
      if (!personMap.has(matchedName)) {
        personMap.set(matchedName, []);
      }
      personMap.get(matchedName)!.push(t);
    });

    const result: PersonilCardGroup[] = [];
    personMap.forEach((tasks, picName) => {
      result.push({ picName, tasks });
    });

    return result.sort((a, b) => a.picName.localeCompare(b.picName));
  };

  // User Submit Form: Otomatis Menyimpan Task Name & Kode Proyek (Kode JC Kosong Dulu)
  const handleSubmitForm = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedFormBiro) return;

    const biroKey = cleanText(selectedFormBiro.biroName);
    const newTask: TaskItem = {
      id: Date.now().toString(),
      biroName: selectedFormBiro.biroName,
      project: formData.kodeProyek,
      taskName: formData.taskName,
      startDate: formData.startDate,
      endDate: formData.endDate,
      pic: formData.nama,
      jo: formData.jo,
      kodeJc: '', // Menunggu diisi Mas Hashfi (Planner)
      status: 'pending',
      createdAt: new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
    };

    const currentList = manualTasks[biroKey] || [];
    const updated = { ...manualTasks, [biroKey]: [newTask, ...currentList] };

    setManualTasks(updated);
    try {
      localStorage.setItem('kpi_form_database', JSON.stringify(updated));
    } catch {
      // Fallback
    }

    setExpandedCards(prev => ({ ...prev, [formData.nama]: true }));

    // Reset Form
    setFormData({
      nama: '',
      kodeProyek: '',
      taskName: '',
      startDate: '',
      endDate: '',
      pic: '',
      jo: ''
    });

    alert(`Job Card berhasil disubmit!\nData Proyek (${newTask.project}) & Task Name telah masuk ke antrean Planner (Mas Hashfi) untuk diberikan Kode Job Card.`);
  };

  // Handler Mas Hashfi: Verifikasi PIN Planner
  const handleVerifyPin = (e: React.FormEvent) => {
    e.preventDefault();
    if (pinInput === PLANNER_PIN) {
      setIsPlannerUnlocked(true);
      setPinError(false);
      setPinInput('');
    } else {
      setPinError(true);
    }
  };

  // Handler Mas Hashfi: Menyimpan Kode JC untuk suatu Job Card (Hanya 1 Kali Input!)
  const handleSaveKodeJcForTask = (taskId: string, biroName: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) {
      alert('Harap masukkan Kode Job Card terlebih dahulu!');
      return;
    }

    const biroKey = cleanText(biroName);
    const currentList = manualTasks[biroKey] || [];
    
    const updatedList = currentList.map(task => {
      if (task.id === taskId) {
        return {
          ...task,
          kodeJc: inputVal,
          status: 'approved' as const
        };
      }
      return task;
    });

    const updated = { ...manualTasks, [biroKey]: updatedList };
    setManualTasks(updated);
    try {
      localStorage.setItem('kpi_form_database', JSON.stringify(updated));
    } catch {
      // Fallback
    }

    // Bersihkan field temporary
    setEditingTaskKode(prev => {
      const next = { ...prev };
      delete next[taskId];
      return next;
    });

    alert(`Kode Job Card "${inputVal}" berhasil disimpan untuk tugas tersebut!`);
  };

  // Hapus Satu Tugas
  const handleDeleteTask = (taskId: string, biroName: string) => {
    if (window.confirm('Hapus penugasan ini?')) {
      const biroKey = cleanText(biroName);
      const currentList = manualTasks[biroKey] || [];
      const updatedList = currentList.filter(t => t.id !== taskId);
      const updated = { ...manualTasks, [biroKey]: updatedList };

      setManualTasks(updated);
      try {
        localStorage.setItem('kpi_form_database', JSON.stringify(updated));
      } catch {
        // Fallback
      }
    }
  };

  // Hapus Seluruh Database Form di Biro Ini
  const handleClearAllBiroData = () => {
    if (!selectedFormBiro) return;
    if (window.confirm(`Hapus semua database Job Card di ${selectedFormBiro.biroName}?`)) {
      const biroKey = cleanText(selectedFormBiro.biroName);
      const updated = { ...manualTasks, [biroKey]: [] };
      setManualTasks(updated);
      try {
        localStorage.setItem('kpi_form_database', JSON.stringify(updated));
      } catch {
        // Fallback
      }
    }
  };

  const toggleAccordion = (picName: string) => {
    setExpandedCards(prev => ({ ...prev, [picName]: !prev[picName] }));
  };

  // Data Rekapitulasi KPI Bulanan
  const handleMonthClick = (biroName: string, month: string) => {
    if (!workbook) {
      alert("File data_kpi.xlsx belum terbaca.");
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

    const personMap = new Map<string, { nip: string; nama: string; effectiveSum: number; overtimeSum: number; idleSum: number }>();

    rawRows.forEach((row, rowIndex) => {
      if (!row || row.length === 0 || rowIndex === 0) return;

      const colA_NIP   = String(row[0] || '').trim();
      const colB_Nama  = String(row[1] || '').trim();
      const colH_Biro  = String(row[7] || '').trim();
      const colK_Eff   = parseValToNumber(row[10]);
      const colL_Ot    = parseValToNumber(row[11]);
      const colM_Idle  = parseValToNumber(row[12]);

      if (isBiroMatch(colH_Biro, biroName) && (colB_Nama || colA_NIP)) {
        const personKey = cleanText(colB_Nama || colA_NIP);

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
      const cleanName = cleanText(person.nama);
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
    setSelectedBiroPage({ biroName, month, data: formattedData });
  };

  const filteredDepartments = (departmentsData || []).filter((dept) => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    (dept.description || '').toLowerCase().includes(searchQuery.toLowerCase())
  );

  const filteredTableData = (selectedBiroPage?.data || []).filter(item => 
    item.nip.toLowerCase().includes(tableSearch.toLowerCase()) ||
    item.nama.toLowerCase().includes(tableSearch.toLowerCase())
  );

  // Sumber Data Form
  const currentBiroMembers = selectedFormBiro ? getBiroMembers(selectedFormBiro.biroName) : [];
  const projectOptions = getJobcardProjects();
  const taskOptions = getJobcardTasks();
  
  // Output Accordion Murni
  const accordionData = selectedFormBiro ? getAccordionOutputForBiro(selectedFormBiro.biroName) : [];
  const filteredAccordionData = accordionData.filter(g => {
    const s = outputSearch.toLowerCase();
    return g.picName.toLowerCase().includes(s) || g.tasks.some(t => t.project.toLowerCase().includes(s) || t.taskName.toLowerCase().includes(s) || (t.kodeJc || '').toLowerCase().includes(s));
  });

  const totalPersonilCount = accordionData.length;
  const totalTasksCount = accordionData.reduce((acc, g) => acc + g.tasks.length, 0);

  // Daftar tugas yang diajukan di biro ini untuk dikelola Planner
  const currentBiroKey = selectedFormBiro ? cleanText(selectedFormBiro.biroName) : '';
  const currentBiroSubmittedTasks = manualTasks[currentBiroKey] || [];
  const pendingTasksCount = currentBiroSubmittedTasks.filter(t => !t.kodeJc).length;

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
        
        {/* LEVEL 1: 6 DEPARTEMEN */}
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

        {/* LEVEL 2: DAFTAR BIRO */}
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

                          {/* 2 Tombol Terpisah: FORM & OUTPUT */}
                          {isDesainDasar && (
                            <div className="flex items-center gap-2">
                              <button
                                onClick={() => {
                                  setSelectedFormBiro({
                                    biroName: biro.name,
                                    deptName: selectedDept.name
                                  });
                                  setFormPageMode('form');
                                }}
                                className="px-3 py-1 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-bold rounded-lg border border-emerald-400/40 shadow-md shadow-emerald-900/30 transition-all duration-150 flex items-center gap-1.5 cursor-pointer hover:scale-105 active:scale-95"
                                title="Buka Form Pengisian Job Card"
                              >
                                <FileText className="w-3.5 h-3.5" />
                                FORM
                              </button>

                              <button
                                onClick={() => {
                                  setSelectedFormBiro({
                                    biroName: biro.name,
                                    deptName: selectedDept.name
                                  });
                                  setFormPageMode('output');
                                }}
                                className="px-3 py-1 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white text-xs font-bold rounded-lg border border-blue-400/40 shadow-md shadow-blue-900/30 transition-all duration-150 flex items-center gap-1.5 cursor-pointer hover:scale-105 active:scale-95"
                                title="Buka Output Hasil Rekapitulasi Personil"
                              >
                                <Layers className="w-3.5 h-3.5" />
                                OUTPUT
                              </button>
                            </div>
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
                              {isLoadingExcel ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : month}
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

        {/* ================= LEVEL TERPISAH: FORM vs OUTPUT ================= */}
        {selectedFormBiro && (
          <div className="space-y-6 animate-fadeIn max-w-5xl mx-auto">
            {/* Header Bersama & Tab Navigasi Pemisah Halaman */}
            <div className="bg-gradient-to-r from-teal-900/40 via-slate-800 to-slate-800 border border-slate-700/80 rounded-2xl p-6 sm:p-8">
              <button
                onClick={() => setSelectedFormBiro(null)}
                className="inline-flex items-center gap-2 text-sm text-blue-400 hover:text-blue-300 font-medium mb-3 transition cursor-pointer"
              >
                <ArrowLeft className="w-4 h-4" /> Kembali ke Daftar Biro
              </button>

              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2 text-xs font-semibold text-blue-400 uppercase tracking-wider mb-1">
                    <FileText className="w-4 h-4" />
                    <span>Manajemen Job Card — {selectedFormBiro.deptName}</span>
                  </div>
                  <h2 className="text-2xl sm:text-3xl font-extrabold text-white">
                    {selectedFormBiro.biroName}
                  </h2>
                  <div className="flex flex-wrap items-center gap-2 mt-1 text-xs text-slate-400">
                    <span>File Master: <b>Struktur</b> & <b>JOBCARD_DESAIN</b></span>
                    <span>•</span>
                    {strukturWorkbook && jobcardWorkbook ? (
                      <span className="text-emerald-400 font-medium flex items-center gap-1">
                        <CheckCircle2 className="w-3.5 h-3.5" /> Data Excel Terhubung
                      </span>
                    ) : (
                      <span className="text-amber-400 font-medium flex items-center gap-1">
                        <AlertCircle className="w-3.5 h-3.5" /> Membaca Excel...
                      </span>
                    )}
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  <div className="px-4 py-2.5 bg-slate-900/80 border border-slate-700 rounded-xl text-center">
                    <div className="text-xl font-black text-emerald-400">{totalPersonilCount}</div>
                    <div className="text-[10px] text-slate-400 uppercase font-semibold">Total Personil</div>
                  </div>
                  <div className="px-4 py-2.5 bg-slate-900/80 border border-slate-700 rounded-xl text-center">
                    <div className="text-xl font-black text-cyan-400">{totalTasksCount}</div>
                    <div className="text-[10px] text-slate-400 uppercase font-semibold">Total Job Card</div>
                  </div>
                </div>
              </div>

              {/* TAB SWITCHER: PINDAH ANTARA HALAMAN FORM & HALAMAN OUTPUT */}
              <div className="mt-6 pt-5 border-t border-slate-700/60 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => setFormPageMode('form')}
                    className={`px-5 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition flex items-center gap-2 cursor-pointer ${
                      formPageMode === 'form'
                        ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-900/40 border border-emerald-500/50'
                        : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-white border border-slate-700'
                    }`}
                  >
                    <FileText className="w-4 h-4" />
                    Halaman Form Pengisian
                  </button>

                  <button
                    onClick={() => setFormPageMode('output')}
                    className={`px-5 py-2.5 rounded-xl text-xs sm:text-sm font-bold transition flex items-center gap-2 cursor-pointer ${
                      formPageMode === 'output'
                        ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/40 border border-blue-500/50'
                        : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-white border border-slate-700'
                    }`}
                  >
                    <Layers className="w-4 h-4" />
                    Halaman Output Rekapitulasi ({totalTasksCount})
                  </button>
                </div>

                {/* Tombol Kelola Kode JC (Planner Mas Hashfi) */}
                <button
                  type="button"
                  onClick={() => {
                    setIsPlannerModalOpen(true);
                    setPinError(false);
                  }}
                  className="px-4 py-2.5 bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-500 hover:to-orange-500 text-white rounded-xl text-xs font-bold border border-amber-400/40 shadow-md shadow-amber-950/40 transition flex items-center gap-2 cursor-pointer"
                >
                  <Lock className="w-4 h-4" />
                  Kelola Kode JC (Planner)
                  {pendingTasksCount > 0 && (
                    <span className="px-2 py-0.2 bg-rose-600 text-white text-[10px] font-bold rounded-full animate-pulse">
                      {pendingTasksCount}
                    </span>
                  )}
                </button>
              </div>
            </div>

            {/* ================= PILIHAN 1: HALAMAN FORM PENGISIAN ================= */}
            {formPageMode === 'form' && (
              <div className="bg-slate-800/70 border border-slate-700/80 rounded-2xl p-6 sm:p-8 shadow-xl animate-fadeIn">
                <div className="mb-6 pb-4 border-b border-slate-700/70">
                  <h3 className="text-lg font-bold text-white flex items-center gap-2">
                    <FileText className="w-5 h-5 text-emerald-400" />
                    Formulir Input Job Card Personil
                  </h3>
                  <p className="text-xs text-slate-400 mt-1">
                    Isi 7 data di bawah ini. Setelah Anda submit, tugas akan langsung tercatat dan masuk ke antrean Planner (Mas Hashfi) untuk penerbitan Kode Job Card.
                  </p>
                </div>

                <form onSubmit={handleSubmitForm} className="space-y-6">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                    {/* 1. NAMA (DROPDOWN STRUKTUR BIRO) */}
                    <div className="md:col-span-2">
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        1. Nama Personil (Sesuai Struktur Biro) <span className="text-rose-400">*</span>
                      </label>
                      <select
                        value={formData.nama}
                        onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      >
                        <option value="">-- Pilih Nama Personil di {selectedFormBiro.biroName} --</option>
                        {currentBiroMembers.map((nama, idx) => (
                          <option key={idx} value={nama}>{nama}</option>
                        ))}
                      </select>
                    </div>

                    {/* 2. KODE PROYEK (DROPDOWN JOBCARD DESAIN SHEET BASIC) */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        2. Kode Proyek <span className="text-rose-400">*</span>
                      </label>
                      <select
                        value={formData.kodeProyek}
                        onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      >
                        <option value="">-- Pilih Kode Proyek --</option>
                        {projectOptions.map((proj, idx) => (
                          <option key={idx} value={proj}>{proj}</option>
                        ))}
                      </select>
                    </div>

                    {/* 7. JO (ESAI KHUSUS ANGKA) */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        7. JO (Khusus Angka) <span className="text-rose-400">*</span>
                      </label>
                      <div className="relative">
                        <Hash className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        <input
                          type="text"
                          inputMode="numeric"
                          value={formData.jo}
                          onChange={(e) => {
                            const digitsOnly = e.target.value.replace(/[^0-9]/g, '');
                            setFormData(prev => ({ ...prev, jo: digitsOnly }));
                          }}
                          placeholder="Contoh: 300426"
                          required
                          className="w-full pl-10 pr-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white font-mono placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                        />
                      </div>
                    </div>

                    {/* 3. TASK NAME (DROPDOWN JOBCARD DESAIN SHEET BASIC) */}
                    <div className="md:col-span-2">
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        3. Task Name / Uraian Pekerjaan <span className="text-rose-400">*</span>
                      </label>
                      <select
                        value={formData.taskName}
                        onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      >
                        <option value="">-- Pilih Task Name / Deskripsi Pekerjaan --</option>
                        {taskOptions.map((task, idx) => (
                          <option key={idx} value={task}>{task}</option>
                        ))}
                      </select>
                    </div>

                    {/* 4. SCHEDULE START DATE (KALENDER) */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        4. Schedule Start Date <span className="text-rose-400">*</span>
                      </label>
                      <input
                        type="date"
                        value={formData.startDate}
                        onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      />
                    </div>

                    {/* 5. SCHEDULE END DATE (KALENDER) */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        5. Schedule End Date <span className="text-rose-400">*</span>
                      </label>
                      <input
                        type="date"
                        value={formData.endDate}
                        onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      />
                    </div>

                    {/* 6. PIC (ESAI / TEKS BEBAS) */}
                    <div className="md:col-span-2">
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-300 mb-2">
                        6. PIC (Ketik Bebas / Esai) <span className="text-rose-400">*</span>
                      </label>
                      <textarea
                        rows={2}
                        value={formData.pic}
                        onChange={(e) => setFormData(prev => ({ ...prev, pic: e.target.value }))}
                        placeholder="Ketikkan nama PIC / personil yang mengerjakan..."
                        required
                        className="w-full px-4 py-3 bg-slate-800 border border-slate-700 rounded-xl text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-emerald-500 leading-relaxed"
                      />
                    </div>
                  </div>

                  {/* Keterangan Status Kode JC */}
                  <div className="p-3 bg-amber-500/10 border border-amber-500/20 rounded-xl flex items-center gap-2.5 text-xs text-amber-300">
                    <Clock className="w-4 h-4 shrink-0 text-amber-400" />
                    <span><b>Catatan:</b> Kode Job Card akan diverifikasi dan diisi oleh Planner (Mas Hashfi) setelah formulir ini Anda simpan.</span>
                  </div>

                  {/* Tombol Simpan */}
                  <div className="pt-4 border-t border-slate-700/70 flex items-center justify-end gap-3">
                    <button
                      type="button"
                      onClick={() => setFormData({
                        nama: '',
                        kodeProyek: '',
                        taskName: '',
                        startDate: '',
                        endDate: '',
                        pic: '',
                        jo: ''
                      })}
                      className="px-5 py-2.5 bg-slate-700 hover:bg-slate-600 text-slate-200 text-sm font-semibold rounded-xl transition cursor-pointer"
                    >
                      Reset Form
                    </button>
                    <button
                      type="submit"
                      className="px-6 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-sm font-bold rounded-xl shadow-lg shadow-emerald-900/30 transition flex items-center gap-2 cursor-pointer hover:scale-105 active:scale-95"
                    >
                      <Send className="w-4 h-4" />
                      Simpan & Kirim ke Planner
                    </button>
                  </div>
                </form>
              </div>
            )}

            {/* ================= PILIHAN 2: HALAMAN OUTPUT REKAPITULASI ================= */}
            {formPageMode === 'output' && (
              <div className="space-y-5 animate-fadeIn">
                <div className="bg-slate-800/80 border border-slate-700/80 rounded-2xl p-5 sm:p-6 shadow-md">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-700/60 pb-4 mb-4">
                    <div>
                      <h3 className="text-base font-bold text-white tracking-wide">
                        [ FORMULIR BIRO: {selectedFormBiro.biroName.toUpperCase()} ]
                      </h3>
                      <p className="text-xs text-slate-400 mt-0.5">
                        Total Personil: <b className="text-emerald-400">{totalPersonilCount} Pegawai</b> | Total Job Card: <b className="text-cyan-400">{totalTasksCount} Tugas</b>
                        {pendingTasksCount > 0 && (
                          <span className="ml-2 text-rose-400 font-semibold">({pendingTasksCount} Menunggu Kode JC)</span>
                        )}
                      </p>
                    </div>

                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        onClick={() => {
                          const all: Record<string, boolean> = {};
                          filteredAccordionData.forEach(g => { all[g.picName] = true; });
                          setExpandedCards(all);
                        }}
                        className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold rounded-lg border border-slate-700 transition cursor-pointer"
                      >
                        Buka Semua
                      </button>
                      <button
                        onClick={() => setExpandedCards({})}
                        className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold rounded-lg border border-slate-700 transition cursor-pointer"
                      >
                        Tutup Semua
                      </button>
                      <button
                        onClick={handleClearAllBiroData}
                        className="px-3 py-1.5 bg-rose-950/40 hover:bg-rose-900 text-rose-300 text-xs font-semibold rounded-lg border border-rose-800/60 transition cursor-pointer flex items-center gap-1"
                        title="Hapus seluruh tugas hasil form di biro ini"
                      >
                        <RotateCcw className="w-3.5 h-3.5" />
                        Kosongkan Database Form
                      </button>
                    </div>
                  </div>

                  {/* Pencarian Personil / Proyek */}
                  <div className="relative">
                    <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Cari nama personil, proyek, atau kode JC..."
                      value={outputSearch}
                      onChange={(e) => setOutputSearch(e.target.value)}
                      className="w-full pl-10 pr-4 py-2 bg-slate-900 border border-slate-700 rounded-xl text-xs sm:text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                </div>

                {/* LIST CARD ACCORDION PER PEGAWAI */}
                {filteredAccordionData.length > 0 ? (
                  <div className="space-y-3.5">
                    {filteredAccordionData.map((person, idx) => {
                      const isExpanded = expandedCards[person.picName] ?? false;
                      const taskCount = person.tasks.length;
                      const isActive = taskCount > 0;

                      return (
                        <div
                          key={idx}
                          className="bg-slate-800/70 border border-slate-700/80 rounded-2xl overflow-hidden shadow-md transition-all duration-200 hover:border-slate-600"
                        >
                          {/* HEADER CARD: KLIK UNTUK BUKA/TUTUP ACCORDION */}
                          <div
                            onClick={() => toggleAccordion(person.picName)}
                            className="p-4 sm:p-5 flex items-center justify-between cursor-pointer select-none bg-slate-800/90 hover:bg-slate-800 transition"
                          >
                            <div className="flex items-center gap-3.5">
                              {/* Icon / Foto */}
                              <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-bold text-white shadow-md ${
                                isActive ? 'bg-gradient-to-br from-emerald-600 to-teal-700 shadow-emerald-900/30' : 'bg-slate-700 text-slate-400'
                              }`}>
                                <User className="w-5 h-5" />
                              </div>

                              {/* Nama & Status */}
                              <div>
                                <div className="flex flex-wrap items-center gap-2">
                                  <h4 className="text-base font-bold text-white tracking-wide">
                                    {person.picName}
                                  </h4>
                                  <span className="text-xs font-semibold text-slate-400">
                                    ({taskCount} Job Card)
                                  </span>
                                  <span className={`px-2 py-0.5 text-[10px] font-bold rounded-md border ${
                                    isActive 
                                      ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30' 
                                      : 'bg-slate-700/40 text-slate-400 border-slate-600'
                                  }`}>
                                    Status: {isActive ? 'Aktif' : 'Kosong'}
                                  </span>
                                </div>
                              </div>
                            </div>

                            {/* Tombol Panah Buka / Tutup */}
                            <div className="p-1.5 rounded-lg bg-slate-700 text-slate-300">
                              {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                            </div>
                          </div>

                          {/* BODY TABEL JOB CARD */}
                          {isExpanded && (
                            <div className="p-4 sm:p-5 border-t border-slate-700/70 bg-slate-900/40 animate-fadeIn">
                              {taskCount > 0 ? (
                                <div className="overflow-x-auto rounded-xl border border-slate-700/60">
                                  <table className="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                      <tr className="bg-slate-800/95 text-slate-300 font-bold uppercase tracking-wider border-b border-slate-700">
                                        <th className="py-3 px-3 w-10 text-center">No</th>
                                        <th className="py-3 px-3 w-40 font-mono text-amber-400">Kode Job Card</th>
                                        <th className="py-3 px-4 w-48">Project</th>
                                        <th className="py-3 px-4">Task Name / Uraian Pekerjaan</th>
                                        <th className="py-3 px-3 text-center">Start Date</th>
                                        <th className="py-3 px-3 text-center">End Date</th>
                                        <th className="py-3 px-3 text-center">JO</th>
                                        <th className="py-3 px-2 w-16 text-center">Aksi</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-700/50 text-slate-200">
                                      {person.tasks.map((task, tIdx) => (
                                        <tr key={task.id} className="hover:bg-slate-800/60 transition">
                                          <td className="py-3.5 px-3 text-center font-mono text-slate-400">
                                            {tIdx + 1}
                                          </td>

                                          {/* Kolom Kode JC */}
                                          <td className="py-3.5 px-3 font-mono">
                                            {task.kodeJc ? (
                                              <span className="px-2 py-1 bg-amber-500/10 text-amber-300 border border-amber-500/30 rounded font-bold text-xs">
                                                {task.kodeJc}
                                              </span>
                                            ) : (
                                              <span className="px-2 py-0.5 bg-rose-500/10 text-rose-300 border border-rose-500/30 rounded text-[11px] font-semibold flex items-center gap-1">
                                                <Clock className="w-3 h-3" /> Menunggu Planner
                                              </span>
                                            )}
                                          </td>

                                          <td className="py-3.5 px-4 font-semibold text-emerald-400">
                                            {task.project}
                                          </td>
                                          <td className="py-3.5 px-4 leading-relaxed text-slate-100">
                                            {task.taskName}
                                          </td>
                                          <td className="py-3.5 px-3 text-center font-mono text-cyan-300">
                                            {task.startDate}
                                          </td>
                                          <td className="py-3.5 px-3 text-center font-mono text-cyan-300">
                                            {task.endDate}
                                          </td>
                                          <td className="py-3.5 px-3 text-center font-mono font-bold text-violet-300">
                                            #{task.jo}
                                          </td>
                                          <td className="py-3.5 px-2 text-center">
                                            <button
                                              onClick={() => handleDeleteTask(task.id, selectedFormBiro.biroName)}
                                              className="p-1.5 text-rose-400 hover:text-white hover:bg-rose-600 rounded-lg transition cursor-pointer"
                                              title="Hapus job card ini"
                                            >
                                              <Trash2 className="w-3.5 h-3.5" />
                                            </button>
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              ) : (
                                <div className="py-6 text-center text-slate-400 text-xs">
                                  Personil ini belum memiliki tugas aktif. Silakan isi melalui <b>Halaman Form Pengisian</b>.
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <div className="py-12 text-center text-slate-400 bg-slate-800/40 border border-slate-700/60 rounded-2xl">
                    Tidak ditemukan data personil untuk <b>{selectedFormBiro.biroName}</b>.
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ================= MODAL KHUSUS PLANNER (MAS HASHFI) ================= */}
        {isPlannerModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
            <div className="bg-slate-800 border border-slate-700 rounded-2xl w-full max-w-3xl shadow-2xl overflow-hidden max-h-[90vh] flex flex-col">
              {/* Header Modal */}
              <div className="p-5 bg-slate-900 border-b border-slate-700 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-amber-500/10 text-amber-400 rounded-lg border border-amber-500/20">
                    <Lock className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="font-bold text-white text-sm">Panel Verifikasi Kode JC (Mas Hashfi - Planner)</h3>
                    <span className="text-[11px] text-slate-400">Pemberian Kode Job Card untuk Form yang Sudah Disubmit User</span>
                  </div>
                </div>
                <button
                  onClick={() => setIsPlannerModalOpen(false)}
                  className="p-1 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Konten Modal */}
              <div className="p-6 overflow-y-auto flex-1">
                {!isPlannerUnlocked ? (
                  /* Form Input PIN */
                  <form onSubmit={handleVerifyPin} className="max-w-sm mx-auto space-y-4 py-6">
                    <div className="text-center">
                      <KeyRound className="w-12 h-12 mx-auto text-amber-400 mb-2" />
                      <h4 className="text-base font-bold text-white">Masukkan PIN Akses Planner</h4>
                      <p className="text-xs text-slate-400 mt-1">
                        Hanya Mas Hashfi yang berwenang mengisi dan merilis Kode Job Card.
                      </p>
                    </div>

                    <div>
                      <input
                        type="password"
                        value={pinInput}
                        onChange={(e) => setPinInput(e.target.value)}
                        placeholder="Ketik PIN (Default: 2026)..."
                        autoFocus
                        required
                        className={`w-full px-4 py-3 bg-slate-900 border rounded-xl text-center text-base font-mono tracking-widest text-white focus:outline-none focus:ring-2 ${
                          pinError ? 'border-rose-500 focus:ring-rose-500' : 'border-slate-700 focus:ring-amber-500'
                        }`}
                      />
                      {pinError && (
                        <span className="text-xs text-rose-400 mt-2 flex items-center justify-center gap-1">
                          <ShieldAlert className="w-3.5 h-3.5" /> PIN salah! Silakan coba lagi (PIN default: 2026).
                        </span>
                      )}
                    </div>

                    <button
                      type="submit"
                      className="w-full py-3 bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold rounded-xl transition cursor-pointer shadow-lg shadow-amber-950/50"
                    >
                      Buka Akses Planner
                    </button>
                  </form>
                ) : (
                  /* Antrean Job Card User yang Sudah Disubmit */
                  <div className="space-y-5 animate-fadeIn">
                    <div className="bg-emerald-500/10 border border-emerald-500/20 p-3 rounded-xl flex items-center justify-between text-xs text-emerald-300">
                      <span className="flex items-center gap-1.5">
                        <CheckCircle2 className="w-4 h-4" /> Mas Hashfi Terhubung — {selectedFormBiro?.biroName}
                      </span>
                      <button
                        onClick={() => setIsPlannerUnlocked(false)}
                        className="text-[11px] underline hover:text-white cursor-pointer"
                      >
                        Kunci Kembali
                      </button>
                    </div>

                    <div>
                      <div className="flex items-center justify-between text-xs text-slate-400 mb-2">
                        <span>Daftar Pengajuan Job Card Masuk ({currentBiroSubmittedTasks.length})</span>
                        <span className="text-[11px] text-amber-400">*Cukup ketik Kode JC satu kali pada baris terkait</span>
                      </div>

                      {currentBiroSubmittedTasks.length > 0 ? (
                        <div className="space-y-3">
                          {currentBiroSubmittedTasks.map((task, idx) => {
                            const isTaskApproved = Boolean(task.kodeJc);
                            const currentInputVal = editingTaskKode[task.id] ?? task.kodeJc ?? '';

                            return (
                              <div 
                                key={task.id}
                                className={`p-4 rounded-xl border transition ${
                                  isTaskApproved 
                                    ? 'bg-slate-900/60 border-slate-700' 
                                    : 'bg-amber-950/20 border-amber-500/40 shadow-sm'
                                }`}
                              >
                                <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                                  <div className="space-y-1">
                                    <div className="flex items-center gap-2">
                                      <span className="text-xs font-mono font-bold text-slate-400">#{idx + 1}</span>
                                      <span className="font-bold text-white text-sm">{task.pic}</span>
                                      <span className="px-2 py-0.5 bg-blue-500/10 text-blue-300 text-[10px] font-semibold rounded border border-blue-500/20">
                                        Proyek: {task.project}
                                      </span>
                                      <span className="px-2 py-0.5 bg-purple-500/10 text-purple-300 text-[10px] font-semibold rounded border border-purple-500/20 font-mono">
                                        JO: #{task.jo}
                                      </span>
                                    </div>
                                    <p className="text-xs text-slate-200 leading-relaxed font-medium">
                                      {task.taskName}
                                    </p>
                                    <span className="text-[10px] text-slate-400 font-mono block">
                                      Jadwal: {task.startDate} s/d {task.endDate}
                                    </span>
                                  </div>

                                  {/* Input Kode JC oleh Mas Hashfi */}
                                  <div className="shrink-0 flex items-center gap-2 self-end sm:self-center">
                                    <input
                                      type="text"
                                      value={currentInputVal}
                                      onChange={(e) => setEditingTaskKode(prev => ({
                                        ...prev,
                                        [task.id]: e.target.value.toUpperCase()
                                      }))}
                                      placeholder="Contoh: JC300426 0001"
                                      className="w-40 px-3 py-1.5 bg-slate-900 border border-slate-600 rounded-lg text-xs font-mono text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-amber-500 uppercase"
                                    />
                                    <button
                                      onClick={() => handleSaveKodeJcForTask(task.id, selectedFormBiro!.biroName)}
                                      className="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 cursor-pointer"
                                      title="Simpan Kode JC"
                                    >
                                      <Check className="w-3.5 h-3.5" />
                                      Simpan
                                    </button>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      ) : (
                        <div className="py-12 text-center text-slate-400 text-xs bg-slate-900/40 border border-slate-700/60 rounded-xl">
                          Belum ada form yang disubmit oleh user pada biro ini.
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* LEVEL 3: TABEL KPI 11 KOLOM PRESISI */}
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