import { useState, useEffect, useCallback, useMemo, ChangeEvent } from 'react';
import { departmentsData, monthList, Department } from './data';
import * as XLSX from 'xlsx';
import { supabase } from './lib/supabase';

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
  FileText, 
  Layers, 
  User, 
  ChevronDown, 
  ChevronUp, 
  Send, 
  Trash2, 
  Lock, 
  KeyRound, 
  X, 
  Check, 
  RotateCcw, 
  Briefcase, 
  HardHat, 
  Upload,
  FileSpreadsheet,
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

const csvMonthMap: Record<string, string> = {
  januari: csvJan,
  februari: csvFeb,
  maret: csvMar,
  april: csvApr,
  mei: csvMei,
  juni: csvJun,
};

const allCsvFiles = import.meta.glob('./**/*.{csv,CSV,txt,TXT}', { 
  query: '?raw', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

// Hanya membaca 3 file Excel utama
const excelGlobUrls = import.meta.glob('./*.xlsx', { 
  query: '?url', 
  import: 'default', 
  eager: true 
}) as Record<string, string>;

export interface ParsedMember {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
  biro: string;
  dept: string;
}

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
  kodeJc: string;
}

interface PersonilCardGroup {
  picName: string;
  status: string;
  jabatan: string;
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
  const b1 = (biro1 || '').toLowerCase().replace('&', ' dan ').trim();
  const b2 = (biro2 || '').toLowerCase().replace('&', ' dan ').trim();
  if (!b1 || !b2) return false;
  if (b1 === b2) return true;

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

  const c1 = cleanText(b1.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  const c2 = cleanText(b2.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  if (c1 && c2) {
    return c1 === c2 || c1.includes(c2) || c2.includes(c1);
  }
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
      // next
    }
  }
  return null;
}

export default function App() {
  const [accessMode, setAccessMode] = useState<'landing' | 'organik' | 'subkon'>('landing');

  // ================= STATE KHUSUS SUBKON =================
  const [subconSelectedDept, setSubconSelectedDept] = useState<Department | null>(null);
  const [subconSelectedBiro, setSubconSelectedBiro] = useState<string | null>(null);
  const [subconPageMode, setSubconPageMode] = useState<'members' | 'form' | 'output'>('members');
  const [subconSearchQuery, setSubconSearchQuery] = useState('');

  // ================= STATE KHUSUS ORGANIK =================
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroPage, setSelectedBiroPage] = useState<SelectedBiroPage | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  const [formPageMode, setFormPageMode] = useState<'form' | 'output'>('form');

  const [searchQuery, setSearchQuery] = useState('');
  const [tableSearch, setTableSearch] = useState('');
  const [outputSearch, setOutputSearch] = useState('');
  
  const [formData, setFormData] = useState({
    nama: '',
    kodeProyek: '',
    taskName: '',
    startDate: '',
    endDate: '',
    pic: '',
    jo: ''
  });

  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>({});

  const loadAllJobCards = useCallback(async () => {
    try {
      const { data, error } = await supabase
        .from('job_cards')
        .select('*')
        .order('created_at', { ascending: false });

      if (error) return;

      if (data) {
        const grouped: { [biroKey: string]: TaskItem[] } = {};
        data.forEach((row: any) => {
          const biroKey = cleanText(row.biro_name || '');
          if (!biroKey) return;
          if (!grouped[biroKey]) grouped[biroKey] = [];
          grouped[biroKey].push({
            id: row.id,
            biroName: row.biro_name || '',
            project: row.project || '',
            taskName: row.task_name || '',
            startDate: row.start_date || '',
            endDate: row.end_date || '',
            pic: row.pic || '',
            jo: row.jo || '',
            kodeJc: row.kode_jc || '',
          });
        });
        setManualTasks(grouped);
      }
    } catch {
      // fallback
    }
  }, []);

  const [isPlannerModalOpen, setIsPlannerModalOpen] = useState(false);
  const [isPlannerUnlocked, setIsPlannerUnlocked] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [pinError, setPinError] = useState(false);
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});

  const PLANNER_PIN = '2026';
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  // HANYA 3 WORKBOOK UTAMA (FILE NOMOR 4 SUDAH DIHAPUS)
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);          // data_kpi.xlsx
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);  // JOBCARD_DESAIN.xlsx
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);          // AKSES AKUN IM4...xlsx
  const [isLoadingExcel, setIsLoadingExcel] = useState<boolean>(true);

  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        setIsLoadingExcel(true);
        let kpiUrl = '';
        let jcUrl = '';
        let im4Url = '';

        Object.entries(excelGlobUrls).forEach(([path, url]) => {
          const pLower = path.toLowerCase();
          if (pLower.includes('kpi')) kpiUrl = url;
          else if (pLower.includes('jobcard')) jcUrl = url;
          else if (pLower.includes('im4') || pLower.includes('drawing') || pLower.includes('akses')) im4Url = url;
        });

        const [wbKpi, wbJc, wbIm4] = await Promise.all([
          fetchSafeWorkbook([kpiUrl, '/data_kpi.xlsx', './data_kpi.xlsx']),
          fetchSafeWorkbook([jcUrl, '/JOBCARD_DESAIN.xlsx', './JOBCARD_DESAIN.xlsx', '/JOBCARD DESAIN.xlsx']),
          fetchSafeWorkbook([
            im4Url,
            '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx',
            './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx',
            '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL.xlsx',
            './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL.xlsx'
          ])
        ]);

        if (wbKpi) setWorkbook(wbKpi);
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
      } catch {
        // fallback
      } finally {
        setIsLoadingExcel(false);
      }
    }
    loadAllExcelFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

  const handleManualUploadExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        setIm4Workbook(wb);
        alert(`File ${file.name} berhasil dibaca! Anggota outsourcing dan organik langsung diperbarui.`);
      } catch {
        alert('Gagal membaca berkas Excel. Pastikan format file .xlsx');
      }
    };
    reader.readAsBinaryString(file);
  };

  // =========================================================================
  // PARSER DINAMIS DARI FILE IM4 (MEMBACA SELURUH 197 ANGGOTA SECARA OTOMATIS)
  // =========================================================================
  const allParsedFromExcel = useMemo<ParsedMember[]>(() => {
    if (!im4Workbook) return [];

    const sheetName = im4Workbook.SheetNames.find(s => s.toLowerCase().includes('education')) || im4Workbook.SheetNames[0];
    const sheet = im4Workbook.Sheets[sheetName];
    if (!sheet) return [];

    const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    
    let headerIdx = -1;
    for (let r = 0; r < Math.min(15, rawRows.length); r++) {
      const rowVals = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
      if (rowVals.includes('nama') && (rowVals.includes('nip') || rowVals.includes('status') || rowVals.includes('jabatan'))) {
        headerIdx = r;
        break;
      }
    }

    if (headerIdx === -1) return [];

    const headers = rawRows[headerIdx].map(v => String(v).trim().toLowerCase());
    const namaCol = headers.findIndex(h => h === 'nama');
    const nipCol = headers.findIndex(h => h === 'nip');
    const statusCol = headers.findIndex(h => h === 'status');
    const unitCol = headers.findIndex(h => h.includes('unit'));
    const jabatanCol = headers.findIndex(h => h.includes('jabatan'));

    let currentDept = '';
    let currentBiro = '';
    const results: ParsedMember[] = [];

    for (let r = headerIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row) continue;

      const nama = String(row[namaCol] || '').trim();
      if (!nama || nama.toLowerCase() === 'nan' || nama.toLowerCase() === 'nama') continue;

      const nip = String(row[nipCol] || '').trim();
      const statusRaw = String(row[statusCol] || '').trim();
      const unit = String(row[unitCol] || '').trim();
      const jabatan = String(row[jabatanCol] || '').trim();

      if (jabatan.toLowerCase().includes('kepala divisi')) {
        currentDept = 'Div. Desain';
        currentBiro = 'Div. Desain';
      } else if (jabatan.toLowerCase().includes('kepala departemen') || jabatan.toLowerCase().includes('kadep')) {
        if (unit && unit.toLowerCase() !== 'nan') currentDept = unit;
        else currentDept = jabatan;
        currentBiro = `Staf ${currentDept}`;
      } else if (jabatan.toLowerCase().includes('kepala biro') || jabatan.toLowerCase().includes('kabiro')) {
        currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').replace(/Kabiro/gi, 'Biro').trim();
        if (unit && unit.toLowerCase() !== 'nan') currentDept = unit;
      }

      results.push({
        nama,
        nip,
        status: statusRaw || 'PKWTT',
        jabatan,
        biro: currentBiro,
        dept: currentDept
      });
    }

    return results;
  }, [im4Workbook]);

  // Saring Anggota Khusus Outsourcing
  const dynamicOutsourcingList = useMemo(() => {
    return allParsedFromExcel.filter(p => p.status.toLowerCase().includes('outsourcing'));
  }, [allParsedFromExcel]);

  const getSubconMembersForBiro = useCallback((biroName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.biro, biroName));
  }, [dynamicOutsourcingList]);

  const getSubconCountForDept = useCallback((deptName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.dept, deptName)).length;
  }, [dynamicOutsourcingList]);

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
              if (isOvertimeFile) current.overtime += numVal;
              else current.reguler += numVal;
              timesheetMap.set(cleanNameKey, current);
            }
          }
        });
      }
    });

    return timesheetMap;
  };

  // Mengambil anggota biro (PKWTT, PKWT, Outsourcing) langsung dari file IM4
  const getBiroMembers = (biroName: string): { nama: string; status: string; jabatan: string }[] => {
    const members = allParsedFromExcel.filter(p => isBiroMatch(p.biro, biroName));
    if (members.length > 0) {
      return members.map(m => ({ nama: m.nama, status: m.status, jabatan: m.jabatan }));
    }
    return [];
  };

  const getJobcardProjects = (): string[] => {
    if (!jobcardWorkbook) return [];
    const projects = new Set<string>();
    jobcardWorkbook.SheetNames.forEach(sheetName => {
      const sheet = jobcardWorkbook.Sheets[sheetName];
      if (!sheet) return;
      const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      rows.forEach((row, idx) => {
        if (idx < 2 || !row) return;
        const p = String(row[2] || '').trim();
        if (p && p.toLowerCase() !== 'nan' && !p.toLowerCase().includes('kode proyek') && !p.toLowerCase().includes('project')) {
          projects.add(p);
        }
      });
    });
    return Array.from(projects).sort();
  };

  const getJobcardTasks = (): string[] => {
    if (!jobcardWorkbook) return [];
    const tasks = new Set<string>();
    jobcardWorkbook.SheetNames.forEach(sheetName => {
      const sheet = jobcardWorkbook.Sheets[sheetName];
      if (!sheet) return;
      const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      rows.forEach((row, idx) => {
        if (idx < 2 || !row) return;
        const t = String(row[3] || '').trim();
        if (t && t.toLowerCase() !== 'nan' && !t.toLowerCase().includes('desc pekerjaan') && !t.toLowerCase().includes('task') && !t.toLowerCase().includes('uraian')) {
          tasks.add(t);
        }
      });
    });
    return Array.from(tasks).sort();
  };

  const getAccordionOutputForBiro = (targetBiroName: string): PersonilCardGroup[] => {
    const biroMembers = getBiroMembers(targetBiroName);
    const personMap = new Map<string, { status: string; jabatan: string; tasks: TaskItem[] }>();

    biroMembers.forEach(m => {
      personMap.set(m.nama, { status: m.status, jabatan: m.jabatan, tasks: [] });
    });

    const biroKey = cleanText(targetBiroName);
    const tasksForThisBiro = manualTasks[biroKey] || [];

    tasksForThisBiro.forEach(t => {
      const matched = biroMembers.find(m => cleanText(m.nama) === cleanText(t.pic));
      const key = matched ? matched.nama : t.pic;
      if (!personMap.has(key)) {
        personMap.set(key, { 
          status: matched?.status || 'Organik', 
          jabatan: matched?.jabatan || '', 
          tasks: [] 
        });
      }
      personMap.get(key)!.tasks.push(t);
    });

    const result: PersonilCardGroup[] = [];
    personMap.forEach((val, picName) => {
      result.push({ 
        picName, 
        status: val.status, 
        jabatan: val.jabatan, 
        tasks: val.tasks 
      });
    });

    return result.sort((a, b) => a.picName.localeCompare(b.picName));
  };

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    const activeBiroName = accessMode === 'subkon' ? subconSelectedBiro : selectedFormBiro?.biroName;
    if (!activeBiroName) return;

    try {
      let validBiroId: string | null = null;
      const { data: biroList } = await supabase.from('biros').select('id, name');
      
      if (biroList && biroList.length > 0) {
        const found = biroList.find(b => isBiroMatch(b.name, activeBiroName));
        validBiroId = found ? found.id : biroList[0].id;
      }

      const { data: insertedRow, error } = await supabase
        .from('job_cards')
        .insert({
          biro_id: validBiroId,
          biro_name: activeBiroName,
          personil_name: formData.nama,
          project_code: formData.kodeProyek,
          project: formData.kodeProyek,
          task_name: formData.taskName,
          start_date: formData.startDate,
          end_date: formData.endDate,
          pic: formData.nama,
          jo: formData.jo,
          kode_jc: '',
          status: 'pending',
        })
        .select()
        .single();

      if (error) {
        alert('Gagal menyimpan ke database: ' + error.message);
        return;
      }

      const biroKey = cleanText(activeBiroName);
      const newTask: TaskItem = {
        id: insertedRow.id,
        biroName: activeBiroName,
        project: formData.kodeProyek,
        taskName: formData.taskName,
        startDate: formData.startDate,
        endDate: formData.endDate,
        pic: formData.nama,
        jo: formData.jo,
        kode_jc: '',
      };

      const currentList = manualTasks[biroKey] || [];
      setManualTasks({ ...manualTasks, [biroKey]: [newTask, ...currentList] });

      setExpandedCards(prev => ({ ...prev, [formData.nama]: true }));
      setFormData({
        nama: '',
        kodeProyek: '',
        taskName: '',
        startDate: '',
        endDate: '',
        pic: '',
        jo: ''
      });

      alert('Job Card berhasil tersimpan ke database online!');
      loadAllJobCards();
    } catch {
      alert('Terjadi kesalahan koneksi database.');
    }
  };

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

  const handleSaveKodeJcForTask = async (taskId: string, biroName: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;

    const { error } = await supabase
      .from('job_cards')
      .update({ kode_jc: inputVal, status: 'approved' })
      .eq('id', taskId);

    if (error) {
      alert('Gagal update kode JC: ' + error.message);
      return;
    }

    const biroKey = cleanText(biroName);
    const currentList = manualTasks[biroKey] || [];
    const updatedList = currentList.map(task => task.id === taskId ? { ...task, kodeJc: inputVal } : task);

    setManualTasks({ ...manualTasks, [biroKey]: updatedList });

    setEditingTaskKode(prev => {
      const next = { ...prev };
      delete next[taskId];
      return next;
    });

    loadAllJobCards();
  };

  const handleDeleteTask = async (taskId: string, biroName: string) => {
    if (window.confirm('Hapus tugas ini?')) {
      const { error } = await supabase
        .from('job_cards')
        .delete()
        .eq('id', taskId);

      if (error) {
        alert('Gagal menghapus: ' + error.message);
        return;
      }

      const biroKey = cleanText(biroName);
      const currentList = manualTasks[biroKey] || [];
      const updatedList = currentList.filter(t => t.id !== taskId);
      setManualTasks({ ...manualTasks, [biroKey]: updatedList });
      loadAllJobCards();
    }
  };

  const handleClearAllBiroData = async (targetBiroName: string) => {
    if (!targetBiroName) return;
    if (window.confirm(`Kosongkan semua data tugas di ${targetBiroName}?`)) {
      const { error } = await supabase
        .from('job_cards')
        .delete()
        .eq('biro_name', targetBiroName);

      if (error) {
        alert('Gagal mengosongkan data: ' + error.message);
        return;
      }

      const biroKey = cleanText(targetBiroName);
      setManualTasks({ ...manualTasks, [biroKey]: [] });
      loadAllJobCards();
    }
  };

  const toggleAccordion = (picName: string) => {
    setExpandedCards(prev => ({ ...prev, [picName]: !prev[picName] }));
  };

  const handleMonthClick = (biroName: string, month: string) => {
    if (!workbook) {
      alert('File data_kpi.xlsx belum terbaca.');
      return;
    }

    const monthPrefix = month.toLowerCase().slice(0, 3);
    const targetSheetName = workbook.SheetNames.find(sheet => {
      const sLower = sheet.toLowerCase().trim();
      return sLower.includes(month.toLowerCase()) || sLower.includes(monthPrefix);
    });

    if (!targetSheetName) {
      alert(`Sheet bulan ${month} tidak ditemukan.`);
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

  const filteredSubconDepartments = (departmentsData || []).filter((dept) => 
    dept.name.toLowerCase().includes(subconSearchQuery.toLowerCase()) ||
    (dept.description || '').toLowerCase().includes(subconSearchQuery.toLowerCase())
  );

  const filteredTableData = (selectedBiroPage?.data || []).filter(item => 
    item.nip.toLowerCase().includes(tableSearch.toLowerCase()) ||
    item.nama.toLowerCase().includes(tableSearch.toLowerCase())
  );

  const currentBiroMembers = selectedFormBiro ? getBiroMembers(selectedFormBiro.biroName) : [];
  const projectOptions = getJobcardProjects();
  const taskOptions = getJobcardTasks();
  
  const accordionData = selectedFormBiro ? getAccordionOutputForBiro(selectedFormBiro.biroName) : [];
  const filteredAccordionData = accordionData.filter(g => {
    const s = outputSearch.toLowerCase();
    return g.picName.toLowerCase().includes(s) || 
           (g.status || '').toLowerCase().includes(s) || 
           g.tasks.some(t => t.project.toLowerCase().includes(s) || t.taskName.toLowerCase().includes(s) || (t.kodeJc || '').toLowerCase().includes(s));
  });

  const totalPersonilCount = accordionData.length;
  const countOutsourcing = accordionData.filter(a => a.status.toLowerCase().includes('outsourcing')).length;
  const countOrganik = totalPersonilCount - countOutsourcing;
  const totalTasksCount = accordionData.reduce((acc, g) => acc + g.tasks.length, 0);

  const currentBiroKey = selectedFormBiro ? cleanText(selectedFormBiro.biroName) : '';
  const currentBiroSubmittedTasks = manualTasks[currentBiroKey] || [];
  const pendingTasksCount = currentBiroSubmittedTasks.filter(t => !t.kodeJc).length;

  // DATA KHUSUS BIRO SUBCON TERPILIH
  const activeSubconMembers = subconSelectedBiro ? getSubconMembersForBiro(subconSelectedBiro) : [];
  const activeSubconKey = subconSelectedBiro ? cleanText(subconSelectedBiro) : '';
  const activeSubconSubmittedTasks = manualTasks[activeSubconKey] || [];
  const activeSubconPendingCount = activeSubconSubmittedTasks.filter(t => !t.kodeJc).length;

  const subconAccordionData = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const members = getSubconMembersForBiro(subconSelectedBiro);
    const pMap = new Map<string, { status: string; jabatan: string; tasks: TaskItem[] }>();

    members.forEach(m => {
      pMap.set(m.nama, { status: 'Outsourcing', jabatan: m.jabatan, tasks: [] });
    });

    activeSubconSubmittedTasks.forEach(t => {
      const found = members.find(m => cleanText(m.nama) === cleanText(t.pic));
      const key = found ? found.nama : t.pic;
      if (!pMap.has(key)) {
        pMap.set(key, { status: 'Outsourcing', jabatan: '', tasks: [] });
      }
      pMap.get(key)!.tasks.push(t);
    });

    const res: PersonilCardGroup[] = [];
    pMap.forEach((v, picName) => {
      res.push({ picName, status: v.status, jabatan: v.jabatan, tasks: v.tasks });
    });
    return res.sort((a, b) => a.picName.localeCompare(b.picName));
  }, [subconSelectedBiro, getSubconMembersForBiro, activeSubconSubmittedTasks]);

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans antialiased">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-14 flex items-center justify-between">
          <div 
            className="flex items-center gap-3 cursor-pointer select-none" 
            onClick={() => {
              setAccessMode('landing');
              setSelectedFormBiro(null);
              setSelectedBiroPage(null);
              setSelectedDept(null);
              setSubconSelectedBiro(null);
              setSubconSelectedDept(null);
            }}
          >
            <div className="p-1.5 bg-blue-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-sm tracking-wide text-white block leading-tight">DIVISI DESAIN</span>
              <span className="text-[10px] text-slate-400">
                {accessMode === 'landing' ? 'Portal Sistem PT PAL' : accessMode === 'organik' ? 'Portal Pegawai Organik' : 'Portal Rekanan & Subkontraktor'}
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {accessMode !== 'landing' && (
              <button
                onClick={() => {
                  setAccessMode('landing');
                  setSelectedFormBiro(null);
                  setSelectedBiroPage(null);
                  setSelectedDept(null);
                  setSubconSelectedBiro(null);
                  setSubconSelectedDept(null);
                }}
                className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                Ganti Akses
              </button>
            )}

            {/* Tombol Back Dinamis untuk Subkon */}
            {accessMode === 'subkon' && subconSelectedBiro ? (
              <button
                onClick={() => setSubconSelectedBiro(null)}
                className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'subkon' && subconSelectedDept ? (
              <button
                onClick={() => setSubconSelectedDept(null)}
                className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" /> Departemen
              </button>
            ) : null}

            {/* Tombol Back Dinamis untuk Organik */}
            {accessMode === 'organik' && selectedFormBiro ? (
              <button
                onClick={() => setSelectedFormBiro(null)}
                className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'organik' && selectedBiroPage ? (
              <button
                onClick={() => setSelectedBiroPage(null)}
                className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'organik' && selectedDept ? (
              <button
                onClick={() => setSelectedDept(null)}
                className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" /> Departemen
              </button>
            ) : null}
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">

        {/* ========================================================================= */}
        {/* TAMPILAN 0: HALAMAN PERTAMA (PORTAL PEMILIHAN AKSES SUBKON & ORGANIK)      */}
        {/* ========================================================================= */}
        {accessMode === 'landing' && (
          <div className="max-w-4xl mx-auto py-12 space-y-10 animate-fadeIn">
            <div className="text-center space-y-3">
              <span className="px-3 py-1 bg-blue-500/10 text-blue-400 text-xs font-semibold rounded-full border border-blue-500/20">
                PT PAL INDONESIA • DIVISI DESAIN 2026
              </span>
              <h1 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight">
                Pilih Akses Portal Sistem
              </h1>
              <p className="text-slate-400 text-sm max-w-lg mx-auto">
                Silakan pilih kategori entitas kerja Anda untuk melanjutkan ke modul Job Card dan Rekapitulasi Kerja.
              </p>
            </div>

            {/* Pilihan 2 Kartu: Organik vs Subkon */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4">
              {/* Opsi 1: Organik */}
              <div
                onClick={() => setAccessMode('organik')}
                className="group relative bg-slate-900 border border-slate-800 hover:border-blue-500/60 rounded-2xl p-7 cursor-pointer transition-all duration-300 hover:shadow-xl hover:shadow-blue-500/10 flex flex-col justify-between"
              >
                <div className="space-y-4">
                  <div className="w-14 h-14 rounded-xl bg-blue-600/10 border border-blue-500/20 flex items-center justify-center text-blue-400 group-hover:bg-blue-600 group-hover:text-white transition-all duration-300">
                    <Briefcase className="w-7 h-7" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-white group-hover:text-blue-400 transition-colors">
                      Pegawai Organik
                    </h3>
                    <span className="text-xs text-blue-400 font-mono">Divisi Desain PT PAL (PKWTT, PKWT)</span>
                  </div>
                  <p className="text-slate-400 text-xs leading-relaxed">
                    Akses 6 Departemen, 19 Biro, Formulir Pengajuan & Verifikasi Job Card Planner, Output Rekapitulasi Personil Organik, serta Evaluasi KPI Bulanan.
                  </p>
                </div>

                <div className="mt-8 pt-4 border-t border-slate-800 flex items-center justify-between text-xs font-semibold text-blue-400">
                  <span>Masuk Portal Organik</span>
                  <ChevronRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
                </div>
              </div>

              {/* Opsi 2: Subkon */}
              <div
                onClick={() => setAccessMode('subkon')}
                className="group relative bg-slate-900 border border-slate-800 hover:border-amber-500/60 rounded-2xl p-7 cursor-pointer transition-all duration-300 hover:shadow-xl hover:shadow-amber-500/10 flex flex-col justify-between"
              >
                <div className="space-y-4">
                  <div className="w-14 h-14 rounded-xl bg-amber-600/10 border border-amber-500/20 flex items-center justify-center text-amber-400 group-hover:bg-amber-600 group-hover:text-white transition-all duration-300">
                    <HardHat className="w-7 h-7" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-white group-hover:text-amber-400 transition-colors">
                      Mitra / Subkontraktor
                    </h3>
                    <span className="text-xs text-amber-400 font-mono">Daftar Anggota Outsourcing per Departemen & Biro</span>
                  </div>
                  <p className="text-slate-400 text-xs leading-relaxed">
                    Akses terkelompok per Departemen dan Biro untuk seluruh personel Outsourcing (Drafter & Desainer) dibaca dinamis dari file Excel IM4.
                  </p>
                </div>

                <div className="mt-8 pt-4 border-t border-slate-800 flex items-center justify-between text-xs font-semibold text-amber-400">
                  <span>Masuk Portal Subkon</span>
                  <ChevronRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAMPILAN MODUL SUBKON: DIKELOMPOKKAN PER DEPARTEMEN -> BIRO -> ANGGOTA   */}
        {/* ========================================================================= */}
        {accessMode === 'subkon' && (
          <div className="space-y-6 animate-fadeIn">
            {/* SUBKON LEVEL 1: PILIH DEPARTEMEN */}
            {!subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-6">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-800 pb-4">
                  <div>
                    <h2 className="text-xl font-bold text-white flex items-center gap-2">
                      <HardHat className="w-5 h-5 text-amber-400" />
                      Departemen Desain (Portal Subkontraktor)
                    </h2>
                    <span className="text-xs text-slate-400">
                      Data {dynamicOutsourcingList.length} personel outsourcing otomatis terurai dari file AKSES AKUN IM4
                    </span>
                  </div>

                  <div className="flex items-center gap-3">
                    <label className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold rounded-lg border border-slate-700 transition flex items-center gap-1.5 cursor-pointer">
                      <Upload className="w-3.5 h-3.5 text-amber-400" />
                      <span>{dynamicOutsourcingList.length > 0 ? 'Ganti File Excel' : 'Pilih File Excel IM4'}</span>
                      <input 
                        type="file" 
                        accept=".xlsx, .xls" 
                        onChange={handleManualUploadExcel}
                        className="hidden" 
                      />
                    </label>

                    <div className="relative w-64">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Cari departemen..."
                        value={subconSearchQuery}
                        onChange={(e) => setSubconSearchQuery(e.target.value)}
                        className="w-full pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                      />
                    </div>
                  </div>
                </div>

                {dynamicOutsourcingList.length === 0 && !isLoadingExcel && (
                  <div className="p-4 bg-amber-500/10 border border-amber-500/30 rounded-xl text-xs text-amber-300 flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="w-4 h-4 text-amber-400" />
                      <span>File Excel IM4 belum terbaca otomatis. Silakan klik tombol "Pilih File Excel IM4" untuk memuat berkas.</span>
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {filteredSubconDepartments.map((dept) => {
                    const IconComponent = iconMap[dept.icon] || Building2;
                    const countSubcon = getSubconCountForDept(dept.name);

                    return (
                      <div
                        key={dept.id}
                        onClick={() => setSubconSelectedDept(dept)}
                        className="bg-slate-900 border border-slate-800 hover:border-amber-500/60 rounded-xl p-5 cursor-pointer transition flex flex-col justify-between hover:shadow-lg hover:shadow-amber-500/5 group"
                      >
                        <div>
                          <div className="flex items-center justify-between mb-3">
                            <div className="p-2.5 bg-amber-500/10 text-amber-400 rounded-lg group-hover:bg-amber-500 group-hover:text-slate-950 transition-colors">
                              <IconComponent className="w-5 h-5" />
                            </div>
                            <span className="text-[10px] font-mono px-2 py-0.5 bg-slate-800 text-amber-400 rounded border border-amber-500/20 font-bold">
                              {countSubcon} Subkon
                            </span>
                          </div>
                          <h3 className="font-semibold text-white text-base group-hover:text-amber-400 transition-colors">
                            {dept.name}
                          </h3>
                        </div>

                        <div className="mt-4 pt-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
                          <span>{dept.biros.length} Biro</span>
                          <span className="text-amber-400 flex items-center gap-0.5 font-medium group-hover:translate-x-1 transition-transform">
                            Buka Biro <ChevronRight className="w-3.5 h-3.5" />
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* SUBKON LEVEL 2: DAFTAR BIRO DI BAWAH DEPARTEMEN TERPILIH */}
            {subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-5 animate-fadeIn">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                  <div>
                    <h2 className="text-lg font-bold text-white flex items-center gap-2">
                      <Building2 className="w-5 h-5 text-amber-400" />
                      {subconSelectedDept.name}
                    </h2>
                    <span className="text-xs text-slate-400">
                      {subconSelectedDept.biros.length} Biro Terdaftar • Total {getSubconCountForDept(subconSelectedDept.name)} Personel Outsourcing (Dinamis Excel)
                    </span>
                  </div>
                </div>

                <div className="space-y-3">
                  {subconSelectedDept.biros.map((biro) => {
                    const biroSubconMembers = getSubconMembersForBiro(biro.name);
                    const countInBiro = biroSubconMembers.length;

                    return (
                      <div
                        key={biro.id}
                        className="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:border-slate-700 transition"
                      >
                        <div className="flex items-center gap-3 flex-wrap">
                          <h4 className="text-sm font-semibold text-white">{biro.name}</h4>
                          <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${
                            countInBiro > 0 
                              ? 'bg-amber-500/15 text-amber-400 border-amber-500/30' 
                              : 'bg-slate-800 text-slate-500 border-slate-700'
                          }`}>
                            {countInBiro} Anggota Outsourcing
                          </span>
                        </div>

                        <div className="flex items-center gap-2">
                          <button
                            onClick={() => {
                              setSubconSelectedBiro(biro.name);
                              setSubconPageMode('members');
                            }}
                            className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold rounded-lg border border-slate-700 transition flex items-center gap-1.5 cursor-pointer"
                          >
                            <Users className="w-3.5 h-3.5 text-amber-400" /> ANGGOTA ({countInBiro})
                          </button>

                          <button
                            onClick={() => {
                              setSubconSelectedBiro(biro.name);
                              setSubconPageMode('form');
                            }}
                            className="px-3 py-1.5 bg-emerald-600/90 hover:bg-emerald-600 text-white text-xs font-bold rounded-lg transition flex items-center gap-1 cursor-pointer"
                          >
                            <FileText className="w-3.5 h-3.5" /> FORM
                          </button>

                          <button
                            onClick={() => {
                              setSubconSelectedBiro(biro.name);
                              setSubconPageMode('output');
                            }}
                            className="px-3 py-1.5 bg-blue-600/90 hover:bg-blue-600 text-white text-xs font-bold rounded-lg transition flex items-center gap-1 cursor-pointer"
                          >
                            <Layers className="w-3.5 h-3.5" /> OUTPUT
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* SUBKON LEVEL 3: HALAMAN BIRO (TABEL ANGGOTA / FORM / OUTPUT) */}
            {subconSelectedBiro && (
              <div className="space-y-4 animate-fadeIn">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-base font-bold text-white flex items-center gap-2">
                      <HardHat className="w-4 h-4 text-amber-400" />
                      {subconSelectedBiro}
                    </h2>
                    <span className="text-xs text-slate-400 font-mono">
                      {subconSelectedDept?.name} • {activeSubconMembers.length} Anggota Outsourcing • {activeSubconSubmittedTasks.length} Tugas
                    </span>
                  </div>

                  <div className="flex items-center gap-2">
                    <div className="bg-slate-950 p-1 rounded-lg border border-slate-800 flex gap-1">
                      <button
                        onClick={() => setSubconPageMode('members')}
                        className={`px-3 py-1 text-xs font-semibold rounded-md transition cursor-pointer ${
                          subconPageMode === 'members' ? 'bg-amber-600 text-white' : 'text-slate-400 hover:text-white'
                        }`}
                      >
                        Anggota ({activeSubconMembers.length})
                      </button>
                      <button
                        onClick={() => setSubconPageMode('form')}
                        className={`px-3 py-1 text-xs font-semibold rounded-md transition cursor-pointer ${
                          subconPageMode === 'form' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
                        }`}
                      >
                        Form
                      </button>
                      <button
                        onClick={() => setSubconPageMode('output')}
                        className={`px-3 py-1 text-xs font-semibold rounded-md transition cursor-pointer ${
                          subconPageMode === 'output' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                        }`}
                      >
                        Output ({activeSubconSubmittedTasks.length})
                      </button>
                    </div>

                    <button
                      type="button"
                      onClick={() => {
                        setIsPlannerModalOpen(true);
                        setPinError(false);
                      }}
                      className="px-2.5 py-1 bg-amber-600/90 hover:bg-amber-600 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer"
                    >
                      <Lock className="w-3 h-3" /> Planner
                      {activeSubconPendingCount > 0 && (
                        <span className="px-1.5 py-0.2 bg-rose-600 text-[10px] font-bold rounded-full">
                          {activeSubconPendingCount}
                        </span>
                      )}
                    </button>
                  </div>
                </div>

                {/* VIEW 1: TABEL DAFTAR ANGGOTA OUTSOURCING DI BIRO INI */}
                {subconPageMode === 'members' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-sm">
                    <div className="p-3.5 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-bold text-white flex items-center gap-2">
                        <Users className="w-4 h-4 text-amber-400" />
                        Daftar Anggota Outsourcing {subconSelectedBiro} (Hasil Baca Excel)
                      </span>
                      <span className="text-xs font-mono text-slate-400">
                        {activeSubconMembers.length} Personel Terdaftar
                      </span>
                    </div>

                    <div className="overflow-x-auto">
                      {activeSubconMembers.length > 0 ? (
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-800/80 text-slate-300 border-b border-slate-700 font-semibold">
                            <tr>
                              <th className="py-2.5 px-3 text-center w-12">#</th>
                              <th className="py-2.5 px-3">Nama Personil</th>
                              <th className="py-2.5 px-3 font-mono">NIP</th>
                              <th className="py-2.5 px-3 text-center">Status</th>
                              <th className="py-2.5 px-3">Jabatan</th>
                              <th className="py-2.5 px-3">Biro Penempatan</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-800 text-slate-200">
                            {activeSubconMembers.map((person, idx) => (
                              <tr key={idx} className="hover:bg-slate-800/40">
                                <td className="py-2.5 px-3 text-center text-slate-500 font-mono">{idx + 1}</td>
                                <td className="py-2.5 px-3 font-medium text-white flex items-center gap-2">
                                  <div className="w-6 h-6 rounded bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center font-bold text-[10px]">
                                    {person.nama.charAt(0)}
                                  </div>
                                  {person.nama}
                                </td>
                                <td className="py-2.5 px-3 font-mono text-slate-400">{person.nip}</td>
                                <td className="py-2.5 px-3 text-center">
                                  <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/15 text-amber-400 border border-amber-500/30">
                                    {person.status}
                                  </span>
                                </td>
                                <td className="py-2.5 px-3 text-cyan-300">{person.jabatan}</td>
                                <td className="py-2.5 px-3 text-slate-300 font-medium">{person.biro}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      ) : (
                        <div className="py-12 text-center text-slate-500 text-xs">
                          Biro ini belum memiliki personel outsourcing terdaftar pada file master Excel
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* VIEW 2: FORMULIR JOB CARD KHUSUS ANGGOTA OUTSOURCING BIRO INI */}
                {subconPageMode === 'form' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">
                            Nama Personil Outsourcing ({subconSelectedBiro})
                          </label>
                          {activeSubconMembers.length > 0 ? (
                            <select
                              value={formData.nama}
                              onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                            >
                              <option value="">Pilih Nama Personil...</option>
                              {activeSubconMembers.map((person, idx) => (
                                <option key={idx} value={person.nama}>
                                  {person.nama} ({person.jabatan})
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.nama}
                              onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                              placeholder="Ketik Nama Personil Outsourcing..."
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">Proyek</label>
                          {projectOptions.length > 0 ? (
                            <select
                              value={formData.kodeProyek}
                              onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                            >
                              <option value="">Pilih Proyek...</option>
                              {projectOptions.map((proj, idx) => (
                                <option key={idx} value={proj}>{proj}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.kodeProyek}
                              onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                              placeholder="Kode Proyek..."
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">JO (Angka)</label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formData.jo}
                            onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                            placeholder="Contoh: 300426"
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white font-mono focus:outline-none focus:ring-1 focus:ring-amber-500"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">Task Name / Uraian Gambar</label>
                          {taskOptions.length > 0 ? (
                            <select
                              value={formData.taskName}
                              onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                            >
                              <option value="">Pilih Task...</option>
                              {taskOptions.map((task, idx) => (
                                <option key={idx} value={task}>{task}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.taskName}
                              onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                              placeholder="Uraian Pekerjaan / Task Name..."
                              required
                              className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">Start Date</label>
                          <input
                            type="date"
                            value={formData.startDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">End Date</label>
                          <input
                            type="date"
                            value={formData.endDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">PIC Subkon</label>
                          <input
                            type="text"
                            value={formData.pic}
                            onChange={(e) => setFormData(prev => ({ ...prev, pic: e.target.value }))}
                            placeholder="Nama PIC Subkon..."
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                          />
                        </div>
                      </div>

                      <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                        <button
                          type="button"
                          onClick={() => setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', pic: '', jo: '' })}
                          className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs rounded-lg transition cursor-pointer"
                        >
                          Reset
                        </button>
                        <button
                          type="submit"
                          className="px-5 py-1.5 bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer"
                        >
                          <Send className="w-3.5 h-3.5" /> Simpan Tugas Subkon
                        </button>
                      </div>
                    </form>
                  </div>
                )}

                {/* VIEW 3: OUTPUT REKAPITULASI TUGAS PERSONEL OUTSOURCING */}
                {subconPageMode === 'output' && (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between gap-3">
                      <div className="relative flex-1 max-w-xs">
                        <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                        <input
                          type="text"
                          placeholder="Cari tugas subkon..."
                          value={outputSearch}
                          onChange={(e) => setOutputSearch(e.target.value)}
                          className="w-full pl-8 pr-3 py-1.5 bg-slate-900 border border-slate-800 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-amber-500"
                        />
                      </div>

                      <div className="flex items-center gap-1.5">
                        <button
                          onClick={() => {
                            const all: Record<string, boolean> = {};
                            subconAccordionData.forEach(g => { all[g.picName] = true; });
                            setExpandedCards(all);
                          }}
                          className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded border border-slate-700 cursor-pointer"
                        >
                          Buka
                        </button>
                        <button
                          onClick={() => setExpandedCards({})}
                          className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded border border-slate-700 cursor-pointer"
                        >
                          Tutup
                        </button>
                        <button
                          onClick={() => handleClearAllBiroData(subconSelectedBiro)}
                          className="p-1.5 text-rose-400 hover:text-rose-300 hover:bg-rose-950/40 rounded border border-rose-800/40 cursor-pointer"
                          title="Reset Data"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>

                    {subconAccordionData.length > 0 ? (
                      <div className="space-y-2">
                        {subconAccordionData.map((person, idx) => {
                          const isExpanded = expandedCards[person.picName] ?? false;
                          const taskCount = person.tasks.length;
                          const isActive = taskCount > 0;

                          return (
                            <div
                              key={idx}
                              className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden"
                            >
                              <div
                                onClick={() => toggleAccordion(person.picName)}
                                className="p-3.5 flex items-center justify-between cursor-pointer select-none hover:bg-slate-800/80 transition"
                              >
                                <div className="flex items-center gap-2.5 flex-wrap">
                                  <div className="w-7 h-7 rounded-lg bg-amber-600 text-white flex items-center justify-center text-xs font-bold">
                                    <User className="w-4 h-4" />
                                  </div>
                                  <span className="font-semibold text-sm text-white">{person.picName}</span>
                                  <span className="text-xs font-mono text-slate-400">({taskCount})</span>

                                  <span className="px-2 py-0.5 text-[10px] font-bold rounded border bg-amber-500/15 text-amber-400 border-amber-500/30">
                                    Outsourcing
                                  </span>

                                  {person.jabatan && (
                                    <span className="text-[11px] text-slate-400 hidden sm:inline">
                                      • {person.jabatan}
                                    </span>
                                  )}

                                  <span className={`px-2 py-0.2 text-[10px] font-bold rounded ${
                                    isActive ? 'bg-emerald-500/10 text-emerald-400' : 'bg-slate-800 text-slate-500'
                                  }`}>
                                    {isActive ? 'Aktif' : 'Kosong'}
                                  </span>
                                </div>

                                <div className="text-slate-400">
                                  {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                                </div>
                              </div>

                              {isExpanded && (
                                <div className="p-3 border-t border-slate-800 bg-slate-950/40">
                                  {taskCount > 0 ? (
                                    <div className="overflow-x-auto">
                                      <table className="w-full text-left text-xs">
                                        <thead>
                                          <tr className="text-slate-400 border-b border-slate-800 font-medium">
                                            <th className="py-2 px-2 w-8 text-center">#</th>
                                            <th className="py-2 px-3 text-amber-400 font-mono">Kode JC</th>
                                            <th className="py-2 px-3">Proyek</th>
                                            <th className="py-2 px-3">Task Name</th>
                                            <th className="py-2 px-3 text-center">Start</th>
                                            <th className="py-2 px-3 text-center">End</th>
                                            <th className="py-2 px-3 text-center font-mono">JO</th>
                                            <th className="py-2 px-2 w-10 text-center">Aksi</th>
                                          </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-800 text-slate-200">
                                          {person.tasks.map((task, tIdx) => (
                                            <tr key={task.id} className="hover:bg-slate-800/40">
                                              <td className="py-2 px-2 text-center text-slate-500 font-mono">{tIdx + 1}</td>
                                              <td className="py-2 px-3 font-mono font-semibold text-amber-300">
                                                {task.kodeJc || <span className="text-rose-400 text-[11px]">Menunggu</span>}
                                              </td>
                                              <td className="py-2 px-3 text-emerald-400 font-medium">{task.project}</td>
                                              <td className="py-2 px-3 text-slate-200">{task.taskName}</td>
                                              <td className="py-2 px-3 text-center font-mono text-cyan-300">{task.startDate}</td>
                                              <td className="py-2 px-3 text-center font-mono text-cyan-300">{task.endDate}</td>
                                              <td className="py-2 px-3 text-center font-mono font-bold text-violet-300">#{task.jo}</td>
                                              <td className="py-2 px-2 text-center">
                                                <button
                                                  onClick={() => handleDeleteTask(task.id, subconSelectedBiro)}
                                                  className="p-1 text-slate-500 hover:text-rose-400 rounded cursor-pointer"
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
                                    <div className="py-4 text-center text-slate-500 text-xs">Belum ada tugas</div>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="py-8 text-center text-slate-500 text-xs bg-slate-900 rounded-xl">
                        Tidak ada anggota outsourcing di biro ini
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAMPILAN MODUL ORGANIK (HALAMAN LEVEL 1, LEVEL 2, FORM & OUTPUT)          */}
        {/* ========================================================================= */}
        {accessMode === 'organik' && (
          <>
            {/* LEVEL 1: 6 DEPARTEMEN */}
            {!selectedDept && !selectedBiroPage && !selectedFormBiro && (
              <div className="space-y-6 animate-fadeIn">
                <div className="flex items-center justify-between gap-4">
                  <div>
                    <h2 className="text-xl font-bold text-white">Departemen Desain</h2>
                    <span className="text-xs text-slate-400">Pilih departemen untuk mengelola Job Card biro</span>
                  </div>
                  <div className="relative w-64">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Cari..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-blue-500"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {filteredDepartments.map((dept) => {
                    const IconComponent = iconMap[dept.icon] || Building2;
                    return (
                      <div
                        key={dept.id}
                        onClick={() => setSelectedDept(dept)}
                        className="bg-slate-800/60 border border-slate-700/60 hover:border-blue-500/50 hover:bg-slate-800 rounded-xl p-5 cursor-pointer transition flex flex-col justify-between"
                      >
                        <div>
                          <div className="flex items-center justify-between mb-3">
                            <div className="p-2.5 bg-blue-500/10 text-blue-400 rounded-lg">
                              <IconComponent className="w-5 h-5" />
                            </div>
                            {dept.code && (
                              <span className="text-[10px] font-mono px-2 py-0.5 bg-slate-700 text-slate-300 rounded border border-slate-600">
                                {dept.code}
                              </span>
                            )}
                          </div>
                          <h3 className="font-semibold text-white text-base">
                            {dept.name}
                          </h3>
                        </div>

                        <div className="mt-4 pt-3 border-t border-slate-700/40 flex items-center justify-between text-xs text-slate-400">
                          <span>{dept.biros.length} Biro</span>
                          <span className="text-blue-400 flex items-center gap-0.5 font-medium">
                            Buka <ChevronRight className="w-3.5 h-3.5" />
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
              <div className="space-y-5 animate-fadeIn">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                  <div>
                    <h2 className="text-lg font-bold text-white">{selectedDept.name}</h2>
                    <span className="text-xs text-slate-400">{selectedDept.biros.length} Biro Terdaftar</span>
                  </div>
                </div>

                <div className="space-y-3">
                  {selectedDept.biros.map((biro) => (
                    <div
                      key={biro.id}
                      className="bg-slate-800/60 border border-slate-700/70 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                    >
                      <div className="flex items-center gap-2 flex-wrap">
                        <h4 className="text-sm font-semibold text-white">{biro.name}</h4>

                        <div className="flex items-center gap-1.5 ml-2">
                          <button
                            onClick={() => {
                              setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name });
                              setFormPageMode('form');
                            }}
                            className="px-2.5 py-1 bg-emerald-600/90 hover:bg-emerald-600 text-white text-[11px] font-bold rounded-md transition flex items-center gap-1 cursor-pointer"
                          >
                            <FileText className="w-3 h-3" /> FORM
                          </button>

                          <button
                            onClick={() => {
                              setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name });
                              setFormPageMode('output');
                            }}
                            className="px-2.5 py-1 bg-blue-600/90 hover:bg-blue-600 text-white text-[11px] font-bold rounded-md transition flex items-center gap-1 cursor-pointer"
                          >
                            <Layers className="w-3 h-3" /> OUTPUT
                          </button>
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5 flex-wrap">
                        {monthList.map((month) => (
                          <button
                            key={month}
                            onClick={() => handleMonthClick(biro.name, month)}
                            disabled={isLoadingExcel}
                            className="px-2.5 py-1 rounded text-xs bg-slate-800 hover:bg-blue-600 text-slate-300 hover:text-white border border-slate-700 transition cursor-pointer"
                          >
                            {month.slice(0, 3)}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* LEVEL FORM vs OUTPUT */}
            {selectedFormBiro && (
              <div className="space-y-4 animate-fadeIn">
                <div className="bg-slate-800/80 border border-slate-700/80 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-base font-bold text-white">{selectedFormBiro.biroName}</h2>
                    <span className="text-xs text-slate-400 font-mono">
                      {totalPersonilCount} Personil ({countOrganik} Organik, {countOutsourcing} Outsourcing) • {totalTasksCount} Tugas
                    </span>
                  </div>

                  <div className="flex items-center gap-2">
                    <div className="bg-slate-900 p-1 rounded-lg border border-slate-700 flex gap-1">
                      <button
                        onClick={() => setFormPageMode('form')}
                        className={`px-3 py-1 text-xs font-semibold rounded-md transition cursor-pointer ${
                          formPageMode === 'form' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
                        }`}
                      >
                        Form
                      </button>
                      <button
                        onClick={() => setFormPageMode('output')}
                        className={`px-3 py-1 text-xs font-semibold rounded-md transition cursor-pointer ${
                          formPageMode === 'output' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                        }`}
                      >
                        Output ({totalTasksCount})
                      </button>
                    </div>

                    <button
                      type="button"
                      onClick={() => {
                        setIsPlannerModalOpen(true);
                        setPinError(false);
                      }}
                      className="px-2.5 py-1 bg-amber-600/90 hover:bg-amber-600 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer"
                    >
                      <Lock className="w-3 h-3" /> Planner
                      {pendingTasksCount > 0 && (
                        <span className="px-1.5 py-0.2 bg-rose-600 text-[10px] font-bold rounded-full">
                          {pendingTasksCount}
                        </span>
                      )}
                    </button>
                  </div>
                </div>

                {/* Form View */}
                {formPageMode === 'form' && (
                  <div className="bg-slate-800/50 border border-slate-700/70 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">Nama Personil</label>
                          {currentBiroMembers.length > 0 ? (
                            <select
                              value={formData.nama}
                              onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                            >
                              <option value="">Pilih Nama Personil...</option>
                              {currentBiroMembers.map((person, idx) => (
                                <option key={idx} value={person.nama}>
                                  {person.nama} ({person.status})
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.nama}
                              onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                              placeholder="Ketik Nama Personil..."
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">Proyek</label>
                          {projectOptions.length > 0 ? (
                            <select
                              value={formData.kodeProyek}
                              onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                            >
                              <option value="">Pilih Proyek...</option>
                              {projectOptions.map((proj, idx) => (
                                <option key={idx} value={proj}>{proj}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.kodeProyek}
                              onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                              placeholder="Kode Proyek..."
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">JO (Angka)</label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formData.jo}
                            onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                            placeholder="Contoh: 300426"
                            required
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white font-mono focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">Task Name</label>
                          {taskOptions.length > 0 ? (
                            <select
                              value={formData.taskName}
                              onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                            >
                              <option value="">Pilih Task...</option>
                              {taskOptions.map((task, idx) => (
                                <option key={idx} value={task}>{task}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={formData.taskName}
                              onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                              placeholder="Uraian Pekerjaan / Task Name..."
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          )}
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">Start Date</label>
                          <input
                            type="date"
                            value={formData.startDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-medium text-slate-300 mb-1">End Date</label>
                          <input
                            type="date"
                            value={formData.endDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-medium text-slate-300 mb-1">PIC</label>
                          <input
                            type="text"
                            value={formData.pic}
                            onChange={(e) => setFormData(prev => ({ ...prev, pic: e.target.value }))}
                            placeholder="Nama PIC..."
                            required
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                      </div>

                      <div className="pt-3 border-t border-slate-700/50 flex justify-end gap-2">
                        <button
                          type="button"
                          onClick={() => setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', pic: '', jo: '' })}
                          className="px-4 py-1.5 bg-slate-700 hover:bg-slate-600 text-slate-200 text-xs rounded-lg transition cursor-pointer"
                        >
                          Reset
                        </button>
                        <button
                          type="submit"
                          className="px-5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer"
                        >
                          <Send className="w-3.5 h-3.5" /> Simpan
                        </button>
                      </div>
                    </form>
                  </div>
                )}

                {/* Output View */}
                {formPageMode === 'output' && (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between gap-3">
                      <div className="relative flex-1 max-w-xs">
                        <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                        <input
                          type="text"
                          placeholder="Cari personil / tugas..."
                          value={outputSearch}
                          onChange={(e) => setOutputSearch(e.target.value)}
                          className="w-full pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-blue-500"
                        />
                      </div>

                      <div className="flex items-center gap-1.5">
                        <button
                          onClick={() => {
                            const all: Record<string, boolean> = {};
                            filteredAccordionData.forEach(g => { all[g.picName] = true; });
                            setExpandedCards(all);
                          }}
                          className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded border border-slate-700 cursor-pointer"
                        >
                          Buka
                        </button>
                        <button
                          onClick={() => setExpandedCards({})}
                          className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded border border-slate-700 cursor-pointer"
                        >
                          Tutup
                        </button>
                        <button
                          onClick={() => handleClearAllBiroData(selectedFormBiro.biroName)}
                          className="p-1.5 text-rose-400 hover:text-rose-300 hover:bg-rose-950/40 rounded border border-rose-800/40 cursor-pointer"
                          title="Reset Data"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>

                    {filteredAccordionData.length > 0 ? (
                      <div className="space-y-2">
                        {filteredAccordionData.map((person, idx) => {
                          const isExpanded = expandedCards[person.picName] ?? false;
                          const taskCount = person.tasks.length;
                          const isActive = taskCount > 0;
                          const isOutsourcing = person.status.toLowerCase().includes('outsourcing');

                          return (
                            <div
                              key={idx}
                              className="bg-slate-800/60 border border-slate-700/70 rounded-xl overflow-hidden"
                            >
                              <div
                                onClick={() => toggleAccordion(person.picName)}
                                className="p-3.5 flex items-center justify-between cursor-pointer select-none hover:bg-slate-800/80 transition"
                              >
                                <div className="flex items-center gap-2.5 flex-wrap">
                                  <div className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-bold ${
                                    isOutsourcing ? 'bg-amber-600 text-white' : isActive ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-400'
                                  }`}>
                                    <User className="w-4 h-4" />
                                  </div>
                                  <span className="font-semibold text-sm text-white">{person.picName}</span>
                                  <span className="text-xs font-mono text-slate-400">({taskCount})</span>

                                  <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${
                                    isOutsourcing 
                                      ? 'bg-amber-500/15 text-amber-400 border-amber-500/30' 
                                      : 'bg-blue-500/15 text-blue-400 border-blue-500/30'
                                  }`}>
                                    {person.status}
                                  </span>

                                  <span className={`px-2 py-0.2 text-[10px] font-bold rounded ${
                                    isActive ? 'bg-emerald-500/10 text-emerald-400' : 'bg-slate-700 text-slate-400'
                                  }`}>
                                    {isActive ? 'Aktif' : 'Kosong'}
                                  </span>
                                </div>

                                <div className="text-slate-400">
                                  {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                                </div>
                              </div>

                              {isExpanded && (
                                <div className="p-3 border-t border-slate-700/60 bg-slate-900/40">
                                  {taskCount > 0 ? (
                                    <div className="overflow-x-auto">
                                      <table className="w-full text-left text-xs">
                                        <thead>
                                          <tr className="text-slate-400 border-b border-slate-700/60 font-medium">
                                            <th className="py-2 px-2 w-8 text-center">#</th>
                                            <th className="py-2 px-3 text-amber-400 font-mono">Kode JC</th>
                                            <th className="py-2 px-3">Proyek</th>
                                            <th className="py-2 px-3">Task Name</th>
                                            <th className="py-2 px-3 text-center">Start</th>
                                            <th className="py-2 px-3 text-center">End</th>
                                            <th className="py-2 px-3 text-center font-mono">JO</th>
                                            <th className="py-2 px-2 w-10 text-center">Aksi</th>
                                          </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-800 text-slate-200">
                                          {person.tasks.map((task, tIdx) => (
                                            <tr key={task.id} className="hover:bg-slate-800/40">
                                              <td className="py-2 px-2 text-center text-slate-500 font-mono">{tIdx + 1}</td>
                                              <td className="py-2 px-3 font-mono font-semibold text-amber-300">
                                                {task.kodeJc || <span className="text-rose-400 text-[11px]">Menunggu</span>}
                                              </td>
                                              <td className="py-2 px-3 text-emerald-400 font-medium">{task.project}</td>
                                              <td className="py-2 px-3 text-slate-200">{task.taskName}</td>
                                              <td className="py-2 px-3 text-center font-mono text-cyan-300">{task.startDate}</td>
                                              <td className="py-2 px-3 text-center font-mono text-cyan-300">{task.endDate}</td>
                                              <td className="py-2 px-3 text-center font-mono font-bold text-violet-300">#{task.jo}</td>
                                              <td className="py-2 px-2 text-center">
                                                <button
                                                  onClick={() => handleDeleteTask(task.id, selectedFormBiro.biroName)}
                                                  className="p-1 text-slate-500 hover:text-rose-400 rounded cursor-pointer"
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
                                    <div className="py-4 text-center text-slate-500 text-xs">Kosong</div>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="py-8 text-center text-slate-500 text-xs bg-slate-800/30 rounded-xl">
                        Data tidak ditemukan
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* LEVEL 3: TABEL KPI */}
            {selectedBiroPage && (
              <div className="space-y-4 animate-fadeIn">
                <div className="bg-slate-800/80 border border-slate-700 rounded-xl p-4 flex items-center justify-between">
                  <div>
                    <h2 className="text-base font-bold text-white">{selectedBiroPage.biroName}</h2>
                    <span className="text-xs text-slate-400 font-mono">Bulan {selectedBiroPage.month} • {filteredTableData.length} Pegawai</span>
                  </div>
                  <button
                    onClick={() => setSelectedBiroPage(null)}
                    className="px-3 py-1 text-xs bg-slate-700 hover:bg-slate-600 text-white rounded-lg cursor-pointer"
                  >
                    Kembali
                  </button>
                </div>

                <div className="flex items-center justify-between gap-4">
                  <div className="relative w-64">
                    <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Cari NIP / Nama..."
                      value={tableSearch}
                      onChange={(e) => setTableSearch(e.target.value)}
                      className="w-full pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none"
                    />
                  </div>
                </div>

                <div className="bg-slate-800/60 border border-slate-700 rounded-xl overflow-hidden">
                  {filteredTableData.length > 0 ? (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs">
                        <thead>
                          <tr className="bg-slate-800 text-slate-300 font-semibold border-b border-slate-700">
                            <th className="py-2.5 px-3 text-center">#</th>
                            <th className="py-2.5 px-3">NIP</th>
                            <th className="py-2.5 px-3">Nama</th>
                            <th className="py-2.5 px-3 text-right">Effective</th>
                            <th className="py-2.5 px-3 text-right">Overtime</th>
                            <th className="py-2.5 px-3 text-right">Idle</th>
                            <th className="py-2.5 px-3 text-right">TS Reguler</th>
                            <th className="py-2.5 px-3 text-right">TS Overtime</th>
                            <th className="py-2.5 px-3 text-center">Terlambat</th>
                            <th className="py-2.5 px-3 text-center">Sakit</th>
                            <th className="py-2.5 px-3 text-center">IPM</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-700/50 text-slate-200">
                          {filteredTableData.map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-700/30">
                              <td className="py-2.5 px-3 text-center text-slate-500 font-mono">{idx + 1}</td>
                              <td className="py-2.5 px-3 font-mono text-blue-400">{row.nip}</td>
                              <td className="py-2.5 px-3 font-medium text-white">{row.nama}</td>
                              <td className="py-2.5 px-3 text-right font-mono text-cyan-400">{row.effectiveHour}</td>
                              <td className="py-2.5 px-3 text-right font-mono text-amber-400">{row.overtimeHour}</td>
                              <td className="py-2.5 px-3 text-right font-mono">{row.idleHour}</td>
                              <td className="py-2.5 px-3 text-right font-mono text-indigo-300">{row.timesheetReguler}%</td>
                              <td className="py-2.5 px-3 text-right font-mono text-violet-300">{row.timesheetOvertime}%</td>
                              <td className="py-2.5 px-3 text-center font-mono text-rose-400">{row.terlambat}</td>
                              <td className="py-2.5 px-3 text-center font-mono text-amber-400">{row.sakit}</td>
                              <td className="py-2.5 px-3 text-center font-mono text-purple-400">{row.ipm}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="py-8 text-center text-slate-500 text-xs">Data tidak ditemukan</div>
                  )}
                </div>
              </div>
            )}
          </>
        )}

        {/* MODAL PLANNER */}
        {isPlannerModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-800 border border-slate-700 rounded-xl w-full max-w-2xl shadow-xl overflow-hidden max-h-[85vh] flex flex-col">
              <div className="p-3.5 bg-slate-900 border-b border-slate-700 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel
                </span>
                <button
                  onClick={() => setIsPlannerModalOpen(false)}
                  className="p-1 text-slate-400 hover:text-white cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-4 overflow-y-auto flex-1">
                {!isPlannerUnlocked ? (
                  <form onSubmit={handleVerifyPin} className="max-w-xs mx-auto space-y-3 py-4">
                    <div className="text-center">
                      <KeyRound className="w-8 h-8 mx-auto text-amber-400 mb-1.5" />
                      <span className="text-xs text-slate-300">PIN Planner</span>
                    </div>

                    <input
                      type="password"
                      value={pinInput}
                      onChange={(e) => setPinInput(e.target.value)}
                      placeholder="PIN (2026)..."
                      autoFocus
                      required
                      className={`w-full px-3 py-2 bg-slate-900 border rounded-lg text-center text-sm font-mono text-white focus:outline-none ${
                        pinError ? 'border-rose-500' : 'border-slate-700'
                      }`}
                    />

                    <button
                      type="submit"
                      className="w-full py-2 bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold rounded-lg transition cursor-pointer"
                    >
                      Buka
                    </button>
                  </form>
                ) : (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between text-xs text-slate-400 pb-2 border-b border-slate-700">
                      <span>Daftar Pengajuan ({currentBiroSubmittedTasks.length})</span>
                      <button
                        onClick={() => setIsPlannerUnlocked(false)}
                        className="text-[11px] underline hover:text-white cursor-pointer"
                      >
                        Kunci
                      </button>
                    </div>

                    {currentBiroSubmittedTasks.length > 0 ? (
                      <div className="space-y-2">
                        {currentBiroSubmittedTasks.map((task, idx) => {
                          const currentVal = editingTaskKode[task.id] ?? task.kodeJc ?? '';

                          return (
                            <div 
                              key={task.id}
                              className="p-3 bg-slate-900/60 border border-slate-700 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-2.5"
                            >
                              <div className="space-y-0.5 text-xs">
                                <div className="flex items-center gap-2">
                                  <span className="font-mono text-slate-500">#{idx + 1}</span>
                                  <span className="font-semibold text-white">{task.pic}</span>
                                  <span className="text-blue-400">{task.project}</span>
                                  <span className="text-purple-400 font-mono">#{task.jo}</span>
                                </div>
                                <div className="text-slate-300 text-[11px]">{task.taskName}</div>
                              </div>

                              <div className="flex items-center gap-1.5 self-end sm:self-center">
                                <input
                                  type="text"
                                  value={currentVal}
                                  onChange={(e) => setEditingTaskKode(prev => ({
                                    ...prev,
                                    [task.id]: e.target.value.toUpperCase()
                                  }))}
                                  placeholder="Kode JC..."
                                  className="w-32 px-2.5 py-1 bg-slate-900 border border-slate-600 rounded text-xs font-mono text-white uppercase focus:outline-none focus:ring-1 focus:ring-amber-500"
                                />
                                <button
                                  onClick={() => handleSaveKodeJcForTask(task.id, selectedFormBiro ? selectedFormBiro.biroName : subconSelectedBiro || '')}
                                  className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white rounded text-xs font-semibold flex items-center gap-1 cursor-pointer"
                                >
                                  <Check className="w-3 h-3" /> Simpan
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="py-6 text-center text-slate-500 text-xs">Tidak ada data</div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

      </main>
    </div>
  );
}