import { useState, useEffect, useCallback, useMemo, ChangeEvent } from 'react';
import { departmentsData, monthList, Department } from './data';
import * as XLSX from 'xlsx';
import { supabase } from './lib/supabase';

// Import File CSV Absensi
import csvJan from './absensi_januari.csv?raw';
import csvFeb from './absensi_februari.csv?raw';
import csvMar from './absensi_maret.csv?raw';
import csvApr from './absensi_april.csv?raw';
import csvMei from './absensi_mei.csv?raw';
import csvJun from './absensi_juni.csv?raw';

import { 
  Building2, 
  Briefcase, 
  HardHat, 
  ArrowLeft, 
  Search, 
  ChevronRight, 
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
  Upload,
  Printer, 
  Calendar, 
  FileCheck,
  Users,
  Wrench,
  CircleDollarSign,
  Truck,
  ShieldCheck,
  Laptop,
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

  // Subkon Navigation State
  const [subconSelectedDept, setSubconSelectedDept] = useState<Department | null>(null);
  const [subconSelectedBiro, setSubconSelectedBiro] = useState<string | null>(null);
  const [subconPageMode, setSubconPageMode] = useState<'members' | 'form' | 'output' | 'release'>('members');
  const [subconSearch, setSubconSearch] = useState('');

  // Organik Navigation State
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
        .order('created_at', { ascending: true });

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
  const [plannerScope, setPlannerScope] = useState<'current' | 'all'>('current');

  const PLANNER_PIN = '2026';
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  // 3 File Excel Utama
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
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
        alert(`File ${file.name} berhasil dibaca.`);
      } catch {
        alert('Gagal membaca file Excel.');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Parser Master Anggota dari Excel IM4
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

  const dynamicOutsourcingList = useMemo(() => {
    return allParsedFromExcel.filter(p => p.status.toLowerCase().includes('outsourcing'));
  }, [allParsedFromExcel]);

  const getSubconMembersForBiro = useCallback((biroName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.biro, biroName));
  }, [dynamicOutsourcingList]);

  const getSubconCountForDept = useCallback((deptName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.dept, deptName)).length;
  }, [dynamicOutsourcingList]);

  // Deteksi Biro Aktif
  const currentActiveBiroName = useMemo(() => {
    if (accessMode === 'subkon') return subconSelectedBiro || '';
    return selectedFormBiro?.biroName || '';
  }, [accessMode, subconSelectedBiro, selectedFormBiro]);

  const currentActiveBiroKey = useMemo(() => {
    return cleanText(currentActiveBiroName);
  }, [currentActiveBiroName]);

  const currentActiveBiroTasks = useMemo(() => {
    return manualTasks[currentActiveBiroKey] || [];
  }, [manualTasks, currentActiveBiroKey]);

  const allSubmittedTasksList = useMemo(() => {
    const list: TaskItem[] = [];
    Object.values(manualTasks).forEach(arr => list.push(...arr));
    return list;
  }, [manualTasks]);

  const plannerTasksToShow = useMemo(() => {
    if (plannerScope === 'all') return allSubmittedTasksList;
    return currentActiveBiroTasks;
  }, [plannerScope, allSubmittedTasksList, currentActiveBiroTasks]);

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
        if (p && p.toLowerCase() !== 'nan' && !p.toLowerCase().includes('kode proyek')) {
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
        if (t && t.toLowerCase() !== 'nan' && !t.toLowerCase().includes('desc pekerjaan')) {
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
        personMap.set(key, { status: matched?.status || 'Organik', jabatan: matched?.jabatan || '', tasks: [] });
      }
      personMap.get(key)!.tasks.push(t);
    });

    const result: PersonilCardGroup[] = [];
    personMap.forEach((val, picName) => {
      result.push({ picName, status: val.status, jabatan: val.jabatan, tasks: val.tasks });
    });

    return result.sort((a, b) => a.picName.localeCompare(b.picName));
  };

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    const activeBiro = currentActiveBiroName;
    if (!activeBiro) return;

    try {
      let validBiroId: string | null = null;
      const { data: biroList } = await supabase.from('biros').select('id, name');
      
      if (biroList && biroList.length > 0) {
        const found = biroList.find(b => isBiroMatch(b.name, activeBiro));
        validBiroId = found ? found.id : biroList[0].id;
      }

      const { data: insertedRow, error } = await supabase
        .from('job_cards')
        .insert({
          biro_id: validBiroId,
          biro_name: activeBiro,
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
        alert('Gagal simpan: ' + error.message);
        return;
      }

      const biroKey = cleanText(activeBiro);
      const newTask: TaskItem = {
        id: insertedRow.id,
        biroName: activeBiro,
        project: formData.kodeProyek,
        taskName: formData.taskName,
        startDate: formData.startDate,
        endDate: formData.endDate,
        pic: formData.nama,
        jo: formData.jo,
        kode_jc: '',
      };

      const currentList = manualTasks[biroKey] || [];
      setManualTasks({ ...manualTasks, [biroKey]: [...currentList, newTask] });

      setExpandedCards(prev => ({ ...prev, [formData.nama]: true }));
      setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', pic: '', jo: '' });

      alert('Tugas tersimpan! Masukkan Kode JC di menu Planner untuk menerbitkan Work Order.');
      loadAllJobCards();
    } catch {
      alert('Koneksi database bermasalah.');
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

  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;

    const { error } = await supabase
      .from('job_cards')
      .update({ kode_jc: inputVal, status: 'approved' })
      .eq('id', taskId);

    if (error) {
      alert('Gagal simpan kode: ' + error.message);
      return;
    }

    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => {
        updated[k] = updated[k].map(t => t.id === taskId ? { ...t, kodeJc: inputVal } : t);
      });
      return updated;
    });

    setEditingTaskKode(prev => {
      const next = { ...prev };
      delete next[taskId];
      return next;
    });

    loadAllJobCards();
  };

  const handleDeleteTask = async (taskId: string) => {
    if (window.confirm('Hapus tugas ini?')) {
      const { error } = await supabase.from('job_cards').delete().eq('id', taskId);
      if (error) return;

      setManualTasks(prev => {
        const updated = { ...prev };
        Object.keys(updated).forEach(k => {
          updated[k] = updated[k].filter(t => t.id !== taskId);
        });
        return updated;
      });
      loadAllJobCards();
    }
  };

  const toggleAccordion = (picName: string) => {
    setExpandedCards(prev => ({ ...prev, [picName]: !prev[picName] }));
  };

  const handleMonthClick = (biroName: string, month: string) => {
    if (!workbook) return;

    const monthPrefix = month.toLowerCase().slice(0, 3);
    const targetSheetName = workbook.SheetNames.find(sheet => {
      const sLower = sheet.toLowerCase().trim();
      return sLower.includes(month.toLowerCase()) || sLower.includes(monthPrefix);
    });

    if (!targetSheetName) return;

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

  const filteredDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const filteredSubconDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(subconSearch.toLowerCase())
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
    return g.picName.toLowerCase().includes(s) || g.tasks.some(t => t.project.toLowerCase().includes(s) || t.taskName.toLowerCase().includes(s) || (t.kodeJc || '').toLowerCase().includes(s));
  });

  const activeSubconMembers = subconSelectedBiro ? getSubconMembersForBiro(subconSelectedBiro) : [];
  const activeSubconPendingCount = currentActiveBiroTasks.filter(t => !t.kodeJc).length;

  // Auto-Numbering Work Order (AA1, AA2...) per Biro
  const approvedSubconPackages = useMemo(() => {
    return currentActiveBiroTasks
      .filter(t => t.kodeJc && t.kodeJc.trim() !== '')
      .map((task, idx) => ({
        ...task,
        packageTitle: `AA${idx + 1}`
      }));
  }, [currentActiveBiroTasks]);

  const subconAccordionData = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const members = getSubconMembersForBiro(subconSelectedBiro);
    const pMap = new Map<string, { status: string; jabatan: string; tasks: TaskItem[] }>();

    members.forEach(m => {
      pMap.set(m.nama, { status: 'Outsourcing', jabatan: m.jabatan, tasks: [] });
    });

    currentActiveBiroTasks.forEach(t => {
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
  }, [subconSelectedBiro, getSubconMembersForBiro, currentActiveBiroTasks]);

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div 
            className="flex items-center gap-2.5 cursor-pointer" 
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
            <div className="flex items-center gap-2">
              <span className="font-bold text-sm text-white">DIVISI DESAIN</span>
              {accessMode !== 'landing' && (
                <span className="text-xs px-2 py-0.5 rounded font-mono bg-slate-800 text-slate-300 border border-slate-700">
                  {accessMode === 'organik' ? 'Organik' : 'Subkontraktor'}
                </span>
              )}
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
                className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer"
              >
                Ganti Portal
              </button>
            )}

            {/* Tombol Kembali Dinamis */}
            {accessMode === 'subkon' && subconSelectedBiro ? (
              <button onClick={() => setSubconSelectedBiro(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'subkon' && subconSelectedDept ? (
              <button onClick={() => setSubconSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                <ArrowLeft className="w-3.5 h-3.5" /> Dept
              </button>
            ) : accessMode === 'organik' && selectedFormBiro ? (
              <button onClick={() => setSelectedFormBiro(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'organik' && selectedBiroPage ? (
              <button onClick={() => setSelectedBiroPage(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                <ArrowLeft className="w-3.5 h-3.5" /> Biro
              </button>
            ) : accessMode === 'organik' && selectedDept ? (
              <button onClick={() => setSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                <ArrowLeft className="w-3.5 h-3.5" /> Dept
              </button>
            ) : null}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-6xl mx-auto px-4 py-8">

        {/* ================= 1. PORTAL DEPAN (LANDING) ================= */}
        {accessMode === 'landing' && (
          <div className="max-w-2xl mx-auto text-center space-y-8 pt-8">
            <div className="space-y-2">
              <h1 className="text-2xl sm:text-3xl font-bold text-white">Sistem Penugasan Job Card</h1>
              <p className="text-slate-400 text-sm">Pilih portal akses kerja Anda</p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div
                onClick={() => setAccessMode('organik')}
                className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-2xl p-6 cursor-pointer text-left transition hover:bg-slate-900/80 group"
              >
                <div className="p-3 bg-blue-500/10 text-blue-400 rounded-xl w-fit mb-4 group-hover:bg-blue-600 group-hover:text-white transition">
                  <Briefcase className="w-6 h-6" />
                </div>
                <h3 className="text-lg font-bold text-white">Pegawai Organik</h3>
                <span className="text-xs text-slate-400">Pegawai PKWTT & PKWT Divisi Desain</span>
              </div>

              <div
                onClick={() => setAccessMode('subkon')}
                className="bg-slate-900 border border-slate-800 hover:border-amber-500 rounded-2xl p-6 cursor-pointer text-left transition hover:bg-slate-900/80 group"
              >
                <div className="p-3 bg-amber-500/10 text-amber-400 rounded-xl w-fit mb-4 group-hover:bg-amber-600 group-hover:text-white transition">
                  <HardHat className="w-6 h-6" />
                </div>
                <h3 className="text-lg font-bold text-white">Mitra / Subkon</h3>
                <span className="text-xs text-slate-400">Personil Outsourcing</span>
              </div>
            </div>
          </div>
        )}

        {/* ================= 2. PORTAL SUBKONTRAKTOR ================= */}
        {accessMode === 'subkon' && (
          <div className="space-y-6">
            {/* Step 1: Pilih Departemen */}
            {!subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                  <div>
                    <h2 className="text-lg font-bold text-white">Pilih Departemen</h2>
                    <span className="text-xs text-slate-400">Total {dynamicOutsourcingList.length} Personel Outsourcing</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Upload className="w-3.5 h-3.5 text-amber-400" />
                      <span>{dynamicOutsourcingList.length > 0 ? 'Update Excel' : 'Upload IM4'}</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadExcel} className="hidden" />
                    </label>

                    <input
                      type="text"
                      placeholder="Cari..."
                      value={subconSearch}
                      onChange={(e) => setSubconSearch(e.target.value)}
                      className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                  {filteredSubconDepartments.map((dept) => {
                    const count = getSubconCountForDept(dept.name);
                    return (
                      <div
                        key={dept.id}
                        onClick={() => setSubconSelectedDept(dept)}
                        className="bg-slate-900 border border-slate-800 hover:border-amber-500/60 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
                      >
                        <div>
                          <h4 className="font-semibold text-white text-sm">{dept.name}</h4>
                          <span className="text-xs text-slate-400">{dept.biros.length} Biro</span>
                        </div>
                        <span className="px-2 py-0.5 bg-amber-500/10 text-amber-400 border border-amber-500/20 rounded font-bold text-xs">
                          {count}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Step 2: Pilih Biro */}
            {subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-4">
                <div className="border-b border-slate-800 pb-2">
                  <h2 className="text-base font-bold text-white">{subconSelectedDept.name}</h2>
                  <span className="text-xs text-slate-400">Pilih biro penugasan</span>
                </div>

                <div className="space-y-2">
                  {subconSelectedDept.biros.map((biro) => {
                    const countInBiro = getSubconMembersForBiro(biro.name).length;
                    const releaseCount = (manualTasks[cleanText(biro.name)] || []).filter(t => t.kodeJc).length;

                    return (
                      <div
                        key={biro.id}
                        className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                      >
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="font-semibold text-white text-sm">{biro.name}</span>
                          <span className="text-xs text-amber-400 font-mono">({countInBiro} Org)</span>
                          {releaseCount > 0 && (
                            <span className="px-1.5 py-0.5 bg-purple-500/20 text-purple-300 rounded text-[10px] font-bold">
                              {releaseCount} Work Order
                            </span>
                          )}
                        </div>

                        <div className="flex items-center gap-1.5 self-end sm:self-auto">
                          <button
                            onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('members'); }}
                            className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs rounded border border-slate-700 cursor-pointer"
                          >
                            Anggota
                          </button>
                          <button
                            onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('form'); }}
                            className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs rounded cursor-pointer"
                          >
                            Form
                          </button>
                          <button
                            onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('output'); }}
                            className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs rounded cursor-pointer"
                          >
                            Output
                          </button>
                          <button
                            onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('release'); }}
                            className="px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white font-semibold text-xs rounded cursor-pointer"
                          >
                            Work Order
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Step 3: Halaman Detail Biro Subkon */}
            {subconSelectedBiro && (
              <div className="space-y-4">
                {/* Header & Switcher Tab */}
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-white text-base">{subconSelectedBiro}</h3>
                    <span className="text-xs text-slate-400 font-mono">{activeSubconMembers.length} Personel Outsourcing</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <div className="bg-slate-950 p-1 rounded-lg border border-slate-800 flex gap-1 text-xs font-semibold">
                      <button
                        onClick={() => setSubconPageMode('members')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${subconPageMode === 'members' ? 'bg-amber-600 text-white' : 'text-slate-400 hover:text-white'}`}
                      >
                        Anggota
                      </button>
                      <button
                        onClick={() => setSubconPageMode('form')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${subconPageMode === 'form' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'}`}
                      >
                        Form
                      </button>
                      <button
                        onClick={() => setSubconPageMode('output')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${subconPageMode === 'output' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'}`}
                      >
                        Output ({currentActiveBiroTasks.length})
                      </button>
                      <button
                        onClick={() => setSubconPageMode('release')}
                        className={`px-3 py-1 rounded transition cursor-pointer flex items-center gap-1 ${subconPageMode === 'release' ? 'bg-purple-600 text-white' : 'text-purple-300 hover:text-white'}`}
                      >
                        <FileCheck className="w-3 h-3 text-amber-300" /> Work Order ({approvedSubconPackages.length})
                      </button>
                    </div>

                    <button
                      onClick={() => setIsPlannerModalOpen(true)}
                      className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
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

                {/* Tab 1: Anggota Outsourcing */}
                {subconPageMode === 'members' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-800 text-slate-300 border-b border-slate-700">
                        <tr>
                          <th className="py-2.5 px-3 text-center w-10">#</th>
                          <th className="py-2.5 px-3">Nama</th>
                          <th className="py-2.5 px-3 font-mono">NIP</th>
                          <th className="py-2.5 px-3">Jabatan</th>
                          <th className="py-2.5 px-3 text-center">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800 text-slate-200">
                        {activeSubconMembers.map((person, idx) => (
                          <tr key={idx} className="hover:bg-slate-800/40">
                            <td className="py-2 px-3 text-center text-slate-500 font-mono">{idx + 1}</td>
                            <td className="py-2 px-3 font-semibold text-white">{person.nama}</td>
                            <td className="py-2 px-3 font-mono text-slate-400">{person.nip}</td>
                            <td className="py-2 px-3 text-cyan-300">{person.jabatan}</td>
                            <td className="py-2 px-3 text-center">
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/10 text-amber-400 border border-amber-500/20">
                                Outsourcing
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                {/* Tab 2: Formulir Penugasan */}
                {subconPageMode === 'form' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Nama Drafter / PIC</label>
                          <select
                            value={formData.nama}
                            onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          >
                            <option value="">Pilih Personel...</option>
                            {activeSubconMembers.map((p, i) => (
                              <option key={i} value={p.nama}>{p.nama} ({p.jabatan})</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Kode Proyek</label>
                          <select
                            value={formData.kodeProyek}
                            onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          >
                            <option value="">Pilih Proyek...</option>
                            {projectOptions.map((p, i) => (
                              <option key={i} value={p}>{p}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Nomor JO</label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formData.jo}
                            onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                            placeholder="Contoh: 300426"
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono focus:outline-none"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Uraian Tugas / Task Name</label>
                          <select
                            value={formData.taskName}
                            onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          >
                            <option value="">Pilih Task...</option>
                            {taskOptions.map((t, i) => (
                              <option key={i} value={t}>{t}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Tanggal Mulai</label>
                          <input
                            type="date"
                            value={formData.startDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          />
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Tanggal Selesai</label>
                          <input
                            type="date"
                            value={formData.endDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">PIC Subkon</label>
                          <input
                            type="text"
                            value={formData.pic}
                            onChange={(e) => setFormData(prev => ({ ...prev, pic: e.target.value }))}
                            placeholder="Nama PIC..."
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          />
                        </div>
                      </div>

                      <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                        <button
                          type="submit"
                          className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-white font-bold text-xs rounded-lg transition cursor-pointer"
                        >
                          Simpan Tugas
                        </button>
                      </div>
                    </form>
                  </div>
                )}

                {/* Tab 3: Output Tugas */}
                {subconPageMode === 'output' && (
                  <div className="space-y-2">
                    {subconAccordionData.map((person, idx) => {
                      const isExpanded = expandedCards[person.picName] ?? false;
                      const taskCount = person.tasks.length;

                      return (
                        <div key={idx} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                          <div
                            onClick={() => toggleAccordion(person.picName)}
                            className="p-3 flex items-center justify-between cursor-pointer hover:bg-slate-800/60"
                          >
                            <div className="flex items-center gap-2">
                              <span className="font-semibold text-white text-sm">{person.picName}</span>
                              <span className="text-xs text-slate-400 font-mono">({taskCount} tugas)</span>
                            </div>
                            <div className="text-slate-400">
                              {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                            </div>
                          </div>

                          {isExpanded && (
                            <div className="p-3 border-t border-slate-800 bg-slate-950/40">
                              {taskCount > 0 ? (
                                <table className="w-full text-left text-xs">
                                  <thead>
                                    <tr className="text-slate-400 border-b border-slate-800">
                                      <th className="py-2 px-2">#</th>
                                      <th className="py-2 px-2 font-mono text-amber-400">Kode JC</th>
                                      <th className="py-2 px-2">Proyek</th>
                                      <th className="py-2 px-2">Task</th>
                                      <th className="py-2 px-2 font-mono">JO</th>
                                      <th className="py-2 px-2 text-center">Aksi</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-800 text-slate-300">
                                    {person.tasks.map((task, tIdx) => (
                                      <tr key={task.id}>
                                        <td className="py-2 px-2 text-slate-500">{tIdx + 1}</td>
                                        <td className="py-2 px-2 font-mono font-bold text-amber-300">
                                          {task.kodeJc || <span className="text-rose-400 font-normal">Menunggu Planner</span>}
                                        </td>
                                        <td className="py-2 px-2 text-emerald-400">{task.project}</td>
                                        <td className="py-2 px-2">{task.taskName}</td>
                                        <td className="py-2 px-2 font-mono">#{task.jo}</td>
                                        <td className="py-2 px-2 text-center">
                                          <button onClick={() => handleDeleteTask(task.id)} className="text-slate-500 hover:text-rose-400 p-1">
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              ) : (
                                <div className="text-xs text-slate-500 py-2">Belum ada tugas</div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Tab 4: Work Order (AA1, AA2...) */}
                {subconPageMode === 'release' && (
                  <div className="space-y-3">
                    {approvedSubconPackages.length > 0 ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {approvedSubconPackages.map((pkg) => (
                          <div key={pkg.id} className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-2">
                            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                              <span className="px-2 py-0.5 bg-purple-600 rounded text-white font-mono font-bold text-sm">
                                {pkg.packageTitle}
                              </span>
                              <span className="font-mono text-xs text-amber-400 font-bold bg-amber-500/10 px-2 py-0.5 rounded border border-amber-500/20">
                                {pkg.kodeJc}
                              </span>
                            </div>
                            <div className="text-xs space-y-1">
                              <div className="text-white font-semibold">{pkg.pic}</div>
                              <div className="text-emerald-400 font-mono">{pkg.project} • #{pkg.jo}</div>
                              <div className="text-slate-300 text-[11px] bg-slate-950 p-2 rounded border border-slate-800">{pkg.taskName}</div>
                            </div>
                            <div className="text-[10px] text-slate-500 pt-2 border-t border-slate-800 flex justify-between font-mono">
                              <span>{pkg.startDate} s/d {pkg.endDate}</span>
                              <button onClick={() => alert(`Cetak work order ${pkg.packageTitle}`)} className="text-amber-400 hover:underline">
                                Cetak
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="bg-slate-900 border border-slate-800 rounded-xl p-8 text-center text-xs text-slate-500">
                        Belum ada Work Order di biro ini.
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ================= 3. PORTAL ORGANIK ================= */}
        {accessMode === 'organik' && (
          <div className="space-y-6">
            {/* Step 1: Departemen Organik */}
            {!selectedDept && !selectedBiroPage && !selectedFormBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                  <div>
                    <h2 className="text-lg font-bold text-white">Departemen Desain</h2>
                    <span className="text-xs text-slate-400">Akses Pegawai Organik</span>
                  </div>
                  <input
                    type="text"
                    placeholder="Cari..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44"
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                  {filteredDepartments.map((dept) => (
                    <div
                      key={dept.id}
                      onClick={() => setSelectedDept(dept)}
                      className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
                    >
                      <div>
                        <h4 className="font-semibold text-white text-sm">{dept.name}</h4>
                        <span className="text-xs text-slate-400">{dept.biros.length} Biro</span>
                      </div>
                      <ChevronRight className="w-4 h-4 text-slate-500" />
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Step 2: Biro Organik */}
            {selectedDept && !selectedBiroPage && !selectedFormBiro && (
              <div className="space-y-4">
                <div className="border-b border-slate-800 pb-2">
                  <h2 className="text-base font-bold text-white">{selectedDept.name}</h2>
                  <span className="text-xs text-slate-400">Daftar Biro & Rekap Bulanan</span>
                </div>

                <div className="space-y-2">
                  {selectedDept.biros.map((biro) => (
                    <div
                      key={biro.id}
                      className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                    >
                      <span className="font-semibold text-white text-sm">{biro.name}</span>

                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('form'); }}
                          className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs rounded cursor-pointer"
                        >
                          Form
                        </button>
                        <button
                          onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('output'); }}
                          className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs rounded cursor-pointer"
                        >
                          Output
                        </button>

                        <div className="h-4 w-px bg-slate-800 mx-1 hidden sm:block" />

                        {monthList.map((month) => (
                          <button
                            key={month}
                            onClick={() => handleMonthClick(biro.name, month)}
                            className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] rounded border border-slate-700 cursor-pointer"
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

            {/* Step 3: Form/Output Organik */}
            {selectedFormBiro && (
              <div className="space-y-4">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex items-center justify-between">
                  <h3 className="font-bold text-white text-sm">{selectedFormBiro.biroName}</h3>
                  <div className="flex gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 text-xs">
                    <button
                      onClick={() => setFormPageMode('form')}
                      className={`px-3 py-1 rounded ${formPageMode === 'form' ? 'bg-emerald-600 text-white' : 'text-slate-400'}`}
                    >
                      Form
                    </button>
                    <button
                      onClick={() => setFormPageMode('output')}
                      className={`px-3 py-1 rounded ${formPageMode === 'output' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}
                    >
                      Output ({currentActiveBiroTasks.length})
                    </button>
                  </div>
                </div>

                {formPageMode === 'form' ? (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Nama Personel</label>
                          <select
                            value={formData.nama}
                            onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                          >
                            <option value="">Pilih Personel...</option>
                            {currentBiroMembers.map((p, i) => (
                              <option key={i} value={p.nama}>{p.nama} ({p.status})</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Proyek</label>
                          <select
                            value={formData.kodeProyek}
                            onChange={(e) => setFormData(prev => ({ ...prev, kodeProyek: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                          >
                            <option value="">Pilih Proyek...</option>
                            {projectOptions.map((p, i) => (
                              <option key={i} value={p}>{p}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Nomor JO</label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formData.jo}
                            onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Task Name</label>
                          <select
                            value={formData.taskName}
                            onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                          >
                            <option value="">Pilih Task...</option>
                            {taskOptions.map((t, i) => (
                              <option key={i} value={t}>{t}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Mulai</label>
                          <input type="date" value={formData.startDate} onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" />
                        </div>
                        <div>
                          <label className="block text-slate-400 mb-1">Selesai</label>
                          <input type="date" value={formData.endDate} onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">PIC</label>
                          <input type="text" value={formData.pic} onChange={(e) => setFormData(prev => ({ ...prev, pic: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" />
                        </div>
                      </div>

                      <div className="flex justify-end pt-2">
                        <button type="submit" className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg cursor-pointer">
                          Simpan
                        </button>
                      </div>
                    </form>
                  </div>
                ) : (
                  <div className="space-y-2">
                    {filteredAccordionData.map((person, idx) => (
                      <div key={idx} className="bg-slate-900 border border-slate-800 rounded-xl p-3">
                        <div className="font-semibold text-white text-sm mb-1">{person.picName} ({person.tasks.length} tugas)</div>
                        {person.tasks.length > 0 && (
                          <div className="text-xs text-slate-400 space-y-1 pt-1">
                            {person.tasks.map((t, i) => (
                              <div key={i} className="flex justify-between border-t border-slate-800/80 pt-1">
                                <span>{t.taskName} ({t.project})</span>
                                <span className="font-mono text-amber-400">{t.kodeJc || 'Pending'}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* Step 4: Rekap KPI Bulanan */}
            {selectedBiroPage && (
              <div className="space-y-3">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex items-center justify-between">
                  <span className="font-bold text-white text-sm">{selectedBiroPage.biroName} — {selectedBiroPage.month}</span>
                  <button onClick={() => setSelectedBiroPage(null)} className="px-3 py-1 bg-slate-800 text-xs rounded border border-slate-700">Tutup</button>
                </div>

                <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-800 text-slate-300 border-b border-slate-700">
                      <tr>
                        <th className="py-2.5 px-3">Nama</th>
                        <th className="py-2.5 px-3 font-mono">NIP</th>
                        <th className="py-2.5 px-3 text-right">Effective</th>
                        <th className="py-2.5 px-3 text-right">Overtime</th>
                        <th className="py-2.5 px-3 text-right">Idle</th>
                        <th className="py-2.5 px-3 text-center">Terlambat</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800 text-slate-300">
                      {filteredTableData.map((row, idx) => (
                        <tr key={idx}>
                          <td className="py-2 px-3 text-white font-medium">{row.nama}</td>
                          <td className="py-2 px-3 font-mono text-slate-400">{row.nip}</td>
                          <td className="py-2 px-3 text-right font-mono text-cyan-400">{row.effectiveHour}</td>
                          <td className="py-2 px-3 text-right font-mono text-amber-400">{row.overtimeHour}</td>
                          <td className="py-2 px-3 text-right font-mono">{row.idleHour}</td>
                          <td className="py-2 px-3 text-center font-mono text-rose-400">{row.terlambat}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= MODAL PLANNER ================= */}
        {isPlannerModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
              <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel
                </span>
                <button onClick={() => setIsPlannerModalOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-4">
                {!isPlannerUnlocked ? (
                  <form onSubmit={handleVerifyPin} className="max-w-xs mx-auto space-y-3 py-6 text-center">
                    <KeyRound className="w-8 h-8 text-amber-400 mx-auto" />
                    <input
                      type="password"
                      value={pinInput}
                      onChange={(e) => setPinInput(e.target.value)}
                      placeholder="PIN (2026)..."
                      autoFocus
                      required
                      className={`w-full px-3 py-2 bg-slate-950 border rounded-lg text-center text-sm font-mono text-white focus:outline-none ${pinError ? 'border-rose-500' : 'border-slate-800'}`}
                    />
                    <button type="submit" className="w-full py-2 bg-amber-600 hover:bg-amber-500 text-white font-bold text-xs rounded-lg cursor-pointer">
                      Buka
                    </button>
                  </form>
                ) : (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between text-xs pb-2 border-b border-slate-800">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => setPlannerScope('current')}
                          className={`px-2 py-0.5 rounded text-[11px] font-semibold cursor-pointer ${plannerScope === 'current' ? 'bg-amber-600 text-white' : 'text-slate-400'}`}
                        >
                          Biro Ini ({currentActiveBiroTasks.length})
                        </button>
                        <button
                          onClick={() => setPlannerScope('all')}
                          className={`px-2 py-0.5 rounded text-[11px] font-semibold cursor-pointer ${plannerScope === 'all' ? 'bg-amber-600 text-white' : 'text-slate-400'}`}
                        >
                          Semua Biro ({allSubmittedTasksList.length})
                        </button>
                      </div>
                      <button onClick={() => setIsPlannerUnlocked(false)} className="text-[11px] text-slate-400 hover:underline">
                        Kunci
                      </button>
                    </div>

                    <div className="max-h-80 overflow-y-auto space-y-2">
                      {plannerTasksToShow.length > 0 ? (
                        plannerTasksToShow.map((task, idx) => (
                          <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                            <div className="min-w-0">
                              <div className="font-semibold text-white truncate">{task.pic} <span className="font-mono text-slate-500 font-normal">#{idx + 1}</span></div>
                              <div className="text-slate-400 text-[11px] truncate">{task.taskName}</div>
                              <div className="text-emerald-400 font-mono text-[10px]">{task.project} • {task.biroName}</div>
                            </div>

                            <div className="flex items-center gap-1.5 shrink-0">
                              <input
                                type="text"
                                value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                                onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                                placeholder="Kode JC..."
                                className="w-28 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase focus:outline-none"
                              />
                              <button
                                onClick={() => handleSaveKodeJcForTask(task.id)}
                                className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white font-bold rounded text-xs cursor-pointer"
                              >
                                Simpan
                              </button>
                            </div>
                          </div>
                        ))
                      ) : (
                        <div className="py-8 text-center text-xs text-slate-500">Tidak ada pengajuan tugas</div>
                      )}
                    </div>
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