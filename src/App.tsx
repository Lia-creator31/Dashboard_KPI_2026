import { useState, useEffect, useCallback, useMemo } from 'react';
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
  FileSpreadsheet, 
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
  BarChart3,
  PieChart,
  Clock,
  UserCheck,
  TrendingUp,
  Briefcase,
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
  const b1 = (biro1 || '').toLowerCase().trim();
  const b2 = (biro2 || '').toLowerCase().trim();
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

  const c1 = cleanText(b1.replace(/biro|departemen|dept|divisi/gi, ''));
  const c2 = cleanText(b2.replace(/biro|departemen|dept|divisi/gi, ''));
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
      // lanjut ke file berikutnya
    }
  }
  return null;
}

export default function App() {
  const [activeMainTab, setActiveMainTab] = useState<'operational' | 'dashboard'>('operational');
  const [dashboardDeptFilter, setDashboardDeptFilter] = useState<string>('ALL');

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
  const [rawJobCards, setRawJobCards] = useState<any[]>([]);

  const loadAllJobCards = useCallback(async () => {
    try {
      const { data, error } = await supabase
        .from('job_cards')
        .select('*')
        .order('created_at', { ascending: false });

      if (error) return;

      if (data) {
        setRawJobCards(data);
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
            '/Struktur_dan_Anggota_Desain_Upd_0826.xlsx'
          ])
        ]);

        if (wbKpi) setWorkbook(wbKpi);
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbStruktur) setStrukturWorkbook(wbStruktur);
      } catch {
        // fallback
      } finally {
        setIsLoadingExcel(false);
      }
    }
    loadAllExcelFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

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

  const getBiroMembers = (biroName: string): string[] => {
    const members: string[] = [];

    if (strukturWorkbook) {
      const sheet = strukturWorkbook.Sheets['CalonPers'] || strukturWorkbook.Sheets[strukturWorkbook.SheetNames[0]];
      if (sheet) {
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        const colBlocks = [
          { dCol: 2, nCol: 3 },
          { dCol: 8, nCol: 9 },
          { dCol: 14, nCol: 15 },
          { dCol: 22, nCol: 23 },
          { dCol: 29, nCol: 30 },
          { dCol: 36, nCol: 37 },
        ];

        colBlocks.forEach(({ dCol, nCol }) => {
          let currentBiro = '';
          for (let r = 10; r < rows.length; r++) {
            const row = rows[r];
            if (!row) continue;
            const valD = String(row[dCol] || '').trim();
            const valN = String(row[nCol] || '').trim();

            if (valD.toLowerCase().includes('biro')) {
              currentBiro = valD;
            } else if (currentBiro && isBiroMatch(currentBiro, biroName)) {
              if (valD && !/^\d+$/.test(valD) && (!valN || valN.toLowerCase() === 'nan') && !valD.toLowerCase().includes('departemen') && !valD.toLowerCase().includes('personil') && !valD.toLowerCase().includes('total')) {
                if (!members.includes(valD) && valD.toLowerCase() !== 'nan') members.push(valD);
              } else if (valN && valN.toLowerCase() !== 'nan' && valN.toUpperCase() !== 'PERSONIL' && !valN.toLowerCase().includes('departemen')) {
                if (!members.includes(valN)) members.push(valN);
              }
            }
          }
        });
      }
    }

    if (members.length === 0 && workbook) {
      workbook.SheetNames.forEach(sName => {
        const s = workbook.Sheets[sName];
        if (!s) return;
        const kRows: any[][] = XLSX.utils.sheet_to_json(s, { header: 1, defval: '' });
        kRows.forEach((r, idx) => {
          if (idx < 1 || !r) return;
          const bCol = String(r[7] || '').trim();
          const namaCol = String(r[1] || '').trim();
          if (namaCol && isBiroMatch(bCol, biroName) && !members.includes(namaCol)) {
            members.push(namaCol);
          }
        });
      });
    }

    return members.sort();
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
    const personMap = new Map<string, TaskItem[]>();

    biroMembers.forEach(m => personMap.set(m, []));
    const biroKey = cleanText(targetBiroName);
    const tasksForThisBiro = manualTasks[biroKey] || [];

    tasksForThisBiro.forEach(t => {
      let matchedName = biroMembers.find(m => cleanText(m) === cleanText(t.pic)) || t.pic;
      if (!personMap.has(matchedName)) personMap.set(matchedName, []);
      personMap.get(matchedName)!.push(t);
    });

    const result: PersonilCardGroup[] = [];
    personMap.forEach((tasks, picName) => result.push({ picName, tasks }));
    return result.sort((a, b) => a.picName.localeCompare(b.picName));
  };

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedFormBiro) return;

    try {
      let validBiroId: string | null = null;
      const { data: biroList } = await supabase.from('biros').select('id, name');
      
      if (biroList && biroList.length > 0) {
        const found = biroList.find(b => isBiroMatch(b.name, selectedFormBiro.biroName));
        validBiroId = found ? found.id : biroList[0].id;
      }

      const { data: insertedRow, error } = await supabase
        .from('job_cards')
        .insert({
          biro_id: validBiroId,
          biro_name: selectedFormBiro.biroName,
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

      const biroKey = cleanText(selectedFormBiro.biroName);
      const newTask: TaskItem = {
        id: insertedRow.id,
        biroName: selectedFormBiro.biroName,
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

  const handleClearAllBiroData = async () => {
    if (!selectedFormBiro) return;
    if (window.confirm(`Kosongkan semua data tugas di ${selectedFormBiro.biroName}?`)) {
      const { error } = await supabase
        .from('job_cards')
        .delete()
        .eq('biro_name', selectedFormBiro.biroName);

      if (error) {
        alert('Gagal mengosongkan data: ' + error.message);
        return;
      }

      const biroKey = cleanText(selectedFormBiro.biroName);
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

  const totalPersonilCount = accordionData.length;
  const totalTasksCount = accordionData.reduce((acc, g) => acc + g.tasks.length, 0);

  const currentBiroKey = selectedFormBiro ? cleanText(selectedFormBiro.biroName) : '';
  const currentBiroSubmittedTasks = manualTasks[currentBiroKey] || [];
  const pendingTasksCount = currentBiroSubmittedTasks.filter(t => !t.kodeJc).length;

  // =========================================================================
  // KOMPUTASI 100% DATA RIIL DENGAN SKALA VISUAL LEBIH BESAR
  // =========================================================================
  const dashboardAnalytics = useMemo(() => {
    const isFiltered = dashboardDeptFilter !== 'ALL';
    const targetDept = departmentsData.find(d => d.id === dashboardDeptFilter);

    const relevantTasks = isFiltered && targetDept
      ? rawJobCards.filter(r => targetDept.biros.some(b => isBiroMatch(b.name, r.biro_name)))
      : rawJobCards;

    const currentUnits = !isFiltered 
      ? departmentsData.map(d => ({ name: d.name, biros: d.biros }))
      : (targetDept ? targetDept.biros.map(b => ({ name: b.name, biros: [b] })) : []);

    // 1. GRAFIK 1: Beban Tugas Riil
    const unitDistribution = currentUnits.map(unit => {
      let count = 0;
      unit.biros.forEach(b => {
        const k = cleanText(b.name);
        count += (manualTasks[k] || []).length;
      });
      const shortName = unit.name
        .replace(/Departemen Desain |Departemen |Biro Desain Dasar |Biro /gi, '')
        .trim();
      return { 
        name: shortName, 
        fullName: unit.name, 
        count 
      };
    });
    const maxUnitCount = Math.max(...unitDistribution.map(u => u.count), 1);

    // 2. GRAFIK 2: Presensi, Terlambat & IPM Riil
    const activeAbsensiMap = parseAbsensiCSV(csvMonthMap['juni'] || csvMonthMap['mei'] || csvMonthMap['januari'] || '');
    
    let totalAllHadir = 0;
    let totalAllPossible = 0;

    const disciplineList = currentUnits.map(unit => {
      let totalTerlambat = 0;
      let totalIpm = 0;
      let totalSakit = 0;
      let memberCount = 0;

      unit.biros.forEach(b => {
        const members = getBiroMembers(b.name);
        memberCount += members.length;
        members.forEach(m => {
          const abs = activeAbsensiMap.get(cleanText(m));
          if (abs) {
            totalTerlambat += abs.terlambat;
            totalIpm += abs.ipm;
            totalSakit += abs.sakit;
          }
        });
      });

      const totalPossibleDays = Math.max(memberCount * 21, 1);
      const totalHadirDays = Math.max(0, totalPossibleDays - totalSakit);
      const kehadiranPct = memberCount > 0 ? Math.min(100, Math.round((totalHadirDays / totalPossibleDays) * 100)) : 0;

      totalAllHadir += totalHadirDays;
      totalAllPossible += totalPossibleDays;

      const shortName = unit.name
        .replace(/Departemen Desain |Departemen |Biro Desain Dasar |Biro /gi, '')
        .trim();

      return {
        name: shortName.length > 12 ? shortName.slice(0, 12) + '…' : shortName,
        fullName: unit.name,
        kehadiranPct,
        terlambatCount: totalTerlambat,
        ipmCount: totalIpm,
      };
    });
    const maxLateIpm = Math.max(...disciplineList.map(d => Math.max(d.terlambatCount, d.ipmCount)), 1);
    const overallAttendanceRate = totalAllPossible > 0 ? Math.round((totalAllHadir / totalAllPossible) * 100) : 95;

    // 3. GRAFIK 3: Utilisasi Personil Riil
    let totalHeadcount = 0;
    const allUnitMembers: string[] = [];
    currentUnits.forEach(u => u.biros.forEach(b => {
      const mems = getBiroMembers(b.name);
      totalHeadcount += mems.length;
      allUnitMembers.push(...mems);
    }));

    const activeAssignedSet = new Set<string>();
    relevantTasks.forEach(t => {
      const picName = (t.pic || t.personil_name || '').trim();
      if (picName) activeAssignedSet.add(cleanText(picName));
    });

    const activePersonilCount = allUnitMembers.filter(m => activeAssignedSet.has(cleanText(m))).length;
    const assignedPercent = totalHeadcount > 0 ? Math.min(100, Math.round((activePersonilCount / totalHeadcount) * 100)) : 0;
    const idlePercent = 100 - assignedPercent;

    // 4. GRAFIK 4: Tren Jam Efektif Riil per Bulan
    const hoursData = monthList.map(mName => {
      const mPrefix = mName.toLowerCase().slice(0, 3);
      let totalEffective = 0;

      if (workbook) {
        const targetSheet = workbook.SheetNames.find(s => {
          const sLow = s.toLowerCase();
          return sLow.includes(mName.toLowerCase()) || sLow.includes(mPrefix);
        });

        if (targetSheet && workbook.Sheets[targetSheet]) {
          const rows: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet], { header: 1, defval: '' });
          rows.forEach((row, rIdx) => {
            if (rIdx < 1 || !row) return;
            const biroCol = String(row[7] || '').trim();
            const effCol = parseValToNumber(row[10]);

            const isMatch = currentUnits.some(u => u.biros.some(b => isBiroMatch(biroCol, b.name)));
            if (isMatch) {
              totalEffective += effCol;
            }
          });
        }
      }

      return {
        label: mName,
        shortLabel: mName.slice(0, 3),
        effective: Math.round(totalEffective)
      };
    });
    const maxHoursVal = Math.max(...hoursData.map(h => h.effective), 1);

    // 5. GRAFIK 5: Porsi Proyek Riil
    const projectMap: Record<string, number> = {};
    relevantTasks.forEach(t => {
      const p = (t.project || 'Umum/Internal').trim();
      projectMap[p] = (projectMap[p] || 0) + 1;
    });

    const projectList = Object.entries(projectMap)
      .map(([project, count]) => ({ project, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 6);

    const totalProjectTasks = projectList.reduce((acc, p) => acc + p.count, 0);
    const projectColors = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ec4899'];

    let currentDeg = 0;
    const conicSegments = totalProjectTasks > 0 ? projectList.map((p, idx) => {
      const deg = (p.count / totalProjectTasks) * 360;
      const start = currentDeg;
      const end = currentDeg + deg;
      currentDeg = end;
      return `${projectColors[idx % projectColors.length]} ${start}deg ${end}deg`;
    }).join(', ') : '#334155 0deg 360deg';

    // 6. GRAFIK 6: Kepatuhan Timesheet Riil
    const activeTimesheetMap = parseTimesheetFolder('juni');
    
    const complianceTrend = currentUnits.map(unit => {
      let sumReguler = 0;
      let sumOvertime = 0;
      let countPerson = 0;

      unit.biros.forEach(b => {
        const members = getBiroMembers(b.name);
        members.forEach(m => {
          const ts = activeTimesheetMap.get(cleanText(m));
          if (ts) {
            sumReguler += ts.reguler;
            sumOvertime += ts.overtime;
            countPerson++;
          }
        });
      });

      const avgReguler = countPerson > 0 ? Math.min(100, Math.round(sumReguler / countPerson)) : 0;
      const avgOvertime = countPerson > 0 ? Math.min(100, Math.round(sumOvertime / countPerson)) : 0;
      const shortName = unit.name
        .replace(/Departemen Desain |Departemen |Biro Desain Dasar |Biro /gi, '')
        .trim();

      return {
        label: shortName.length > 10 ? shortName.slice(0, 10) + '…' : shortName,
        fullName: unit.name,
        regular: avgReguler,
        overtime: avgOvertime,
      };
    });

    return {
      unitDistribution,
      maxUnitCount,
      disciplineList,
      maxLateIpm,
      overallAttendanceRate,
      totalHeadcount,
      activePersonilCount,
      assignedPercent,
      idlePercent,
      hoursData,
      maxHoursVal,
      projectList,
      projectColors,
      totalProjectTasks,
      conicSegments,
      complianceTrend,
      totalJobCards: relevantTasks.length
    };
  }, [dashboardDeptFilter, departmentsData, rawJobCards, manualTasks, workbook, allCsvFiles]);

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans antialiased">
      {/* Top Navbar */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-14 flex items-center justify-between">
          <div 
            className="flex items-center gap-3 cursor-pointer select-none" 
            onClick={() => {
              setSelectedFormBiro(null);
              setSelectedBiroPage(null);
              setSelectedDept(null);
              setActiveMainTab('operational');
            }}
          >
            <div className="p-1.5 bg-blue-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-sm tracking-wide text-white block leading-tight">DIVISI DESAIN</span>
              <span className="text-[10px] text-slate-400">Executive & Operational System</span>
            </div>
          </div>

          {/* Switcher Tab Utama */}
          <div className="flex items-center gap-1 bg-slate-800 p-1 rounded-lg border border-slate-700">
            <button
              onClick={() => {
                setActiveMainTab('operational');
                setSelectedFormBiro(null);
                setSelectedBiroPage(null);
              }}
              className={`px-3.5 py-1 text-xs font-semibold rounded-md transition cursor-pointer flex items-center gap-1.5 ${
                activeMainTab === 'operational' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Operasional Biro
            </button>
            <button
              onClick={() => {
                setActiveMainTab('dashboard');
                setSelectedFormBiro(null);
                setSelectedBiroPage(null);
              }}
              className={`px-3.5 py-1 text-xs font-semibold rounded-md transition cursor-pointer flex items-center gap-1.5 ${
                activeMainTab === 'dashboard' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5" />
              Dashboard Grafis
            </button>
          </div>

          {selectedFormBiro ? (
            <button
              onClick={() => setSelectedFormBiro(null)}
              className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Biro
            </button>
          ) : selectedBiroPage ? (
            <button
              onClick={() => setSelectedBiroPage(null)}
              className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Biro
            </button>
          ) : selectedDept ? (
            <button
              onClick={() => setSelectedDept(null)}
              className="flex items-center gap-1.5 px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 transition cursor-pointer"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Departemen
            </button>
          ) : null}
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">

        {/* ========================================================================= */}
        {/* TAMPILAN 1: DASHBOARD EKSEKUTIF GRAFIS LEBIH BESAR & DETAIL (2 KOLOM)     */}
        {/* ========================================================================= */}
        {activeMainTab === 'dashboard' && (
          <div className="space-y-6 animate-fadeIn">
            {/* Header & Filter Divisi -> Departemen */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-900 border border-slate-800 p-5 rounded-2xl shadow-sm">
              <div>
                <h2 className="text-xl font-bold text-white tracking-wide flex items-center gap-2">
                  <BarChart3 className="w-5 h-5 text-blue-500" />
                  Workforce & Attendance Realtime Dashboard
                </h2>
                <span className="text-xs text-slate-400">Divisi Desain — Panel Pemantauan Terintegrasi Eksekutif</span>
              </div>

              {/* Filter Hierarki Departemen */}
              <div className="flex items-center gap-2.5">
                <span className="text-xs text-slate-400 font-semibold">Filter Unit:</span>
                <select
                  value={dashboardDeptFilter}
                  onChange={(e) => setDashboardDeptFilter(e.target.value)}
                  className="px-3.5 py-2 bg-slate-800 border border-slate-700 rounded-xl text-xs font-semibold text-white focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer"
                >
                  <option value="ALL">Semua Departemen (Divisi Level)</option>
                  {departmentsData.map(d => (
                    <option key={d.id} value={d.id}>{d.name}</option>
                  ))}
                </select>
              </div>
            </div>

            {/* PITA RINGKASAN METRIK EKSEKUTIF (BARU) */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-slate-900 border border-slate-800/80 p-4 rounded-xl flex items-center gap-3">
                <div className="p-3 bg-blue-500/10 text-blue-400 rounded-xl">
                  <Briefcase className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] text-slate-400 font-medium block">Total Job Card</span>
                  <span className="text-2xl font-black font-mono text-white">{dashboardAnalytics.totalJobCards}</span>
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800/80 p-4 rounded-xl flex items-center gap-3">
                <div className="p-3 bg-purple-500/10 text-purple-400 rounded-xl">
                  <Users className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] text-slate-400 font-medium block">Personil Teralokasi</span>
                  <span className="text-2xl font-black font-mono text-purple-400">{dashboardAnalytics.activePersonilCount} <span className="text-xs text-slate-500 font-normal">/ {dashboardAnalytics.totalHeadcount}</span></span>
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800/80 p-4 rounded-xl flex items-center gap-3">
                <div className="p-3 bg-cyan-500/10 text-cyan-400 rounded-xl">
                  <Clock className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] text-slate-400 font-medium block">Puncak Jam Efektif</span>
                  <span className="text-2xl font-black font-mono text-cyan-400">{dashboardAnalytics.maxHoursVal.toLocaleString('id-ID')} <span className="text-xs text-slate-500 font-normal">jam</span></span>
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800/80 p-4 rounded-xl flex items-center gap-3">
                <div className="p-3 bg-emerald-500/10 text-emerald-400 rounded-xl">
                  <TrendingUp className="w-5 h-5" />
                </div>
                <div>
                  <span className="text-[11px] text-slate-400 font-medium block">Rata-rata Presensi</span>
                  <span className="text-2xl font-black font-mono text-emerald-400">{dashboardAnalytics.overallAttendanceRate}%</span>
                </div>
              </div>
            </div>

            {/* GRID 6 PANEL GRAFIK LEGA (2 KOLOM) */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

              {/* GRAFIK 1 (Kiri Atas): Horizontal Bar Chart */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-5">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <BarChart3 className="w-4 h-4 text-blue-400" />
                        {dashboardDeptFilter === 'ALL' ? 'Distribusi Beban Tugas per Departemen' : 'Distribusi Beban Tugas per Biro'}
                      </span>
                      <span className="text-xs text-slate-400">Total akumulasi Job Card riil yang tersimpan</span>
                    </div>
                    <span className="text-xs px-2.5 py-1 bg-blue-500/10 text-blue-400 font-mono font-bold rounded-lg">Database Riil</span>
                  </div>

                  <div className="space-y-4 py-2">
                    {dashboardAnalytics.unitDistribution.map((unit, idx) => {
                      const widthPercent = (unit.count / dashboardAnalytics.maxUnitCount) * 100;
                      return (
                        <div key={idx} className="space-y-1.5">
                          <div className="flex justify-between text-xs font-medium">
                            <span className="text-slate-200 truncate max-w-[280px]" title={unit.fullName}>{unit.name}</span>
                            <span className="text-white font-mono font-bold bg-slate-800 px-2 py-0.5 rounded border border-slate-700">
                              {unit.count} <span className="text-slate-400 font-normal">tugas</span>
                            </span>
                          </div>
                          <div className="w-full bg-slate-800/80 h-3 rounded-full overflow-hidden">
                            <div 
                              className="bg-gradient-to-r from-blue-600 to-cyan-500 h-full rounded-full transition-all duration-700" 
                              style={{ width: `${unit.count > 0 ? Math.max(widthPercent, 5) : 0}%` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span>Jumlah Unit: {dashboardAnalytics.unitDistribution.length}</span>
                  <span className="text-blue-400 font-bold">Beban Puncak: {dashboardAnalytics.maxUnitCount} Tugas</span>
                </div>
              </div>

              {/* GRAFIK 2 (Kanan Atas): Grouped Bar Presensi, Terlambat, IPM Riil */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between border-b border-slate-800 pb-3 mb-5 gap-2">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <UserCheck className="w-4 h-4 text-emerald-400" />
                        {dashboardDeptFilter === 'ALL' ? 'Kedisiplinan & Presensi per Departemen' : 'Kedisiplinan & Presensi per Biro'}
                      </span>
                      <span className="text-xs text-slate-400">Perbandingan kehadiran, keterlambatan, dan izin pulang</span>
                    </div>

                    <div className="flex items-center gap-2.5 text-xs font-mono">
                      <span className="flex items-center gap-1.5 text-emerald-400"><div className="w-2.5 h-2.5 bg-emerald-500 rounded" /> Hadir(%)</span>
                      <span className="flex items-center gap-1.5 text-amber-400"><div className="w-2.5 h-2.5 bg-amber-500 rounded" /> Tlbt</span>
                      <span className="flex items-center gap-1.5 text-purple-400"><div className="w-2.5 h-2.5 bg-purple-500 rounded" /> IPM</span>
                    </div>
                  </div>

                  {/* Area Batang Lebih Tinggi (h-60) dengan Grid Lines */}
                  <div className="relative h-60 pt-4 px-2">
                    {/* Garis Panduan Nilai (Grid Lines) */}
                    <div className="absolute inset-x-2 inset-y-4 flex flex-col justify-between pointer-events-none opacity-15">
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-slate-400 w-full" />
                    </div>

                    <div className="relative h-full flex items-end justify-between gap-3">
                      {dashboardAnalytics.disciplineList.map((item, idx) => (
                        <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end">
                          <div className="w-full flex items-end justify-center gap-1 h-48">
                            {/* Bar Kehadiran */}
                            <div className="w-1/3 flex flex-col items-center justify-end h-full">
                              <span className="text-[10px] font-mono font-bold text-emerald-300 mb-1">{item.kehadiranPct}%</span>
                              <div 
                                className="w-full bg-emerald-500 rounded-t-md transition-all duration-500 hover:bg-emerald-400" 
                                style={{ height: `${item.kehadiranPct * 0.85}%` }}
                                title={`${item.fullName} — Kehadiran: ${item.kehadiranPct}%`}
                              />
                            </div>

                            {/* Bar Terlambat */}
                            <div className="w-1/3 flex flex-col items-center justify-end h-full">
                              <span className="text-[10px] font-mono font-bold text-amber-300 mb-1">{item.terlambatCount}</span>
                              <div 
                                className="w-full bg-amber-500 rounded-t-md transition-all duration-500 hover:bg-amber-400" 
                                style={{ height: `${item.terlambatCount > 0 ? Math.max((item.terlambatCount / dashboardAnalytics.maxLateIpm) * 85, 8) : 0}%` }}
                                title={`${item.fullName} — Terlambat: ${item.terlambatCount} kali`}
                              />
                            </div>

                            {/* Bar IPM */}
                            <div className="w-1/3 flex flex-col items-center justify-end h-full">
                              <span className="text-[10px] font-mono font-bold text-purple-300 mb-1">{item.ipmCount}</span>
                              <div 
                                className="w-full bg-purple-500 rounded-t-md transition-all duration-500 hover:bg-purple-400" 
                                style={{ height: `${item.ipmCount > 0 ? Math.max((item.ipmCount / dashboardAnalytics.maxLateIpm) * 85, 6) : 0}%` }}
                                title={`${item.fullName} — IPM: ${item.ipmCount} kali`}
                              />
                            </div>
                          </div>
                          <span className="text-xs font-mono text-slate-300 truncate max-w-[75px] mt-2 text-center" title={item.fullName}>{item.name}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span className="text-emerald-400">Kepatuhan Terintegrasi</span>
                  <span className="text-slate-400">File CSV Absensi Riil</span>
                </div>
              </div>

              {/* GRAFIK 3 (Kiri Bawah): Ring Donut Gauge Lebih Besar */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-5">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <Users className="w-4 h-4 text-purple-400" />
                        Tingkat Utilisasi Personil Desain
                      </span>
                      <span className="text-xs text-slate-400">Rasio personil yang memiliki beban tugas aktif</span>
                    </div>
                    <span className="text-xs px-2.5 py-1 bg-purple-500/10 text-purple-400 font-mono font-bold rounded-lg">Kapabilitas SDM</span>
                  </div>

                  <div className="flex flex-col sm:flex-row items-center justify-around gap-6 py-4">
                    {/* Ring Donut Besar (w-48 h-48) */}
                    <div className="relative w-48 h-48 shrink-0 flex items-center justify-center">
                      <div 
                        className="w-full h-full rounded-full transition-all duration-700 shadow-inner"
                        style={{
                          background: `conic-gradient(#8b5cf6 0% ${dashboardAnalytics.assignedPercent}%, #334155 ${dashboardAnalytics.assignedPercent}% 100%)`
                        }}
                      />
                      <div className="absolute w-32 h-32 bg-slate-900 rounded-full flex flex-col items-center justify-center shadow-lg border border-slate-800">
                        <span className="text-3xl font-black text-white font-mono">{dashboardAnalytics.assignedPercent}%</span>
                        <span className="text-[11px] text-purple-400 font-semibold tracking-wide uppercase mt-0.5">Teralokasi</span>
                      </div>
                    </div>

                    {/* Keterangan Angka Detail */}
                    <div className="space-y-3 flex-1 w-full max-w-xs">
                      <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700/60 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <div className="w-3 h-3 bg-purple-500 rounded-full" />
                          <span className="text-xs text-slate-300 font-medium">Personil Ditugaskan</span>
                        </div>
                        <span className="text-sm font-mono font-bold text-white">{dashboardAnalytics.activePersonilCount} Org</span>
                      </div>

                      <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700/60 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <div className="w-3 h-3 bg-slate-600 rounded-full" />
                          <span className="text-xs text-slate-300 font-medium">Personil Standby</span>
                        </div>
                        <span className="text-sm font-mono font-bold text-slate-400">{Math.max(0, dashboardAnalytics.totalHeadcount - dashboardAnalytics.activePersonilCount)} Org</span>
                      </div>

                      <div className="p-3 bg-slate-800/30 rounded-xl border border-slate-700/40 flex items-center justify-between">
                        <span className="text-xs text-slate-400">Total Personil Master</span>
                        <span className="text-sm font-mono font-bold text-purple-300">{dashboardAnalytics.totalHeadcount} Org</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span className="text-purple-400">Kapasitas Kerja</span>
                  <span className="text-slate-300">{dashboardAnalytics.assignedPercent >= 70 ? 'Alokasi Optimal' : 'Perlu Penugasan'}</span>
                </div>
              </div>

              {/* GRAFIK 4 (Kanan Bawah): Tren Jam Efektif Riil Lebih Tinggi */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-5">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <Clock className="w-4 h-4 text-cyan-400" />
                        Tren Jam Efektif Riil Bulanan
                      </span>
                      <span className="text-xs text-slate-400">Akumulasi Effective Hour per lembar data_kpi.xlsx</span>
                    </div>
                    <span className="text-xs px-2.5 py-1 bg-cyan-500/10 text-cyan-400 font-mono font-bold rounded-lg">data_kpi.xlsx</span>
                  </div>

                  {/* Area Batang Vertikal Lebih Tinggi (h-60) */}
                  <div className="relative h-60 pt-4 px-2">
                    <div className="absolute inset-x-2 inset-y-4 flex flex-col justify-between pointer-events-none opacity-15">
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-slate-400 w-full" />
                    </div>

                    <div className="relative h-full flex items-end justify-between gap-3">
                      {dashboardAnalytics.hoursData.map((h, idx) => (
                        <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end">
                          <div className="w-full flex flex-col items-center justify-end h-48">
                            <span className="text-[10px] font-mono font-bold text-cyan-300 mb-1">
                              {h.effective > 0 ? h.effective.toLocaleString('id-ID') : '0'}
                            </span>
                            <div 
                              className="w-10 bg-gradient-to-t from-cyan-600 to-cyan-400 rounded-t-md transition-all duration-500 hover:brightness-110" 
                              style={{ height: `${h.effective > 0 ? (h.effective / dashboardAnalytics.maxHoursVal) * 85 : 4}%` }}
                              title={`${h.label}: ${h.effective.toLocaleString('id-ID')} Jam`}
                            />
                          </div>
                          <span className="text-xs font-mono text-slate-300 mt-2">{h.shortLabel}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span>Semester I 2026</span>
                  <span className="text-cyan-400 font-bold">Puncak: {dashboardAnalytics.maxHoursVal.toLocaleString('id-ID')} Jam</span>
                </div>
              </div>

              {/* GRAFIK 5 (Kiri Bawah): Porsi Proyek Kapal Riil */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-5">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <PieChart className="w-4 h-4 text-amber-400" />
                        Porsi Beban per Proyek Kapal Riil
                      </span>
                      <span className="text-xs text-slate-400">Komposisi sebaran proyek dari input formulir Job Card</span>
                    </div>
                    <span className="text-xs px-2.5 py-1 bg-amber-500/10 text-amber-400 font-mono font-bold rounded-lg">Proyek Aktif</span>
                  </div>

                  {dashboardAnalytics.projectList.length > 0 ? (
                    <div className="flex flex-col sm:flex-row items-center justify-around gap-6 py-2">
                      <div className="relative w-44 h-44 shrink-0 flex items-center justify-center">
                        <div 
                          className="w-full h-full rounded-full transition-all duration-700 shadow-md"
                          style={{
                            background: `conic-gradient(${dashboardAnalytics.conicSegments})`
                          }}
                        />
                        <div className="absolute w-28 h-28 bg-slate-900 rounded-full flex flex-col items-center justify-center border border-slate-800">
                          <span className="text-xl font-bold text-white font-mono">{dashboardAnalytics.totalProjectTasks}</span>
                          <span className="text-[10px] text-slate-400 uppercase font-semibold">Total Tugas</span>
                        </div>
                      </div>

                      <div className="space-y-2.5 flex-1 w-full">
                        {dashboardAnalytics.projectList.map((p, idx) => {
                          const percent = dashboardAnalytics.totalProjectTasks > 0 
                            ? Math.round((p.count / dashboardAnalytics.totalProjectTasks) * 100) 
                            : 0;
                          return (
                            <div key={idx} className="flex items-center justify-between p-2 rounded-lg bg-slate-800/40 border border-slate-700/40">
                              <div className="flex items-center gap-2 truncate max-w-[180px]">
                                <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: dashboardAnalytics.projectColors[idx % dashboardAnalytics.projectColors.length] }} />
                                <span className="text-xs text-slate-200 font-medium truncate" title={p.project}>{p.project}</span>
                              </div>
                              <span className="text-xs font-mono text-white font-bold">{p.count} <span className="text-slate-400 font-normal">({percent}%)</span></span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ) : (
                    <div className="py-16 text-center text-slate-500 text-xs font-mono">
                      Belum ada Job Card proyek terdaftar di unit ini
                    </div>
                  )}
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span>Klasifikasi Proyek</span>
                  <span className="text-amber-400 font-bold">{dashboardAnalytics.projectList.length} Proyek Terdata</span>
                </div>
              </div>

              {/* GRAFIK 6 (Kanan Bawah): Kepatuhan Timesheet Riil Lebih Tinggi */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between border-b border-slate-800 pb-3 mb-5 gap-2">
                    <div>
                      <span className="text-sm font-bold text-white flex items-center gap-2">
                        <ShieldCheck className="w-4 h-4 text-rose-400" />
                        Kepatuhan Pengisian Timesheet Riil
                      </span>
                      <span className="text-xs text-slate-400">Rata-rata persentase pengisian jam reguler vs jam lembur</span>
                    </div>

                    <div className="flex items-center gap-2.5 text-xs font-mono">
                      <span className="flex items-center gap-1.5 text-indigo-400"><div className="w-2.5 h-2.5 bg-indigo-500 rounded" /> Reg(%)</span>
                      <span className="flex items-center gap-1.5 text-rose-400"><div className="w-2.5 h-2.5 bg-rose-500 rounded" /> Lbr(%)</span>
                    </div>
                  </div>

                  {/* Area Batang Lebih Tinggi (h-60) */}
                  <div className="relative h-60 pt-4 px-2">
                    <div className="absolute inset-x-2 inset-y-4 flex flex-col justify-between pointer-events-none opacity-15">
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-dashed border-slate-400 w-full" />
                      <div className="border-b border-slate-400 w-full" />
                    </div>

                    <div className="relative h-full flex items-end justify-between gap-3">
                      {dashboardAnalytics.complianceTrend.map((c, idx) => (
                        <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end">
                          <div className="w-full flex items-end justify-center gap-1.5 h-48">
                            {/* Bar Reguler */}
                            <div className="w-1/2 flex flex-col items-center justify-end h-full">
                              <span className="text-[10px] font-mono text-indigo-300 font-bold mb-1">{c.regular}%</span>
                              <div 
                                className="w-full bg-indigo-500 rounded-t-md transition-all duration-500 hover:bg-indigo-400" 
                                style={{ height: `${c.regular * 0.85}%` }}
                                title={`${c.fullName} — Reguler: ${c.regular}%`}
                              />
                            </div>

                            {/* Bar Lembur */}
                            <div className="w-1/2 flex flex-col items-center justify-end h-full">
                              <span className="text-[10px] font-mono text-rose-300 font-bold mb-1">{c.overtime}%</span>
                              <div 
                                className="w-full bg-rose-500 rounded-t-md transition-all duration-500 hover:bg-rose-400" 
                                style={{ height: `${Math.min(c.overtime * 0.85, 85)}%` }}
                                title={`${c.fullName} — Lembur: ${c.overtime}%`}
                              />
                            </div>
                          </div>
                          <span className="text-xs font-mono text-slate-300 truncate max-w-[75px] mt-2 text-center cursor-help" title={c.fullName}>{c.label}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="mt-6 pt-4 border-t border-slate-800 flex justify-between text-xs text-slate-400 font-mono">
                  <span>Folder CSV Timesheet</span>
                  <span className="text-indigo-400 font-bold">Terpantau Sistematis</span>
                </div>
              </div>

            </div>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAMPILAN 2: MODE OPERASIONAL BIRO (LEVEL 1 / LEVEL 2 / FORM / OUTPUT)     */}
        {/* ========================================================================= */}
        {activeMainTab === 'operational' && (
          <>
            {/* LEVEL 1: 6 DEPARTEMEN */}
            {!selectedDept && !selectedBiroPage && !selectedFormBiro && (
              <div className="space-y-6">
                <div className="flex items-center justify-between gap-4">
                  <h2 className="text-xl font-bold text-white">Departemen Desain</h2>
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
              <div className="space-y-5">
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
              <div className="space-y-4">
                <div className="bg-slate-800/80 border border-slate-700/80 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-base font-bold text-white">{selectedFormBiro.biroName}</h2>
                    <span className="text-xs text-slate-400 font-mono">
                      {totalPersonilCount} Personil • {totalTasksCount} Tugas
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
                          <label className="block text-xs font-medium text-slate-300 mb-1">Nama</label>
                          {currentBiroMembers.length > 0 ? (
                            <select
                              value={formData.nama}
                              onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                              required
                              className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                            >
                              <option value="">Pilih Nama...</option>
                              {currentBiroMembers.map((nama, idx) => (
                                <option key={idx} value={nama}>{nama}</option>
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
                          placeholder="Cari..."
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
                          onClick={handleClearAllBiroData}
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

                          return (
                            <div
                              key={idx}
                              className="bg-slate-800/60 border border-slate-700/70 rounded-xl overflow-hidden"
                            >
                              <div
                                onClick={() => toggleAccordion(person.picName)}
                                className="p-3.5 flex items-center justify-between cursor-pointer select-none hover:bg-slate-800/80 transition"
                              >
                                <div className="flex items-center gap-2.5">
                                  <div className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-bold ${
                                    isActive ? 'bg-emerald-600 text-white' : 'bg-slate-700 text-slate-400'
                                  }`}>
                                    <User className="w-4 h-4" />
                                  </div>
                                  <span className="font-semibold text-sm text-white">{person.picName}</span>
                                  <span className="text-xs font-mono text-slate-400">({taskCount})</span>
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
              <div className="space-y-4">
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

        {/* ================= MODAL KHUSUS PLANNER (MAS HASHFI) ================= */}
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
                                  onClick={() => handleSaveKodeJcForTask(task.id, selectedFormBiro!.biroName)}
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