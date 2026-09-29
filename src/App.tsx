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
  ChevronDown, 
  ChevronUp, 
  Trash2, 
  Lock, 
  KeyRound, 
  X, 
  Check, 
  Upload,
  Printer, 
  FileCheck,
  Clock,
  Wrench,
  CircleDollarSign,
  Users,
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

// File Excel utama
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
  rev?: string;
  realJo?: string;
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

// Inisial Unik Kode Biro untuk Work Order Subkon
function getBiroPrefix(biroName: string): string {
  const b = (biroName || '').toLowerCase().replace('&', ' dan ').trim();
  if (b.includes('dokumen') || (b.includes('perencanaan') && b.includes('biro'))) return 'DP';
  if (b.includes('logistik')) return 'DL';
  if (b.includes('administrasi')) return 'DA';
  if (b.includes('pengembangan')) return 'PD';
  if (b.includes('non kapal')) return 'NK';
  if (b.includes('kapal selam') || b.includes('submarine')) return 'KS';
  if (b.includes('kapal permukaan') || b.includes('surface')) return 'KP';
  if (b.includes('struktur') && b.includes('lambung')) return 'SL';
  if (b.includes('akomodasi')) return 'AK';
  if (b.includes('perlengkapan') && b.includes('lambung')) return 'PL';
  if (b.includes('produksi') && b.includes('lambung')) return 'PR';
  if (b.includes('propulsi')) return 'SP';
  if (b.includes('pengaturan') || b.includes('permesinan')) return 'PP';
  if (b.includes('hvac') || b.includes('geladak')) return 'HV';
  if (b.includes('listrik') || b.includes('kelistrikan')) return 'SK';
  if (b.includes('kontrol') || b.includes('otomasi')) return 'KO';
  if (b.includes('elektronika')) return 'SE';
  if (b.includes('hps')) return 'HP';

  const words = b.replace(/biro|desain|dasar/gi, '').split(/\s+/).filter(Boolean);
  if (words.length >= 2) return (words[0][0] + words[1][0]).toUpperCase();
  if (words.length === 1 && words[0].length >= 2) return words[0].slice(0, 2).toUpperCase();
  return 'WO';
}

function cleanProjectString(raw: any): string {
  if (!raw) return '';
  let s = String(raw)
    .replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ')
    .trim();
  
  s = s.replace(/[\.,;:\-_/\\\s]+$/, '').trim();
  s = s.toUpperCase();

  s = s.replace(/([A-Z0-9])O(\d+)$/, '$10$2');
  s = s.replace(/(\d)0P(\d)/, '$1OP$2');

  if (/^W[O0]{2,}\d+/.test(s)) {
    s = 'W' + s.slice(1).replace(/[O0]/g, '0');
  }

  return s.trim();
}

function getProjectNormKey(s: string): string {
  const k = (s || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  return k.replace(/0/g, 'o');
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
  const [subconPageMode, setSubconPageMode] = useState<'members' | 'form' | 'release'>('members');
  const [subconSearch, setSubconSearch] = useState('');

  // Organik Navigation State
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroPage, setSelectedBiroPage] = useState<SelectedBiroPage | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  const [formPageMode, setFormPageMode] = useState<'members' | 'form'>('members');

  const [searchQuery, setSearchQuery] = useState('');
  const [tableSearch, setTableSearch] = useState('');
  
  const [formData, setFormData] = useState({
    nama: '',
    kodeProyek: '',
    taskName: '',
    startDate: '',
    endDate: '',
    pic: '',
    jo: '',
    rev: '0',
    realJo: '',
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
            rev: row.rev || '0',
            realJo: row.real_jo || row.realJo || '',
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

  const toggleAccordion = (cardKey: string) => {
    setExpandedCards(prev => ({ ...prev, [cardKey]: !prev[cardKey] }));
  };

  // 3 File Excel Utama
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);

  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        let kpiUrl = '';
        let jcUrl = '';
        let im4Url = '';
        let realisasiUrl = '';

        Object.entries(excelGlobUrls).forEach(([path, url]) => {
          const pLower = path.toLowerCase();
          if (pLower.includes('kpi')) kpiUrl = url;
          else if (pLower.includes('jobcard') && !pLower.includes('realisasi')) jcUrl = url;
          else if (pLower.includes('im4') || pLower.includes('drawing') || pLower.includes('akses')) im4Url = url;
          else if (pLower.includes('realisasi')) realisasiUrl = url;
        });

        const [wbKpi, wbJc, wbIm4, wbRealisasi] = await Promise.all([
          fetchSafeWorkbook([kpiUrl, '/data_kpi.xlsx', './data_kpi.xlsx']),
          fetchSafeWorkbook([jcUrl, '/JOBCARD_DESAIN.xlsx', './JOBCARD_DESAIN.xlsx', '/JOBCARD DESAIN.xlsx']),
          fetchSafeWorkbook([
            im4Url,
            '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx',
            './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx',
            '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL.xlsx',
            './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL.xlsx'
          ]),
          fetchSafeWorkbook([
            realisasiUrl,
            '/Realisasi JO.xlsx',
            './Realisasi JO.xlsx',
            '/Realisasi_JO.xlsx',
            './Realisasi_JO.xlsx'
          ])
        ]);

        if (wbKpi) setWorkbook(wbKpi);
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
        if (wbRealisasi) setRealisasiWorkbook(wbRealisasi);
      } catch {
        // fallback
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

  const handleManualUploadRealisasi = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        setRealisasiWorkbook(wb);
        alert(`File ${file.name} berhasil dibaca! Real JO langsung terhitung otomatis.`);
      } catch {
        alert('Gagal membaca file Realisasi JO.xlsx.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const parseValToNumber = (val: any): number => {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    const str = String(val).trim().replace('%', '').replace(',', '.');
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  // Parser Realisasi JO (Effective + Overtime)
  const realisasiMap = useMemo(() => {
    const map = new Map<string, number>();
    if (!realisasiWorkbook) return map;

    realisasiWorkbook.SheetNames.forEach(sheetName => {
      const sheet = realisasiWorkbook.Sheets[sheetName];
      if (!sheet) return;

      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (rawRows.length === 0) return;

      let headerIdx = -1;
      let jobcardCol = -1;
      let effCol = -1;
      let otCol = -1;

      for (let r = 0; r < Math.min(5, rawRows.length); r++) {
        const row = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
        const jcIdx = row.findIndex(c => c.includes('jobcard'));
        const eIdx = row.findIndex(c => c.includes('effective'));
        const oIdx = row.findIndex(c => c.includes('overtime'));

        if (jcIdx !== -1 && (eIdx !== -1 || oIdx !== -1)) {
          headerIdx = r;
          jobcardCol = jcIdx;
          effCol = eIdx;
          otCol = oIdx;
          break;
        }
      }

      if (jobcardCol === -1) jobcardCol = 16;
      if (effCol === -1) effCol = 11;
      if (otCol === -1) otCol = 12;

      for (let r = Math.max(headerIdx + 1, 1); r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row) continue;

        const rawJc = String(row[jobcardCol] || '').trim();
        if (!rawJc || rawJc.toLowerCase() === 'nan' || rawJc.toLowerCase().includes('jobcard')) continue;

        const eff = effCol !== -1 ? parseValToNumber(row[effCol]) : 0;
        const ot = otCol !== -1 ? parseValToNumber(row[otCol]) : 0;
        const total = eff + ot;

        const key = cleanText(rawJc);
        if (key) {
          map.set(key, (map.get(key) || 0) + total);
        }
      }
    });

    return map;
  }, [realisasiWorkbook]);

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

  // HANYA MENGAMBIL PEGAWAI ORGANIK (PKWTT & PKWT)
  const getBiroMembers = useCallback((biroName: string): { nama: string; status: string; jabatan: string }[] => {
    const members = allParsedFromExcel.filter(
      p => isBiroMatch(p.biro, biroName) && !p.status.toLowerCase().includes('outsourcing')
    );
    if (members.length > 0) {
      return members.map(m => ({ nama: m.nama, status: m.status, jabatan: m.jabatan }));
    }
    return [];
  }, [allParsedFromExcel]);

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

  // Work Order Khusus Subkon
  const subconWorkOrders = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const prefix = getBiroPrefix(subconSelectedBiro);
    return currentActiveBiroTasks.map((task, idx) => ({
      ...task,
      packageTitle: `${prefix}${idx + 1}`
    }));
  }, [subconSelectedBiro, currentActiveBiroTasks]);

  // Deduplikasi Ketat Kode Proyek
  const projectOptions = useMemo((): string[] => {
    const projectMap = new Map<string, string>();

    if (jobcardWorkbook) {
      jobcardWorkbook.SheetNames.forEach(sheetName => {
        const sheet = jobcardWorkbook.Sheets[sheetName];
        if (!sheet) return;
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        let projCol = 2;
        for (let r = 0; r < Math.min(5, rows.length); r++) {
          const rowVals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
          const foundIdx = rowVals.findIndex(v => v.includes('proyek') || v.includes('project'));
          if (foundIdx !== -1) {
            projCol = foundIdx;
            break;
          }
        }

        rows.forEach((row, idx) => {
          if (idx < 1 || !row) return;
          const raw = String(row[projCol] || '').trim();
          if (
            !raw || 
            raw.toLowerCase() === 'nan' || 
            raw.toLowerCase().includes('kode proyek') || 
            raw.toLowerCase() === 'proyek' || 
            raw.toLowerCase() === 'project'
          ) return;

          const cleaned = cleanProjectString(raw);
          const normKey = getProjectNormKey(cleaned);

          if (cleaned && normKey && !projectMap.has(normKey)) {
            projectMap.set(normKey, cleaned);
          }
        });
      });
    }

    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        const raw = String(t.project || '').trim();
        if (raw) {
          const cleaned = cleanProjectString(raw);
          const normKey = getProjectNormKey(cleaned);
          if (cleaned && normKey && !projectMap.has(normKey)) {
            projectMap.set(normKey, cleaned);
          }
        }
      });
    });

    return Array.from(projectMap.values()).sort((a, b) => a.localeCompare(b));
  }, [jobcardWorkbook, manualTasks]);

  // Deduplikasi Ketat Deskripsi Tugas
  const taskOptions = useMemo((): string[] => {
    if (!jobcardWorkbook) return [];
    const taskMap = new Map<string, string>();

    jobcardWorkbook.SheetNames.forEach(sheetName => {
      const sheet = jobcardWorkbook.Sheets[sheetName];
      if (!sheet) return;
      const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      let taskCol = 3;
      for (let r = 0; r < Math.min(5, rows.length); r++) {
        const rowVals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
        const foundIdx = rowVals.findIndex(v => v.includes('task') || v.includes('deskripsi') || v.includes('uraian') || v.includes('pekerjaan'));
        if (foundIdx !== -1) {
          taskCol = foundIdx;
          break;
        }
      }

      rows.forEach((row, idx) => {
        if (idx < 1 || !row) return;
        const raw = String(row[taskCol] || '').replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ').trim();
        if (
          !raw || 
          raw.toLowerCase() === 'nan' || 
          raw.toLowerCase().includes('desc pekerjaan') || 
          raw.toLowerCase().includes('task name') || 
          raw.toLowerCase() === 'deskripsi'
        ) return;

        const cleaned = raw.replace(/[\.,;:\-_/\\\s]+$/, '').trim();
        const normKey = cleanText(cleaned);

        if (cleaned && normKey && !taskMap.has(normKey)) {
          taskMap.set(normKey, cleaned);
        }
      });
    });

    return Array.from(taskMap.values()).sort((a, b) => a.localeCompare(b));
  }, [jobcardWorkbook]);

  // Submit Form: Tanpa PIC manual
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

      const biroKey = cleanText(activeBiro);
      const currentList = manualTasks[biroKey] || [];

      // Otomatisasi Kode Work Order Subkon
      let autoKode = '';
      if (accessMode === 'subkon') {
        const prefix = getBiroPrefix(activeBiro);
        autoKode = `${prefix}${currentList.length + 1}`;
      }

      const insertPayload: any = {
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
        kode_jc: autoKode,
        status: accessMode === 'subkon' ? 'approved' : 'pending',
      };

      if (formData.rev) {
        insertPayload.rev = formData.rev;
      }

      let insertedRow: any = null;
      const { data: resData, error } = await supabase
        .from('job_cards')
        .insert(insertPayload)
        .select()
        .single();

      if (error) {
        if (error.message?.includes('rev') || (error as any).details?.includes('rev')) {
          delete insertPayload.rev;
          const { data: retryData, error: retryError } = await supabase
            .from('job_cards')
            .insert(insertPayload)
            .select()
            .single();
          if (retryError) {
            alert('Gagal simpan: ' + retryError.message);
            return;
          }
          insertedRow = retryData;
        } else {
          alert('Gagal simpan: ' + error.message);
          return;
        }
      } else {
        insertedRow = resData;
      }

      const newTask: TaskItem = {
        id: insertedRow.id,
        biroName: activeBiro,
        project: formData.kodeProyek,
        taskName: formData.taskName,
        startDate: formData.startDate,
        endDate: formData.endDate,
        pic: formData.nama,
        jo: formData.jo,
        kode_jc: autoKode,
        rev: formData.rev || '0',
      };

      setManualTasks({ ...manualTasks, [biroKey]: [...currentList, newTask] });

      setExpandedCards(prev => ({ ...prev, [formData.nama]: true }));
      setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', pic: '', jo: '', rev: '0', realJo: '' });

      if (accessMode === 'subkon') {
        alert(`Tugas tersimpan! Work Order "${autoKode}" langsung terbit dan dapat dilihat di bawah nama personil.`);
      } else {
        alert('Tugas tersimpan! Buka Planner untuk approval Jobcard.');
      }
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
      alert('Gagal simpan Jobcard: ' + error.message);
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

  const filteredDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const filteredSubconDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(subconSearch.toLowerCase())
  );

  const currentBiroMembers = selectedFormBiro ? getBiroMembers(selectedFormBiro.biroName) : [];
  const activeSubconMembers = subconSelectedBiro ? getSubconMembersForBiro(subconSelectedBiro) : [];

  const organicBiroTasks = useMemo(() => {
    return currentActiveBiroTasks.filter(t => 
      currentBiroMembers.some(m => cleanText(m.nama) === cleanText(t.pic))
    );
  }, [currentActiveBiroTasks, currentBiroMembers]);

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
                    const releaseCount = (manualTasks[cleanText(biro.name)] || []).length;

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

            {/* Step 3: Halaman Biro Subkon */}
            {subconSelectedBiro && (
              <div className="space-y-4">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-white text-base">{subconSelectedBiro}</h3>
                    <span className="text-xs text-slate-400 font-mono">
                      {activeSubconMembers.length} Personel • Kode Biro: <b className="text-amber-400">{getBiroPrefix(subconSelectedBiro)}</b>
                    </span>
                  </div>

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
                      onClick={() => setSubconPageMode('release')}
                      className={`px-3 py-1 rounded transition cursor-pointer flex items-center gap-1 ${subconPageMode === 'release' ? 'bg-purple-600 text-white' : 'text-purple-300 hover:text-white'}`}
                    >
                      <FileCheck className="w-3 h-3 text-amber-300" /> Work Order ({subconWorkOrders.length})
                    </button>
                  </div>
                </div>

                {/* TAB 1: ANGGOTA SUBKON */}
                {subconPageMode === 'members' && (
                  <div className="space-y-2.5">
                    {activeSubconMembers.length > 0 ? (
                      activeSubconMembers.map((person, idx) => {
                        const personTasks = currentActiveBiroTasks.filter(
                          t => cleanText(t.pic) === cleanText(person.nama)
                        );
                        const isExpanded = !!expandedCards[person.nama];

                        return (
                          <div 
                            key={idx} 
                            className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden transition-all duration-200"
                          >
                            <div 
                              onClick={() => toggleAccordion(person.nama)}
                              className="p-3.5 flex items-center justify-between cursor-pointer select-none hover:bg-slate-800/40 transition"
                            >
                              <div className="flex items-center gap-3">
                                <div className="w-8 h-8 rounded-lg bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center font-bold text-xs">
                                  {person.nama.charAt(0)}
                                </div>
                                <div>
                                  <div className="font-semibold text-white text-sm">{person.nama}</div>
                                  <div className="text-xs text-slate-400 font-mono">
                                    NIP: {person.nip || '-'} • <span className="text-cyan-400">{person.jabatan}</span>
                                  </div>
                                </div>
                              </div>

                              <div className="flex items-center gap-3">
                                <span className="text-xs font-mono text-slate-400 hidden sm:inline">
                                  {personTasks.length} Tugas
                                </span>
                                <div className={`p-1.5 rounded-lg bg-slate-950 border border-slate-800 text-slate-400 transition-transform duration-200 ${isExpanded ? 'text-amber-400' : ''}`}>
                                  {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                                </div>
                              </div>
                            </div>

                            {isExpanded && (
                              <div className="p-3.5 border-t border-slate-800/80 bg-slate-950/60 animate-fadeIn">
                                {personTasks.length > 0 ? (
                                  <div className="overflow-x-auto">
                                    <table className="w-full text-left text-xs">
                                      <thead>
                                        <tr className="text-slate-400 border-b border-slate-800">
                                          <th className="py-2 px-2.5 w-8">#</th>
                                          <th className="py-2 px-2.5 font-mono text-amber-400">Kode WO</th>
                                          <th className="py-2 px-2.5">Proyek</th>
                                          <th className="py-2 px-2.5">Deskripsi</th>
                                          <th className="py-2 px-2.5 font-mono">Jadwal</th>
                                          <th className="py-2 px-2.5 font-mono">JO</th>
                                          <th className="py-2 px-2.5 text-center w-12">Aksi</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-800 text-slate-300">
                                        {personTasks.map((task, tIdx) => {
                                          const isSameProjectAsAbove = tIdx > 0 && task.project === personTasks[tIdx - 1].project;

                                          return (
                                            <tr key={task.id} className="hover:bg-slate-900/40">
                                              <td className="py-2 px-2.5 text-slate-500 font-mono">{tIdx + 1}</td>
                                              <td className="py-2 px-2.5 font-mono font-bold text-amber-300">
                                                {task.kodeJc || `${getBiroPrefix(subconSelectedBiro)}${tIdx + 1}`}
                                              </td>
                                              <td className="py-2 px-2.5 font-medium">
                                                {isSameProjectAsAbove ? (
                                                  <span className="text-slate-500 font-mono text-[11px]" title={task.project}>
                                                    — s.d.a —
                                                  </span>
                                                ) : (
                                                  <span className="text-emerald-400">{task.project}</span>
                                                )}
                                              </td>
                                              <td className="py-2 px-2.5 text-slate-200">{task.taskName}</td>
                                              <td className="py-2 px-2.5 font-mono text-[11px] text-slate-400">
                                                {task.startDate} s/d {task.endDate}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-violet-300">#{task.jo}</td>
                                              <td className="py-2 px-2.5 text-center">
                                                <button 
                                                  onClick={(e) => {
                                                    e.stopPropagation();
                                                    handleDeleteTask(task.id);
                                                  }} 
                                                  className="text-slate-500 hover:text-rose-400 p-1 cursor-pointer"
                                                  title="Hapus Tugas"
                                                >
                                                  <Trash2 className="w-3.5 h-3.5" />
                                                </button>
                                              </td>
                                            </tr>
                                          );
                                        })}
                                      </tbody>
                                    </table>
                                  </div>
                                ) : (
                                  <div className="text-xs text-slate-500 py-3 text-center italic">
                                    Belum ada tugas. Buka tab <b>Form</b> di atas untuk menambahkan penugasan.
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })
                    ) : (
                      <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
                        Tidak ada anggota di biro ini
                      </div>
                    )}
                  </div>
                )}

                {/* TAB 2: FORM SUBKON */}
                {subconPageMode === 'form' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Nama Drafter / Personel</label>
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
                          <label className="block text-slate-400 mb-1">Deskripsi</label>
                          <select
                            value={formData.taskName}
                            onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none"
                          >
                            <option value="">Pilih Deskripsi Pekerjaan...</option>
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
                            style={{ colorScheme: 'dark' }}
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer"
                          />
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Tanggal Selesai</label>
                          <input
                            type="date"
                            value={formData.endDate}
                            onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                            required
                            style={{ colorScheme: 'dark' }}
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer"
                          />
                        </div>
                      </div>

                      <div className="pt-3 border-t border-slate-800 flex justify-end">
                        <button
                          type="submit"
                          className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-white font-bold text-xs rounded-lg transition cursor-pointer"
                        >
                          Simpan Tugas & Terbitkan Work Order
                        </button>
                      </div>
                    </form>
                  </div>
                )}

                {/* TAB 3: WORK ORDER SUBKON */}
                {subconPageMode === 'release' && (
                  <div className="space-y-3">
                    {subconWorkOrders.length > 0 ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {subconWorkOrders.map((pkg) => (
                          <div key={pkg.id} className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-2">
                            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                              <span className="px-2.5 py-0.5 bg-purple-600 rounded text-white font-mono font-bold text-sm tracking-wide shadow-sm">
                                {pkg.packageTitle}
                              </span>
                              <span className="font-mono text-xs text-amber-400 font-bold bg-amber-500/10 px-2 py-0.5 rounded border border-amber-500/20">
                                {pkg.kodeJc || pkg.packageTitle}
                              </span>
                            </div>
                            <div className="text-xs space-y-1">
                              <div className="text-white font-semibold">{pkg.pic}</div>
                              <div className="text-emerald-400 font-mono">{pkg.project} • #{pkg.jo}</div>
                              <div className="text-slate-300 text-[11px] bg-slate-950 p-2 rounded border border-slate-800">{pkg.taskName}</div>
                            </div>
                            <div className="text-[10px] text-slate-500 pt-2 border-t border-slate-800 flex justify-between font-mono">
                              <span>{pkg.startDate} s/d {pkg.endDate}</span>
                              <button onClick={() => alert(`Mencetak Work Order ${pkg.packageTitle}...`)} className="text-amber-400 hover:underline cursor-pointer flex items-center gap-1">
                                <Printer className="w-3 h-3" /> Cetak
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="bg-slate-900 border border-slate-800 rounded-xl p-8 text-center text-xs text-slate-500">
                        Belum ada Work Order di biro ini. Tambahkan tugas di tab <b>Form</b> untuk langsung menerbitkan Work Order.
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
                    <span className="text-xs text-slate-400">Akses Pegawai Organik (PKWTT & PKWT)</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Clock className="w-3.5 h-3.5 text-emerald-400" />
                      <span>{realisasiMap.size > 0 ? `Realisasi JO: OK (${realisasiMap.size})` : 'Upload Realisasi JO.xlsx'}</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
                    </label>

                    <input
                      type="text"
                      placeholder="Cari..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44"
                    />
                  </div>
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
                  <span className="text-xs text-slate-400">Daftar Biro Penugasan Pegawai Organik</span>
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
                          onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('members'); }}
                          className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs rounded cursor-pointer"
                        >
                          Anggota
                        </button>
                        <button
                          onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('form'); }}
                          className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs rounded cursor-pointer"
                        >
                          Form
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Step 3: Halaman Detail Biro Organik */}
            {selectedFormBiro && (
              <div className="space-y-4">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-white text-sm">{selectedFormBiro.biroName}</h3>
                    <span className="text-xs text-slate-400 font-mono">
                      {currentBiroMembers.length} Pegawai Organik • Real JO: {realisasiMap.size > 0 ? `${realisasiMap.size} Jobcard Terdaftar` : 'Excel Realisasi Belum Dimuat'}
                    </span>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <label className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Clock className="w-3 h-3 text-emerald-400" />
                      <span>{realisasiMap.size > 0 ? 'Update Realisasi JO' : 'Upload Realisasi JO'}</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
                    </label>

                    <div className="flex gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 text-xs">
                      <button
                        onClick={() => setFormPageMode('members')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${formPageMode === 'members' ? 'bg-blue-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}
                      >
                        Anggota ({currentBiroMembers.length})
                      </button>
                      <button
                        onClick={() => setFormPageMode('form')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${formPageMode === 'form' ? 'bg-emerald-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}
                      >
                        Form
                      </button>
                    </div>

                    <button
                      onClick={() => setIsPlannerModalOpen(true)}
                      className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
                    >
                      <Lock className="w-3 h-3" /> Planner
                    </button>
                  </div>
                </div>

                {/* TAB 1: ANGGOTA ORGANIK (TABEL KOLOM: JOBCARD, PLAN JO & REAL JO DARI EXCEL) */}
                {formPageMode === 'members' ? (
                  <div className="space-y-2.5">
                    {currentBiroMembers.length > 0 ? (
                      currentBiroMembers.map((person, idx) => {
                        const personTasks = currentActiveBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
                        const isExpanded = !!expandedCards[person.nama];

                        return (
                          <div key={idx} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden transition-all duration-200">
                            {/* Baris Nama Personel (Header Dropdown) */}
                            <div 
                              onClick={() => toggleAccordion(person.nama)}
                              className="p-3.5 flex items-center justify-between cursor-pointer select-none hover:bg-slate-800/40 transition"
                            >
                              <div className="flex items-center gap-3">
                                <div className="w-8 h-8 rounded-lg bg-blue-500/10 border border-blue-500/20 text-blue-400 flex items-center justify-center font-bold text-xs">
                                  {person.nama.charAt(0)}
                                </div>
                                <div>
                                  <div className="font-semibold text-white text-sm">
                                    {person.nama} <span className="text-xs font-mono text-cyan-400 font-normal">({person.status})</span>
                                  </div>
                                  <div className="text-xs text-slate-400 font-mono">
                                    NIP: {person.nip || '-'} • <span className="text-slate-400">{person.jabatan}</span>
                                  </div>
                                </div>
                              </div>

                              <div className="flex items-center gap-3">
                                <span className="text-xs font-mono text-slate-400 hidden sm:inline">
                                  {personTasks.length} Tugas
                                </span>
                                <div className={`p-1.5 rounded-lg bg-slate-950 border border-slate-800 text-slate-400 transition-transform duration-200 ${isExpanded ? 'text-blue-400' : ''}`}>
                                  {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                                </div>
                              </div>
                            </div>

                            {/* OUTPUT TABEL BERKOLOM */}
                            {isExpanded && (
                              <div className="p-3.5 border-t border-slate-800/80 bg-slate-950/60 animate-fadeIn">
                                {personTasks.length > 0 ? (
                                  <div className="overflow-x-auto">
                                    <table className="w-full text-left text-xs">
                                      <thead>
                                        <tr className="text-slate-400 border-b border-slate-800">
                                          <th className="py-2 px-2.5 w-8">#</th>
                                          <th className="py-2 px-2.5 font-mono text-amber-400">Jobcard</th>
                                          <th className="py-2 px-2.5">Proyek</th>
                                          <th className="py-2 px-2.5">Deskripsi</th>
                                          <th className="py-2 px-2.5 text-center font-mono">Rev</th>
                                          <th className="py-2 px-2.5 font-mono text-cyan-400">Plan Start</th>
                                          <th className="py-2 px-2.5 font-mono text-cyan-400">Plan Finish</th>
                                          <th className="py-2 px-2.5 font-mono text-violet-300">Plan JO</th>
                                          <th className="py-2 px-2.5 font-mono text-emerald-400">Real JO</th>
                                          <th className="py-2 px-2.5 text-center w-12">Aksi</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-800 text-slate-300">
                                        {personTasks.map((task, tIdx) => {
                                          const isSameProjectAsAbove = tIdx > 0 && task.project === personTasks[tIdx - 1].project;
                                          const jcKey = cleanText(task.kodeJc || '');
                                          const calculatedRealHours = jcKey ? realisasiMap.get(jcKey) : undefined;

                                          return (
                                            <tr key={task.id} className="hover:bg-slate-900/40">
                                              <td className="py-2 px-2.5 text-slate-500 font-mono">{tIdx + 1}</td>
                                              <td className="py-2 px-2.5 font-mono font-bold text-amber-300">
                                                {task.kodeJc || <span className="text-rose-400 font-normal">Menunggu Planner</span>}
                                              </td>
                                              <td className="py-2 px-2.5 font-medium">
                                                {isSameProjectAsAbove ? (
                                                  <span className="text-slate-500 font-mono text-[11px]" title={task.project}>
                                                    — s.d.a —
                                                  </span>
                                                ) : (
                                                  <span className="text-emerald-400">{task.project}</span>
                                                )}
                                              </td>
                                              <td className="py-2 px-2.5 text-slate-200">{task.taskName}</td>
                                              <td className="py-2 px-2.5 text-center font-mono text-slate-300">
                                                {task.rev || '0'}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-[11px] text-slate-300">
                                                {task.startDate}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-[11px] text-slate-300">
                                                {task.endDate}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-violet-300">#{task.jo}</td>
                                              
                                              {/* Kolom Real JO: Nilai Otomatis dari Excel Realisasi JO.xlsx */}
                                              <td className="py-2 px-2.5 font-mono font-bold">
                                                {calculatedRealHours !== undefined ? (
                                                  <span className="px-2 py-0.5 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded">
                                                    {Math.round(calculatedRealHours * 100) / 100} Jam
                                                  </span>
                                                ) : task.kodeJc ? (
                                                  <span className="text-slate-500 font-normal">0 Jam</span>
                                                ) : (
                                                  <span className="text-slate-600 font-normal">-</span>
                                                )}
                                              </td>

                                              <td className="py-2 px-2.5 text-center">
                                                <button 
                                                  onClick={(e) => {
                                                    e.stopPropagation();
                                                    handleDeleteTask(task.id);
                                                  }} 
                                                  className="text-slate-500 hover:text-rose-400 p-1 cursor-pointer"
                                                  title="Hapus Tugas"
                                                >
                                                  <Trash2 className="w-3.5 h-3.5" />
                                                </button>
                                              </td>
                                            </tr>
                                          );
                                        })}
                                      </tbody>
                                    </table>
                                  </div>
                                ) : (
                                  <div className="text-xs text-slate-500 py-3 text-center italic">
                                    Belum ada tugas. Buka tab <b>Form</b> di atas untuk menambahkan penugasan.
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })
                    ) : (
                      <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
                        Tidak ada personil organik di biro ini
                      </div>
                    )}
                  </div>
                ) : (
                  /* TAB 2: FORM ORGANIK (IKON KALENDER TERBUKA & JELAS SAAT DIKLIK) */
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Nama Personel (Organik)</label>
                          <select
                            value={formData.nama}
                            onChange={(e) => setFormData(prev => ({ ...prev, nama: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                          >
                            <option value="">Pilih Personel Organik...</option>
                            {currentBiroMembers.map((p, i) => (
                              <option key={i} value={p.nama}>{p.nama} ({p.status})</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Kode Proyek</label>
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

                        {/* Plan JO */}
                        <div>
                          <label className="block text-slate-400 mb-1">Plan JO</label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formData.jo}
                            onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                            placeholder="Contoh: 300426"
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                          />
                        </div>

                        {/* Deskripsi */}
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Deskripsi</label>
                          <select
                            value={formData.taskName}
                            onChange={(e) => setFormData(prev => ({ ...prev, taskName: e.target.value }))}
                            required
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                          >
                            <option value="">Pilih Deskripsi Pekerjaan...</option>
                            {taskOptions.map((t, i) => (
                              <option key={i} value={t}>{t}</option>
                            ))}
                          </select>
                        </div>

                        {/* Rev */}
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Rev (Revisi)</label>
                          <input
                            type="text"
                            value={formData.rev}
                            onChange={(e) => setFormData(prev => ({ ...prev, rev: e.target.value }))}
                            placeholder="0"
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                          />
                        </div>

                        {/* Plan Start (Kalender Asli Putih Terang, Langsung Muncul Saat Diklik) */}
                        <div>
                          <label className="block text-slate-400 mb-1">Plan Start</label>
                          <input 
                            type="date" 
                            value={formData.startDate} 
                            onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))} 
                            required 
                            style={{ colorScheme: 'dark' }}
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" 
                          />
                        </div>

                        {/* Plan Finish (Kalender Asli Putih Terang, Langsung Muncul Saat Diklik) */}
                        <div>
                          <label className="block text-slate-400 mb-1">Plan Finish</label>
                          <input 
                            type="date" 
                            value={formData.endDate} 
                            onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))} 
                            required 
                            style={{ colorScheme: 'dark' }}
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" 
                          />
                        </div>
                      </div>

                      <div className="flex justify-end pt-2">
                        <button type="submit" className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg cursor-pointer">
                          Simpan Tugas
                        </button>
                      </div>
                    </form>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ================= MODAL PLANNER (INPUT JOBCARD ORGANIK) ================= */}
        {isPlannerModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
              <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel — Approval Jobcard Organik
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
                      <span className="text-slate-400">Daftar Pengajuan Jobcard ({organicBiroTasks.length})</span>
                      <button onClick={() => setIsPlannerUnlocked(false)} className="text-[11px] text-slate-400 hover:underline cursor-pointer">
                        Kunci
                      </button>
                    </div>

                    <div className="max-h-80 overflow-y-auto space-y-2">
                      {organicBiroTasks.length > 0 ? (
                        organicBiroTasks.map((task, idx) => (
                          <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                            <div className="min-w-0">
                              <div className="font-semibold text-white truncate">
                                {task.pic} <span className="font-mono text-slate-500 font-normal">#{idx + 1}</span>
                                {task.rev && <span className="ml-2 font-mono text-[10px] text-cyan-400">Rev.{task.rev}</span>}
                              </div>
                              <div className="text-slate-400 text-[11px] truncate">{task.taskName}</div>
                              <div className="text-emerald-400 font-mono text-[10px]">{task.project} • {task.biroName}</div>
                            </div>

                            <div className="flex items-center gap-1.5 shrink-0">
                              <input
                                type="text"
                                value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                                onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                                placeholder="Jobcard (misal: JC020926 39833)..."
                                className="w-48 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase focus:outline-none"
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
                        <div className="py-8 text-center text-xs text-slate-500">Tidak ada pengajuan tugas organik</div>
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