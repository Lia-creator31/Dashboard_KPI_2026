import { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
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
  Pencil,
  Trash2, 
  Lock, 
  KeyRound, 
  X, 
  Check, 
  Upload, 
  Printer, 
  FileCheck, 
  Clock, 
  Sparkles, 
  FileSpreadsheet,
  Users, 
  Laptop, 
  LucideIcon 
} from 'lucide-react';

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

const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';

interface DriveSheetInfo {
  id: string;
  projectKey: string;
  title: string;
}

const GOOGLE_DRIVE_SHEETS: DriveSheetInfo[] = [
  { id: '1CBL96MejQnswfg_sLIkwkK9K594S9hDWoSGDBIeIBNo', projectKey: 'M000313', title: 'Drawing Control GEOMARIN V (M000313)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000304', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000305', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '19qiMBb1b7qmeRXkFNTbh-eKtVEKDUTgshloa116HQ48', projectKey: 'W000308', title: 'Drawing Control LPD UAE USED (W000308)' },
  { id: '16rrjyU60BmzRIYFL3I1WRLkw_AfJPUwi8vZiJx2LYCg', projectKey: 'W000314', title: 'Drawing Control KSR 20 M ALU (W000314)' },
  { id: '1uv2D4ECJ0z-NWP8byzXTkkM17MKD6f0r-RCtKjpG5Hs', projectKey: 'M000312', title: 'Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312)' },
  { id: '1eTO-dvVd3psdP4mFQc2ZQW_Z-lyooCb2TVjR4mYTbDw', projectKey: 'W000311', title: 'Drawing Control KSSR 18 M (W000311)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000306', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000307', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1l0ZoTCctHrmkruIkC4Gk1NtoqsugMVGABrsf3yb6PHo', projectKey: 'KSOT', title: 'Drawing Control KSOT' },
  { id: '1jGNHh9HhSrPlEv7JZvtW8ToMU2ADrT0yE87tk9E53zY', projectKey: 'LCU', title: 'Drawing Control LCU LD PN 124 M' },
  { id: '1XsrklFhcUkjgfEjuT4jDRWx6Uufsj8Jt', projectKey: 'FFBNW', title: 'Drawing Control FFBNW FRIGATE 140 M' },
];

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
  release?: string;
}

interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
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

interface SelectOption {
  value: string;
  label: string;
}

function SearchableSelect({
  options,
  value,
  onChange,
  placeholder = 'Ketik untuk mencari...',
  required = false,
}: {
  options: (string | SelectOption)[];
  value: string;
  onChange: (val: string) => void;
  placeholder?: string;
  required?: boolean;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  const normalizedOptions = useMemo<SelectOption[]>(() => {
    if (!options) return [];
    return options.map(opt => typeof opt === 'string' ? { value: opt, label: opt } : opt);
  }, [options]);

  const selectedOption = useMemo(() => {
    return normalizedOptions.find(opt => opt.value === value);
  }, [normalizedOptions, value]);

  useEffect(() => {
    if (!isOpen) {
      setSearch(selectedOption ? selectedOption.label : '');
    }
  }, [selectedOption, isOpen]);

  const filtered = useMemo(() => {
    if (!search || !isOpen) return normalizedOptions;
    const s = search.toLowerCase();
    return normalizedOptions.filter(opt =>
      opt.label.toLowerCase().includes(s) || opt.value.toLowerCase().includes(s)
    );
  }, [normalizedOptions, search, isOpen]);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
        setIsOpen(false);
        setSearch(selectedOption ? selectedOption.label : '');
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [selectedOption]);

  return (
    <div ref={wrapperRef} className="relative w-full">
      <div className="relative">
        <input
          type="text"
          value={isOpen ? search : (selectedOption ? selectedOption.label : '')}
          onChange={(e) => {
            setSearch(e.target.value);
            if (!isOpen) setIsOpen(true);
          }}
          onFocus={() => {
            setIsOpen(true);
            setSearch('');
          }}
          placeholder={placeholder}
          required={required && !value}
          className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-blue-500 pr-8 text-xs cursor-text"
        />
        <div className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
          <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-200 ${isOpen ? 'rotate-180 text-blue-400' : ''}`} />
        </div>
      </div>

      {isOpen && (
        <div className="absolute left-0 right-0 top-full mt-1 z-50 bg-slate-900 border border-slate-700 rounded-lg shadow-2xl max-h-56 overflow-y-auto divide-y divide-slate-800 animate-fadeIn text-xs">
          {filtered.length > 0 ? (
            filtered.map((opt, i) => (
              <div
                key={i}
                onClick={() => {
                  onChange(opt.value);
                  setSearch(opt.label);
                  setIsOpen(false);
                }}
                className={`px-3 py-2 cursor-pointer hover:bg-blue-600 hover:text-white transition flex items-center justify-between ${opt.value === value ? 'bg-slate-800 text-blue-400 font-semibold' : 'text-slate-200'}`}
              >
                <span className="truncate">{opt.label}</span>
                {opt.value === value && <Check className="w-3.5 h-3.5 shrink-0 ml-2 text-blue-400" />}
              </div>
            ))
          ) : (
            <div className="px-3 py-3 text-center text-slate-500 italic">
              Tidak ditemukan "{search}"
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

const indoMonthsMap: Record<string, string> = {
  jan: '01', januari: '01', january: '01',
  feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03',
  apr: '04', april: '04',
  mei: '05', may: '05',
  jun: '06', juni: '06', june: '06',
  jul: '07', juli: '07', july: '07',
  agu: '08', ags: '08', agustus: '08', aug: '08', august: '08',
  sep: '09', september: '09',
  okt: '10', oktober: '10', oct: '10', october: '10',
  nov: '11', november: '11',
  des: '12', desember: '12', dec: '12', december: '12'
};

function parseToStandardDate(val: any): string {
  if (!val) return '';
  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'nan') return '';

  str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();

  if (str.toLowerCase().includes('t') && str.length >= 10 && /^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.slice(0, 10);
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

  const mJs = str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (mJs) {
    const mon = indoMonthsMap[mJs[1].toLowerCase()] || indoMonthsMap[mJs[1].toLowerCase().slice(0, 3)] || '01';
    return `${mJs[3]}-${mon}-${mJs[2].padStart(2, '0')}`;
  }

  const mDdMon = str.match(/^(\d{1,2})[-/\s]+([A-Za-z]+)[-/\s]+(\d{2,4})/);
  if (mDdMon) {
    let y = mDdMon[3];
    if (y.length === 2) y = '20' + y;
    const mon = indoMonthsMap[mDdMon[2].toLowerCase()] || indoMonthsMap[mDdMon[2].toLowerCase().slice(0, 3)] || '01';
    return `${y}-${mon}-${mDdMon[1].padStart(2, '0')}`;
  }

  const mYMon = str.match(/^(\d{4})[-/\s]+([A-Za-z]+)[-/\s]+(\d{1,2})/);
  if (mYMon) {
    const mon = indoMonthsMap[mYMon[2].toLowerCase()] || indoMonthsMap[mYMon[2].toLowerCase().slice(0, 3)] || '01';
    return `${mYMon[1]}-${mon}-${mYMon[3].padStart(2, '0')}`;
  }

  const mDmy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mDmy) {
    return `${mDmy[3]}-${mDmy[2].padStart(2, '0')}-${mDmy[1].padStart(2, '0')}`;
  }

  const mYmd = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (mYmd) {
    return `${mYmd[1]}-${mYmd[2].padStart(2, '0')}-${mYmd[3].padStart(2, '0')}`;
  }

  const num = Number(str);
  if (!isNaN(num) && num > 30000 && num < 60000) {
    const d = new Date(Math.round((num - 25569) * 86400 * 1000));
    return d.toISOString().split('T')[0];
  }

  return str;
}

function formatDisplayDate(val: any): string {
  if (!val) return '-';
  const iso = parseToStandardDate(val);
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return `${m[3]}-${m[2]}-${m[1]}`;
  }
  const clean = String(val).replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  return clean || '-';
}

function getLocalRev(id: string, defaultVal: string = '0'): string {
  try {
    const map = JSON.parse(localStorage.getItem('task_rev_map') || '{}');
    return map[id] !== undefined ? String(map[id]) : defaultVal;
  } catch {
    return defaultVal;
  }
}

function saveLocalRev(id: string, val: string) {
  try {
    const map = JSON.parse(localStorage.getItem('task_rev_map') || '{}');
    map[id] = val;
    localStorage.setItem('task_rev_map', JSON.stringify(map));
  } catch {}
}

function getLocalRelease(id: string, defaultVal: string = ''): string {
  try {
    const map = JSON.parse(localStorage.getItem('task_release_map') || '{}');
    return map[id] !== undefined ? String(map[id]) : defaultVal;
  } catch {
    return defaultVal;
  }
}

function saveLocalRelease(id: string, val: string) {
  try {
    const map = JSON.parse(localStorage.getItem('task_release_map') || '{}');
    map[id] = val;
    localStorage.setItem('task_release_map', JSON.stringify(map));
  } catch {}
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
    } catch {}
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
    release: '',
  });

  const [editingTask, setEditingTask] = useState<TaskItem | null>(null);
  const [editFormData, setEditFormData] = useState({
    project: '',
    taskName: '',
    startDate: '',
    endDate: '',
    jo: '',
    kodeJc: '',
    rev: '0',
    release: '',
  });

  const [editingReleaseId, setEditingReleaseId] = useState<string | null>(null);
  const [editingReleaseVal, setEditingReleaseVal] = useState<string>('');

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

          const persistedRev = (row.rev !== undefined && row.rev !== null && String(row.rev).trim() !== '')
            ? String(row.rev)
            : getLocalRev(row.id, '0');

          const persistedRelease = (row.release !== undefined && row.release !== null && String(row.release).trim() !== '')
            ? String(row.release)
            : getLocalRelease(row.id, '');

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
            rev: persistedRev,
            release: persistedRelease,
            realJo: row.real_jo || row.realJo || '',
          });
        });
        setManualTasks(grouped);
      }
    } catch {}
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

  // State File Excel Utama
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);

  const [drawingControlMap, setDrawingControlMap] = useState<Record<string, DrawingControlRow[]>>({});
  const [activeDrawingSheetTitle, setActiveDrawingSheetTitle] = useState<string>('');
  const [isFetchingDrawing, setIsFetchingDrawing] = useState<boolean>(false);

  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        let jcUrl = '';
        let im4Url = '';
        let realisasiUrl = '';

        Object.entries(excelGlobUrls).forEach(([path, url]) => {
          const pLower = path.toLowerCase();
          if (pLower.includes('jobcard') && !pLower.includes('realisasi')) jcUrl = url;
          else if (pLower.includes('im4') || pLower.includes('drawing') || pLower.includes('akses')) im4Url = url;
          else if (pLower.includes('realisasi')) realisasiUrl = url;
        });

        const [wbJc, wbIm4, wbRealisasi] = await Promise.all([
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

        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
        if (wbRealisasi) setRealisasiWorkbook(wbRealisasi);
      } catch {}
    }
    loadAllExcelFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

  const handleUpdateIm4Excel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        setIm4Workbook(wb);
        alert(`Master Personel "${file.name}" berhasil diperbarui!`);
      } catch {
        alert('Gagal membaca file Excel Personel IM4.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleUpdateJobcardExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        setJobcardWorkbook(wb);
        alert(`Katalog Jobcard "${file.name}" berhasil diperbarui!`);
      } catch {
        alert('Gagal membaca file JOBCARD_DESAIN.xlsx.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const fetchDrawingControlForProject = useCallback(async (projectCode: string) => {
    const cleanProj = cleanText(projectCode);
    if (!cleanProj) return;

    setIsFetchingDrawing(true);

    try {
      const url = `${GAS_DRAWING_API_URL}?project=${encodeURIComponent(projectCode.trim())}`;
      const res = await fetch(url);
      
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data) && json.data.length > 0) {
          const rowsList: DrawingControlRow[] = json.data.map((item: any) => ({
            noDwg: String(item.noDwg || '').trim(),
            drawingName: String(item.drawingName || '').trim(),
            fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
            rev: String(item.rev || '0').trim(),
            finishDate: String(item.finishDate || '').trim(),
          }));

          setDrawingControlMap(prev => ({ ...prev, [cleanProj]: rowsList }));
          setActiveDrawingSheetTitle(json.fileName || `Drawing Control ${projectCode}`);
          setIsFetchingDrawing(false);
          return;
        }
      }
    } catch (err) {
      console.warn('Gagal koneksi ke Google Apps Script, mencoba fallback...', err);
    }

    const match = GOOGLE_DRIVE_SHEETS.find(s => 
      cleanText(s.projectKey) === cleanProj || 
      cleanProj.includes(cleanText(s.projectKey)) || 
      cleanText(s.title).includes(cleanProj)
    );

    if (match) {
      try {
        const csvUrl = `https://docs.google.com/spreadsheets/d/${match.id}/gviz/tq?tqx=out:csv&sheet=Drawing%20Control%20(2)`;
        const res = await fetch(csvUrl);
        if (res.ok) {
          const csvText = await res.text();
          const wb = XLSX.read(csvText, { type: 'string' });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          
          const rowsList: DrawingControlRow[] = [];
          for (let r = 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row) continue;
            const noDwg = String(row[0] || '').trim();
            const dwgName = String(row[1] || '').trim();
            const rev = String(row[14] || '0').trim();
            const finishRaw = String(row[15] || '').trim();

            if (!noDwg && !dwgName) continue;
            const fullDeskripsi = (noDwg && dwgName) ? `${noDwg}-${dwgName}` : (dwgName || noDwg);
            rowsList.push({ noDwg, drawingName: dwgName, fullDeskripsi, rev, finishDate: finishRaw });
          }

          if (rowsList.length > 0) {
            setDrawingControlMap(prev => ({ ...prev, [cleanProj]: rowsList }));
            setActiveDrawingSheetTitle(match.title);
          }
        }
      } catch {}
    }

    setIsFetchingDrawing(false);
  }, []);

  const handleProjectChange = (newProject: string) => {
    setFormData(prev => ({ ...prev, kodeProyek: newProject, taskName: '', release: '' }));
    fetchDrawingControlForProject(newProject);
  };

  const handleDeskripsiChange = (selectedDesc: string) => {
    const cleanProj = cleanText(formData.kodeProyek || '');
    const rowsForThisProj = drawingControlMap[cleanProj] || [];

    let autoRelease = '';
    let autoRev = formData.rev || '0';

    if (rowsForThisProj.length > 0 && selectedDesc) {
      const cleanTarget = cleanText(selectedDesc);

      const matches = rowsForThisProj.filter(r => {
        const cFull = cleanText(r.fullDeskripsi);
        const cDwg = cleanText(r.noDwg);
        const cName = cleanText(r.drawingName);

        if (cFull === cleanTarget) return true;
        if (cName && (cleanTarget === cName || cleanTarget.includes(cName))) return true;
        if (cDwg && cleanTarget.includes(cDwg)) return true;
        return false;
      });

      if (matches.length > 0) {
        const lastRow = matches[matches.length - 1];
        if (lastRow.finishDate) {
          autoRelease = parseToStandardDate(lastRow.finishDate);
        }
        if (lastRow.rev) {
          autoRev = lastRow.rev;
        }
      }
    }

    setFormData(prev => ({
      ...prev,
      taskName: selectedDesc,
      release: autoRelease || prev.release,
      rev: autoRev,
    }));
  };

  const handleQuickSaveRelease = async (taskId: string, newVal: string) => {
    const cleanDate = parseToStandardDate(newVal);

    saveLocalRelease(taskId, cleanDate);

    try {
      await supabase
        .from('job_cards')
        .update({ release: cleanDate })
        .eq('id', taskId);
    } catch {}

    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => {
        updated[k] = updated[k].map(t => t.id === taskId ? { ...t, release: cleanDate } : t);
      });
      return updated;
    });

    setEditingReleaseId(null);
  };

  const handleSyncReleaseFromDrive = async (task: TaskItem) => {
    const cleanProj = cleanText(task.project);
    let rows = drawingControlMap[cleanProj];

    if (!rows || rows.length === 0) {
      try {
        const url = `${GAS_DRAWING_API_URL}?project=${encodeURIComponent(task.project.trim())}`;
        const res = await fetch(url);
        if (res.ok) {
          const json = await res.json();
          if (json.success && Array.isArray(json.data)) {
            rows = json.data.map((item: any) => ({
              noDwg: String(item.noDwg || '').trim(),
              drawingName: String(item.drawingName || '').trim(),
              fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
              rev: String(item.rev || '0').trim(),
              finishDate: String(item.finishDate || '').trim(),
            }));
            setDrawingControlMap(prev => ({ ...prev, [cleanProj]: rows }));
          }
        }
      } catch (err) {
        console.error(err);
      }
    }

    if (rows && rows.length > 0) {
      const cleanTarget = cleanText(task.taskName);
      const matches = rows.filter(r => {
        const cFull = cleanText(r.fullDeskripsi);
        const cDwg = cleanText(r.noDwg);
        const cName = cleanText(r.drawingName);
        return cFull === cleanTarget || (cName && cleanTarget.includes(cName)) || (cDwg && cleanTarget.includes(cDwg));
      });

      if (matches.length > 0) {
        const lastRow = matches[matches.length - 1];
        if (lastRow.finishDate) {
          const standardDate = parseToStandardDate(lastRow.finishDate);
          await handleQuickSaveRelease(task.id, standardDate);
          alert(`Berhasil sinkronisasi dari Google Drive: ${formatDisplayDate(standardDate)}`);
          return;
        }
      }
    }
    alert(`Tidak ditemukan data FINISH DATE di Google Drive untuk gambar "${task.taskName}". Silakan edit langsung dengan tombol pensil.`);
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

  const getBiroMembers = useCallback((biroName: string): { nama: string; status: string; jabatan: string }[] => {
    const members = allParsedFromExcel.filter(
      p => isBiroMatch(p.biro, biroName) && !p.status.toLowerCase().includes('outsourcing')
    );
    if (members.length > 0) {
      return members.map(m => ({ nama: m.nama, status: m.status, jabatan: m.jabatan }));
    }
    return [];
  }, [allParsedFromExcel]);

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

  const subconWorkOrders = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const prefix = getBiroPrefix(subconSelectedBiro);
    return currentActiveBiroTasks.map((task, idx) => ({
      ...task,
      packageTitle: `${prefix}${idx + 1}`
    }));
  }, [subconSelectedBiro, currentActiveBiroTasks]);

  const projectOptions = useMemo((): string[] => {
    const projectMap = new Map<string, string>();

    GOOGLE_DRIVE_SHEETS.forEach(s => {
      const cleaned = cleanProjectString(s.projectKey);
      const normKey = getProjectNormKey(cleaned);
      if (cleaned && normKey && !projectMap.has(normKey)) {
        projectMap.set(normKey, cleaned);
      }
    });

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

  const dynamicTaskOptions = useMemo((): string[] => {
    const cleanProj = cleanText(formData.kodeProyek || '');
    const rowsForThisProj = drawingControlMap[cleanProj] || [];
    const taskMap = new Map<string, string>();

    rowsForThisProj.forEach(r => {
      if (r.fullDeskripsi) {
        const norm = cleanText(r.fullDeskripsi);
        if (!taskMap.has(norm)) {
          taskMap.set(norm, r.fullDeskripsi);
        }
      }
    });

    if (jobcardWorkbook) {
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
            raw.toLowerCase() === 'deskripsi'
          ) return;

          const cleaned = raw.replace(/[\.,;:\-_/\\\s]+$/, '').trim();
          const normKey = cleanText(cleaned);

          if (cleaned && normKey && !taskMap.has(normKey)) {
            taskMap.set(normKey, cleaned);
          }
        });
      });
    }

    return Array.from(taskMap.values()).sort((a, b) => a.localeCompare(b));
  }, [formData.kodeProyek, drawingControlMap, jobcardWorkbook]);

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

      let autoKode = '';
      if (accessMode === 'subkon') {
        const prefix = getBiroPrefix(activeBiro);
        autoKode = `${prefix}${currentList.length + 1}`;
      }

      const revVal = (formData.rev && formData.rev.trim() !== '') ? formData.rev.trim() : '0';
      const releaseVal = formData.release ? formData.release.trim() : '';

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
        rev: revVal,
        release: releaseVal,
      };

      let insertedRow: any = null;
      const { data: resData, error } = await supabase
        .from('job_cards')
        .insert(insertPayload)
        .select()
        .single();

      if (error) {
        delete insertPayload.rev;
        delete insertPayload.release;
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
        insertedRow = resData;
      }

      saveLocalRev(insertedRow.id, revVal);
      saveLocalRelease(insertedRow.id, releaseVal);

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
        rev: revVal,
        release: releaseVal,
      };

      setManualTasks({ ...manualTasks, [biroKey]: [...currentList, newTask] });

      setExpandedCards(prev => ({ ...prev, [formData.nama]: true }));
      setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', pic: '', jo: '', rev: '0', realJo: '', release: '' });

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

  const handleOpenEdit = (task: TaskItem) => {
    setEditingTask(task);
    setEditFormData({
      project: task.project || '',
      taskName: task.taskName || '',
      startDate: task.startDate || '',
      endDate: task.endDate || '',
      jo: task.jo || '',
      kodeJc: task.kodeJc || '',
      rev: task.rev || '0',
      release: task.release || '',
    });
  };

  const handleSaveEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingTask) return;

    const revVal = (editFormData.rev && editFormData.rev.trim() !== '') ? editFormData.rev.trim() : '0';
    const releaseVal = editFormData.release ? editFormData.release.trim() : '';

    const updatePayload: any = {
      project: editFormData.project,
      project_code: editFormData.project,
      task_name: editFormData.taskName,
      start_date: editFormData.startDate,
      end_date: editFormData.endDate,
      jo: editFormData.jo,
      kode_jc: editFormData.kodeJc,
      rev: revVal,
      release: releaseVal,
    };

    let { error } = await supabase
      .from('job_cards')
      .update(updatePayload)
      .eq('id', editingTask.id);

    if (error && (error.message?.includes('rev') || error.message?.includes('release') || (error as any).details?.includes('rev') || (error as any).details?.includes('release'))) {
      delete updatePayload.rev;
      delete updatePayload.release;
      const retry = await supabase.from('job_cards').update(updatePayload).eq('id', editingTask.id);
      error = retry.error;
    }

    if (error) {
      alert('Gagal mengupdate: ' + error.message);
      return;
    }

    saveLocalRev(editingTask.id, revVal);
    saveLocalRelease(editingTask.id, releaseVal);

    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => {
        updated[k] = updated[k].map(t => {
          if (t.id === editingTask.id) {
            return {
              ...t,
              project: editFormData.project,
              taskName: editFormData.taskName,
              startDate: editFormData.startDate,
              endDate: editFormData.endDate,
              jo: editFormData.jo,
              kodeJc: editFormData.kodeJc,
              rev: revVal,
              release: releaseVal,
            };
          }
          return t;
        });
      });
      return updated;
    });

    setEditingTask(null);
    alert('Penugasan berhasil diperbarui!');
    loadAllJobCards();
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
            {!subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
                  <div>
                    <h2 className="text-lg font-bold text-white">Pilih Departemen</h2>
                    <span className="text-xs text-slate-400">Total {dynamicOutsourcingList.length} Personel Outsourcing</span>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Users className="w-3.5 h-3.5 text-amber-400" />
                      <span>Update Personel IM4</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleUpdateIm4Excel} className="hidden" />
                    </label>

                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" />
                      <span>Update Jobcard Desain</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleUpdateJobcardExcel} className="hidden" />
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
                                          <th className="py-2 px-2.5 text-center w-20">Aksi</th>
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
                                                {formatDisplayDate(task.startDate)} s/d {formatDisplayDate(task.endDate)}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-violet-300">
                                                {task.jo ? String(task.jo).replace(/^#+/, '') : '-'}
                                              </td>
                                              <td className="py-2 px-2.5 text-center">
                                                <div className="flex items-center justify-center gap-1.5">
                                                  <button 
                                                    onClick={(e) => {
                                                      e.stopPropagation();
                                                      handleOpenEdit(task);
                                                    }} 
                                                    className="p-1 rounded bg-slate-850 hover:bg-amber-500/20 text-slate-400 hover:text-amber-300 transition cursor-pointer"
                                                    title="Edit Tugas"
                                                  >
                                                    <Pencil className="w-3.5 h-3.5" />
                                                  </button>
                                                  <button 
                                                    onClick={(e) => {
                                                      e.stopPropagation();
                                                      handleDeleteTask(task.id);
                                                    }} 
                                                    className="p-1 rounded bg-slate-850 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 transition cursor-pointer"
                                                    title="Hapus Tugas"
                                                  >
                                                    <Trash2 className="w-3.5 h-3.5" />
                                                  </button>
                                                </div>
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
                        Tidak ada anggota di biro ini (Silakan upload Master IM4 terlebih dahulu)
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
                          <SearchableSelect
                            options={activeSubconMembers.map(p => ({ value: p.nama, label: `${p.nama} (${p.jabatan})` }))}
                            value={formData.nama}
                            onChange={(val) => setFormData(prev => ({ ...prev, nama: val }))}
                            placeholder="Ketik nama drafter / personel..."
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Kode Proyek</label>
                          <SearchableSelect
                            options={projectOptions}
                            value={formData.kodeProyek}
                            onChange={(val) => handleProjectChange(val)}
                            placeholder="Ketik kode proyek..."
                            required
                          />
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
                          <SearchableSelect
                            options={dynamicTaskOptions}
                            value={formData.taskName}
                            onChange={(val) => handleDeskripsiChange(val)}
                            placeholder="Ketik nama gambar / deskripsi tugas..."
                            required
                          />
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
                              <div className="text-emerald-400 font-mono">{pkg.project} • {pkg.jo}</div>
                              <div className="text-slate-300 text-[11px] bg-slate-950 p-2 rounded border border-slate-800">{pkg.taskName}</div>
                            </div>
                            <div className="text-[10px] text-slate-500 pt-2 border-t border-slate-800 flex justify-between font-mono">
                              <span>{formatDisplayDate(pkg.startDate)} s/d {formatDisplayDate(pkg.endDate)}</span>
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
            {!selectedDept && !selectedBiroPage && !selectedFormBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
                  <div>
                    <h2 className="text-lg font-bold text-white">Departemen Desain</h2>
                    <span className="text-xs text-slate-400">Akses Pegawai Organik (PKWTT & PKWT)</span>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    {/* Tombol Update File Master Personel IM4 */}
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Users className="w-3.5 h-3.5 text-blue-400" />
                      <span>Update Personel IM4</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleUpdateIm4Excel} className="hidden" />
                    </label>

                    {/* Tombol Update File Master JOBCARD_DESAIN */}
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" />
                      <span>Update Jobcard Desain</span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleUpdateJobcardExcel} className="hidden" />
                    </label>

                    {/* Tombol Update Realisasi JO.xlsx */}
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

            {selectedFormBiro && (
              <div className="space-y-4">
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-white text-sm">{selectedFormBiro.biroName}</h3>
                    <div className="text-xs text-slate-400 font-mono flex items-center gap-2 mt-0.5">
                      <span>{currentBiroMembers.length} Pegawai Organik</span>
                      {activeDrawingSheetTitle && (
                        <span className="text-emerald-400 flex items-center gap-1">
                          • <Sparkles className="w-3 h-3 text-amber-400" /> {activeDrawingSheetTitle}
                        </span>
                      )}
                    </div>
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

                {/* TAB 1: ANGGOTA ORGANIK */}
                {formPageMode === 'members' ? (
                  <div className="space-y-2.5">
                    {currentBiroMembers.length > 0 ? (
                      currentBiroMembers.map((person, idx) => {
                        const personTasks = currentActiveBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
                        const isExpanded = !!expandedCards[person.nama];

                        return (
                          <div key={idx} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden transition-all duration-200">
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
                                          <th className="py-2 px-2.5 font-mono text-cyan-300">Release</th>
                                          <th className="py-2 px-2.5 text-center w-20">Aksi</th>
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
                                                {formatDisplayDate(task.startDate)}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-[11px] text-slate-300">
                                                {formatDisplayDate(task.endDate)}
                                              </td>
                                              <td className="py-2 px-2.5 font-mono text-violet-300">
                                                {task.jo ? String(task.jo).replace(/^#+/, '') : '-'}
                                              </td>
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
                                              <td className="py-2 px-2.5 font-mono text-[11px]">
                                                {editingReleaseId === task.id ? (
                                                  <div className="flex items-center gap-1">
                                                    <input
                                                      type="date"
                                                      value={editingReleaseVal}
                                                      onChange={(e) => setEditingReleaseVal(e.target.value)}
                                                      style={{ colorScheme: 'dark' }}
                                                      className="px-1.5 py-0.5 bg-slate-900 border border-blue-500 rounded text-xs text-white"
                                                      autoFocus
                                                    />
                                                    <button
                                                      onClick={() => handleQuickSaveRelease(task.id, editingReleaseVal)}
                                                      className="p-1 text-emerald-400 hover:text-emerald-300 hover:bg-emerald-500/20 rounded cursor-pointer"
                                                      title="Simpan"
                                                    >
                                                      <Check className="w-3 h-3" />
                                                    </button>
                                                    <button
                                                      onClick={() => setEditingReleaseId(null)}
                                                      className="p-1 text-slate-400 hover:text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer"
                                                      title="Batal"
                                                    >
                                                      <X className="w-3 h-3" />
                                                    </button>
                                                  </div>
                                                ) : (
                                                  <div className="flex items-center gap-1.5 group">
                                                    <span className="text-cyan-300 font-semibold">
                                                      {formatDisplayDate(task.release)}
                                                    </span>
                                                    <div className="flex items-center gap-1 opacity-70 group-hover:opacity-100 transition">
                                                      <button
                                                        onClick={(e) => {
                                                          e.stopPropagation();
                                                          setEditingReleaseId(task.id);
                                                          setEditingReleaseVal(parseToStandardDate(task.release) || '');
                                                        }}
                                                        className="p-0.5 rounded text-slate-400 hover:text-cyan-300 hover:bg-slate-800 transition cursor-pointer"
                                                        title="Edit Langsung Tanggal Release"
                                                      >
                                                        <Pencil className="w-3 h-3" />
                                                      </button>

                                                      <button
                                                        onClick={(e) => {
                                                          e.stopPropagation();
                                                          handleSyncReleaseFromDrive(task);
                                                        }}
                                                        className="p-0.5 rounded text-slate-400 hover:text-amber-400 hover:bg-slate-800 transition cursor-pointer"
                                                        title="Tarik Otomatis dari Link Google Drive"
                                                      >
                                                        <Sparkles className="w-3 h-3" />
                                                      </button>
                                                    </div>
                                                  </div>
                                                )}
                                              </td>

                                              <td className="py-2 px-2.5 text-center">
                                                <div className="flex items-center justify-center gap-1.5">
                                                  <button 
                                                    onClick={(e) => {
                                                      e.stopPropagation();
                                                      handleOpenEdit(task);
                                                    }} 
                                                    className="p-1 rounded bg-slate-850 hover:bg-blue-500/20 text-slate-400 hover:text-blue-400 transition cursor-pointer"
                                                    title="Edit Tugas"
                                                  >
                                                    <Pencil className="w-3.5 h-3.5" />
                                                  </button>
                                                  <button 
                                                    onClick={(e) => {
                                                      e.stopPropagation();
                                                      handleDeleteTask(task.id);
                                                    }} 
                                                    className="p-1 rounded bg-slate-850 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 transition cursor-pointer"
                                                    title="Hapus Tugas"
                                                  >
                                                    <Trash2 className="w-3.5 h-3.5" />
                                                  </button>
                                                </div>
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
                        Tidak ada personil organik di biro ini (Silakan upload Master IM4 terlebih dahulu)
                      </div>
                    )}
                  </div>
                ) : (
                  /* TAB 2: FORM ORGANIK */
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-slate-400 mb-1">Nama Personel (Organik)</label>
                          <SearchableSelect
                            options={currentBiroMembers.map(p => ({ value: p.nama, label: `${p.nama} (${p.status})` }))}
                            value={formData.nama}
                            onChange={(val) => setFormData(prev => ({ ...prev, nama: val }))}
                            placeholder="Ketik nama personel organik..."
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Kode Proyek</label>
                          <SearchableSelect
                            options={projectOptions}
                            value={formData.kodeProyek}
                            onChange={(val) => handleProjectChange(val)}
                            placeholder="Ketik kode proyek (contoh: M000313)..."
                            required
                          />
                        </div>

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

                        <div className="md:col-span-2">
                          <div className="flex items-center justify-between mb-1">
                            <label className="text-slate-400">Deskripsi</label>
                            {isFetchingDrawing && (
                              <span className="text-[10px] text-amber-400 animate-pulse font-mono">
                                Mengambil dari Google Drive...
                              </span>
                            )}
                          </div>
                          <SearchableSelect
                            options={dynamicTaskOptions}
                            value={formData.taskName}
                            onChange={(val) => handleDeskripsiChange(val)}
                            placeholder="Ketik no dwg atau nama gambar (misal: Shaft Protection)..."
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-slate-400 mb-1">Rev (Revisi)</label>
                          <input
                            type="text"
                            value={formData.rev}
                            onChange={(e) => setFormData(prev => ({ ...prev, rev: e.target.value }))}
                            placeholder="0"
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                          />
                        </div>

                        <div>
                          <div className="flex items-center justify-between mb-1">
                            <label className="text-slate-400">Release (Tanggal-Bulan-Tahun)</label>
                            <div className="flex items-center gap-1.5">
                              {formData.release && (
                                <span className="text-[10px] text-cyan-300 font-mono font-semibold">
                                  {formatDisplayDate(formData.release)}
                                </span>
                              )}
                              <button
                                type="button"
                                onClick={() => handleDeskripsiChange(formData.taskName)}
                                className="px-1.5 py-0.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-400 border border-amber-500/30 rounded text-[10px] flex items-center gap-1 cursor-pointer"
                                title="Tarik ulang dari Link Drive"
                              >
                                <Sparkles className="w-2.5 h-2.5" /> Dari Drive
                              </button>
                            </div>
                          </div>
                          <input 
                            type="date" 
                            value={formData.release} 
                            onChange={(e) => setFormData(prev => ({ ...prev, release: e.target.value }))} 
                            style={{ colorScheme: 'dark' }}
                            className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" 
                          />
                        </div>

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

        {/* ================= MODAL EDIT PENUGASAN ================= */}
        {editingTask && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg overflow-hidden shadow-2xl animate-fadeIn">
              <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Pencil className="w-3.5 h-3.5 text-blue-400" /> Edit Penugasan — {editingTask.pic}
                </span>
                <button onClick={() => setEditingTask(null)} className="text-slate-400 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>

              <form onSubmit={handleSaveEdit} className="p-4 space-y-3.5 text-xs">
                <div>
                  <label className="block text-slate-400 mb-1">Kode Proyek</label>
                  <SearchableSelect
                    options={projectOptions}
                    value={editFormData.project}
                    onChange={(val) => {
                      setEditFormData(prev => ({ ...prev, project: val }));
                      fetchDrawingControlForProject(val);
                    }}
                    placeholder="Ketik kode proyek..."
                    required
                  />
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-400 mb-1">
                      {accessMode === 'organik' ? 'Plan JO' : 'Nomor JO'}
                    </label>
                    <input
                      type="text"
                      inputMode="numeric"
                      value={editFormData.jo}
                      onChange={(e) => setEditFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9]/g, '') }))}
                      required
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-400 mb-1">
                      {accessMode === 'organik' ? 'Jobcard' : 'Kode WO'}
                    </label>
                    <input
                      type="text"
                      value={editFormData.kodeJc}
                      onChange={(e) => setEditFormData(prev => ({ ...prev, kodeJc: e.target.value.toUpperCase() }))}
                      placeholder={accessMode === 'organik' ? 'JC...' : 'WO...'}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono uppercase"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Deskripsi</label>
                  <SearchableSelect
                    options={dynamicTaskOptions}
                    value={editFormData.taskName}
                    onChange={(val) => {
                      const cleanProj = cleanText(editFormData.project || '');
                      const rows = drawingControlMap[cleanProj] || [];
                      let newRelease = editFormData.release;
                      let newRev = editFormData.rev;

                      const cleanTarget = cleanText(val);
                      const matches = rows.filter(r => cleanText(r.fullDeskripsi) === cleanTarget || cleanText(r.drawingName) === cleanTarget || (r.noDwg && cleanTarget.includes(cleanText(r.noDwg))));
                      if (matches.length > 0) {
                        const last = matches[matches.length - 1];
                        if (last.finishDate) newRelease = parseToStandardDate(last.finishDate);
                        if (last.rev) newRev = last.rev;
                      }

                      setEditFormData(prev => ({ ...prev, taskName: val, release: newRelease, rev: newRev }));
                    }}
                    placeholder="Ketik deskripsi gambar..."
                    required
                  />
                </div>

                {accessMode === 'organik' && (
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-400 mb-1">Rev (Revisi)</label>
                      <input
                        type="text"
                        value={editFormData.rev}
                        onChange={(e) => setEditFormData(prev => ({ ...prev, rev: e.target.value }))}
                        placeholder="0"
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Release</label>
                      <input
                        type="date"
                        value={editFormData.release}
                        onChange={(e) => setEditFormData(prev => ({ ...prev, release: e.target.value }))}
                        style={{ colorScheme: 'dark' }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer"
                      />
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-400 mb-1">
                      {accessMode === 'organik' ? 'Plan Start' : 'Tanggal Mulai'}
                    </label>
                    <input
                      type="date"
                      value={editFormData.startDate}
                      onChange={(e) => setEditFormData(prev => ({ ...prev, startDate: e.target.value }))}
                      required
                      style={{ colorScheme: 'dark' }}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-400 mb-1">
                      {accessMode === 'organik' ? 'Plan Finish' : 'Tanggal Selesai'}
                    </label>
                    <input
                      type="date"
                      value={editFormData.endDate}
                      onChange={(e) => setEditFormData(prev => ({ ...prev, endDate: e.target.value }))}
                      required
                      style={{ colorScheme: 'dark' }}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer"
                    />
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setEditingTask(null)}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold rounded-lg cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg cursor-pointer"
                  >
                    Simpan Perubahan
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* ================= MODAL PLANNER ================= */}
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