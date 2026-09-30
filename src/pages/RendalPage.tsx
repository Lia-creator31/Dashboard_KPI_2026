import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  Building2, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, Lock, X, Check, Printer, FileCheck, Clock, Sparkles, 
  FileSpreadsheet, Users, Search, KeyRound, Upload, Briefcase, HardHat
} from 'lucide-react';

// --- Konstanta Global ---
const excelGlobUrls = import.meta.glob('./*.xlsx', {
  query: '?url',
  import: 'default',
  eager: true
}) as Record<string, string>;

const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';

const GOOGLE_DRIVE_SHEETS = [
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

// --- Interfaces ---
export interface ParsedMember {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
  biro: string;
  dept: string;
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

interface SelectOption {
  value: string;
  label: string;
}

interface RendalPageProps {
  user: UserSession;
  onLogout: () => void;
}

// --- Helper Functions (LENGKAP dari App.tsx) ---
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
  if (str.toLowerCase().includes('t') && str.length >= 10 && /^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const mJs = str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (mJs) {
    const mon = indoMonthsMap[mJs[1].toLowerCase()] || indoMonthsMap[mJs[1].toLowerCase().slice(0, 3)] || '01';
    return `${mJs[3]}-${mon}-${mJs[2].padStart(2, '0')}`;
  }
  const mDdMon = str.match(/^(\d{1,2})[-/\s]+([A-Za-z]+)[-/\s]+(\d{2,4})/);
  if (mDdMon) {
    let y = mDdMon[3]; if (y.length === 2) y = '20' + y;
    const mon = indoMonthsMap[mDdMon[2].toLowerCase()] || indoMonthsMap[mDdMon[2].toLowerCase().slice(0, 3)] || '01';
    return `${y}-${mon}-${mDdMon[1].padStart(2, '0')}`;
  }
  const mYMon = str.match(/^(\d{4})[-/\s]+([A-Za-z]+)[-/\s]+(\d{1,2})/);
  if (mYMon) {
    const mon = indoMonthsMap[mYMon[2].toLowerCase()] || indoMonthsMap[mYMon[2].toLowerCase().slice(0, 3)] || '01';
    return `${mYMon[1]}-${mon}-${mYMon[3].padStart(2, '0')}`;
  }
  const mDmy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mDmy) return `${mDmy[3]}-${mDmy[2].padStart(2, '0')}-${mDmy[1].padStart(2, '0')}`;
  const mYmd = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (mYmd) return `${mYmd[1]}-${mYmd[2].padStart(2, '0')}-${mYmd[3].padStart(2, '0')}`;
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
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  const clean = String(val).replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  return clean || '-';
}

function getLocalRev(id: string, defaultVal: string = '0'): string {
  try {
    const map = JSON.parse(localStorage.getItem('task_rev_map') || '{}');
    return map[id] !== undefined ? String(map[id]) : defaultVal;
  } catch { return defaultVal; }
}

function getLocalRelease(id: string, defaultVal: string = ''): string {
  try {
    const map = JSON.parse(localStorage.getItem('task_release_map') || '{}');
    return map[id] !== undefined ? String(map[id]) : defaultVal;
  } catch { return defaultVal; }
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
  if (c1 && c2) return c1 === c2 || c1.includes(c2) || c2.includes(c1);
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
    } catch {}
  }
  return null;
}

// --- Komponen SearchableSelect ---
function SearchableSelect({ options, value, onChange, placeholder = 'Ketik untuk mencari...', required = false }: { options: (string | SelectOption)[]; value: string; onChange: (val: string) => void; placeholder?: string; required?: boolean; }) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);
  const normalizedOptions = useMemo<SelectOption[]>(() => {
    if (!options) return [];
    return options.map(opt => typeof opt === 'string' ? { value: opt, label: opt } : opt);
  }, [options]);
  const selectedOption = useMemo(() => normalizedOptions.find(opt => opt.value === value), [normalizedOptions, value]);
  useEffect(() => { if (!isOpen) setSearch(selectedOption ? selectedOption.label : ''); }, [selectedOption, isOpen]);
  const filtered = useMemo(() => {
    if (!search || !isOpen) return normalizedOptions;
    const s = search.toLowerCase();
    return normalizedOptions.filter(opt => opt.label.toLowerCase().includes(s) || opt.value.toLowerCase().includes(s));
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
        <input type="text" value={isOpen ? search : (selectedOption ? selectedOption.label : '')}
          onChange={(e) => { setSearch(e.target.value); if (!isOpen) setIsOpen(true); }}
          onFocus={() => { setIsOpen(true); setSearch(''); }}
          placeholder={placeholder} required={required && !value}
          className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-blue-500 pr-8 text-xs cursor-text" />
        <div className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
          <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-200 ${isOpen ? 'rotate-180 text-blue-400' : ''}`} />
        </div>
      </div>
      {isOpen && (
        <div className="absolute left-0 right-0 top-full mt-1 z-50 bg-slate-900 border border-slate-700 rounded-lg shadow-2xl max-h-56 overflow-y-auto divide-y divide-slate-800 animate-fadeIn text-xs">
          {filtered.length > 0 ? (
            filtered.map((opt, i) => (
              <div key={i} onClick={() => { onChange(opt.value); setSearch(opt.label); setIsOpen(false); }}
                className={`px-3 py-2 cursor-pointer hover:bg-blue-600 hover:text-white transition flex items-center justify-between ${opt.value === value ? 'bg-slate-800 text-blue-400 font-semibold' : 'text-slate-200'}`}>
                <span className="truncate">{opt.label}</span>
                {opt.value === value && <Check className="w-3.5 h-3.5 shrink-0 ml-2 text-blue-400" />}
              </div>
            ))
          ) : (
            <div className="px-3 py-3 text-center text-slate-500 italic">Tidak ditemukan "{search}"</div>
          )}
        </div>
      )}
    </div>
  );
}

// --- Komponen Utama RendalPage ---
export default function RendalPage({ user, onLogout }: RendalPageProps) {
  // State Navigasi
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroName, setSelectedBiroName] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [memberTab, setMemberTab] = useState<'organic' | 'outsourcing'>('organic');
  
  // State Data Master (Excel)
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);
  
  // State Tasks & UI
  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>({});
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});
  
  // State Planner Modal
  const [isPlannerOpen, setIsPlannerOpen] = useState(false);
  
  // State Release Editing
  const [editingReleaseId, setEditingReleaseId] = useState<string | null>(null);
  const [editingReleaseVal, setEditingReleaseVal] = useState<string>('');

  // State Drawing Control
  const [drawingControlMap, setDrawingControlMap] = useState<Record<string, DrawingControlRow[]>>({});
  const [isFetchingDrawing, setIsFetchingDrawing] = useState<boolean>(false);

  // Load All Job Cards from Supabase
  const loadAllJobCards = useCallback(async () => {
    try {
      const { data, error } = await supabase.from('job_cards').select('*').order('created_at', { ascending: true });
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
            realJo: row.real_jc || row.realJo || '',
          });
        });
        setManualTasks(grouped);
      }
    } catch {}
  }, []);

  // Initial Load Excel Files
  useEffect(() => {
    async function initMasterFiles() {
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
          fetchSafeWorkbook([jcUrl, '/JOBCARD_DESAIN.xlsx', './JOBCARD_DESAIN.xlsx']),
          fetchSafeWorkbook([im4Url, '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx', './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx']),
          fetchSafeWorkbook([realisasiUrl, '/Realisasi JO.xlsx', './Realisasi JO.xlsx'])
        ]);

        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
        if (wbRealisasi) setRealisasiWorkbook(wbRealisasi);
      } catch {}
    }
    initMasterFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

  // Handlers Upload Excel
  const handleUpdateIm4Excel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try { setIm4Workbook(XLSX.read(evt.target?.result, { type: 'binary' })); alert(`Master Personel "${file.name}" berhasil diperbarui!`); } 
      catch { alert('Gagal membaca file Excel Personel IM4.'); }
    };
    reader.readAsBinaryString(file);
  };

  const handleUpdateJobcardExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try { setJobcardWorkbook(XLSX.read(evt.target?.result, { type: 'binary' })); alert(`Katalog Jobcard "${file.name}" berhasil diperbarui!`); } 
      catch { alert('Gagal membaca file JOBCARD_DESAIN.xlsx.'); }
    };
    reader.readAsBinaryString(file);
  };

  const handleManualUploadRealisasi = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try { setRealisasiWorkbook(XLSX.read(evt.target?.result, { type: 'binary' })); alert(`File ${file.name} berhasil dibaca!`); } 
      catch { alert('Gagal membaca file Realisasi JO.xlsx.'); }
    };
    reader.readAsBinaryString(file);
  };

  // Parse Realisasi Map
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
    try {
      realisasiWorkbook.SheetNames.forEach(sheetName => {
        const sheet = realisasiWorkbook.Sheets[sheetName]; if (!sheet) return;
        const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }); if (rawRows.length === 0) return;
        let headerIdx = -1, jobcardCol = -1, effCol = -1, otCol = -1;
        for (let r = 0; r < Math.min(5, rawRows.length); r++) {
          const row = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
          const jcIdx = row.findIndex(c => c.includes('jobcard'));
          const eIdx = row.findIndex(c => c.includes('effective'));
          const oIdx = row.findIndex(c => c.includes('overtime'));
          if (jcIdx !== -1 && (eIdx !== -1 || oIdx !== -1)) { headerIdx = r; jobcardCol = jcIdx; effCol = eIdx; otCol = oIdx; break; }
        }
        if (jobcardCol === -1) jobcardCol = 16; if (effCol === -1) effCol = 11; if (otCol === -1) otCol = 12;
        for (let r = Math.max(headerIdx + 1, 1); r < rawRows.length; r++) {
          const row = rawRows[r]; if (!row) continue;
          const rawJc = String(row[jobcardCol] || '').trim();
          if (!rawJc || rawJc.toLowerCase() === 'nan' || rawJc.toLowerCase().includes('jobcard')) continue;
          const eff = effCol !== -1 ? parseValToNumber(row[effCol]) : 0;
          const ot = otCol !== -1 ? parseValToNumber(row[otCol]) : 0;
          const total = eff + ot;
          const key = cleanText(rawJc); if (key) map.set(key, (map.get(key) || 0) + total);
        }
      });
    } catch {}
    return map;
  }, [realisasiWorkbook]);

  // Parse IM4 Members
  const allParsedFromExcel = useMemo<ParsedMember[]>(() => {
    if (!im4Workbook) return [];
    try {
      const sheetName = im4Workbook.SheetNames.find(s => s.toLowerCase().includes('education')) || im4Workbook.SheetNames[0];
      const sheet = im4Workbook.Sheets[sheetName]; if (!sheet) return [];
      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      let headerIdx = -1;
      for (let r = 0; r < Math.min(15, rawRows.length); r++) {
        const rowVals = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
        if (rowVals.includes('nama') && (rowVals.includes('nip') || rowVals.includes('status') || rowVals.includes('jabatan'))) { headerIdx = r; break; }
      }
      if (headerIdx === -1) return [];
      const headers = rawRows[headerIdx].map(v => String(v).trim().toLowerCase());
      const namaCol = headers.findIndex(h => h === 'nama');
      const nipCol = headers.findIndex(h => h === 'nip');
      const statusCol = headers.findIndex(h => h === 'status');
      const unitCol = headers.findIndex(h => h.includes('unit'));
      const jabatanCol = headers.findIndex(h => h.includes('jabatan'));
      let currentDept = '', currentBiro = '';
      const results: ParsedMember[] = [];
      for (let r = headerIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r]; if (!row) continue;
        const nama = String(row[namaCol] || '').trim(); if (!nama || nama.toLowerCase() === 'nan' || nama.toLowerCase() === 'nama') continue;
        const nip = String(row[nipCol] || '').trim();
        const statusRaw = String(row[statusCol] || '').trim();
        const unit = String(row[unitCol] || '').trim();
        const jabatan = String(row[jabatanCol] || '').trim();
        if (jabatan.toLowerCase().includes('kepala divisi')) { currentDept = 'Div. Desain'; currentBiro = 'Div. Desain'; } 
        else if (jabatan.toLowerCase().includes('kepala departemen') || jabatan.toLowerCase().includes('kadep')) {
          if (unit && unit.toLowerCase() !== 'nan') currentDept = unit; else currentDept = jabatan; currentBiro = `Staf ${currentDept}`;
        } else if (jabatan.toLowerCase().includes('kepala biro') || jabatan.toLowerCase().includes('kabiro')) {
          currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').replace(/Kabiro/gi, 'Biro').trim();
          if (unit && unit.toLowerCase() !== 'nan') currentDept = unit;
        }
        results.push({ nama, nip, status: statusRaw || 'PKWTT', jabatan, biro: currentBiro, dept: currentDept });
      }
      return results;
    } catch { return []; }
  }, [im4Workbook]);

  // ✅ PISAHKAN: Organik (PKWT/PKWTT) vs Outsourcing
  const getOrganicBiroMembers = useCallback((biroName: string) => {
    return allParsedFromExcel.filter(p => 
      isBiroMatch(p.biro, biroName) && 
      !p.status.toLowerCase().includes('outsourcing')
    );
  }, [allParsedFromExcel]);

  const getOutsourcingBiroMembers = useCallback((biroName: string) => {
    return allParsedFromExcel.filter(p => 
      isBiroMatch(p.biro, biroName) && 
      p.status.toLowerCase().includes('outsourcing')
    );
  }, [allParsedFromExcel]);

  const toggleAccordion = (cardKey: string) => {
    setExpandedCards(prev => ({ ...prev, [cardKey]: !prev[cardKey] }));
  };

  // Save Kode JC (Approve Planner)
  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;
    const { error } = await supabase.from('job_cards').update({ kode_jc: inputVal, status: 'approved' }).eq('id', taskId);
    if (error) { alert('Gagal simpan Jobcard: ' + error.message); return; }
    
    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => {
        updated[k] = updated[k].map(t => t.id === taskId ? { ...t, kodeJc: inputVal } : t);
      });
      return updated;
    });
    setEditingTaskKode(prev => { const next = { ...prev }; delete next[taskId]; return next; });
    loadAllJobCards();
  };

  // Quick Save Release
  const handleQuickSaveRelease = async (taskId: string, newVal: string) => {
    const cleanDate = parseToStandardDate(newVal);
    saveLocalRelease(taskId, cleanDate);
    try {
      await supabase.from('job_cards').update({ release: cleanDate }).eq('id', taskId);
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

  // Sync Release from Drive
  const handleSyncReleaseFromDrive = async (task: TaskItem) => {
    const cleanProj = cleanText(task.project);
    let rows = drawingControlMap[cleanProj];

    if (!rows || rows.length === 0) {
      setIsFetchingDrawing(true);
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
      } catch (err) { console.error(err); }
      finally { setIsFetchingDrawing(false); }
    }

    if (rows && rows.length > 0) {
      const cleanTarget = cleanText(task.taskName);
      const currentTaskRev = String(task.rev || '0').trim();
      
      const exactMatch = rows.find(r => {
        const cFull = cleanText(r.fullDeskripsi);
        const cDwg = cleanText(r.noDwg);
        const cName = cleanText(r.drawingName);
        const rowRev = String(r.rev || '0').trim();
        const isRevMatch = rowRev === currentTaskRev;
        const isNameMatch = cleanTarget.includes(cDwg) || cleanTarget.includes(cName) || cFull.includes(cleanTarget) || cleanTarget.includes(cFull);
        return isNameMatch && isRevMatch;
      });

      if (exactMatch && exactMatch.finishDate) {
        const standardDate = parseToStandardDate(exactMatch.finishDate);
        await handleQuickSaveRelease(task.id, standardDate);
        alert(`Berhasil sinkronisasi dari Google Drive: ${formatDisplayDate(standardDate)}`);
        return;
      }
    }
    alert(`Tidak ditemukan data FINISH DATE di Google Drive untuk gambar "${task.taskName}" dengan Rev ${task.rev}.`);
  };

  const filteredDepartments = (departmentsData || []).filter(dept =>
    dept.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const currentBiroTasks = selectedBiroName ? (manualTasks[cleanText(selectedBiroName)] || []) : [];
  const currentOrganicMembers = selectedBiroName ? getOrganicBiroMembers(selectedBiroName) : [];
  const currentOutsourcingMembers = selectedBiroName ? getOutsourcingBiroMembers(selectedBiroName) : [];
  
  const pendingTasksCount = useMemo(() => {
    let count = 0;
    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        if (!t.kodeJc || t.kodeJc.trim() === '') count++;
      });
    });
    return count;
  }, [manualTasks]);

  // Fungsi untuk render tabel tugas (dipakai di organik & outsourcing)
  const renderTaskTable = (members: ParsedMember[]) => (
    <div className="space-y-2.5">
      {members.length > 0 ? (
        members.map((person) => {
          const personTasks = currentBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
          const isExpanded = !!expandedCards[person.nama];

          return (
            <div key={`${person.nama}-${person.nip}`} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
              <div
                onClick={() => toggleAccordion(person.nama)}
                className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition"
              >
                <div className="flex items-center gap-3">
                  <div className={`w-8 h-8 rounded-lg flex items-center justify-center font-bold text-xs ${
                    memberTab === 'organic' 
                      ? 'bg-blue-500/10 border border-blue-500/20 text-blue-400' 
                      : 'bg-amber-500/10 border border-amber-500/20 text-amber-400'
                  }`}>
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
                  <span className="text-xs font-mono text-slate-400">{personTasks.length} Tugas</span>
                  <div className={`p-1 rounded bg-slate-950 border border-slate-800 text-slate-400 ${isExpanded ? (memberTab === 'organic' ? 'text-blue-400' : 'text-amber-400') : ''}`}>
                    {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                  </div>
                </div>
              </div>

              {isExpanded && (
                <div className="p-3.5 border-t border-slate-800 bg-slate-950/60">
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
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800 text-slate-300">
                          {personTasks.map((task, tIdx) => {
                            const jcKey = cleanText(task.kodeJc || '');
                            const calculatedRealHours = jcKey ? realisasiMap.get(jcKey) : undefined;

                            return (
                              <tr key={task.id} className="hover:bg-slate-900/40">
                                <td className="py-2 px-2.5 text-slate-500 font-mono">{tIdx + 1}</td>
                                <td className="py-2 px-2.5 font-mono font-bold text-amber-300">
                                  {task.kodeJc || <span className="text-rose-400 font-normal">Pending</span>}
                                </td>
                                <td className="py-2 px-2.5 text-emerald-400">{task.project}</td>
                                <td className="py-2 px-2.5 text-slate-200 truncate max-w-[200px]" title={task.taskName}>{task.taskName}</td>
                                <td className="py-2 px-2.5 text-center font-mono">{task.rev || '0'}</td>
                                <td className="py-2 px-2.5 font-mono text-[11px]">{formatDisplayDate(task.startDate)}</td>
                                <td className="py-2 px-2.5 font-mono text-[11px]">{formatDisplayDate(task.endDate)}</td>
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
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="text-xs text-slate-500 py-3 text-center italic">Belum ada tugas untuk personel ini.</div>
                  )}
                </div>
              )}
            </div>
          );
        })
      ) : (
        <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
          {memberTab === 'organic' 
            ? 'Tidak ada pegawai organik (PKWT/PKWTT) terdaftar di biro ini.' 
            : 'Tidak ada personel outsourcing terdaftar di biro ini.'}
        </div>
      )}
    </div>
  );

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5 cursor-pointer" onClick={() => { setSelectedDept(null); setSelectedBiroName(null); }}>
            <div className="p-1.5 bg-purple-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <span className="font-bold text-sm text-white">PORTAL ADMIN RENDAL — <span className="text-cyan-400">{user.nama}</span></span>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setIsPlannerOpen(true)}
              className="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer relative"
            >
              <Lock className="w-3.5 h-3.5" />
              <span>Planner Panel</span>
              {pendingTasksCount > 0 && (
                <span className="absolute -top-1.5 -right-1.5 bg-rose-500 text-white text-[10px] font-bold px-1.5 py-0.2 rounded-full border border-slate-950 animate-pulse">
                  {pendingTasksCount}
                </span>
              )}
            </button>
            <button
              onClick={onLogout}
              className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer"
            >
              Keluar
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-6xl mx-auto px-4 py-8 space-y-6">
        
        {/* Toolbar Upload & Search */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
          <div>
            <h2 className="text-lg font-bold text-white">Monitoring Seluruh Departemen & Biro</h2>
            <span className="text-xs text-slate-400">Akses lintas seluruh karyawan organik dan subkontraktor</span>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
              <Users className="w-3.5 h-3.5 text-blue-400" />
              <span>Update IM4</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleUpdateIm4Excel} className="hidden" />
            </label>
            <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
              <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" />
              <span>Update Jobcard</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleUpdateJobcardExcel} className="hidden" />
            </label>
            <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
              <Clock className="w-3.5 h-3.5 text-emerald-400" />
              <span>{realisasiMap.size > 0 ? `Realisasi: OK (${realisasiMap.size})` : 'Upload Realisasi JO'}</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
            </label>
            <input
              type="text"
              placeholder="Cari Departemen..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44"
            />
          </div>
        </div>

        {/* 1. Level Departemen */}
        {!selectedDept && (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
            {filteredDepartments.map((dept) => (
              <div
                key={dept.id}
                onClick={() => setSelectedDept(dept)}
                className="bg-slate-900 border border-slate-800 hover:border-purple-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
              >
                <div>
                  <h4 className="font-semibold text-white text-sm">{dept.name}</h4>
                  <span className="text-xs text-slate-400">{dept.biros.length} Biro</span>
                </div>
                <ChevronRight className="w-4 h-4 text-slate-500" />
              </div>
            ))}
          </div>
        )}

        {/* 2. Level Biro */}
        {selectedDept && !selectedBiroName && (
          <div className="space-y-4">
            <button onClick={() => setSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Departemen
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{selectedDept.name}</h2>
              <span className="text-xs text-slate-400">Pilih biro untuk melihat detail penugasan</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {selectedDept.biros.map((biro) => (
                <div
                  key={biro.id}
                  onClick={() => { setSelectedBiroName(biro.name); setMemberTab('organic'); }}
                  className="bg-slate-900 border border-slate-800 hover:border-purple-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
                >
                  <span className="font-semibold text-white text-sm">{biro.name}</span>
                  <ChevronRight className="w-4 h-4 text-slate-500" />
                </div>
              ))}
            </div>
          </div>
        )}

        {/* 3. Level Daftar Anggota & Tabel Tugas (DENGAN TAB ORGANIK / OUTSOURCING) */}
        {selectedDept && selectedBiroName && (
          <div className="space-y-4">
            <button onClick={() => setSelectedBiroName(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Daftar Biro
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{selectedBiroName}</h2>
              <span className="text-xs text-slate-400">
                {currentOrganicMembers.length} Organik • {currentOutsourcingMembers.length} Outsourcing
              </span>
            </div>

            {/* TAB SWITCHER */}
            <div className="bg-slate-950 p-1 rounded-lg border border-slate-800 flex gap-1 text-xs font-semibold w-fit">
              <button
                onClick={() => setMemberTab('organic')}
                className={`px-4 py-2 rounded transition cursor-pointer flex items-center gap-2 ${
                  memberTab === 'organic' 
                    ? 'bg-blue-600 text-white font-bold shadow' 
                    : 'text-slate-400 hover:text-white hover:bg-slate-900'
                }`}
              >
                <Briefcase className="w-3.5 h-3.5" />
                Organik (PKWT/PKWTT)
                <span className="bg-white/10 px-1.5 py-0.5 rounded text-[10px]">{currentOrganicMembers.length}</span>
              </button>
              <button
                onClick={() => setMemberTab('outsourcing')}
                className={`px-4 py-2 rounded transition cursor-pointer flex items-center gap-2 ${
                  memberTab === 'outsourcing' 
                    ? 'bg-amber-600 text-white font-bold shadow' 
                    : 'text-slate-400 hover:text-white hover:bg-slate-900'
                }`}
              >
                <HardHat className="w-3.5 h-3.5" />
                Outsourcing
                <span className="bg-white/10 px-1.5 py-0.5 rounded text-[10px]">{currentOutsourcingMembers.length}</span>
              </button>
            </div>

            {/* KONTEN TAB */}
            {memberTab === 'organic' && renderTaskTable(currentOrganicMembers)}
            {memberTab === 'outsourcing' && renderTaskTable(currentOutsourcingMembers)}
          </div>
        )}
      </main>

      {/* Modal Planner Panel */}
      {isPlannerOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
            <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <span className="font-bold text-xs text-white flex items-center gap-1.5">
                <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel — Approval Jobcard ({pendingTasksCount} Menunggu)
              </span>
              <button onClick={() => setIsPlannerOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4 max-h-96 overflow-y-auto space-y-2">
              {Object.values(manualTasks).flatMap(tasks => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').length > 0 ? (
                Object.values(manualTasks).flatMap(tasks => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').map((task, idx) => (
                  <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                    <div>
                      <div className="font-semibold text-white">{task.pic} <span className="font-mono text-slate-500">#{idx + 1}</span></div>
                      <div className="text-slate-400 text-[11px]">{task.taskName}</div>
                      <div className="text-emerald-400 font-mono text-[10px]">{task.project} • {task.biroName}</div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <input
                        type="text"
                        value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                        onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                        placeholder="No Jobcard..."
                        className="w-36 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase focus:outline-none"
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
                <div className="py-8 text-center text-xs text-slate-500">Tidak ada pengajuan tugas yang menunggu approval.</div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}