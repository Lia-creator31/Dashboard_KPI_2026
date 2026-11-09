import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import KpiDashboardPage from './KpiDashboardPage';
import ImportJobcardPage from './ImportJobcardPage';
import RekapAbsensiDivisiPage, { DivisiGroup } from './RekapAbsensiDivisiPage';
import { projectCandidates, dedupeRows } from '../lib/projectCandidates';
import {
  Building2, Briefcase, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, X, Check, Printer, FileCheck, Clock, Sparkles,
  FileSpreadsheet, Users, Search, TrendingUp
} from 'lucide-react';

const excelGlobUrls = import.meta.glob('./*.xlsx', {
  query: '?url',
  import: 'default',
  eager: true
}) as Record<string, string>;

const parentExcelGlobUrls = import.meta.glob('../*.xlsx', {
  query: '?url',
  import: 'default',
  eager: true
}) as Record<string, string>;

const MONTHS: { key: string; label: string; num: string }[] = [];
const absensiCsvGlob = import.meta.glob('../absensi_*.csv', {
  query: '?raw',
  import: 'default',
  eager: true
}) as Record<string, string>;

const csvMonthMap: Record<string, string> = {};
Object.entries(absensiCsvGlob).forEach(([path, raw]) => {
  const m = MONTHS.find(x => path.toLowerCase().includes(x.key));
  if (m) csvMonthMap[m.key] = raw;
});

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
  status?: string;
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

interface SelectedFormPage {
  biroName: string;
  deptName: string;
}

interface RendalPageProps {
  user: UserSession;
  onLogout: () => void;
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

export function cleanText(str: string): string {
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

export function formatDisplayDate(val: any): string {
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
  if ((b1.includes('kapal selam') || b1.includes('submarine') || b1.includes('scorpne')) && (b2.includes('kapal selam') || b2.includes('submarine') || b2.includes('scorpne'))) return true;
  if (b1.includes('non kapal') && b2.includes('non kapal')) return true;
  if ((b1.includes('kapal permukaan') || b1.includes('surface')) && (b2.includes('kapal permukaan') || b2.includes('surface'))) return true;
  const c1 = cleanText(b1.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  const c2 = cleanText(b2.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  if (c1 && c2) return c1 === c2 || c1.includes(c2) || c2.includes(c1);
  return false;
}

function isPimpinanDivDept(jabatan: string): boolean {
  const j = (jabatan || '').toLowerCase();
  return j.includes('kepala divisi') || j.includes('kepala departemen') || /\bkadep\b/.test(j) || /\bkadiv\b/.test(j);
}

function getBiroPrefix(biroName: string): string {
  const b = (biroName || '').toLowerCase().replace('&', ' dan ').trim();
  if (b.includes('dokumen') || (b.includes('perencanaan') && b.includes('biro'))) return 'DP';
  if (b.includes('logistik')) return 'DL'; if (b.includes('administrasi')) return 'DA'; if (b.includes('pengembangan')) return 'PD';
  if (b.includes('non kapal')) return 'NK'; if (b.includes('kapal selam') || b.includes('submarine')) return 'KS';
  if (b.includes('kapal permukaan') || b.includes('surface')) return 'KP'; if (b.includes('struktur') && b.includes('lambung')) return 'SL';
  if (b.includes('akomodasi')) return 'AK'; if (b.includes('perlengkapan') && b.includes('lambung')) return 'PL';
  if (b.includes('produksi') && b.includes('lambung')) return 'PR'; if (b.includes('propulsi')) return 'SP';
  if (b.includes('pengaturan') || b.includes('permesinan')) return 'PP'; if (b.includes('hvac') || b.includes('geladak')) return 'HV';
  if (b.includes('listrik') || b.includes('kelistrikan')) return 'SK'; if (b.includes('kontrol') || b.includes('otomasi')) return 'KO';
  if (b.includes('elektronika')) return 'SE'; if (b.includes('hps')) return 'HP';
  const words = b.replace(/biro|desain|dasar/gi, '').split(/\s+/).filter(Boolean);
  if (words.length >= 2) return (words[0][0] + words[1][0]).toUpperCase();
  if (words.length === 1 && words[0].length >= 2) return words[0].slice(0, 2).toUpperCase();
  return 'WO';
}

function cleanProjectString(raw: any): string {
  if (!raw) return '';
  let s = String(raw).replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ').trim();
  s = s.replace(/[\.,;:\-_/\\\s]+$/, '').trim(); s = s.toUpperCase();
  s = s.replace(/([A-Z0-9])O(\d+)$/, '$10$2'); s = s.replace(/(\d)0P(\d)/, '$1OP$2');
  if (/^W[O0]{2,}\d+/.test(s)) s = 'W' + s.slice(1).replace(/[O0]/g, '0');
  return s.trim();
}

function getProjectNormKey(s: string): string {
  const k = (s || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  return k.replace(/0/g, 'o');
}

async function fetchDrawingRowsRaw(projectCode: string): Promise<DrawingControlRow[]> {
  const cleanProj = cleanText(projectCode);
  if (!cleanProj) return [];
  try {
    const res = await fetch(`${GAS_DRAWING_API_URL}?project=${encodeURIComponent(projectCode.trim())}`);
    if (res.ok) {
      const json = await res.json();
      if (json.success && Array.isArray(json.data) && json.data.length > 0) {
        return json.data.map((item: any) => ({
          noDwg: String(item.noDwg || '').trim(), drawingName: String(item.drawingName || '').trim(),
          fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
          rev: String(item.rev || '0').trim(), finishDate: String(item.finishDate || '').trim(),
        }));
      }
    }
  } catch { /* lanjut ke cadangan CSV */ }
  const match = GOOGLE_DRIVE_SHEETS.find(s => cleanText(s.projectKey) === cleanProj || cleanProj.includes(cleanText(s.projectKey)));
  if (match) {
    try {
      const res = await fetch(`https://docs.google.com/spreadsheets/d/${match.id}/gviz/tq?tqx=out:csv&sheet=Drawing%20Control%20(2)`);
      if (res.ok) {
        const wb = XLSX.read(await res.text(), { type: 'string' });
        const rawRows: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
        const rows: DrawingControlRow[] = [];
        for (let r = 1; r < rawRows.length; r++) {
          const row = rawRows[r]; if (!row) continue;
          const noDwg = String(row[0] || '').trim(); const dwgName = String(row[1] || '').trim();
          if (!noDwg && !dwgName) continue;
          rows.push({
            noDwg, drawingName: dwgName,
            fullDeskripsi: (noDwg && dwgName) ? `${noDwg}-${dwgName}` : (dwgName || noDwg),
            rev: String(row[14] || '0').trim(), finishDate: String(row[15] || '').trim(),
          });
        }
        if (rows.length > 0) return rows;
      }
    } catch { /* tidak ada data */ }
  }
  return [];
}

const drawingRowsCache = new Map<string, DrawingControlRow[]>();

function isKnownDriveProject(projectCode: string): boolean {
  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  return cands.some(c => {
    const cp = cleanText(c);
    return cp && GOOGLE_DRIVE_SHEETS.some(s => cleanText(s.projectKey) === cp || cp.includes(cleanText(s.projectKey)));
  });
}

async function fetchDrawingRowsFor(projectCode: string): Promise<DrawingControlRow[]> {
  const cacheKey = cleanText(projectCode);
  if (cacheKey && drawingRowsCache.has(cacheKey)) {
    const cached = drawingRowsCache.get(cacheKey)!;
    if (cached.length > 0) return cached;
  }
  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  const results = await Promise.all(cands.map(c => fetchDrawingRowsRaw(c).catch(() => [] as DrawingControlRow[])));
  const deduped = dedupeRows(results.flat());
  if (cacheKey && deduped.length > 0) drawingRowsCache.set(cacheKey, deduped);
  return deduped;
}

function normalizeDrawingSpacing(desc: string): string {
  return (desc || '')
    .replace(/(\b[A-Z]{1,4}\d{2,}[A-Z0-9]*)\s*\.\s*([A-Z0-9]+)/gi, '$1.$2')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractDrawingNumber(desc: string): string {
  let s = normalizeDrawingSpacing(desc);
  s = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-\s_:]\s*/i, '');
  const m = s.match(/^[A-Z]{1,4}\d{3,}[A-Z0-9]*(?:\.[A-Z0-9]+)*/i);
  if (m) return m[0].toUpperCase();
  const mMid = s.match(/\b[A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+\b/i);
  if (mMid) return mMid[0].toUpperCase();
  const first = s.split(/\s+-\s+|\s+/)[0] || '';
  return /\d/.test(first) && first.length >= 4 ? first.toUpperCase() : '';
}

const drawingKey = (v: string) => cleanText(extractDrawingNumber(v) || v);

function findReleaseInRows(rows: DrawingControlRow[], taskName: string, rev: string | number): string {
  if (!rows || rows.length === 0 || !taskName) return '';
  const wanted = drawingKey(taskName);
  const cleanTarget = cleanText(taskName);
  if (!wanted && !cleanTarget) return '';
  const toNum = (v: any) => { const n = parseInt(String(v ?? '').replace(/[^0-9]/g, ''), 10); return isNaN(n) ? 0 : n; };
  const wantRev = toNum(rev);
  const hit = rows.find(r => {
    if (toNum(r.rev) !== wantRev) return false;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(parseToStandardDate(r.finishDate))) return false;
    const rKey = drawingKey(r.noDwg);
    if (wanted && rKey === wanted) return true;
    const cDwg = cleanText(r.noDwg);
    if (cDwg && cDwg.length >= 5 && cleanTarget.includes(cDwg)) return true;
    return false;
  });
  return hit ? parseToStandardDate(hit.finishDate) : '';
}

const MASTER_BUCKET = 'master-files';
type MasterKey = 'jobcard' | 'im4' | 'realisasi' | 'timeoff';

async function saveMasterFile(key: MasterKey, file: File) {
  const { error } = await supabase.storage
    .from(MASTER_BUCKET)
    .upload(`${key}.xlsx`, file, { upsert: true, contentType: file.type || 'application/octet-stream' });
  return error;
}

async function loadMasterFile(key: MasterKey): Promise<XLSX.WorkBook | null> {
  try {
    const { data, error } = await supabase.storage.from(MASTER_BUCKET).download(`${key}.xlsx`);
    if (error || !data) return null;
    return XLSX.read(await data.arrayBuffer(), { type: 'array' });
  } catch { return null; }
}

const HOUR_PLANNER_FILE = 'hourplanner.json';
interface HourPlannerEntry {
  kode: string;
  hour: number;
  planner: string;
  start: string;
  end: string;
}
type HourPlannerMap = Record<string, HourPlannerEntry>;

async function loadHourPlanner(): Promise<HourPlannerMap> {
  try {
    const { data, error } = await supabase.storage.from(MASTER_BUCKET).download(HOUR_PLANNER_FILE);
    if (error || !data) return {};
    const parsed = JSON.parse(await data.text());
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch { return {}; }
}

async function saveHourPlanner(map: HourPlannerMap) {
  const blob = new Blob([JSON.stringify(map)], { type: 'application/octet-stream' });
  const { error } = await supabase.storage
    .from(MASTER_BUCKET)
    .upload(HOUR_PLANNER_FILE, blob, { upsert: true, contentType: 'application/octet-stream' });
  return error;
}

function parseHourPlannerWorkbook(wb: XLSX.WorkBook): { entries: HourPlannerMap; rowCount: number; error?: string } {
  const entries: HourPlannerMap = {};
  let rowCount = 0;
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return { entries, rowCount, error: 'Sheet tidak ditemukan.' };
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  let headerIdx = -1, nameCol = -1, hourCol = -1, plannerCol = -1, startCol = -1, endCol = -1;
  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const h = (rows[r] || []).map(v => String(v).trim().toLowerCase());
    const n = h.findIndex(c => c === 'name' || c === 'jobcard' || c === 'kode jobcard');
    const hr = h.findIndex(c => c.includes('hour') && c.includes('plan'));
    if (n !== -1 && hr !== -1) {
      headerIdx = r; nameCol = n; hourCol = hr;
      plannerCol = h.findIndex(c => c === 'planner');
      startCol = h.findIndex(c => c.includes('start'));
      endCol = h.findIndex(c => c.includes('end'));
      break;
    }
  }
  if (headerIdx === -1) return { entries, rowCount, error: 'Kolom "Name" dan "Hour Planned" tidak ditemukan di baris judul.' };
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r]; if (!row) continue;
    const kode = String(row[nameCol] ?? '').replace(/\s+/g, ' ').trim();
    if (!kode || kode.toLowerCase() === 'nan') continue;
    const hourRaw = String(row[hourCol] ?? '').trim().replace(',', '.');
    const hour = parseFloat(hourRaw);
    if (isNaN(hour)) continue;
    const key = cleanText(kode);
    if (!key) continue;
    rowCount++;
    entries[key] = {
      kode,
      hour,
      planner: plannerCol !== -1 ? String(row[plannerCol] ?? '').trim() : '',
      start: startCol !== -1 ? parseToStandardDate(row[startCol]) : '',
      end: endCol !== -1 ? parseToStandardDate(row[endCol]) : '',
    };
  }
  return { entries, rowCount };
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

interface AbsensiRow {
  nip: string;
  nama: string;
  planned: number;
  effective: number;
  overtime: number;
  idle: number;
  terlambat: number;
  sakit: number;
}

const absNum = (v: any): number => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const n = parseFloat(String(v).trim().replace('%', '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
};

export const fmtNum = (n: number) => String(Math.round(n * 100) / 100);

function findMonthSheet(wb: XLSX.WorkBook | null, monthKey: string): string | undefined {
  if (!wb) return undefined;
  return wb.SheetNames.find(s => {
    const l = s.toLowerCase().trim();
    return l.includes(monthKey) || l.includes(monthKey.slice(0, 3));
  });
}

function buildRekap(wb: XLSX.WorkBook | null, biroName: string, monthKey: string, csv: string): AbsensiRow[] {
  const sheetName = findMonthSheet(wb, monthKey);
  if (!wb || !sheetName) return [];
  const abs = new Map<string, { terlambat: number; sakit: number }>();
  (csv || '').split(/\r?\n/).forEach((line, i) => {
    if (i === 0 || !line.trim()) return;
    const p = line.split(',');
    if (p.length < 6) return;
    const k = cleanText(p[1] || '');
    if (k) abs.set(k, { terlambat: absNum(p[p.length - 3]), sakit: absNum(p[p.length - 2]) });
  });
  const rows: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
  const head = (rows[0] || []).map(v => String(v).trim().toLowerCase());
  const planIdx = head.findIndex(h => h.includes('plan'));
  const PLAN_COL = planIdx >= 0 ? planIdx : 9;
  const people = new Map<string, AbsensiRow>();
  rows.forEach((row, i) => {
    if (i === 0 || !row) return;
    const nip = String(row[0] || '').trim();
    const nama = String(row[1] || '').trim();
    if (!(nama || nip)) return;
    if (!isBiroMatch(String(row[7] || ''), biroName)) return;
    const key = cleanText(nama || nip);
    const p = people.get(key) || { nip: nip || '-', nama: nama || '-', planned: 0, effective: 0, overtime: 0, idle: 0, terlambat: 0, sakit: 0 };
    p.planned += absNum(row[PLAN_COL]);
    p.effective += absNum(row[10]);
    p.overtime += absNum(row[11]);
    p.idle += absNum(row[12]);
    people.set(key, p);
  });
  return Array.from(people.values()).map(p => {
    const a = abs.get(cleanText(p.nama));
    return { ...p, terlambat: a ? a.terlambat : 0, sakit: a ? a.sakit : 0 };
  });
}

export function AbsensiBiroView({ biroName, deptName, wb, monthKey, onChangeMonth, onUploadKpi }: {
  biroName: string;
  deptName: string;
  wb: XLSX.WorkBook | null;
  monthKey: string | null;
  onChangeMonth: (k: string | null) => void;
  onUploadKpi: (e: ChangeEvent<HTMLInputElement>) => void;
}) {
  const [q, setQ] = useState('');
  const month = MONTHS.find(m => m.key === monthKey) || null;
  const allRows = useMemo(
    () => (month ? buildRekap(wb, biroName, month.key, csvMonthMap[month.key] || '') : []),
    [wb, biroName, month]
  );
  const rows = useMemo(() => {
    const s = q.trim().toLowerCase();
    return s ? allRows.filter(r => r.nip.toLowerCase().includes(s) || r.nama.toLowerCase().includes(s)) : allRows;
  }, [allRows, q]);

  if (!month) {
    return (
      <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 space-y-4">
        <div className="flex items-start justify-between gap-3 flex-wrap">
        </div>
        {!wb && (
          <div className="text-[11px] text-amber-400">
            File data_kpi.xlsx belum termuat. Letakkan di folder src/ atau public/, atau upload manual dengan tombol di atas.
          </div>
        )}
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
          {MONTHS.map(m => {
            const ada = !!findMonthSheet(wb, m.key);
            return (
              <button key={m.key} disabled={!ada} onClick={() => onChangeMonth(m.key)}
                className={`p-4 rounded-xl border text-left transition ${ada ? 'bg-slate-950 border-slate-700 hover:border-blue-500 cursor-pointer' : 'bg-slate-950/40 border-slate-800 opacity-40 cursor-not-allowed'}`}>
                <div className="font-semibold text-white text-sm">{m.label}</div>
                <div className="text-[11px] text-slate-500 font-mono">{ada ? 'Data tersedia' : 'Sheet belum ada'}</div>
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 flex items-start justify-between gap-4 flex-wrap">
        <div className="space-y-1">
          <button onClick={() => { onChangeMonth(null); setQ(''); }} className="text-xs text-blue-400 hover:underline flex items-center gap-1 cursor-pointer">
            <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Pilihan Bulan
          </button>
          <div className="text-[11px] font-bold text-blue-400 uppercase tracking-wide">{deptName}</div>
          <h3 className="text-xl font-bold text-white">{biroName}</h3>
          <div className="text-xs text-slate-400">
            Rekapitulasi Jam Kerja &amp; Absensi Pegawai — <span className="text-blue-400 font-semibold">Bulan {month.label}</span>
          </div>
        </div>
        <div className="px-5 py-3 bg-slate-950 border border-slate-800 rounded-xl text-right">
          <div className="text-2xl font-bold text-blue-400">{allRows.length}</div>
          <div className="text-[10px] text-slate-400 uppercase tracking-wide">Total Pegawai</div>
        </div>
      </div>
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="Cari NIP atau Nama Pegawai..."
          className="px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-72" />
        <span className="text-[11px] text-slate-400">Menampilkan <b className="text-white">{rows.length}</b> data pegawai</span>
      </div>
      <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-x-auto">
        <table className="w-full text-left text-xs">
          <thead>
            <tr className="text-slate-400 border-b border-slate-800 uppercase text-[11px]">
              <th className="py-3 px-3 w-10">No</th>
              <th className="py-3 px-3">Nama Pegawai</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800 font-mono">
            {rows.length > 0 ? rows.map((r, i) => (
              <tr key={`${r.nip}-${i}`} className="hover:bg-slate-800/40">
                <td className="py-3 px-3 text-slate-500">{i + 1}</td>
                <td className="py-3 px-3 font-sans font-semibold text-white">{r.nama}</td>
              </tr>
            )) : (
              <tr><td colSpan={2} className="py-8 text-center font-sans text-slate-500">
                {allRows.length === 0
                  ? `Tidak ada data biro ini di sheet ${month.label} pada data_kpi.xlsx (cek kolom H = nama biro).`
                  : 'Tidak ada data yang cocok dengan pencarian.'}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ===== TIME OFF (cuti & sakit) dari TIME_OFF_DIVISI_DESAIN.xlsx ===== */
export interface TimeOffEntry {
  nip: string;
  nama: string;
  type: string;
  kind: 'cuti' | 'sakit' | 'panjang' | 'khusus';
  start: string;
  end: string;
  days: number;
  status: string;
}

export const normNip = (s: any) => String(s ?? '').replace(/\D/g, '').replace(/^0+/, '');
const isoLocal = (d: Date) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
export const shortType = (t: string) =>
  t.replace(/\s*\(.*?\)\s*/g, ' ').replace(/\s*\d{4}\s*$/, '').trim();

function classifyTimeOff(type: string): TimeOffEntry['kind'] {
  const t = type.toLowerCase().trim();
  if (t.startsWith('sakit')) return 'sakit';
  if (t.includes('kewajiban negara') || t.includes('persiapan pensiun') || t === 'cuti melahirkan') return 'panjang';
  if (t.includes('umroh') || t.includes('umrah') || t.includes('haji')) return 'khusus';
  if (t.includes('menikah') || t.includes('pernikahan') || t.includes('kawin')) return 'khusus';
  if (t.includes('istri melahirkan') || t.includes('suami melahirkan') ||
      (t.includes('melahirkan') && (t.includes('istri') || t.includes('suami')))) return 'khusus';
  if ((t.includes('keluarga') || t.includes('anak') || t.includes('istri') ||
       t.includes('suami') || t.includes('orang tua') || t.includes('ayah') ||
       t.includes('ibu') || t.includes('bapak') || t.includes('mamah') ||
       t.includes('papa') || t.includes('saudara')) &&
      (t.includes('sakit') || t.includes('rawat inap') || t.includes('meninggal') ||
       t.includes('duka') || t.includes('wafat') || t.includes('meninggal dunia'))) return 'khusus';
  return 'cuti';
}

export function parseTimeOff(wb: XLSX.WorkBook | null): TimeOffEntry[] {
  if (!wb) return [];
  const out: TimeOffEntry[] = [];
  const isoRe = /^\d{4}-\d{2}-\d{2}$/;
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) return;
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length === 0) return;
    let headerIdx = -1, cEmp = -1, cType = -1, cStart = -1, cEnd = -1, cDur = -1, cStatus = -1;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const head = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      const emp = head.findIndex(h => h === 'employee' || h === 'pegawai' || h === 'nama' || h.includes('employee'));
      const typ = head.findIndex(h => h === 'time off type' || h.includes('time off') || h === 'tipe' || h === 'jenis cuti');
      const st = head.findIndex(h => h === 'start date' || h.includes('start'));
      const en = head.findIndex(h => h === 'end date' || h.includes('end'));
      if (emp >= 0 && typ >= 0 && st >= 0 && en >= 0) {
        headerIdx = r;
        cEmp = emp; cType = typ; cStart = st; cEnd = en;
        cDur = head.findIndex(h => h === 'duration' || h.includes('duration') || h.includes('durasi'));
        cStatus = head.findIndex(h => h === 'status' || h.includes('state'));
        break;
      }
    }
    if (headerIdx < 0) return;
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const emp = String(r[cEmp] || '').trim(); if (!emp) continue;
      const dash = emp.indexOf('-');
      const start = parseToStandardDate(r[cStart]);
      if (!isoRe.test(start)) continue;
      const endRaw = parseToStandardDate(r[cEnd]);
      const type = String(r[cType] || '').trim();
      const rawStatus = cStatus >= 0 ? String(r[cStatus] || '').trim() : 'Approved';
      const normStatus = /^approved|disetujui|validate|confirm$/i.test(rawStatus) ? 'Approved' : rawStatus;
      out.push({
        nip: dash > 0 ? emp.slice(0, dash).trim() : '',
        nama: dash > 0 ? emp.slice(dash + 1).trim() : emp,
        type,
        kind: classifyTimeOff(type),
        start,
        end: isoRe.test(endRaw) ? endRaw : start,
        days: (cDur >= 0 ? parseInt(String(r[cDur] || '').replace(/[^0-9]/g, ''), 10) : 0) || 1,
        status: normStatus,
      });
    }
  });
  return out;
}

function getWeekRange(base: Date = new Date()) {
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const mon = new Date(d); mon.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  const fri = new Date(mon); fri.setDate(mon.getDate() + 4);
  return { from: isoLocal(mon), to: isoLocal(fri) };
}

/* ===== SISA CUTI dari Google Sheet Form Alokasi Cuti ===== */
const SISA_CUTI_SHEET_ID = '1c6TgsfSRrLJ9QZkTQR8aIMe-fBgq89HxmoUdEcqU4b4';
const SISA_CUTI_GID = '1997448035';

export interface SisaCutiInfo {
  nama: string;
  sisa: number;
  filledAt: string;
  rencana: Record<string, string>; // Rencana 1, Rencana 2, dst. dari Google Sheet
  rowIdx?: number;
}

export interface SisaCutiData {
  byName: Map<string, SisaCutiInfo>;
}

function parseSheetTimestamp(val: any, order: 'dmy' | 'mdy'): string {
  const s = String(val ?? '').trim();
  const m = s.match(/^(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/);
  if (!m) return parseToStandardDate(s.split(' ')[0]);
  const a = parseInt(m[1], 10), b = parseInt(m[2], 10), c = parseInt(m[3], 10);
  const pad = (n: number) => String(n).padStart(2, '0');
  if (a > 31) return `${a}-${pad(b)}-${pad(c)}`;
  let day = order === 'dmy' ? a : b;
  let mon = order === 'dmy' ? b : a;
  if (mon > 12 && day <= 12) { const tmp = day; day = mon; mon = tmp; }
  const res = `${c}-${pad(mon)}-${pad(day)}`;
  const todayIso = isoLocal(new Date());
  if (res > todayIso && day <= 12 && mon <= 12) {
    const swapped = `${c}-${pad(day)}-${pad(mon)}`;
    if (swapped <= todayIso) return swapped;
  }
  return res;
}

export function countDaysAfter(e: TimeOffEntry, filledAt: string): number {
  if (!filledAt) return e.days;
  if (e.start > filledAt) return e.days;
  if (e.end <= filledAt) return 0;
  let n = 0;
  const d = new Date(filledAt + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  const endD = new Date(e.end + 'T00:00:00');
  while (d <= endD) {
    const w = d.getDay();
    if (w !== 0 && w !== 6) n++;
    d.setDate(d.getDate() + 1);
  }
  return Math.min(n, e.days);
}

// Batas awal perhitungan pengurang Sisa Cuti dari file Excel Time Off yang diupload:
// Dihitung mulai dari awal bulan berjalan (tanggal 1 bulan ini).
// Jika saat ini masih berada di minggu awal bulan berjalan (Senin minggu kemarin jatuh sebelum tanggal 1 bulan ini),
// maka dihitung mulai dari Senin minggu kemarin.
export function getTimeOffCutoffStart(base: Date = new Date()): string {
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const firstOfMonth = new Date(d.getFullYear(), d.getMonth(), 1);
  const monThisWeek = new Date(d);
  monThisWeek.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  const monPrevWeek = new Date(monThisWeek);
  monPrevWeek.setDate(monThisWeek.getDate() - 7);
  const start = monPrevWeek < firstOfMonth ? monPrevWeek : firstOfMonth;
  return isoLocal(start);
}

export function countDaysFrom(e: TimeOffEntry, startDate: string): number {
  if (!startDate) return e.days;
  if (e.end < startDate) return 0;
  if (e.start >= startDate) return e.days;
  let n = 0;
  const d = new Date(startDate + 'T00:00:00');
  const endD = new Date(e.end + 'T00:00:00');
  while (d <= endD) {
    const w = d.getDay();
    if (w !== 0 && w !== 6) n++;
    d.setDate(d.getDate() + 1);
  }
  return Math.min(n, e.days);
}

// Filter entri cuti/izin/sakit untuk tahun tertentu (misal 2026).
// Jika pegawai sudah memiliki transaksi cuti tahun berjalan (misal "Cuti Tahunan PKWT November 2026"),
// maka sisa cuti periode kontrak tahun sebelumnya ("Cuti PKWT November 2025") tidak dihitung sebagai realisasi tahun berjalan.
export function filterLeavesForYear(mine: TimeOffEntry[], targetYear: number = new Date().getFullYear()): TimeOffEntry[] {
  const targetYearStr = String(targetYear);
  const prevYearStr = String(targetYear - 1);
  const hasTargetYearLeave = mine.some(x => x.type.includes(targetYearStr));

  return mine.filter(e => {
    if (targetYear === 2026) {
      const [y, mm] = e.start.split('-');
      const n = parseInt(mm, 10);
      if (y !== '2026' || n < 8 || n > 12) return false;
    } else {
      if (!e.start.startsWith(targetYearStr)) return false;
    }

    if (hasTargetYearLeave && e.type.includes(prevYearStr) && !e.type.includes(targetYearStr)) {
      return false;
    }

    return true;
  });
}

async function fetchSisaCutiCsv(): Promise<{ text: string; order: 'dmy' | 'mdy' } | null> {
  const cb = Date.now();
  const sources: { url: string; order: 'dmy' | 'mdy' }[] = [
    { url: `https://docs.google.com/spreadsheets/d/${SISA_CUTI_SHEET_ID}/gviz/tq?tqx=out:csv&gid=${SISA_CUTI_GID}&headers=1&_t=${cb}`, order: 'dmy' },
    { url: `https://docs.google.com/spreadsheets/d/${SISA_CUTI_SHEET_ID}/export?format=csv&gid=${SISA_CUTI_GID}&_t=${cb}`, order: 'dmy' },
    { url: `https://docs.google.com/spreadsheets/d/${SISA_CUTI_SHEET_ID}/pub?gid=${SISA_CUTI_GID}&single=true&output=csv&_t=${cb}`, order: 'dmy' },
  ];
  for (const s of sources) {
    try {
      const res = await fetch(s.url, { cache: 'no-store' });
      if (res.ok) {
        const text = await res.text();
        const low = (text || '').trim().toLowerCase();
        if (
          low &&
          !low.startsWith('<!doctype') &&
          !low.startsWith('<html') &&
          (low.includes('nama personel') || low.includes('sisa cuti') || low.includes('timestamp') || low.includes('cap waktu'))
        ) {
          return { text, order: s.order };
        }
      }
    } catch (err) {
      console.warn('Gagal fetch:', s.url, err);
    }
  }
  return null;
}

function parseSisaNumber(raw: string): number | null {
  const m = String(raw ?? '').replace(',', '.').match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
}

function normalizeRencanaCell(raw: any): string {
  const s = String(raw ?? '').trim();
  if (!s || s === '-') return '-';
  // Serial date Excel/Sheets, contoh: 46337.0000462963
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const n = parseFloat(s);
    if (n > 30000 && n < 60000) {
      const d = new Date(Math.round((Math.floor(n) - 25569) * 86400 * 1000));
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      return `${dd}/${mm}/${d.getUTCFullYear()}`;
    }
  }
  return s;
}

async function fetchSisaCuti(): Promise<SisaCutiData | null> {
  try {
    const src = await fetchSisaCutiCsv();
    if (!src) return null;
    const wb = XLSX.read(src.text, { type: 'string' });
    const rows: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    if (rows.length < 2) return null;

    const head = (rows[0] || []).map(v => String(v).trim().toLowerCase());
    let cTs = head.findIndex(h => h.includes('timestamp') || h.includes('cap waktu'));
    let cName = head.findIndex(h => h.includes('nama personel'));
    let cSisa = head.findIndex(h => h.includes('jumlah sisa cuti'));

    // Cari semua kolom yang mengandung kata "rencana" (case-insensitive)
    const rencanaCols: { index: number; name: string }[] = [];
    head.forEach((h, idx) => {
      if (h.includes('rencana')) {
        rencanaCols.push({ index: idx, name: String(rows[0][idx]).trim() });
      }
    });

    if (cTs < 0) cTs = 0;
    if (cName < 0) cName = 2;
    if (cSisa < 0) cSisa = 5;

    // Deteksi otomatis urutan tanggal (dmy vs mdy) dari seluruh baris Timestamp
    let dmyHits = 0, mdyHits = 0;
    for (let i = 1; i < rows.length; i++) {
      const s = String(rows[i]?.[cTs] ?? '').trim();
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (m) {
        const a = parseInt(m[1], 10), b = parseInt(m[2], 10);
        if (a > 12 && b <= 12) dmyHits++;
        else if (b > 12 && a <= 12) mdyHits++;
      }
    }
    const detectedOrder: 'dmy' | 'mdy' = mdyHits > dmyHits ? 'mdy' : dmyHits > mdyHits ? 'dmy' : src.order;

    const byName = new Map<string, SisaCutiInfo>();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const sisaNum = parseSisaNumber(String(r[cSisa] ?? ''));
      if (sisaNum === null) continue;
      const nm = cleanText(String(r[cName] ?? ''));
      if (!nm) continue;

      const rencanaData: Record<string, string> = {};
      rencanaCols.forEach(col => {
        rencanaData[col.name] = normalizeRencanaCell(r[col.index]);
      });

      const info: SisaCutiInfo = {
        nama: String(r[cName] ?? '').trim(),
        sisa: sisaNum,
        filledAt: parseSheetTimestamp(r[cTs], detectedOrder),
        rencana: rencanaData,
        rowIdx: i,
      };

      const old = byName.get(nm);
      // Di Google Sheets, baris yang berada di bawah (i lebih besar)
      // adalah respons formulir / inputan update terbaru dari personel tersebut.
      if (!old || i >= (old.rowIdx ?? 0) || (info.filledAt && old.filledAt && info.filledAt >= old.filledAt)) {
        byName.set(nm, info);
      }
    }
    return { byName };
  } catch { return null; }
}

const normNameToken = (t: string) =>
  t.replace(/ch/g, 'h').replace(/(.)\1+/g, '$1').replace(/^mu/, 'mo');
const nameTokens = (name: string): string[] =>
  String(name || '').toLowerCase().split(/[^a-z]+/).filter(Boolean).map(normNameToken);

function sameNameFuzzy(a: string, b: string): boolean {
  const ta = nameTokens(a), tb = nameTokens(b);
  if (!ta.length || !tb.length) return false;
  const [short, long] = ta.length <= tb.length ? [ta, tb] : [tb, ta];
  if (short.length === 1 && long.length > 1) {
    return short[0].length >= 3 && long.includes(short[0]);
  }
  const used = new Set<number>();
  for (const x of short) {
    let idx = -1;
    for (let i = 0; i < long.length; i++) {
      if (!used.has(i) && long[i] === x) { idx = i; break; }
    }
    if (idx < 0 && !(x.length === 1 && short.length === 1)) {
      for (let i = 0; i < long.length; i++) {
        if (!used.has(i) && (long[i].startsWith(x) || x.startsWith(long[i]))) { idx = i; break; }
      }
    }
    if (idx < 0) return false;
    used.add(idx);
  }
  return true;
}

export function findSisaCuti(data: SisaCutiData | null, nama: string, allNames: string[] = []): SisaCutiInfo | undefined {
  if (!data) return undefined;
  const targetTokens = nameTokens(nama);
  let best: SisaCutiInfo | undefined;
  let bestExactLen = false;
  data.byName.forEach(info => {
    if (!sameNameFuzzy(info.nama, nama)) return;
    const infoTokens = nameTokens(info.nama);
    if (infoTokens.length === 1 && targetTokens.length > 1) {
      const kandidat = allNames.filter(n => sameNameFuzzy(info.nama, n));
      if (kandidat.length > 1) return;
    }
    const isExactLen = infoTokens.length === targetTokens.length;
    const isNewer = !best ||
      (isExactLen && !bestExactLen) ||
      (isExactLen === bestExactLen && (
        (info.rowIdx !== undefined && best.rowIdx !== undefined && info.rowIdx > best.rowIdx) ||
        info.filledAt >= best.filledAt
      ));
    if (isNewer) {
      best = info;
      bestExactLen = isExactLen;
    }
  });
  return best;
}

/* ===== REALISASI JO TANPA JOBCARD + ISI OTOMATIS KODE JOBCARD ===== */
interface RealisasiRow {
  nip: string;
  nama: string;
  project: string;
  jobcard: string;   // kosong jika belum ada Jobcard
  taskName: string;  // kolom "Jobcard/Task Name"
  hours: number;     // Effective + Overtime
}

function splitUserCell(raw: any): { nip: string; nama: string } {
  const s = String(raw ?? '').trim();
  const m = s.match(/^(\d[\d\s]*?)\s*-\s*(.+)$/);
  if (m) return { nip: m[1].trim(), nama: m[2].trim() };
  return { nip: '', nama: s };
}

function extractProjectCodeFromCell(raw: any): string {
  const s = String(raw ?? '').trim();
  if (!s || s === '-' || s.toLowerCase() === 'nan') return '';
  const firstToken = s.split(/[\s-]+/)[0] || '';
  return cleanProjectString(firstToken || s);
}

// Baca file Realisasi per BARIS (tidak membuang baris tanpa Jobcard).
function parseRealisasiRows(wb: XLSX.WorkBook | null): RealisasiRow[] {
  const out: RealisasiRow[] = [];
  if (!wb) return out;
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName]; if (!sheet) return;
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length === 0) return;
    let headerIdx = -1;
    let cNip = -1, cName = 1, cProj = 5, cJc = 16, cEff = 11, cOt = 12, cTask = 22;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const h = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      const jc = h.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
      const ef = h.findIndex(c => c.includes('effective'));
      const ot = h.findIndex(c => c.includes('overtime'));
      if (jc !== -1 && (ef !== -1 || ot !== -1)) {
        headerIdx = r; cJc = jc; cEff = ef; cOt = ot;
        const nm = h.findIndex(c => c === 'employee/name' || c === 'nama' || c === 'name' || c === 'user' || c === 'employee' || c === 'pegawai');
        const np = h.findIndex(c => c === 'employee/nip' || c === 'nip');
        const pj = h.findIndex(c => c === 'project/code' || c === 'project' || c === 'proyek');
        const tk = h.findIndex(c => c === 'jobcard/task name' || c === 'description' || c === 'deskripsi' || c === 'task name');
        if (nm !== -1) cName = nm;
        if (np !== -1) cNip = np;
        if (pj !== -1) cProj = pj;
        if (tk !== -1) cTask = tk;
        break;
      }
    }
    if (headerIdx === -1) return;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const userParsed = splitUserCell(row[cName]);
      const nama = userParsed.nama;
      if (!nama) continue;
      const nip = cNip !== -1 ? String(row[cNip] ?? '').trim() : userParsed.nip;
      const jcRaw = String(row[cJc] ?? '').trim();
      const jobcard = (jcRaw.toLowerCase() === 'nan' || jcRaw.toLowerCase().includes('jobcard')) ? '' : jcRaw;
      const hours = (cEff !== -1 ? absNum(row[cEff]) : 0) + (cOt !== -1 ? absNum(row[cOt]) : 0);
      out.push({
        nip,
        nama,
        project: extractProjectCodeFromCell(row[cProj]),
        jobcard,
        taskName: String(row[cTask] ?? '').trim(),
        hours,
      });
    }
  });
  return out;
}

interface NoJcSummary {
  nama: string;
  hours: number;
  rows: number;
  byProject: Record<string, number>;
}

// Total jam realisasi yang belum punya Jobcard, per pegawai.
function summarizeNoJobcard(rows: RealisasiRow[]): Map<string, NoJcSummary> {
  const m = new Map<string, NoJcSummary>();
  rows.filter(r => !r.jobcard).forEach(r => {
    const k = cleanText(r.nama); if (!k) return;
    const cur = m.get(k) || { nama: r.nama, hours: 0, rows: 0, byProject: {} };
    cur.hours += r.hours;
    cur.rows += 1;
    const p = r.project || '-';
    cur.byProject[p] = (cur.byProject[p] || 0) + r.hours;
    m.set(k, cur);
  });
  return m;
}

// "R0 - EC123C.6 - PROD. DWG ..." -> 0
function revFromTaskName(s: string): number | null {
  const m = (s || '').trim().match(/^R(?:EV)?\.?\s*(\d{1,2})\b/i);
  return m ? parseInt(m[1], 10) : null;
}

// Cocokkan baris Excel yang SUDAH punya Jobcard ke tugas Form yang kode-nya kosong.
// Syarat cocok: nama sama + proyek sama + nomor gambar sama (+ Rev sama bila terbaca).
// Cadangan: jika pegawai itu hanya punya SATU tugas menunggu di proyek tsb dan
// Excel hanya punya SATU kode Jobcard untuk nama+proyek itu, pasangkan keduanya.
// Beberapa personel dalam proyek yang sama diperbolehkan memakai kode Jobcard yang sama
// (pembatas "taken" berlaku per personel, bukan global).
// Jika ada lebih dari satu kemungkinan kode, dilewati (tidak ditebak).
function matchJobcardsToTasks(
  rows: RealisasiRow[],
  pending: TaskItem[],
  usedPersonCodes: Set<string>
): Map<string, string> {
  const withJc = rows.filter(r => r.jobcard);
  const result = new Map<string, string>();
  const taken = new Set(usedPersonCodes);
  const projKey = (p: string) => getProjectNormKey(cleanProjectString(p));
  const toNum = (v: any) => { const n = parseInt(String(v ?? '').replace(/[^0-9]/g, ''), 10); return isNaN(n) ? 0 : n; };

  const pendingCount = new Map<string, number>();
  pending.forEach(t => {
    const k = `${cleanText(t.pic)}|${projKey(t.project)}`;
    pendingCount.set(k, (pendingCount.get(k) || 0) + 1);
  });

  pending.forEach(t => {
    const key = `${cleanText(t.pic)}|${projKey(t.project)}`;
    const cands = withJc.filter(r => sameNameFuzzy(r.nama, t.pic) && projKey(r.project) === projKey(t.project));
    if (cands.length === 0) return;

    const wantDwg = drawingKey(t.taskName);
    const wantRev = toNum(t.rev);
    let pick = cands.filter(r => {
      if (!wantDwg || drawingKey(r.taskName) !== wantDwg) return false;
      const rv = revFromTaskName(r.taskName);
      return rv === null || rv === wantRev;
    });
    const matchedByDwg = pick.length > 0;
    if (!matchedByDwg && pendingCount.get(key) === 1) pick = cands;

    const personCodeKey = (c: string) => `${cleanText(t.pic)}|${cleanText(c)}`;
    const codes = Array.from(new Set(pick.map(r => r.jobcard.replace(/\s+/g, ' ').trim().toUpperCase())))
      .filter(c => matchedByDwg || !taken.has(personCodeKey(c)));
    if (codes.length !== 1) return;

    result.set(t.id, codes[0]);
    taken.add(personCodeKey(codes[0]));
  });
  return result;
}

/* ===== Parser isi kolom "Rencana" (teks dari Google Sheet) menjadi rentang tanggal =====
   Format yang dikenali:
   - tanggal tunggal : "20-10-2026", "2026-10-20", "20 Oktober 2026"
   - rentang          : "20-10-2026 s/d 22-10-2026", "20 Okt 2026 sampai 22 Okt 2026"
   - rentang pendek   : "20-22 Oktober 2026"
   - beberapa tanggal dipisah koma / baris baru
   Format tanggal angka dianggap dd-mm-yyyy. */
const ISO_RE = /^\d{4}-\d{2}-\d{2}$/;

function parseRencanaRanges(raw: string): { start: string; end: string }[] {
  const s = String(raw || '').trim();
  if (!s || s === '-') return [];
  const out: { start: string; end: string }[] = [];
  const mon = '[A-Za-z]{3,9}';

  // a) "20-22 Oktober 2026"
  const rShort = new RegExp(`(\\d{1,2})\\s*(?:-|–|s/d|sd|sampai)\\s*(\\d{1,2})\\s+(${mon})\\s+(\\d{4})`, 'gi');
  for (const m of s.matchAll(rShort)) {
    const a = parseToStandardDate(`${m[1]} ${m[3]} ${m[4]}`);
    const b = parseToStandardDate(`${m[2]} ${m[3]} ${m[4]}`);
    if (ISO_RE.test(a) && ISO_RE.test(b)) out.push({ start: a, end: b });
  }
  const rest = s.replace(rShort, ' ');

  // b) tanggal lengkap (tunggal atau pasangan awal-akhir)
  const dateRe = new RegExp(
    `\\d{4}-\\d{1,2}-\\d{1,2}|\\d{1,2}[/-]\\d{1,2}[/-]\\d{4}|\\d{1,2}\\s+${mon}\\s+\\d{4}`, 'g');
  const dates = (rest.match(dateRe) || []).map(parseToStandardDate).filter(d => ISO_RE.test(d));
  const between = rest.replace(dateRe, '|');
  const isRange = dates.length === 2 && /s\/d|\bsd\b|sampai|hingga|–|\s-\s/i.test(between);
  if (isRange) {
    const [a, b] = dates[0] <= dates[1] ? [dates[0], dates[1]] : [dates[1], dates[0]];
    out.push({ start: a, end: b });
  } else {
    dates.forEach(d => out.push({ start: d, end: d }));
  }
  return out;
}

const isKadep = (jabatan: string) => {
  const j = (jabatan || '').toLowerCase();
  return j.includes('kepala departemen') || /\bkadep\b/.test(j);
};

const deptKey = (s: string) =>
  cleanText(
    (s || '').toLowerCase().replace('&', ' dan ')
      .replace(/\b(kepala|departemen|dept|divisi|dan|kadep|biro)\b/gi, ' ')
  );

function getKadepForDept(all: ParsedMember[], deptName: string): ParsedMember[] {
  const key = deptKey(deptName);
  if (!key) return [];
  const seen = new Set<string>();
  return all.filter(p => {
    if (!isKadep(p.jabatan)) return false;
    if (deptKey(p.dept) !== key && deptKey(p.jabatan) !== key) return false;
    const id = p.nip || p.nama;
    if (seen.has(id)) return false;
    seen.add(id);
    return true;
  });
}

export function CutiSakitPanel({ biroName, members, wb, sisaCuti, title, kadepIds }: {
  biroName: string;
  members: ParsedMember[];
  wb: XLSX.WorkBook | null;
  sisaCuti: SisaCutiData | null;
  title?: string;
  kadepIds?: Set<string>;
}) {
  const entries = useMemo(() => parseTimeOff(wb), [wb]);
  const week = useMemo(() => getWeekRange(), []);
  const today = useMemo(() => isoLocal(new Date()), []);
  const year = String(new Date().getFullYear());

  // Ambil semua key "Rencana" yang unik dari data sisa cuti untuk header tabel
  const rencanaKeys = useMemo(() => {
    if (!sisaCuti) return [];
    const keys = new Set<string>();
    sisaCuti.byName.forEach(info => {
      if (info.rencana) {
        Object.keys(info.rencana).forEach(k => keys.add(k));
      }
    });
    return Array.from(keys).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
    );
  }, [sisaCuti]);

  const data = useMemo(() => {
    const currentYear = new Date().getFullYear();
    const cutoffStart = getTimeOffCutoffStart();
    return members.map(m => {
      const nip = normNip(m.nip), nm = cleanText(m.nama);
      const mine = entries.filter(e =>
        e.status === 'Approved' && ((nip && normNip(e.nip) === nip) || cleanText(e.nama) === nm)
      );
      const inYear = filterLeavesForYear(mine, currentYear);
      const sum = (k: TimeOffEntry['kind']) => inYear.filter(e => e.kind === k).reduce((s, e) => s + e.days, 0);
      const realisasiCuti = sum('cuti');
      const cutiKhusus = inYear.filter(e => e.kind === 'khusus');
      const info: SisaCutiInfo | undefined = findSisaCuti(sisaCuti, m.nama, members.map(x => x.nama));
      const cutiSetelahIsi = info
        ? inYear.filter(e => e.kind === 'cuti')
            .reduce((s, e) => s + countDaysFrom(e, cutoffStart), 0)
        : 0;
      const sisa = info ? info.sisa - cutiSetelahIsi : undefined;
      return {
        m,
        sisa,
        filledAt: info?.filledAt,
        cutiSetelahIsi,
        cuti: realisasiCuti,
        sakit: sum('sakit'),
        sakitKali: inYear.filter(e => e.kind === 'sakit').length,
        panjang: inYear.filter(e => e.kind === 'panjang'),
        khusus: cutiKhusus,
        minggu: mine.filter(e => e.start <= week.to && e.end >= week.from),
        rencana: info?.rencana || {},
      };
    });
  }, [members, entries, week, sisaCuti]);

  const totalCuti = data.reduce((s, d) => s + d.cuti, 0);
  const totalSakit = data.reduce((s, d) => s + d.sakit, 0);
  const orangMinggu = data.filter(d => d.minggu.length > 0);

  const kindLabel = (k: TimeOffEntry['kind']) => {
    if (k === 'sakit') return 'Sakit';
    if (k === 'panjang') return 'Cuti panjang';
    if (k === 'khusus') return 'Cuti khusus';
    return 'Cuti';
  };

  const kindClass = (k: TimeOffEntry['kind']) => {
    if (k === 'sakit') return 'bg-amber-500/10 text-amber-400 border-amber-500/20';
    if (k === 'panjang') return 'bg-violet-500/10 text-violet-300 border-violet-500/20';
    if (k === 'khusus') return 'bg-orange-500/10 text-orange-300 border-orange-500/20';
    return 'bg-cyan-500/10 text-cyan-300 border-cyan-500/20';
  };

  return (
    <div className="space-y-4">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div>
          <h4 className="text-sm font-bold text-white">{title ?? <>Cuti &amp; Sakit Tahun {year}</>}</h4>
          <span className="text-xs text-slate-400">{biroName}</span>
        </div>
      </div>
      {!wb ? (
        <div className="text-[11px] text-amber-400 bg-slate-900 border border-slate-800 rounded-xl p-4">
          Data Time Off belum termuat. Upload file TIME_OFF_DIVISI_DESAIN.xlsx lewat tombol "Upload Time Off" di halaman Departemen Desain.
        </div>
      ) : (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {[
              { label: 'Anggota', value: data.length, color: 'text-blue-400' },
              { label: 'Tidak Hadir Minggu Ini', value: orangMinggu.length, color: 'text-cyan-400' },
              { label: 'Total Hari Cuti', value: totalCuti, color: 'text-emerald-400' },
              { label: 'Total Hari Sakit', value: totalSakit, color: 'text-amber-400' },
            ].map(c => (
              <div key={c.label} className="bg-slate-900 border border-slate-800 rounded-xl px-4 py-3">
                <div className={`text-2xl font-bold ${c.color}`}>{c.value}</div>
                <div className="text-[10px] text-slate-400 uppercase tracking-wide">{c.label}</div>
              </div>
            ))}
          </div>
          <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-3">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <h5 className="text-xs font-bold text-white uppercase tracking-wide">Tidak Masuk Minggu Ini</h5>
              <span className="text-[11px] text-slate-400 font-mono">{formatDisplayDate(week.from)} s/d {formatDisplayDate(week.to)}</span>
            </div>
            {orangMinggu.length > 0 ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                {orangMinggu.flatMap(d => d.minggu.map((e, i) => (
                  <div key={`${d.m.nip}-${i}`} className="flex items-center justify-between gap-3 p-3 bg-slate-950 border border-slate-800 rounded-lg">
                    <div className="min-w-0">
                      <div className="text-sm font-semibold text-white truncate">{d.m.nama}</div>
                      <div className="text-[11px] text-slate-400 truncate">
                        {shortType(e.type)} • {formatDisplayDate(e.start)} s/d {formatDisplayDate(e.end)}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      {e.start <= today && e.end >= today && (
                        <span className="px-2 py-0.5 text-[10px] font-bold rounded border bg-rose-500/10 text-rose-400 border-rose-500/20">Hari ini</span>
                      )}
                      <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${kindClass(e.kind)}`}>{kindLabel(e.kind)}</span>
                    </div>
                  </div>
                )))}
              </div>
            ) : (
              <div className="text-xs text-slate-500 italic text-center py-3">Tidak ada anggota yang cuti minggu ini.</div>
            )}
          </div>
          <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead>
                <tr className="text-slate-400 border-b border-slate-800 uppercase text-[11px]">
                  <th className="py-3 px-3 w-10">No</th>
                  <th className="py-3 px-3">Nama Pegawai</th>
                  <th className="py-3 px-3">Realisasi Cuti {year}</th>
                  <th className="py-3 px-8 text-center text-sky-400 bg-sky-500/5">Sisa Cuti</th>
                  <th className="py-3 px-3 text-center text-amber-400 bg-amber-500/5">Sakit</th>
                  {/* Kolom-kolom Rencana dinamis setelah Sakit */}
                  {rencanaKeys.map(key => (
                    <th key={key} className="py-3 px-3 text-center text-blue-400 bg-blue-500/5 whitespace-nowrap">
                      {key}
                    </th>
                  ))}
                  <th className="py-3 px-3">Minggu Ini</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800">
                {data.length > 0 ? data.map((d, i) => (
                  <tr key={`${d.m.nip}-${i}`} className="hover:bg-slate-800/40">
                    <td className="py-3 px-3 text-slate-500 font-mono">{i + 1}</td>
                    <td className="py-3 px-3 font-semibold text-white">
                      {d.m.nama}
                      {kadepIds?.has(d.m.nip || d.m.nama) && (
                        <span className="ml-2 px-1.5 py-0.5 text-[10px] font-bold rounded border bg-purple-500/10 text-purple-300 border-purple-500/20">Kadep</span>
                      )}
                    </td>
                    <td className="py-3 px-3">
                      <span className="font-mono font-bold text-emerald-400">{d.cuti} hari</span>
                      {d.panjang.length > 0 && (
                        <div className="mt-1 text-[10px] text-violet-300">
                          + {d.panjang.map(e => `${shortType(e.type)} ${e.days} hari`).join(', ')} (tidak dihitung)
                        </div>
                      )}
                      {d.khusus.length > 0 && (
                        <div className="mt-1 text-[10px] text-orange-300">
                          + {d.khusus.map(e => `${shortType(e.type)} ${e.days} hari`).join(', ')} (tidak dihitung)
                        </div>
                      )}
                    </td>
                    <td className="py-3 px-3 text-center bg-sky-500/5">
                      {sisaCuti === null
                        ? <span className="text-slate-600 text-[11px]">memuat...</span>
                        : d.sisa !== undefined
                        ? (
                          <div>
                            <span className="font-mono font-bold text-sky-400">{d.sisa} hari</span>
                            <div className="text-[10px] text-slate-500">
                              isi form {formatDisplayDate(d.filledAt)}{d.cutiSetelahIsi > 0 ? ` • −${d.cutiSetelahIsi}` : ''}
                            </div>
                          </div>
                        )
                        : <span className="text-slate-600" title="Belum mengisi form alokasi cuti">-</span>}
                    </td>
                    <td className="py-3 px-3 text-center bg-amber-500/5">
                      <div className="font-mono font-bold text-amber-400">{d.sakit} hari</div>
                      <div className="text-[10px] text-slate-500">{d.sakitKali}x</div>
                    </td>
                    {/* Isi data Rencana per orang */}
                    {rencanaKeys.map(key => (
                      <td key={key} className="py-3 px-3 text-center bg-blue-500/5 font-mono text-xs text-blue-300 whitespace-nowrap">
                        {d.rencana[key] || '-'}
                      </td>
                    ))}
                    <td className="py-3 px-3">
                      {d.minggu.length > 0
                        ? <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${kindClass(d.minggu[0].kind)}`}>{kindLabel(d.minggu[0].kind)}</span>
                        : <span className="text-slate-600">-</span>}
                    </td>
                  </tr>
                )) : (
                  <tr><td colSpan={6 + rencanaKeys.length} className="py-8 text-center text-slate-500">Tidak ada anggota organik di biro ini.</td></tr>
                )}
              </tbody>
            </table>
          </div>
          <div className="text-[10px] text-slate-500">
            Dihitung dari status Approved. Durasi = hari kerja.{' '}
            <b className="text-violet-300">Cuti panjang</b> (melahirkan, kewajiban negara, pensiun) dan{' '}
            <b className="text-orange-300">cuti khusus</b> (umroh, menikah, istri melahirkan, keluarga inti sakit/rawat inap)
            tidak masuk realisasi cuti.{' '}
            Sisa Cuti diambil dari Google Sheet "Form Pengisian Alokasi Cuti 2026" (isian terakhir bila nama ganda),
            lalu dikurangi cuti biasa pada file Time Off mulai dari <b>awal bulan berjalan</b> (atau mulai dari <b>minggu kemarin</b> jika saat ini berada di minggu awal bulan). Tanda "-" berarti belum mengisi form.
          </div>
        </>
      )}
    </div>
  );
}

// ===== Panel "Tidak Masuk Minggu Ini" per Departemen =====
// Sumber data:
//   1) Time Off Approved dari TIME_OFF_DIVISI_DESAIN.xlsx
//   2) Rencana cuti (Rencana 1, Rencana 2, ...) dari Google Sheet Form Alokasi Cuti
function DeptWeeklyTimeOffPanel({ deptName, members, wb, sisaCuti }: {
  deptName: string;
  members: ParsedMember[];
  wb: XLSX.WorkBook | null;
  sisaCuti: SisaCutiData | null;
}) {
  const entries = useMemo(() => parseTimeOff(wb), [wb]);
  const week = useMemo(() => getWeekRange(), []);
  const today = useMemo(() => isoLocal(new Date()), []);

  type WeekItem = {
    key: string;
    nama: string;
    label: string;
    start: string;
    end: string;
    kind: TimeOffEntry['kind'] | 'rencana';
  };

  const items = useMemo<WeekItem[]>(() => {
    const allNames = members.map(x => x.nama);
    const list: WeekItem[] = [];
    members.forEach(m => {
      const nip = normNip(m.nip), nm = cleanText(m.nama);
      const mine = entries.filter(e =>
        e.status === 'Approved' && ((nip && normNip(e.nip) === nip) || cleanText(e.nama) === nm)
      );

      // 1) Time Off Approved yang beririsan dengan minggu ini
      mine
        .filter(e => e.start <= week.to && e.end >= week.from)
        .forEach((e, i) => list.push({
          key: `to-${m.nip || m.nama}-${i}`,
          nama: m.nama,
          label: shortType(e.type),
          start: e.start,
          end: e.end,
          kind: e.kind,
        }));

      // 2) Rencana dari Google Sheet
      const info = findSisaCuti(sisaCuti, m.nama, allNames);
      if (!info?.rencana) return;
      Object.entries(info.rencana).forEach(([rKey, rVal]) => {
        parseRencanaRanges(rVal).forEach((r, j) => {
          if (!(r.start <= week.to && r.end >= week.from)) return;
          // Lewati jika sudah tercatat di Time Off Approved agar tidak dobel
          if (mine.some(e => e.start <= r.end && e.end >= r.start)) return;
          list.push({
            key: `rc-${m.nip || m.nama}-${rKey}-${j}`,
            nama: m.nama,
            label: rKey,
            start: r.start,
            end: r.end,
            kind: 'rencana',
          });
        });
      });
    });
    return list;
  }, [members, entries, week, sisaCuti]);

  const kindLabel = (k: WeekItem['kind']) => {
    if (k === 'sakit') return 'Sakit';
    if (k === 'panjang') return 'Cuti panjang';
    if (k === 'khusus') return 'Cuti khusus';
    if (k === 'rencana') return 'Rencana';
    return 'Cuti';
  };

  const kindClass = (k: WeekItem['kind']) => {
    if (k === 'sakit') return 'bg-amber-500/10 text-amber-400 border-amber-500/20';
    if (k === 'panjang') return 'bg-violet-500/10 text-violet-300 border-violet-500/20';
    if (k === 'khusus') return 'bg-orange-500/10 text-orange-300 border-orange-500/20';
    if (k === 'rencana') return 'bg-blue-500/10 text-blue-300 border-blue-500/30 border-dashed';
    return 'bg-cyan-500/10 text-cyan-300 border-cyan-500/20';
  };

  if (!wb && !sisaCuti) return null;

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-3 mb-4">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <h5 className="text-xs font-bold text-white uppercase tracking-wide flex items-center gap-2">
          <Clock className="w-3.5 h-3.5 text-cyan-400" />
          Tidak Masuk Minggu Ini ({deptName})
        </h5>
        <span className="text-[11px] text-slate-400 font-mono">{formatDisplayDate(week.from)} s/d {formatDisplayDate(week.to)}</span>
      </div>
      {items.length > 0 ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2">
          {items.map(it => (
            <div key={it.key} className="flex items-center justify-between gap-3 p-3 bg-slate-950 border border-slate-800 rounded-lg">
              <div className="min-w-0">
                <div className="text-sm font-semibold text-white truncate">{it.nama}</div>
                <div className="text-[11px] text-slate-400 break-words">
                  {it.label} • {formatDisplayDate(it.start)} s/d {formatDisplayDate(it.end)}
                </div>
              </div>
              <div className="flex items-center gap-1.5 shrink-0">
                {it.start <= today && it.end >= today && (
                  <span className="px-2 py-0.5 text-[10px] font-bold rounded border bg-rose-500/10 text-rose-400 border-rose-500/20">Hari ini</span>
                )}
                <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${kindClass(it.kind)}`}>{kindLabel(it.kind)}</span>
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className="text-xs text-slate-500 italic text-center py-3">Tidak ada anggota departemen ini yang cuti minggu ini.</div>
      )}
      <div className="text-[10px] text-slate-500">
        <b className="text-blue-300">Rencana</b> = dari Google Sheet Form Alokasi Cuti (belum tentu sudah disetujui di Time Off).
      </div>
    </div>
  );
}

export default function RendalPage({ user, onLogout }: RendalPageProps) {
  const [accessMode, setAccessMode] = useState<'landing' | 'organik' | 'subkon'>('landing');
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  const [formPageMode, setFormPageMode] = useState<'members' | 'form' | 'absensi'>('members');
  const [absensiMonth, setAbsensiMonth] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [subconSelectedDept, setSubconSelectedDept] = useState<Department | null>(null);
  const [subconSelectedBiro, setSubconSelectedBiro] = useState<string | null>(null);
  const [subconPageMode, setSubconPageMode] = useState<'members' | 'form' | 'release'>('members');
  const [subconSearch, setSubconSearch] = useState('');
  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [kpiWorkbook, setKpiWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [timeOffWorkbook, setTimeOffWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>({});
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});
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
  const [formData, setFormData] = useState({
    nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '',
    pic: '', jo: '', rev: '0', realJo: '', release: '',
  });
  const [isPlannerOpen, setIsPlannerOpen] = useState(false);
  const [showKpi, setShowKpi] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showRekapDivisi, setShowRekapDivisi] = useState(user.role === 'rekap');
  const [editingReleaseId, setEditingReleaseId] = useState<string | null>(null);
  const [editingReleaseVal, setEditingReleaseVal] = useState<string>('');
  const [drawingControlMap, setDrawingControlMap] = useState<Record<string, DrawingControlRow[]>>({});
  const [activeDrawingSheetTitle, setActiveDrawingSheetTitle] = useState<string>('');
  const [isFetchingDrawing, setIsFetchingDrawing] = useState<boolean>(false);
  const [hourPlannerMap, setHourPlannerMap] = useState<HourPlannerMap>({});
  const [isImportingHour, setIsImportingHour] = useState(false);
  const [isImportingUnified, setIsImportingUnified] = useState(false);
  const [sisaCuti, setSisaCuti] = useState<SisaCutiData | null>(null);

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
          const persistedRev = (row.rev !== undefined && row.rev !== null && String(row.rev).trim() !== '') ? String(row.rev) : getLocalRev(row.id, '0');
          const persistedRelease = (row.release !== undefined && row.release !== null && String(row.release).trim() !== '') ? String(row.release) : getLocalRelease(row.id, '');
          grouped[biroKey].push({
            id: row.id, biroName: row.biro_name || '', project: row.project || '', taskName: row.task_name || '',
            startDate: row.start_date || '', endDate: row.end_date || '', pic: row.pic || '', jo: row.jo || '',
            kodeJc: row.kode_jc || '', rev: persistedRev, release: persistedRelease, realJo: row.real_jc || row.realJo || '',
            status: row.status || '',
          });
        });
        setManualTasks(grouped);
      }
    } catch {}
  }, []);

  useEffect(() => {
    async function initMasterFiles() {
      try {
        let im4Url = '';
        let realisasiUrl = '';
        let kpiUrl = '';
        let timeOffUrl = '';
        Object.entries(excelGlobUrls).forEach(([path, url]) => {
          const pLower = path.toLowerCase();
          if (pLower.includes('kpi')) kpiUrl = url;
          else if (pLower.includes('im4') || pLower.includes('drawing') || pLower.includes('akses')) im4Url = url;
          else if (pLower.includes('realisasi')) realisasiUrl = url;
        });
        Object.entries(parentExcelGlobUrls).forEach(([path, url]) => {
          const p = path.toLowerCase();
          if (p.includes('kpi')) kpiUrl = url;
          else if (p.includes('time_off') || p.includes('timeoff')) timeOffUrl = url;
        });
        const [wbJc, wbIm4, wbRealisasi] = await Promise.all([
          loadMasterFile('jobcard'),
          loadMasterFile('im4').then(wb => wb || fetchSafeWorkbook([im4Url, '/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx', './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx'])),
          loadMasterFile('realisasi').then(wb => wb || fetchSafeWorkbook([realisasiUrl, '/Realisasi JO.xlsx', './Realisasi JO.xlsx'])),
        ]);
        setHourPlannerMap(await loadHourPlanner());
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
        if (wbRealisasi) setRealisasiWorkbook(wbRealisasi);
        const wbKpi = await fetchSafeWorkbook([kpiUrl, '/data_kpi.xlsx', './data_kpi.xlsx']);
        if (wbKpi) setKpiWorkbook(wbKpi);
        const wbTimeOff = await loadMasterFile('timeoff')
          .then(wb => wb || fetchSafeWorkbook([timeOffUrl, '/TIME_OFF_DIVISI_DESAIN.xlsx']));
        if (wbTimeOff) setTimeOffWorkbook(wbTimeOff);
      } catch {}
    }
    initMasterFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

  const refreshSisaCuti = useCallback(async () => {
    const data = await fetchSisaCuti();
    if (data) setSisaCuti(data);
    return data;
  }, []);

  useEffect(() => {
    let alive = true;
    const load = async () => {
      const data = await fetchSisaCuti();
      if (alive && data) setSisaCuti(data);
    };
    load();
    const timer = setInterval(load, 2 * 60 * 1000);
    return () => { alive = false; clearInterval(timer); };
  }, []);

  const autoTried = useRef<Set<string>>(new Set());
  const autoRunning = useRef(false);
  const autoPending = useRef(false);
  const [autoSyncMsg, setAutoSyncMsg] = useState('');

  const autoSyncReleases = useCallback(async (retryAll: boolean) => {
    if (autoRunning.current) {
      autoPending.current = true;
      return;
    }
    const all = ([] as TaskItem[]).concat(...Object.values(manualTasks));
    const todo = all.filter(t =>
      t.project && t.taskName &&
      !/^\d{4}-\d{2}-\d{2}$/.test(parseToStandardDate(t.release)) &&
      (retryAll || !autoTried.current.has(`${t.id}|${t.rev || '0'}`)));
    if (todo.length === 0) return;
    autoRunning.current = true;
    let totalFound = 0;
    try {
      const projects = Array.from(new Set(todo.map(t => t.project))).sort((a, b) =>
        (isKnownDriveProject(b) ? 1 : 0) - (isKnownDriveProject(a) ? 1 : 0)
      );
      for (let i = 0; i < projects.length; i += 4) {
        const batchProjs = projects.slice(i, i + 4);
        const rowsByProject = new Map<string, DrawingControlRow[]>();
        await Promise.all(batchProjs.map(async p => {
          try { rowsByProject.set(p, await fetchDrawingRowsFor(p)); } catch { rowsByProject.set(p, []); }
        }));
        const batchFound = new Map<string, string>();
        todo.filter(t => batchProjs.includes(t.project)).forEach(t => {
          autoTried.current.add(`${t.id}|${t.rev || '0'}`);
          const rel = findReleaseInRows(rowsByProject.get(t.project) || [], t.taskName, t.rev || '0');
          if (rel) batchFound.set(t.id, rel);
        });
        if (batchFound.size > 0) {
          totalFound += batchFound.size;
          const ids = Array.from(batchFound.keys());
          for (let j = 0; j < ids.length; j += 10) {
            await Promise.all(ids.slice(j, j + 10).map(async id => {
              const rel = batchFound.get(id)!;
              saveLocalRelease(id, rel);
              try { await supabase.from('job_cards').update({ release: rel }).eq('id', id); } catch {}
            }));
          }
          setManualTasks(prev => {
            const updated = { ...prev };
            Object.keys(updated).forEach(k => {
              updated[k] = updated[k].map(t => batchFound.has(t.id) ? { ...t, release: batchFound.get(t.id)! } : t);
            });
            return updated;
          });
        }
      }
      if (totalFound > 0) {
        setAutoSyncMsg(`${totalFound} tanggal Release terisi otomatis dari Drawing Control`);
        setTimeout(() => setAutoSyncMsg(''), 8000);
      }
    } finally {
      autoRunning.current = false;
      if (autoPending.current) {
        autoPending.current = false;
        setTimeout(() => autoSyncReleases(false), 150);
      }
    }
  }, [manualTasks]);

  useEffect(() => { autoSyncReleases(false); }, [autoSyncReleases]);
  useEffect(() => {
    const timer = setInterval(() => autoSyncReleases(true), 10 * 60 * 1000);
    return () => clearInterval(timer);
  }, [autoSyncReleases]);

  const handleUpdateIm4Excel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        setIm4Workbook(XLSX.read(evt.target?.result, { type: 'binary' }));
        const err = await saveMasterFile('im4', file);
        alert(err
          ? `File terbaca, tapi gagal disimpan ke server: ${err.message}`
          : `Master Personel "${file.name}" berhasil diperbarui dan tersimpan!`);
      } catch { alert('Gagal membaca file Excel Personel IM4.'); }
    };
    reader.readAsBinaryString(file);
  };

  const handleUpdateJobcardExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        setJobcardWorkbook(wb);
        const err = await saveMasterFile('jobcard', file);
        if (err) {
          alert(`File "${file.name}" terbaca, tapi gagal disimpan ke server: ${err.message}`);
        } else {
          alert(`Katalog Jobcard "${file.name}" berhasil diperbarui dan tersimpan permanen di server!`);
        }
      }
      catch { alert('Gagal membaca file JOBCARD_DESAIN.xlsx.'); }
    };
    reader.readAsBinaryString(file);
  };

  const handleUploadKpiExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const file = input.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        setKpiWorkbook(XLSX.read(evt.target?.result, { type: 'binary' }));
        alert(`File "${file.name}" berhasil dibaca untuk Rekap Absensi.`);
      } catch { alert('Gagal membaca file data_kpi.xlsx.'); }
      input.value = '';
    };
    reader.readAsBinaryString(file);
  };

  const handleUploadTimeOffExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const file = input.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        setTimeOffWorkbook(XLSX.read(evt.target?.result, { type: 'binary' }));
        const err = await saveMasterFile('timeoff', file);
        alert(err
          ? `File terbaca, tapi gagal disimpan ke server: ${err.message}`
          : `Data Time Off "${file.name}" berhasil diperbarui dan tersimpan!`);
      } catch { alert('Gagal membaca file Time Off.'); }
      input.value = '';
    };
    reader.readAsBinaryString(file);
  };

  const handleImportHourPlanner = async (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const file = input.files?.[0]; if (!file) return;
    setIsImportingHour(true);
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const { entries, rowCount, error } = parseHourPlannerWorkbook(wb);
      if (error) { alert(`Gagal membaca "${file.name}": ${error}`); return; }
      if (rowCount === 0) { alert(`Tidak ada baris Hour Planner yang valid di "${file.name}".`); return; }
      const merged: HourPlannerMap = { ...(await loadHourPlanner()), ...entries };
      const saveErr = await saveHourPlanner(merged);
      setHourPlannerMap(merged);
      const allTasks = ([] as TaskItem[]).concat(...Object.values(manualTasks));
      const toUpdate = allTasks.filter(t => t.kodeJc && entries[cleanText(t.kodeJc)]);
      let okCount = 0, failCount = 0;
      const newJo = new Map<string, string>();
      for (let i = 0; i < toUpdate.length; i += 10) {
        await Promise.all(toUpdate.slice(i, i + 10).map(async t => {
          const hour = String(entries[cleanText(t.kodeJc)].hour);
          const { error: upErr } = await supabase.from('job_cards').update({ jo: hour }).eq('id', t.id);
          if (upErr) failCount++;
          else { okCount++; newJo.set(t.id, hour); }
        }));
      }
      if (newJo.size > 0) {
        setManualTasks(prev => {
          const updated = { ...prev };
          Object.keys(updated).forEach(k => {
            updated[k] = updated[k].map(t => newJo.has(t.id) ? { ...t, jo: newJo.get(t.id)! } : t);
          });
          return updated;
        });
      }
      const linkedKeys = new Set(toUpdate.map(t => cleanText(t.kodeJc)));
      const noJobcard = Object.keys(entries).filter(k => !linkedKeys.has(k)).length;
      alert(
        `Hour Planner "${file.name}" diimpor: ${rowCount} baris.\n` +
        `• Plan JO terisi pada job card yang sudah ada: ${okCount}\n` +
        `• Belum ada job card dengan kode itu: ${noJobcard} (tetap disimpan, Plan JO terisi otomatis saat kode Jobcard diisi di Planner)` +
        (failCount ? `\nPERHATIAN: ${failCount} job card gagal diperbarui Plan JO-nya (cek tipe kolom jo di Supabase).` : '') +
        (saveErr ? `\nPERHATIAN: gagal disimpan permanen ke server (${saveErr.message}). Data hanya ada di sesi ini.` : '')
      );
    } catch {
      alert('Gagal membaca file Excel Hour Planner.');
    } finally {
      setIsImportingHour(false);
      input.value = '';
    }
  };

  // Impor 1 file terpadu (Hour Planner + Realisasi JO + Jobcard + Deskripsi):
  // 1) Plan JO diambil dari nilai tertinggi (MAX Planned Hour) per Jobcard
  // 2) Semua baris tetap dimasukkan untuk Real JO (Effective + Overtime)
  // 3) Plan Start / Plan Finish diabaikan dulu (kosong)
  // Otomatis menghapus data lama di job_cards & menggantinya dengan data baru.
  const handleImportUnifiedData = async (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const file = input.files?.[0];
    if (!file) return;
    if (!window.confirm(`Apakah Anda yakin ingin mengimpor "${file.name}"?\n\nSemua data penugasan/Jobcard lama akan DIHAPUS otomatis dan digantikan dengan data baru dari file ini.`)) {
      input.value = '';
      return;
    }
    setIsImportingUnified(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });

      const allMasterBiros = (departmentsData || []).flatMap(d => d.biros.map(b => b.name));
      const resolveBiroName = (rawWc: string, memberBiro?: string): string => {
        const stripWc = (s: string) => s.replace(/^[0-9A-Z]{3,10}\s*-\s*/i, '').trim();
        const candidates = [memberBiro || '', stripWc(rawWc || ''), rawWc || ''].filter(Boolean);
        for (const cand of candidates) {
          const exact = allMasterBiros.find(b => cleanText(b) === cleanText(cand));
          if (exact) return exact;
          const fuzzy = allMasterBiros.find(b => isBiroMatch(b, cand));
          if (fuzzy) return fuzzy;
        }
        return candidates[0] || 'Tanpa Biro';
      };

      const resolveIm4Member = (nip: string, rawName: string): ParsedMember | undefined => {
        const nNip = normNip(nip);
        if (nNip) {
          const byNip = allParsedFromExcel.find(m => normNip(m.nip) === nNip);
          if (byNip) return byNip;
        }
        const cName = cleanText(rawName);
        if (!cName) return undefined;
        const exact = allParsedFromExcel.find(m => cleanText(m.nama) === cName);
        if (exact) return exact;
        const fuzzy = allParsedFromExcel.filter(m => sameNameFuzzy(m.nama, rawName));
        return fuzzy.length === 1 ? fuzzy[0] : undefined;
      };

      interface UnifiedTaskGroup {
        biroName: string;
        pic: string;
        project: string;
        taskName: string;
        kodeJc: string;
        startDate: string;
        maxPlanJo: number;
        hasPlanJo: boolean;
        realJo: number;
        rev: string;
      }

      const dateFromJobcardCode = (jc: string): string => {
        const m = (jc || '').match(/^JC(\d{2})(\d{2})(\d{2})\b/i);
        if (!m) return '';
        return `20${m[3]}-${m[2]}-${m[1]}`;
      };

      const taskGroups = new Map<string, UnifiedTaskGroup>();
      const newHpMap: HourPlannerMap = {};
      let totalRowsRead = 0;

      wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        if (!sheet) return;
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (rows.length === 0) return;

        let headerIdx = -1;
        let cDate = -1, cUser = -1, cJc = -1, cPlan = -1, cEff = -1, cOt = -1, cQty = -1, cProj = -1, cDesc = -1, cWc = -1;
        for (let r = 0; r < Math.min(10, rows.length); r++) {
          const h = (rows[r] || []).map(v => String(v).trim().toLowerCase());
          const jc = h.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
          const us = h.findIndex(c => c === 'user' || c === 'employee/name' || c === 'nama' || c === 'name' || c === 'employee' || c === 'pegawai');
          if (jc !== -1 && us !== -1) {
            headerIdx = r;
            cJc = jc;
            cUser = us;
            cDate = h.findIndex(c => c === 'date' || c === 'tanggal' || c.includes('start'));
            cPlan = h.findIndex(c => (c.includes('plan') && c.includes('hour')) || c === 'plan jo' || c === 'planned hour');
            cEff = h.findIndex(c => c.includes('effective'));
            cOt = h.findIndex(c => c.includes('overtime'));
            cQty = h.findIndex(c => c === 'quantity' || c === 'qty');
            cProj = h.findIndex(c => c === 'project' || c === 'project/code' || c === 'proyek');
            cDesc = h.findIndex(c => c === 'description' || c === 'deskripsi' || c === 'jobcard/task name' || c === 'task name');
            cWc = h.findIndex(c => c === 'work center' || c === 'biro' || c.includes('work center'));
            break;
          }
        }
        if (headerIdx === -1) return;

        for (let r = headerIdx + 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row) continue;
          const userParsed = splitUserCell(row[cUser]);
          if (!userParsed.nama) continue;
          totalRowsRead++;

          const jcRaw = String(row[cJc] ?? '').replace(/\s+/g, ' ').trim();
          const kodeJc = (!jcRaw || jcRaw.toLowerCase() === 'nan' || jcRaw.toLowerCase().includes('jobcard')) ? '' : jcRaw.toUpperCase();
          if (!kodeJc) continue; // baris tanpa kode Jobcard tetap dihitung di ringkasan "Realisasi tanpa Jobcard"

          const matchedMember = resolveIm4Member(userParsed.nip, userParsed.nama);
          const pic = matchedMember ? matchedMember.nama : userParsed.nama;
          const rawWc = cWc !== -1 ? String(row[cWc] ?? '').trim() : '';
          const biroName = resolveBiroName(rawWc, matchedMember?.biro);

          const project = cProj !== -1 ? extractProjectCodeFromCell(row[cProj]) : '';
          const desc = cDesc !== -1 ? normalizeDrawingSpacing(String(row[cDesc] ?? '')) : '';
          const detectedRev = revFromTaskName(desc);
          const parsedRowDate = cDate !== -1 ? parseToStandardDate(row[cDate]) : '';
          const rowDate = /^\d{4}-\d{2}-\d{2}$/.test(parsedRowDate) ? parsedRowDate : dateFromJobcardCode(kodeJc);

          const planValRaw = cPlan !== -1 ? String(row[cPlan] ?? '').trim().replace(',', '.') : '';
          const planNum = planValRaw !== '' ? parseFloat(planValRaw) : NaN;
          const hasValidPlan = !isNaN(planNum);

          const eff = cEff !== -1 ? absNum(row[cEff]) : 0;
          const ot = cOt !== -1 ? absNum(row[cOt]) : 0;
          const qty = cQty !== -1 ? absNum(row[cQty]) : 0;
          const hours = (eff + ot) > 0 ? (eff + ot) : qty;

          // Simpan nilai tertinggi Planned Hour ke HourPlannerMap per kode Jobcard
          const jcKey = cleanText(kodeJc);
          if (jcKey && hasValidPlan) {
            const prevHp = newHpMap[jcKey];
            const bestStart = prevHp?.start && rowDate ? (rowDate < prevHp.start ? rowDate : prevHp.start) : (rowDate || prevHp?.start || '');
            if (!prevHp || planNum > prevHp.hour) {
              newHpMap[jcKey] = {
                kode: kodeJc,
                hour: planNum,
                planner: prevHp?.planner || '',
                start: bestStart,
                end: '',
              };
            } else if (bestStart && bestStart !== prevHp.start) {
              prevHp.start = bestStart;
            }
          }

          // Kelompokkan per (PIC, Jobcard)
          const groupKey = `${cleanText(pic)}|${jcKey}`;
          const existing = taskGroups.get(groupKey);
          if (!existing) {
            taskGroups.set(groupKey, {
              biroName,
              pic,
              project,
              taskName: desc,
              kodeJc,
              startDate: rowDate,
              maxPlanJo: hasValidPlan ? planNum : 0,
              hasPlanJo: hasValidPlan,
              realJo: hours,
              rev: detectedRev !== null ? String(detectedRev) : '0',
            });
          } else {
            if (rowDate && (!existing.startDate || rowDate < existing.startDate)) {
              existing.startDate = rowDate;
            }
            if (hasValidPlan && (!existing.hasPlanJo || planNum > existing.maxPlanJo)) {
              existing.maxPlanJo = planNum;
              existing.hasPlanJo = true;
            }
            existing.realJo += hours;
            if (!existing.project && project) existing.project = project;
            if (desc && (!existing.taskName || (!extractDrawingNumber(existing.taskName) && extractDrawingNumber(desc)))) {
              existing.taskName = desc;
            }
            if (detectedRev !== null) {
              existing.rev = String(detectedRev);
            }
          }
        }
      });

      if (totalRowsRead === 0 || taskGroups.size === 0) {
        alert(`Tidak ditemukan baris data Jobcard terpadu yang valid di "${file.name}". Pastikan terdapat kolom User, Jobcard, Planned Hour, Effective Hours, Overtime Hours, Project, dan Description.`);
        return;
      }

      // 1. Baca data lama di tabel job_cards (simpan release yang sudah pernah terisi agar tidak hilang) lalu hapus
      const { data: oldRows, error: fetchOldErr } = await supabase.from('job_cards').select('*');
      if (fetchOldErr) {
        alert('Gagal membaca data lama untuk dihapus: ' + fetchOldErr.message);
        return;
      }
      const prevReleaseByJc = new Map<string, string>();
      const prevReleaseByDwg = new Map<string, string>();
      (oldRows || []).forEach((r: any) => {
        const rel = parseToStandardDate(r.release || getLocalRelease(r.id, ''));
        if (!/^\d{4}-\d{2}-\d{2}$/.test(rel)) return;
        const jcKey = cleanText(r.kode_jc || '');
        if (jcKey) prevReleaseByJc.set(jcKey, rel);
        const dKey = drawingKey(r.task_name || '');
        const pKey = cleanText(r.project || '');
        const rv = String(r.rev ?? '0').trim() || '0';
        if (pKey && dKey) prevReleaseByDwg.set(`${pKey}|${dKey}|${rv}`, rel);
      });

      // Ambil juga Release dari Drawing Control (Google Drive) untuk proyek yang ada
      const uniqueProjs = Array.from(new Set(Array.from(taskGroups.values()).map(g => g.project).filter(Boolean))).sort((a, b) =>
        (isKnownDriveProject(b) ? 1 : 0) - (isKnownDriveProject(a) ? 1 : 0)
      );
      const driveRowsByProj = new Map<string, DrawingControlRow[]>();
      for (let i = 0; i < uniqueProjs.length; i += 4) {
        await Promise.all(uniqueProjs.slice(i, i + 4).map(async p => {
          try { driveRowsByProj.set(p, await fetchDrawingRowsFor(p)); } catch { driveRowsByProj.set(p, []); }
        }));
      }

      const oldIds = (oldRows || []).map((r: any) => r.id).filter(Boolean);
      for (let i = 0; i < oldIds.length; i += 100) {
        const chunk = oldIds.slice(i, i + 100);
        const { error: delErr } = await supabase.from('job_cards').delete().in('id', chunk);
        if (delErr) {
          alert('Gagal menghapus data lama: ' + delErr.message);
          return;
        }
      }
      try {
        localStorage.removeItem('task_rev_map');
        localStorage.removeItem('task_release_map');
      } catch {}

      // 2. Petakan biro_id dari tabel biros di Supabase
      const { data: biroList } = await supabase.from('biros').select('id, name');
      const getBiroId = (bName: string): string | null => {
        if (!biroList || biroList.length === 0) return null;
        const found = biroList.find((b: any) => isBiroMatch(b.name, bName));
        return found ? found.id : biroList[0].id;
      };

      // 3. Insert data baru ke job_cards (beserta tanggal Release dari Drive / data sebelumnya)
      const payloads = Array.from(taskGroups.values()).map(g => {
        const joStr = g.hasPlanJo && g.maxPlanJo > 0 ? String(fmtNum(g.maxPlanJo)) : (g.hasPlanJo ? String(fmtNum(g.maxPlanJo)) : '');
        const driveRel = findReleaseInRows(driveRowsByProj.get(g.project) || [], g.taskName, g.rev);
        const fallbackRel = prevReleaseByJc.get(cleanText(g.kodeJc)) || prevReleaseByDwg.get(`${cleanText(g.project)}|${drawingKey(g.taskName)}|${g.rev}`) || '';
        const releaseVal = driveRel || fallbackRel;
        return {
          biro_id: getBiroId(g.biroName),
          biro_name: g.biroName,
          personil_name: g.pic,
          pic: g.pic,
          project_code: g.project,
          project: g.project,
          task_name: g.taskName,
          start_date: g.startDate,
          end_date: '',
          jo: joStr,
          kode_jc: g.kodeJc,
          status: 'approved',
          rev: g.rev,
          release: releaseVal,
        };
      });

      let insertedCount = 0;
      for (let i = 0; i < payloads.length; i += 50) {
        const batch = payloads.slice(i, i + 50);
        let { data: insRows, error: insErr } = await supabase.from('job_cards').insert(batch).select('id, kode_jc, pic');
        if (insErr && (insErr.message?.includes('rev') || insErr.message?.includes('release') || (insErr as any).details?.includes('rev') || (insErr as any).details?.includes('release'))) {
          const stripped = batch.map(({ rev, release, ...rest }) => rest);
          const retry = await supabase.from('job_cards').insert(stripped).select('id, kode_jc, pic');
          insRows = retry.data;
          insErr = retry.error;
        }
        if (insErr) {
          alert(`Gagal menyimpan sebagian data baru (batch ${Math.floor(i / 50) + 1}): ${insErr.message}`);
          break;
        }
        (insRows || []).forEach((ir: any, idx: number) => {
          const rel = batch[idx]?.release;
          if (ir?.id && rel) saveLocalRelease(ir.id, rel);
        });
        insertedCount += batch.length;
      }

      // 4. Simpan Hour Planner baru & file Realisasi baru (menggantikan yang lama)
      setHourPlannerMap(newHpMap);
      await saveHourPlanner(newHpMap);

      setRealisasiWorkbook(wb);
      const xlsxOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const realisasiFile = new File([xlsxOut], 'realisasi.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      await saveMasterFile('realisasi', realisasiFile);

      // 5. Muat ulang daftar job_cards
      autoTried.current.clear();
      await loadAllJobCards();

      alert(
        `Data Terpadu "${file.name}" berhasil diimpor!\n\n` +
        `• Data lama dihapus: ${oldIds.length} tugas\n` +
        `• Tugas Jobcard baru tersimpan: ${insertedCount} tugas\n` +
        `• Plan JO (nilai tertinggi) tercatat untuk ${Object.keys(newHpMap).length} kode Jobcard\n` +
        `• Real JO langsung terhitung dari seluruh baris (${totalRowsRead} baris log harian)`
      );
    } catch (err: any) {
      alert('Gagal membaca atau memproses file Data Terpadu: ' + (err?.message || 'Error tidak diketahui'));
    } finally {
      setIsImportingUnified(false);
      input.value = '';
    }
  };

  // Isi kode Jobcard otomatis pada tugas Form yang masih kosong (nama + proyek + nomor gambar).
  const autoAssignJobcardsFromRealisasi = useCallback(async (wb: XLSX.WorkBook): Promise<number> => {
    const rows = parseRealisasiRows(wb);
    const all = ([] as TaskItem[]).concat(...Object.values(manualTasks));
    const pending = all.filter(t => !(t.kodeJc || '').trim() && t.pic && t.project);
    if (pending.length === 0) return 0;
    const used = new Set(
      all
        .filter(t => (t.kodeJc || '').trim() && t.pic)
        .map(t => `${cleanText(t.pic)}|${cleanText(t.kodeJc || '')}`)
    );
    const matched = matchJobcardsToTasks(rows, pending, used);
    if (matched.size === 0) return 0;

    const done = new Map<string, { kode: string; jo?: string }>();
    const ids = Array.from(matched.keys());
    for (let i = 0; i < ids.length; i += 10) {
      await Promise.all(ids.slice(i, i + 10).map(async id => {
        const kode = matched.get(id)!;
        const hp = hourPlannerMap[cleanText(kode)];
        const payload: any = { kode_jc: kode, status: 'approved' };
        if (hp) payload.jo = String(hp.hour); // sama seperti saat Planner menyimpan kode
        const { error } = await supabase.from('job_cards').update(payload).eq('id', id);
        if (!error) done.set(id, { kode, jo: hp ? String(hp.hour) : undefined });
      }));
    }
    if (done.size > 0) {
      setManualTasks(prev => {
        const updated = { ...prev };
        Object.keys(updated).forEach(k => {
          updated[k] = updated[k].map(t => {
            const d = done.get(t.id);
            return d ? { ...t, kodeJc: d.kode, status: 'approved', ...(d.jo !== undefined ? { jo: d.jo } : {}) } : t;
          });
        });
        return updated;
      });
      loadAllJobCards();
    }
    return done.size;
  }, [manualTasks, hourPlannerMap, loadAllJobCards]);

  const autoJcTriedRef = useRef<Set<string>>(new Set());
  const autoJcRunningRef = useRef(false);
  useEffect(() => {
    autoJcTriedRef.current.clear();
  }, [realisasiWorkbook]);

  useEffect(() => {
    if (!realisasiWorkbook || autoJcRunningRef.current) return;
    const all = ([] as TaskItem[]).concat(...Object.values(manualTasks));
    const untriedPending = all.filter(
      t => !(t.kodeJc || '').trim() && t.pic && t.project && !autoJcTriedRef.current.has(t.id)
    );
    if (untriedPending.length === 0) return;
    untriedPending.forEach(t => autoJcTriedRef.current.add(t.id));
    autoJcRunningRef.current = true;
    autoAssignJobcardsFromRealisasi(realisasiWorkbook).finally(() => {
      autoJcRunningRef.current = false;
    });
  }, [realisasiWorkbook, manualTasks, autoAssignJobcardsFromRealisasi]);

  const handleManualUploadRealisasi = (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.target;
    const file = input.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        setRealisasiWorkbook(wb);
        const err = await saveMasterFile('realisasi', file);
        const filled = await autoAssignJobcardsFromRealisasi(wb);
        const info = filled > 0 ? `\n${filled} tugas terisi kode Jobcard otomatis dari file ini.` : '';
        if (err) {
          alert(`File "${file.name}" terbaca, tapi gagal disimpan ke server: ${err.message}${info}`);
        } else {
          alert(`File "${file.name}" berhasil diperbarui dan tersimpan permanen! Real JO langsung terhitung otomatis.${info}`);
        }
      }
      catch { alert('Gagal membaca file Realisasi JO.xlsx.'); }
      input.value = '';
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
    try {
      realisasiWorkbook.SheetNames.forEach(sheetName => {
        const sheet = realisasiWorkbook.Sheets[sheetName]; if (!sheet) return;
        const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }); if (rawRows.length === 0) return;
        let headerIdx = -1, jobcardCol = -1, effCol = -1, otCol = -1;
        for (let r = 0; r < Math.min(5, rawRows.length); r++) {
          const row = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
          let jcIdx = row.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
          if (jcIdx === -1) jcIdx = row.findIndex(c => c.includes('jobcard'));
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

  // Jam realisasi per pegawai yang BELUM punya kode Jobcard (tetap ditampilkan di tab Anggota).
  const realisasiRows = useMemo(() => parseRealisasiRows(realisasiWorkbook), [realisasiWorkbook]);
  const realisasiNoJc = useMemo(() => summarizeNoJobcard(realisasiRows), [realisasiRows]);

  // Hitung Real JO per personel pada kode Jobcard (mendukung kode Jobcard bersama dalam proyek yang sama).
  const getRealHoursForTask = useCallback((pic: string, kodeJc: string): number | undefined => {
    const jcKey = cleanText(kodeJc || '');
    if (!jcKey) return undefined;
    const jcRows = realisasiRows.filter(r => cleanText(r.jobcard) === jcKey);
    if (jcRows.length > 0 && (pic || '').trim()) {
      const picClean = cleanText(pic);
      const personRows = jcRows.filter(r => cleanText(r.nama) === picClean || sameNameFuzzy(r.nama, pic));
      if (personRows.length > 0) {
        return personRows.reduce((sum, r) => sum + r.hours, 0);
      }
      return undefined;
    }
    return realisasiMap.get(jcKey);
  }, [realisasiRows, realisasiMap]);

  const getNoJcFor = (nama: string): NoJcSummary | undefined => {
    const direct = realisasiNoJc.get(cleanText(nama));
    if (direct) return direct;
    for (const v of realisasiNoJc.values()) {
      if (sameNameFuzzy(v.nama, nama)) return v;
    }
    return undefined;
  };

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

  const dynamicOutsourcingList = useMemo(() => {
    return allParsedFromExcel.filter(p => p.status.toLowerCase().includes('outsourcing'));
  }, [allParsedFromExcel]);

  const getSubconMembersForBiro = useCallback((biroName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.biro, biroName));
  }, [dynamicOutsourcingList]);

  const getSubconCountForDept = useCallback((deptName: string): number => {
    try {
      const dept = (departmentsData || []).find(d => d.name === deptName);
      if (!dept) return 0;
      const uniq = new Set<string>();
      dept.biros.forEach(biro => {
        getSubconMembersForBiro(biro.name).forEach(p => uniq.add(p.nip || p.nama));
      });
      return uniq.size;
    } catch { return 0; }
  }, [getSubconMembersForBiro]);

  const getBiroMembers = useCallback((biroName: string) => {
    return allParsedFromExcel.filter(p =>
      isBiroMatch(p.biro, biroName) &&
      !p.status.toLowerCase().includes('outsourcing') &&
      !isPimpinanDivDept(p.jabatan));
  }, [allParsedFromExcel]);

  const toggleAccordion = (cardKey: string) => {
    setExpandedCards(prev => ({ ...prev, [cardKey]: !prev[cardKey] }));
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
    let releaseVal = editFormData.release ? editFormData.release.trim() : '';
    const keyChanged = editingTask.project !== editFormData.project
      || editingTask.taskName !== editFormData.taskName
      || String(editingTask.rev || '0') !== revVal;
    if (keyChanged) {
      const cached = drawingControlMap[cleanText(editFormData.project)];
      const rows = cached && cached.length > 0 ? cached : await fetchDrawingRowsFor(editFormData.project);
      releaseVal = findReleaseInRows(rows, editFormData.taskName, revVal);
    }
    const trimmedKodeJc = (editFormData.kodeJc || '').replace(/\s+/g, ' ').trim().toUpperCase();
    const hpEntry = trimmedKodeJc ? hourPlannerMap[cleanText(trimmedKodeJc)] : undefined;
    const resolvedJo = ((!editFormData.jo || editFormData.jo === '0' || editFormData.jo === '-') && hpEntry)
      ? String(hpEntry.hour)
      : editFormData.jo;
    const nextStatus = trimmedKodeJc ? 'approved' : (editingTask.status || 'pending');
    const updatePayload: any = {
      project: editFormData.project,
      project_code: editFormData.project,
      task_name: editFormData.taskName,
      start_date: editFormData.startDate,
      end_date: editFormData.endDate,
      jo: resolvedJo,
      kode_jc: trimmedKodeJc,
      status: nextStatus,
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
              jo: resolvedJo,
              kodeJc: trimmedKodeJc,
              status: nextStatus,
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

  const handleDeleteTask = async (taskId: string) => {
    if (window.confirm('Apakah Anda yakin ingin menghapus tugas ini?')) {
      const { error } = await supabase.from('job_cards').delete().eq('id', taskId);
      if (error) {
        alert('Gagal menghapus: ' + error.message);
        return;
      }
      setManualTasks(prev => {
        const updated = { ...prev };
        Object.keys(updated).forEach(k => {
          updated[k] = updated[k].filter(t => t.id !== taskId);
        });
        return updated;
      });
      alert('Tugas berhasil dihapus!');
      loadAllJobCards();
    }
  };

  const fetchDrawingControlForProject = useCallback(async (projectCode: string) => {
    const cleanProj = cleanText(projectCode); if (!cleanProj) return;
    setIsFetchingDrawing(true);
    try {
      const rows = await fetchDrawingRowsFor(projectCode);
      if (rows.length > 0) {
        setDrawingControlMap(prev => ({ ...prev, [cleanProj]: rows }));
        setActiveDrawingSheetTitle(`Drawing Control ${projectCode}`);
      }
    } finally {
      setIsFetchingDrawing(false);
    }
  }, []);

  const handleProjectChange = (newProject: string) => {
    setFormData(prev => ({ ...prev, kodeProyek: newProject, taskName: '', release: '' }));
    fetchDrawingControlForProject(newProject);
  };

  const handleDeskripsiChange = (selectedDesc: string) => {
    const cleanProj = cleanText(formData.kodeProyek || '');
    const rowsForThisProj = drawingControlMap[cleanProj] || [];
    let autoRelease = ''; let autoRev = formData.rev || '0';
    if (rowsForThisProj.length > 0 && selectedDesc) {
      const cleanTarget = cleanText(selectedDesc);
      const currentFormRev = String(formData.rev || '0').trim();
      const exactMatch = rowsForThisProj.find(r => {
        const cFull = cleanText(r.fullDeskripsi); const cDwg = cleanText(r.noDwg); const cName = cleanText(r.drawingName);
        const rowRev = String(r.rev || '0').trim();
        const isRevMatch = rowRev === currentFormRev;
        const isNameMatch = cleanTarget.includes(cDwg) || cleanTarget.includes(cName) || cFull.includes(cleanTarget) || cleanTarget.includes(cFull);
        return isNameMatch && isRevMatch;
      });
      if (exactMatch && exactMatch.finishDate) { autoRelease = parseToStandardDate(exactMatch.finishDate); if (exactMatch.rev) autoRev = exactMatch.rev; }
      else { autoRelease = ''; }
    }
    setFormData(prev => ({ ...prev, taskName: selectedDesc, release: autoRelease, rev: autoRev }));
  };

  const handleQuickSaveRelease = async (taskId: string, newVal: string) => {
    const cleanDate = parseToStandardDate(newVal);
    saveLocalRelease(taskId, cleanDate);
    try { await supabase.from('job_cards').update({ release: cleanDate }).eq('id', taskId); } catch {}
    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => { updated[k] = updated[k].map(t => t.id === taskId ? { ...t, release: cleanDate } : t); });
      return updated;
    });
    setEditingReleaseId(null);
  };

  const handleSyncReleaseFromDrive = async (task: TaskItem) => {
    const cleanProj = cleanText(task.project);
    let rows: DrawingControlRow[] | undefined = drawingControlMap[cleanProj];
    if (!rows || rows.length === 0) {
      rows = await fetchDrawingRowsFor(task.project);
      if (rows.length > 0) {
        const fetched = rows;
        setDrawingControlMap(prev => ({ ...prev, [cleanProj]: fetched }));
      }
    }
    const rel = findReleaseInRows(rows || [], task.taskName, task.rev || '0');
    if (rel) {
      await handleQuickSaveRelease(task.id, rel);
      alert(`Berhasil sinkronisasi dari Google Drive: ${formatDisplayDate(rel)}`);
      return;
    }
    alert(`Tidak ditemukan data FINISH DATE di Google Drive untuk gambar "${task.taskName}" dengan Rev ${task.rev}.`);
  };

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    const activeBiro = accessMode === 'subkon' ? subconSelectedBiro : selectedFormBiro?.biroName;
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
      let releaseVal = '';
      if (formData.kodeProyek && formData.taskName) {
        const cp = cleanText(formData.kodeProyek);
        const cached = drawingControlMap[cp];
        const rows = cached && cached.length > 0 ? cached : await fetchDrawingRowsFor(formData.kodeProyek);
        releaseVal = findReleaseInRows(rows, formData.taskName, revVal);
      }
      const insertPayload: any = {
        biro_id: validBiroId, biro_name: activeBiro, personil_name: formData.nama, project_code: formData.kodeProyek,
        project: formData.kodeProyek, task_name: formData.taskName, start_date: formData.startDate, end_date: formData.endDate,
        pic: formData.nama, jo: formData.jo, kode_jc: autoKode, status: accessMode === 'subkon' ? 'approved' : 'pending',
        rev: revVal, release: releaseVal,
      };
      let insertedRow: any = null;
      const { data: resData, error } = await supabase.from('job_cards').insert(insertPayload).select().single();
      if (error) {
        delete insertPayload.rev; delete insertPayload.release;
        const { data: retryData, error: retryError } = await supabase.from('job_cards').insert(insertPayload).select().single();
        if (retryError) { alert('Gagal simpan: ' + retryError.message); return; }
        insertedRow = retryData;
      } else { insertedRow = resData; }
      if (releaseVal) saveLocalRelease(insertedRow.id, releaseVal);
      const newTask: TaskItem = {
        id: insertedRow.id, biroName: activeBiro, project: formData.kodeProyek, taskName: formData.taskName,
        startDate: formData.startDate, endDate: formData.endDate, pic: formData.nama, jo: formData.jo,
        kodeJc: autoKode, rev: revVal, release: releaseVal,
        status: accessMode === 'subkon' ? 'approved' : 'pending',
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
    } catch { alert('Koneksi database bermasalah.'); }
  };

  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;
    const hp = hourPlannerMap[cleanText(inputVal)];
    const payload: any = { kode_jc: inputVal, status: 'approved' };
    if (hp) payload.jo = String(hp.hour);
    const { error } = await supabase.from('job_cards').update(payload).eq('id', taskId);
    if (error) { alert('Gagal simpan Jobcard: ' + error.message); return; }
    setManualTasks(prev => {
      const updated = { ...prev };
      Object.keys(updated).forEach(k => {
        updated[k] = updated[k].map(t => t.id === taskId
          ? { ...t, kodeJc: inputVal, status: 'approved', ...(hp ? { jo: String(hp.hour) } : {}) }
          : t);
      });
      return updated;
    });
    setEditingTaskKode(prev => { const next = { ...prev }; delete next[taskId]; return next; });
    loadAllJobCards();
  };

  const filteredDepartments = (departmentsData || []).filter(dept => dept.name.toLowerCase().includes(searchQuery.toLowerCase()));
  const filteredSubconDepartments = (departmentsData || []).filter(dept => dept.name.toLowerCase().includes(subconSearch.toLowerCase()));

  const currentActiveBiroName = useMemo(() => {
    if (accessMode === 'subkon') return subconSelectedBiro || '';
    return selectedFormBiro?.biroName || '';
  }, [accessMode, subconSelectedBiro, selectedFormBiro]);

  const currentActiveBiroKey = useMemo(() => cleanText(currentActiveBiroName), [currentActiveBiroName]);
  const currentActiveBiroTasks = useMemo(() => manualTasks[currentActiveBiroKey] || [], [manualTasks, currentActiveBiroKey]);

  // Memastikan data Kadep muncul saat diklik
  const currentBiroMembers = useMemo(() => {
    if (!selectedFormBiro) return [];
    if (selectedFormBiro.biroName.startsWith('Kepala ')) {
      return getKadepForDept(allParsedFromExcel, selectedFormBiro.deptName);
    }
    return getBiroMembers(selectedFormBiro.biroName);
  }, [selectedFormBiro, allParsedFromExcel]);

  // Gabungan semua anggota departemen untuk Panel Mingguan
  const allDeptMembers = useMemo(() => {
    if (!selectedDept) return [];
    const kadeps = getKadepForDept(allParsedFromExcel, selectedDept.name);
    const biroMembers = selectedDept.biros.flatMap(biro => getBiroMembers(biro.name));
    const seen = new Set<string>();
    return [...kadeps, ...biroMembers].filter(p => {
      const id = p.nip || p.nama;
      if (seen.has(id)) return false;
      seen.add(id);
      return true;
    });
  }, [selectedDept, allParsedFromExcel]);

  // Grup Kadep + Staf per departemen untuk halaman Rekap Absensi Divisi
  const divisiGroups = useMemo<DivisiGroup[]>(() => {
    return (departmentsData || []).map(dept => {
      const kadep = getKadepForDept(allParsedFromExcel, dept.name);
      const seen = new Set<string>(kadep.map(p => p.nip || p.nama));
      const staf = dept.biros.flatMap(b => getBiroMembers(b.name)).filter(p => {
        const id = p.nip || p.nama;
        if (seen.has(id)) return false;
        seen.add(id);
        return true;
      });
      return { deptName: dept.name, kadep, staf };
    }).filter(g => g.kadep.length + g.staf.length > 0);
  }, [allParsedFromExcel, getBiroMembers]);

  const activeSubconMembers = useMemo(() => {
    if (!subconSelectedBiro) return [];
    return getSubconMembersForBiro(subconSelectedBiro);
  }, [subconSelectedBiro, getSubconMembersForBiro]);

  const subconWorkOrders = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const prefix = getBiroPrefix(subconSelectedBiro);
    const subconNames = getSubconMembersForBiro(subconSelectedBiro).map(m => cleanText(m.nama));
    const filteredTasks = currentActiveBiroTasks.filter(t => subconNames.includes(cleanText(t.pic)));
    return filteredTasks.map((task, idx) => ({ ...task, packageTitle: `${prefix}${idx + 1}` }));
  }, [subconSelectedBiro, currentActiveBiroTasks, getSubconMembersForBiro]);

  const projectOptions = useMemo((): string[] => {
    const projectMap = new Map<string, string>();
    GOOGLE_DRIVE_SHEETS.forEach(s => {
      const cleaned = cleanProjectString(s.projectKey);
      const normKey = getProjectNormKey(cleaned);
      if (cleaned && normKey && !projectMap.has(normKey)) projectMap.set(normKey, cleaned);
    });
    if (jobcardWorkbook) {
      try {
        jobcardWorkbook.SheetNames.forEach(sheetName => {
          const sheet = jobcardWorkbook.Sheets[sheetName]; if (!sheet) return;
          const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          let projCol = 2;
          for (let r = 0; r < Math.min(5, rows.length); r++) {
            const rowVals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
            const foundIdx = rowVals.findIndex(v => v.includes('proyek') || v.includes('project'));
            if (foundIdx !== -1) { projCol = foundIdx; break; }
          }
          rows.forEach((row, idx) => {
            if (idx < 1 || !row) return;
            const raw = String(row[projCol] || '').trim();
            if (!raw || raw.toLowerCase() === 'nan' || raw.toLowerCase().includes('kode proyek') || raw.toLowerCase() === 'proyek' || raw.toLowerCase() === 'project') return;
            const cleaned = cleanProjectString(raw);
            const normKey = getProjectNormKey(cleaned);
            if (cleaned && normKey && !projectMap.has(normKey)) projectMap.set(normKey, cleaned);
          });
        });
      } catch {}
    }
    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        const raw = String(t.project || '').trim();
        if (raw) {
          const cleaned = cleanProjectString(raw);
          const normKey = getProjectNormKey(cleaned);
          if (cleaned && normKey && !projectMap.has(normKey)) projectMap.set(normKey, cleaned);
        }
      });
    });
    return Array.from(projectMap.values()).sort((a, b) => a.localeCompare(b));
  }, [jobcardWorkbook, manualTasks]);

  const dynamicTaskOptions = useMemo((): string[] => {
    const taskMap = new Map<string, string>();
    if (jobcardWorkbook) {
      try {
        jobcardWorkbook.SheetNames.forEach(sheetName => {
          const sheet = jobcardWorkbook.Sheets[sheetName];
          if (!sheet) return;
          const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          let taskCol = -1;
          for (let r = 0; r < Math.min(15, rows.length); r++) {
            const rowVals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
            const foundIdx = rowVals.findIndex(v =>
              v.includes('task') ||
              v.includes('deskripsi') ||
              v.includes('uraian') ||
              v.includes('pekerjaan') ||
              v.includes('nama gambar') ||
              v.includes('judul')
            );
            if (foundIdx !== -1) {
              taskCol = foundIdx;
              break;
            }
          }
          if (taskCol === -1) taskCol = 3;
          rows.forEach((row, idx) => {
            if (idx < 1 || !row) return;
            const raw = String(row[taskCol] || '').replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ').trim();
            if (!raw || raw.toLowerCase() === 'nan' || raw.toLowerCase().includes('deskripsi') || raw.toLowerCase() === 'task' || raw.toLowerCase() === 'task name' || raw.toUpperCase().startsWith('REVISI KE - NO DOC')) return;
            const cleaned = raw.replace(/[\.,;:\-_/\\\s]+$/, '').trim();
            const normKey = cleanText(cleaned);
            if (cleaned && normKey && !taskMap.has(normKey)) {
              taskMap.set(normKey, cleaned);
            }
          });
        });
      } catch (err) {
        console.error("Gagal membaca jobcardWorkbook untuk opsi deskripsi:", err);
      }
    }
    return Array.from(taskMap.values()).sort((a, b) => a.localeCompare(b));
  }, [jobcardWorkbook]);

  // Planner hanya menampilkan tugas hasil isian Form (status 'pending') yang belum punya kode Jobcard.
  // Baris dari Excel Realisasi JO tidak pernah masuk ke sini.
  const pendingTasksByBiro = useMemo((): [string, TaskItem[]][] => {
    const groups: Record<string, TaskItem[]> = {};
    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        if (t.status !== 'pending') return;
        if (!t.pic || !t.project) return;
        if (!t.kodeJc || t.kodeJc.trim() === '') {
          const key = t.biroName || 'Tanpa Biro';
          if (!groups[key]) groups[key] = [];
          groups[key].push(t);
        }
      });
    });
    return Object.entries(groups).sort(([a], [b]) => a.localeCompare(b));
  }, [manualTasks]);

  const pendingTasksCount = useMemo(() => {
    return pendingTasksByBiro.reduce((total, [, list]) => total + list.length, 0);
  }, [pendingTasksByBiro]);

  if (showKpi) {
    return <KpiDashboardPage user={user} onLogout={onLogout} onBack={() => setShowKpi(false)} />;
  }
  if (showImport) {
    return <ImportJobcardPage user={user} onLogout={onLogout} onBack={() => setShowImport(false)} />;
  }
  if (showRekapDivisi) {
    return (
      <RekapAbsensiDivisiPage
        groups={divisiGroups}
        wb={timeOffWorkbook}
        sisaCuti={sisaCuti}
        onBack={() => (user.role === 'rekap' ? onLogout() : setShowRekapDivisi(false))}
        onUploadTimeOff={handleUploadTimeOffExcel}
        onRefreshSisaCuti={refreshSisaCuti}
      />
    );
  }

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5 cursor-pointer" onClick={() => { setAccessMode('landing'); setSelectedDept(null); setSelectedFormBiro(null); setSubconSelectedBiro(null); setSubconSelectedDept(null); }}>
            <div className="p-1.5 bg-purple-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <div className="flex items-center gap-2">
              <span className="font-bold text-sm text-white">PORTAL ADMIN RENDAL</span>
              {accessMode !== 'landing' && (
                <span className="text-xs px-2 py-0.5 rounded font-mono bg-slate-800 text-slate-300 border border-slate-700">
                  {accessMode === 'organik' ? 'Organik' : 'Subkon'}
                </span>
              )}
            </div>
          </div>
          <div className="flex items-center gap-2">
            {autoSyncMsg && (
              <span className="px-2.5 py-1 text-[11px] bg-cyan-500/10 text-cyan-300 rounded-lg border border-cyan-500/30 flex items-center gap-1">
                <Sparkles className="w-3 h-3" /> {autoSyncMsg}
              </span>
            )}
            {accessMode !== 'landing' && (
              <button onClick={() => { setAccessMode('landing'); setSelectedDept(null); setSelectedFormBiro(null); setSubconSelectedBiro(null); setSubconSelectedDept(null); }}
                className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
                Ganti Portal
              </button>
            )}
            <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
              Keluar
            </button>
          </div>
        </div>
      </header>
      <main className="max-w-6xl mx-auto px-4 py-8">
        {accessMode === 'landing' && (
          <div className="max-w-2xl mx-auto text-center space-y-8 pt-8">
            <div className="space-y-2">
              <h1 className="text-2xl sm:text-3xl font-bold text-white">Sistem Penugasan Job Card</h1>
              <p className="text-slate-400 text-sm">Pilih portal akses kerja Anda</p>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div onClick={() => setAccessMode('organik')}
                className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-2xl p-6 cursor-pointer text-left transition hover:bg-slate-900/80 group">
                <div className="p-3 bg-blue-500/10 text-blue-400 rounded-xl w-fit mb-4 group-hover:bg-blue-600 group-hover:text-white transition">
                  <Briefcase className="w-6 h-6" />
                </div>
                <h3 className="text-lg font-bold text-white">Pegawai Organik</h3>
                <span className="text-xs text-slate-400">Pegawai PKWTT & PKWT Divisi Desain</span>
              </div>
              <div onClick={() => setAccessMode('subkon')}
                className="bg-slate-900 border border-slate-800 hover:border-amber-500 rounded-2xl p-6 cursor-pointer text-left transition hover:bg-slate-900/80 group">
                <div className="p-3 bg-amber-500/10 text-amber-400 rounded-xl w-fit mb-4 group-hover:bg-amber-600 group-hover:text-white transition">
                  <HardHat className="w-6 h-6" />
                </div>
                <h3 className="text-lg font-bold text-white">Mitra / Subkon</h3>
                <span className="text-xs text-slate-400">Personil Outsourcing</span>
              </div>
            </div>
          </div>
        )}

        {accessMode === 'subkon' && !subconSelectedDept && !subconSelectedBiro && (
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
                <input type="text" placeholder="Cari..." value={subconSearch} onChange={(e) => setSubconSearch(e.target.value)}
                  className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44" />
              </div>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
              {filteredSubconDepartments.map((dept) => {
                const count = getSubconCountForDept(dept.name);
                return (
                  <div key={dept.id} onClick={() => setSubconSelectedDept(dept)}
                    className="bg-slate-900 border border-slate-800 hover:border-amber-500/60 rounded-xl p-4 cursor-pointer transition flex items-center justify-between">
                    <div>
                      <h4 className="font-semibold text-white text-sm">{dept.name}</h4>
                      <span className="text-xs text-slate-400">{dept.biros.length} Biro</span>
                    </div>
                    <span className="px-2 py-0.5 bg-amber-500/10 text-amber-400 border border-amber-500/20 rounded font-bold text-xs">{count}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {accessMode === 'subkon' && subconSelectedDept && !subconSelectedBiro && (
          <div className="space-y-4">
            <button onClick={() => setSubconSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{subconSelectedDept.name}</h2>
              <span className="text-xs text-slate-400">Pilih biro penugasan</span>
            </div>
            <div className="space-y-2">
              {subconSelectedDept.biros.map((biro) => {
                const countInBiro = getSubconMembersForBiro(biro.name).length;
                return (
                  <div key={biro.id} className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-semibold text-white text-sm">{biro.name}</span>
                      <span className="text-xs text-amber-400 font-mono">({countInBiro} Org)</span>
                    </div>
                    <div className="flex items-center gap-1.5 self-end sm:self-auto">
                      <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('members'); }}
                        className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs rounded border border-slate-700 cursor-pointer">Anggota</button>
                      <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('form'); }}
                        className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs rounded cursor-pointer">Form</button>
                      <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('release'); }}
                        className="px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white font-semibold text-xs rounded cursor-pointer">Work Order</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {accessMode === 'subkon' && subconSelectedBiro && (
          <div className="space-y-4">
            <button onClick={() => setSubconSelectedBiro(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali
            </button>
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h3 className="font-bold text-white text-base">{subconSelectedBiro}</h3>
                <span className="text-xs text-slate-400 font-mono">{activeSubconMembers.length} Personel Outsourcing • Kode Biro: <b className="text-amber-400">{getBiroPrefix(subconSelectedBiro)}</b></span>
              </div>
              <div className="bg-slate-950 p-1 rounded-lg border border-slate-800 flex gap-1 text-xs font-semibold">
                <button onClick={() => setSubconPageMode('members')}
                  className={`px-3 py-1 rounded transition cursor-pointer ${subconPageMode === 'members' ? 'bg-amber-600 text-white font-bold shadow' : 'text-slate-400 hover:text-white hover:bg-slate-900'}`}>Anggota</button>
                <button onClick={() => setSubconPageMode('form')}
                  className={`px-3 py-1 rounded transition cursor-pointer ${subconPageMode === 'form' ? 'bg-emerald-600 text-white font-bold shadow' : 'text-slate-400 hover:text-white hover:bg-slate-900'}`}>Form</button>
                <button onClick={() => setSubconPageMode('release')}
                  className={`px-3 py-1 rounded transition cursor-pointer flex items-center gap-1 ${subconPageMode === 'release' ? 'bg-purple-600 text-white font-bold shadow' : 'text-purple-300 hover:text-white hover:bg-slate-900'}`}>
                  <FileCheck className="w-3 h-3 text-amber-300" /> Work Order ({subconWorkOrders.length})
                </button>
              </div>
            </div>
            {subconPageMode === 'members' && (
              <div className="space-y-2.5">
                {activeSubconMembers.length > 0 ? activeSubconMembers.map((person) => {
                  const personTasks = currentActiveBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
                  const isExpanded = !!expandedCards[person.nama];
                  return (
                    <div key={`${person.nama}-${person.nip}`} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                      <div onClick={() => toggleAccordion(person.nama)} className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition">
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-lg bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center font-bold text-xs">{person.nama.charAt(0)}</div>
                          <div>
                            <div className="font-semibold text-white text-sm">{person.nama}</div>
                            <div className="text-xs text-slate-400 font-mono">NIP: {person.nip || '-'} • <span className="text-cyan-400">{person.jabatan}</span></div>
                          </div>
                        </div>
                        <div className="flex items-center gap-3">
                          <span className="text-xs font-mono text-slate-400 hidden sm:inline">{personTasks.length} Tugas</span>
                          <div className={`p-1.5 rounded-lg bg-slate-950 border border-slate-800 text-slate-400 transition-transform duration-200 ${isExpanded ? 'text-amber-400' : ''}`}>
                            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                          </div>
                        </div>
                      </div>
                      {isExpanded && (
                        <div className="p-3.5 border-t border-slate-800 bg-slate-950/60">
                          {personTasks.length > 0 ? (
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs">
                                <thead><tr className="text-slate-400 border-b border-slate-800">
                                  <th className="py-2 px-2.5 w-8">#</th>
                                  <th className="py-2 px-2.5 font-mono text-amber-400">Kode WO</th>
                                  <th className="py-2 px-2.5">Proyek</th>
                                  <th className="py-2 px-2.5">Deskripsi</th>
                                  <th className="py-2 px-2.5 font-mono">Jadwal</th>
                                  <th className="py-2 px-2.5 font-mono">JO</th>
                                  <th className="py-2 px-2.5 text-center w-20">Aksi</th>
                                </tr></thead>
                                <tbody className="divide-y divide-slate-800 text-slate-300">
                                  {personTasks.map((task, tIdx) => (
                                    <tr key={task.id} className="hover:bg-slate-900/40">
                                      <td className="py-2 px-2.5 text-slate-500 font-mono">{tIdx + 1}</td>
                                      <td className="py-2 px-2.5 font-mono font-bold text-amber-300">{task.kodeJc || `${getBiroPrefix(subconSelectedBiro)}${tIdx + 1}`}</td>
                                      <td className="py-2 px-2.5 text-emerald-400">{task.project}</td>
                                      <td className="py-2 px-2.5 text-slate-200">{task.taskName}</td>
                                      <td className="py-2 px-2.5 font-mono text-[11px] text-slate-400">{formatDisplayDate(task.startDate)} s/d {formatDisplayDate(task.endDate)}</td>
                                      <td className="py-2 px-2.5 font-mono text-violet-300">{task.jo ? String(task.jo).replace(/^#+/, '') : '-'}</td>
                                      <td className="py-2 px-2.5 text-center">
                                        <div className="flex items-center justify-center gap-1.5">
                                          <button
                                            onClick={(e) => { e.stopPropagation(); handleOpenEdit(task); }}
                                            className="p-1 rounded bg-slate-800 hover:bg-blue-500/20 text-slate-400 hover:text-blue-400 transition cursor-pointer"
                                            title="Edit Tugas"
                                          >
                                            <Pencil className="w-3.5 h-3.5" />
                                          </button>
                                          <button
                                            onClick={(e) => { e.stopPropagation(); handleDeleteTask(task.id); }}
                                            className="p-1 rounded bg-slate-800 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 transition cursor-pointer"
                                            title="Hapus Tugas"
                                          >
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        </div>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : <div className="text-xs text-slate-500 py-3 text-center italic">Belum ada tugas.</div>}
                        </div>
                      )}
                    </div>
                  );
                }) : <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">Tidak ada personel outsourcing di biro ini</div>}
              </div>
            )}
            {subconPageMode === 'form' && (
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <form onSubmit={handleSubmitForm} className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                    <div className="md:col-span-2">
                      <label className="block text-slate-400 mb-1">Nama Drafter / Personel (Outsourcing)</label>
                      <SearchableSelect options={activeSubconMembers.map(p => ({ value: p.nama, label: `${p.nama} (${p.jabatan})` }))}
                        value={formData.nama} onChange={(val) => setFormData(prev => ({ ...prev, nama: val }))}
                        placeholder="Ketik nama personel outsourcing..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Kode Proyek</label>
                      <SearchableSelect options={projectOptions} value={formData.kodeProyek} onChange={(val) => handleProjectChange(val)}
                        placeholder="Ketik kode proyek..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Nomor JO</label>
                      <input type="text" inputMode="decimal" value={formData.jo}
                        onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9.]/g, '') }))}
                        placeholder="Contoh: 300426" required
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
                    </div>
                    <div className="md:col-span-2">
                      <label className="block text-slate-400 mb-1">Deskripsi</label>
                      <SearchableSelect options={dynamicTaskOptions} value={formData.taskName} onChange={handleDeskripsiChange}
                        placeholder="Ketik nama gambar / deskripsi tugas..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Tanggal Mulai</label>
                      <input type="date" value={formData.startDate} onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                        required style={{ colorScheme: 'dark' }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Tanggal Selesai</label>
                      <input type="date" value={formData.endDate} onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                        required style={{ colorScheme: 'dark' }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" />
                    </div>
                  </div>
                  <div className="pt-3 border-t border-slate-800 flex justify-end">
                    <button type="submit" className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-white font-bold text-xs rounded-lg transition cursor-pointer">
                      Simpan Tugas & Terbitkan Work Order
                    </button>
                  </div>
                </form>
              </div>
            )}
            {subconPageMode === 'release' && (
              <div className="space-y-3">
                {subconWorkOrders.length > 0 ? (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    {subconWorkOrders.map((pkg) => (
                      <div key={pkg.id} className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-2">
                        <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                          <span className="px-2.5 py-0.5 bg-purple-600 rounded text-white font-mono font-bold text-sm tracking-wide shadow-sm">{pkg.packageTitle}</span>
                          <span className="font-mono text-xs text-amber-400 font-bold bg-amber-500/10 px-2 py-0.5 rounded border border-amber-500/20">{pkg.kodeJc || pkg.packageTitle}</span>
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
                ) : <div className="bg-slate-900 border border-slate-800 rounded-xl p-8 text-center text-xs text-slate-500">Belum ada Work Order outsourcing di biro ini.</div>}
              </div>
            )}
          </div>
        )}

        {accessMode === 'organik' && !selectedDept && !selectedFormBiro && (
          <div className="space-y-5">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
              <div>
                <h2 className="text-lg font-bold text-white">Departemen Desain</h2>
                <span className="text-xs text-slate-400">Akses Pegawai Organik (PKWTT & PKWT)</span>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                  <Users className="w-3.5 h-3.5 text-blue-400" /><span>Update Personel IM4</span>
                  <input type="file" accept=".xlsx, .xls" onChange={handleUpdateIm4Excel} className="hidden" />
                </label>
                <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" /><span>Update Jobcard Desain</span>
                  <input type="file" accept=".xlsx, .xls" onChange={handleUpdateJobcardExcel} className="hidden" />
                </label>
                <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                  <Clock className="w-3.5 h-3.5 text-emerald-400" /><span>{realisasiMap.size > 0 ? `Realisasi JO: OK (${realisasiMap.size})` : 'Upload Realisasi JO.xlsx'}</span>
                  <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
                </label>
                <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-violet-400" />
                  <span>{timeOffWorkbook ? 'Update Time Off' : 'Upload Time Off'}</span>
                  <input type="file" accept=".xlsx, .xls" onChange={handleUploadTimeOffExcel} className="hidden" />
                </label>
                <button
                  onClick={() => setIsPlannerOpen(true)}
                  className="relative px-2.5 py-1.5 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  title={pendingTasksCount > 0 ? `${pendingTasksCount} tugas menunggu kode Jobcard` : 'Tidak ada tugas yang menunggu kode Jobcard'}
                >
                  <FileCheck className="w-3.5 h-3.5" /> Planner
                  {pendingTasksCount > 0 && (
                    <span className="absolute -top-1.5 -right-1.5 min-w-[18px] h-[18px] px-1 rounded-full bg-rose-500 text-white text-[10px] font-bold flex items-center justify-center border-2 border-slate-950">
                      {pendingTasksCount > 99 ? '99+' : pendingTasksCount}
                    </span>
                  )}
                </button>
                <button
                  onClick={() => setShowKpi(true)}
                  className="px-2.5 py-1.5 bg-purple-600 hover:bg-purple-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  title="Buka Dashboard KPI Individual"
                >
                  <TrendingUp className="w-3.5 h-3.5" /> Dashboard KPI
                </button>
                <button
                  onClick={() => setShowImport(true)}
                  className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  title="Impor Jobcard bulanan dari JOBCARD_DESAIN.xlsx"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" /> Import Jobcard
                </button>
                <label
                  className="px-2.5 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  title="Impor Hour Planner (Hour_Planner.xlsx): kolom Hour Planned masuk ke Plan JO sesuai kode Jobcard"
                >
                  <Clock className="w-3.5 h-3.5" />
                  {isImportingHour ? 'Mengimpor...' : `Import Hour Planner${Object.keys(hourPlannerMap).length > 0 ? ` (${Object.keys(hourPlannerMap).length})` : ''}`}
                  <input type="file" accept=".xlsx, .xls" onChange={handleImportHourPlanner} disabled={isImportingHour} className="hidden" />
                </label>
                <label
                  className="px-2.5 py-1.5 bg-rose-600 hover:bg-rose-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer shadow-sm"
                  title="Impor 1 file gabungan (TARIKAN_TIMESHEET.xlsx: Hour Planner + Realisasi JO + Jobcard + Deskripsi). Otomatis menghapus data lama & mengganti dengan data baru."
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  {isImportingUnified ? 'Memproses Data Terpadu...' : 'Import Data Terpadu (Reset)'}
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleImportUnifiedData} disabled={isImportingUnified} className="hidden" />
                </label>
                <input type="text" placeholder="Cari..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                  className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44" />
              </div>
            </div>

            {/* Tombol Rekap Absensi Divisi Desain */}
            <button
              onClick={() => setShowRekapDivisi(true)}
              className="w-full bg-gradient-to-r from-purple-600/20 to-blue-600/10 hover:from-purple-600/30 border border-purple-500/30 hover:border-purple-400 rounded-xl p-4 flex items-center justify-between transition cursor-pointer group text-left"
            >
              <div className="flex items-center gap-3">
                <div className="p-2.5 bg-purple-600 rounded-lg text-white"><Clock className="w-5 h-5" /></div>
                <div>
                  <div className="font-bold text-white text-sm">Rekap Absensi Divisi Desain</div>
                  <div className="text-xs text-slate-400">Cuti, sakit, dan rencana cuti seluruh departemen: Kadep dan staf</div>
                </div>
              </div>
              <ChevronRight className="w-5 h-5 text-purple-300 group-hover:translate-x-1 transition" />
            </button>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
              {filteredDepartments.map((dept) => (
                <div key={dept.id} onClick={() => setSelectedDept(dept)}
                  className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between">
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

        {/* ===== TAMPILAN DETAIL DEPARTEMEN (Panel Mingguan & Tombol Kadep) ===== */}
        {accessMode === 'organik' && selectedDept && !selectedFormBiro && (
          <div className="space-y-4">
            <button onClick={() => setSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{selectedDept.name}</h2>
              <span className="text-xs text-slate-400">Daftar Biro Penugasan Pegawai Organik</span>
            </div>

            {/* 1. Panel Tidak Masuk Minggu Ini (Time Off + Rencana dari Google Sheet) */}
            <DeptWeeklyTimeOffPanel
              deptName={selectedDept.name}
              members={allDeptMembers}
              wb={timeOffWorkbook}
              sisaCuti={sisaCuti}
            />

            {/* 2. Baris Kadep (Hanya Tombol Rekap Absensi) */}
            {(() => {
              const kadepList = getKadepForDept(allParsedFromExcel, selectedDept.name);
              if (kadepList.length > 0) {
                return (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-2">
                    <div className="flex items-center gap-2">
                      <div className="w-8 h-8 rounded-lg bg-purple-500/10 border border-purple-500/20 text-purple-400 flex items-center justify-center font-bold text-xs">K</div>
                      <div>
                        <span className="font-semibold text-white text-sm">Kepala {selectedDept.name}</span>
                        <div className="text-[11px] text-slate-400">{kadepList.length} Pegawai</div>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <button
                        onClick={() => {
                          setSelectedFormBiro({ biroName: `Kepala ${selectedDept.name}`, deptName: selectedDept.name });
                          setFormPageMode('absensi');
                        }}
                        className="px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white font-semibold text-xs rounded cursor-pointer flex items-center gap-1"
                      >
                        <Clock className="w-3 h-3" /> Rekap Absensi
                      </button>
                    </div>
                  </div>
                );
              }
              return null;
            })()}

            {/* 3. Daftar Biro */}
            <div className="space-y-2">
              {selectedDept.biros.map((biro) => (
                <div key={biro.id} className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <span className="font-semibold text-white text-sm">{biro.name}</span>
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <button onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('members'); }}
                      className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs rounded cursor-pointer">Anggota</button>
                    <button onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('form'); }}
                      className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs rounded cursor-pointer">Form</button>
                    <button onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('absensi'); setAbsensiMonth(null); }}
                      className="px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white font-semibold text-xs rounded cursor-pointer">Rekap Absensi</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ===== TAMPILAN DETAIL BIRO / KADEP ===== */}
        {accessMode === 'organik' && selectedFormBiro && (
          <div className="space-y-4">
            <button onClick={() => setSelectedFormBiro(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali
            </button>
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h3 className="font-bold text-white text-sm">{selectedFormBiro.biroName}</h3>
                <div className="text-xs text-slate-400 font-mono flex items-center gap-2 mt-0.5">
                  <span>{currentBiroMembers.length} Pegawai Organik</span>
                  {activeDrawingSheetTitle && (
                    <span className="text-emerald-400 flex items-center gap-1">• <Sparkles className="w-3 h-3 text-amber-400" /> {activeDrawingSheetTitle}</span>
                  )}
                </div>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <label className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                  <Clock className="w-3 h-3 text-emerald-400" /><span>{realisasiMap.size > 0 ? 'Update Realisasi JO' : 'Upload Realisasi JO'}</span>
                  <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
                </label>
                <div className="flex gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 text-xs">
                  {/* Sembunyikan tab Anggota & Form khusus untuk Kadep */}
                  {!selectedFormBiro.biroName.startsWith('Kepala ') && (
                    <>
                      <button onClick={() => setFormPageMode('members')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${formPageMode === 'members' ? 'bg-blue-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}>
                        Anggota ({currentBiroMembers.length})
                      </button>
                      <button onClick={() => setFormPageMode('form')}
                        className={`px-3 py-1 rounded transition cursor-pointer ${formPageMode === 'form' ? 'bg-emerald-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}>
                        Form
                      </button>
                    </>
                  )}
                  <button onClick={() => setFormPageMode('absensi')}
                    className={`px-3 py-1 rounded transition cursor-pointer ${formPageMode === 'absensi' ? 'bg-purple-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}>
                    Rekap Absensi
                  </button>
                </div>
              </div>
            </div>

            {formPageMode === 'members' && (
              <div className="space-y-2.5">
                {currentBiroMembers.length > 0 ? currentBiroMembers.map((person) => {
                  const personTasks = currentActiveBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
                  const isExpanded = !!expandedCards[person.nama];
                  return (
                    <div key={`${person.nama}-${person.nip}`} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                      <div onClick={() => toggleAccordion(person.nama)} className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition">
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-lg bg-blue-500/10 border border-blue-500/20 text-blue-400 flex items-center justify-center font-bold text-xs">{person.nama.charAt(0)}</div>
                          <div>
                            <div className="font-semibold text-white text-sm">{person.nama} <span className="text-xs font-mono text-cyan-400 font-normal">({person.status})</span></div>
                            <div className="text-xs text-slate-400 font-mono">NIP: {person.nip || '-'} • <span className="text-slate-400">{person.jabatan}</span></div>
                          </div>
                        </div>
                        <div className="flex items-center gap-3">
                          <span className="text-xs font-mono text-slate-400 hidden sm:inline">{personTasks.length} Tugas</span>
                          <div className={`p-1.5 rounded-lg bg-slate-950 border border-slate-800 text-slate-400 transition-transform duration-200 ${isExpanded ? 'text-blue-400' : ''}`}>
                            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                          </div>
                        </div>
                      </div>
                      {isExpanded && (
                        <div className="p-3.5 border-t border-slate-800 bg-slate-950/60">
                          {/* Jam realisasi yang belum punya kode Jobcard (dari file Realisasi JO) */}
                          {(() => {
                            const nj = getNoJcFor(person.nama);
                            if (!nj) return null;
                            return (
                              <div className="mb-3 px-3 py-2 rounded-lg bg-slate-900 border border-slate-800 text-[11px] text-slate-300">
                                <span className="text-emerald-400 font-semibold">Realisasi tanpa Jobcard: </span>
                                <b className="font-mono text-white">{fmtNum(nj.hours)} jam</b>
                                <span className="text-slate-500"> ({nj.rows} baris)</span>
                                <div className="mt-1 text-[10px] text-slate-500 font-mono break-words">
                                  {Object.entries(nj.byProject)
                                    .sort((a, b) => b[1] - a[1])
                                    .map(([p, h]) => `${p}: ${fmtNum(h)} jam`)
                                    .join(' • ')}
                                </div>
                              </div>
                            );
                          })()}
                          {personTasks.length > 0 ? (
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs">
                                <thead><tr className="text-slate-400 border-b border-slate-800">
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
                                </tr></thead>
                                <tbody className="divide-y divide-slate-800 text-slate-300">
                                  {personTasks.map((task, tIdx) => {
                                    const jcKey = cleanText(task.kodeJc || '');
                                    const calculatedRealHours = getRealHoursForTask(task.pic, task.kodeJc);
                                    const rawJo = task.jo ? String(task.jo).replace(/^#+/, '').trim() : '';
                                    const displayPlanJo = (rawJo && rawJo !== '-')
                                      ? rawJo
                                      : (jcKey && hourPlannerMap[jcKey] ? String(hourPlannerMap[jcKey].hour) : '-');
                                    return (
                                      <tr key={task.id} className="hover:bg-slate-900/40">
                                        <td className="py-2 px-2.5 text-slate-500 font-mono">{tIdx + 1}</td>
                                        <td className="py-2 px-2.5 font-mono font-bold text-amber-300">{task.kodeJc || <span className="text-rose-400 font-normal">Menunggu Planner</span>}</td>
                                        <td className="py-2 px-2.5 text-emerald-400">{task.project}</td>
                                        <td className="py-2 px-2.5 text-slate-200">{task.taskName}</td>
                                        <td className="py-2 px-2.5 text-center font-mono text-slate-300">{task.rev || '0'}</td>
                                        <td className="py-2 px-2.5 font-mono text-[11px] text-slate-300">{formatDisplayDate(task.startDate || (jcKey && hourPlannerMap[jcKey]?.start))}</td>
                                        <td className="py-2 px-2.5 font-mono text-[11px] text-slate-300">{formatDisplayDate(task.endDate)}</td>
                                        <td className="py-2 px-2.5 font-mono text-violet-300">{displayPlanJo}</td>
                                        <td className="py-2 px-2.5 font-mono font-bold">
                                          {calculatedRealHours !== undefined ? (
                                            <span className="px-2 py-0.5 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded">{Math.round(calculatedRealHours * 100) / 100} Jam</span>
                                          ) : task.kodeJc ? <span className="text-slate-500 font-normal">0 Jam</span> : <span className="text-slate-600 font-normal">-</span>}
                                        </td>
                                        <td className="py-2 px-2.5 font-mono text-[11px]">
                                          {editingReleaseId === task.id ? (
                                            <div className="flex items-center gap-1">
                                              <input type="date" value={editingReleaseVal} onChange={(e) => setEditingReleaseVal(e.target.value)}
                                                style={{ colorScheme: 'dark' }} className="px-1.5 py-0.5 bg-slate-900 border border-blue-500 rounded text-xs text-white" autoFocus />
                                              <button onClick={() => handleQuickSaveRelease(task.id, editingReleaseVal)} className="p-1 text-emerald-400 hover:text-emerald-300 hover:bg-emerald-500/20 rounded cursor-pointer" title="Simpan">
                                                <Check className="w-3 h-3" />
                                              </button>
                                              <button onClick={() => setEditingReleaseId(null)} className="p-1 text-slate-400 hover:text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer" title="Batal">
                                                <X className="w-3 h-3" />
                                              </button>
                                            </div>
                                          ) : (
                                            <div className="flex items-center gap-1.5 group">
                                              <span className="text-cyan-300 font-semibold">{formatDisplayDate(task.release)}</span>
                                              <div className="flex items-center gap-1 opacity-70 group-hover:opacity-100 transition">
                                                <button onClick={(e) => { e.stopPropagation(); setEditingReleaseId(task.id); setEditingReleaseVal(parseToStandardDate(task.release) || ''); }}
                                                  className="p-0.5 rounded text-slate-400 hover:text-cyan-300 hover:bg-slate-800 transition cursor-pointer" title="Edit Langsung Tanggal Release">
                                                  <Pencil className="w-3 h-3" />
                                                </button>
                                                <button onClick={(e) => { e.stopPropagation(); handleSyncReleaseFromDrive(task); }}
                                                  className="p-0.5 rounded text-slate-400 hover:text-amber-400 hover:bg-slate-800 transition cursor-pointer" title="Tarik Otomatis dari Link Google Drive">
                                                  <Sparkles className="w-3 h-3" />
                                                </button>
                                              </div>
                                            </div>
                                          )}
                                        </td>
                                        <td className="py-2 px-2.5 text-center">
                                          <div className="flex items-center justify-center gap-1.5">
                                            <button
                                              onClick={(e) => { e.stopPropagation(); handleOpenEdit(task); }}
                                              className="p-1 rounded bg-slate-800 hover:bg-blue-500/20 text-slate-400 hover:text-blue-400 transition cursor-pointer"
                                              title="Edit Tugas"
                                            >
                                              <Pencil className="w-3.5 h-3.5" />
                                            </button>
                                            <button
                                              onClick={(e) => { e.stopPropagation(); handleDeleteTask(task.id); }}
                                              className="p-1 rounded bg-slate-800 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 transition cursor-pointer"
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
                          ) : <div className="text-xs text-slate-500 py-3 text-center italic">Belum ada tugas. Buka tab <b>Form</b> di atas untuk menambahkan penugasan.</div>}
                        </div>
                      )}
                    </div>
                  );
                }) : <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">Tidak ada personil organik di biro ini (Silakan upload Master IM4 terlebih dahulu)</div>}
              </div>
            )}

            {formPageMode === 'form' && (
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="md:col-span-2">
                      <label className="block text-slate-400 mb-1">Nama Personel (Organik)</label>
                      <SearchableSelect options={currentBiroMembers.map(p => ({ value: p.nama, label: `${p.nama} (${p.status})` }))}
                        value={formData.nama} onChange={(val) => setFormData(prev => ({ ...prev, nama: val }))}
                        placeholder="Ketik nama personel organik..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Kode Proyek</label>
                      <SearchableSelect options={projectOptions} value={formData.kodeProyek} onChange={(val) => handleProjectChange(val)}
                        placeholder="Ketik kode proyek (contoh: M000313)..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Plan JO</label>
                      <input type="text" inputMode="decimal" value={formData.jo}
                        onChange={(e) => setFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9.]/g, '') }))}
                        placeholder="Contoh: 300426" required
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
                    </div>
                    <div className="md:col-span-2">
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-slate-400">Deskripsi</label>
                        {isFetchingDrawing && <span className="text-[10px] text-amber-400 animate-pulse font-mono">Mengambil dari Google Drive...</span>}
                      </div>
                      <SearchableSelect options={dynamicTaskOptions} value={formData.taskName} onChange={handleDeskripsiChange}
                        placeholder="Ketik no dwg atau nama gambar (misal: Shaft Protection)..." required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Rev (Revisi)</label>
                      <input type="text" value={formData.rev}
                        onChange={(e) => {
                          const newRev = e.target.value;
                          const cleanProj = cleanText(formData.kodeProyek || '');
                          const rows = drawingControlMap[cleanProj] || [];
                          let autoRel = '';
                          const exactMatch = rows.find(r => {
                            const cFull = cleanText(r.fullDeskripsi); const cName = cleanText(r.drawingName);
                            const isRevMatch = String(r.rev || '0').trim() === String(newRev).trim();
                            const isNameMatch = cFull === cleanText(formData.taskName) || (cName && cleanText(formData.taskName).includes(cName));
                            return isNameMatch && isRevMatch;
                          });
                          if (exactMatch && exactMatch.finishDate) autoRel = parseToStandardDate(exactMatch.finishDate);
                          else autoRel = '';
                          setFormData(prev => ({ ...prev, rev: newRev, release: autoRel }));
                        }}
                        placeholder="0" className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
                    </div>
                    <div className="hidden">
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-slate-400">Release (Tanggal-Bulan-Tahun)</label>
                        <div className="flex items-center gap-1.5">
                          {formData.release && <span className="text-[10px] text-cyan-300 font-mono font-semibold">{formatDisplayDate(formData.release)}</span>}
                          <button type="button" onClick={() => handleDeskripsiChange(formData.taskName)}
                            className="px-1.5 py-0.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-400 border border-amber-500/30 rounded text-[10px] flex items-center gap-1 cursor-pointer" title="Tarik ulang dari Link Drive">
                            <Sparkles className="w-2.5 h-2.5" /> Dari Drive
                          </button>
                        </div>
                      </div>
                      <input type="date" value={formData.release} onChange={(e) => setFormData(prev => ({ ...prev, release: e.target.value }))}
                        style={{ colorScheme: 'dark' }} className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Plan Start</label>
                      <input type="date" value={formData.startDate} onChange={(e) => setFormData(prev => ({ ...prev, startDate: e.target.value }))}
                        required style={{ colorScheme: 'dark' }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Plan Finish</label>
                      <input type="date" value={formData.endDate} onChange={(e) => setFormData(prev => ({ ...prev, endDate: e.target.value }))}
                        required style={{ colorScheme: 'dark' }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none cursor-pointer" />
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

            {formPageMode === 'absensi' && (
              <>
                <CutiSakitPanel
                  biroName={selectedFormBiro.biroName}
                  members={currentBiroMembers}
                  wb={timeOffWorkbook}
                  sisaCuti={sisaCuti}
                />
              </>
            )}
          </div>
        )}
      </main>

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
                    inputMode="decimal"
                    value={editFormData.jo}
                    onChange={(e) => setEditFormData(prev => ({ ...prev, jo: e.target.value.replace(/[^0-9.]/g, '') }))}
                    required={accessMode !== 'organik'}
                    placeholder={accessMode === 'organik' ? 'Otomatis dari Hour Planner / ketik manual' : ''}
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
                    onChange={(e) => {
                      const val = e.target.value.toUpperCase();
                      const hp = accessMode === 'organik' ? hourPlannerMap[cleanText(val)] : undefined;
                      setEditFormData(prev => ({
                        ...prev,
                        kodeJc: val,
                        jo: ((!prev.jo || prev.jo === '0' || prev.jo === '-') && hp) ? String(hp.hour) : prev.jo,
                      }));
                    }}
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
                    let newRelease = '';
                    let newRev = editFormData.rev;
                    const cleanTarget = cleanText(val);
                    const currentEditRev = String(editFormData.rev || '0').trim();
                    const exactMatch = rows.find(r => {
                      const cFull = cleanText(r.fullDeskripsi);
                      const cDwg = cleanText(r.noDwg);
                      const cName = cleanText(r.drawingName);
                      const rowRev = String(r.rev || '0').trim();
                      const isRevMatch = rowRev === currentEditRev;
                      const isNameMatch = cleanTarget.includes(cDwg) || cleanTarget.includes(cName) || cFull.includes(cleanTarget) || cleanTarget.includes(cFull);
                      return isNameMatch && isRevMatch;
                    });
                    if (exactMatch && exactMatch.finishDate) {
                      newRelease = parseToStandardDate(exactMatch.finishDate);
                      if (exactMatch.rev) newRev = exactMatch.rev;
                    } else {
                      newRelease = '';
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
                      onChange={(e) => {
                        const newRev = e.target.value;
                        const cleanProj = cleanText(editFormData.project || '');
                        const rows = drawingControlMap[cleanProj] || [];
                        let autoRel = '';
                        const exactMatch = rows.find(r => {
                          const cFull = cleanText(r.fullDeskripsi);
                          const cDwg = cleanText(r.noDwg);
                          const cName = cleanText(r.drawingName);
                          const rowRev = String(r.rev || '0').trim();
                          const isRevMatch = rowRev === String(newRev).trim();
                          const isNameMatch = cleanText(editFormData.taskName).includes(cDwg) || cleanText(editFormData.taskName).includes(cName) || cFull.includes(cleanText(editFormData.taskName));
                          return isNameMatch && isRevMatch;
                        });
                        if (exactMatch && exactMatch.finishDate) {
                          autoRel = parseToStandardDate(exactMatch.finishDate);
                        } else {
                          autoRel = '';
                        }
                        setEditFormData(prev => ({ ...prev, rev: newRev, release: autoRel }));
                      }}
                      placeholder="0"
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>
                  <div className="hidden">
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

      {isPlannerOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-2xl overflow-hidden shadow-2xl">
            <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <span className="font-bold text-xs text-white flex items-center gap-1.5">
                <FileCheck className="w-3.5 h-3.5 text-amber-400" /> Planner Panel — Approval Jobcard Semua Biro ({pendingTasksCount} Menunggu)
              </span>
              <button onClick={() => setIsPlannerOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4">
              <div className="space-y-3">
                <div className="flex items-center justify-between text-xs pb-2 border-b border-slate-800">
                  <span className="text-slate-400">Daftar Pengajuan Jobcard ({pendingTasksCount}) dari {pendingTasksByBiro.length} Biro</span>
                </div>
                <div className="max-h-96 overflow-y-auto space-y-4 pr-1">
                  {pendingTasksByBiro.length > 0 ? pendingTasksByBiro.map(([biroName, list]) => (
                    <div key={biroName} className="space-y-2">
                      <div className="flex items-center justify-between text-[11px] px-1">
                        <span className="font-bold text-slate-300">{biroName}</span>
                        <span className="font-mono text-amber-400">{list.length} menunggu</span>
                      </div>
                      {list.map((task, idx) => (
                        <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                          <div>
                            <div className="font-semibold text-white">{task.pic} <span className="font-mono text-slate-500">#{idx + 1}</span>{task.rev && <span className="ml-2 font-mono text-[10px] text-cyan-400">Rev.{task.rev}</span>}</div>
                            <div className="text-slate-400 text-[11px]">{task.taskName}</div>
                            <div className="text-emerald-400 font-mono text-[10px]">{task.project}</div>
                          </div>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <input type="text" value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                              onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                              placeholder="Jobcard (misal: JC020926 39833)..."
                              className="w-48 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase focus:outline-none" />
                            <button onClick={() => handleSaveKodeJcForTask(task.id)}
                              className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white font-bold rounded text-xs cursor-pointer">Simpan</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )) : <div className="py-8 text-center text-xs text-slate-500">Semua tugas sudah memiliki kode Jobcard</div>}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}