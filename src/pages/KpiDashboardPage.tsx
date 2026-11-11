import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { departmentsData } from '../data';
import { projectCandidates, dedupeRows } from '../lib/projectCandidates';
import { Building2, ArrowLeft, ChevronRight, RefreshCw, Users, Layers } from 'lucide-react';

interface KpiDashboardPageProps {
  user: UserSession;
  onLogout: () => void;
  onBack?: () => void;
}

/* ============================================================
   PENGATURAN ATURAN HITUNG (ubah di sini bila aturan berubah)
   ============================================================ */
const WEIGHT_A = 0.5;                 // bobot Efisiensi JO (50%)
const WEIGHT_B = 0.5;                 // bobot Release Drawing (50% - semua revisi)
const CAP_EFFICIENCY = Infinity;      // tanpa batas: A tampil apa adanya
const A_ONLY_FROM_RELEASED = false;   // false = A tetap dihitung walau drawing belum release
const A_ONLY_REV0 = false;            // true = jam kerja hanya dari drawing Rev.0
const REQUIRE_JOBCARD_CODE = true;    // jobcard tanpa kode (Menunggu Planner) belum valid
const COUNT_DRAFT_TIMESHEET = false;  // timesheet Draft tidak dihitung sebagai Real JO
const PARTIAL_KPI = false;            // true = bila hanya A atau B ada, KPI memakai komponen yang ada
const INCLUDE_REALISASI_JOBCARDS = true; // true = Jobcard yang hanya ada di file Realisasi JO (belum diinput lewat Form) ikut tampil sebagai penugasan
const HOUR_PLANNER_FILE = 'hourplanner.json';

const MASTER_BUCKET = 'master-files';
const OTHER_DEPT_ID = '__other__';

interface JobCardRow {
  id: string;
  biro: string;
  pic: string;
  project: string;
  taskName: string;
  startDate: string;
  endDate: string;
  plannedJo: number;
  kodeJc: string;
  rev: string;
  release: string;
  realJo: number;
  fromRealisasi?: boolean;   // true = diambil dari file Realisasi JO (tidak ada di job_cards)
}

// Anggota organik dari master IM4 (dipakai agar semua anggota biro tampil walau belum punya job card)
interface MemberRow {
  nama: string;
  biro: string;
  dept: string;
}

interface PersonScore {
  name: string;
  biro: string;
  cards: JobCardRow[];
  a: number | null;
  b: number | null;
  kpi: number | null;
  targetRev0: number;
  onTimeRev0: number;
  pendingCount: number;
}

interface BiroScore {
  id: string;
  name: string;
  persons: PersonScore[];
  kpi: number | null;
  a: number | null;
  b: number | null;
  scored: number;
}

interface DeptScore {
  id: string;
  name: string;
  biros: BiroScore[];
  kpi: number | null;
  a: number | null;
  b: number | null;
  personCount: number;
  scored: number;
}

interface ProjectScore {
  project: string;
  persons: PersonScore[];
  kpi: number | null;
  a: number | null;
  b: number | null;
  scored: number;
  total: number;
  released: number;
}

/* ============================ Helper ============================ */
function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

// Sama dengan logika di KabiroPage
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

function normalizeDate(val: any): string {
  if (!val) return '';
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return '';
}

function formatDisplayDate(val: any): string {
  const iso = normalizeDate(val);
  if (!iso) return '-';
  const [y, m, d] = iso.split('-');
  return `${d}-${m}-${y}`;
}

function getLocalMap(key: string): Record<string, any> {
  try { return JSON.parse(localStorage.getItem(key) || '{}'); } catch { return {}; }
}

function saveLocalRelease(id: string, val: string) {
  try {
    const map = JSON.parse(localStorage.getItem('task_release_map') || '{}');
    map[id] = val;
    localStorage.setItem('task_release_map', JSON.stringify(map));
  } catch {}
}

/* ===== Sinkronisasi Release Drawing otomatis dari Google Drive ===== */
const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';
const GOOGLE_DRIVE_SHEETS = [
  { id: '1CBL96MejQnswfg_sLIkwkK9K594S9hDWoSGDBIeIBNo', projectKey: 'M000313', title: 'Drawing Control GEOMARIN V (M000313)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000304', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000305', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '19qiMBb1b7qmeRXkFNTbh-eKtVEKDUTgshloa116HQ48', projectKey: 'W000308', title: 'Drawing Control LPD UAE USED (W000308)' },
  { id: '16rrjyU60BmzRIYFL3I1WRLkw_AfJPUwi8vZiJx2LYCg', projectKey: 'W000314', title: 'Drawing Control KSR 20 M ALU (W000314)' },
  { id: '1uv2D4ECJ0z-NWP8byzXTkkM17MKD6f0r-RCtKjpG5Hs', projectKey: 'M000312', title: 'Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312)' },
  { id: '1eTO-dvVd3psdP4mFQc2ZQW_Z-lyooCb2TVjR4mYTbDw', projectKey: 'W000311', title: 'Drawing Control KSSR 18 M (W000311)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKjpAMfo', projectKey: 'W000306', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000307', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1l0ZoTCctHrmkruIkC4Gk1NtoqsugMVGABrsf3yb6PHo', projectKey: 'KSOT', title: 'Drawing Control KSOT' },
  { id: '1jGNHh9HhSrPlEv7JZvtW8ToMU2ADrT0yE87tk9E53zY', projectKey: 'LCU', title: 'Drawing Control LCU LD PN 124 M' },
  { id: '1XsrklFhcUkjgfEjuT4jDRWx6Uufsj8Jt', projectKey: 'FFBNW', title: 'Drawing Control FFBNW FRIGATE 140 M' },
];

interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
}

const indoMonthsMap: Record<string, string> = {
  jan: '01', januari: '01', january: '01', feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03', apr: '04', april: '04', mei: '05', may: '05',
  jun: '06', juni: '06', june: '06', jul: '07', juli: '07', july: '07',
  agu: '08', ags: '08', agustus: '08', aug: '08', august: '08',
  sep: '09', september: '09', okt: '10', oktober: '10', oct: '10', october: '10',
  nov: '11', november: '11', des: '12', desember: '12', dec: '12', december: '12'
};

function parseDriveDate(val: any): string {
  if (!val) return '';
  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'nan') return '';
  str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
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
    return new Date(Math.round((num - 25569) * 86400 * 1000)).toISOString().split('T')[0];
  }
  return '';
}

const kpiDrawingCache = new Map<string, DrawingControlRow[]>();

function isKnownDriveProject(projectCode: string): boolean {
  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  return cands.some(c => {
    const cp = cleanText(c);
    return cp && GOOGLE_DRIVE_SHEETS.some(s => cleanText(s.projectKey) === cp || cp.includes(cleanText(s.projectKey)));
  });
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
          noDwg: String(item.noDwg || '').trim(),
          drawingName: String(item.drawingName || '').trim(),
          fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
          rev: String(item.rev || '0').trim(),
          finishDate: String(item.finishDate || '').trim(),
        }));
      }
    }
  } catch {}
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
          const noDwg = String(row[0] || '').trim();
          const dwgName = String(row[1] || '').trim();
          if (!noDwg && !dwgName) continue;
          rows.push({
            noDwg, drawingName: dwgName,
            fullDeskripsi: (noDwg && dwgName) ? `${noDwg}-${dwgName}` : (dwgName || noDwg),
            rev: String(row[14] || '0').trim(),
            finishDate: String(row[15] || '').trim(),
          });
        }
        if (rows.length > 0) return rows;
      }
    } catch {}
  }
  return [];
}

async function fetchDrawingRowsFor(projectCode: string): Promise<DrawingControlRow[]> {
  const cacheKey = cleanText(projectCode);
  if (cacheKey && kpiDrawingCache.has(cacheKey)) {
    const cached = kpiDrawingCache.get(cacheKey)!;
    if (cached.length > 0) return cached;
  }
  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  const results = await Promise.all(cands.map(c => fetchDrawingRowsRaw(c).catch(() => [] as DrawingControlRow[])));
  const deduped = dedupeRows(results.flat());
  if (cacheKey && deduped.length > 0) kpiDrawingCache.set(cacheKey, deduped);
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
    if (!/^\d{4}-\d{2}-\d{2}$/.test(parseDriveDate(r.finishDate))) return false;
    const rKey = drawingKey(r.noDwg);
    if (wanted && rKey === wanted) return true;
    const cDwg = cleanText(r.noDwg);
    if (cDwg && cDwg.length >= 5 && cleanTarget.includes(cDwg)) return true;
    return false;
  });
  return hit ? parseDriveDate(hit.finishDate) : '';
}

function parseNumber(val: any): number {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const n = parseFloat(String(val).trim().replace('%', '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function avg(values: (number | null)[]): number | null {
  const v = values.filter((x): x is number => x !== null);
  return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null;
}

function getCategory(score: number | null): { label: string; color: string } {
  if (score === null) return { label: 'Belum lengkap', color: '#64748b' };
  if (score >= 90) return { label: 'Excellent', color: '#34d399' };
  if (score >= 80) return { label: 'Good', color: '#4cc2ff' };
  if (score >= 70) return { label: 'Fair', color: '#fbbf24' };
  return { label: 'Poor', color: '#fb7185' };
}

const fmt = (v: number | null, digits = 1) => (v === null ? '—' : v.toFixed(digits).replace(/\.0$/, ''));

/* Real JO dari file Realisasi JO.xlsx, kunci = kode jobcard */
function buildRealisasiMap(wb: XLSX.WorkBook | null): Map<string, number> {
  const map = new Map<string, number>();
  if (!wb) return map;
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName]; if (!sheet) return;
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows.length) return;
    let headerIdx = -1, jcCol = -1, effCol = -1, otCol = -1;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      const row = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      let j = row.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
      if (j === -1) j = row.findIndex(c => c.includes('jobcard'));
      const e = row.findIndex(c => c.includes('effective'));
      const o = row.findIndex(c => c.includes('overtime'));
      if (j !== -1 && (e !== -1 || o !== -1)) { headerIdx = r; jcCol = j; effCol = e; otCol = o; break; }
    }
    if (jcCol === -1) jcCol = 16;
    if (effCol === -1) effCol = 11;
    if (otCol === -1) otCol = 12;
    for (let r = Math.max(headerIdx + 1, 1); r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const rawJc = String(row[jcCol] || '').trim();
      if (!rawJc || rawJc.toLowerCase() === 'nan' || rawJc.toLowerCase().includes('jobcard')) continue;
      const total = parseNumber(row[effCol]) + parseNumber(row[otCol]);
      const key = cleanText(rawJc);
      if (key) map.set(key, (map.get(key) || 0) + total);
    }
  });
  return map;
}

/* Pencocokan nama longgar (sama dengan RendalPage): "Fendy" = "Fendi", "Luqman Chakim" = "Lukman Hakim", dst. */
const normNameToken = (t: string) =>
  t.replace(/ch/g, 'h').replace(/(.)\1+/g, '$1').replace(/^mu/, 'mo').replace(/y$/, 'i').replace(/q/g, 'k');
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

/* Baris Realisasi JO per pegawai (dipakai agar Real JO dihitung per orang walau satu kode Jobcard dipakai bersama) */
interface RealisasiRow {
  nama: string;
  project: string;
  jobcard: string;
  taskName: string;
  hours: number;
}

function splitUserCell(raw: any): string {
  const s = String(raw ?? '').trim();
  const m = s.match(/^(\d[\d\s]*?)\s*-\s*(.+)$/);
  return m ? m[2].trim() : s;
}

function extractProjectCodeFromCell(raw: any): string {
  const s = String(raw ?? '').trim();
  if (!s || s === '-' || s.toLowerCase() === 'nan') return '';
  return (s.split(/[\s-]+/)[0] || s).trim().toUpperCase();
}

function parseRealisasiRows(wb: XLSX.WorkBook | null): RealisasiRow[] {
  const out: RealisasiRow[] = [];
  if (!wb) return out;
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName]; if (!sheet) return;
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows.length) return;
    let headerIdx = -1, cName = -1, cProj = -1, cJc = -1, cEff = -1, cOt = -1, cTask = -1;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const h = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      const jc = h.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
      const ef = h.findIndex(c => c.includes('effective'));
      const ot = h.findIndex(c => c.includes('overtime'));
      if (jc !== -1 && (ef !== -1 || ot !== -1)) {
        headerIdx = r; cJc = jc; cEff = ef; cOt = ot;
        cName = h.findIndex(c => c === 'employee/name' || c === 'nama' || c === 'name' || c === 'user' || c === 'employee' || c === 'pegawai');
        cProj = h.findIndex(c => c === 'project/code' || c === 'project' || c === 'proyek');
        cTask = h.findIndex(c => c === 'work order/description' || c === 'jobcard/task name' || c === 'description' || c === 'deskripsi' || c === 'task name');
        break;
      }
    }
    if (headerIdx === -1 || cName === -1) return;   // tanpa kolom nama -> tidak bisa per pegawai
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const nama = splitUserCell(row[cName]);
      if (!nama) continue;
      const jcRaw = String(row[cJc] ?? '').trim();
      if (!jcRaw || jcRaw.toLowerCase() === 'nan' || jcRaw.toLowerCase().includes('jobcard')) continue;
      out.push({
        nama,
        project: cProj !== -1 ? extractProjectCodeFromCell(row[cProj]) : '',
        jobcard: jcRaw.replace(/\s+/g, ' ').toUpperCase(),
        taskName: cTask !== -1 ? String(row[cTask] ?? '').trim() : '',
        hours: (cEff !== -1 ? Math.abs(parseNumber(row[cEff])) : 0) + (cOt !== -1 ? Math.abs(parseNumber(row[cOt])) : 0),
      });
    }
  });
  return out;
}

async function loadHourPlanner(): Promise<Record<string, { hour: number; start?: string; end?: string }>> {
  try {
    const { data, error } = await supabase.storage.from(MASTER_BUCKET).download(HOUR_PLANNER_FILE);
    if (error || !data) return {};
    const parsed = JSON.parse(await data.text());
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch { return {}; }
}

/* Baca master IM4: daftar nama organik & outsourcing, plus daftar anggota organik beserta biro/departemennya
   (penanda biro/departemen berjalan, logikanya sama dengan RendalPage) */
function parseMemberNames(wb: XLSX.WorkBook | null): { organic: Set<string>; outsourcing: Set<string>; members: MemberRow[] } {
  const organic = new Set<string>();
  const outsourcing = new Set<string>();
  const members: MemberRow[] = [];
  if (!wb) return { organic, outsourcing, members };
  try {
    const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes('education')) || wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName]; if (!sheet) return { organic, outsourcing, members };
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    let headerIdx = -1;
    for (let r = 0; r < Math.min(15, rows.length); r++) {
      const vals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      if (vals.includes('nama') && (vals.includes('nip') || vals.includes('status') || vals.includes('jabatan'))) { headerIdx = r; break; }
    }
    if (headerIdx === -1) return { organic, outsourcing, members };
    const headers = rows[headerIdx].map(v => String(v).trim().toLowerCase());
    const namaCol = headers.findIndex(h => h === 'nama');
    const statusCol = headers.findIndex(h => h === 'status');
    const unitCol = headers.findIndex(h => h.includes('unit'));
    const jabatanCol = headers.findIndex(h => h.includes('jabatan'));
    let currentDept = '', currentBiro = '';
    const seen = new Set<string>();
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const nama = String(row[namaCol] || '').trim();
      if (!nama || nama.toLowerCase() === 'nan' || nama.toLowerCase() === 'nama') continue;
      const unit = String(row[unitCol] || '').trim();
      const jabatan = String(row[jabatanCol] || '').trim();
      const jl = jabatan.toLowerCase();
      const hasUnit = unit !== '' && unit.toLowerCase() !== 'nan';

      if (jl.includes('kepala divisi')) { currentDept = 'Div. Desain'; currentBiro = 'Div. Desain'; }
      else if (jl.includes('kepala departemen') || jl.includes('kadep')) {
        currentDept = hasUnit ? unit : jabatan;
        currentBiro = `Staf ${currentDept}`;
      } else if (jl.includes('kepala biro') || jl.includes('kabiro')) {
        currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').replace(/Kabiro/gi, 'Biro').trim();
        if (hasUnit) currentDept = unit;
      }

      const status = String(row[statusCol] || '').trim().toLowerCase();
      const key = cleanText(nama);
      if (status.includes('outsourcing')) {
        outsourcing.add(key);
      } else {
        organic.add(key);
        if (!seen.has(key)) { seen.add(key); members.push({ nama, biro: currentBiro, dept: currentDept }); }
      }
    }
  } catch { /* abaikan */ }
  return { organic, outsourcing, members };
}

async function loadIm4Workbook(): Promise<XLSX.WorkBook | null> {
  try {
    const { data, error } = await supabase.storage.from(MASTER_BUCKET).download('im4.xlsx');
    if (!error && data) return XLSX.read(await data.arrayBuffer(), { type: 'array' });
  } catch { /* lanjut ke fallback */ }
  for (const p of ['/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx']) {
    try {
      const res = await fetch(p);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      const b = new Uint8Array(buf.slice(0, 4));
      if (b[0] === 80 && b[1] === 75 && b[2] === 3 && b[3] === 4) return XLSX.read(buf, { type: 'array' });
    } catch { /* abaikan */ }
  }
  return null;
}

/* ====================== Perhitungan skor ====================== */
function scorePerson(name: string, biro: string, cards: JobCardRow[]): PersonScore {
  const valid = cards.filter(c => !REQUIRE_JOBCARD_CODE || c.kodeJc.trim() !== '');

  // A = Planned JO / Actual JO x 100%
  const aCards = valid.filter(c =>
    c.plannedJo > 0 && c.realJo > 0 &&
    (!A_ONLY_FROM_RELEASED || c.release !== '') &&
    (!A_ONLY_REV0 || Number(c.rev) === 0)
  );
  const sumPlan = aCards.reduce((s, c) => s + c.plannedJo, 0);
  const sumReal = aCards.reduce((s, c) => s + c.realJo, 0);
  const a = aCards.length ? Math.min((sumPlan / sumReal) * 100, CAP_EFFICIENCY) : null;

  // B = jumlah drawing yang sudah ada tanggal release / total list drawing x 100% (semua revisi)
  const targets = valid.filter(c => !c.fromRealisasi);   // jobcard dari Realisasi saja tidak punya daftar drawing
  const onTime = targets.filter(c => c.release !== '');   // = drawing yang sudah release
  const b = targets.length ? (onTime.length / targets.length) * 100 : null;

  let kpi: number | null = null;
  if (a !== null && b !== null) kpi = WEIGHT_A * a + WEIGHT_B * b;
  else if (PARTIAL_KPI && a !== null) kpi = a;
  else if (PARTIAL_KPI && b !== null) kpi = b;

  return {
    name, biro, cards, a, b, kpi,
    targetRev0: targets.length,
    onTimeRev0: onTime.length,
    pendingCount: cards.length - valid.length,
  };
}

// memberNames = anggota biro dari master IM4 yang belum punya job card; tetap ditampilkan (skor "Belum lengkap")
function buildBiroScore(id: string, name: string, cards: JobCardRow[], memberNames: string[] = []): BiroScore {
  const byPerson = new Map<string, JobCardRow[]>();
  cards.forEach(c => {
    const pk = cleanText(c.pic);
    if (!byPerson.has(pk)) byPerson.set(pk, []);
    byPerson.get(pk)!.push(c);
  });
  const persons: PersonScore[] = [];
  byPerson.forEach(list => persons.push(scorePerson(list[0].pic, name, list)));

  const seen = new Set(byPerson.keys());
  memberNames.forEach(n => {
    const k = cleanText(n);
    if (!seen.has(k)) { seen.add(k); persons.push(scorePerson(n, name, [])); }
  });

  persons.sort((x, y) => (y.kpi ?? -1) - (x.kpi ?? -1) || x.name.localeCompare(y.name));
  return {
    id, name, persons,
    kpi: avg(persons.map(p => p.kpi)),
    a: avg(persons.map(p => p.a)),
    b: avg(persons.map(p => p.b)),
    scored: persons.filter(p => p.kpi !== null).length,
  };
}

function buildDeptScore(id: string, name: string, biros: BiroScore[]): DeptScore {
  const allPersons = biros.flatMap(b => b.persons);
  return {
    id, name, biros,
    kpi: avg(allPersons.map(p => p.kpi)),
    a: avg(allPersons.map(p => p.a)),
    b: avg(allPersons.map(p => p.b)),
    personCount: allPersons.length,
    scored: allPersons.filter(p => p.kpi !== null).length,
  };
}

/* KPI per proyek untuk sekumpulan jobcard (biro / departemen / divisi).
   Tiap orang dihitung KPI-nya di proyek itu, lalu dirata-rata. */
function buildProjectScores(cards: JobCardRow[]): ProjectScore[] {
  const byProj = new Map<string, Map<string, JobCardRow[]>>();
  cards.forEach(c => {
    const proj = c.project || '(tanpa proyek)';
    if (!byProj.has(proj)) byProj.set(proj, new Map());
    const pm = byProj.get(proj)!;
    const pk = cleanText(c.pic);
    if (!pm.has(pk)) pm.set(pk, []);
    pm.get(pk)!.push(c);
  });
  const out: ProjectScore[] = [];
  byProj.forEach((pm, project) => {
    const persons: PersonScore[] = [];
    pm.forEach(list => persons.push(scorePerson(list[0].pic, list[0].biro, list)));
    persons.sort((x, y) => (y.kpi ?? -1) - (x.kpi ?? -1) || x.name.localeCompare(y.name));
    let total = 0, released = 0;
    persons.forEach(pr => {
      const valid = pr.cards.filter(c => (!REQUIRE_JOBCARD_CODE || c.kodeJc.trim() !== '') && !c.fromRealisasi);
      total += valid.length;
      released += valid.filter(c => c.release !== '').length;
    });
    out.push({
      project, persons,
      kpi: avg(persons.map(x => x.kpi)),
      a: avg(persons.map(x => x.a)),
      b: avg(persons.map(x => x.b)),
      scored: persons.filter(x => x.kpi !== null).length,
      total, released,
    });
  });
  return out.sort((x, y) => x.project.localeCompare(y.project));
}

/* ========================= Komponen UI ========================= */
function Donut({ value, size = 120, color, text, stroke = 10 }: {
  value: number | null; size?: number; color: string; text?: string; stroke?: number;
}) {
  const r = 54;
  const c = 2 * Math.PI * r;
  const dash = value === null ? 0 : (Math.max(0, Math.min(100, value)) / 100) * c;
  const label = text ?? (value === null ? '—' : String(Math.round(value)));
  return (
    <svg width={size} height={size} viewBox="0 0 120 120" role="img" aria-label={label}>
      <circle cx="60" cy="60" r={r} fill="none" stroke="#24324f" strokeWidth={stroke} />
      {dash > 0 && (
        <circle cx="60" cy="60" r={r} fill="none" stroke={color} strokeWidth={stroke} strokeLinecap="round"
          strokeDasharray={`${dash} ${c}`} transform="rotate(-90 60 60)" />
      )}
      <text x="60" y="60" textAnchor="middle" dominantBaseline="central" fontSize={size >= 150 ? 34 : size >= 100 ? 26 : 20}
        fontWeight={800} fill="#e6edf7">{label}</text>
    </svg>
  );
}

function Pill({ score }: { score: number | null }) {
  const cat = getCategory(score);
  return (
    <span className="text-[10px] font-bold px-2.5 py-0.5 rounded-full text-slate-950 whitespace-nowrap" style={{ background: cat.color }}>
      {cat.label}
    </span>
  );
}

// Jumlah Plan JO dan Real JO dari sekumpulan jobcard (hanya yang sudah punya kode Jobcard bila aturannya demikian)
function sumJo(cards: JobCardRow[]) {
  const valid = cards.filter(c => !REQUIRE_JOBCARD_CODE || c.kodeJc.trim() !== '');
  return {
    plan: valid.reduce((s, c) => s + c.plannedJo, 0),
    real: valid.reduce((s, c) => s + c.realJo, 0),
  };
}

const round2 = (v: number) => Math.round(v * 100) / 100;

function ProjectBarChart({ rows, title, className = '' }: { rows: ProjectScore[]; title: string; className?: string }) {
  const H = 180;
  const [selected, setSelected] = useState<string | null>(null);

  const detail = useMemo(() => {
    const r = rows.find(x => x.project === selected);
    if (!r) return null;
    const total = sumJo(r.persons.flatMap(p => p.cards));
    const persons = r.persons
      .map(p => ({ name: p.name, biro: p.biro, ...sumJo(p.cards) }))
      .filter(p => p.plan > 0 || p.real > 0)
      .sort((x, y) => y.plan - x.plan);
    return { r, ...total, persons };
  }, [rows, selected]);

  const pctOf = (plan: number, real: number) => (plan > 0 ? (real / plan) * 100 : null);

  return (
    <div className={`bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3 min-w-0 ${className}`}>
      <div className="flex items-center justify-between flex-wrap gap-2">
        <h2 className="text-sm font-bold text-white flex items-center gap-2">
          <Layers className="w-4 h-4 text-blue-400" /> {title} ({rows.length} proyek)
        </h2>
        <span className="text-[10px] text-slate-500">Rata-rata KPI individu per proyek · klik batang untuk detail Plan JO &amp; Real JO</span>
      </div>
      {rows.length === 0 ? (
        <div className="text-xs text-slate-500 italic py-10 text-center">Belum ada data proyek pada periode ini.</div>
      ) : (
        <div className="overflow-x-auto pb-1">
          <div className="flex gap-2" style={{ minWidth: rows.length * 64 + 36 }}>
            {/* Sumbu Y */}
            <div className="relative shrink-0 w-7" style={{ height: H }}>
              {[100, 75, 50, 25, 0].map(t => (
                <span key={t} className="absolute right-0 text-[9px] text-slate-500 -translate-y-1/2" style={{ top: `${100 - t}%` }}>{t}</span>
              ))}
            </div>
            <div className="flex-1 min-w-0">
              {/* Area batang */}
              <div className="relative flex gap-3" style={{ height: H }}>
                {[0, 25, 50, 75, 100].map(t => (
                  <div key={t} className="absolute left-0 right-0 border-t border-slate-800 pointer-events-none" style={{ top: `${100 - t}%` }} />
                ))}
                {rows.map(r => {
                  const cat = getCategory(r.kpi);
                  const isSel = selected === r.project;
                  const dim = selected !== null && !isSel;
                  const tip = `${r.project}\nKPI ${fmt(r.kpi)} (${cat.label})\nA ${fmt(r.a)}${r.a !== null ? '%' : ''} · B ${fmt(r.b)}${r.b !== null ? '%' : ''}\n${r.released}/${r.total} drawing release\n${r.scored}/${r.persons.length} personel terhitung\n(klik untuk detail Plan JO & Real JO)`;
                  return (
                    <div key={r.project} title={tip}
                      onClick={() => setSelected(prev => (prev === r.project ? null : r.project))}
                      className="relative flex-1 min-w-[52px] h-full cursor-pointer transition-opacity"
                      style={{ opacity: dim ? 0.45 : 1 }}>
                      {r.kpi !== null ? (
                        <>
                          <div className={`absolute bottom-0 left-1/2 -translate-x-1/2 w-full max-w-[44px] rounded-t-md ${isSel ? 'ring-2 ring-white' : ''}`}
                            style={{ height: `${Math.max(0, Math.min(100, r.kpi))}%`, background: cat.color }} />
                          <span className="absolute left-0 right-0 text-center text-[10px] font-bold text-slate-100"
                            style={{ bottom: `calc(${Math.max(0, Math.min(100, r.kpi))}% + 2px)` }}>{fmt(r.kpi, 0)}</span>
                        </>
                      ) : (
                        <span className="absolute left-0 right-0 bottom-1 text-center text-[10px] text-slate-500">—</span>
                      )}
                    </div>
                  );
                })}
              </div>
              {/* Label proyek */}
              <div className="flex gap-3 mt-1.5">
                {rows.map(r => (
                  <div key={r.project} title={r.project}
                    onClick={() => setSelected(prev => (prev === r.project ? null : r.project))}
                    className={`flex-1 min-w-[52px] text-center text-[10px] font-mono truncate cursor-pointer ${selected === r.project ? 'text-white font-bold' : 'text-emerald-400'}`}>
                    {r.project}
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ===== Detail proyek yang diklik ===== */}
      {detail && (() => {
        const pct = pctOf(detail.plan, detail.real);
        const over = pct !== null && pct > 100;
        return (
          <div className="p-4 bg-slate-950 border border-slate-800 rounded-xl space-y-4">
            <div className="flex items-center justify-between gap-3">
              <div className="font-mono font-bold text-emerald-400 text-sm">{detail.r.project}</div>
              <button onClick={() => setSelected(null)}
                className="px-2.5 py-1 text-[11px] bg-slate-800 hover:bg-slate-700 text-slate-300 rounded border border-slate-700 cursor-pointer">Tutup</button>
            </div>

            <div className="grid grid-cols-3 gap-3 text-center">
              <div className="bg-slate-900 border border-slate-800 rounded-lg p-3">
                <div className="text-lg font-extrabold font-mono text-violet-300">{detail.plan ? round2(detail.plan) : '-'}</div>
                <div className="text-[10px] text-slate-400">Plan JO (jam)</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-lg p-3">
                <div className="text-lg font-extrabold font-mono text-emerald-400">{detail.real ? round2(detail.real) : '-'}</div>
                <div className="text-[10px] text-slate-400">Real JO (jam)</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-lg p-3">
                <div className={`text-lg font-extrabold font-mono ${over ? 'text-rose-300' : 'text-white'}`}>{pct === null ? '—' : `${Math.round(pct)}%`}</div>
                <div className="text-[10px] text-slate-400">Real JO ÷ Plan JO</div>
              </div>
            </div>

            {pct !== null && (
              <div className="space-y-1">
                <div className="h-2.5 bg-slate-800 rounded-full overflow-hidden">
                  <div className="h-2.5 rounded-full" style={{ width: `${Math.min(pct, 100)}%`, background: over ? '#fb7185' : '#34d399' }} />
                </div>
                <div className="text-[10px] text-slate-500">
                  {over ? 'Jam realisasi melebihi rencana (di atas 100%).' : 'Jam realisasi dibanding rencana.'}
                </div>
              </div>
            )}

            <div className="space-y-1.5">
              <div className="text-[11px] font-bold text-slate-300">Per personel ({detail.persons.length})</div>
              {detail.persons.length === 0 && (
                <div className="text-[11px] text-slate-500 italic">Belum ada Plan JO atau Real JO pada proyek ini.</div>
              )}
              {detail.persons.map(p => {
                const pp = pctOf(p.plan, p.real);
                return (
                  <div key={p.name} className="grid grid-cols-12 gap-2 items-center text-[11px] py-1 border-b border-slate-800/60 last:border-0">
                    <div className="col-span-5 text-slate-200 truncate" title={`${p.name} · ${p.biro}`}>{p.name}</div>
                    <div className="col-span-3 text-right font-mono text-violet-300">{p.plan ? round2(p.plan) : '-'}</div>
                    <div className="col-span-2 text-right font-mono text-emerald-400">{p.real ? round2(p.real) : '-'}</div>
                    <div className={`col-span-2 text-right font-mono font-bold ${pp !== null && pp > 100 ? 'text-rose-300' : 'text-slate-200'}`}>
                      {pp === null ? '—' : `${Math.round(pp)}%`}
                    </div>
                  </div>
                );
              })}
              <div className="grid grid-cols-12 gap-2 text-[10px] text-slate-500 pt-1">
                <div className="col-span-5" />
                <div className="col-span-3 text-right">Plan JO</div>
                <div className="col-span-2 text-right">Real JO</div>
                <div className="col-span-2 text-right">%</div>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

/* ============================= Halaman ============================= */
export default function KpiDashboardPage({ user, onLogout, onBack }: KpiDashboardPageProps) {
  const [cards, setCards] = useState<JobCardRow[]>([]);
  const [members, setMembers] = useState<MemberRow[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState('');
  const [period, setPeriod] = useState('all');
  const [selectedDept, setSelectedDept] = useState<string | null>(null);
  const [selectedBiro, setSelectedBiro] = useState<string | null>(null);
  const [selectedPerson, setSelectedPerson] = useState<string | null>(null);
  const [expandedProjects, setExpandedProjects] = useState<Record<string, boolean>>({});
  const toggleProject = (k: string) => setExpandedProjects(prev => ({ ...prev, [k]: !prev[k] }));

  const loadData = useCallback(async () => {
    setIsLoading(true);
    setErrorMsg('');
    try {
      const [jcRes, tsRes, lineRes, wbRes, im4Wb, hourPlanner] = await Promise.all([
        supabase.from('job_cards').select('*'),
        supabase.from('outsourcing_timesheets').select('id, status'),
        supabase.from('outsourcing_timesheet_lines').select('timesheet_id, work_order_code, effective_hours, overtime_hours'),
        supabase.storage.from(MASTER_BUCKET).download('realisasi.xlsx'),
        loadIm4Workbook(),
        loadHourPlanner(),
      ]);

      if (jcRes.error) throw jcRes.error;

      // Daftar pegawai organik (dan anggota per biro) dari master IM4
      const { organic, outsourcing, members: memberList } = parseMemberNames(im4Wb);
      setMembers(memberList);
      if (organic.size === 0) {
        setCards([]);
        setErrorMsg('Master IM4 belum termuat, sehingga daftar pegawai organik tidak bisa ditentukan. Upload master IM4 terlebih dahulu.');
        return;
      }

      // Real JO dari Excel Realisasi JO
      let excelMap = new Map<string, number>();
      let realRows: RealisasiRow[] = [];
      if (!wbRes.error && wbRes.data) {
        try {
          const realWb = XLSX.read(await wbRes.data.arrayBuffer(), { type: 'array' });
          excelMap = buildRealisasiMap(realWb);
          realRows = parseRealisasiRows(realWb);
        } catch { /* abaikan */ }
      }

      // Cocokkan nama (job card / Realisasi) ke nama anggota organik IM4 (persis dulu, lalu longgar)
      const memberByKey = new Map<string, MemberRow>(memberList.map(m => [cleanText(m.nama), m] as [string, MemberRow]));
      const resolveMember = (name: string): MemberRow | null => {
        const k = cleanText(name);
        if (!k || outsourcing.has(k)) return null;
        const direct = memberByKey.get(k);
        if (direct) return direct;
        const cands = memberList.filter(m => sameNameFuzzy(m.nama, name));
        return cands.length === 1 ? cands[0] : null;
      };

      // Jam Realisasi per pegawai per Jobcard (kunci: nama IM4 | kode Jobcard)
      const personHours = new Map<string, number>();
      const jcWithNamedRows = new Set<string>();
      realRows.forEach(r => {
        const jcKey = cleanText(r.jobcard);
        jcWithNamedRows.add(jcKey);
        const m = resolveMember(r.nama);
        const who = cleanText(m ? m.nama : r.nama);
        const k = `${who}|${jcKey}`;
        personHours.set(k, (personHours.get(k) || 0) + r.hours);
      });
      const realHoursFor = (pic: string, jcKey: string): number => {
        if (!jcKey) return 0;
        if (!jcWithNamedRows.has(jcKey)) return excelMap.get(jcKey) || 0;   // file tanpa kolom nama
        return personHours.get(`${cleanText(pic)}|${jcKey}`) || 0;          // hanya jam milik orang ini
      };
      const planFromHourPlanner = (jcKey: string): number => {
        const hp = jcKey ? hourPlanner[jcKey] : undefined;
        return hp ? parseNumber(hp.hour) : 0;
      };

      // Real JO dari timesheet (kunci = kode Work Order = kode_jc)
      const statusById = new Map<string, string>();
      (tsRes.data || []).forEach((t: any) => statusById.set(t.id, t.status || 'Draft'));
      const tsMap = new Map<string, number>();
      (lineRes.data || []).forEach((l: any) => {
        const st = statusById.get(l.timesheet_id) || 'Draft';
        if (!COUNT_DRAFT_TIMESHEET && st === 'Draft') return;
        const key = cleanText(l.work_order_code || '');
        if (!key) return;
        tsMap.set(key, (tsMap.get(key) || 0) + parseNumber(l.effective_hours) + parseNumber(l.overtime_hours));
      });

      const revLocal = getLocalMap('task_rev_map');
      const releaseLocal = getLocalMap('task_release_map');

      const rows: JobCardRow[] = (jcRes.data || []).map((row: any) => {
        const kodeJc = String(row.kode_jc || '').trim();
        const key = cleanText(kodeJc);
        const rawRev = row.rev !== undefined && row.rev !== null && String(row.rev).trim() !== ''
          ? String(row.rev) : String(revLocal[row.id] ?? '0');
        const rawRelease = row.release !== undefined && row.release !== null && String(row.release).trim() !== ''
          ? row.release : releaseLocal[row.id];
        const rawPic = String(row.pic || row.personil_name || '').trim();
        const member = resolveMember(rawPic);
        const pic = member ? member.nama : rawPic;
        const plannedJo = parseNumber(row.jo) || planFromHourPlanner(key);
        const hp = key ? hourPlanner[key] : undefined;
        return {
          id: row.id,
          biro: row.biro_name || 'Tanpa Biro',
          pic,
          project: row.project || '',
          taskName: row.task_name || '',
          startDate: normalizeDate(row.start_date || hp?.start),
          endDate: normalizeDate(row.end_date || hp?.end),
          plannedJo,
          kodeJc,
          rev: rawRev.trim() || '0',
          release: normalizeDate(rawRelease),
          realJo: key ? realHoursFor(pic, key) + (tsMap.get(key) || 0) : 0,
        };
      }).filter((r: JobCardRow) => {
        if (r.pic === '') return false;
        const k = cleanText(r.pic);
        return organic.has(k) && !outsourcing.has(k);   // hanya pegawai organik
      });

      // Jobcard yang ada di Realisasi JO tetapi belum diinput lewat Form -> tetap tampil sebagai penugasan
      if (INCLUDE_REALISASI_JOBCARDS && realRows.length > 0) {
        const existing = new Set(rows.filter(r => r.kodeJc).map(r => `${cleanText(r.pic)}|${cleanText(r.kodeJc)}`));
        const extra = new Map<string, JobCardRow>();
        realRows.forEach(r => {
          const m = resolveMember(r.nama);
          if (!m) return;   // bukan pegawai organik IM4
          const jcKey = cleanText(r.jobcard);
          const k = `${cleanText(m.nama)}|${jcKey}`;
          if (existing.has(k) || extra.has(k)) return;
          const hp = hourPlanner[jcKey];
          extra.set(k, {
            id: `realisasi-${k}`,
            biro: m.biro || 'Tanpa Biro',
            pic: m.nama,
            project: r.project,
            taskName: r.taskName,
            startDate: normalizeDate(hp?.start),
            endDate: normalizeDate(hp?.end),
            plannedJo: planFromHourPlanner(jcKey),
            kodeJc: r.jobcard,
            rev: '0',
            release: '',
            realJo: realHoursFor(m.nama, jcKey) + (tsMap.get(jcKey) || 0),
            fromRealisasi: true,
          });
        });
        rows.push(...extra.values());
      }

      setCards(rows);

      // Sinkronisasi otomatis tanggal Release dari Google Drive untuk kartu yang belum punya release
      const missingRelease = rows.filter(r => r.project && r.taskName && !r.release);
      if (missingRelease.length > 0) {
        const projects = Array.from(new Set(missingRelease.map(r => r.project))).sort((a, b) =>
          (isKnownDriveProject(b) ? 1 : 0) - (isKnownDriveProject(a) ? 1 : 0)
        );
        (async () => {
          for (let i = 0; i < projects.length; i += 4) {
            const batchProjs = projects.slice(i, i + 4);
            const rowsByProj = new Map<string, DrawingControlRow[]>();
            await Promise.all(batchProjs.map(async p => {
              try { rowsByProj.set(p, await fetchDrawingRowsFor(p)); } catch { rowsByProj.set(p, []); }
            }));
            const batchFound = new Map<string, string>();
            missingRelease.filter(r => batchProjs.includes(r.project)).forEach(r => {
              const rel = findReleaseInRows(rowsByProj.get(r.project) || [], r.taskName, r.rev || '0');
              if (rel) batchFound.set(r.id, rel);
            });
            if (batchFound.size > 0) {
              setCards(prev => prev.map(c => batchFound.has(c.id) ? { ...c, release: batchFound.get(c.id)! } : c));
              const ids = Array.from(batchFound.keys());
              for (let j = 0; j < ids.length; j += 10) {
                await Promise.all(ids.slice(j, j + 10).map(async id => {
                  const rel = batchFound.get(id)!;
                  saveLocalRelease(id, rel);
                  if (!id.startsWith('realisasi-')) {
                    try { await supabase.from('job_cards').update({ release: rel }).eq('id', id); } catch {}
                  }
                }));
              }
            }
          }
        })();
      }
    } catch (err: any) {
      console.error('Gagal memuat data KPI:', err);
      setErrorMsg(err?.message || 'Gagal memuat data KPI.');
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => { loadData(); }, [loadData]);

  const periodOptions = useMemo(() => {
    const set = new Set<string>();
    cards.forEach(c => {
      const refDate = c.endDate || c.startDate;
      if (refDate) set.add(refDate.slice(0, 7));
    });
    return Array.from(set).sort().reverse();
  }, [cards]);

  const periodLabel = (p: string) => {
    const [y, m] = p.split('-');
    const names = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return `${names[Number(m) - 1]} ${y}`;
  };

  // Struktur: Departemen -> Biro -> Individu
  const deptScores = useMemo<DeptScore[]>(() => {
    const filtered = period === 'all' ? cards : cards.filter(c => (c.endDate || c.startDate).startsWith(period));
    const depts = departmentsData || [];

    // Setiap job card dipetakan ke SATU biro master (yang pertama cocok) agar tidak terhitung ganda
    const bucket = new Map<string, JobCardRow[]>();       // key: deptId|biroId
    const unmapped = new Map<string, JobCardRow[]>();     // key: nama biro mentah

    filtered.forEach(c => {
      let placed = false;
      for (const d of depts) {
        const b = d.biros.find(bb => isBiroMatch(c.biro, bb.name));
        if (b) {
          const k = `${d.id}|${b.id}`;
          if (!bucket.has(k)) bucket.set(k, []);
          bucket.get(k)!.push(c);
          placed = true;
          break;
        }
      }
      if (!placed) {
        if (!unmapped.has(c.biro)) unmapped.set(c.biro, []);
        unmapped.get(c.biro)!.push(c);
      }
    });

    // Anggota IM4 yang belum punya job card pada periode ini -> dipetakan ke biro master lewat biro di IM4
    const withCards = new Set(filtered.map(c => cleanText(c.pic)));
    const memberBucket = new Map<string, string[]>();     // key: deptId|biroId
    const memberUnmapped = new Map<string, string[]>();   // key: nama biro mentah dari IM4
    members.filter(m => !withCards.has(cleanText(m.nama))).forEach(m => {
      let placed = false;
      for (const d of depts) {
        const b = d.biros.find(bb => isBiroMatch(m.biro, bb.name));
        if (b) {
          const k = `${d.id}|${b.id}`;
          if (!memberBucket.has(k)) memberBucket.set(k, []);
          memberBucket.get(k)!.push(m.nama);
          placed = true;
          break;
        }
      }
      if (!placed) {
        const k = m.biro || 'Tanpa Biro';
        if (!memberUnmapped.has(k)) memberUnmapped.set(k, []);
        memberUnmapped.get(k)!.push(m.nama);
      }
    });

    const result: DeptScore[] = depts.map(d => {
      const biros = d.biros.map(b => {
        const k = `${d.id}|${b.id}`;
        return buildBiroScore(String(b.id), b.name, bucket.get(k) || [], memberBucket.get(k) || []);
      });
      return buildDeptScore(String(d.id), d.name, biros);
    });

    // "Lainnya" hanya muncul bila ada JOB CARD yang biro-nya tidak dikenali (peringatan).
    // Anggota IM4 tanpa job card yang biro-nya tidak cocok tidak lagi ditampung di sini.
    const otherNames = new Set<string>([...unmapped.keys()]);
    if (otherNames.size > 0) {
      const biros: BiroScore[] = [];
      otherNames.forEach(name =>
        biros.push(buildBiroScore(`other-${cleanText(name)}`, name, unmapped.get(name) || [], memberUnmapped.get(name) || [])));
      biros.sort((x, y) => x.name.localeCompare(y.name));
      result.push(buildDeptScore(OTHER_DEPT_ID, 'Lainnya (belum terpetakan)', biros));
    }
    return result;
  }, [cards, period, members]);

  const activeDept = deptScores.find(d => d.id === selectedDept) || null;
  const activeBiro = activeDept?.biros.find(b => b.id === selectedBiro) || null;
  const activePerson = activeBiro?.persons.find(p => p.name === selectedPerson) || null;

  const allPersons = deptScores.flatMap(d => d.biros.flatMap(b => b.persons));
  const divisionProjects = useMemo(
    () => buildProjectScores(deptScores.flatMap(d => d.biros.flatMap(b => b.persons.flatMap(p => p.cards)))),
    [deptScores]
  );
  const deptProjects = useMemo(
    () => activeDept ? buildProjectScores(activeDept.biros.flatMap(b => b.persons.flatMap(p => p.cards))) : [],
    [activeDept]
  );
  const biroProjects = useMemo(
    () => activeBiro ? buildProjectScores(activeBiro.persons.flatMap(p => p.cards)) : [],
    [activeBiro]
  );
  const divisionKpi = avg(allPersons.map(p => p.kpi));

  const goBack = () => {
    if (selectedPerson) setSelectedPerson(null);
    else if (selectedBiro) setSelectedBiro(null);
    else if (selectedDept) setSelectedDept(null);
  };

  const legend = [
    { label: 'Excellent', range: '≥ 90', color: '#34d399' },
    { label: 'Good', range: '80–89', color: '#4cc2ff' },
    { label: 'Fair', range: '70–79', color: '#fbbf24' },
    { label: 'Poor', range: '< 70', color: '#fb7185' },
  ];

  const cardStatus = (c: JobCardRow): { text: string; cls: string } => {
    if (REQUIRE_JOBCARD_CODE && !c.kodeJc) return { text: 'Menunggu Planner', cls: 'text-rose-300' };
    if (c.fromRealisasi) return { text: 'Dari Realisasi JO (belum diinput di Form)', cls: 'text-sky-300' };
    if (!c.release) return { text: 'Belum release', cls: 'text-amber-300' };
    return { text: `Sudah release (Rev ${c.rev})`, cls: 'text-emerald-300' };
  };

  const pageTitle = activePerson ? activePerson.name
    : activeBiro ? activeBiro.name
    : activeDept ? activeDept.name
    : 'KPI Individual — Per Departemen';

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-purple-600 rounded-lg text-white"><Building2 className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white">DASHBOARD KPI INDIVIDUAL</span>
            <span className="text-xs text-slate-400 hidden sm:inline">— {user.nama}</span>
          </div>
          <div className="flex items-center gap-2">
            {onBack && (
              <button onClick={onBack}
                className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
                Ke Portal
              </button>
            )}
            <button onClick={loadData} title="Muat ulang data"
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
            </button>
            <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
              Keluar
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-6">
        {/* Judul + filter */}
        <div className="flex items-start justify-between flex-wrap gap-3">
          <div>
            {/* Breadcrumb */}
            <div className="text-[11px] text-slate-500 flex items-center gap-1 flex-wrap mb-1">
              <button onClick={() => { setSelectedDept(null); setSelectedBiro(null); setSelectedPerson(null); }}
                className="hover:text-slate-300 cursor-pointer">Semua Departemen</button>
              {activeDept && (<>
                <ChevronRight className="w-3 h-3" />
                <button onClick={() => { setSelectedBiro(null); setSelectedPerson(null); }}
                  className="hover:text-slate-300 cursor-pointer">{activeDept.name}</button>
              </>)}
              {activeBiro && (<>
                <ChevronRight className="w-3 h-3" />
                <button onClick={() => setSelectedPerson(null)}
                  className="hover:text-slate-300 cursor-pointer">{activeBiro.name}</button>
              </>)}
              {activePerson && (<>
                <ChevronRight className="w-3 h-3" />
                <span className="text-slate-300">{activePerson.name}</span>
              </>)}
            </div>
            <h1 className="text-2xl font-extrabold text-white">{pageTitle}</h1>
            <p className="text-xs text-slate-400 mt-1">
              Pegawai organik · KPI = ({WEIGHT_A * 100}% × A Efisiensi JO) + ({WEIGHT_B * 100}% × B Release Drawing)
            </p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <select value={period} onChange={(e) => setPeriod(e.target.value)}
              className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white cursor-pointer">
              <option value="all">Semua periode</option>
              {periodOptions.map(p => <option key={p} value={p}>{periodLabel(p)} (Plan Finish)</option>)}
            </select>
            {legend.map(l => (
              <span key={l.label} className="text-[10px] font-bold px-2.5 py-0.5 rounded-full text-slate-950" style={{ background: l.color }}>
                {l.label} {l.range}
              </span>
            ))}
          </div>
        </div>

        {selectedDept && (
          <button onClick={goBack}
            className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
            <ArrowLeft className="w-3.5 h-3.5" /> Kembali
          </button>
        )}

        {errorMsg && <div className="p-3 bg-rose-500/10 border border-rose-500/30 rounded-lg text-xs text-rose-300">{errorMsg}</div>}
        {isLoading && <div className="text-center py-12 text-slate-400 text-sm">Memuat data...</div>}

        {/* ===== LEVEL 1: SEMUA DEPARTEMEN ===== */}
        {!isLoading && !selectedDept && (
          <>
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <div className="lg:col-span-4 bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center text-center gap-4">
              <Donut value={divisionKpi} size={150} color={getCategory(divisionKpi).color} text={fmt(divisionKpi, 0)} />
              <div className="space-y-1">
                <div className="text-xs text-slate-400">Rata-rata KPI seluruh divisi</div>
                <Pill score={divisionKpi} />
                <div className="text-[11px] text-slate-500">
                  {deptScores.length} departemen · {deptScores.reduce((s, d) => s + d.biros.length, 0)} biro · {allPersons.length} personel organik
                </div>
              </div>
            </div>
            <ProjectBarChart key="division" className="lg:col-span-8" title="KPI per proyek — Divisi Desain" rows={divisionProjects} />
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {deptScores.map(d => (
                <div key={d.id} onClick={() => setSelectedDept(d.id)}
                  className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-2xl p-5 cursor-pointer transition flex flex-col items-center gap-3">
                  <div className="w-full flex items-center justify-between gap-2">
                    <h3 className="font-bold text-sm text-white">{d.name}</h3>
                    <ChevronRight className="w-4 h-4 text-slate-500 shrink-0" />
                  </div>
                  <Donut value={d.kpi} size={130} color={getCategory(d.kpi).color} text={fmt(d.kpi, 0)} />
                  <Pill score={d.kpi} />
                  <div className="w-full grid grid-cols-4 text-center text-[11px] text-slate-400 pt-2 border-t border-slate-800">
                    <div><div className="font-bold text-slate-200">{fmt(d.a, 0)}{d.a !== null && '%'}</div>A</div>
                    <div><div className="font-bold text-slate-200">{fmt(d.b, 0)}{d.b !== null && '%'}</div>B</div>
                    <div><div className="font-bold text-slate-200">{d.biros.length}</div>Biro</div>
                    <div><div className="font-bold text-slate-200">{d.scored}/{d.personCount}</div>Terhitung</div>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {/* ===== LEVEL 2: SATU DEPARTEMEN -> DAFTAR BIRO ===== */}
        {!isLoading && activeDept && !selectedBiro && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <div className="lg:col-span-4 bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-4">
              <Donut value={activeDept.kpi} size={170} color={getCategory(activeDept.kpi).color} text={fmt(activeDept.kpi, 0)} />
              <Pill score={activeDept.kpi} />
              <div className="text-[11px] text-slate-400 text-center">
                Rata-rata KPI departemen · {activeDept.scored} dari {activeDept.personCount} personel terhitung
              </div>
              <div className="flex gap-6">
                <div className="text-center">
                  <Donut value={activeDept.a} size={80} color="#34d399" text={activeDept.a === null ? '—' : `${Math.round(activeDept.a)}%`} />
                  <div className="text-[11px] text-slate-400 mt-1">A · Efisiensi JO</div>
                </div>
                <div className="text-center">
                  <Donut value={activeDept.b} size={80} color="#fbbf24" text={activeDept.b === null ? '—' : `${Math.round(activeDept.b)}%`} />
                  <div className="text-[11px] text-slate-400 mt-1">B · Release Drawing</div>
                </div>
              </div>
            </div>

            <ProjectBarChart key={`dept-${activeDept.id}`} className="lg:col-span-8" title={`KPI per proyek — ${activeDept.name}`} rows={deptProjects} />

            <div className="lg:col-span-12 bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <Layers className="w-4 h-4 text-blue-400" /> Skor KPI per biro ({activeDept.biros.length})
              </h2>
              {activeDept.biros.length === 0 && (
                <div className="text-xs text-slate-500 italic py-4 text-center">Departemen ini belum memiliki biro.</div>
              )}
              {activeDept.biros.map(b => {
                const cat = getCategory(b.kpi);
                return (
                  <div key={b.id} onClick={() => setSelectedBiro(b.id)}
                    className="p-3 bg-slate-950 border border-slate-800 hover:border-blue-500 rounded-xl cursor-pointer transition space-y-2">
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-semibold text-sm text-white">{b.name}</span>
                      <div className="flex items-center gap-3">
                        <span className="font-extrabold text-white">{fmt(b.kpi)}</span>
                        <Pill score={b.kpi} />
                        <ChevronRight className="w-4 h-4 text-slate-500" />
                      </div>
                    </div>
                    <div className="h-2.5 bg-slate-800 rounded-full overflow-hidden">
                      <div className="h-2.5 rounded-full" style={{ width: `${b.kpi ?? 0}%`, background: cat.color }} />
                    </div>
                    <div className="text-[11px] text-slate-500">
                      A {fmt(b.a)}{b.a !== null && '%'} · B {fmt(b.b)}{b.b !== null && '%'} · {b.scored}/{b.persons.length} personel terhitung
                      {b.persons.length === 0 && <span className="text-slate-600"> · belum ada anggota</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ===== LEVEL 3: SATU BIRO -> DAFTAR INDIVIDU ===== */}
        {!isLoading && activeBiro && !selectedPerson && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <div className="lg:col-span-5 bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-4">
              <Donut value={activeBiro.kpi} size={190} color={getCategory(activeBiro.kpi).color} text={fmt(activeBiro.kpi, 0)} />
              <Pill score={activeBiro.kpi} />
              <div className="text-[11px] text-slate-400">Rata-rata KPI biro · {activeBiro.scored} dari {activeBiro.persons.length} personel terhitung</div>
              <div className="flex gap-8">
                <div className="text-center">
                  <Donut value={activeBiro.a} size={90} color="#34d399" text={activeBiro.a === null ? '—' : `${Math.round(activeBiro.a)}%`} />
                  <div className="text-[11px] text-slate-400 mt-1">A · Efisiensi JO</div>
                </div>
                <div className="text-center">
                  <Donut value={activeBiro.b} size={90} color="#fbbf24" text={activeBiro.b === null ? '—' : `${Math.round(activeBiro.b)}%`} />
                  <div className="text-[11px] text-slate-400 mt-1">B · Release Drawing</div>
                </div>
              </div>
            </div>

            <ProjectBarChart key={`biro-${activeBiro.id}`} className="lg:col-span-7" title={`KPI per proyek — ${activeBiro.name}`} rows={biroProjects} />

            <div className="lg:col-span-12 bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
              <h2 className="text-sm font-bold text-white flex items-center gap-2"><Users className="w-4 h-4 text-blue-400" /> Skor KPI per personel ({activeBiro.persons.length})</h2>
              {activeBiro.persons.length === 0 && (
                <div className="text-xs text-slate-500 italic py-4 text-center">Belum ada anggota di biro ini.</div>
              )}
              {activeBiro.persons.map(p => {
                const cat = getCategory(p.kpi);
                return (
                  <div key={p.name} onClick={() => setSelectedPerson(p.name)}
                    className="p-3 bg-slate-950 border border-slate-800 hover:border-blue-500 rounded-xl cursor-pointer transition space-y-2">
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-semibold text-sm text-white">{p.name}</span>
                      <div className="flex items-center gap-3">
                        <span className="font-extrabold text-white">{fmt(p.kpi)}</span>
                        <Pill score={p.kpi} />
                      </div>
                    </div>
                    <div className="h-2.5 bg-slate-800 rounded-full overflow-hidden">
                      <div className="h-2.5 rounded-full" style={{ width: `${p.kpi ?? 0}%`, background: cat.color }} />
                    </div>
                    <div className="text-[11px] text-slate-500">
                      A {fmt(p.a)}{p.a !== null && '%'} · B {fmt(p.b)}{p.b !== null && '%'}
                      {p.cards.length === 0 && <span className="text-slate-600"> · belum ada penugasan</span>}
                      {p.pendingCount > 0 && <span className="text-rose-300"> · {p.pendingCount} menunggu Planner</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ===== LEVEL 4: SATU INDIVIDU ===== */}
        {!isLoading && activePerson && (
          <div className="space-y-5">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.a} size={130} color="#34d399" text={activePerson.a === null ? '—' : `${Math.round(activePerson.a)}%`} />
                <div className="text-xs font-bold text-white">A · Efisiensi JO</div>
                <div className="text-[11px] text-slate-400">
                  {activePerson.a === null ? 'Belum ada jobcard dengan Plan JO dan Real JO' : 'Planned JO ÷ Actual JO'}
                </div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.b} size={130} color="#fbbf24" text={activePerson.b === null ? '—' : `${Math.round(activePerson.b)}%`} />
                <div className="text-xs font-bold text-white">B · Release Drawing</div>
                <div className="text-[11px] text-slate-400">
                  {activePerson.targetRev0 === 0
                    ? 'Belum ada drawing dalam daftar'
                    : `${activePerson.onTimeRev0} dari ${activePerson.targetRev0} drawing sudah release`}
                </div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.kpi} size={130} color={getCategory(activePerson.kpi).color} text={fmt(activePerson.kpi, 0)} />
                <div className="text-xs font-bold text-white">Skor KPI Individual</div>
                <Pill score={activePerson.kpi} />
                {activePerson.kpi === null && <div className="text-[11px] text-slate-400">Menunggu data A dan B</div>}
              </div>
            </div>

            {/* Rincian per proyek */}
            {(() => {
              const valid = activePerson.cards.filter(c => !REQUIRE_JOBCARD_CODE || c.kodeJc.trim() !== '');
              const map = new Map<string, { project: string; plan: number; real: number; total: number; released: number; cards: JobCardRow[] }>();
              valid.forEach(c => {
                const k = c.project || '(tanpa proyek)';
                if (!map.has(k)) map.set(k, { project: k, plan: 0, real: 0, total: 0, released: 0, cards: [] });
                const r = map.get(k)!;
                r.cards.push(c);
                if (!c.fromRealisasi) {
                  r.total += 1;
                  if (c.release !== '') r.released += 1;
                }
                if (c.plannedJo > 0 && c.realJo > 0 && (!A_ONLY_FROM_RELEASED || c.release !== '') && (!A_ONLY_REV0 || Number(c.rev) === 0)) {
                  r.plan += c.plannedJo;
                  r.real += c.realJo;
                }
              });
              const rows = Array.from(map.values()).sort((x, y) => x.project.localeCompare(y.project));
              if (rows.length === 0) return null;
              const maxJo = Math.max(1, ...rows.flatMap(r => [r.plan, r.real]));
              return (
                <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <h2 className="text-sm font-bold text-white">Rincian per proyek ({rows.length})</h2>
                    <div className="flex items-center gap-3 text-[10px] text-slate-400">
                      <span className="flex items-center gap-1"><i className="w-2.5 h-2.5 rounded-sm inline-block" style={{ background: '#a78bfa' }} /> Plan JO</span>
                      <span className="flex items-center gap-1"><i className="w-2.5 h-2.5 rounded-sm inline-block" style={{ background: '#34d399' }} /> Real JO</span>
                      <span className="flex items-center gap-1"><i className="w-2.5 h-2.5 rounded-sm inline-block" style={{ background: '#fbbf24' }} /> Release</span>
                    </div>
                  </div>
                  <div className="space-y-3">
                    {rows.map(r => {
                      const ps = scorePerson(activePerson.name, activePerson.biro, r.cards);
                      const eff = r.plan > 0 && r.real > 0 ? Math.min((r.plan / r.real) * 100, CAP_EFFICIENCY) : null;
                      const relPct = r.total ? (r.released / r.total) * 100 : 0;
                      return (
                        <div key={r.project} className="space-y-2">
                        <div onClick={() => toggleProject(r.project)}
                          className="grid grid-cols-12 gap-3 items-center text-xs cursor-pointer hover:bg-slate-950/60 rounded-lg p-1.5 -m-1.5 transition">
                          <div className="col-span-12 sm:col-span-2 font-mono font-bold text-emerald-400 flex items-center gap-1">
                            <ChevronRight className={`w-3.5 h-3.5 text-slate-500 transition-transform ${expandedProjects[r.project] ? 'rotate-90 text-blue-400' : ''}`} />
                            {r.project}
                            <span className="text-[10px] font-normal text-slate-500">({r.total})</span>
                            <span className="text-[10px] font-bold px-1.5 py-0.5 rounded-full text-slate-950 ml-1" style={{ background: getCategory(ps.kpi).color }}>
                              KPI {fmt(ps.kpi, 0)}
                            </span>
                          </div>
                          <div className="col-span-12 sm:col-span-6 space-y-1">
                            <div className="flex items-center gap-2">
                              <div className="flex-1 h-2.5 bg-slate-800 rounded-full overflow-hidden">
                                <div className="h-2.5 rounded-full" style={{ width: `${(r.plan / maxJo) * 100}%`, background: '#a78bfa' }} />
                              </div>
                              <span className="w-14 text-right font-mono text-violet-300">{r.plan ? Math.round(r.plan * 100) / 100 : '-'}</span>
                            </div>
                            <div className="flex items-center gap-2">
                              <div className="flex-1 h-2.5 bg-slate-800 rounded-full overflow-hidden">
                                <div className="h-2.5 rounded-full" style={{ width: `${(r.real / maxJo) * 100}%`, background: '#34d399' }} />
                              </div>
                              <span className="w-14 text-right font-mono text-emerald-400">{r.real ? Math.round(r.real * 100) / 100 : '-'}</span>
                            </div>
                          </div>
                          <div className="col-span-6 sm:col-span-1 text-center">
                            <div className="font-bold text-slate-200">{eff === null ? '—' : `${Math.round(eff)}%`}</div>
                            <div className="text-[10px] text-slate-500">A</div>
                          </div>
                          <div className="col-span-6 sm:col-span-3 space-y-1">
                            <div className="h-2.5 bg-slate-800 rounded-full overflow-hidden">
                              <div className="h-2.5 rounded-full" style={{ width: `${relPct}%`, background: '#fbbf24' }} />
                            </div>
                            <div className="text-[10px] text-slate-400 font-mono">{r.released}/{r.total} release</div>
                          </div>
                        </div>

                        {expandedProjects[r.project] && (
                          <div className="ml-4 pl-3 border-l border-slate-700 space-y-2">
                            {(() => {
                              const maxCard = Math.max(1, ...r.cards.flatMap(c => [c.plannedJo, c.realJo]));
                              return r.cards.map(c => {
                                const st = cardStatus(c);
                                return (
                                  <div key={c.id} className="p-2.5 bg-slate-950 border border-slate-800 rounded-lg space-y-1.5 text-xs">
                                    <div className="flex items-start justify-between gap-3">
                                      <div className="text-slate-200">{c.taskName || '-'}</div>
                                      <div className="flex items-center gap-2 shrink-0">
                                        <span className="font-mono text-[10px] text-slate-400">Rev {c.rev}</span>
                                        <span className={`text-[10px] font-semibold ${st.cls}`}>{st.text}</span>
                                      </div>
                                    </div>
                                    <div className="flex items-center gap-2">
                                      <div className="flex-1 h-2 bg-slate-800 rounded-full overflow-hidden">
                                        <div className="h-2 rounded-full" style={{ width: `${(c.plannedJo / maxCard) * 100}%`, background: '#a78bfa' }} />
                                      </div>
                                      <span className="w-12 text-right font-mono text-violet-300">{c.plannedJo || '-'}</span>
                                    </div>
                                    <div className="flex items-center gap-2">
                                      <div className="flex-1 h-2 bg-slate-800 rounded-full overflow-hidden">
                                        <div className="h-2 rounded-full" style={{ width: `${(c.realJo / maxCard) * 100}%`, background: '#34d399' }} />
                                      </div>
                                      <span className="w-12 text-right font-mono text-emerald-400">{c.realJo ? Math.round(c.realJo * 100) / 100 : '-'}</span>
                                    </div>
                                    <div className="text-[10px] text-slate-500 font-mono">
                                      {c.kodeJc} · Plan Finish {formatDisplayDate(c.endDate)} · Release {formatDisplayDate(c.release)}
                                    </div>
                                  </div>
                                );
                              });
                            })()}
                          </div>
                        )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
              <h2 className="text-sm font-bold text-white">Data pekerjaan ({activePerson.cards.length})</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="text-slate-400 border-b border-slate-800">
                      <th className="py-2 px-2.5 text-amber-400 font-mono">Jobcard</th>
                      <th className="py-2 px-2.5">Proyek</th>
                      <th className="py-2 px-2.5">Deskripsi</th>
                      <th className="py-2 px-2.5 text-center">Rev</th>
                      <th className="py-2 px-2.5 text-right">Plan JO</th>
                      <th className="py-2 px-2.5 text-right">Real JO</th>
                      <th className="py-2 px-2.5">Plan Start</th>
                      <th className="py-2 px-2.5">Plan Finish</th>
                      <th className="py-2 px-2.5">Release</th>
                      <th className="py-2 px-2.5">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
                    {activePerson.cards.length === 0 && (
                      <tr><td colSpan={10} className="py-6 text-center text-slate-500 italic">Belum ada penugasan pada periode terpilih.</td></tr>
                    )}
                    {activePerson.cards.map(c => {
                      const st = cardStatus(c);
                      return (
                        <tr key={c.id} className="hover:bg-slate-950/50">
                          <td className="py-2 px-2.5 font-mono font-bold text-amber-300">{c.kodeJc || <span className="text-rose-400 font-normal">-</span>}</td>
                          <td className="py-2 px-2.5 text-emerald-400">{c.project}</td>
                          <td className="py-2 px-2.5 text-slate-200">{c.taskName}</td>
                          <td className="py-2 px-2.5 text-center font-mono">{c.rev}</td>
                          <td className="py-2 px-2.5 text-right font-mono text-violet-300">{c.plannedJo || '-'}</td>
                          <td className="py-2 px-2.5 text-right font-mono text-emerald-400">{c.realJo ? Math.round(c.realJo * 100) / 100 : '-'}</td>
                          <td className="py-2 px-2.5 font-mono">{formatDisplayDate(c.startDate)}</td>
                          <td className="py-2 px-2.5 font-mono">{formatDisplayDate(c.endDate)}</td>
                          <td className="py-2 px-2.5 font-mono text-cyan-300">{formatDisplayDate(c.release)}</td>
                          <td className={`py-2 px-2.5 font-semibold ${st.cls}`}>{st.text}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}