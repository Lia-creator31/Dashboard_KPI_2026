import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { departmentsData } from '../data';
import { projectCandidates, dedupeRows, resolveDrawingProject, findDrawingProjectInDescription, GOOGLE_DRIVE_SHEETS, findMatchingDriveSheet, ALL_DRIVE_PROJECT_KEYS, isCorpProject, fetchDrawingControlRowsForProject, DrawingControlRow, parseDrawingControlDate } from '../lib/projectCandidates';
import { Building2, ArrowLeft, ChevronRight, RefreshCw, Users, Layers, ArrowUpDown, FileSpreadsheet, FileText, Clock, X } from 'lucide-react';
import { getAllTimesheetData, getPersonMonthlyTimesheet, TIMESHEET_STORAGE_KEY, loadTimesheetFromDatabase } from '../lib/timesheetAccomplishment';

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
  typeDrawing?: string;
  fromRealisasi?: boolean;   // true = diambil dari file Realisasi JO (tidak ada di job_cards)
  noJobcardDesc?: boolean;   // true = baris tanpa kode jobcard, deskripsi dari WO/Description
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
const parseDriveDate = parseDrawingControlDate;

function isKnownDriveProject(projectCode: string): boolean {
  if (findMatchingDriveSheet(projectCode)) return true;
  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  return cands.some(c => !!findMatchingDriveSheet(c));
}

async function fetchDrawingRowsFor(projectCode: string): Promise<DrawingControlRow[]> {
  return fetchDrawingControlRowsForProject(projectCode);
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
  // Hilangkan kode proyek di awal deskripsi (mis. "W000314 - E25291.01100_ R0_..." atau "KCR60M - E25291.01100...") agar yang terambil adalah nomor gambarnya
  s = s.replace(/^[WM]\d{4,6}\s*[-–—_:/]\s*/i, '');
  s = s.replace(/^(?:KSOT|LCU|W?FFBNW\d*)\s*[-–—_:/]\s*/i, '');
  s = s.replace(/^[A-Z]{2,12}\d{0,6}\s*[-–—_:/]\s*(?=[A-Z]{1,4}\d{2,}[A-Z0-9]*\.)/i, '');
  s = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-\s_:]\s*/i, '');
  // Prioritaskan nomor gambar berformat titik (mis. E25291.01100, EC123C.6) di mana pun posisinya
  const mDot = s.match(/\b[A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+/i);
  if (mDot) return mDot[0].toUpperCase();
  const m = s.match(/^[A-Z]{1,4}\d{3,}[A-Z0-9]*(?:\.[A-Z0-9]+)*/i);
  if (m) return m[0].toUpperCase();
  const first = s.split(/\s+-\s+|\s+|_/)[0] || '';
  return /\d/.test(first) && first.length >= 4 ? first.toUpperCase() : '';
}

const drawingKey = (v: string) => cleanText(extractDrawingNumber(v) || v);

function findReleaseInRows(rows: DrawingControlRow[], taskName: string, rev: string | number): string {
  if (!rows || rows.length === 0 || !taskName) return '';
  const wanted = drawingKey(taskName);
  // Hilangkan kode proyek di awal deskripsi (mis. "M000313 - Bonjean Table" -> "Bonjean Table") untuk pencocokan nama gambar
  const strippedTask = (taskName || '')
    .replace(/^[WM]\d{4,6}\s*[-–—_:/]\s*/i, '')
    .replace(/^[A-Z0-9]{2,12}(?:-[A-Z0-9]{2,8})?\s*[-–—_:/]\s*/i, '')
    .trim();
  const cleanTarget = cleanText(taskName);
  const cleanStripped = cleanText(strippedTask);
  if (!wanted && !cleanTarget) return '';
  const toNum = (v: any) => { const n = parseInt(String(v ?? '').replace(/[^0-9]/g, ''), 10); return isNaN(n) ? 0 : n; };
  const wantRev = toNum(rev);
  const hit = rows.find(r => {
    if (toNum(r.rev) !== wantRev) return false;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(parseDriveDate(r.finishDate))) return false;
    const rKey = drawingKey(r.noDwg);
    if (wanted && rKey && rKey === wanted) return true;
    const cDwg = cleanText(r.noDwg);
    if (cDwg && cDwg.length >= 5 && cleanTarget.includes(cDwg)) return true;
    // Cadangan: jika di deskripsi tertulis Nama Gambar (mis. "M000313 - Bonjean Table"), cocokkan dengan kolom DRAWING NAME
    const cName = cleanText(r.drawingName);
    if (cName && cName.length >= 5 && (cName === cleanStripped || cleanTarget.includes(cName) || (cleanStripped.length >= 5 && cName.includes(cleanStripped)))) return true;
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
      let j = row.findIndex(c => c === 'work order/job card/display name' || c === 'job card/display name' || c === 'jobcard/display name');
      if (j === -1) j = row.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state'));
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
  woDesc: string;      // Work Order/Description (selalu diisi, bahkan saat tanpa jobcard)
  noJobcard: boolean;  // true = baris ini tidak punya kode jobcard
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
    let headerIdx = -1, cName = -1, cProj = -1, cJc = -1, cEff = -1, cOt = -1, cTask = -1, cDwgCtrl = -1;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const h = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      const jcExact = h.findIndex(c => c === 'work order/job card/display name' || c === 'job card/display name' || c === 'jobcard/display name');
      const jc = jcExact !== -1 ? jcExact : h.findIndex(c => c.includes('jobcard') && !c.includes('task') && !c.includes('state') && !c.includes('drawing'));
      const ef = h.findIndex(c => c.includes('effective'));
      const ot = h.findIndex(c => c.includes('overtime'));
      if (jc !== -1 && (ef !== -1 || ot !== -1)) {
        headerIdx = r; cJc = jc; cEff = ef; cOt = ot;
        cName = h.findIndex(c => c === 'employee/name' || c === 'nama' || c === 'name' || c === 'user' || c === 'employee' || c === 'pegawai');
        cProj = h.findIndex(c => c === 'project/code' || c === 'project' || c === 'proyek');
        cTask = h.findIndex(c => c === 'work order/description' || c === 'jobcard/task name' || c === 'description' || c === 'deskripsi' || c === 'task name');
        // Deskripsi utama: Jobcard/Drawing Control (fallback ke Work Order/Description per baris)
        cDwgCtrl = h.findIndex(c => c === 'jobcard/drawing control' || c === 'work order/job card/drawing control' || c === 'job card/drawing control' || c === 'drawing control');
        break;
      }
    }
    if (headerIdx === -1 || cName === -1) return;   // tanpa kolom nama -> tidak bisa per pegawai

    // Cari kolom Work Order/Description untuk deskripsi fallback (terutama saat tanpa jobcard)
    let cWoDesc = -1;
    if (headerIdx !== -1) {
      const h = (rows[headerIdx] || []).map(v => String(v).trim().toLowerCase());
      cWoDesc = h.findIndex(c => c === 'work order/description' || c === 'wo/description' || c === 'description');
      // Jangan bentrok dengan cTask yang mungkin sudah diset ke kolom yang sama
    }

    // Kelompokkan jam per (nama, jobcard) untuk baris DENGAN jobcard
    // Untuk baris TANPA jobcard, kumpulkan deskripsi unik per (nama, project)
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const nama = splitUserCell(row[cName]);
      if (!nama) continue;
      const jcRaw = String(row[cJc] ?? '').trim();
      const woDesc = (cWoDesc !== -1 ? String(row[cWoDesc] ?? '').trim() : '') ||
                     (cTask !== -1 ? String(row[cTask] ?? '').trim() : '');
      const isEmptyJc = !jcRaw || jcRaw.toLowerCase() === 'nan' || jcRaw.toLowerCase().includes('jobcard');
      if (isEmptyJc) {
        // Baris tanpa jobcard: rekam sebagai entri noJobcard untuk ditampilkan sebagai deskripsi
        const proj = cProj !== -1 ? extractProjectCodeFromCell(row[cProj]) : '';
        if (!woDesc) continue; // skip baris benar-benar kosong
        out.push({
          nama,
          project: proj,
          jobcard: '',
          taskName: woDesc,
          woDesc,
          noJobcard: true,
          hours: (cEff !== -1 ? Math.abs(parseNumber(row[cEff])) : 0) + (cOt !== -1 ? Math.abs(parseNumber(row[cOt])) : 0),
        });
        continue;
      }
      out.push({
        nama,
        project: cProj !== -1 ? extractProjectCodeFromCell(row[cProj]) : '',
        jobcard: jcRaw.replace(/\s+/g, ' ').toUpperCase(),
        taskName: (cDwgCtrl !== -1 ? String(row[cDwgCtrl] ?? '').trim() : '') || woDesc,
        woDesc,
        noJobcard: false,
        hours: (cEff !== -1 ? Math.abs(parseNumber(row[cEff])) : 0) + (cOt !== -1 ? Math.abs(parseNumber(row[cOt])) : 0),
      });
    }
  });
  return out;
}

/* Ekstrak kode proyek dari deskripsi CORP2026 (misal "W000314 - E25291.01100..." -> "W000314") */
function extractProjectFromDesc(desc: string): string {
  if (!desc) return '';
  const known = findDrawingProjectInDescription(desc);
  if (known) return known;
  // Cari kode proyek standar: 1 huruf + angka (M000312, W000307, E260N002, dsb.)
  const m = desc.match(/\b([A-Z]\d{4,}[A-Z0-9]*)\b/);
  if (m) return m[1];
  // Fallback: token pertama yang mengandung angka
  const first = desc.split(/[\s\-–—]+/)[0] || '';
  return /[A-Z].*\d/.test(first) && first.length >= 4 ? first.toUpperCase() : '';
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
  // Hanya mengecualikan baris deskripsi kegiatan yang benar-benar tanpa kode Jobcard (noJobcardDesc)
  const targets = valid.filter(c => !(c as any).noJobcardDesc);
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
      const valid = pr.cards.filter(c => (!REQUIRE_JOBCARD_CODE || c.kodeJc.trim() !== '') && !(c as any).noJobcardDesc);
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
  const [showAllPersonnel, setShowAllPersonnel] = useState(false);
  const [personnelSortOrder, setPersonnelSortOrder] = useState<'desc' | 'asc'>('desc');
  const [personnelSearch, setPersonnelSearch] = useState('');
  const [timesheetAllData, setTimesheetAllData] = useState(() => getAllTimesheetData());
  const [timesheetDetailModal, setTimesheetDetailModal] = useState<{
    month: string;
    personName: string;
    kind: 'overtime' | 'regular';
    accomplishment: number;
    otHours?: number;
    tsHours?: number;
    expectedTs?: number;
    totalTs?: number;
  } | null>(null);

  useEffect(() => {
    loadTimesheetFromDatabase().then(data => {
      if (data && Object.keys(data).length > 0) {
        setTimesheetAllData(data);
      }
    });
    const handleStorage = (e: StorageEvent) => {
      if (e.key === TIMESHEET_STORAGE_KEY) {
        setTimesheetAllData(getAllTimesheetData());
      }
    };
    window.addEventListener('storage', handleStorage);
    return () => window.removeEventListener('storage', handleStorage);
  }, []);
  const toggleProject = (k: string) => setExpandedProjects(prev => ({ ...prev, [k]: !prev[k] }));

  const loadData = useCallback(async () => {
    setIsLoading(true);
    setErrorMsg('');
    try {
      // Ambil SELURUH baris job_cards dengan paginasi (karena Supabase membatasi 1000 baris per query jika tanpa range)
      const fetchAllJcRows = async () => {
        const all: any[] = [];
        for (let from = 0; ; from += 1000) {
          const { data, error } = await supabase.from('job_cards').select('*').range(from, from + 999);
          if (error) throw error;
          all.push(...(data || []));
          if (!data || data.length < 1000) break;
        }
        return all;
      };

      const [jcAllData, tsRes, lineRes, wbRes, im4Wb, hourPlanner, dbTimesheet] = await Promise.all([
        fetchAllJcRows(),
        supabase.from('outsourcing_timesheets').select('id, status'),
        supabase.from('outsourcing_timesheet_lines').select('timesheet_id, work_order_code, effective_hours, overtime_hours'),
        supabase.storage.from(MASTER_BUCKET).download('realisasi.xlsx'),
        loadIm4Workbook(),
        loadHourPlanner(),
        loadTimesheetFromDatabase(),
      ]);
      if (dbTimesheet && Object.keys(dbTimesheet).length > 0) {
        setTimesheetAllData(dbTimesheet);
      }

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
      const typeLocal = getLocalMap('task_type_map');

      const rows: JobCardRow[] = (jcAllData || []).map((row: any) => {
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
          typeDrawing: row.type || typeLocal[row.id] || 'Drawing',
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

        // === Baris DENGAN jobcard: sudah ada kode Jobcard dari file yang diimpor ===
        realRows.filter(r => !r.noJobcard).forEach(r => {
          const m = resolveMember(r.nama);
          if (!m) return;   // bukan pegawai organik IM4
          const jcKey = cleanText(r.jobcard);
          const k = `${cleanText(m.nama)}|${jcKey}`;
          if (existing.has(k)) return;
          if (extra.has(k)) {
            // Jika baris sebelumnya di extra deskripsinya belum punya nomor gambar dan baris ini punya, perbarui deskripsinya
            const cur = extra.get(k)!;
            if (r.taskName && (!cur.taskName || (!extractDrawingNumber(cur.taskName) && extractDrawingNumber(r.taskName)))) {
              cur.taskName = r.taskName;
            }
            return;
          }
          const hp = hourPlanner[jcKey];
          const revMatch = (r.taskName || '').match(/(?:^|[\s_-])R(?:EV)?\.?\s*(\d{1,2})(?:$|[\s_-])/i);
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
            rev: revMatch ? String(parseInt(revMatch[1], 10)) : '0',
            release: '',
            realJo: realHoursFor(m.nama, jcKey) + (tsMap.get(jcKey) || 0),
            typeDrawing: 'Drawing',
            fromRealisasi: false,
          });
        });

        // === Baris TANPA jobcard: tampilkan sebagai penugasan dengan deskripsi ===
        // Khusus CORP2026: baca deskripsi, cari kode proyek dari Drawing Control Drive
        // Kelompokkan per (nama, project, woDesc) agar tidak duplikat per hari
        const noJcSeen = new Map<string, JobCardRow>();
        realRows.filter(r => r.noJobcard && r.woDesc).forEach(r => {
          const m = resolveMember(r.nama);
          if (!m) return;

          // Tentukan proyek: dari kolom Project atau ekstrak dari deskripsi
          let proj = r.project;
          const isCorp = proj.toUpperCase().startsWith('CORP');

          // Untuk CORP2026: coba ekstrak kode proyek dari deskripsi (misal "W000307 - INPUT PDM...")
          let derivedProjFromDesc = '';
          if (isCorp) {
            derivedProjFromDesc = extractProjectFromDesc(r.woDesc);
          }

          const effectiveProj = derivedProjFromDesc || proj;
          // Kunci unik: nama + project efektif + deskripsi (dedupe baris harian)
          const descKey = cleanText(r.woDesc.slice(0, 60));
          const k = `nojc|${cleanText(m.nama)}|${cleanText(effectiveProj)}|${descKey}`;
          if (noJcSeen.has(k)) {
            // Akumulasi jam
            const ex = noJcSeen.get(k)!;
            ex.realJo += r.hours;
            return;
          }
          noJcSeen.set(k, {
            id: `nojc-${k}`,
            biro: m.biro || 'Tanpa Biro',
            pic: m.nama,
            project: effectiveProj,
            taskName: r.woDesc,       // deskripsi sebagai nama task
            startDate: '',
            endDate: '',
            plannedJo: 0,
            kodeJc: '',               // tidak punya kode jobcard
            rev: '0',
            release: '',
            realJo: r.hours,
            typeDrawing: 'Drawing',
            fromRealisasi: true,
            noJobcardDesc: true,      // flag: ini baris tanpa jobcard dengan deskripsi
          } as any);
        });
        rows.push(...extra.values(), ...noJcSeen.values());
      }

      setCards(rows);

      // Sinkronisasi otomatis tanggal Release dari Google Drive untuk kartu yang belum punya release
      // 1) Jika proyek adalah CORP2026 dan di deskripsi ada nama proyek (mis. W000314), cari di proyek tsb.
      // 2) Khusus CORP2026 yang tidak menulis kode proyek di deskripsi tapi punya nomor gambar (mis. "E90961.92009 - LIGHTWEIGHT ESTIMATION"),
      //    cari satu per satu di seluruh proyek Drawing Control yang ada di link Google Drive.
      const missingRelease = rows.filter(r => (r.project || r.taskName) && r.taskName && !r.release);
      if (missingRelease.length > 0) {
        const targetProjOf = (r: JobCardRow) => resolveDrawingProject(r.project, r.taskName);
        const needsAllDriveSearch = missingRelease.some(r => isCorpProject(targetProjOf(r)) && !!extractDrawingNumber(r.taskName));
        const baseProjs = missingRelease.map(targetProjOf).filter(p => p && !isCorpProject(p));
        const projects = Array.from(new Set(needsAllDriveSearch ? [...baseProjs, ...ALL_DRIVE_PROJECT_KEYS] : baseProjs)).sort((a, b) =>
          (isKnownDriveProject(b) ? 1 : 0) - (isKnownDriveProject(a) ? 1 : 0)
        );
        (async () => {
          const rowsByProj = new Map<string, DrawingControlRow[]>();
          for (let i = 0; i < projects.length; i += 4) {
            const batchProjs = projects.slice(i, i + 4);
            await Promise.all(batchProjs.map(async p => {
              try { rowsByProj.set(p, await fetchDrawingRowsFor(p)); } catch { rowsByProj.set(p, []); }
            }));
            const batchFound = new Map<string, string>();
            missingRelease.forEach(r => {
              if (r.release || batchFound.has(r.id)) return;
              const effProj = targetProjOf(r);
              if (!isCorpProject(effProj)) {
                if (!batchProjs.includes(effProj)) return;
                const rel = findReleaseInRows(rowsByProj.get(effProj) || [], r.taskName, r.rev || '0');
                if (rel) batchFound.set(r.id, rel);
              } else if (extractDrawingNumber(r.taskName)) {
                // Proyek masih CORP2026 tapi ada nomor gambar -> cek satu per satu di setiap proyek Drive yang sudah dimuat
                for (const p of batchProjs) {
                  const rel = findReleaseInRows(rowsByProj.get(p) || [], r.taskName, r.rev || '0');
                  if (rel) { batchFound.set(r.id, rel); break; }
                }
              }
            });
            if (batchFound.size > 0) {
              batchFound.forEach((rel, id) => {
                const card = missingRelease.find(x => x.id === id);
                if (card) card.release = rel;
              });
              setCards(prev => prev.map(c => batchFound.has(c.id) ? { ...c, release: batchFound.get(c.id)! } : c));
              const ids = Array.from(batchFound.keys());
              for (let j = 0; j < ids.length; j += 10) {
                await Promise.all(ids.slice(j, j + 10).map(async id => {
                  const rel = batchFound.get(id)!;
                  saveLocalRelease(id, rel);
                  if (!id.startsWith('realisasi-') && !id.startsWith('nojc-')) {
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

  const allPersonsWithMeta = useMemo(() => {
    return deptScores.flatMap(d =>
      d.biros.flatMap(b =>
        b.persons.map(p => ({ person: p, deptId: d.id, deptName: d.name, biroId: b.id, biroName: b.name }))
      )
    );
  }, [deptScores]);

  const sortedAllPersons = useMemo(() => {
    const q = personnelSearch.trim().toLowerCase();
    const filtered = q
      ? allPersonsWithMeta.filter(x =>
          x.person.name.toLowerCase().includes(q) ||
          x.biroName.toLowerCase().includes(q) ||
          x.deptName.toLowerCase().includes(q)
        )
      : [...allPersonsWithMeta];
    filtered.sort((x, y) => {
      const kx = x.person.kpi;
      const ky = y.person.kpi;
      // Personel yang belum punya nilai KPI (null) selalu ditaruh di bagian paling bawah
      if (kx === null && ky === null) return x.person.name.localeCompare(y.person.name);
      if (kx === null) return 1;
      if (ky === null) return -1;
      return personnelSortOrder === 'desc'
        ? (ky - kx) || x.person.name.localeCompare(y.person.name)
        : (kx - ky) || x.person.name.localeCompare(y.person.name);
    });
    return filtered;
  }, [allPersonsWithMeta, personnelSortOrder, personnelSearch]);
  const allPersons = useMemo(() => allPersonsWithMeta.map(x => x.person), [allPersonsWithMeta]);

  const buildPersonnelExportData = () => {
    const periodText = period === 'all' ? 'Semua Periode' : periodLabel(period);
    const orderText = personnelSortOrder === 'desc' ? 'Tertinggi ke Terendah' : 'Terendah ke Tertinggi';
    const head = [
      'No',
      'Nama Personel',
      'Biro',
      'Departemen',
      'A - Efisiensi JO (%)',
      'B - Release Drawing (%)',
      'Drawing Release',
      'Skor KPI',
      'Kategori',
      'Jumlah Tugas',
    ];
    const rows = sortedAllPersons.map((item, idx) => {
      const p = item.person;
      const cat = getCategory(p.kpi);
      const descOnlyCount = p.cards.filter(c => !!(c as any).noJobcardDesc).length;
      const realCardCount = p.cards.length - descOnlyCount;
      return [
        idx + 1,
        p.name,
        item.biroName,
        item.deptName,
        p.a !== null ? Number(p.a.toFixed(1)) : '-',
        p.b !== null ? Number(p.b.toFixed(1)) : '-',
        p.targetRev0 > 0 ? `${p.onTimeRev0}/${p.targetRev0}` : '-',
        p.kpi !== null ? Number(p.kpi.toFixed(1)) : '-',
        cat.label,
        realCardCount,
      ];
    });
    return { head, rows, periodText, orderText };
  };

  const exportPersonnelExcel = () => {
    const { head, rows, periodText, orderText } = buildPersonnelExportData();
    const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 32 },
      { wch: 36 },
      { wch: 40 },
      { wch: 20 },
      { wch: 22 },
      { wch: 16 },
      { wch: 12 },
      { wch: 16 },
      { wch: 14 },
    ];
    const wbOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbOut, ws, 'Peringkat KPI');
    const safePeriod = periodText.replace(/\s+/g, '_');
    const safeOrder = personnelSortOrder === 'desc' ? 'Tertinggi_Terendah' : 'Terendah_Tertinggi';
    XLSX.writeFile(wbOut, `Peringkat_KPI_Personel_Divisi_Desain_${safeOrder}_${safePeriod}.xlsx`);
    void orderText;
  };

  const exportPersonnelPdf = () => {
    const { head, rows, periodText, orderText } = buildPersonnelExportData();
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });

    doc.setFontSize(13);
    doc.text('PERINGKAT KPI SELURUH PERSONEL DIVISI DESAIN', 40, 34);
    doc.setFontSize(8.5);
    doc.text(
      `Urutan: ${orderText}  |  Periode: ${periodText}  |  Total Personel: ${rows.length} orang  |  Dicetak: ${new Date().toLocaleDateString('id-ID')}`,
      40,
      49
    );

    autoTable(doc, {
      head: [head],
      body: rows,
      startY: 58,
      theme: 'grid',
      styles: { fontSize: 7.5, cellPadding: 3.5, valign: 'middle' },
      headStyles: { fillColor: [30, 41, 59], textColor: 255, halign: 'center' },
      columnStyles: {
        0: { halign: 'center', cellWidth: 28 },
        4: { halign: 'right', cellWidth: 65 },
        5: { halign: 'right', cellWidth: 70 },
        6: { halign: 'center', cellWidth: 55 },
        7: { halign: 'right', cellWidth: 50, fontStyle: 'bold' },
        8: { halign: 'center', cellWidth: 65 },
        9: { halign: 'center', cellWidth: 48 },
      },
      didDrawPage: () => {
        const p = doc.getNumberOfPages();
        doc.setFontSize(8);
        doc.text(`Hal. ${p}`, doc.internal.pageSize.getWidth() - 60, doc.internal.pageSize.getHeight() - 18);
      },
    });

    const safePeriod = periodText.replace(/\s+/g, '_');
    const safeOrder = personnelSortOrder === 'desc' ? 'Tertinggi_Terendah' : 'Terendah_Tertinggi';
    doc.save(`Peringkat_KPI_Personel_Divisi_Desain_${safeOrder}_${safePeriod}.pdf`);
  };

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
    if ((c as any).noJobcardDesc) return { text: 'Deskripsi kegiatan (tanpa Jobcard)', cls: 'text-slate-400' };
    if (REQUIRE_JOBCARD_CODE && !c.kodeJc) return { text: 'Menunggu Planner', cls: 'text-rose-300' };
    if (!c.release) return { text: 'Belum release', cls: 'text-amber-300' };
    return { text: `Sudah release (Rev ${c.rev})`, cls: 'text-emerald-300' };
  };

  const pageTitle = activePerson ? activePerson.name
    : activeBiro ? activeBiro.name
    : activeDept ? activeDept.name
    : 'KPI Individual — Per Departemen';

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/95 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-blue-600/15 border border-blue-500/30 rounded-md text-blue-400"><Building2 className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white tracking-tight">DIVISI DESAIN</span>
            <span className="text-slate-700">|</span>
            <span className="text-xs font-medium text-slate-300">Dashboard KPI Individual</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs text-slate-400 hidden sm:inline mr-1">{user.nama}</span>
            {onBack && (
              <button onClick={onBack}
                className="px-3 py-1.5 text-xs bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-md border border-slate-700 cursor-pointer transition flex items-center gap-1">
                <ArrowLeft className="w-3.5 h-3.5" /> Ke Portal
              </button>
            )}
            <button onClick={loadData} title="Muat ulang data"
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-md border border-slate-700 cursor-pointer transition">
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
            </button>
            <button onClick={onLogout} className="px-3 py-1.5 text-xs bg-slate-800 hover:bg-rose-600/20 text-slate-300 hover:text-rose-300 rounded-md border border-slate-700 hover:border-rose-500/30 cursor-pointer transition">
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

            {/* Tombol & Panel Peringkat KPI Seluruh Personel Divisi Desain */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 space-y-4">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-blue-500/10 text-blue-400 rounded-xl border border-blue-500/20">
                    <Users className="w-4 h-4" />
                  </div>
                  <div>
                    <h2 className="text-sm font-bold text-white">Peringkat KPI Personel ({allPersons.length} orang)</h2>
                  </div>
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  {showAllPersonnel && (
                    <>
                      <input
                        type="text"
                        value={personnelSearch}
                        onChange={(e) => setPersonnelSearch(e.target.value)}
                        placeholder="Cari nama, biro, departemen..."
                        className="px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-blue-500 w-56"
                      />
                      <button
                        onClick={() => setPersonnelSortOrder(prev => (prev === 'desc' ? 'asc' : 'desc'))}
                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs rounded-lg flex items-center gap-1.5 cursor-pointer transition"
                        title="Balik urutan peringkat KPI"
                      >
                        <ArrowUpDown className="w-3.5 h-3.5" />
                        {personnelSortOrder === 'desc' ? 'Tertinggi → Terendah' : 'Terendah → Tertinggi'}
                      </button>
                      <button
                        onClick={exportPersonnelExcel}
                        className="px-3 py-1.5 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 font-semibold text-xs rounded-lg flex items-center gap-1.5 cursor-pointer transition"
                        title="Download Peringkat KPI ke file Excel (.xlsx)"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        Export Excel
                      </button>
                      <button
                        onClick={exportPersonnelPdf}
                        className="px-3 py-1.5 bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 border border-rose-500/30 font-semibold text-xs rounded-lg flex items-center gap-1.5 cursor-pointer transition"
                        title="Download Peringkat KPI ke file PDF (.pdf)"
                      >
                        <FileText className="w-3.5 h-3.5" />
                        Export PDF
                      </button>
                    </>
                  )}
                  <button
                    onClick={() => setShowAllPersonnel(prev => !prev)}
                    className={`px-3.5 py-1.5 rounded-lg text-xs font-bold border cursor-pointer transition flex items-center gap-1.5 ${
                      showAllPersonnel
                        ? 'bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700'
                        : 'bg-emerald-600 hover:bg-emerald-500 text-white border-emerald-500'
                    }`}
                  >
                    <Users className="w-3.5 h-3.5" />
                    {showAllPersonnel ? 'Sembunyikan Daftar Personel' : `Lihat Semua Personel (${allPersons.length})`}
                  </button>
                </div>
              </div>

              {showAllPersonnel && (
                <div className="space-y-2 pt-2 border-t border-slate-800 max-h-[540px] overflow-y-auto pr-1">
                  {sortedAllPersons.length === 0 ? (
                    <div className="text-xs text-slate-500 italic py-6 text-center">Tidak ada personel yang cocok dengan pencarian.</div>
                  ) : (
                    sortedAllPersons.map((item, idx) => {
                      const p = item.person;
                      const cat = getCategory(p.kpi);
                      const descOnlyCount = p.cards.filter(c => !!(c as any).noJobcardDesc).length;
                      const realCardCount = p.cards.length - descOnlyCount;
                      return (
                        <div
                          key={`${item.deptId}-${item.biroId}-${p.name}`}
                          onClick={() => {
                            setSelectedDept(item.deptId);
                            setSelectedBiro(item.biroId);
                            setSelectedPerson(p.name);
                          }}
                          className="p-3 bg-slate-950 border border-slate-800 hover:border-blue-500 rounded-xl cursor-pointer transition space-y-2"
                        >
                          <div className="flex items-center justify-between gap-3 flex-wrap">
                            <div className="flex items-center gap-2.5 min-w-0">
                              <span className="w-7 h-7 rounded-lg bg-slate-900 border border-slate-800 text-slate-300 font-mono text-xs font-bold flex items-center justify-center shrink-0">
                                {idx + 1}
                              </span>
                              <div className="min-w-0">
                                <div className="font-semibold text-sm text-white truncate">{p.name}</div>
                                <div className="text-[11px] text-slate-400 truncate">
                                  {item.biroName} · <span className="text-slate-500">{item.deptName}</span>
                                </div>
                              </div>
                            </div>
                            <div className="flex items-center gap-3 shrink-0">
                              <span className="font-extrabold text-white text-sm">{fmt(p.kpi)}</span>
                              <Pill score={p.kpi} />
                              <ChevronRight className="w-4 h-4 text-slate-500" />
                            </div>
                          </div>
                          <div className="h-2 bg-slate-800 rounded-full overflow-hidden">
                            <div className="h-2 rounded-full" style={{ width: `${Math.max(0, Math.min(100, p.kpi ?? 0))}%`, background: cat.color }} />
                          </div>
                          <div className="text-[11px] text-slate-500">
                            A {fmt(p.a)}{p.a !== null && '%'} · B {fmt(p.b)}{p.b !== null && '%'}
                            {realCardCount > 0 && <span className="text-slate-400"> · {realCardCount} tugas</span>}
                            {p.cards.length === 0 && <span className="text-slate-600"> · belum ada penugasan</span>}
                            {realCardCount === 0 && descOnlyCount > 0 && <span className="text-slate-400"> · {descOnlyCount} deskripsi kegiatan</span>}
                            {realCardCount > 0 && descOnlyCount > 0 && <span className="text-slate-400"> · +{descOnlyCount} deskripsi kegiatan</span>}
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              )}
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
                // Hitung berapa kartu yang hanya berisi deskripsi (tanpa jobcard)
                const descOnlyCount = p.cards.filter(c => !!(c as any).noJobcardDesc).length;
                const realCardCount = p.cards.length - descOnlyCount;
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
                      {realCardCount === 0 && descOnlyCount > 0 && <span className="text-slate-400"> · {descOnlyCount} deskripsi kegiatan (belum ada Jobcard)</span>}
                      {realCardCount > 0 && descOnlyCount > 0 && <span className="text-slate-400"> · +{descOnlyCount} deskripsi kegiatan</span>}
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
                if (!(c as any).noJobcardDesc) {
                  r.total += 1;
                  if (c.release !== '') r.released += 1;
                }
                if (c.plannedJo > 0 && c.realJo > 0 && (!A_ONLY_FROM_RELEASED || c.release !== '') && (!A_ONLY_REV0 || Number(c.rev) === 0)) {
                  r.plan += c.plannedJo;
                  r.real += c.realJo;
                }
              });
              const rows = Array.from(map.values()).sort((x, y) => x.project.localeCompare(y.project));
              const personTsMonths = getPersonMonthlyTimesheet(activePerson.name, timesheetAllData);
              if (rows.length === 0 && personTsMonths.length === 0) return null;
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
                  {rows.length > 0 ? (
                    <div className="space-y-4">
                      {/* Diagram Batang Rincian per Proyek */}
                      <div className="bg-slate-950 border border-slate-800/80 rounded-xl p-4">
                        <div className="relative flex">
                          {/* Sumbu Y */}
                          <div className="flex flex-col justify-between h-44 mt-6 pr-2 text-[9px] font-mono text-slate-500 text-right w-10 shrink-0">
                            <span>{Math.round(maxJo)} jam</span>
                            <span>{Math.round(maxJo * 0.75)}</span>
                            <span>{Math.round(maxJo * 0.5)}</span>
                            <span>{Math.round(maxJo * 0.25)}</span>
                            <span>0</span>
                          </div>

                          <div className="relative flex-1 overflow-x-auto pt-6">
                            {/* Garis Grid Horizontal */}
                            <div className="absolute inset-x-0 top-6 h-44 flex flex-col justify-between pointer-events-none">
                              {[0, 1, 2, 3, 4].map(i => <div key={i} className="border-t border-dashed border-slate-800/90" />)}
                            </div>

                            {/* Batang Diagram per Proyek */}
                            <div className="relative flex items-end gap-8 h-44 px-3 min-w-max">
                              {rows.map(r => {
                                const ps = scorePerson(activePerson.name, activePerson.biro, r.cards);
                                const eff = r.plan > 0 && r.real > 0 ? Math.min((r.plan / r.real) * 100, CAP_EFFICIENCY) : null;
                                const planH = maxJo > 0 ? `${Math.max(3, Math.min(100, (r.plan / maxJo) * 100))}%` : '3%';
                                const realH = maxJo > 0 ? `${Math.max(3, Math.min(100, (r.real / maxJo) * 100))}%` : '3%';
                                const isExpanded = !!expandedProjects[r.project];

                                return (
                                  <div
                                    key={r.project}
                                    onClick={() => toggleProject(r.project)}
                                    className={`group flex flex-col items-center cursor-pointer p-1.5 rounded-lg transition ${
                                      isExpanded ? 'bg-slate-900/80 ring-1 ring-blue-500/50' : 'hover:bg-slate-900/50'
                                    }`}
                                    title={`Klik untuk melihat rincian jobcard proyek ${r.project}`}
                                  >
                                    <div className="flex items-end gap-2 h-44 w-28 justify-center">
                                      {/* Batang Plan JO */}
                                      <div className="flex flex-col items-center justify-end h-full w-10">
                                        <div className="relative w-full" style={{ height: planH }}>
                                          <span className="absolute -top-4 left-1/2 -translate-x-1/2 whitespace-nowrap text-[9px] font-mono font-bold text-violet-300">
                                            {r.plan ? Math.round(r.plan * 10) / 10 : '0'}
                                          </span>
                                          <div className="w-full h-full rounded-t bg-violet-400/85 group-hover:bg-violet-300 transition shadow-sm" />
                                        </div>
                                      </div>

                                      {/* Batang Real JO */}
                                      <div className="flex flex-col items-center justify-end h-full w-10">
                                        <div className="relative w-full" style={{ height: realH }}>
                                          <span className="absolute -top-4 left-1/2 -translate-x-1/2 whitespace-nowrap text-[9px] font-mono font-bold text-emerald-300">
                                            {r.real ? Math.round(r.real * 10) / 10 : '0'}
                                          </span>
                                          <div className="w-full h-full rounded-t bg-emerald-400/85 group-hover:bg-emerald-300 transition shadow-sm" />
                                        </div>
                                      </div>
                                    </div>

                                    {/* Label Bawah Proyek */}
                                    <div className="mt-2 text-center w-28">
                                      <div className="font-mono font-bold text-emerald-400 text-xs truncate flex items-center justify-center gap-1">
                                        <ChevronRight className={`w-3 h-3 text-slate-500 transition-transform ${isExpanded ? 'rotate-90 text-blue-400' : ''}`} />
                                        {r.project}
                                      </div>
                                      <div className="flex items-center justify-center gap-1 mt-1">
                                        <span className="text-[9px] font-bold px-1.5 py-0.2 rounded-full text-slate-950" style={{ background: getCategory(ps.kpi).color }}>
                                          KPI {fmt(ps.kpi, 0)}
                                        </span>
                                        <span className="text-[9px] font-mono text-slate-400">
                                          {eff === null ? '—' : `${Math.round(eff)}%`}
                                        </span>
                                      </div>
                                      <div className="text-[9px] text-amber-400 font-mono mt-0.5">
                                        {r.released}/{r.total} release
                                      </div>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Detail Pekerjaan per Proyek yang Di-expand */}
                      {rows.some(r => expandedProjects[r.project]) && (
                        <div className="space-y-3 pt-2">
                          {rows.filter(r => expandedProjects[r.project]).map(r => {
                            const maxCard = Math.max(1, ...r.cards.flatMap(c => [c.plannedJo, c.realJo]));
                            return (
                              <div key={r.project} className="bg-slate-950/70 border border-slate-800 rounded-xl p-3.5 space-y-2.5">
                                <div className="flex items-center justify-between border-b border-slate-800/80 pb-2 text-xs">
                                  <div className="font-mono font-bold text-emerald-400 flex items-center gap-1.5">
                                    <span className="w-2 h-2 rounded-full bg-emerald-400" />
                                    Jobcard Proyek {r.project} ({r.cards.length})
                                  </div>
                                  <button
                                    onClick={() => toggleProject(r.project)}
                                    className="text-slate-400 hover:text-white text-[11px] cursor-pointer"
                                  >
                                    Tutup
                                  </button>
                                </div>
                                <div className="space-y-2">
                                  {r.cards.map(c => {
                                    const st = cardStatus(c);
                                    const isDescOnly = !!(c as any).noJobcardDesc;
                                    return (
                                      <div key={c.id} className={`p-2.5 bg-slate-900 border rounded-lg space-y-1.5 text-xs ${isDescOnly ? 'border-slate-700/50 opacity-75' : 'border-slate-800'}`}>
                                        <div className="flex items-start justify-between gap-3">
                                          <div className="text-slate-200 font-medium">{c.taskName || '-'}</div>
                                          <div className="flex items-center gap-2 shrink-0">
                                            {!isDescOnly && <span className="font-mono text-[10px] text-slate-400">Rev {c.rev}</span>}
                                            <span className={`text-[10px] font-semibold ${st.cls}`}>{st.text}</span>
                                          </div>
                                        </div>
                                        {!isDescOnly && (
                                          <>
                                            <div className="grid grid-cols-2 gap-2 text-[11px] font-mono pt-1">
                                              <div className="flex items-center gap-1.5">
                                                <span className="text-slate-400">Plan JO:</span>
                                                <span className="text-violet-300 font-bold">{c.plannedJo || '-'}</span>
                                              </div>
                                              <div className="flex items-center gap-1.5">
                                                <span className="text-slate-400">Real JO:</span>
                                                <span className="text-emerald-400 font-bold">{c.realJo ? Math.round(c.realJo * 100) / 100 : '-'}</span>
                                              </div>
                                            </div>
                                            <div className="text-[10px] text-slate-500 font-mono">
                                              {c.kodeJc} · Plan Finish {formatDisplayDate(c.endDate)} · Release {formatDisplayDate(c.release)}
                                            </div>
                                          </>
                                        )}
                                        {isDescOnly && c.realJo > 0 && (
                                          <div className="text-[10px] text-slate-500 font-mono">
                                            Real JO: {Math.round(c.realJo * 100) / 100} jam
                                            {c.release && ` · Drawing release: ${formatDisplayDate(c.release)}`}
                                          </div>
                                        )}
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="text-xs text-slate-500 italic py-1">Belum ada penugasan proyek pada periode terpilih.</div>
                  )}

                  {/* Section: Accomplishment Timesheet Bulanan */}
                  <div className="pt-4 border-t border-slate-800 space-y-3">
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div>
                        <div className="text-xs font-bold text-white flex items-center gap-1.5">
                          <Clock className="w-3.5 h-3.5 text-emerald-400" />
                          Accomplishment Timesheet Bulanan
                        </div>
                        <div className="text-[10px] text-slate-400">
                          Klik batang diagram untuk melihat detail jam/hari kerja
                        </div>
                      </div>
                      <div className="flex items-center gap-3 text-[10px] text-slate-400">
                        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-sky-400 inline-block" /> Tidak Overtime</span>
                        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-purple-400 inline-block" /> Overtime</span>
                      </div>
                    </div>

                    {personTsMonths.length === 0 ? (
                      <div className="text-xs text-slate-500 italic py-1">
                        Belum ada data timesheet bulanan untuk {activePerson.name}. Gunakan tombol <b>Folder Timesheet</b> di Data Master untuk mengimpor folder bulan.
                      </div>
                    ) : (
                      <div className="bg-slate-950 border border-slate-800/80 rounded-xl p-4">
                        <div className="relative flex">
                          {/* Sumbu Y */}
                          <div className="flex flex-col justify-between h-48 mt-5 pr-2 text-[9px] font-mono text-slate-500 text-right w-8 shrink-0">
                            <span>100%</span><span>75%</span><span>50%</span><span>25%</span><span>0%</span>
                          </div>
                          <div className="relative flex-1 overflow-x-auto pt-5">
                            {/* Garis bantu */}
                            <div className="absolute inset-x-0 top-5 h-48 flex flex-col justify-between pointer-events-none">
                              {[0, 1, 2, 3, 4].map(i => <div key={i} className="border-t border-dashed border-slate-800" />)}
                            </div>
                            <div className="relative flex items-end gap-4 h-48 px-2 min-w-max">
                              {personTsMonths.map(pm => {
                                const reg = pm.regular;
                                const ot = pm.overtime;
                                const barH = (v: number) => `${Math.max(2, Math.min(100, v))}%`;
                                return (
                                  <div key={pm.month} className="flex items-end gap-1.5 h-full min-w-[56px] justify-center">
                                    {/* Bar Tidak Overtime */}
                                    <div className="flex flex-col items-center justify-end h-full w-6">
                                      {reg ? (
                                        <div className="relative w-full" style={{ height: barH(reg.accomplishment) }}><span className="absolute -top-4 left-1/2 -translate-x-1/2 whitespace-nowrap text-[9px] font-mono font-bold text-sky-300">{Math.round(reg.accomplishment * 10) / 10}%</span>
                                          <button
                                            type="button"
                                            onClick={() => setTimesheetDetailModal({
                                              month: pm.month,
                                              personName: activePerson.name,
                                              kind: 'regular',
                                              accomplishment: reg.accomplishment,
                                              expectedTs: reg.expectedTotalTimesheet,
                                              totalTs: reg.totalOfTimesheet,
                                            })}
                                            title={`${pm.month} · Tidak Overtime ${Math.round(reg.accomplishment * 10) / 10}% — klik untuk detail`}
                                            className="w-full h-full rounded-t bg-sky-400/80 hover:bg-sky-300 transition cursor-pointer" /></div>
                                      ) : (
                                        <span className="text-[9px] text-slate-600 font-mono">-</span>
                                      )}
                                    </div>
                                    {/* Bar Overtime */}
                                    <div className="flex flex-col items-center justify-end h-full w-6">
                                      {ot ? (
                                        <div className="relative w-full" style={{ height: barH(ot.accomplishment) }}><span className="absolute -top-4 left-1/2 -translate-x-1/2 whitespace-nowrap text-[9px] font-mono font-bold text-purple-300">{Math.round(ot.accomplishment * 10) / 10}%</span>
                                          <button
                                            type="button"
                                            onClick={() => setTimesheetDetailModal({
                                              month: pm.month,
                                              personName: activePerson.name,
                                              kind: 'overtime',
                                              accomplishment: ot.accomplishment,
                                              otHours: ot.totalOvertimeHours,
                                              tsHours: ot.totalTimesheetHours,
                                            })}
                                            title={`${pm.month} · Overtime ${Math.round(ot.accomplishment * 10) / 10}% — klik untuk detail`}
                                            className="w-full h-full rounded-t bg-purple-400/80 hover:bg-purple-300 transition cursor-pointer" /></div>
                                      ) : (
                                        <span className="text-[9px] text-slate-600 font-mono">-</span>
                                      )}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                            {/* Label bulan (sumbu X) */}
                            <div className="flex gap-4 px-2 pt-1.5 border-t border-slate-700 min-w-max">
                              {personTsMonths.map(pm => (
                                <div key={pm.month} className="min-w-[56px] text-center text-[10px] font-semibold text-slate-300">{pm.month}</div>
                              ))}
                            </div>
                          </div>
                        </div>
                      </div>
                    )}
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
                      <th className="py-2 px-2.5">Type Drawing</th>
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
                      <tr><td colSpan={11} className="py-6 text-center text-slate-500 italic">Belum ada penugasan pada periode terpilih.</td></tr>
                    )}
                    {activePerson.cards.map(c => {
                      const st = cardStatus(c);
                      const isDescOnly = !!(c as any).noJobcardDesc;
                      return (
                        <tr key={c.id} className={`hover:bg-slate-950/50 ${isDescOnly ? 'opacity-70' : ''}`}>
                          <td className="py-2 px-2.5 font-mono font-bold text-amber-300">
                            {c.kodeJc || <span className={isDescOnly ? 'text-slate-500 font-normal text-[10px]' : 'text-rose-400 font-normal'}>
                              {isDescOnly ? '(deskripsi)' : '-'}
                            </span>}
                          </td>
                          <td className="py-2 px-2.5">
                            <span className={`text-[10px] px-1.5 py-0.5 rounded font-medium border ${
                              c.typeDrawing === 'Evatek' ? 'bg-purple-500/10 text-purple-300 border-purple-500/30' :
                              c.typeDrawing === 'Tender' ? 'bg-amber-500/10 text-amber-300 border-amber-500/30' :
                              c.typeDrawing === 'Modeling' ? 'bg-emerald-500/10 text-emerald-300 border-emerald-500/30' :
                              'bg-blue-500/10 text-blue-300 border-blue-500/30'
                            }`}>
                              {c.typeDrawing || 'Drawing'}
                            </span>
                          </td>
                          <td className="py-2 px-2.5 text-emerald-400">{c.project}</td>
                          <td className="py-2 px-2.5 text-slate-200">{c.taskName}</td>
                          <td className="py-2 px-2.5 text-center font-mono">{isDescOnly ? '-' : c.rev}</td>
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

      {/* Modal Detail Saat Nilai Accomplishment (%) Dikelik */}
      {timesheetDetailModal && (
        <div
          className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4"
          onClick={() => setTimesheetDetailModal(null)}
        >
          <div
            className="bg-slate-900 border border-slate-700 rounded-2xl max-w-sm w-full p-5 space-y-4 shadow-2xl"
            onClick={e => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span
                    className={`px-2 py-0.5 text-[10px] font-bold uppercase rounded border ${
                      timesheetDetailModal.kind === 'overtime'
                        ? 'bg-purple-500/15 text-purple-300 border-purple-500/30'
                        : 'bg-sky-500/15 text-sky-300 border-sky-500/30'
                    }`}
                  >
                    {timesheetDetailModal.kind === 'overtime' ? 'Overtime' : 'Tidak Overtime (Reguler)'}
                  </span>
                  <span className="text-xs font-semibold text-blue-400">Bulan {timesheetDetailModal.month}</span>
                </div>
                <h3 className="text-sm font-bold text-white mt-1.5">{timesheetDetailModal.personName}</h3>
              </div>
              <button
                onClick={() => setTimesheetDetailModal(null)}
                className="p-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 space-y-2.5">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="text-xs text-slate-400">Accomplishment (%)</span>
                <span className="text-base font-bold font-mono text-emerald-400">
                  {Math.round(timesheetDetailModal.accomplishment * 10) / 10}%
                </span>
              </div>

              {timesheetDetailModal.kind === 'overtime' ? (
                <>
                  <div className="flex items-center justify-between text-xs">
                    <span className="text-slate-300">Total Overtime Hours</span>
                    <span className="font-bold font-mono text-white">
                      {(Math.round((timesheetDetailModal.otHours ?? 0) * 100) / 100).toLocaleString('id-ID')} jam
                    </span>
                  </div>
                  <div className="flex items-center justify-between text-xs">
                    <span className="text-slate-300">Total Timesheet Hours</span>
                    <span className="font-bold font-mono text-purple-300">
                      {(Math.round((timesheetDetailModal.tsHours ?? 0) * 100) / 100).toLocaleString('id-ID')} jam
                    </span>
                  </div>
                </>
              ) : (
                <>
                  <div className="flex items-center justify-between text-xs">
                    <span className="text-slate-300">Expected Total of Timesheet</span>
                    <span className="font-bold font-mono text-white">
                      {(Math.round((timesheetDetailModal.expectedTs ?? 0) * 100) / 100).toLocaleString('id-ID')} hari
                    </span>
                  </div>
                  <div className="flex items-center justify-between text-xs">
                    <span className="text-slate-300">Total of Timesheet</span>
                    <span className="font-bold font-mono text-sky-300">
                      {(Math.round((timesheetDetailModal.totalTs ?? 0) * 100) / 100).toLocaleString('id-ID')} hari
                    </span>
                  </div>
                </>
              )}
            </div>

            <div className="flex justify-end">
              <button
                onClick={() => setTimesheetDetailModal(null)}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg cursor-pointer"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}