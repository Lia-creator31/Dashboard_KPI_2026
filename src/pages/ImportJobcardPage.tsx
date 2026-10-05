import React, { useState, useEffect, useMemo, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { ArrowLeft, FileSpreadsheet, Save, AlertTriangle, Check } from 'lucide-react';

interface ImportJobcardPageProps {
  user: UserSession;
  onLogout: () => void;
  onBack?: () => void;
}

/* ============================================================
   PEMETAAN SHEET EXCEL -> DEPARTEMEN (ubah di sini bila berubah)
   ============================================================ */
const SHEET_DEPT: Record<string, string> = {
  'desaindasar': 'Departemen Desain Dasar',
  'spl(1)': 'Departemen Struktur dan Perlengkapan Lambung',
  'spl(2)': 'Departemen Struktur dan Perlengkapan Lambung',
  'spl(3)': 'Departemen Struktur dan Perlengkapan Lambung',
  'spl(4)': 'Departemen Struktur dan Perlengkapan Lambung',
  'mo': 'Departemen Struktur & Perlengkapan Permesinan',
  'eo': 'Departemen Perlengkapan Listrik & Elektronika',
};
const DEPT_ORDER = [
  'Departemen Desain Dasar',
  'Departemen Struktur dan Perlengkapan Lambung',
  'Departemen Struktur & Perlengkapan Permesinan',
  'Departemen Perlengkapan Listrik & Elektronika',
];

// Kolom Excel (indeks mulai 0): B=1 Tgl permintaan, C=2 Project, D=3 Task Name, E=4 Work Center,
// F=5 Start, G=6 End, H=7 Total personil, I=8 PIC, J=9 JO, K=10 Kode Job Card
const COL = { project: 2, task: 3, workCenter: 4, start: 5, end: 6, pic: 8, jo: 9, kode: 10, reqDate: 1 };

const MONTHS = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

const excelGlobUrls = import.meta.glob('./*.xlsx', {
  query: '?url',
  import: 'default',
  eager: true
}) as Record<string, string>;

interface ParsedMember {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
  biro: string;
  dept: string;
}

interface ImportRow {
  key: string;
  sheet: string;
  deptName: string;
  excelRow: number;
  reqDate: string;
  pic: string;
  person: string;
  biroName: string;
  project: string;
  taskName: string;
  startDate: string;
  endDate: string;
  jo: string;
  kodeJc: string;
  rev: string;
  revDetected: boolean;
  approx: boolean;
  include: boolean;
}

/* ============================ Helper ============================ */
function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

const pad = (n: number) => String(n).padStart(2, '0');

const indoMonthsMap: Record<string, string> = {
  jan: '01', januari: '01', january: '01', feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03', apr: '04', april: '04', mei: '05', may: '05',
  jun: '06', juni: '06', june: '06', jul: '07', juli: '07', july: '07',
  agu: '08', ags: '08', agustus: '08', aug: '08', august: '08',
  sep: '09', september: '09', okt: '10', oktober: '10', oct: '10', october: '10',
  nov: '11', november: '11', des: '12', desember: '12', dec: '12', december: '12'
};

function parseTextDate(val: string): string {
  let str = val.trim();
  if (!str) return '';
  str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
  const mDdMon = str.match(/^(\d{1,2})[-/\s.]+([A-Za-z]+)[-/\s.]+(\d{2,4})/);
  if (mDdMon) {
    let y = mDdMon[3]; if (y.length === 2) y = '20' + y;
    const key = mDdMon[2].toLowerCase();
    const mon = indoMonthsMap[key] || indoMonthsMap[key.slice(0, 3)];
    if (mon) return `${y}-${mon}-${mDdMon[1].padStart(2, '0')}`;
  }
  const mDmy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (mDmy) return `${mDmy[3]}-${mDmy[2].padStart(2, '0')}-${mDmy[1].padStart(2, '0')}`;
  return '';
}

// Tanggal dari sel Excel: angka serial, objek Date, atau teks
function parseCellDate(v: any): string {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return `${d.y}-${pad(d.m)}-${pad(d.d)}`;
    }
    return '';
  }
  return parseTextDate(String(v));
}

function cleanProjectString(raw: any): string {
  if (!raw) return '';
  let s = String(raw).replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ').trim();
  s = s.replace(/\s*-\s*/g, '-');
  s = s.replace(/[\.,;:\-_/\\\s]+$/, '').trim().toUpperCase();
  s = s.replace(/([A-Z0-9])O(\d+)$/, '$10$2');
  if (/^W[O0]{2,}\d+/.test(s)) s = 'W' + s.slice(1).replace(/[O0]/g, '0');
  return s.trim();
}

// Revisi dibaca dari nama task: "R2 - ...", "..._R0", "... - R1", "REV 2 - ..."
function detectRev(name: string): { rev: string; detected: boolean } {
  const s = name.trim();
  let m = s.match(/^R\s*(\d+)\s*[-–_ ]/i);
  if (!m) m = s.match(/[\s_-]R\s*(\d+)\s*$/i);
  if (!m) m = s.match(/^REV\.?\s*(\d+)\b/i);
  if (m) return { rev: String(parseInt(m[1], 10)), detected: true };
  return { rev: '0', detected: false };
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

const normDept = (s: string) => cleanText((s || '').replace(/&/g, 'dan'));

const nameTokens = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(Boolean);

const isMultiPerson = (pic: string) => /[,&/+]|\sdan\s|\sdkk\b/i.test(pic);

// Dua kata dianggap sama bila identik, atau salah satunya awalan/singkatan dari yang lain
// (mis. "moh" ~ "mohammad", "ihya" ~ "ihyail", "r" ~ "rahadian")
const tokEq = (a: string, b: string) => a === b || a.startsWith(b) || b.startsWith(a);

// Cocokkan teks PIC ke satu personel secara toleran:
// - nama persis -> langsung cocok
// - selain itu, hitung berapa kata PIC yang cocok dengan tiap personel
//   (PIC 1 kata: butuh 1 kata cocok; PIC 2+ kata: butuh minimal 2 kata cocok)
// - dipilih personel dengan kata cocok terbanyak; bila seri / tidak ada -> null (pilih manual)
function matchPerson(pic: string, pool: ParsedMember[]): ParsedMember | null {
  const cp = cleanText(pic);
  if (!cp || isMultiPerson(pic)) return null;
  const exact = pool.filter(p => cleanText(p.nama) === cp);
  if (exact.length === 1) return exact[0];
  const pt = nameTokens(pic);
  if (!pt.length) return null;
  const need = pt.length === 1 ? 1 : 2;
  const scored = pool
    .map(p => {
      const nt = nameTokens(p.nama);
      return { p, hit: pt.filter(t => nt.some(n => tokEq(t, n))).length };
    })
    .filter(x => x.hit >= need);
  if (!scored.length) return null;
  const best = Math.max(...scored.map(x => x.hit));
  const top = scored.filter(x => x.hit === best);
  return top.length === 1 ? top[0].p : null;
}

function parseIm4(wb: XLSX.WorkBook | null): ParsedMember[] {
  if (!wb) return [];
  try {
    const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes('education')) || wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName]; if (!sheet) return [];
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
      const nama = String(row[namaCol] || '').trim();
      if (!nama || nama.toLowerCase() === 'nan' || nama.toLowerCase() === 'nama') continue;
      const nip = String(row[nipCol] || '').trim();
      const statusRaw = String(row[statusCol] || '').trim();
      const unit = String(row[unitCol] || '').trim();
      const jabatan = String(row[jabatanCol] || '').trim();
      if (jabatan.toLowerCase().includes('kepala divisi')) { currentDept = 'Div. Desain'; currentBiro = 'Div. Desain'; }
      else if (jabatan.toLowerCase().includes('kepala departemen') || jabatan.toLowerCase().includes('kadep')) {
        if (unit && unit.toLowerCase() !== 'nan') currentDept = unit; else currentDept = jabatan;
        currentBiro = `Staf ${currentDept}`;
      } else if (jabatan.toLowerCase().includes('kepala biro') || jabatan.toLowerCase().includes('kabiro')) {
        currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').replace(/Kabiro/gi, 'Biro').trim();
        if (unit && unit.toLowerCase() !== 'nan') currentDept = unit;
      }
      results.push({ nama, nip, status: statusRaw || 'PKWTT', jabatan, biro: currentBiro, dept: currentDept });
    }
    return results;
  } catch { return []; }
}

async function loadIm4Workbook(): Promise<XLSX.WorkBook | null> {
  try {
    const { data, error } = await supabase.storage.from('master-files').download('im4.xlsx');
    if (!error && data) return XLSX.read(await data.arrayBuffer(), { type: 'array' });
  } catch { /* lanjut ke fallback */ }
  const urls: string[] = [];
  Object.entries(excelGlobUrls).forEach(([path, url]) => {
    const p = path.toLowerCase();
    if (p.includes('im4') || p.includes('drawing') || p.includes('akses')) urls.push(url);
  });
  urls.push('/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx');
  for (const u of urls) {
    try {
      const res = await fetch(u);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      const b = new Uint8Array(buf.slice(0, 4));
      if (b[0] === 80 && b[1] === 75 && b[2] === 3 && b[3] === 4) return XLSX.read(buf, { type: 'array' });
    } catch { /* coba berikutnya */ }
  }
  return null;
}

async function fetchAllJobCards(): Promise<any[]> {
  const all: any[] = [];
  for (let from = 0; ; from += 1000) {
    const { data, error } = await supabase.from('job_cards').select('kode_jc, pic, project, task_name, start_date').range(from, from + 999);
    if (error) throw error;
    all.push(...(data || []));
    if (!data || data.length < 1000) break;
  }
  return all;
}

const dupKey = (kode: string, pic: string, project: string, task: string, start: string) =>
  kode ? `k|${cleanText(kode)}|${cleanText(pic)}` : `t|${cleanText(pic)}|${cleanText(project)}|${cleanText(task)}|${(start || '').slice(0, 10)}`;

/* ============================= Halaman ============================= */
export default function ImportJobcardPage({ user, onLogout, onBack }: ImportJobcardPageProps) {
  const [members, setMembers] = useState<ParsedMember[]>([]);
  const [im4Loaded, setIm4Loaded] = useState(false);
  const [jobWb, setJobWb] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState('');
  const [month, setMonth] = useState(9);
  const [year, setYear] = useState(2026);
  const [rows, setRows] = useState<ImportRow[]>([]);
  const [activeDept, setActiveDept] = useState(DEPT_ORDER[0]);
  const [savedKeys, setSavedKeys] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [resultMsg, setResultMsg] = useState('');

  useEffect(() => {
    loadIm4Workbook().then(wb => { setMembers(parseIm4(wb)); setIm4Loaded(true); });
  }, []);

  const deptByName = useMemo(() => {
    const m = new Map<string, Department>();
    DEPT_ORDER.forEach(name => {
      const d = (departmentsData || []).find(x => normDept(x.name) === normDept(name));
      if (d) m.set(name, d);
    });
    return m;
  }, []);

  const poolByDept = useMemo(() => {
    const m = new Map<string, ParsedMember[]>();
    DEPT_ORDER.forEach(name => {
      const d = deptByName.get(name);
      const seen = new Set<string>();
      const pool = members.filter(p => {
        if (!d || !d.biros.some(b => isBiroMatch(p.biro, b.name))) return false;
        const k = p.nip || p.nama;
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
      m.set(name, pool.sort((a, b) => a.nama.localeCompare(b.nama)));
    });
    return m;
  }, [members, deptByName]);

  // Semua personel (unik) - cadangan bila nama PIC tidak ada di daftar departemen sheet-nya
  const allPool = useMemo(() => {
    const seen = new Set<string>();
    return members.filter(p => {
      const k = p.nip || p.nama;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }, [members]);

  const resolveBiroName = (person: ParsedMember, deptName: string): string => {
    const d = deptByName.get(deptName);
    const inDept = d?.biros.find(b => isBiroMatch(person.biro, b.name));
    if (inDept) return inDept.name;
    for (const dep of departmentsData || []) {
      const b = dep.biros.find(x => isBiroMatch(person.biro, x.name));
      if (b) return b.name;
    }
    return '';
  };

  // Bangun ulang baris impor bila file / bulan / daftar personel berubah.
  // Hanya baris dengan Tgl permintaan (kolom B), Start, dan End semuanya di bulan & tahun terpilih yang diambil.
  useEffect(() => {
    if (!jobWb) { setRows([]); return; }
    const ym = `${year}-${pad(month)}`;
    const built: ImportRow[] = [];
    jobWb.SheetNames.forEach(sn => {
      const deptName = SHEET_DEPT[sn.toLowerCase().replace(/\s+/g, '')];
      if (!deptName) return;
      const ws = jobWb.Sheets[sn]; if (!ws) return;
      const pool = poolByDept.get(deptName) || [];
      const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      for (let r = 2; r < data.length; r++) {
        const row = data[r]; if (!row) continue;
        if (String(row[COL.reqDate] ?? '').toUpperCase().includes('CONTOH')) continue;

        // Filter utama: kolom B (Tgl permintaan) harus jatuh di bulan terpilih
        const reqDate = parseCellDate(row[COL.reqDate]);
        if (reqDate.slice(0, 7) !== ym) continue;

        const taskName = String(row[COL.task] ?? '').replace(/\s+/g, ' ').trim();
        const kodeJc = String(row[COL.kode] ?? '').replace(/\s+/g, ' ').trim().toUpperCase();
        if (!taskName && !kodeJc) continue;

        const startDate = parseCellDate(row[COL.start]);
        const endDate = parseCellDate(row[COL.end]);
        // Start dan End juga harus di bulan yang sama
        if (startDate.slice(0, 7) !== ym || endDate.slice(0, 7) !== ym) continue;
        const pic = String(row[COL.pic] ?? '').replace(/\s+/g, ' ').trim();
        const person = matchPerson(pic, pool) || matchPerson(pic, allPool);
        const joNum = parseFloat(String(row[COL.jo] ?? '').replace(',', '.'));
        const { rev, detected } = detectRev(taskName);
        built.push({
          key: `${sn}#${r}`,
          sheet: sn,
          deptName,
          excelRow: r + 1,
          reqDate,
          pic,
          person: person ? person.nama : '',
          biroName: person ? resolveBiroName(person, deptName) : '',
          project: cleanProjectString(row[COL.project]),
          taskName,
          startDate,
          endDate,
          jo: isFinite(joNum) && joNum > 0 ? String(Math.round(joNum * 100) / 100) : '',
          kodeJc,
          rev,
          revDetected: detected,
          approx: !!person && cleanText(person.nama) !== cleanText(pic),
          include: true,
        });
      }
    });
    setRows(built);
    setSavedKeys(new Set());
    setResultMsg('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [jobWb, month, year, poolByDept, allPool]);

  const handleFile = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        setJobWb(XLSX.read(evt.target?.result as ArrayBuffer, { type: 'array' }));
        setFileName(file.name);
      } catch { alert('Gagal membaca file Excel.'); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const updateRow = (key: string, patch: Partial<ImportRow>) =>
    setRows(prev => prev.map(r => (r.key === key ? { ...r, ...patch } : r)));

  const changePerson = (key: string, deptName: string, nama: string) => {
    const p = members.find(m => m.nama === nama);
    updateRow(key, { person: p ? p.nama : '', biroName: p ? resolveBiroName(p, deptName) : '', approx: false });
  };

  const blockers = (r: ImportRow): string[] => {
    const b: string[] = [];
    if (!r.person) b.push(r.pic ? (isMultiPerson(r.pic) ? 'PIC lebih dari satu orang, pilih personel' : 'PIC tidak cocok, pilih personel') : 'PIC kosong, pilih personel');
    else if (!r.biroName) b.push('Biro personel tidak dikenali');
    if (!r.project) b.push('Proyek kosong');
    if (!r.taskName) b.push('Deskripsi kosong');
    if (!r.startDate || !r.endDate) b.push('Tanggal start/finish kosong');
    else if (r.endDate < r.startDate) b.push('Tanggal selesai sebelum mulai');
    return b;
  };

  const notes = (r: ImportRow): string[] => {
    const n: string[] = [];
    if (r.approx) n.push('Nama dicocokkan perkiraan, mohon dicek');
    if (!r.revDetected) n.push('Rev tidak terdeteksi (0)');
    if (!r.kodeJc) n.push('Tanpa kode Jobcard (menunggu Planner)');
    return n;
  };

  const readyRows = rows.filter(r => r.include && !savedKeys.has(r.key) && blockers(r).length === 0);

  const handleSave = async () => {
    if (!readyRows.length) { alert('Tidak ada baris yang siap disimpan.'); return; }
    if (!window.confirm(`Simpan ${readyRows.length} penugasan ke database?`)) return;
    setIsSaving(true);
    setResultMsg('');
    try {
      const { data: biroList } = await supabase.from('biros').select('id, name');
      const existing = await fetchAllJobCards();
      const seen = new Set<string>(existing.map(e => dupKey(String(e.kode_jc || ''), String(e.pic || ''), String(e.project || ''), String(e.task_name || ''), String(e.start_date || ''))));

      let skipped = 0;
      const fresh: ImportRow[] = [];
      readyRows.forEach(r => {
        const k = dupKey(r.kodeJc, r.person, r.project, r.taskName, r.startDate);
        if (seen.has(k)) { skipped++; return; }
        seen.add(k);
        fresh.push(r);
      });

const payloadOf = (r: ImportRow, withRev: boolean) => {
        const biro = (biroList || []).find((b: any) => isBiroMatch(b.name, r.biroName)) || (biroList || [])[0];
        const p: any = {
          biro_id: biro ? biro.id : null,
          biro_name: r.biroName,
          personil_name: r.person,
          project_code: r.project,
          project: r.project,
          task_name: r.taskName,
          start_date: r.startDate,
          end_date: r.endDate,
          pic: r.person,
          jo: r.jo,
          kode_jc: r.kodeJc,
          status: r.kodeJc ? 'approved' : 'pending',
        };
        // Otomatis menyertakan rev dan mengosongkan release awal saat di-import
        if (withRev) { 
          p.rev = r.rev || '0'; 
          p.release = ''; // Release diisi otomatis kosong saat import, nanti diisi via form Rendal bila sudah rilis
        }
        return p;
      };

      let inserted = 0;
      const savedNow: string[] = [];
      let failMsg = '';
      for (let i = 0; i < fresh.length; i += 50) {
        const batch = fresh.slice(i, i + 50);
        let { error } = await supabase.from('job_cards').insert(batch.map(r => payloadOf(r, true)));
        if (error && /rev|release/i.test(error.message || '')) {
          ({ error } = await supabase.from('job_cards').insert(batch.map(r => payloadOf(r, false))));
        }
        if (error) { failMsg = error.message; break; }
        inserted += batch.length;
        batch.forEach(r => savedNow.push(r.key));
      }

      setSavedKeys(prev => new Set([...Array.from(prev), ...savedNow]));
      setResultMsg(
        `Tersimpan ${inserted} penugasan` +
        (skipped ? `, ${skipped} dilewati karena sudah ada di database` : '') +
        (failMsg ? `. Berhenti karena error: ${failMsg}` : '.')
      );
    } catch (err: any) {
      setResultMsg('Gagal menyimpan: ' + (err?.message || 'kesalahan tidak diketahui'));
    } finally {
      setIsSaving(false);
    }
  };

  const deptRows = (deptName: string) => rows.filter(r => r.deptName === deptName);

  const activeRows = useMemo(() => rows.filter(r => r.deptName === activeDept), [rows, activeDept]);
  const groups = useMemo(() => {
    const map = new Map<string, ImportRow[]>();
    activeRows.forEach(r => {
      const k = r.person || '';
      if (!map.has(k)) map.set(k, []);
      map.get(k)!.push(r);
    });
    return Array.from(map.entries()).sort(([a], [b]) => (a === '' ? -1 : b === '' ? 1 : a.localeCompare(b)));
  }, [activeRows]);

  const activePool = poolByDept.get(activeDept) || [];

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            {onBack && (
              <button onClick={onBack} title="Kembali"
                className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
                <ArrowLeft className="w-4 h-4" />
              </button>
            )}
            <div className="p-1.5 bg-emerald-600 rounded-lg text-white"><FileSpreadsheet className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white">IMPORT JOBCARD BULANAN</span>
            <span className="text-xs text-slate-400">— {user.nama}</span>
          </div>
          <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
            Keluar
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-5">
        {/* Pengaturan */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <div className="flex items-end gap-3 flex-wrap text-xs">
            <label className="px-3 py-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer">
              <FileSpreadsheet className="w-4 h-4" /> {fileName ? 'Ganti file' : 'Pilih JOBCARD_DESAIN.xlsx'}
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} className="hidden" />
            </label>
            <div>
              <div className="text-slate-400 mb-1">Bulan (semua tanggal)</div>
              <select value={month} onChange={(e) => setMonth(Number(e.target.value))}
                className="px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white cursor-pointer">
                {MONTHS.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
              </select>
            </div>
            <div>
              <div className="text-slate-400 mb-1">Tahun</div>
              <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value) || year)}
                className="w-24 px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
            </div>
          </div>
          <div className="text-[11px] text-slate-400 space-y-1">
            <div>
              {fileName ? <>File: <b className="text-slate-200">{fileName}</b> · </> : null}
              Master IM4: {im4Loaded ? (members.length > 0 ? <b className="text-emerald-400">{members.length} personel terbaca</b> : <b className="text-rose-400">tidak terbaca, pencocokan nama tidak bisa berjalan</b>) : 'memuat...'}
            </div>
            <div>Hanya baris yang Tgl permintaan (kolom B), Start, dan End-nya semua berada di bulan dan tahun terpilih yang diambil. Nama PIC dicocokkan otomatis ke personel di departemennya. PIC kosong, tidak cocok, atau berisi beberapa nama harus dipilih manual (atau dilewati).</div>
          </div>
        </div>

        {!jobWb && (
          <div className="py-16 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
            Pilih file JOBCARD_DESAIN.xlsx untuk mulai. Hanya baris yang Tgl permintaan, Start, dan End-nya semua di bulan yang dipilih yang akan ditampilkan.
          </div>
        )}

        {jobWb && (
          <>
            {/* Tab departemen */}
            <div className="flex gap-2 flex-wrap">
              {DEPT_ORDER.map(d => {
                const list = deptRows(d);
                const ready = list.filter(r => r.include && !savedKeys.has(r.key) && blockers(r).length === 0).length;
                const need = list.filter(r => blockers(r).length > 0).length;
                return (
                  <button key={d} onClick={() => setActiveDept(d)}
                    className={`px-3 py-2 rounded-xl border text-left text-xs cursor-pointer transition ${activeDept === d ? 'bg-slate-800 border-blue-500' : 'bg-slate-900 border-slate-800 hover:border-slate-600'}`}>
                    <div className="font-semibold text-white">{d.replace('Departemen ', '')}</div>
                    <div className="text-[10px] text-slate-400 mt-0.5">
                      {list.length} baris · <span className="text-emerald-400">{ready} siap</span> · <span className={need ? 'text-amber-400' : 'text-slate-500'}>{need} perlu dicek</span>
                    </div>
                  </button>
                );
              })}
            </div>

            {/* Isi departemen */}
            <div className="space-y-4">
              {groups.length === 0 && (
                <div className="py-10 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
                  Tidak ada baris yang seluruh tanggalnya (permintaan, start, end) di {MONTHS[month - 1]} {year} pada departemen ini.
                </div>
              )}
              {groups.map(([person, list]) => {
                const member = members.find(m => m.nama === person);
                const totalJo = list.reduce((s, r) => s + (parseFloat(r.jo) || 0), 0);
                return (
                  <div key={person || '__manual'} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                    <div className="px-4 py-2.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between gap-3 flex-wrap">
                      <div className="flex items-center gap-2 text-sm">
                        {person ? (
                          <>
                            <span className="font-semibold text-white">{person}</span>
                            {member && <span className="text-xs font-mono text-cyan-400">({member.status})</span>}
                            {list[0].biroName && <span className="text-xs text-slate-400">· {list[0].biroName}</span>}
                          </>
                        ) : (
                          <span className="font-semibold text-amber-400 flex items-center gap-1.5"><AlertTriangle className="w-4 h-4" /> Perlu pilih personel</span>
                        )}
                      </div>
                      <span className="text-[11px] font-mono text-slate-400">{list.length} tugas · Plan JO {Math.round(totalJo * 100) / 100}</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-[11px]">
                        <thead>
                          <tr className="text-slate-400 border-b border-slate-800">
                            <th className="py-2 px-2 w-8"></th>
                            <th className="py-2 px-2">Sheet</th>
                            <th className="py-2 px-2">PIC di Excel</th>
                            <th className="py-2 px-2">Personel</th>
                            <th className="py-2 px-2">Proyek</th>
                            <th className="py-2 px-2">Deskripsi</th>
                            <th className="py-2 px-2 text-center">Rev</th>
                            <th className="py-2 px-2">Plan Start</th>
                            <th className="py-2 px-2">Plan Finish</th>
                            <th className="py-2 px-2 text-right">Plan JO</th>
                            <th className="py-2 px-2">Jobcard</th>
                            <th className="py-2 px-2">Catatan</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800 text-slate-300">
                          {list.map(r => {
                            const bl = blockers(r);
                            const nt = notes(r);
                            const saved = savedKeys.has(r.key);
                            const options = r.person && !activePool.some(p => p.nama === r.person)
                              ? [...activePool, ...members.filter(m => m.nama === r.person)]
                              : activePool;
                            return (
                              <tr key={r.key} className={`align-top ${saved ? 'opacity-50' : ''}`}>
                                <td className="py-2 px-2">
                                  {saved ? <Check className="w-4 h-4 text-emerald-400" /> : (
                                    <input type="checkbox" checked={r.include && bl.length === 0} disabled={bl.length > 0}
                                      onChange={(e) => updateRow(r.key, { include: e.target.checked })} className="cursor-pointer" />
                                  )}
                                </td>
                                <td className="py-2 px-2 whitespace-nowrap text-slate-400">{r.sheet} <span className="text-slate-600">#{r.excelRow}</span></td>
                                <td className="py-2 px-2 text-slate-400 max-w-[130px]">{r.pic || <span className="text-slate-600">-</span>}</td>
                                <td className="py-2 px-2">
                                  <select value={r.person} onChange={(e) => changePerson(r.key, r.deptName, e.target.value)}
                                    className="w-44 px-1.5 py-1 bg-slate-950 border border-slate-700 rounded text-white cursor-pointer">
                                    <option value="">-- Pilih personel --</option>
                                    {options.map(p => <option key={p.nip || p.nama} value={p.nama}>{p.nama}</option>)}
                                  </select>
                                </td>
                                <td className="py-2 px-2 font-mono text-emerald-400 whitespace-nowrap">{r.project || '-'}</td>
                                <td className="py-2 px-2 text-slate-200 min-w-[220px]">{r.taskName || '-'}</td>
                                <td className="py-2 px-2 text-center">
                                  <input type="text" value={r.rev} onChange={(e) => updateRow(r.key, { rev: e.target.value.replace(/[^0-9]/g, ''), revDetected: true })}
                                    className="w-10 px-1 py-1 text-center bg-slate-950 border border-slate-700 rounded text-white font-mono" />
                                </td>
                                <td className="py-2 px-2">
                                  <input type="date" value={r.startDate} onChange={(e) => updateRow(r.key, { startDate: e.target.value })}
                                    style={{ colorScheme: 'dark' }} className="px-1 py-1 bg-slate-950 border border-slate-700 rounded text-white font-mono" />
                                </td>
                                <td className="py-2 px-2">
                                  <input type="date" value={r.endDate} onChange={(e) => updateRow(r.key, { endDate: e.target.value })}
                                    style={{ colorScheme: 'dark' }} className="px-1 py-1 bg-slate-950 border border-slate-700 rounded text-white font-mono" />
                                </td>
                                <td className="py-2 px-2 text-right font-mono text-violet-300">{r.jo || '-'}</td>
                                <td className="py-2 px-2 font-mono font-bold text-amber-300 whitespace-nowrap">{r.kodeJc || <span className="text-rose-400 font-normal">-</span>}</td>
                                <td className="py-2 px-2 min-w-[160px]">
                                  {bl.map(b => <div key={b} className="text-rose-300">{b}</div>)}
                                  {nt.map(n => <div key={n} className="text-slate-500">{n}</div>)}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Bar simpan */}
            <div className="sticky bottom-4 bg-slate-900/95 backdrop-blur border border-slate-700 rounded-2xl p-4 flex items-center justify-between gap-3 flex-wrap shadow-2xl">
              <div className="text-xs">
                <div className="text-slate-200">
                  <b>{readyRows.length}</b> penugasan siap disimpan dari <b>{rows.length}</b> baris {MONTHS[month - 1]} {year}
                </div>
                {resultMsg && <div className="text-emerald-300 mt-1">{resultMsg}</div>}
              </div>
              <button onClick={handleSave} disabled={isSaving || readyRows.length === 0}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-lg flex items-center gap-1.5 cursor-pointer">
                <Save className="w-4 h-4" /> {isSaving ? 'Menyimpan...' : 'Simpan ke Database'}
              </button>
            </div>
          </>
        )}
      </main>
    </div>
  );
}