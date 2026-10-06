import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { ArrowLeft, RefreshCw, Save, Check, AlertTriangle } from 'lucide-react';

interface SyncReleasePageProps {
  user: UserSession;
  onLogout: () => void;
  onBack?: () => void;
}

/* ===== Drawing Control (Google Drive): ambil Release otomatis ===== */
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

interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
}

const dcClean = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();

const DC_MONTHS: Record<string, string> = {
  jan: '01', januari: '01', january: '01', feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03', apr: '04', april: '04', mei: '05', may: '05',
  jun: '06', juni: '06', june: '06', jul: '07', juli: '07', july: '07',
  agu: '08', ags: '08', agustus: '08', aug: '08', august: '08',
  sep: '09', september: '09', okt: '10', oktober: '10', oct: '10', october: '10',
  nov: '11', november: '11', des: '12', desember: '12', dec: '12', december: '12',
};

// Tanggal teks apa pun (mis. "2026-Sep-10", "10 Sep 2026", "10/09/2026") -> "YYYY-MM-DD"
function parseToStandardDate(val: any): string {
  if (!val) return '';
  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'nan') return '';
  str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
  const mon = (k: string) => DC_MONTHS[k.toLowerCase()] || DC_MONTHS[k.toLowerCase().slice(0, 3)] || '';
  const mYMon = str.match(/^(\d{4})[-/\s]+([A-Za-z]+)[-/\s]+(\d{1,2})/);
  if (mYMon && mon(mYMon[2])) return `${mYMon[1]}-${mon(mYMon[2])}-${mYMon[3].padStart(2, '0')}`;
  const mDdMon = str.match(/^(\d{1,2})[-/\s]+([A-Za-z]+)[-/\s]+(\d{2,4})/);
  if (mDdMon && mon(mDdMon[2])) {
    let y = mDdMon[3]; if (y.length === 2) y = '20' + y;
    return `${y}-${mon(mDdMon[2])}-${mDdMon[1].padStart(2, '0')}`;
  }
  const mDmy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mDmy) return `${mDmy[3]}-${mDmy[2].padStart(2, '0')}-${mDmy[1].padStart(2, '0')}`;
  const mJs = str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (mJs && mon(mJs[1])) return `${mJs[3]}-${mon(mJs[1])}-${mJs[2].padStart(2, '0')}`;
  const num = Number(str);
  if (!isNaN(num) && num > 30000 && num < 60000) {
    return new Date(Math.round((num - 25569) * 86400 * 1000)).toISOString().split('T')[0];
  }
  return '';
}

const dcCache = new Map<string, DrawingControlRow[]>();

// Ambil seluruh baris Drawing Control sebuah proyek: Google Apps Script dulu, lalu CSV Google Sheets
async function fetchDrawingRows(project: string): Promise<DrawingControlRow[]> {
  const proj = dcClean(project);
  if (!proj) return [];
  if (dcCache.has(proj)) return dcCache.get(proj)!;

  try {
    const res = await fetch(`${GAS_DRAWING_API_URL}?project=${encodeURIComponent(project.trim())}`);
    if (res.ok) {
      const json = await res.json();
      if (json.success && Array.isArray(json.data) && json.data.length > 0) {
        const rows: DrawingControlRow[] = json.data.map((item: any) => ({
          noDwg: String(item.noDwg || '').trim(),
          drawingName: String(item.drawingName || '').trim(),
          fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
          rev: String(item.rev || '0').trim(),
          finishDate: String(item.finishDate || '').trim(),
        }));
        dcCache.set(proj, rows);
        return rows;
      }
    }
  } catch { /* lanjut ke cadangan */ }

  const match = GOOGLE_DRIVE_SHEETS.find(s => dcClean(s.projectKey) === proj || proj.includes(dcClean(s.projectKey)));
  if (match) {
    try {
      const res = await fetch(`https://docs.google.com/spreadsheets/d/${match.id}/gviz/tq?tqx=out:csv&sheet=Drawing%20Control%20(2)`);
      if (res.ok) {
        const wb = XLSX.read(await res.text(), { type: 'string' });
        const raw: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
        const rows: DrawingControlRow[] = [];
        for (let r = 1; r < raw.length; r++) {
          const row = raw[r]; if (!row) continue;
          const noDwg = String(row[0] || '').trim();
          const name = String(row[1] || '').trim();
          if (!noDwg && !name) continue;
          rows.push({
            noDwg, drawingName: name,
            fullDeskripsi: noDwg && name ? `${noDwg}-${name}` : (name || noDwg),
            rev: String(row[14] || '0').trim(),
            finishDate: String(row[15] || '').trim(),
          });
        }
        if (rows.length > 0) { dcCache.set(proj, rows); return rows; }
      }
    } catch { /* tidak ada data */ }
  }
  return [];
}

const dcRevNum = (v: any) => {
  const n = parseInt(String(v ?? '').replace(/[^0-9]/g, ''), 10);
  return isNaN(n) ? 0 : n;
};

// Cari tanggal Release (FINISH DATE) untuk gambar + Rev tertentu. Hasil "YYYY-MM-DD", atau '' bila tidak ada.
function extractDrawingNumber(desc: string): string {
  let s = (desc || '').replace(/\s+/g, ' ').trim();
  s = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-–_:]\s*/i, '');
  const m = s.match(/^[A-Z]{1,4}\d{3,}[A-Z0-9]*(?:\.[A-Z0-9]+)*/i);
  if (m) return m[0].toUpperCase();
  const first = s.split(/\s+-\s+|\s+/)[0] || '';
  return /\d/.test(first) && first.length >= 4 ? first.toUpperCase() : '';
}
const drawingKey = (v: string) => dcClean(extractDrawingNumber(v) || v);

function findReleaseDate(rows: DrawingControlRow[], taskName: string, rev: string | number): string {
  if (!rows.length || !taskName) return '';
  const wanted = drawingKey(taskName);
  if (!wanted) return '';
  const wantRev = dcRevNum(rev);
  const hit = rows.find(r => dcRevNum(r.rev) === wantRev && drawingKey(r.noDwg) === wanted && parseToStandardDate(r.finishDate));
  return hit ? parseToStandardDate(hit.finishDate) : '';
}


function clearDrawingCache() { dcCache.clear(); }

const MONTH_NAMES = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

// "W0000304 - 305" -> "W000304" (satu huruf + 6 digit); kode lain dibiarkan
function normalizeProject(raw: string): string {
  const m = (raw || '').toUpperCase().replace(/\s+/g, '').match(/^([A-Z])0*(\d+)/);
  if (!m || m[2].length > 6) return '';
  return m[1] + m[2].padStart(6, '0');
}

async function loadProject(raw: string): Promise<{ rows: DrawingControlRow[]; usedCode: string }> {
  let rows = await fetchDrawingRows(raw);
  if (rows.length) return { rows, usedCode: raw };
  const alt = normalizeProject(raw);
  if (alt && dcClean(alt) !== dcClean(raw)) {
    rows = await fetchDrawingRows(alt);
    if (rows.length) return { rows, usedCode: alt };
  }
  return { rows: [], usedCode: '' };
}

function getLocalRev(id: string): string {
  try { const m = JSON.parse(localStorage.getItem('task_rev_map') || '{}'); return m[id] !== undefined ? String(m[id]) : ''; } catch { return ''; }
}
function getLocalRelease(id: string): string {
  try { const m = JSON.parse(localStorage.getItem('task_release_map') || '{}'); return m[id] !== undefined ? String(m[id]) : ''; } catch { return ''; }
}
function saveLocalRelease(id: string, val: string) {
  try { const m = JSON.parse(localStorage.getItem('task_release_map') || '{}'); m[id] = val; localStorage.setItem('task_release_map', JSON.stringify(m)); } catch { /* abaikan */ }
}

const showDate = (iso: string) => {
  const m = (iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : '-';
};

interface CardRow {
  id: string;
  biro: string;
  pic: string;
  project: string;
  task: string;
  rev: string;
  release: string;     // tersimpan sekarang (YYYY-MM-DD atau '')
  startDate: string;
  kode: string;
}

interface DriveState {
  state: 'loading' | 'ok' | 'empty';
  rows: DrawingControlRow[];
  usedCode: string;
}

type Status = 'loading' | 'new' | 'same' | 'diff' | 'noproj' | 'notfound';

const STATUS_LABEL: Record<Status, string> = {
  loading: 'Membaca Drive...',
  new: 'Baru dari Drive',
  same: 'Sudah sama',
  diff: 'Beda dengan Drive',
  noproj: 'Proyek tidak ada di Drive',
  notfound: 'Gambar / Rev tidak ada di Drive',
};
const STATUS_CLASS: Record<Status, string> = {
  loading: 'text-slate-500',
  new: 'text-emerald-400',
  same: 'text-slate-400',
  diff: 'text-amber-400',
  noproj: 'text-rose-300',
  notfound: 'text-rose-300',
};

export default function SyncReleasePage({ user, onLogout, onBack }: SyncReleasePageProps) {
  const [cards, setCards] = useState<CardRow[]>([]);
  const [loadingCards, setLoadingCards] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [driveMap, setDriveMap] = useState<Record<string, DriveState>>({});
  const [month, setMonth] = useState(0);              // 0 = semua bulan
  const [year, setYear] = useState(2026);
  const [statusFilter, setStatusFilter] = useState<'all' | Status>('all');
  const [search, setSearch] = useState('');
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [resultMsg, setResultMsg] = useState('');
  const [reloadTick, setReloadTick] = useState(0);

  // 1) Ambil semua job card dari database
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingCards(true);
      setLoadError('');
      try {
        const all: any[] = [];
        for (let from = 0; ; from += 1000) {
          const { data, error } = await supabase.from('job_cards').select('*').order('created_at', { ascending: true }).range(from, from + 999);
          if (error) throw error;
          all.push(...(data || []));
          if (!data || data.length < 1000) break;
        }
        if (cancelled) return;
        setCards(all.map((row: any) => ({
          id: String(row.id),
          biro: row.biro_name || '',
          pic: row.pic || row.personil_name || '',
          project: String(row.project || row.project_code || '').trim(),
          task: row.task_name || '',
          rev: (row.rev !== undefined && row.rev !== null && String(row.rev).trim() !== '') ? String(row.rev) : (getLocalRev(String(row.id)) || '0'),
          release: parseToStandardDate((row.release !== undefined && row.release !== null && String(row.release).trim() !== '') ? row.release : getLocalRelease(String(row.id))),
          startDate: String(row.start_date || '').slice(0, 10),
          kode: row.kode_jc || '',
        })));
      } catch (e: any) {
        if (!cancelled) setLoadError(e?.message || 'Gagal memuat job card');
      } finally {
        if (!cancelled) setLoadingCards(false);
      }
    })();
    return () => { cancelled = true; };
  }, [reloadTick]);

  // 2) Filter bulan (berdasarkan Plan Start)
  const monthCards = useMemo(() => {
    if (!month) return cards;
    const ym = `${year}-${String(month).padStart(2, '0')}`;
    return cards.filter(c => c.startDate.slice(0, 7) === ym);
  }, [cards, month, year]);

  // 3) Baca Drawing Control otomatis untuk setiap proyek yang tampil (tanpa klik apa pun)
  useEffect(() => {
    const needed = Array.from(new Set(monthCards.map(c => c.project).filter(Boolean)));
    const todo = needed.filter(p => !driveMap[p]);
    if (!todo.length) return;
    let cancelled = false;
    setDriveMap(prev => {
      const next = { ...prev };
      todo.forEach(p => { next[p] = { state: 'loading', rows: [], usedCode: '' }; });
      return next;
    });
    (async () => {
      for (let i = 0; i < todo.length; i += 4) {
        const chunk = todo.slice(i, i + 4);
        const results = await Promise.all(chunk.map(async p => {
          try { return [p, await loadProject(p)] as const; }
          catch { return [p, { rows: [] as DrawingControlRow[], usedCode: '' }] as const; }
        }));
        if (cancelled) return;
        setDriveMap(prev => {
          const next = { ...prev };
          results.forEach(([p, r]) => { next[p] = { state: r.rows.length ? 'ok' : 'empty', rows: r.rows, usedCode: r.usedCode }; });
          return next;
        });
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [monthCards]);

  // 4) Cocokkan tiap job card dengan Drawing Control
  const analyzed = useMemo(() => {
    return monthCards.map(c => {
      const d = c.project ? driveMap[c.project] : undefined;
      let status: Status = 'loading';
      let drive = '';
      if (!c.project || (d && d.state === 'empty')) status = 'noproj';
      else if (d && d.state === 'ok') {
        drive = findReleaseDate(d.rows, c.task, c.rev);
        if (!drive) status = 'notfound';
        else if (!c.release) status = 'new';
        else status = c.release === drive ? 'same' : 'diff';
      }
      return { card: c, drive, status };
    });
  }, [monthCards, driveMap]);

  const counts = useMemo(() => {
    const k: Record<string, number> = { all: analyzed.length };
    analyzed.forEach(a => { k[a.status] = (k[a.status] || 0) + 1; });
    return k;
  }, [analyzed]);

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    return analyzed.filter(a =>
      (statusFilter === 'all' || a.status === statusFilter) &&
      (!q || `${a.card.pic} ${a.card.project} ${a.card.task} ${a.card.kode}`.toLowerCase().includes(q)));
  }, [analyzed, statusFilter, search]);

  // Otomatis centang baris "Baru dari Drive" (sekali per baris, agar centang manual tidak ditimpa)
  const autoSeen = useRef<Set<string>>(new Set());
  useEffect(() => {
    const add: string[] = [];
    analyzed.forEach(a => {
      if (a.status === 'new' && !autoSeen.current.has(a.card.id)) { autoSeen.current.add(a.card.id); add.push(a.card.id); }
    });
    if (add.length) setSelected(prev => { const n = new Set(prev); add.forEach(id => n.add(id)); return n; });
  }, [analyzed]);

  const toggle = (id: string) => setSelected(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const applicable = analyzed.filter(a => (a.status === 'new' || a.status === 'diff') && selected.has(a.card.id));

  const handleApply = async () => {
    if (!applicable.length) return;
    if (!window.confirm(`Simpan ${applicable.length} tanggal Release ke database?`)) return;
    setIsSaving(true);
    setResultMsg('');
    let ok = 0, local = 0, failed = 0;
    const done = new Map<string, string>();
    for (let i = 0; i < applicable.length; i += 10) {
      await Promise.all(applicable.slice(i, i + 10).map(async a => {
        const { error } = await supabase.from('job_cards').update({ release: a.drive }).eq('id', a.card.id);
        if (!error) { ok++; done.set(a.card.id, a.drive); saveLocalRelease(a.card.id, a.drive); }
        else if (/release/i.test(error.message || '')) { saveLocalRelease(a.card.id, a.drive); local++; done.set(a.card.id, a.drive); }
        else failed++;
      }));
    }
    setCards(prev => prev.map(c => done.has(c.id) ? { ...c, release: done.get(c.id)! } : c));
    setSelected(prev => { const n = new Set(prev); done.forEach((_, id) => n.delete(id)); return n; });
    setResultMsg(
      `Tersimpan ${ok} Release ke database` +
      (local ? `, ${local} hanya di peramban ini (kolom release belum ada di database)` : '') +
      (failed ? `, ${failed} gagal` : '') + '.'
    );
    setIsSaving(false);
  };

  const handleRefreshDrive = () => {
    clearDrawingCache();
    setDriveMap({});
    setResultMsg('');
  };

  const driveSummary = useMemo(() => {
    return Object.entries(driveMap).map(([p, d]) => ({ p, ...d })).sort((a, b) => a.p.localeCompare(b.p));
  }, [driveMap]);

  const chip = (key: 'all' | Status, label: string) => (
    <button key={key} onClick={() => setStatusFilter(key)}
      className={`px-2.5 py-1 rounded-lg border text-[11px] cursor-pointer ${statusFilter === key ? 'bg-slate-800 border-blue-500 text-white' : 'bg-slate-900 border-slate-800 text-slate-400 hover:border-slate-600'}`}>
      {label} <b>{counts[key] || 0}</b>
    </button>
  );

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
            <div className="p-1.5 bg-cyan-600 rounded-lg text-white"><RefreshCw className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white">SINKRON RELEASE DARI DRAWING CONTROL</span>
            <span className="text-xs text-slate-400">— {user.nama}</span>
          </div>
          <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
            Keluar
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-5">
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
          <div className="flex items-end gap-3 flex-wrap text-xs">
            <div>
              <div className="text-slate-400 mb-1">Bulan (Plan Start)</div>
              <select value={month} onChange={(e) => setMonth(Number(e.target.value))}
                className="px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white cursor-pointer">
                <option value={0}>Semua bulan</option>
                {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
              </select>
            </div>
            <div>
              <div className="text-slate-400 mb-1">Tahun</div>
              <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value) || year)} disabled={!month}
                className="w-24 px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono disabled:opacity-40" />
            </div>
            <div>
              <div className="text-slate-400 mb-1">Cari</div>
              <input type="text" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Nama, proyek, gambar..."
                className="w-56 px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" />
            </div>
            <button onClick={handleRefreshDrive}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg border border-slate-700 flex items-center gap-1.5 cursor-pointer">
              <RefreshCw className="w-3.5 h-3.5" /> Baca ulang Drive
            </button>
            <button onClick={() => { setReloadTick(t => t + 1); }}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg border border-slate-700 cursor-pointer">
              Muat ulang database
            </button>
          </div>
          <div className="text-[11px] text-slate-400">
            Drawing Control dibaca otomatis dari Google Drive untuk setiap proyek yang tampil. Release dicocokkan lewat <b>nomor gambar di deskripsi</b> dan <b>Rev</b> yang sama.
            Baris berstatus "Baru dari Drive" otomatis tercentang; "Beda dengan Drive" perlu Anda centang sendiri bila ingin menimpa.
          </div>
          {driveSummary.length > 0 && (
            <div className="flex flex-wrap gap-1.5 text-[10px] font-mono">
              {driveSummary.map(d => (
                <span key={d.p} className={`px-2 py-0.5 rounded border ${d.state === 'ok' ? 'border-emerald-500/30 text-emerald-300' : d.state === 'loading' ? 'border-slate-700 text-slate-400' : 'border-rose-500/30 text-rose-300'}`}>
                  {d.p}{d.usedCode && d.usedCode !== d.p ? ` → ${d.usedCode}` : ''}: {d.state === 'ok' ? `${d.rows.length} gambar` : d.state === 'loading' ? 'membaca...' : 'tidak ada data'}
                </span>
              ))}
            </div>
          )}
        </div>

        {loadError && <div className="p-3 bg-rose-500/10 border border-rose-500/30 rounded-lg text-xs text-rose-300">{loadError}</div>}

        <div className="flex gap-2 flex-wrap">
          {chip('all', 'Semua')}
          {chip('new', STATUS_LABEL.new)}
          {chip('diff', STATUS_LABEL.diff)}
          {chip('same', STATUS_LABEL.same)}
          {chip('notfound', STATUS_LABEL.notfound)}
          {chip('noproj', STATUS_LABEL.noproj)}
        </div>

        <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-[11px]">
              <thead>
                <tr className="text-slate-400 border-b border-slate-800 bg-slate-950">
                  <th className="py-2 px-2 w-8"></th>
                  <th className="py-2 px-2">Personel</th>
                  <th className="py-2 px-2">Jobcard</th>
                  <th className="py-2 px-2">Proyek</th>
                  <th className="py-2 px-2">Deskripsi</th>
                  <th className="py-2 px-2 text-center">Rev</th>
                  <th className="py-2 px-2">Release sekarang</th>
                  <th className="py-2 px-2">Release di Drive</th>
                  <th className="py-2 px-2">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800 text-slate-300">
                {loadingCards && <tr><td colSpan={9} className="py-10 text-center text-slate-500">Memuat job card...</td></tr>}
                {!loadingCards && visible.length === 0 && <tr><td colSpan={9} className="py-10 text-center text-slate-500">Tidak ada baris.</td></tr>}
                {visible.slice(0, 500).map(({ card: c, drive, status }) => {
                  const canPick = status === 'new' || status === 'diff';
                  return (
                    <tr key={c.id} className="align-top">
                      <td className="py-2 px-2">
                        {canPick && <input type="checkbox" checked={selected.has(c.id)} onChange={() => toggle(c.id)} className="cursor-pointer" />}
                      </td>
                      <td className="py-2 px-2 whitespace-nowrap">{c.pic || '-'}</td>
                      <td className="py-2 px-2 font-mono text-amber-300 whitespace-nowrap">{c.kode || '-'}</td>
                      <td className="py-2 px-2 font-mono text-emerald-400 whitespace-nowrap">{c.project || '-'}</td>
                      <td className="py-2 px-2 text-slate-200 min-w-[240px]">{c.task || '-'}</td>
                      <td className="py-2 px-2 text-center font-mono">{c.rev}</td>
                      <td className="py-2 px-2 font-mono text-cyan-300 whitespace-nowrap">{showDate(c.release)}</td>
                      <td className="py-2 px-2 font-mono whitespace-nowrap">{drive ? <span className="text-emerald-300 font-semibold">{showDate(drive)}</span> : '-'}</td>
                      <td className={`py-2 px-2 ${STATUS_CLASS[status]}`}>
                        {status === 'diff' && <AlertTriangle className="w-3 h-3 inline mr-1" />}
                        {status === 'same' && <Check className="w-3 h-3 inline mr-1" />}
                        {STATUS_LABEL[status]}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {visible.length > 500 && <div className="p-2 text-center text-[11px] text-slate-500 border-t border-slate-800">Menampilkan 500 dari {visible.length} baris. Persempit dengan filter bulan atau pencarian.</div>}
        </div>

        <div className="sticky bottom-4 bg-slate-900/95 backdrop-blur border border-slate-700 rounded-2xl p-4 flex items-center justify-between gap-3 flex-wrap shadow-2xl">
          <div className="text-xs">
            <div className="text-slate-200"><b>{applicable.length}</b> tanggal Release siap diterapkan dari <b>{analyzed.length}</b> job card</div>
            {resultMsg && <div className="text-emerald-300 mt-1">{resultMsg}</div>}
          </div>
          <button onClick={handleApply} disabled={isSaving || applicable.length === 0}
            className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-lg flex items-center gap-1.5 cursor-pointer">
            <Save className="w-4 h-4" /> {isSaving ? 'Menyimpan...' : 'Terapkan Release ke Database'}
          </button>
        </div>
      </main>
    </div>
  );
}