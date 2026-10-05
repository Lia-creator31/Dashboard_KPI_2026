import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { Building2, ArrowLeft, ChevronRight, RefreshCw, Users } from 'lucide-react';

interface KpiDashboardPageProps {
  user: UserSession;
  onLogout: () => void;
  onBack?: () => void;
}

/* ============================================================
   PENGATURAN ATURAN HITUNG (ubah di sini bila aturan berubah)
   ============================================================ */
const WEIGHT_A = 0.6;                 // bobot Efisiensi JO
const WEIGHT_B = 0.4;                 // bobot Pencapaian Drawing Rev.0
const CAP_EFFICIENCY = 100;           // A dibatasi maksimal 100% (Real JO jauh < Plan JO tidak membuat skor melonjak)
const A_ONLY_FROM_RELEASED = true;    // A hanya dihitung dari jobcard yang drawing-nya sudah released (jam sudah final)
const A_ONLY_REV0 = false;            // true = jam kerja hanya dari drawing Rev.0
const REQUIRE_JOBCARD_CODE = true;    // jobcard tanpa kode (Menunggu Planner) belum dianggap valid
const COUNT_DRAFT_TIMESHEET = false;  // timesheet berstatus Draft tidak dihitung sebagai Real JO
const PARTIAL_KPI = false;            // true = bila hanya A atau B yang ada, KPI memakai komponen yang ada saja

const MASTER_BUCKET = 'master-files';

interface JobCardRow {
  id: string;
  biro: string;
  pic: string;
  project: string;
  taskName: string;
  endDate: string;
  plannedJo: number;
  kodeJc: string;
  rev: string;
  release: string;
  realJo: number;
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
  name: string;
  persons: PersonScore[];
  kpi: number | null;
  a: number | null;
  b: number | null;
  scored: number;
}

/* ============================ Helper ============================ */
function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
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

/* Real JO dari file Realisasi JO.xlsx (sama dengan logika di RendalPage), kunci = kode jobcard */
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
      const j = row.findIndex(c => c.includes('jobcard'));
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

  // B = Drawing Rev.0 released tepat waktu / total target Rev.0 x 100%
  const targets = valid.filter(c => Number(c.rev) === 0);
  const onTime = targets.filter(c => c.release !== '' && c.endDate !== '' && c.release <= c.endDate);
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

/* ============================= Halaman ============================= */
export default function KpiDashboardPage({ user, onLogout, onBack }: KpiDashboardPageProps) {
  const [cards, setCards] = useState<JobCardRow[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState('');
  const [period, setPeriod] = useState('all');
  const [selectedBiro, setSelectedBiro] = useState<string | null>(null);
  const [selectedPerson, setSelectedPerson] = useState<string | null>(null);

  const loadData = useCallback(async () => {
    setIsLoading(true);
    setErrorMsg('');
    try {
      const [jcRes, tsRes, lineRes, wbRes] = await Promise.all([
        supabase.from('job_cards').select('*'),
        supabase.from('outsourcing_timesheets').select('id, status'),
        supabase.from('outsourcing_timesheet_lines').select('timesheet_id, work_order_code, effective_hours, overtime_hours'),
        supabase.storage.from(MASTER_BUCKET).download('realisasi.xlsx'),
      ]);

      if (jcRes.error) throw jcRes.error;

      // Real JO dari Excel Realisasi JO (pegawai organik)
      let excelMap = new Map<string, number>();
      if (!wbRes.error && wbRes.data) {
        try { excelMap = buildRealisasiMap(XLSX.read(await wbRes.data.arrayBuffer(), { type: 'array' })); } catch { /* abaikan */ }
      }

      // Real JO dari timesheet outsourcing (kunci = kode Work Order = kode_jc)
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
        return {
          id: row.id,
          biro: row.biro_name || 'Tanpa Biro',
          pic: String(row.pic || row.personil_name || '').trim(),
          project: row.project || '',
          taskName: row.task_name || '',
          endDate: normalizeDate(row.end_date),
          plannedJo: parseNumber(row.jo),
          kodeJc,
          rev: rawRev.trim() || '0',
          release: normalizeDate(rawRelease),
          realJo: key ? (excelMap.get(key) || 0) + (tsMap.get(key) || 0) : 0,
        };
      }).filter((r: JobCardRow) => r.pic !== '');

      setCards(rows);
    } catch (err: any) {
      console.error('Gagal memuat data KPI:', err);
      setErrorMsg(err?.message || 'Gagal memuat data KPI.');
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => { loadData(); }, [loadData]);

  // Periode = bulan dari Plan Finish
  const periodOptions = useMemo(() => {
    const set = new Set<string>();
    cards.forEach(c => { if (c.endDate) set.add(c.endDate.slice(0, 7)); });
    return Array.from(set).sort().reverse();
  }, [cards]);

  const periodLabel = (p: string) => {
    const [y, m] = p.split('-');
    const names = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return `${names[Number(m) - 1]} ${y}`;
  };

  const biroScores = useMemo<BiroScore[]>(() => {
    const filtered = period === 'all' ? cards : cards.filter(c => c.endDate.startsWith(period));
    const byBiro = new Map<string, Map<string, JobCardRow[]>>();
    filtered.forEach(c => {
      if (!byBiro.has(c.biro)) byBiro.set(c.biro, new Map());
      const pm = byBiro.get(c.biro)!;
      const pk = cleanText(c.pic);
      if (!pm.has(pk)) pm.set(pk, []);
      pm.get(pk)!.push(c);
    });
    const result: BiroScore[] = [];
    byBiro.forEach((pm, biroName) => {
      const persons: PersonScore[] = [];
      pm.forEach(list => persons.push(scorePerson(list[0].pic, biroName, list)));
      persons.sort((x, y) => (y.kpi ?? -1) - (x.kpi ?? -1) || x.name.localeCompare(y.name));
      result.push({
        name: biroName,
        persons,
        kpi: avg(persons.map(p => p.kpi)),
        a: avg(persons.map(p => p.a)),
        b: avg(persons.map(p => p.b)),
        scored: persons.filter(p => p.kpi !== null).length,
      });
    });
    return result.sort((x, y) => x.name.localeCompare(y.name));
  }, [cards, period]);

  const activeBiro = biroScores.find(b => b.name === selectedBiro) || null;
  const activePerson = activeBiro?.persons.find(p => p.name === selectedPerson) || null;
  const divisionKpi = avg(biroScores.map(b => b.kpi));

  const legend = [
    { label: 'Excellent', range: '≥ 90', color: '#34d399' },
    { label: 'Good', range: '80–89', color: '#4cc2ff' },
    { label: 'Fair', range: '70–79', color: '#fbbf24' },
    { label: 'Poor', range: '< 70', color: '#fb7185' },
  ];

  const cardStatus = (c: JobCardRow): { text: string; cls: string } => {
    if (REQUIRE_JOBCARD_CODE && !c.kodeJc) return { text: 'Menunggu Planner', cls: 'text-rose-300' };
    if (Number(c.rev) !== 0) return { text: `Rev ${c.rev} · tidak dihitung di B`, cls: 'text-slate-400' };
    if (!c.release) return { text: 'Belum release', cls: 'text-amber-300' };
    return c.release <= c.endDate
      ? { text: 'Tepat waktu', cls: 'text-emerald-300' }
      : { text: 'Terlambat', cls: 'text-rose-300' };
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            {onBack && (
              <button onClick={onBack} title="Kembali"
                className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
                <ArrowLeft className="w-4 h-4" />
              </button>
            )}
            <div className="p-1.5 bg-purple-600 rounded-lg text-white"><Building2 className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white">DASHBOARD KPI INDIVIDUAL</span>
            <span className="text-xs text-slate-400">— {user.nama}</span>
          </div>
          <div className="flex items-center gap-2">
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
            <h1 className="text-2xl font-extrabold text-white">
              {activePerson ? activePerson.name : activeBiro ? activeBiro.name : 'KPI Individual — Per Biro'}
            </h1>
            <p className="text-xs text-slate-400 mt-1">
              KPI = ({WEIGHT_A * 100}% × A Efisiensi JO) + ({WEIGHT_B * 100}% × B Pencapaian Drawing Rev.0)
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

        {(selectedBiro || selectedPerson) && (
          <button
            onClick={() => (selectedPerson ? setSelectedPerson(null) : setSelectedBiro(null))}
            className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer"
          >
            <ArrowLeft className="w-3.5 h-3.5" /> Kembali
          </button>
        )}

        {errorMsg && <div className="p-3 bg-rose-500/10 border border-rose-500/30 rounded-lg text-xs text-rose-300">{errorMsg}</div>}
        {isLoading && <div className="text-center py-12 text-slate-400 text-sm">Memuat data...</div>}

        {/* ===== LEVEL 1: SEMUA BIRO ===== */}
        {!isLoading && !selectedBiro && (
          <>
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex items-center gap-6 flex-wrap">
              <Donut value={divisionKpi} size={110} color={getCategory(divisionKpi).color} text={fmt(divisionKpi, 0)} />
              <div className="space-y-1">
                <div className="text-xs text-slate-400">Rata-rata KPI seluruh biro</div>
                <Pill score={divisionKpi} />
                <div className="text-[11px] text-slate-500">{biroScores.length} biro · {biroScores.reduce((s, b) => s + b.persons.length, 0)} personel</div>
              </div>
            </div>

            {biroScores.length === 0 ? (
              <div className="py-12 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">Belum ada data penugasan pada periode ini.</div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {biroScores.map(b => (
                  <div key={b.name} onClick={() => setSelectedBiro(b.name)}
                    className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-2xl p-5 cursor-pointer transition flex flex-col items-center gap-3">
                    <div className="w-full flex items-center justify-between">
                      <h3 className="font-bold text-sm text-white">{b.name}</h3>
                      <ChevronRight className="w-4 h-4 text-slate-500" />
                    </div>
                    <Donut value={b.kpi} size={130} color={getCategory(b.kpi).color} text={fmt(b.kpi, 0)} />
                    <Pill score={b.kpi} />
                    <div className="w-full grid grid-cols-3 text-center text-[11px] text-slate-400 pt-2 border-t border-slate-800">
                      <div><div className="font-bold text-slate-200">{fmt(b.a, 0)}{b.a !== null && '%'}</div>A</div>
                      <div><div className="font-bold text-slate-200">{fmt(b.b, 0)}{b.b !== null && '%'}</div>B</div>
                      <div><div className="font-bold text-slate-200">{b.scored}/{b.persons.length}</div>Terhitung</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {/* ===== LEVEL 2: SATU BIRO ===== */}
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
                  <div className="text-[11px] text-slate-400 mt-1">B · Drawing Rev.0</div>
                </div>
              </div>
            </div>

            <div className="lg:col-span-7 bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
              <h2 className="text-sm font-bold text-white flex items-center gap-2"><Users className="w-4 h-4 text-blue-400" /> Skor KPI per personel</h2>
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
                      {p.pendingCount > 0 && <span className="text-rose-300"> · {p.pendingCount} menunggu Planner</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ===== LEVEL 3: SATU INDIVIDU ===== */}
        {!isLoading && activePerson && (
          <div className="space-y-5">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.a} size={130} color="#34d399" text={activePerson.a === null ? '—' : `${Math.round(activePerson.a)}%`} />
                <div className="text-xs font-bold text-white">A · Efisiensi JO</div>
                <div className="text-[11px] text-slate-400">
                  {activePerson.a === null ? 'Belum ada jobcard released dengan Real JO' : 'Planned JO ÷ Actual JO (maks 100%)'}
                </div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.b} size={130} color="#fbbf24" text={activePerson.b === null ? '—' : `${Math.round(activePerson.b)}%`} />
                <div className="text-xs font-bold text-white">B · Drawing Rev.0</div>
                <div className="text-[11px] text-slate-400">
                  {activePerson.targetRev0 === 0
                    ? 'Belum ada target Rev.0'
                    : `${activePerson.onTimeRev0} dari ${activePerson.targetRev0} drawing Rev.0 tepat waktu`}
                </div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col items-center gap-2 text-center">
                <Donut value={activePerson.kpi} size={130} color={getCategory(activePerson.kpi).color} text={fmt(activePerson.kpi, 0)} />
                <div className="text-xs font-bold text-white">Skor KPI Individual</div>
                <Pill score={activePerson.kpi} />
                {activePerson.kpi === null && <div className="text-[11px] text-slate-400">Menunggu data A dan B</div>}
              </div>
            </div>

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
                      <th className="py-2 px-2.5">Plan Finish</th>
                      <th className="py-2 px-2.5">Release</th>
                      <th className="py-2 px-2.5">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
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