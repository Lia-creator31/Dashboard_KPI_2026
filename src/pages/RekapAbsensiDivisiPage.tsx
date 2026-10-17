import React, { useMemo, useState, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { ArrowLeft, Search, Download, ChevronDown, ChevronUp, Users, FileSpreadsheet } from 'lucide-react';
import type { ParsedMember } from './RendalPage';

export interface DivisiGroup {
  deptName: string;
  kadep: ParsedMember[];
  staf: ParsedMember[];
}

interface MonthItem { key: string; label: string; num: string }

interface Props {
  groups: DivisiGroup[];
  wb: XLSX.WorkBook | null;
  months: MonthItem[];
  getCsv: (monthKey: string) => string;
  onBack: () => void;
  onUploadKpi: (e: ChangeEvent<HTMLInputElement>) => void;
}

interface Stat { planned: number; effective: number; overtime: number; idle: number; terlambat: number; sakit: number; found: boolean }

const cleanText = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
const normNip = (s: any) => String(s ?? '').replace(/\D/g, '').replace(/^0+/, '');
const absNum = (v: any): number => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const n = parseFloat(String(v).trim().replace('%', '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
};
const fmt = (n: number) => String(Math.round(n * 100) / 100);
const EMPTY: Stat = { planned: 0, effective: 0, overtime: 0, idle: 0, terlambat: 0, sakit: 0, found: false };

function findMonthSheet(wb: XLSX.WorkBook | null, monthKey: string) {
  if (!wb) return undefined;
  return wb.SheetNames.find(s => {
    const l = s.toLowerCase().trim();
    return l.includes(monthKey) || l.includes(monthKey.slice(0, 3));
  });
}

/** Bangun lookup nilai per orang (by NIP & by nama) dari sheet bulan + CSV absensi */
function buildLookup(wb: XLSX.WorkBook | null, monthKey: string, csv: string) {
  const byNip = new Map<string, Stat>();
  const byName = new Map<string, Stat>();
  const sheetName = findMonthSheet(wb, monthKey);
  if (!wb || !sheetName) return { byNip, byName };

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

  rows.forEach((row, i) => {
    if (i === 0 || !row) return;
    const nip = normNip(row[0]);
    const nama = cleanText(String(row[1] || ''));
    if (!nip && !nama) return;
    const keyName = nama || nip;
    const cur = byName.get(keyName) || { ...EMPTY, found: true };
    cur.planned += absNum(row[PLAN_COL]);
    cur.effective += absNum(row[10]);
    cur.overtime += absNum(row[11]);
    cur.idle += absNum(row[12]);
    byName.set(keyName, cur);
    if (nip) byNip.set(nip, cur);
  });

  byName.forEach((v, k) => {
    const a = abs.get(k);
    if (a) { v.terlambat = a.terlambat; v.sakit = a.sakit; }
  });
  return { byNip, byName };
}

export default function RekapAbsensiDivisiPage({ groups, wb, months, getCsv, onBack, onUploadKpi }: Props) {
  const [monthKey, setMonthKey] = useState<string | null>(null);
  const [q, setQ] = useState('');
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  const month = months.find(m => m.key === monthKey) || null;

  const lookup = useMemo(
    () => (month ? buildLookup(wb, month.key, getCsv(month.key)) : null),
    [wb, month, getCsv]
  );

  const statOf = (m: ParsedMember): Stat => {
    if (!lookup) return EMPTY;
    return lookup.byNip.get(normNip(m.nip)) || lookup.byName.get(cleanText(m.nama)) || EMPTY;
  };

  const filteredGroups = useMemo(() => {
    const s = q.trim().toLowerCase();
    const match = (m: ParsedMember) => !s || m.nama.toLowerCase().includes(s) || (m.nip || '').toLowerCase().includes(s);
    return groups
      .map(g => ({ ...g, kadep: g.kadep.filter(match), staf: g.staf.filter(match) }))
      .filter(g => g.kadep.length + g.staf.length > 0);
  }, [groups, q]);

  const totals = useMemo(() => {
    const t = { pegawai: 0, planned: 0, effective: 0, overtime: 0, idle: 0 };
    filteredGroups.forEach(g => [...g.kadep, ...g.staf].forEach(m => {
      const s = statOf(m);
      t.pegawai++; t.planned += s.planned; t.effective += s.effective; t.overtime += s.overtime; t.idle += s.idle;
    }));
    return t;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filteredGroups, lookup]);

  const exportExcel = () => {
    if (!month) return;
    const aoa: any[][] = [['Departemen', 'Peran', 'NIP', 'Nama', 'Jabatan', 'Planned', 'Effective', 'Overtime', 'Idle', 'Terlambat', 'Sakit']];
    filteredGroups.forEach(g => {
      [...g.kadep.map(m => ['Kadep', m] as const), ...g.staf.map(m => ['Staf', m] as const)].forEach(([peran, m]) => {
        const s = statOf(m);
        aoa.push([g.deptName, peran, m.nip, m.nama, m.jabatan, s.planned, s.effective, s.overtime, s.idle, s.terlambat, s.sakit]);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, ws, month.label.slice(0, 30));
    XLSX.writeFile(book, `Rekap_Absensi_Divisi_Desain_${month.label}.xlsx`);
  };

  const Row = ({ m, i, peran }: { m: ParsedMember; i: number; peran: 'Kadep' | 'Staf' }) => {
    const s = statOf(m);
    return (
      <tr className="hover:bg-slate-800/40">
        <td className="py-2.5 px-3 text-slate-500 font-mono">{i + 1}</td>
        <td className="py-2.5 px-3">
          <div className="font-semibold text-white">{m.nama}</div>
          <div className="text-[10px] text-slate-500 font-mono">{m.nip || '-'}</div>
        </td>
        <td className="py-2.5 px-3">
          <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${peran === 'Kadep'
            ? 'bg-purple-500/10 text-purple-300 border-purple-500/20'
            : 'bg-slate-800 text-slate-300 border-slate-700'}`}>{peran}</span>
        </td>
        <td className="py-2.5 px-3 text-center font-mono text-slate-300">{s.found ? fmt(s.planned) : '-'}</td>
        <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{s.found ? fmt(s.effective) : '-'}</td>
        <td className="py-2.5 px-3 text-center font-mono text-cyan-400">{s.found ? fmt(s.overtime) : '-'}</td>
        <td className="py-2.5 px-3 text-center font-mono text-rose-400">{s.found ? fmt(s.idle) : '-'}</td>
        <td className="py-2.5 px-3 text-center font-mono text-amber-400">{s.found ? fmt(s.terlambat) : '-'}</td>
        <td className="py-2.5 px-3 text-center font-mono text-amber-400">{s.found ? fmt(s.sakit) : '-'}</td>
      </tr>
    );
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <button onClick={onBack} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
            <ArrowLeft className="w-3.5 h-3.5" /> Kembali
          </button>
          <span className="font-bold text-sm text-white">REKAP ABSENSI DIVISI DESAIN</span>
          <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
            <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" /> Upload data_kpi.xlsx
            <input type="file" accept=".xlsx, .xls" onChange={onUploadKpi} className="hidden" />
          </label>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-4">
        {!month ? (
          <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 space-y-4">
            <div>
              <h2 className="text-lg font-bold text-white">Pilih Bulan</h2>
              <span className="text-xs text-slate-400">Rekap seluruh departemen: Kadep dan staf</span>
            </div>
            {!wb && <div className="text-[11px] text-amber-400">File data_kpi.xlsx belum termuat. Upload lewat tombol di kanan atas.</div>}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
              {months.map(m => {
                const ada = !!findMonthSheet(wb, m.key);
                return (
                  <button key={m.key} disabled={!ada} onClick={() => setMonthKey(m.key)}
                    className={`p-4 rounded-xl border text-left transition ${ada ? 'bg-slate-950 border-slate-700 hover:border-purple-500 cursor-pointer' : 'bg-slate-950/40 border-slate-800 opacity-40 cursor-not-allowed'}`}>
                    <div className="font-semibold text-white text-sm">{m.label}</div>
                    <div className="text-[11px] text-slate-500 font-mono">{ada ? 'Data tersedia' : 'Sheet belum ada'}</div>
                  </button>
                );
              })}
            </div>
          </div>
        ) : (
          <>
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 flex items-start justify-between gap-4 flex-wrap">
              <div className="space-y-1">
                <button onClick={() => { setMonthKey(null); setQ(''); }} className="text-xs text-purple-400 hover:underline flex items-center gap-1 cursor-pointer">
                  <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Pilihan Bulan
                </button>
                <h3 className="text-xl font-bold text-white">Divisi Desain</h3>
                <div className="text-xs text-slate-400">Bulan <span className="text-purple-400 font-semibold">{month.label}</span></div>
              </div>
              <button onClick={exportExcel} className="px-3 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer">
                <Download className="w-3.5 h-3.5" /> Export Excel
              </button>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
              {[
                { l: 'Pegawai', v: totals.pegawai, c: 'text-blue-400' },
                { l: 'Planned', v: fmt(totals.planned), c: 'text-slate-200' },
                { l: 'Effective', v: fmt(totals.effective), c: 'text-emerald-400' },
                { l: 'Overtime', v: fmt(totals.overtime), c: 'text-cyan-400' },
                { l: 'Idle', v: fmt(totals.idle), c: 'text-rose-400' },
              ].map(c => (
                <div key={c.l} className="bg-slate-900 border border-slate-800 rounded-xl px-4 py-3">
                  <div className={`text-2xl font-bold ${c.c}`}>{c.v}</div>
                  <div className="text-[10px] text-slate-400 uppercase tracking-wide">{c.l}</div>
                </div>
              ))}
            </div>

            <div className="relative w-72">
              <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
              <input value={q} onChange={e => setQ(e.target.value)} placeholder="Cari NIP atau Nama..."
                className="w-full pl-8 pr-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-purple-500" />
            </div>

            {filteredGroups.map(g => {
              const all = [...g.kadep, ...g.staf];
              const isOpen = !collapsed[g.deptName];
              return (
                <div key={g.deptName} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                  <div onClick={() => setCollapsed(p => ({ ...p, [g.deptName]: isOpen }))}
                    className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition">
                    <div className="flex items-center gap-2">
                      <Users className="w-4 h-4 text-purple-400" />
                      <span className="font-bold text-white text-sm">{g.deptName}</span>
                      <span className="text-[11px] text-slate-400 font-mono">({all.length} orang)</span>
                    </div>
                    {isOpen ? <ChevronUp className="w-4 h-4 text-slate-400" /> : <ChevronDown className="w-4 h-4 text-slate-400" />}
                  </div>
                  {isOpen && (
                    <div className="overflow-x-auto border-t border-slate-800">
                      <table className="w-full text-left text-xs">
                        <thead>
                          <tr className="text-slate-400 border-b border-slate-800 uppercase text-[11px]">
                            <th className="py-2.5 px-3 w-10">No</th>
                            <th className="py-2.5 px-3">Nama Pegawai</th>
                            <th className="py-2.5 px-3">Peran</th>
                            <th className="py-2.5 px-3 text-center">Planned</th>
                            <th className="py-2.5 px-3 text-center">Effective</th>
                            <th className="py-2.5 px-3 text-center">Overtime</th>
                            <th className="py-2.5 px-3 text-center">Idle</th>
                            <th className="py-2.5 px-3 text-center">Terlambat</th>
                            <th className="py-2.5 px-3 text-center">Sakit</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800">
                          {g.kadep.map((m, i) => <Row key={`k-${m.nip}-${i}`} m={m} i={i} peran="Kadep" />)}
                          {g.staf.map((m, i) => <Row key={`s-${m.nip}-${i}`} m={m} i={g.kadep.length + i} peran="Staf" />)}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })}
            {filteredGroups.length === 0 && (
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-8 text-center text-xs text-slate-500">
                Tidak ada data yang cocok.
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
}