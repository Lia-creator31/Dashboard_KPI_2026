import React, { useMemo, useState, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { ArrowLeft, Search, ChevronDown, ChevronUp, Users, FileSpreadsheet } from 'lucide-react';
import { CutiSakitPanel } from './RendalPage';
import type { ParsedMember, SisaCutiData } from './RendalPage';

export interface DivisiGroup {
  deptName: string;
  kadep: ParsedMember[];
  staf: ParsedMember[];
}

interface Props {
  groups: DivisiGroup[];
  wb: XLSX.WorkBook | null;          // TIME_OFF_DIVISI_DESAIN.xlsx
  sisaCuti: SisaCutiData | null;     // Google Sheet Form Alokasi Cuti (termasuk Rencana Cuti 1, 2, ...)
  onBack: () => void;
  onUploadTimeOff: (e: ChangeEvent<HTMLInputElement>) => void;
}

const idOf = (m: ParsedMember) => m.nip || m.nama;

export default function RekapAbsensiDivisiPage({ groups, wb, sisaCuti, onBack, onUploadTimeOff }: Props) {
  const [q, setQ] = useState('');
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});

  const filteredGroups = useMemo(() => {
    const s = q.trim().toLowerCase();
    const match = (m: ParsedMember) =>
      !s || m.nama.toLowerCase().includes(s) || (m.nip || '').toLowerCase().includes(s);
    return groups
      .map(g => ({ ...g, kadep: g.kadep.filter(match), staf: g.staf.filter(match) }))
      .filter(g => g.kadep.length + g.staf.length > 0);
  }, [groups, q]);

  const totalPegawai = useMemo(
    () => filteredGroups.reduce((s, g) => s + g.kadep.length + g.staf.length, 0),
    [filteredGroups]
  );

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <button onClick={onBack} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
            <ArrowLeft className="w-3.5 h-3.5" /> Kembali
          </button>
          <span className="font-bold text-sm text-white">REKAP ABSENSI DIVISI DESAIN</span>
          <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
            <FileSpreadsheet className="w-3.5 h-3.5 text-violet-400" />
            {wb ? 'Update Time Off' : 'Upload Time Off'}
            <input type="file" accept=".xlsx, .xls" onChange={onUploadTimeOff} className="hidden" />
          </label>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-4">
        <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 flex items-start justify-between gap-4 flex-wrap">
          <div className="space-y-1">
            <h3 className="text-xl font-bold text-white">Divisi Desain</h3>
            <div className="text-xs text-slate-400">
              Cuti, Sakit &amp; Rencana Cuti — seluruh departemen, Kadep dan staf
            </div>
          </div>
          <div className="px-5 py-3 bg-slate-950 border border-slate-800 rounded-xl text-right">
            <div className="text-2xl font-bold text-blue-400">{totalPegawai}</div>
            <div className="text-[10px] text-slate-400 uppercase tracking-wide">Total Pegawai</div>
          </div>
        </div>

        <div className="relative w-72">
          <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
          <input
            value={q}
            onChange={e => setQ(e.target.value)}
            placeholder="Cari NIP atau Nama..."
            className="w-full pl-8 pr-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-purple-500"
          />
        </div>

        {filteredGroups.map(g => {
          const members = [...g.kadep, ...g.staf]; // Kadep di atas, staf di bawah
          const kadepIds = new Set(g.kadep.map(idOf));
          const isOpen = !collapsed[g.deptName];
          return (
            <div key={g.deptName} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
              <div
                onClick={() => setCollapsed(p => ({ ...p, [g.deptName]: isOpen }))}
                className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition"
              >
                <div className="flex items-center gap-2">
                  <Users className="w-4 h-4 text-purple-400" />
                  <span className="font-bold text-white text-sm">{g.deptName}</span>
                  <span className="text-[11px] text-slate-400 font-mono">
                    ({g.kadep.length} Kadep, {g.staf.length} staf)
                  </span>
                </div>
                {isOpen ? <ChevronUp className="w-4 h-4 text-slate-400" /> : <ChevronDown className="w-4 h-4 text-slate-400" />}
              </div>
              {isOpen && (
                <div className="p-4 border-t border-slate-800">
                  <CutiSakitPanel
                    biroName={g.deptName}
                    members={members}
                    wb={wb}
                    sisaCuti={sisaCuti}
                    title="Cuti, Sakit & Rencana Cuti"
                    kadepIds={kadepIds}
                  />
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
      </main>
    </div>
  );
}