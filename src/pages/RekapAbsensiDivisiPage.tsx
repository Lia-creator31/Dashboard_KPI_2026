import React, { useMemo, useState, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ArrowLeft, Search, FileSpreadsheet, FileDown, FileText } from 'lucide-react';
import {
  parseTimeOff, findSisaCuti, countDaysAfter, normNip, cleanText, shortType, formatDisplayDate,
} from './RendalPage';
import type { ParsedMember, SisaCutiData, SisaCutiInfo, TimeOffEntry } from './RendalPage';

export interface DivisiGroup {
  deptName: string;
  kadep: ParsedMember[];
  staf: ParsedMember[];
}

interface Props {
  groups: DivisiGroup[];
  wb: XLSX.WorkBook | null;          // TIME_OFF_DIVISI_DESAIN.xlsx
  sisaCuti: SisaCutiData | null;     // Google Sheet Form Alokasi Cuti (Rencana Cuti 1, 2, ...)
  onBack: () => void;
  onUploadTimeOff: (e: ChangeEvent<HTMLInputElement>) => void;
}

interface Row {
  m: ParsedMember;
  isKadep: boolean;
  cuti: number;
  panjang: TimeOffEntry[];
  khusus: TimeOffEntry[];
  sakit: number;
  sakitKali: number;
  sisa?: number;
  filledAt?: string;
  cutiSetelahIsi: number;
  rencana: Record<string, string>;
}

type TabKey = 'realisasi' | 'plan';

const idOf = (m: ParsedMember) => m.nip || m.nama;

export default function RekapAbsensiDivisiPage({ groups, wb, sisaCuti, onBack, onUploadTimeOff }: Props) {
  const [q, setQ] = useState('');
  const [tab, setTab] = useState<TabKey>('realisasi');
  const year = String(new Date().getFullYear());
  const entries = useMemo(() => parseTimeOff(wb), [wb]);

  // Kolom "Rencana Cuti 1, 2, ..." dinamis dari Google Sheet (hanya dipakai di tab Plan Cuti)
  const rencanaKeys = useMemo(() => {
    if (!sisaCuti) return [] as string[];
    const keys = new Set<string>();
    sisaCuti.byName.forEach(info => Object.keys(info.rencana || {}).forEach(k => keys.add(k)));
    return Array.from(keys).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
    );
  }, [sisaCuti]);

  const sections = useMemo(() => {
    const s = q.trim().toLowerCase();
    const match = (m: ParsedMember) =>
      !s || m.nama.toLowerCase().includes(s) || (m.nip || '').toLowerCase().includes(s);
    const currentYear = new Date().getFullYear();

    return groups
      .map(g => {
        const kadepIds = new Set(g.kadep.map(idOf));
        const allMembers = [...g.kadep, ...g.staf];
        const allNames = allMembers.map(x => x.nama);
        const rows: Row[] = allMembers.filter(match).map(m => {
          const nip = normNip(m.nip), nm = cleanText(m.nama);
          const mine = entries.filter(e =>
            e.status === 'Approved' && ((nip && normNip(e.nip) === nip) || cleanText(e.nama) === nm));
          const inYear = currentYear === 2026
            ? mine.filter(e => {
                const [y, mm] = e.start.split('-');
                const n = parseInt(mm, 10);
                return y === '2026' && n >= 8 && n <= 12;
              })
            : mine.filter(e => e.start.startsWith(String(currentYear)));
          const sum = (k: TimeOffEntry['kind']) =>
            inYear.filter(e => e.kind === k).reduce((a, e) => a + e.days, 0);
          const info: SisaCutiInfo | undefined = findSisaCuti(sisaCuti, m.nama, allNames);
          const cutiSetelahIsi = info
            ? inYear.filter(e => e.kind === 'cuti').reduce((a, e) => a + countDaysAfter(e, info.filledAt), 0)
            : 0;
          return {
            m,
            isKadep: kadepIds.has(idOf(m)),
            cuti: sum('cuti'),
            panjang: inYear.filter(e => e.kind === 'panjang'),
            khusus: inYear.filter(e => e.kind === 'khusus'),
            sakit: sum('sakit'),
            sakitKali: inYear.filter(e => e.kind === 'sakit').length,
            sisa: info ? info.sisa - cutiSetelahIsi : undefined,
            filledAt: info?.filledAt,
            cutiSetelahIsi,
            rencana: info?.rencana || {},
          };
        });
        return { deptName: g.deptName, rows };
      })
      .filter(sec => sec.rows.length > 0);
  }, [groups, entries, sisaCuti, q]);

  const showRencana = tab === 'plan';
  const colSpan = 5 + (showRencana ? rencanaKeys.length : 0);
  let no = 0;

  // ===== EXPORT (Excel & PDF) =====
  // Mengikuti tab aktif (Realisasi Cuti / Plan Cuti) dan filter pencarian yang sedang dipakai.
  const buildExport = () => {
    const isPlan = tab === 'plan';
    const head = [
      'No', 'Nama Pegawai', 'Departemen', `Realisasi Cuti ${year} (hari)`, 'Keterangan',
      'Sisa Cuti (hari)', 'Sakit (hari)',
      ...(isPlan ? rencanaKeys : []),
    ];
    const rows: (string | number)[][] = [];
    const groupsOut: { deptName: string; rows: (string | number)[][] }[] = [];
    let n = 0;

    sections.forEach(sec => {
      const secRows: (string | number)[][] = [];
      sec.rows.forEach(d => {
        n += 1;
        const ket = [
          ...d.panjang.map(e => `${shortType(e.type)} ${e.days} hari`),
          ...d.khusus.map(e => `${shortType(e.type)} ${e.days} hari`),
        ].join(', ');
        const row: (string | number)[] = [
          n,
          d.m.nama + (d.isKadep ? ' (Kadep)' : ''),
          sec.deptName,
          d.cuti,
          ket ? `${ket} (tidak dihitung)` : '',
          d.sisa !== undefined ? d.sisa : '-',
          d.sakit,
          ...(isPlan ? rencanaKeys.map(k => d.rencana[k] || '-') : []),
        ];
        rows.push(row);
        secRows.push(row);
      });
      groupsOut.push({ deptName: sec.deptName, rows: secRows });
    });

    const title = isPlan ? 'Plan Cuti' : 'Realisasi Cuti';
    return { head, rows, groupsOut, title };
  };

  const exportExcel = () => {
    const { head, rows, title } = buildExport();
    const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
    ws['!cols'] = head.map((h, i) => ({
      wch: i === 1 ? 32 : i === 2 ? 38 : i === 4 ? 36 : Math.max(10, h.length + 2),
    }));
    const wbOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbOut, ws, title);
    XLSX.writeFile(wbOut, `Rekap_${title.replace(' ', '_')}_Divisi_Desain_${year}.xlsx`);
  };

  const exportPdf = () => {
    const { head, groupsOut, title } = buildExport();
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });

    doc.setFontSize(13);
    doc.text(`REKAP ${title.toUpperCase()} DIVISI DESAIN ${year}`, 40, 36);
    doc.setFontSize(8);
    doc.text(`Dicetak: ${new Date().toLocaleDateString('id-ID')}`, 40, 50);

    // Departemen jadi baris pemisah, jadi kolom Departemen (indeks 2) dibuang di PDF
    const pdfHead = head.filter((_, i) => i !== 2);
    const body: any[] = [];
    groupsOut.forEach(g => {
      body.push([{
        content: `${g.deptName} (${g.rows.length} orang)`,
        colSpan: pdfHead.length,
        styles: { fillColor: [233, 225, 250], textColor: [76, 29, 149], fontStyle: 'bold' },
      }]);
      g.rows.forEach(r => body.push(r.filter((_, i) => i !== 2)));
    });

    autoTable(doc, {
      head: [pdfHead],
      body,
      startY: 60,
      theme: 'grid',
      styles: { fontSize: 7.5, cellPadding: 3, valign: 'middle' },
      headStyles: { fillColor: [30, 41, 59], textColor: 255, halign: 'center' },
      columnStyles: { 0: { halign: 'center', cellWidth: 28 } },
      didDrawPage: () => {
        const p = doc.getNumberOfPages();
        doc.setFontSize(8);
        doc.text(`Hal. ${p}`, doc.internal.pageSize.getWidth() - 60, doc.internal.pageSize.getHeight() - 20);
      },
    });

    doc.save(`Rekap_${title.replace(' ', '_')}_Divisi_Desain_${year}.pdf`);
  };

  const tabClass = (active: boolean, activeColor: string) =>
    `px-4 py-2 text-xs font-semibold rounded-lg border transition cursor-pointer ${
      active
        ? `${activeColor} text-white border-transparent shadow`
        : 'bg-slate-900 text-slate-400 border-slate-700 hover:text-white hover:border-slate-500'
    }`;

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
        {/* Dua tombol tampilan */}
        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={() => setTab('realisasi')}
            className={tabClass(tab === 'realisasi', 'bg-emerald-600')}
          >
            Realisasi Cuti
          </button>
          <button
            onClick={() => setTab('plan')}
            className={tabClass(tab === 'plan', 'bg-blue-600')}
          >
            Plan Cuti
          </button>
        </div>

        {/* Pencarian + tombol export */}
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div className="relative w-72">
            <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
            <input
              value={q}
              onChange={e => setQ(e.target.value)}
              placeholder="Cari NIP atau Nama..."
              className="w-full pl-8 pr-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:border-purple-500"
            />
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={exportExcel}
              disabled={!wb || sections.length === 0}
              className="px-3 py-2 text-xs font-semibold rounded-lg border border-emerald-700 bg-emerald-600/10 text-emerald-300 hover:bg-emerald-600/20 flex items-center gap-1.5 cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <FileDown className="w-3.5 h-3.5" /> Export Excel
            </button>
            <button
              onClick={exportPdf}
              disabled={!wb || sections.length === 0}
              className="px-3 py-2 text-xs font-semibold rounded-lg border border-red-700 bg-red-600/10 text-red-300 hover:bg-red-600/20 flex items-center gap-1.5 cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <FileText className="w-3.5 h-3.5" /> Export PDF
            </button>
          </div>
        </div>

        {!wb ? (
          <div className="text-[11px] text-amber-400 bg-slate-900 border border-slate-800 rounded-xl p-4">
            Data Time Off belum termuat. Upload file TIME_OFF_DIVISI_DESAIN.xlsx lewat tombol di kanan atas.
          </div>
        ) : (
          <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-auto max-h-[calc(100vh-13rem)]">
            <table className="w-full min-w-max text-left text-xs">
              <thead>
                <tr className="text-slate-400 uppercase text-[11px]">
                  <th className="sticky top-0 left-0 z-30 bg-slate-900 border-b border-slate-800 py-3 px-3 w-10 min-w-[40px] max-w-[40px]">No</th>
                  <th className="sticky top-0 left-10 z-30 bg-slate-900 border-b border-slate-800 py-3 px-3 min-w-[220px]">Nama Pegawai</th>
                  <th className="sticky top-0 z-20 bg-slate-900 border-b border-slate-800 py-3 px-3">Realisasi Cuti {year}</th>
                  <th className="sticky top-0 z-20 bg-sky-950 border-b border-slate-800 py-3 px-8 text-center text-sky-400">Sisa Cuti</th>
                  <th className="sticky top-0 z-20 bg-amber-950 border-b border-slate-800 py-3 px-3 text-center text-amber-400">Sakit</th>
                  {showRencana && rencanaKeys.map(key => (
                    <th key={key} className="sticky top-0 z-20 bg-blue-950 border-b border-slate-800 py-3 px-3 text-center text-blue-400 whitespace-nowrap">{key}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sections.length === 0 && (
                  <tr><td colSpan={colSpan} className="py-8 text-center text-slate-500">Tidak ada data yang cocok.</td></tr>
                )}
                {sections.map(sec => (
                  <React.Fragment key={sec.deptName}>
                    <tr className="bg-purple-500/10 border-y border-purple-500/30">
                      <td colSpan={colSpan} className="p-0">
                        <div className="sticky left-0 w-fit py-2.5 px-3">
                          <span className="text-xs font-bold text-purple-300 uppercase tracking-wide">{sec.deptName}</span>
                          <span className="ml-2 text-[11px] text-slate-400 font-mono">({sec.rows.length} orang)</span>
                        </div>
                      </td>
                    </tr>
                    {sec.rows.map((d, i) => {
                      no += 1;
                      return (
                        <tr key={`${d.m.nip}-${i}`} className="group hover:bg-slate-800/40 border-b border-slate-800">
                          <td className="sticky left-0 z-10 bg-slate-900 group-hover:bg-slate-800 py-3 px-3 w-10 min-w-[40px] max-w-[40px] text-slate-500 font-mono">{no}</td>
                          <td className="sticky left-10 z-10 bg-slate-900 group-hover:bg-slate-800 py-3 px-3 min-w-[220px] font-semibold text-white">
                            {d.m.nama}
                            {d.isKadep && (
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
                          {showRencana && rencanaKeys.map(key => (
                            <td key={key} className="py-3 px-3 text-center bg-blue-500/5 font-mono text-xs text-blue-300 whitespace-nowrap">
                              {d.rencana[key] || '-'}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </main>
    </div>
  );
}