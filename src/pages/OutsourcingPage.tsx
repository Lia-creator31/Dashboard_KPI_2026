import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  Building2, Briefcase, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, Lock, X, Check, Printer, FileCheck, Clock, Sparkles, 
  FileSpreadsheet, Users, Search, Plus, Calendar
} from 'lucide-react';

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
}

interface TimesheetLine {
  id: string;
  date: string;
  workOrder: string;
  description: string;
  effectiveHours: number;
  overtimeHours: number;
}

interface TimesheetHeader {
  id: string;
  code: string;
  startDate: string;
  endDate: string;
  responsible: string;
  unitKerja: string;
  status: string;
  lines: TimesheetLine[];
}

function cleanText(str: string): string {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

function parseToStandardDate(val: any): string {
  if (!val) return '';
  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'nan') return '';
  return str.slice(0, 10);
}

function formatDisplayDate(val: any): string {
  if (!val) return '-';
  const iso = parseToStandardDate(val);
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return `${m[3]}-${m[2]}-${m[1]}`;
  }
  return String(val) || '-';
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  const [viewMode, setViewMode] = useState<'list' | 'create'>('list');
  
  // State Form Header Timesheet
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [unitKerja, setUnitKerja] = useState('Biro Dukungan & Administrasi');
  
  // State Lines Timesheet
  const [lines, setLines] = useState<TimesheetLine[]>([]);
  
  // State Modal Add Line
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [lineDate, setLineDate] = useState('');
  const [lineWorkOrder, setLineWorkOrder] = useState('');
  const [lineDescription, setLineDescription] = useState('');
  const [lineEffective, setLineEffective] = useState('8');
  const [lineOvertime, setLineOvertime] = useState('0');

  // State daftar timesheet tersimpan
  const [savedTimesheets, setSavedTimesheets] = useState<TimesheetHeader[]>([
    {
      id: '1',
      code: 'TIM-2609301065844',
      startDate: '2026-09-28',
      endDate: '2026-09-30',
      responsible: user.nama || 'Nur Zakiyyah',
      unitKerja: 'Biro Dukungan & Administrasi',
      status: 'Approved',
      lines: []
    }
  ]);

  // Load data job cards / work orders yang diterbitkan untuk user ini dari Supabase
  const [availableWorkOrders, setAvailableWorkOrders] = useState<any[]>([]);

  useEffect(() => {
    async function fetchWorkOrders() {
      try {
        const { data, error } = await supabase.from('job_cards').select('*');
        if (data) {
          // Filter fleksibel mencocokkan nama user yang login (misal: "Hari Priyono")
          const myWo = data.filter(item => {
            const picClean = cleanText(item.pic || item.personil_name || '');
            const userClean = cleanText(user.nama || '');
            return picClean.includes(userClean) || userClean.includes(picClean);
          });
          setAvailableWorkOrders(myWo);
        }
      } catch {}
    }
    fetchWorkOrders();
  }, [user.nama]);

  // Filter work order berdasarkan tanggal harian (lineDate) yang dipilih di modal
  const filteredWorkOrdersForModal = useMemo(() => {
    if (!lineDate) return availableWorkOrders;
    
    return availableWorkOrders.filter(wo => {
      const woStart = parseToStandardDate(wo.start_date || wo.startDate);
      const woEnd = parseToStandardDate(wo.end_date || wo.endDate);
      
      if (woStart && woEnd) {
        return lineDate >= woStart && lineDate <= woEnd;
      }
      return true;
    });
  }, [availableWorkOrders, lineDate]);

  const handleAddLineSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!lineDate || !lineWorkOrder) {
      alert('Tanggal dan Work Order wajib diisi!');
      return;
    }

    const newLine: TimesheetLine = {
      id: Math.random().toString(36).substring(2, 9),
      date: lineDate,
      workOrder: lineWorkOrder,
      description: lineDescription || 'Pekerjaan Desain & Drafting',
      effectiveHours: parseFloat(lineEffective) || 8,
      overtimeHours: parseFloat(lineOvertime) || 0,
    };

    setLines(prev => [...prev, newLine]);
    setIsModalOpen(false);
    // Reset modal form
    setLineWorkOrder('');
    setLineDescription('');
    setLineEffective('8');
    setLineOvertime('0');
  };

  const handleSaveTimesheet = () => {
    if (!startDate || !endDate) {
      alert('Harap tentukan Tanggal Mulai dan Tanggal Selesai terlebih dahulu.');
      return;
    }
    if (lines.length === 0) {
      alert('Tambahkan setidaknya satu baris timesheet (Add a line).');
      return;
    }

    const randomCode = `TIM-${Math.floor(1000000000000 + Math.random() * 9000000000000)}`;
    const newTimesheet: TimesheetHeader = {
      id: Date.now().toString(),
      code: randomCode,
      startDate,
      endDate,
      responsible: user.nama,
      unitKerja,
      status: 'Approved',
      lines
    };

    setSavedTimesheets(prev => [newTimesheet, ...prev]);
    setViewMode('list');
    setLines([]);
    setStartDate('');
    setEndDate('');
    alert(`Timesheet berhasil disimpan dengan nomor: ${randomCode}`);
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-amber-600 rounded-lg text-white">
              <HardHat className="w-4 h-4" />
            </div>
            <span className="font-bold text-sm text-white">PORTAL MITRA / OUTSOURCING — <span className="text-amber-400">{user.nama}</span></span>
          </div>
          <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
            Keluar
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-6">
        
        {viewMode === 'list' ? (
          /* ================= VIEW 1: KARTU TIMESHEET ================= */
          <div className="space-y-6">
            <div className="flex items-center justify-between border-b border-slate-800 pb-4">
              <div>
                <h2 className="text-xl font-bold text-white">Project Timesheet</h2>
                <span className="text-xs text-slate-400">Dokumen rekapitulasi jam kerja dan work order harian</span>
              </div>
              <button
                onClick={() => setViewMode('create')}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-lg shadow flex items-center gap-1.5 cursor-pointer"
              >
                <Plus className="w-4 h-4" /> Create
              </button>
            </div>

            {/* Grid Kartu Timesheet */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
              {savedTimesheets.map((ts) => (
                <div key={ts.id} className="bg-slate-900 border border-slate-800 hover:border-amber-500/50 rounded-xl p-4 space-y-3 transition shadow-lg">
                  <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                    <span className="font-mono font-bold text-amber-400 text-xs tracking-wider">{ts.code}</span>
                    <span className="text-[10px] px-2 py-0.5 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded font-semibold">
                      {ts.status}
                    </span>
                  </div>
                  <div className="text-xs space-y-1.5 font-mono text-slate-300">
                    <div><span className="text-slate-500">Responsible:</span> <span className="text-white">{ts.responsible}</span></div>
                    <div className="flex justify-between"><span className="text-slate-500">Start Date:</span> <span className="text-cyan-300">{formatDisplayDate(ts.startDate)}</span></div>
                    <div className="flex justify-between"><span className="text-slate-500">End Date:</span> <span className="text-cyan-300">{formatDisplayDate(ts.endDate)}</span></div>
                  </div>
                  <div className="pt-2 border-t border-slate-800/80 text-[11px] text-slate-400 flex justify-between">
                    <span>Total Baris: {ts.lines.length} Line(s)</span>
                    <span className="text-purple-400 font-semibold">{ts.unitKerja}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          /* ================= VIEW 2: FORM CREATE TIMESHEET ================= */
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <button onClick={() => setViewMode('list')} className="p-1.5 bg-slate-800 hover:bg-slate-700 rounded-lg text-slate-300 cursor-pointer">
                  <ArrowLeft className="w-4 h-4" />
                </button>
                <h2 className="text-lg font-bold text-white">New Timesheet</h2>
              </div>
              <button
                onClick={handleSaveTimesheet}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-lg cursor-pointer"
              >
                Simpan & Terbitkan
              </button>
            </div>

            {/* Header Informasi */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs bg-slate-950 p-4 rounded-xl border border-slate-800">
              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">User</span>
                  <span className="font-semibold text-white">{user.nama}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Date Start</span>
                  <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" />
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Date End</span>
                  <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" />
                </div>
              </div>

              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Unit Kerja</span>
                  <span className="font-semibold text-cyan-400">{unitKerja}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">Divisi</span>
                  <span className="font-semibold text-white">71000 - Divisi Desain</span>
                </div>
              </div>
            </div>

            {/* Tab Project Timesheet Line */}
            <div className="space-y-3">
              <div className="border-b border-slate-800 pb-2">
                <span className="text-xs font-bold text-slate-300 uppercase tracking-wider bg-slate-800 px-3 py-1.5 rounded-t-lg">
                  Project Timesheet Line
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="text-slate-400 border-b border-slate-800 font-mono">
                      <th className="py-2.5 px-3">Date</th>
                      <th className="py-2.5 px-3">Work Order</th>
                      <th className="py-2.5 px-3">Description</th>
                      <th className="py-2.5 px-3 text-center">Effective Hours</th>
                      <th className="py-2.5 px-3 text-center">Overtime Hours</th>
                      <th className="py-2.5 px-3 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
                    {lines.map((l, idx) => (
                      <tr key={l.id} className="hover:bg-slate-950/50">
                        <td className="py-2.5 px-3 font-mono text-cyan-300">{formatDisplayDate(l.date)}</td>
                        <td className="py-2.5 px-3 font-mono font-bold text-amber-300">{l.workOrder}</td>
                        <td className="py-2.5 px-3 text-slate-200">{l.description}</td>
                        <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{l.effectiveHours} Jam</td>
                        <td className="py-2.5 px-3 text-center font-mono text-purple-400">{l.overtimeHours} Jam</td>
                        <td className="py-2.5 px-3 text-center">
                          <button onClick={() => setLines(lines.filter(item => item.id !== l.id))} className="text-rose-400 hover:text-rose-300 cursor-pointer">
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Tombol Add a Line */}
              <button
                onClick={() => {
                  if (!startDate || !endDate) {
                    alert('Mohon isi Date Start dan Date End terlebih dahulu!');
                    return;
                  }
                  setLineDate(startDate);
                  setIsModalOpen(true);
                }}
                className="text-blue-400 hover:text-blue-300 text-xs font-semibold cursor-pointer pt-2 flex items-center gap-1"
              >
                + Add a line
              </button>
            </div>
          </div>
        )}
      </main>

      {/* ================= MODAL CREATE LINES ================= */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
            <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <span className="font-bold text-xs text-white">Create Lines (Timesheet Harian)</span>
              <button onClick={() => setIsModalOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleAddLineSubmit} className="p-5 space-y-4 text-xs">
              <div className="grid grid-cols-1 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1">Date (Tanggal Harian)</label>
                  <input
                    type="date"
                    min={startDate}
                    max={endDate}
                    value={lineDate}
                    onChange={(e) => setLineDate(e.target.value)}
                    required
                    style={{ colorScheme: 'dark' }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono cursor-pointer"
                  />
                  <span className="text-[10px] text-slate-500 mt-0.5 block">Harus berada di antara rentang Start Date & End Date.</span>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Work Order (Terbitan Kabiro untuk {user.nama})</label>
                    <select
                      value={lineWorkOrder}
                      onChange={(e) => {
                        setLineWorkOrder(e.target.value);
                        const found = filteredWorkOrdersForModal.find(w => w.kode_jc === e.target.value);
                        if (found) setLineDescription(found.task_name || '');
                      }}
                      required
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    >
                      <option value="">-- Pilih Work Order --</option>
                      {filteredWorkOrdersForModal.map((wo, idx) => (
                        <option key={wo.id || idx} value={wo.kode_jc}>
                          {wo.kode_jc} - {wo.task_name || wo.project} ({formatDisplayDate(wo.start_date)} s/d {formatDisplayDate(wo.end_date)})
                        </option>
                      ))}
                    </select>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Description</label>
                  <input
                    type="text"
                    value={lineDescription}
                    onChange={(e) => setLineDescription(e.target.value)}
                    placeholder="Uraian pekerjaan harian..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                  />
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-400 mb-1">Effective Hours (Jam)</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="24"
                      value={lineEffective}
                      onChange={(e) => setLineEffective(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">Overtime Hours (Jam Lembur)</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="12"
                      value={lineOvertime}
                      onChange={(e) => setLineOvertime(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                    />
                  </div>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                <button type="button" onClick={() => setIsModalOpen(false)} className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold rounded-lg cursor-pointer">
                  Discard
                </button>
                <button type="submit" className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg cursor-pointer">
                  Save & Close
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}