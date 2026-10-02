import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  HardHat, ArrowLeft, X, Check, Plus, Trash2, Pencil,
  Calendar, FileCheck, Clock, Printer
} from 'lucide-react';

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
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
  release?: string;
}

interface TimesheetLine {
  id: string;
  date: string;
  workOrderId: string;
  workOrderCode: string;
  description: string;
  effectiveHours: number;
  overtimeHours: number;
  day: string;
}

interface TimesheetHeader {
  id: string;
  code: string;
  userId: string;
  userName: string;
  dateStart: string;
  dateEnd: string;
  unitKerja: string;
  divisi: string;
  status: 'Draft' | 'Submitted' | 'Approved';
  lines: TimesheetLine[];
  createdAt: string;
}

// Helper: Normalisasi tanggal ke format YYYY-MM-DD
function normalizeDate(dateStr: string): string {
  if (!dateStr) return '';
  const clean = String(dateStr).trim();
  // Sudah YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) return clean;
  // Format DD-MM-YYYY
  if (/^\d{2}-\d{2}-\d{4}$/.test(clean)) {
    const parts = clean.split('-');
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  // Format DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(clean)) {
    const parts = clean.split('/');
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return clean;
}

// Helper: Format tanggal untuk tampilan DD-MM-YYYY
function formatDisplayDate(dateStr: string): string {
  const normalized = normalizeDate(dateStr);
  if (!normalized) return '-';
  const parts = normalized.split('-');
  if (parts.length === 3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
  return dateStr;
}

// Helper: Dapatkan nama hari dari tanggal
function getDayName(dateStr: string): string {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const normalized = normalizeDate(dateStr);
  if (!normalized) return '';
  const date = new Date(normalized);
  return days[date.getDay()];
}

// Helper: Generate kode timesheet unik
function generateTimesheetCode(): string {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const rand = Math.floor(Math.random() * 900000) + 100000;
  return `TIM-${yy}${mm}${dd}${rand}`;
}

// Helper: Cek apakah nama cocok (flexible matching)
function isNameMatch(dbName: string, userName: string): boolean {
  if (!dbName || !userName) return false;
  const db = dbName.toLowerCase().trim();
  const user = userName.toLowerCase().trim();
  
  // Exact match
  if (db === user) return true;
  
  // User format: "NIP - Nama", DB hanya "Nama"
  if (user.includes(' - ')) {
    const userNamaOnly = user.split(' - ')[1].trim();
    if (db === userNamaOnly) return true;
    if (db.includes(userNamaOnly) || userNamaOnly.includes(db)) return true;
  }
  
  // Partial match
  if (db.includes(user) || user.includes(db)) return true;
  
  return false;
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  const [viewMode, setViewMode] = useState<'list' | 'create'>('list');
  
  // State untuk daftar timesheet
  const [timesheets, setTimesheets] = useState<TimesheetHeader[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  
  // State form header timesheet
  const [dateStart, setDateStart] = useState('');
  const [dateEnd, setDateEnd] = useState('');
  const [unitKerja, setUnitKerja] = useState('Biro Dukungan & Administrasi');
  const [divisi, setDivisi] = useState('71000 - Divisi Desain');
  
  // State lines timesheet
  const [lines, setLines] = useState<TimesheetLine[]>([]);
  
  // State modal Add Line
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [lineDate, setLineDate] = useState('');
  const [lineWorkOrderId, setLineWorkOrderId] = useState('');
  const [lineDescription, setLineDescription] = useState('');
  const [lineEffective, setLineEffective] = useState('8');
  const [lineOvertime, setLineOvertime] = useState('0');
  const [editingLineId, setEditingLineId] = useState<string | null>(null);
  
  // State work orders yang tersedia
  const [availableWorkOrders, setAvailableWorkOrders] = useState<TaskItem[]>([]);

  // ============ LOAD TIMESHEET DARI SUPABASE ============
  const loadTimesheets = useCallback(async () => {
    setIsLoading(true);
    try {
      const { data, error } = await supabase
        .from('outsourcing_timesheets')
        .select('*')
        .eq('user_id', user.nip)
        .order('created_at', { ascending: false });

      if (error) {
        console.error('Error loading timesheets:', error);
        setTimesheets([]);
        return;
      }

      // Load lines untuk setiap timesheet
      const timesheetsWithLines: TimesheetHeader[] = [];
      for (const ts of (data || [])) {
        const { data: linesData } = await supabase
          .from('outsourcing_timesheet_lines')
          .select('*')
          .eq('timesheet_id', ts.id)
          .order('date', { ascending: true });

        timesheetsWithLines.push({
          id: ts.id,
          code: ts.code || `TIM-${ts.id}`,
          userId: ts.user_id,
          userName: ts.user_name || user.nama,
          dateStart: ts.date_start,
          dateEnd: ts.date_end,
          unitKerja: ts.unit_kerja || 'Biro Dukungan & Administrasi',
          divisi: ts.divisi || '71000 - Divisi Desain',
          status: ts.status || 'Draft',
          lines: (linesData || []).map((l: any) => ({
            id: l.id,
            date: l.date,
            workOrderId: l.work_order_id,
            workOrderCode: l.work_order_code || '',
            description: l.description || '',
            effectiveHours: Number(l.effective_hours) || 0,
            overtimeHours: Number(l.overtime_hours) || 0,
            day: l.day || getDayName(l.date),
          })),
          createdAt: ts.created_at,
        });
      }

      setTimesheets(timesheetsWithLines);
    } catch (err) {
      console.error('Error:', err);
    } finally {
      setIsLoading(false);
    }
  }, [user.nip, user.nama]);

  // ============ LOAD WORK ORDER YANG TERSEDIA ============
  const loadAvailableWorkOrders = useCallback(async (selectedDate?: string) => {
    try {
      console.log('🔍 Mencari work order untuk:', user.nama, '| NIP:', user.nip);
      
      const { data, error } = await supabase
        .from('job_cards')
        .select('*')
        .eq('status', 'approved');

      if (error) {
        console.error('❌ Error Supabase:', error);
        return;
      }

      console.log(`✅ Ditemukan ${data?.length || 0} work order approved`);

      // Filter di client-side
      const filteredData = (data || []).filter((wo: any) => {
        // A. Cek Nama PIC (flexible matching)
        const isNameMatch = isNameMatch(wo.pic, user.nama) || 
                            isNameMatch(wo.pic, user.nip) ||
                            isNameMatch(wo.personil_name, user.nama);

        // B. Cek Tanggal (jika ada selectedDate)
        let isDateMatch = true;
        if (selectedDate) {
          const woStart = normalizeDate(wo.start_date);
          const woEnd = normalizeDate(wo.end_date);
          const selected = normalizeDate(selectedDate);
          
          if (woStart && woEnd && selected) {
            isDateMatch = selected >= woStart && selected <= woEnd;
          }
        }

        if (isNameMatch) {
          console.log(`📝 WO ${wo.kode_jc || wo.kodeJc}: ${wo.pic} | Range: ${wo.start_date} - ${wo.end_date} | Match: ${isDateMatch}`);
        }

        return isNameMatch && isDateMatch;
      });

      console.log('🎯 Work order yang lolos filter:', filteredData.length, filteredData);
      setAvailableWorkOrders(filteredData);
      
    } catch (error) {
      console.error('❌ Gagal memuat work order:', error);
    }
  }, [user.nama, user.nip]);

  useEffect(() => {
    loadTimesheets();
  }, [loadTimesheets]);

  // ============ HANDLER CREATE TIMESHEET ============
  const handleCreateTimesheet = () => {
    setViewMode('create');
    setDateStart('');
    setDateEnd('');
    setLines([]);
    setUnitKerja('Biro Dukungan & Administrasi');
    setDivisi('71000 - Divisi Desain');
  };

  // ============ HANDLER ADD LINE ============
  const handleAddLine = () => {
    if (!dateStart || !dateEnd) {
      alert('Mohon isi Date Start dan Date End terlebih dahulu!');
      return;
    }
    setEditingLineId(null);
    setLineDate(dateStart);
    setLineWorkOrderId('');
    setLineDescription('');
    setLineEffective('8');
    setLineOvertime('0');
    setIsModalOpen(true);
    // Load work orders untuk tanggal default
    loadAvailableWorkOrders(dateStart);
  };

  // ============ HANDLER SAAT TANGGAL DI MODAL BERUBAH ============
  const handleLineDateChange = (newDate: string) => {
    setLineDate(newDate);
    if (newDate) {
      loadAvailableWorkOrders(newDate);
    }
  };

  // ============ HANDLER SAVE LINE ============
  const handleSaveLine = () => {
    if (!lineDate) {
      alert('Tanggal wajib diisi!');
      return;
    }
    if (!lineWorkOrderId) {
      alert('Work Order wajib dipilih!');
      return;
    }

    const selectedWo = availableWorkOrders.find(wo => wo.id === lineWorkOrderId);
    if (!selectedWo) {
      alert('Work Order tidak valid!');
      return;
    }

    const newLine: TimesheetLine = {
      id: editingLineId || `line_${Date.now()}`,
      date: lineDate,
      workOrderId: lineWorkOrderId,
      workOrderCode: selectedWo.kode_jc || selectedWo.kodeJc || '',
      description: lineDescription || selectedWo.taskName || '',
      effectiveHours: parseFloat(lineEffective) || 0,
      overtimeHours: parseFloat(lineOvertime) || 0,
      day: getDayName(lineDate),
    };

    if (editingLineId) {
      setLines(prev => prev.map(l => l.id === editingLineId ? newLine : l));
    } else {
      setLines(prev => [...prev, newLine]);
    }

    setIsModalOpen(false);
    setEditingLineId(null);
  };

  // ============ HANDLER EDIT LINE ============
  const handleEditLine = (line: TimesheetLine) => {
    setEditingLineId(line.id);
    setLineDate(line.date);
    setLineWorkOrderId(line.workOrderId);
    setLineDescription(line.description);
    setLineEffective(String(line.effectiveHours));
    setLineOvertime(String(line.overtimeHours));
    setIsModalOpen(true);
    loadAvailableWorkOrders(line.date);
  };

  // ============ HANDLER DELETE LINE ============
  const handleDeleteLine = (lineId: string) => {
    if (confirm('Hapus baris ini?')) {
      setLines(prev => prev.filter(l => l.id !== lineId));
    }
  };

  // ============ HANDLER SAVE TIMESHEET ============
  const handleSaveTimesheet = async () => {
    if (!dateStart || !dateEnd) {
      alert('Tanggal mulai dan tanggal akhir harus diisi!');
      return;
    }
    if (lines.length === 0) {
      alert('Tambahkan setidaknya satu baris timesheet!');
      return;
    }

    try {
      const code = generateTimesheetCode();
      
      // Insert header timesheet
      const { data: tsData, error: tsError } = await supabase
        .from('outsourcing_timesheets')
        .insert({
          code: code,
          user_id: user.nip,
          user_name: user.nama,
          date_start: dateStart,
          date_end: dateEnd,
          unit_kerja: unitKerja,
          divisi: divisi,
          status: 'Draft',
        })
        .select()
        .single();

      if (tsError) throw tsError;

      // Insert lines
      const linesToInsert = lines.map(l => ({
        timesheet_id: tsData.id,
        date: l.date,
        work_order_id: l.workOrderId,
        work_order_code: l.workOrderCode,
        description: l.description,
        effective_hours: l.effectiveHours,
        overtime_hours: l.overtimeHours,
        day: l.day,
      }));

      const { error: linesError } = await supabase
        .from('outsourcing_timesheet_lines')
        .insert(linesToInsert);

      if (linesError) throw linesError;

      alert(`Timesheet berhasil disimpan dengan nomor: ${code}`);
      setViewMode('list');
      loadTimesheets();
    } catch (error: any) {
      console.error('Error saving timesheet:', error);
      alert('Gagal menyimpan timesheet: ' + error.message);
    }
  };

  // ============ HANDLER SUBMIT TIMESHEET ============
  const handleSubmitTimesheet = async (tsId: string) => {
    if (!confirm('Submit timesheet ini untuk approval?')) return;
    
    try {
      const { error } = await supabase
        .from('outsourcing_timesheets')
        .update({ status: 'Submitted' })
        .eq('id', tsId);

      if (error) throw error;
      alert('Timesheet berhasil disubmit!');
      loadTimesheets();
    } catch (error: any) {
      alert('Gagal submit: ' + error.message);
    }
  };

  // ============ HITUNG TOTAL JAM ============
  const totalHours = useMemo(() => {
    return lines.reduce((acc, line) => acc + line.effectiveHours + line.overtimeHours, 0);
  }, [lines]);

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-amber-600 rounded-lg text-white">
              <HardHat className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-sm text-white">PORTAL MITRA / OUTSOURCING</span>
              <span className="text-xs text-slate-400 ml-2">— {user.nama}</span>
            </div>
          </div>
          <button 
            onClick={onLogout} 
            className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer"
          >
            Keluar
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-6">
        {/* ================= VIEW 1: DAFTAR TIMESHEET ================= */}
        {viewMode === 'list' && (
          <div className="space-y-6">
            <div className="flex items-center justify-between border-b border-slate-800 pb-4">
              <div>
                <h2 className="text-xl font-bold text-white">Project Timesheet</h2>
                <p className="text-xs text-slate-400">Dokumen rekapitulasi jam kerja dan work order harian</p>
              </div>
              <button
                onClick={handleCreateTimesheet}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-lg shadow flex items-center gap-1.5 cursor-pointer"
              >
                <Plus className="w-4 h-4" /> Create
              </button>
            </div>

            {isLoading ? (
              <div className="text-center py-12 text-slate-400">Memuat data...</div>
            ) : timesheets.length === 0 ? (
              <div className="text-center py-12 bg-slate-900 border border-slate-800 rounded-xl">
                <FileCheck className="w-12 h-12 text-slate-600 mx-auto mb-3" />
                <p className="text-slate-400 text-sm">Belum ada timesheet. Klik <b>Create</b> untuk membuat yang baru.</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
                {timesheets.map((ts) => (
                  <div 
                    key={ts.id} 
                    className="bg-slate-900 border border-slate-800 hover:border-amber-500/50 rounded-xl p-4 space-y-3 transition shadow-lg"
                  >
                    <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                      <span className="font-mono font-bold text-amber-400 text-xs tracking-wider">{ts.code}</span>
                      <span className={`text-[10px] px-2 py-0.5 rounded font-semibold ${
                        ts.status === 'Approved' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' :
                        ts.status === 'Submitted' ? 'bg-blue-500/10 text-blue-400 border border-blue-500/20' :
                        'bg-slate-700 text-slate-300'
                      }`}>
                        {ts.status}
                      </span>
                    </div>
                    <div className="text-xs space-y-1.5 font-mono text-slate-300">
                      <div><span className="text-slate-500">Responsible:</span> <span className="text-white">{ts.userName}</span></div>
                      <div className="flex justify-between"><span className="text-slate-500">Start Date:</span> <span className="text-cyan-300">{formatDisplayDate(ts.dateStart)}</span></div>
                      <div className="flex justify-between"><span className="text-slate-500">End Date:</span> <span className="text-cyan-300">{formatDisplayDate(ts.dateEnd)}</span></div>
                    </div>
                    <div className="pt-2 border-t border-slate-800/80 text-[11px] text-slate-400 flex justify-between">
                      <span>Total Baris: <b className="text-white">{ts.lines.length}</b> Line(s)</span>
                      <span className="text-purple-400 font-semibold">{ts.lines.reduce((a, l) => a + l.effectiveHours + l.overtimeHours, 0).toFixed(1)} Jam</span>
                    </div>
                    {ts.status === 'Draft' && (
                      <button
                        onClick={() => handleSubmitTimesheet(ts.id)}
                        className="w-full px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs font-semibold cursor-pointer"
                      >
                        Submit for Approval
                      </button>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ================= VIEW 2: FORM CREATE TIMESHEET ================= */}
        {viewMode === 'create' && (
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => setViewMode('list')} 
                  className="p-1.5 bg-slate-800 hover:bg-slate-700 rounded-lg text-slate-300 cursor-pointer"
                >
                  <ArrowLeft className="w-4 h-4" />
                </button>
                <h2 className="text-lg font-bold text-white">New Timesheet</h2>
              </div>
              <button
                onClick={handleSaveTimesheet}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-lg cursor-pointer flex items-center gap-1.5"
              >
                <Check className="w-4 h-4" /> Simpan & Terbitkan
              </button>
            </div>

            {/* Header Informasi */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs bg-slate-950 p-4 rounded-xl border border-slate-800">
              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-slate-400">User</span>
                  <span className="font-semibold text-white">{user.nip} - {user.nama}</span>
                </div>
                <div className="flex justify-between items-center gap-2">
                  <span className="text-slate-400">Date Start</span>
                  <input 
                    type="date" 
                    value={dateStart} 
                    onChange={(e) => setDateStart(e.target.value)} 
                    style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" 
                  />
                </div>
                <div className="flex justify-between items-center gap-2">
                  <span className="text-slate-400">Date End</span>
                  <input 
                    type="date" 
                    value={dateEnd} 
                    onChange={(e) => setDateEnd(e.target.value)} 
                    style={{ colorScheme: 'dark' }}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono cursor-pointer" 
                  />
                </div>
              </div>

              <div className="space-y-3">
                <div className="flex justify-between items-center gap-2">
                  <span className="text-slate-400">Unit Kerja</span>
                  <input 
                    type="text" 
                    value={unitKerja} 
                    onChange={(e) => setUnitKerja(e.target.value)}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-cyan-400 font-semibold" 
                  />
                </div>
                <div className="flex justify-between items-center gap-2">
                  <span className="text-slate-400">Divisi</span>
                  <input 
                    type="text" 
                    value={divisi} 
                    onChange={(e) => setDivisi(e.target.value)}
                    className="px-3 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-semibold" 
                  />
                </div>
              </div>
            </div>

            {/* Project Timesheet Line */}
            <div className="space-y-3">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="text-xs font-bold text-slate-300 uppercase tracking-wider bg-slate-800 px-3 py-1.5 rounded-t-lg">
                  Project Timesheet Line
                </span>
                <button
                  onClick={handleAddLine}
                  className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs cursor-pointer flex items-center gap-1"
                >
                  <Plus className="w-3 h-3" /> Add a line
                </button>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="text-slate-400 border-b border-slate-800 font-mono">
                      <th className="py-2.5 px-3">Date</th>
                      <th className="py-2.5 px-3">Day</th>
                      <th className="py-2.5 px-3">Work Order</th>
                      <th className="py-2.5 px-3">Description</th>
                      <th className="py-2.5 px-3 text-center">Effective Hours</th>
                      <th className="py-2.5 px-3 text-center">Overtime Hours</th>
                      <th className="py-2.5 px-3 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
                    {lines.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="px-3 py-8 text-center text-slate-500 italic">
                          Belum ada baris. Klik "Add a line" untuk menambahkan.
                        </td>
                      </tr>
                    ) : (
                      lines.map((l) => (
                        <tr key={l.id} className="hover:bg-slate-950/50">
                          <td className="py-2.5 px-3 font-mono text-cyan-300">{formatDisplayDate(l.date)}</td>
                          <td className="py-2.5 px-3 text-slate-400">{l.day}</td>
                          <td className="py-2.5 px-3 font-mono font-bold text-amber-300">{l.workOrderCode}</td>
                          <td className="py-2.5 px-3 text-slate-200">{l.description}</td>
                          <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{l.effectiveHours} Jam</td>
                          <td className="py-2.5 px-3 text-center font-mono text-purple-400">{l.overtimeHours} Jam</td>
                          <td className="py-2.5 px-3 text-center">
                            <div className="flex justify-center gap-1">
                              <button 
                                onClick={() => handleEditLine(l)} 
                                className="p-1 text-blue-400 hover:bg-blue-500/20 rounded cursor-pointer"
                                title="Edit"
                              >
                                <Pencil className="w-3 h-3" />
                              </button>
                              <button 
                                onClick={() => handleDeleteLine(l.id)} 
                                className="p-1 text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer"
                                title="Hapus"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                  {lines.length > 0 && (
                    <tfoot className="border-t-2 border-slate-700">
                      <tr className="bg-slate-950/50">
                        <td colSpan={4} className="py-2.5 px-3 text-right font-bold text-white">Total Hours:</td>
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-emerald-400">{totalHours.toFixed(1)} Jam</td>
                        <td colSpan={2}></td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* ================= MODAL CREATE/EDIT LINE ================= */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
            <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <span className="font-bold text-xs text-white flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-amber-400" /> 
                {editingLineId ? 'Edit Line' : 'Create Lines (Timesheet Harian)'}
              </span>
              <button 
                onClick={() => { setIsModalOpen(false); setEditingLineId(null); }} 
                className="text-slate-400 hover:text-white cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-4 text-xs">
              <div className="grid grid-cols-1 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1">Date (Tanggal Harian) <span className="text-rose-400">*</span></label>
                  <input
                    type="date"
                    min={dateStart}
                    max={dateEnd}
                    value={lineDate}
                    onChange={(e) => handleLineDateChange(e.target.value)}
                    style={{ colorScheme: 'dark' }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono cursor-pointer"
                  />
                  <span className="text-[10px] text-slate-500 mt-0.5 block">
                    Harus berada di antara rentang {formatDisplayDate(dateStart)} s/d {formatDisplayDate(dateEnd)}.
                  </span>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Work Order (Terbitan Kabiro) <span className="text-rose-400">*</span></label>
                  <select
                    value={lineWorkOrderId}
                    onChange={(e) => {
                      setLineWorkOrderId(e.target.value);
                      const found = availableWorkOrders.find(w => w.id === e.target.value);
                      if (found) setLineDescription(found.taskName || '');
                    }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                  >
                    <option value="">-- Pilih Work Order --</option>
                    {availableWorkOrders.length === 0 && (
                      <option value="" disabled>Tidak ada work order untuk tanggal ini</option>
                    )}
                    {availableWorkOrders.map((wo) => (
                      <option key={wo.id} value={wo.id}>
                        {wo.kode_jc || wo.kodeJc || '-'} - {wo.taskName || wo.task_name || wo.project} ({formatDisplayDate(wo.startDate)} s/d {formatDisplayDate(wo.endDate)})
                      </option>
                    ))}
                  </select>
                  {availableWorkOrders.length === 0 && lineDate && (
                    <span className="text-[10px] text-amber-400 mt-0.5 block">
                      ⚠️ Tidak ada work order yang tersedia untuk tanggal {formatDisplayDate(lineDate)}. Coba pilih tanggal lain.
                    </span>
                  )}
                </div>

                <div>
                  <label className="block text-slate-400 mb-1">Description</label>
                  <textarea
                    value={lineDescription}
                    onChange={(e) => setLineDescription(e.target.value)}
                    placeholder="Uraian pekerjaan harian..."
                    rows={2}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white resize-none"
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

                {lineDate && (
                  <div className="bg-slate-950 p-2 rounded border border-slate-800 text-center">
                    <span className="text-slate-400">Day: </span>
                    <span className="text-cyan-400 font-semibold">{getDayName(lineDate)}</span>
                  </div>
                )}
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                <button 
                  type="button" 
                  onClick={() => { setIsModalOpen(false); setEditingLineId(null); }} 
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold rounded-lg cursor-pointer"
                >
                  Discard
                </button>
                <button 
                  type="button"
                  onClick={handleSaveLine} 
                  className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg cursor-pointer"
                >
                  Save & Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}