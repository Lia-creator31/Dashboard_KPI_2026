import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  HardHat, ArrowLeft, X, Check, Plus, Trash2, Pencil,
  Calendar, FileCheck, Printer, ChevronLeft
} from 'lucide-react';

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
}

interface TimesheetLine {
  id: string;
  date: string;
  workOrderId: string;
  workOrderCode: string;
  description: string;
  effectiveHours: number;
  overtimeHours: number;
  plannedHour?: string;
  idleHours?: string;
  type?: string;
  project?: string;
  workCenter?: string;
  wbsActivity?: string;
  day?: string;
  analyticAccount?: string;
}

interface TimesheetHeader {
  id: string;
  code: string;
  userId: string;
  userName: string;
  nip: string;
  dateStart: string;
  dateEnd: string;
  unitKerja: string;
  divisi: string;
  status: 'Draft' | 'Submitted' | 'Approved';
  lines: TimesheetLine[];
  createdAt: string;
}

// Helper: Normalisasi tanggal
function normalizeDate(dateStr: string): string {
  if (!dateStr) return '';
  const clean = String(dateStr).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) return clean;
  if (/^\d{2}-\d{2}-\d{4}$/.test(clean)) {
    const parts = clean.split('-');
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return clean;
}

function formatDisplayDate(dateStr: string): string {
  const normalized = normalizeDate(dateStr);
  if (!normalized) return '-';
  const parts = normalized.split('-');
  if (parts.length === 3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
  return dateStr;
}

function getDayName(dateStr: string): string {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const normalized = normalizeDate(dateStr);
  if (!normalized) return '';
  const date = new Date(normalized);
  return days[date.getDay()];
}

function generateTimesheetCode(): string {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const rand = Math.floor(Math.random() * 900000) + 100000;
  return `TIM-${yy}${mm}${dd}${rand}`;
}

// Helper: ambil deskripsi dari work order (job_cards).
// Jika nama kolom di tabel Anda berbeda, tambahkan di sini.
function getWoDescription(wo: any): string {
  return (
    wo?.deskripsi ||
    wo?.description ||
    wo?.uraian ||
    wo?.task_name ||
    wo?.taskName ||
    ''
  );
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  const [viewMode, setViewMode] = useState<'list' | 'create' | 'detail'>('list');
  const [selectedTimesheet, setSelectedTimesheet] = useState<TimesheetHeader | null>(null);

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

  // State work orders
  const [availableWorkOrders, setAvailableWorkOrders] = useState<any[]>([]);

  // Load timesheet dari Supabase
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
          nip: user.nip,
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
            plannedHour: l.planned_hour || '08:00',
            idleHours: l.idle_hours || '00:00',
            type: l.type || 'Biasa',
            project: l.project || '',
            workCenter: l.work_center || '',
            wbsActivity: l.wbs_activity || '',
            day: l.day || getDayName(l.date),
            analyticAccount: l.analytic_account || '',
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

  // Load work orders
  const loadAvailableWorkOrders = useCallback(async (selectedDate?: string) => {
    try {
      const { data, error } = await supabase
        .from('job_cards')
        .select('*')
        .eq('status', 'approved');

      if (error) {
        console.error('Error Supabase:', error);
        return;
      }

      const filteredData = (data || []).filter((wo: any) => {
        const isNameMatch = wo.pic?.toLowerCase().includes(user.nama.toLowerCase()) ||
                            user.nama.toLowerCase().includes(wo.pic?.toLowerCase()) ||
                            wo.personil_name?.toLowerCase().includes(user.nama.toLowerCase());

        let isDateMatch = true;
        if (selectedDate) {
          const woStart = normalizeDate(wo.start_date || wo.startDate);
          const woEnd = normalizeDate(wo.end_date || wo.endDate);
          const selected = normalizeDate(selectedDate);

          if (woStart && woEnd && selected) {
            isDateMatch = selected >= woStart && selected <= woEnd;
          }
        }

        return isNameMatch && isDateMatch;
      });

      setAvailableWorkOrders(filteredData);

    } catch (error) {
      console.error('Gagal memuat work order:', error);
    }
  }, [user.nama]);

  useEffect(() => {
    loadTimesheets();
  }, [loadTimesheets]);

  // Handler untuk melihat detail timesheet
  const handleViewDetail = (timesheet: TimesheetHeader) => {
    setSelectedTimesheet(timesheet);
    setViewMode('detail');
  };

  // Handler create timesheet
  const handleCreateTimesheet = () => {
    setViewMode('create');
    setDateStart('');
    setDateEnd('');
    setLines([]);
    setUnitKerja('Biro Dukungan & Administrasi');
    setDivisi('71000 - Divisi Desain');
  };

  // Handler add line
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
    loadAvailableWorkOrders(dateStart);
  };

  // Handler saat tanggal di modal berubah
  const handleLineDateChange = (newDate: string) => {
    setLineDate(newDate);
    setLineWorkOrderId('');
    setLineDescription('');
    if (newDate) {
      loadAvailableWorkOrders(newDate);
    }
  };

  // Handler save line
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
      description: lineDescription.trim() || getWoDescription(selectedWo),
      effectiveHours: parseFloat(lineEffective) || 0,
      overtimeHours: parseFloat(lineOvertime) || 0,
      plannedHour: '08:00',
      idleHours: '00:00',
      type: 'Biasa',
      project: selectedWo.project || '',
      workCenter: 'ICD303 - Biro Dukungan & Administrasi',
      wbsActivity: 'Level 6 ICAJO Jam Orang Divisi Desain',
      day: getDayName(lineDate),
      analyticAccount: `CORP2026\nCORPORATE BUDGET 2026`,
    };

    if (editingLineId) {
      setLines(prev => prev.map(l => l.id === editingLineId ? newLine : l));
    } else {
      setLines(prev => [...prev, newLine]);
    }

    setIsModalOpen(false);
    setEditingLineId(null);
  };

  // Handler edit line
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

  // Handler delete line
  const handleDeleteLine = (lineId: string) => {
    if (confirm('Hapus baris ini?')) {
      setLines(prev => prev.filter(l => l.id !== lineId));
    }
  };

  // Handler save timesheet
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

      const linesToInsert = lines.map(l => ({
        timesheet_id: tsData.id,
        date: l.date,
        work_order_id: l.workOrderId,
        work_order_code: l.workOrderCode,
        description: l.description,
        effective_hours: l.effectiveHours,
        overtime_hours: l.overtimeHours,
        planned_hour: l.plannedHour || '08:00',
        idle_hours: l.idleHours || '00:00',
        type: l.type || 'Biasa',
        project: l.project || '',
        work_center: l.workCenter || '',
        wbs_activity: l.wbsActivity || '',
        day: l.day || getDayName(l.date),
        analytic_account: l.analyticAccount || '',
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

  // Handler submit timesheet
  const handleSubmitTimesheet = async (tsId: string) => {
    if (!confirm('Submit timesheet ini untuk approval?')) return;

    try {
      const { error } = await supabase
        .from('outsourcing_timesheets')
        .update({ status: 'Submitted' })
        .eq('id', tsId);

      if (error) throw error;
      alert('Timesheet berhasil disubmit!');
      setViewMode('list');
      setSelectedTimesheet(null);
      loadTimesheets();
    } catch (error: any) {
      alert('Gagal submit: ' + error.message);
    }
  };

  // Hitung total jam
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
                    onClick={() => handleViewDetail(ts)}
                    className="bg-slate-900 border border-slate-800 hover:border-amber-500/50 rounded-xl p-4 space-y-3 transition shadow-lg cursor-pointer"
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
                        onClick={(e) => {
                          e.stopPropagation();
                          handleSubmitTimesheet(ts.id);
                        }}
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

        {/* ================= VIEW 2: DETAIL TIMESHEET ================= */}
        {viewMode === 'detail' && selectedTimesheet && (
          <div className="space-y-6">
            {/* Breadcrumb & Actions */}
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => { setViewMode('list'); setSelectedTimesheet(null); }}
                  className="p-1.5 bg-slate-800 hover:bg-slate-700 rounded-lg text-slate-300 cursor-pointer"
                >
                  <ArrowLeft className="w-4 h-4" />
                </button>
                <div className="flex items-center gap-2 text-sm">
                  <span className="text-slate-400">Project Timesheet</span>
                  <ChevronLeft className="w-3 h-3 text-slate-600 rotate-180" />
                  <span className="font-mono font-bold text-amber-400">{selectedTimesheet.code}</span>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => window.print()}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs cursor-pointer flex items-center gap-1"
                >
                  <Printer className="w-3 h-3" /> Print
                </button>
                {selectedTimesheet.status === 'Draft' && (
                  <button
                    onClick={() => handleSubmitTimesheet(selectedTimesheet.id)}
                    className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs cursor-pointer"
                  >
                    Submit
                  </button>
                )}
              </div>
            </div>

            {/* Main Card */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-2xl">
              {/* Header Info */}
              <div className="mb-6">
                <h1 className="text-2xl font-bold text-white mb-4">{selectedTimesheet.code}</h1>
                <div className="grid grid-cols-2 gap-6 text-xs">
                  <div className="space-y-2">
                    <div className="flex">
                      <span className="text-slate-400 w-32">User</span>
                      <span className="text-white font-mono">{selectedTimesheet.nip} - {selectedTimesheet.userName}</span>
                    </div>
                    <div className="flex">
                      <span className="text-slate-400 w-32">Date Start</span>
                      <span className="text-white font-mono">{formatDisplayDate(selectedTimesheet.dateStart)}</span>
                    </div>
                    <div className="flex">
                      <span className="text-slate-400 w-32">Date End</span>
                      <span className="text-white font-mono">{formatDisplayDate(selectedTimesheet.dateEnd)}</span>
                    </div>
                  </div>
                  <div className="space-y-2">
                    <div className="flex">
                      <span className="text-slate-400 w-32">Unit Kerja</span>
                      <span className="text-cyan-400">{selectedTimesheet.unitKerja}</span>
                    </div>
                    <div className="flex">
                      <span className="text-slate-400 w-32">Divisi</span>
                      <span className="text-white">{selectedTimesheet.divisi}</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Status Badge */}
              <div className="mb-4 flex justify-end">
                <span className={`px-3 py-1 rounded text-xs font-semibold ${
                  selectedTimesheet.status === 'Approved' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' :
                  selectedTimesheet.status === 'Submitted' ? 'bg-blue-500/10 text-blue-400 border border-blue-500/20' :
                  'bg-slate-700 text-slate-300'
                }`}>
                  {selectedTimesheet.status}
                </span>
              </div>

              {/* Table Header */}
              <div className="mb-2">
                <span className="text-xs font-bold text-slate-300 uppercase tracking-wider bg-slate-800 px-3 py-1.5 rounded-t-lg inline-block">
                  Project Timesheet Line
                </span>
              </div>

              {/* Detailed Table (sama dengan kolom di form) */}
              <div className="overflow-x-auto border border-slate-800 rounded-lg">
                <table className="w-full text-xs">
                  <thead className="bg-slate-950 text-slate-400">
                    <tr>
                      <th className="py-2.5 px-3 text-left border-b border-slate-800">Date</th>
                      <th className="py-2.5 px-3 text-left border-b border-slate-800">Work Order</th>
                      <th className="py-2.5 px-3 text-left border-b border-slate-800">Description</th>
                      <th className="py-2.5 px-3 text-center border-b border-slate-800">Effective Hours</th>
                      <th className="py-2.5 px-3 text-center border-b border-slate-800">Overtime Hours</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300">
                    {selectedTimesheet.lines.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-3 py-8 text-center text-slate-500 italic">
                          Tidak ada baris timesheet
                        </td>
                      </tr>
                    ) : (
                      selectedTimesheet.lines.map((line) => (
                        <tr key={line.id} className="hover:bg-slate-950/50">
                          <td className="py-2.5 px-3 font-mono text-cyan-300">{formatDisplayDate(line.date)}</td>
                          <td className="py-2.5 px-3 font-mono font-bold text-amber-300">{line.workOrderCode}</td>
                          <td className="py-2.5 px-3 text-slate-200">{line.description || '-'}</td>
                          <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{line.effectiveHours.toFixed(2)}</td>
                          <td className="py-2.5 px-3 text-center font-mono text-purple-400">{line.overtimeHours.toFixed(2)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                  {selectedTimesheet.lines.length > 0 && (
                    <tfoot className="bg-slate-950 border-t-2 border-slate-700">
                      <tr>
                        <td colSpan={3} className="py-2.5 px-3 text-right font-bold text-white">Total:</td>
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-emerald-400">
                          {selectedTimesheet.lines.reduce((a, l) => a + l.effectiveHours, 0).toFixed(2)}
                        </td>
                        <td className="py-2.5 px-3 text-center font-mono font-bold text-purple-400">
                          {selectedTimesheet.lines.reduce((a, l) => a + l.overtimeHours, 0).toFixed(2)}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ================= VIEW 3: FORM CREATE TIMESHEET ================= */}
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
                        <td colSpan={6} className="px-3 py-8 text-center text-slate-500 italic">
                          Belum ada baris. Klik "Add a line" untuk menambahkan.
                        </td>
                      </tr>
                    ) : (
                      lines.map((l) => (
                        <tr key={l.id} className="hover:bg-slate-950/50">
                          <td className="py-2.5 px-3 font-mono text-cyan-300">{formatDisplayDate(l.date)}</td>
                          <td className="py-2.5 px-3 font-mono font-bold text-amber-300">{l.workOrderCode}</td>
                          <td className="py-2.5 px-3 text-slate-200">{l.description || '-'}</td>
                          <td className="py-2.5 px-3 text-center font-mono text-emerald-400">{l.effectiveHours} Jam</td>
                          <td className="py-2.5 px-3 text-center font-mono text-purple-400">{l.overtimeHours} Jam</td>
                          <td className="py-2.5 px-3 text-center">
                            <div className="flex justify-center gap-1">
                              <button
                                onClick={() => handleEditLine(l)}
                                className="p-1 text-blue-400 hover:bg-blue-500/20 rounded cursor-pointer"
                              >
                                <Pencil className="w-3 h-3" />
                              </button>
                              <button
                                onClick={() => handleDeleteLine(l.id)}
                                className="p-1 text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer"
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
                        <td colSpan={3} className="py-2.5 px-3 text-right font-bold text-white">Total Hours:</td>
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
                      // Deskripsi work order disalin ke form timesheet, tetap bisa diedit
                      setLineDescription(found ? getWoDescription(found) : '');
                    }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono"
                  >
                    <option value="">-- Pilih Work Order --</option>
                    {availableWorkOrders.map((wo) => (
                      <option key={wo.id} value={wo.id}>
                        {wo.kode_jc || wo.kodeJc || '-'} - {getWoDescription(wo) || wo.project || '-'}
                      </option>
                    ))}
                  </select>
                  {availableWorkOrders.length === 0 && lineDate && (
                    <span className="text-[10px] text-amber-400 mt-0.5 block">
                      Tidak ada work order yang tersedia untuk tanggal {formatDisplayDate(lineDate)}.
                    </span>
                  )}
                </div>

                {/* Description timesheet: terisi dari work order, bisa diedit */}
                <div>
                  <label className="block text-slate-400 mb-1">Description</label>
                  <textarea
                    value={lineDescription}
                    onChange={(e) => setLineDescription(e.target.value)}
                    placeholder="Uraian pekerjaan harian (terisi otomatis dari work order, bisa diubah)..."
                    rows={3}
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