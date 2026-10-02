import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { departmentsData, Department } from '../data';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  Building2, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, X, Check, Clock, Plus, Calendar, Briefcase
} from 'lucide-react';

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
  id?: string;
  date: string;
  workOrderId: string;
  workOrderCode: string;
  description: string;
  effectiveHours: string;
  overtimeHours: string;
  plannedHour?: string;
  idleHours?: string;
  type?: string;
  project?: string;
  wbsActivity?: string;
  analyticAccount?: string;
  day?: string;
  isLembur?: boolean;
}

interface Timesheet {
  id: string;
  kodeTimesheet: string;
  userId: string;
  userName: string;
  dateStart: string;
  dateEnd: string;
  status: 'draft' | 'submitted' | 'approved';
  lines: TimesheetLine[];
  createdAt: string;
}

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
}

function generateTimesheetCode(): string {
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 1000);
  return `TIM-${timestamp}${random}`;
}

function formatDateToDay(dateStr: string): string {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const date = new Date(dateStr);
  return days[date.getDay()];
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiro, setSelectedBiro] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  
  // State untuk Timesheet
  const [timesheets, setTimesheets] = useState<Timesheet[]>([]);
  const [isCreatingTimesheet, setIsCreatingTimesheet] = useState(false);
  const [currentTimesheet, setCurrentTimesheet] = useState<Partial<Timesheet>>({
    dateStart: '',
    dateEnd: '',
    lines: [],
  });
  const [isAddingLine, setIsAddingLine] = useState(false);
  const [currentLine, setCurrentLine] = useState<Partial<TimesheetLine>>({
    date: '',
    effectiveHours: '',
    overtimeHours: '',
    plannedHour: '00:00',
    idleHours: '00:00',
  });
  const [availableWorkOrders, setAvailableWorkOrders] = useState<TaskItem[]>([]);
  const [selectedLineIndex, setSelectedLineIndex] = useState<number | null>(null);

  // Load Timesheet dari Supabase
  const loadTimesheets = useCallback(async () => {
    try {
      const { data, error } = await supabase
        .from('outsourcing_timesheets')
        .select('*')
        .eq('user_id', user.nip)
        .order('created_at', { ascending: false });
      
      if (error) throw error;
      
      // Load lines untuk setiap timesheet
      const timesheetsWithLines = await Promise.all(
        (data || []).map(async (ts) => {
          const { data: lines } = await supabase
            .from('outsourcing_timesheet_lines')
            .select('*')
            .eq('timesheet_id', ts.id);
          
          return {
            ...ts,
            lines: lines || [],
          };
        })
      );
      
      setTimesheets(timesheetsWithLines as Timesheet[]);
    } catch (error) {
      console.error('Error loading timesheets:', error);
    }
  }, [user.nip]);

  // Load Work Order yang tersedia untuk user ini
const loadAvailableWorkOrders = useCallback(async (startDate: string, endDate: string) => {
  try {
    // Query yang lebih flexible - ambil semua work order untuk user ini
    let query = supabase
      .from('job_cards')
      .select('*')
      .eq('status', 'approved');
    
    // Filter berdasarkan nama PIC (tanpa NIP)
    const userNameOnly = user.nama.split(' - ').pop() || user.nama;
    query = query.ilike('pic', `%${userNameOnly}%`);
    
    // Atau filter berdasarkan NIP juga
    const userNip = user.nip;
    
    const { data, error } = await query;
    
    if (error) throw error;
    
    // Filter di client-side berdasarkan tanggal
    const filteredData = (data || []).filter(wo => {
      const woStart = new Date(wo.start_date);
      const woEnd = new Date(wo.end_date);
      const selectedDate = new Date(startDate);
      
      // Cek apakah tanggal yang dipilih ada dalam range work order
      return selectedDate >= woStart && selectedDate <= woEnd;
    });
    
    setAvailableWorkOrders(filteredData);
  } catch (error) {
    console.error('Error loading work orders:', error);
  }
}, [user.nama, user.nip]);

  useEffect(() => {
    loadTimesheets();
  }, [loadTimesheets]);

  // Handler untuk membuat timesheet baru
  const handleCreateTimesheet = () => {
    setIsCreatingTimesheet(true);
    setCurrentTimesheet({
      dateStart: '',
      dateEnd: '',
      lines: [],
    });
  };

  // Handler untuk menyimpan timesheet
  const handleSaveTimesheet = async () => {
    if (!currentTimesheet.dateStart || !currentTimesheet.dateEnd) {
      alert('Tanggal mulai dan tanggal akhir harus diisi!');
      return;
    }

    try {
      const newTimesheet = {
        kode_timesheet: generateTimesheetCode(),
        user_id: user.nip,
        user_name: user.nama,
        date_start: currentTimesheet.dateStart,
        date_end: currentTimesheet.dateEnd,
        status: 'draft',
      };

      const { data, error } = await supabase
        .from('outsourcing_timesheets')
        .insert(newTimesheet)
        .select()
        .single();

      if (error) throw error;

      // Simpan lines
      if (currentTimesheet.lines && currentTimesheet.lines.length > 0) {
        const linesToInsert = currentTimesheet.lines.map(line => ({
          timesheet_id: data.id,
          date: line.date,
          work_order_id: line.workOrderId,
          work_order_code: line.workOrderCode,
          description: line.description,
          effective_hours: line.effectiveHours,
          overtime_hours: line.overtimeHours,
          planned_hour: line.plannedHour || '00:00',
          idle_hours: line.idleHours || '00:00',
        }));

        await supabase.from('outsourcing_timesheet_lines').insert(linesToInsert);
      }

      alert('Timesheet berhasil disimpan!');
      setIsCreatingTimesheet(false);
      loadTimesheets();
    } catch (error) {
      console.error('Error saving timesheet:', error);
      alert('Gagal menyimpan timesheet');
    }
  };

  // Handler untuk add line
  const handleAddLine = () => {
    if (!currentTimesheet.dateStart || !currentTimesheet.dateEnd) {
      alert('Silakan isi tanggal mulai dan tanggal akhir terlebih dahulu!');
      return;
    }
    
    setIsAddingLine(true);
    loadAvailableWorkOrders(currentTimesheet.dateStart!, currentTimesheet.dateEnd!);
    setCurrentLine({
      date: '',
      effectiveHours: '',
      overtimeHours: '',
      plannedHour: '00:00',
      idleHours: '00:00',
    });
  };

  // Handler untuk menyimpan line
  const handleSaveLine = () => {
    if (!currentLine.date || !currentLine.workOrderId) {
      alert('Tanggal dan Work Order harus diisi!');
      return;
    }

    const workOrder = availableWorkOrders.find(wo => wo.id === currentLine.workOrderId);
    
    const newLine: TimesheetLine = {
      date: currentLine.date!,
      workOrderId: currentLine.workOrderId!,
      workOrderCode: workOrder?.kodeJc || '',
      description: currentLine.description || '',
      effectiveHours: currentLine.effectiveHours || '0',
      overtimeHours: currentLine.overtimeHours || '0',
      plannedHour: currentLine.plannedHour || '00:00',
      idleHours: currentLine.idleHours || '00:00',
      day: formatDateToDay(currentLine.date),
    };

    if (selectedLineIndex !== null) {
      // Update existing line
      const updatedLines = [...(currentTimesheet.lines || [])];
      updatedLines[selectedLineIndex] = newLine;
      setCurrentTimesheet({ ...currentTimesheet, lines: updatedLines });
    } else {
      // Add new line
      setCurrentTimesheet({
        ...currentTimesheet,
        lines: [...(currentTimesheet.lines || []), newLine],
      });
    }

    setIsAddingLine(false);
    setSelectedLineIndex(null);
    setCurrentLine({
      date: '',
      effectiveHours: '',
      overtimeHours: '',
      plannedHour: '00:00',
      idleHours: '00:00',
    });
  };

  // Handler untuk edit line
  const handleEditLine = (index: number) => {
    const line = currentTimesheet.lines?.[index];
    if (line) {
      setCurrentLine(line);
      setSelectedLineIndex(index);
      setIsAddingLine(true);
      loadAvailableWorkOrders(currentTimesheet.dateStart!, currentTimesheet.dateEnd!);
    }
  };

  // Handler untuk delete line
  const handleDeleteLine = (index: number) => {
    if (confirm('Hapus baris ini?')) {
      const updatedLines = currentTimesheet.lines?.filter((_, i) => i !== index) || [];
      setCurrentTimesheet({ ...currentTimesheet, lines: updatedLines });
    }
  };

  // Handler untuk submit timesheet
  const handleSubmitTimesheet = async (timesheetId: string) => {
    try {
      await supabase
        .from('outsourcing_timesheets')
        .update({ status: 'submitted' })
        .eq('id', timesheetId);
      
      alert('Timesheet berhasil disubmit!');
      loadTimesheets();
    } catch (error) {
      console.error('Error submitting timesheet:', error);
      alert('Gagal submit timesheet');
    }
  };

  // Filter work order berdasarkan tanggal
  const getWorkOrdersForDate = (date: string) => {
    return availableWorkOrders.filter(wo => {
      return date >= wo.startDate && date <= wo.endDate;
    });
  };

  // Render halaman utama outsourcing
  if (!selectedDept) {
    return (
      <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
        <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
          <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
            <div className="flex items-center gap-2.5">
              <div className="p-1.5 bg-amber-600 rounded-lg text-white">
                <HardHat className="w-4 h-4" />
              </div>
              <span className="font-bold text-sm text-white">OUTSOURCING TIMESHEET</span>
            </div>
            <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
              Keluar
            </button>
          </div>
        </header>

        <main className="max-w-6xl mx-auto px-4 py-8">
          <div className="flex items-center justify-between mb-6">
            <div>
              <h1 className="text-2xl font-bold text-white">Project Timesheet</h1>
              <p className="text-slate-400 text-sm">Kelola timesheet kerja Anda</p>
            </div>
            <button
              onClick={handleCreateTimesheet}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-sm font-semibold flex items-center gap-2 cursor-pointer"
            >
              <Plus className="w-4 h-4" /> Create
            </button>
          </div>

          {/* Daftar Timesheet */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {timesheets.map((ts) => (
              <div key={ts.id} className="bg-slate-900 border border-slate-800 rounded-xl p-4 hover:border-blue-500 transition">
                <div className="flex items-start justify-between mb-3">
                  <h3 className="font-mono font-bold text-blue-400">{ts.kodeTimesheet}</h3>
                  <span className={`px-2 py-1 rounded text-xs font-semibold ${
                    ts.status === 'approved' ? 'bg-emerald-500/10 text-emerald-400' :
                    ts.status === 'submitted' ? 'bg-amber-500/10 text-amber-400' :
                    'bg-slate-700 text-slate-300'
                  }`}>
                    {ts.status === 'approved' ? 'Approved' : ts.status === 'submitted' ? 'Submitted' : 'Draft'}
                  </span>
                </div>
                
                <div className="space-y-2 text-xs">
                  <div>
                    <span className="text-slate-400">Responsible:</span>
                    <p className="text-white font-medium">{ts.userName}</p>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-slate-400">Start Date</span>
                    <span className="text-emerald-400">{ts.dateStart}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-slate-400">End Date</span>
                    <span className="text-emerald-400">{ts.dateEnd}</span>
                  </div>
                  <div className="flex justify-between pt-2 border-t border-slate-800">
                    <span className="text-slate-400">Total Lines</span>
                    <span className="text-white">{ts.lines?.length || 0}</span>
                  </div>
                </div>

                {ts.status === 'draft' && (
                  <div className="mt-4 flex gap-2">
                    <button
                      onClick={() => handleSubmitTimesheet(ts.id)}
                      className="flex-1 px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs cursor-pointer"
                    >
                      Submit
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        </main>

        {/* Modal Create Timesheet */}
        {isCreatingTimesheet && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[90vh] overflow-y-auto">
              <div className="p-4 border-b border-slate-800 flex items-center justify-between">
                <h2 className="text-lg font-bold text-white">New</h2>
                <button onClick={() => setIsCreatingTimesheet(false)} className="text-slate-400 hover:text-white cursor-pointer">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-6 space-y-6">
                {/* Header Info */}
                <div className="grid grid-cols-2 gap-4 text-sm">
                  <div>
                    <label className="block text-slate-400 mb-1">User</label>
                    <input
                      type="text"
                      value={`${user.nip} - ${user.nama}`}
                      readOnly
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">Unit Kerja</label>
                    <input
                      type="text"
                      value="Outsourcing"
                      readOnly
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">Date Start</label>
                    <input
                      type="date"
                      value={currentTimesheet.dateStart || ''}
                      onChange={(e) => setCurrentTimesheet({ ...currentTimesheet, dateStart: e.target.value })}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">Date End</label>
                    <input
                      type="date"
                      value={currentTimesheet.dateEnd || ''}
                      onChange={(e) => setCurrentTimesheet({ ...currentTimesheet, dateEnd: e.target.value })}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white"
                    />
                  </div>
                </div>

                {/* Project Timesheet Line */}
                <div>
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="text-sm font-semibold text-slate-300">Project Timesheet Line</h3>
                    <button
                      onClick={handleAddLine}
                      className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs cursor-pointer flex items-center gap-1"
                    >
                      <Plus className="w-3 h-3" /> Add a line
                    </button>
                  </div>

                  <div className="bg-slate-950 border border-slate-800 rounded-lg overflow-hidden">
                    <table className="w-full text-xs">
                      <thead className="bg-slate-900">
                        <tr>
                          <th className="px-3 py-2 text-left text-slate-400">Date</th>
                          <th className="px-3 py-2 text-left text-slate-400">Work Order</th>
                          <th className="px-3 py-2 text-left text-slate-400">Description</th>
                          <th className="px-3 py-2 text-right text-slate-400">Effective Hours</th>
                          <th className="px-3 py-2 text-right text-slate-400">Overtime Hours</th>
                          <th className="px-3 py-2 text-center text-slate-400">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800">
                        {currentTimesheet.lines?.map((line, index) => (
                          <tr key={index} className="hover:bg-slate-900/50">
                            <td className="px-3 py-2 text-white">{line.date}</td>
                            <td className="px-3 py-2 text-emerald-400 font-mono">{line.workOrderCode}</td>
                            <td className="px-3 py-2 text-slate-300">{line.description}</td>
                            <td className="px-3 py-2 text-right text-white">{line.effectiveHours}</td>
                            <td className="px-3 py-2 text-right text-amber-400">{line.overtimeHours}</td>
                            <td className="px-3 py-2 text-center">
                              <div className="flex justify-center gap-1">
                                <button
                                  onClick={() => handleEditLine(index)}
                                  className="p-1 text-blue-400 hover:bg-blue-500/20 rounded cursor-pointer"
                                >
                                  <Pencil className="w-3 h-3" />
                                </button>
                                <button
                                  onClick={() => handleDeleteLine(index)}
                                  className="p-1 text-rose-400 hover:bg-rose-500/20 rounded cursor-pointer"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                        {(!currentTimesheet.lines || currentTimesheet.lines.length === 0) && (
                          <tr>
                            <td colSpan={6} className="px-3 py-8 text-center text-slate-500 text-xs">
                              Click "Add a line" to add timesheet entries
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Total Hours */}
                <div className="flex justify-end">
                  <div className="text-sm">
                    <span className="text-slate-400">Total Hours: </span>
                    <span className="font-bold text-white">
                      {currentTimesheet.lines?.reduce((acc, line) => {
                        const eff = parseFloat(line.effectiveHours) || 0;
                        const ot = parseFloat(line.overtimeHours) || 0;
                        return acc + eff + ot;
                      }, 0).toFixed(2) || '0.00'}
                    </span>
                  </div>
                </div>

                {/* Action Buttons */}
                <div className="flex justify-end gap-3 pt-4 border-t border-slate-800">
                  <button
                    onClick={() => setIsCreatingTimesheet(false)}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
                  >
                    Discard
                  </button>
                  <button
                    onClick={handleSaveTimesheet}
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg cursor-pointer"
                  >
                    Save
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Modal Add/Edit Line */}
        {isAddingLine && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
              <div className="p-4 border-b border-slate-800 flex items-center justify-between">
                <h2 className="text-lg font-bold text-white">Create Lines</h2>
                <button onClick={() => { setIsAddingLine(false); setSelectedLineIndex(null); }} className="text-slate-400 hover:text-white cursor-pointer">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-6 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">Type</label>
                    <select className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm">
                      <option value="">Select Type</option>
                      <option value="regular">Regular</option>
                      <option value="overtime">Overtime</option>
                    </select>
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">Work Center</label>
                    <input
                      type="text"
                      value="Outsourcing"
                      readOnly
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
                    />
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">User</label>
                    <input
                      type="text"
                      value={`${user.nip} - ${user.nama}`}
                      readOnly
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
                    />
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">Date</label>
                    <input
                      type="date"
                      value={currentLine.date || ''}
                      onChange={(e) => setCurrentLine({ ...currentLine, date: e.target.value })}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
                    />
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">Work Order</label>
                    <select
                      value={currentLine.workOrderId || ''}
                      onChange={(e) => {
                        const wo = availableWorkOrders.find(w => w.id === e.target.value);
                        setCurrentLine({
                          ...currentLine,
                          workOrderId: e.target.value,
                          description: wo?.taskName || '',
                        });
                      }}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
                    >
                      <option value="">Select Work Order</option>
                      {availableWorkOrders.map((wo) => (
                        <option key={wo.id} value={wo.id}>
                          {wo.kodeJc} - {wo.taskName}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="col-span-2">
                    <label className="block text-slate-400 mb-1 text-sm">Description</label>
                    <textarea
                      value={currentLine.description || ''}
                      onChange={(e) => setCurrentLine({ ...currentLine, description: e.target.value })}
                      rows={3}
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-400 mb-1 text-sm">Effective Hours</label>
                    <input
                      type="text"
                      value={currentLine.effectiveHours || ''}
                      onChange={(e) => setCurrentLine({ ...currentLine, effectiveHours: e.target.value })}
                      placeholder="00:00"
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-400 mb-1 text-sm">Overtime Hours</label>
                    <input
                      type="text"
                      value={currentLine.overtimeHours || ''}
                      onChange={(e) => setCurrentLine({ ...currentLine, overtimeHours: e.target.value })}
                      placeholder="00:00"
                      className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white text-sm"
                    />
                  </div>

                  {currentLine.date && (
                    <div className="col-span-2">
                      <label className="block text-slate-400 mb-1 text-sm">Day</label>
                      <input
                        type="text"
                        value={formatDateToDay(currentLine.date)}
                        readOnly
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-slate-400 text-sm"
                      />
                    </div>
                  )}
                </div>

                <div className="flex justify-end gap-3 pt-4 border-t border-slate-800">
                  <button
                    onClick={() => { setIsAddingLine(false); setSelectedLineIndex(null); }}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleSaveLine}
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg cursor-pointer"
                  >
                    Save Line
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // Render halaman lain (jika perlu navigasi tambahan)
  return null;
}