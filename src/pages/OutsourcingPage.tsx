import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import {
  Building2, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, X, Check, Clock, Users, Plus, Calendar,
  Save, RotateCcw
} from 'lucide-react';

// ... (fungsi cleanText, parseToStandardDate, formatDisplayDate, isBiroMatch tetap sama)

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

function isBiroMatch(biro1: string, biro2: string): boolean {
  const b1 = (biro1 || '').toLowerCase().replace('&', ' dan ').trim();
  const b2 = (biro2 || '').toLowerCase().replace('&', ' dan ').trim();
  if (!b1 || !b2) return false;
  if (b1 === b2) return true;
  const c1 = cleanText(b1.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  const c2 = cleanText(b2.replace(/biro|departemen|dept|divisi|dan/gi, ''));
  if (c1 && c2) {
    return c1 === c2 || c1.includes(c2) || c2.includes(c1);
  }
  return false;
}

// Interface untuk TaskItem (Work Order dari Kabiro)
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

// Interface untuk Timesheet Line (detail per tanggal)
interface TimesheetLine {
  id?: string;
  date: string;
  type: string;
  workOrderId: string;
  workOrderCode: string;
  workOrderDesc: string;
  plannedHour: string;
  effectiveHours: string;
  overtimeHours: string;
  idleHours: string;
  hourSpend: string;
  description?: string;
}

// Interface untuk Timesheet Header
interface Timesheet {
  id: string;
  timCode: string;
  userId: string;
  userName: string;
  unitKerja: string;
  divisi: string;
  dateStart: string;
  dateEnd: string;
  status: 'draft' | 'submitted' | 'approved';
  lines: TimesheetLine[];
  createdAt: string;
}

interface ParsedMember {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
  biro: string;
  dept: string;
}

interface OutsourcingPageProps {
  user: UserSession;
  onLogout: () => void;
}

export default function OutsourcingPage({ user, onLogout }: OutsourcingPageProps) {
  // State untuk navigasi
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiro, setSelectedBiro] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  
  // State untuk Timesheet
  const [timesheets, setTimesheets] = useState<Timesheet[]>([]);
  const [workOrders, setWorkOrders] = useState<TaskItem[]>([]);
  const [isCreateModalOpen, setIsCreateModalOpen] = useState(false);
  const [isLineModalOpen, setIsLineModalOpen] = useState(false);
  const [editingTimesheet, setEditingTimesheet] = useState<Timesheet | null>(null);
  
  // State untuk form Timesheet baru
  const [newTimesheet, setNewTimesheet] = useState({
    dateStart: '',
    dateEnd: '',
    unitKerja: '',
    divisi: '',
  });
  
  // State untuk line timesheet
  const [currentLine, setCurrentLine] = useState<Partial<TimesheetLine>>({
    date: '',
    type: 'Regular',
    plannedHour: '00:00',
    effectiveHours: '00:00',
    overtimeHours: '00:00',
    idleHours: '00:00',
    hourSpend: '00:00',
  });
  
  // State untuk lines yang sedang dibuat
  const [pendingLines, setPendingLines] = useState<TimesheetLine[]>([]);
  
  // Load data
  useEffect(() => {
    loadTimesheets();
    loadWorkOrders();
  }, []);
  
  const loadTimesheets = async () => {
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
          } as Timesheet;
        })
      );
      
      setTimesheets(timesheetsWithLines);
    } catch (error) {
      console.error('Error loading timesheets:', error);
    }
  };
  
  const loadWorkOrders = async () => {
    try {
      // Load work orders yang dibuat oleh Kabiro untuk user ini
      const { data, error } = await supabase
        .from('job_cards')
        .select('*')
        .eq('pic', user.nama)
        .or(`status.eq.approved,status.eq.pending`);
      
      if (error) throw error;
      
      setWorkOrders(data || []);
    } catch (error) {
      console.error('Error loading work orders:', error);
    }
  };
  
  // Generate TIM Code unik
  const generateTimCode = () => {
    const timestamp = Date.now();
    const random = Math.floor(Math.random() * 1000);
    return `TIM-${timestamp}${random}`;
  };
  
  // Filter work orders yang available untuk tanggal tertentu
  const getAvailableWorkOrders = (date: string) => {
    return workOrders.filter(wo => {
      const woStart = new Date(wo.startDate);
      const woEnd = new Date(wo.endDate);
      const checkDate = new Date(date);
      return checkDate >= woStart && checkDate <= woEnd;
    });
  };
  
  // Handler untuk membuat timesheet baru
  const handleCreateTimesheet = async () => {
    if (!newTimesheet.dateStart || !newTimesheet.dateEnd) {
      alert('Mohon isi Date Start dan Date End');
      return;
    }
    
    if (pendingLines.length === 0) {
      alert('Mohon tambahkan至少 satu line timesheet');
      return;
    }
    
    try {
      const timCode = generateTimCode();
      
      // Insert timesheet header
      const { data: tsData, error: tsError } = await supabase
        .from('outsourcing_timesheets')
        .insert({
          tim_code: timCode,
          user_id: user.nip,
          user_name: user.nama,
          unit_kerja: newTimesheet.unitKerja || user.nama,
          divisi: newTimesheet.divisi || 'Outsourcing',
          date_start: newTimesheet.dateStart,
          date_end: newTimesheet.dateEnd,
          status: 'draft',
          created_at: new Date().toISOString(),
        })
        .select()
        .single();
      
      if (tsError) throw tsError;
      
      // Insert lines
      const linesToInsert = pendingLines.map(line => ({
        timesheet_id: tsData.id,
        date: line.date,
        type: line.type,
        work_order_id: line.workOrderId,
        work_order_code: line.workOrderCode,
        work_order_desc: line.workOrderDesc,
        planned_hour: line.plannedHour,
        effective_hours: line.effectiveHours,
        overtime_hours: line.overtimeHours,
        idle_hours: line.idleHours,
        hour_spend: line.hourSpend,
        description: line.description,
      }));
      
      const { error: linesError } = await supabase
        .from('outsourcing_timesheet_lines')
        .insert(linesToInsert);
      
      if (linesError) throw linesError;
      
      alert('Timesheet berhasil dibuat!');
      setIsCreateModalOpen(false);
      setNewTimesheet({ dateStart: '', dateEnd: '', unitKerja: '', divisi: '' });
      setPendingLines([]);
      loadTimesheets();
    } catch (error) {
      console.error('Error creating timesheet:', error);
      alert('Gagal membuat timesheet');
    }
  };
  
  // Handler untuk add line
  const handleAddLine = () => {
    if (!currentLine.date || !currentLine.workOrderId) {
      alert('Mohon isi Tanggal dan Work Order');
      return;
    }
    
    // Cek apakah tanggal sudah ada di pending lines
    const existingLineIndex = pendingLines.findIndex(l => l.date === currentLine.date);
    
    const lineData: TimesheetLine = {
      date: currentLine.date || '',
      type: currentLine.type || 'Regular',
      workOrderId: currentLine.workOrderId || '',
      workOrderCode: currentLine.workOrderCode || '',
      workOrderDesc: currentLine.workOrderDesc || '',
      plannedHour: currentLine.plannedHour || '00:00',
      effectiveHours: currentLine.effectiveHours || '00:00',
      overtimeHours: currentLine.overtimeHours || '00:00',
      idleHours: currentLine.idleHours || '00:00',
      hourSpend: currentLine.hourSpend || '00:00',
      description: currentLine.description,
    };
    
    if (existingLineIndex >= 0) {
      // Update line yang sudah ada
      const updatedLines = [...pendingLines];
      updatedLines[existingLineIndex] = lineData;
      setPendingLines(updatedLines);
    } else {
      // Add line baru
      setPendingLines([...pendingLines, lineData]);
    }
    
    setIsLineModalOpen(false);
    setCurrentLine({
      date: '',
      type: 'Regular',
      plannedHour: '00:00',
      effectiveHours: '00:00',
      overtimeHours: '00:00',
      idleHours: '00:00',
      hourSpend: '00:00',
    });
  };
  
  // Handler untuk submit timesheet
  const handleSubmitTimesheet = async (timesheetId: string) => {
    try {
      const { error } = await supabase
        .from('outsourcing_timesheets')
        .update({ status: 'submitted' })
        .eq('id', timesheetId);
      
      if (error) throw error;
      
      alert('Timesheet berhasil disubmit!');
      loadTimesheets();
    } catch (error) {
      console.error('Error submitting timesheet:', error);
      alert('Gagal submit timesheet');
    }
  };
  
  // Handler untuk delete timesheet
  const handleDeleteTimesheet = async (timesheetId: string) => {
    if (!window.confirm('Apakah Anda yakin ingin menghapus timesheet ini?')) return;
    
    try {
      // Delete lines dulu
      await supabase
        .from('outsourcing_timesheet_lines')
        .delete()
        .eq('timesheet_id', timesheetId);
      
      // Delete header
      const { error } = await supabase
        .from('outsourcing_timesheets')
        .delete()
        .eq('id', timesheetId);
      
      if (error) throw error;
      
      alert('Timesheet berhasil dihapus');
      loadTimesheets();
    } catch (error) {
      console.error('Error deleting timesheet:', error);
      alert('Gagal menghapus timesheet');
    }
  };
  
  // Filter work orders berdasarkan tanggal
  const getFilteredWorkOrders = (date: string) => {
    if (!date) return [];
    return workOrders.filter(wo => {
      const woStart = new Date(wo.startDate);
      const woEnd = new Date(wo.endDate);
      const checkDate = new Date(date);
      return checkDate >= woStart && checkDate <= woEnd;
    });
  };
  
  // Hitung total hours
  const calculateTotalHours = (lines: TimesheetLine[]) => {
    const total = lines.reduce((acc, line) => {
      const hours = parseFloat(line.hourSpend.replace(':', '.')) || 0;
      return acc + hours;
    }, 0);
    return total.toFixed(2);
  };
  
  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-amber-600 rounded-lg text-white">
              <Users className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-sm text-white">OUTSOURCING TIMESHEET</span>
              <div className="text-xs text-slate-400">{user.nama} - {user.nip}</div>
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
      
      <main className="max-w-7xl mx-auto px-4 py-8">
        {/* Header Actions */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold text-white">Project Timesheet</h1>
            <p className="text-slate-400 text-sm">Kelola timesheet kerja outsourcing Anda</p>
          </div>
          <button
            onClick={() => setIsCreateModalOpen(true)}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-sm font-semibold flex items-center gap-2 cursor-pointer"
          >
            <Plus className="w-4 h-4" />
            Create New
          </button>
        </div>
        
        {/* Timesheet List */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {timesheets.map((ts) => (
            <div
              key={ts.id}
              className="bg-slate-900 border border-slate-800 rounded-xl p-4 hover:border-blue-500/50 transition cursor-pointer"
            >
              <div className="flex items-start justify-between mb-3">
                <div>
                  <h3 className="font-bold text-white text-lg">{ts.timCode}</h3>
                  <span className={`inline-block px-2 py-0.5 rounded text-xs font-semibold mt-1 ${
                    ts.status === 'approved' ? 'bg-emerald-500/20 text-emerald-400' :
                    ts.status === 'submitted' ? 'bg-amber-500/20 text-amber-400' :
                    'bg-slate-700 text-slate-300'
                  }`}>
                    {ts.status.charAt(0).toUpperCase() + ts.status.slice(1)}
                  </span>
                </div>
                <div className="flex gap-1">
                  {ts.status === 'draft' && (
                    <>
                      <button
                        onClick={() => handleSubmitTimesheet(ts.id)}
                        className="p-1.5 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-400 rounded cursor-pointer"
                        title="Submit"
                      >
                        <Check className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => handleDeleteTimesheet(ts.id)}
                        className="p-1.5 bg-rose-600/20 hover:bg-rose-600/30 text-rose-400 rounded cursor-pointer"
                        title="Delete"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </>
                  )}
                </div>
              </div>
              
              <div className="space-y-2 text-sm">
                <div className="flex justify-between">
                  <span className="text-slate-400">Responsible:</span>
                  <span className="text-white font-mono text-xs">{ts.userName}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">Start Date:</span>
                  <span className="text-emerald-400">{formatDisplayDate(ts.dateStart)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">End Date:</span>
                  <span className="text-emerald-400">{formatDisplayDate(ts.dateEnd)}</span>
                </div>
                <div className="flex justify-between pt-2 border-t border-slate-800">
                  <span className="text-slate-400">Total Lines:</span>
                  <span className="text-blue-400 font-semibold">{ts.lines.length} days</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-400">Total Hours:</span>
                  <span className="text-amber-400 font-bold">{calculateTotalHours(ts.lines)} hrs</span>
                </div>
              </div>
            </div>
          ))}
          
          {timesheets.length === 0 && (
            <div className="col-span-full py-12 text-center text-slate-500">
              <Clock className="w-12 h-12 mx-auto mb-3 opacity-50" />
              <p>Belum ada timesheet. Klik "Create New" untuk membuat timesheet.</p>
            </div>
          )}
        </div>
      </main>
      
      {/* Modal Create Timesheet */}
      {isCreateModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl overflow-hidden shadow-2xl">
            <div className="p-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <h2 className="font-bold text-white">New Timesheet</h2>
              <button onClick={() => setIsCreateModalOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 space-y-6">
              {/* Header Info */}
              <div className="grid grid-cols-2 gap-4 text-sm">
                <div>
                  <label className="block text-slate-400 mb-1">User</label>
                  <div className="px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white">
                    {user.nip} - {user.nama}
                  </div>
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Unit Kerja</label>
                  <input
                    type="text"
                    value={newTimesheet.unitKerja}
                    onChange={(e) => setNewTimesheet({...newTimesheet, unitKerja: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                    placeholder="Unit Kerja"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Date Start</label>
                  <input
                    type="date"
                    value={newTimesheet.dateStart}
                    onChange={(e) => setNewTimesheet({...newTimesheet, dateStart: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Date End</label>
                  <input
                    type="date"
                    value={newTimesheet.dateEnd}
                    onChange={(e) => setNewTimesheet({...newTimesheet, dateEnd: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
              </div>
              
              {/* Lines Section */}
              <div>
                <div className="flex items-center justify-between mb-3">
                  <h3 className="font-semibold text-white">Project Timesheet Line</h3>
                  <button
                    onClick={() => setIsLineModalOpen(true)}
                    className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs rounded flex items-center gap-1 cursor-pointer"
                  >
                    <Plus className="w-3 h-3" />
                    Add a line
                  </button>
                </div>
                
                <div className="bg-slate-950 border border-slate-800 rounded-lg overflow-hidden">
                  <table className="w-full text-xs">
                    <thead className="bg-slate-900 text-slate-400">
                      <tr>
                        <th className="px-3 py-2 text-left">Date</th>
                        <th className="px-3 py-2 text-left">Work Order</th>
                        <th className="px-3 py-2 text-left">Description</th>
                        <th className="px-3 py-2 text-right">Planned</th>
                        <th className="px-3 py-2 text-right">Effective</th>
                        <th className="px-3 py-2 text-right">Hour Spend</th>
                        <th className="px-3 py-2 text-center">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800">
                      {pendingLines.map((line, idx) => (
                        <tr key={idx} className="text-slate-300">
                          <td className="px-3 py-2">{formatDisplayDate(line.date)}</td>
                          <td className="px-3 py-2 font-mono text-amber-400">{line.workOrderCode}</td>
                          <td className="px-3 py-2 text-slate-400">{line.workOrderDesc}</td>
                          <td className="px-3 py-2 text-right font-mono">{line.plannedHour}</td>
                          <td className="px-3 py-2 text-right font-mono">{line.effectiveHours}</td>
                          <td className="px-3 py-2 text-right font-mono text-emerald-400">{line.hourSpend}</td>
                          <td className="px-3 py-2 text-center">
                            <button
                              onClick={() => {
                                const newLines = pendingLines.filter((_, i) => i !== idx);
                                setPendingLines(newLines);
                              }}
                              className="text-rose-400 hover:text-rose-300 cursor-pointer"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          </td>
                        </tr>
                      ))}
                      {pendingLines.length === 0 && (
                        <tr>
                          <td colSpan={7} className="px-3 py-8 text-center text-slate-500">
                            Belum ada line. Klik "Add a line" untuk menambahkan.
                          </td>
                        </tr>
                      )}
                    </tbody>
                    <tfoot className="bg-slate-900 font-semibold">
                      <tr>
                        <td colSpan={5} className="px-3 py-2 text-right text-slate-400">Total:</td>
                        <td className="px-3 py-2 text-right text-emerald-400 font-mono">
                          {calculateTotalHours(pendingLines)}
                        </td>
                        <td></td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
              
              {/* Actions */}
              <div className="flex justify-end gap-2 pt-4 border-t border-slate-800">
                <button
                  onClick={() => {
                    setIsCreateModalOpen(false);
                    setPendingLines([]);
                    setNewTimesheet({ dateStart: '', dateEnd: '', unitKerja: '', divisi: '' });
                  }}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  onClick={handleCreateTimesheet}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded cursor-pointer flex items-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  Save Timesheet
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      
      {/* Modal Create Line */}
      {isLineModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-2xl overflow-hidden shadow-2xl max-h-[90vh] overflow-y-auto">
            <div className="p-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <h2 className="font-bold text-white">Create Lines</h2>
              <button onClick={() => setIsLineModalOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 space-y-4 text-sm">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-slate-400 mb-1">Type</label>
                  <select
                    value={currentLine.type}
                    onChange={(e) => setCurrentLine({...currentLine, type: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  >
                    <option value="Regular">Regular</option>
                    <option value="Overtime">Overtime</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Date</label>
                  <input
                    type="date"
                    value={currentLine.date}
                    onChange={(e) => {
                      const date = e.target.value;
                      setCurrentLine({
                        ...currentLine,
                        date,
                        workOrderId: '',
                        workOrderCode: '',
                        workOrderDesc: '',
                      });
                    }}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
              </div>
              
              <div>
                <label className="block text-slate-400 mb-1">Work Order</label>
                <select
                  value={currentLine.workOrderId}
                  onChange={(e) => {
                    const wo = workOrders.find(w => w.id === e.target.value);
                    if (wo) {
                      setCurrentLine({
                        ...currentLine,
                        workOrderId: wo.id,
                        workOrderCode: wo.kodeJc,
                        workOrderDesc: wo.taskName,
                      });
                    }
                  }}
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  disabled={!currentLine.date}
                >
                  <option value="">Pilih Work Order</option>
                  {getFilteredWorkOrders(currentLine.date).map(wo => (
                    <option key={wo.id} value={wo.id}>
                      {wo.kodeJc} - {wo.taskName} ({formatDisplayDate(wo.startDate)} - {formatDisplayDate(wo.endDate)})
                    </option>
                  ))}
                </select>
                {!currentLine.date && (
                  <p className="text-xs text-amber-400 mt-1">Pilih tanggal terlebih dahulu</p>
                )}
              </div>
              
              <div>
                <label className="block text-slate-400 mb-1">Description</label>
                <textarea
                  value={currentLine.description || ''}
                  onChange={(e) => setCurrentLine({...currentLine, description: e.target.value})}
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  rows={2}
                />
              </div>
              
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-slate-400 mb-1">Planned Hour</label>
                  <input
                    type="time"
                    value={currentLine.plannedHour}
                    onChange={(e) => setCurrentLine({...currentLine, plannedHour: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Effective Hours</label>
                  <input
                    type="time"
                    value={currentLine.effectiveHours}
                    onChange={(e) => setCurrentLine({...currentLine, effectiveHours: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Overtime Hours</label>
                  <input
                    type="time"
                    value={currentLine.overtimeHours}
                    onChange={(e) => setCurrentLine({...currentLine, overtimeHours: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Idle Hours</label>
                  <input
                    type="time"
                    value={currentLine.idleHours}
                    onChange={(e) => setCurrentLine({...currentLine, idleHours: e.target.value})}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded text-white"
                  />
                </div>
              </div>
              
              <div className="flex justify-end gap-2 pt-4 border-t border-slate-800">
                <button
                  onClick={() => setIsLineModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  onClick={handleAddLine}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded cursor-pointer flex items-center gap-2"
                >
                  <Check className="w-4 h-4" />
                  Add Line
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}