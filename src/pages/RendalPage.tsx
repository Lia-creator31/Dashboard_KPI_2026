import React, { useState, useEffect, useCallback, useMemo, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { 
  Building2, ArrowLeft, Search, ChevronRight, ChevronDown, ChevronUp, 
  Pencil, Trash2, Check, Printer, FileCheck, Clock, Sparkles, FileSpreadsheet, Users, Lock, X 
} from 'lucide-react';

interface ParsedMember {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
  biro: string;
  dept: string;
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

interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
}

interface RendalPageProps {
  onBackToLogin: () => void;
}

const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';

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
  return cleanText(b1) === cleanText(b2);
}

export default function RendalPage({ onBackToLogin }: RendalPageProps) {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroName, setSelectedBiroName] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [manualTasks, setManualTasks] = useState<TaskItem[]>([]);
  const [drawingControlMap, setDrawingControlMap] = useState<Record<string, DrawingControlRow[]>>({});
  
  const [isPlannerOpen, setIsPlannerOpen] = useState(false);
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  const loadAllJobCards = useCallback(async () => {
    try {
      const { data, error } = await supabase.from('job_cards').select('*').order('created_at', { ascending: true });
      if (error) return;
      if (data) {
        const formatted: TaskItem[] = data.map((row: any) => ({
          id: row.id,
          biroName: row.biro_name || '',
          project: row.project || '',
          taskName: row.task_name || '',
          startDate: row.start_date || '',
          endDate: row.end_date || '',
          pic: row.pic || '',
          jo: row.jo || '',
          kodeJc: row.kode_jc || '',
          rev: String(row.rev || '0'),
          release: row.release || '',
        }));
        setManualTasks(formatted);
      }
    } catch {}
  }, []);

  useEffect(() => {
    loadAllJobCards();
  }, [loadAllJobCards]);

  const toggleAccordion = (cardKey: string) => {
    setExpandedCards(prev => ({ ...prev, [cardKey]: !prev[cardKey] }));
  };

  const handleManualUploadRealisasi = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        setRealisasiWorkbook(wb);
        alert(`File ${file.name} berhasil dibaca!`);
      } catch {
        alert('Gagal membaca file Realisasi.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleUpdateIm4Excel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        setIm4Workbook(wb);
        alert(`Master Personel "${file.name}" berhasil diperbarui!`);
      } catch {
        alert('Gagal membaca file IM4.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const allParsedFromExcel = useMemo<ParsedMember[]>(() => {
    if (!im4Workbook) return [];
    try {
      const sheetName = im4Workbook.SheetNames.find(s => s.toLowerCase().includes('education')) || im4Workbook.SheetNames[0];
      const sheet = im4Workbook.Sheets[sheetName];
      if (!sheet) return [];
      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      
      let headerIdx = -1;
      for (let r = 0; r < Math.min(15, rawRows.length); r++) {
        const rowVals = (rawRows[r] || []).map(v => String(v).trim().toLowerCase());
        if (rowVals.includes('nama') && (rowVals.includes('nip') || rowVals.includes('status'))) {
          headerIdx = r;
          break;
        }
      }
      if (headerIdx === -1) return [];

      const headers = rawRows[headerIdx].map(v => String(v).trim().toLowerCase());
      const namaCol = headers.findIndex(h => h === 'nama');
      const nipCol = headers.findIndex(h => h === 'nip');
      const statusCol = headers.findIndex(h => h === 'status');
      const unitCol = headers.findIndex(h => h.includes('unit'));
      const jabatanCol = headers.findIndex(h => h.includes('jabatan'));

      let currentDept = '';
      let currentBiro = '';
      const results: ParsedMember[] = [];

      for (let r = headerIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row) continue;
        const nama = String(row[namaCol] || '').trim();
        if (!nama || nama.toLowerCase() === 'nan') continue;

        const nip = String(row[nipCol] || '').trim();
        const statusRaw = String(row[statusCol] || '').trim();
        const unit = String(row[unitCol] || '').trim();
        const jabatan = String(row[jabatanCol] || '').trim();

        results.push({
          nama,
          nip,
          status: statusRaw || 'PKWTT',
          jabatan,
          biro: currentBiro || 'Biro Umum',
          dept: currentDept || 'Departemen Umum'
        });
      }
      return results;
    } catch {
      return [];
    }
  }, [im4Workbook]);

  // Hitung jumlah tugas yang menumpuk/belum disetujui jobcard-nya (Pending/Menunggu Planner)
  const pendingTasksCount = useMemo(() => {
    return manualTasks.filter(t => !t.kodeJc || t.kodeJc.trim() === '').length;
  }, [manualTasks]);

  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;

    const { error } = await supabase.from('job_cards').update({ kode_jc: inputVal, status: 'approved' }).eq('id', taskId);
    if (error) {
      alert('Gagal simpan Jobcard: ' + error.message);
      return;
    }
    alert('Jobcard berhasil disimpan!');
    loadAllJobCards();
  };

  const filteredDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5 cursor-pointer" onClick={onBackToLogin}>
            <div className="p-1.5 bg-purple-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <span className="font-bold text-sm text-white">PORTAL RENDAL</span>
          </div>

          <div className="flex items-center gap-2">
            {/* Tombol Planner dengan Badge Angka Tugas Menumpuk */}
            <button
              onClick={() => setIsPlannerOpen(true)}
              className="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 cursor-pointer relative"
            >
              <Lock className="w-3.5 h-3.5" /> 
              <span>Planner Panel</span>
              {pendingTasksCount > 0 && (
                <span className="absolute -top-1.5 -right-1.5 bg-rose-500 text-white text-[10px] font-bold px-1.5 py-0.2 rounded-full border border-slate-950 animate-pulse">
                  {pendingTasksCount}
                </span>
              )}
            </button>

            <button
              onClick={onBackToLogin}
              className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer"
            >
              Keluar
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-6">
        {/* Kontrol Utama File Excel */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
          <div>
            <h2 className="text-lg font-bold text-white">Manajemen Data & Master</h2>
            <span className="text-xs text-slate-400">Akses seluruh data karyawan organik dan subkon</span>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
              <Users className="w-3.5 h-3.5 text-blue-400" />
              <span>Update IM4</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleUpdateIm4Excel} className="hidden" />
            </label>

            <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
              <Clock className="w-3.5 h-3.5 text-emerald-400" />
              <span>Upload Realisasi JO</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleManualUploadRealisasi} className="hidden" />
            </label>

            <input
              type="text"
              placeholder="Cari Departemen..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white focus:outline-none w-44"
            />
          </div>
        </div>

        {/* Daftar Departemen */}
        {!selectedDept && (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
            {filteredDepartments.map((dept) => (
              <div
                key={dept.id}
                onClick={() => setSelectedDept(dept)}
                className="bg-slate-900 border border-slate-800 hover:border-purple-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
              >
                <div>
                  <h4 className="font-semibold text-white text-sm">{dept.name}</h4>
                  <span className="text-xs text-slate-400">{dept.biros.length} Biro</span>
                </div>
                <ChevronRight className="w-4 h-4 text-slate-500" />
              </div>
            ))}
          </div>
        )}

        {/* Modal / Panel Planner (Tanpa Password) */}
        {isPlannerOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
              <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel — Approval Jobcard ({pendingTasksCount} Menunggu)
                </span>
                <button onClick={() => setIsPlannerOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-4 max-h-96 overflow-y-auto space-y-2">
                {manualTasks.length > 0 ? (
                  manualTasks.map((task, idx) => (
                    <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                      <div className="min-w-0">
                        <div className="font-semibold text-white truncate">
                          {task.pic} <span className="font-mono text-slate-500">#{idx + 1}</span>
                        </div>
                        <div className="text-slate-400 text-[11px] truncate">{task.taskName}</div>
                        <div className="text-emerald-400 font-mono text-[10px]">{task.project} • {task.biroName}</div>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0">
                        <input
                          type="text"
                          value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                          onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                          placeholder="No Jobcard..."
                          className="w-36 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase focus:outline-none"
                        />
                        <button
                          onClick={() => handleSaveKodeJcForTask(task.id)}
                          className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 text-white font-bold rounded text-xs cursor-pointer"
                        >
                          Simpan
                        </button>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="py-8 text-center text-xs text-slate-500">Tidak ada pengajuan tugas</div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}