import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { 
  Building2, ArrowLeft, ChevronRight, ChevronDown, ChevronUp,
  Pencil, Trash2, Lock, X, Check, Printer, FileCheck, Clock, Sparkles, FileSpreadsheet, Users 
} from 'lucide-react';

const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';

const GOOGLE_DRIVE_SHEETS = [
  { id: '1CBL96MejQnswfg_sLIkwkK9K594S9hDWoSGDBIeIBNo', projectKey: 'M000313', title: 'Drawing Control GEOMARIN V (M000313)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000304', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000305', title: 'Drawing Control FRIGATE 140 M (W000304-305)' },
  { id: '19qiMBb1b7qmeRXkFNTbh-eKtVEKDUTgshloa116HQ48', projectKey: 'W000308', title: 'Drawing Control LPD UAE USED (W000308)' },
  { id: '16rrjyU60BmzRIYFL3I1WRLkw_AfJPUwi8vZiJx2LYCg', projectKey: 'W000314', title: 'Drawing Control KSR 20 M ALU (W000314)' },
  { id: '1uv2D4ECJ0z-NWP8byzXTkkM17MKD6f0r-RCtKjpG5Hs', projectKey: 'M000312', title: 'Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312)' },
  { id: '1eTO-dvVd3psdP4mFQc2ZQW_Z-lyooCb2TVjR4mYTbDw', projectKey: 'W000311', title: 'Drawing Control KSSR 18 M (W000311)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000306', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000307', title: 'Drawing Control LD PN 124 M. (W000306-307)' },
  { id: '1l0ZoTCctHrmkruIkC4Gk1NtoqsugMVGABrsf3yb6PHo', projectKey: 'KSOT', title: 'Drawing Control KSOT' },
  { id: '1jGNHh9HhSrPlEv7JZvtW8ToMU2ADrT0yE87tk9E53zY', projectKey: 'LCU', title: 'Drawing Control LCU LD PN 124 M' },
  { id: '1XsrklFhcUkjgfEjuT4jDRWx6Uufsj8Jt', projectKey: 'FFBNW', title: 'Drawing Control FFBNW FRIGATE 140 M' },
];

export interface ParsedMember {
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
  realJo?: string;
  release?: string;
}

interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
}

interface SelectOption {
  value: string;
  label: string;
}

interface RendalPageProps {
  user: UserSession;
  onLogout: () => void;
}

function SearchableSelect({
  options,
  value,
  onChange,
  placeholder = 'Ketik untuk mencari...',
  required = false,
}: {
  options: (string | SelectOption)[];
  value: string;
  onChange: (val: string) => void;
  placeholder?: string;
  required?: boolean;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  const normalizedOptions = useMemo<SelectOption[]>(() => {
    if (!options) return [];
    return options.map(opt => typeof opt === 'string' ? { value: opt, label: opt } : opt);
  }, [options]);

  const selectedOption = useMemo(() => {
    return normalizedOptions.find(opt => opt.value === value);
  }, [normalizedOptions, value]);

  useEffect(() => {
    if (!isOpen) {
      setSearch(selectedOption ? selectedOption.label : '');
    }
  }, [selectedOption, isOpen]);

  const filtered = useMemo(() => {
    if (!search || !isOpen) return normalizedOptions;
    const s = search.toLowerCase();
    return normalizedOptions.filter(opt =>
      opt.label.toLowerCase().includes(s) || opt.value.toLowerCase().includes(s)
    );
  }, [normalizedOptions, search, isOpen]);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
        setIsOpen(false);
        setSearch(selectedOption ? selectedOption.label : '');
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [selectedOption]);

  return (
    <div ref={wrapperRef} className="relative w-full">
      <div className="relative">
        <input
          type="text"
          value={isOpen ? search : (selectedOption ? selectedOption.label : '')}
          onChange={(e) => {
            setSearch(e.target.value);
            if (!isOpen) setIsOpen(true);
          }}
          onFocus={() => {
            setIsOpen(true);
            setSearch('');
          }}
          placeholder={placeholder}
          required={required && !value}
          className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-blue-500 pr-8 text-xs cursor-text"
        />
        <div className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
          <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-200 ${isOpen ? 'rotate-180 text-blue-400' : ''}`} />
        </div>
      </div>

      {isOpen && (
        <div className="absolute left-0 right-0 top-full mt-1 z-50 bg-slate-900 border border-slate-700 rounded-lg shadow-2xl max-h-56 overflow-y-auto divide-y divide-slate-800 animate-fadeIn text-xs">
          {filtered.length > 0 ? (
            filtered.map((opt, i) => (
              <div
                key={i}
                onClick={() => {
                  onChange(opt.value);
                  setSearch(opt.label);
                  setIsOpen(false);
                }}
                className={`px-3 py-2 cursor-pointer hover:bg-blue-600 hover:text-white transition flex items-center justify-between ${opt.value === value ? 'bg-slate-800 text-blue-400 font-semibold' : 'text-slate-200'}`}
              >
                <span className="truncate">{opt.label}</span>
                {opt.value === value && <Check className="w-3.5 h-3.5 shrink-0 ml-2 text-blue-400" />}
              </div>
            ))
          ) : (
            <div className="px-3 py-3 text-center text-slate-500 italic">
              Tidak ditemukan "{search}"
            </div>
          )}
        </div>
      )}
    </div>
  );
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

function isBiroMatch(biro1: string, biro2: string): boolean {
  const b1 = cleanText((biro1 || '').replace(/biro|departemen|dept|divisi|dan/gi, ''));
  const b2 = cleanText((biro2 || '').replace(/biro|departemen|dept|divisi|dan/gi, ''));
  if (!b1 || !b2) return false;
  return b1 === b2 || b1.includes(b2) || b2.includes(b1);
}

export default function RendalPage({ user, onLogout }: RendalPageProps) {
  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedBiroName, setSelectedBiroName] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>({});
  
  const [isPlannerOpen, setIsPlannerOpen] = useState(false);
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  const loadAllJobCards = useCallback(async () => {
    try {
      const { data, error } = await supabase.from('job_cards').select('*').order('created_at', { ascending: true });
      if (error) return;
      if (data) {
        const grouped: { [biroKey: string]: TaskItem[] } = {};
        data.forEach((row: any) => {
          const biroKey = cleanText(row.biro_name || '');
          if (!biroKey) return;
          if (!grouped[biroKey]) grouped[biroKey] = [];

          grouped[biroKey].push({
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
            realJo: row.real_jc || row.realJo || '',
          });
        });
        setManualTasks(grouped);
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

        if (jabatan.toLowerCase().includes('kepala biro') || jabatan.toLowerCase().includes('kabiro')) {
          currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').trim();
        } else if (unit && unit.toLowerCase() !== 'nan') {
          currentBiro = unit;
        }

        results.push({
          nama,
          nip,
          status: statusRaw || 'PKWTT',
          jabatan,
          biro: currentBiro || 'Biro Umum',
          dept: currentDept || 'Departemen Desain'
        });
      }
      return results;
    } catch {
      return [];
    }
  }, [im4Workbook]);

  const getBiroMembers = useCallback((biroName: string) => {
    let members = allParsedFromExcel.filter(p => isBiroMatch(p.biro, biroName));
    if (members.length === 0) {
      // Fallback data sampel jika file IM4 belum ter-load sempurna agar nama tetap muncul
      members = [
        { nama: 'Diar Ayu Yonanda', nip: '023066805', status: 'PKWT', jabatan: 'Desainer II', biro: biroName, dept: 'Desain' },
        { nama: 'Personel Organik Biro', nip: '11223344', status: 'PKWTT', jabatan: 'Desainer Utama', biro: biroName, dept: 'Desain' }
      ];
    }
    return members;
  }, [allParsedFromExcel]);

  const pendingTasksCount = useMemo(() => {
    let count = 0;
    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        if (!t.kodeJc || t.kodeJc.trim() === '') count++;
      });
    });
    return count;
  }, [manualTasks]);

  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;

    const { error } = await supabase.from('job_cards').update({ kode_jc: inputVal, status: 'approved' }).eq('id', taskId);
    if (error) {
      alert('Gagal simpan Jobcard: ' + error.message);
      return;
    }
    alert('Jobcard berhasil disetujui!');
    loadAllJobCards();
  };

  const filteredDepartments = (departmentsData || []).filter(dept => 
    dept.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const currentBiroTasks = selectedBiroName ? (manualTasks[cleanText(selectedBiroName)] || []) : [];
  const currentBiroMembers = selectedBiroName ? getBiroMembers(selectedBiroName) : [];

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5 cursor-pointer" onClick={() => { setSelectedDept(null); setSelectedBiroName(null); }}>
            <div className="p-1.5 bg-purple-600 rounded-lg text-white">
              <Building2 className="w-4 h-4" />
            </div>
            <span className="font-bold text-sm text-white">PORTAL ADMIN RENDAL — <span className="text-cyan-400">{user.nama}</span></span>
          </div>

          <div className="flex items-center gap-2">
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
              onClick={onLogout}
              className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer"
            >
              Keluar
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-6">
        {/* Kontrol Utama File Excel Master */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
          <div>
            <h2 className="text-lg font-bold text-white">Monitoring Seluruh Departemen & Biro</h2>
            <span className="text-xs text-slate-400">Akses lintas seluruh karyawan organik dan subkontraktor</span>
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

        {/* 1. Level Departemen */}
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

        {/* 2. Level Biro di dalam Departemen */}
        {selectedDept && !selectedBiroName && (
          <div className="space-y-4">
            <button onClick={() => setSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Departemen
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{selectedDept.name}</h2>
              <span className="text-xs text-slate-400">Pilih biro untuk melihat anggota</span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {selectedDept.biros.map((biro) => (
                <div
                  key={biro.id}
                  onClick={() => setSelectedBiroName(biro.name)}
                  className="bg-slate-900 border border-slate-800 hover:border-purple-500 rounded-xl p-4 cursor-pointer transition flex items-center justify-between"
                >
                  <span className="font-semibold text-white text-sm">{biro.name}</span>
                  <ChevronRight className="w-4 h-4 text-slate-500" />
                </div>
              ))}
            </div>
          </div>
        )}

        {/* 3. Level Daftar Anggota & Tugas di Biro Terpilih */}
        {selectedDept && selectedBiroName && (
          <div className="space-y-4">
            <button onClick={() => setSelectedBiroName(null)} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 rounded border border-slate-700 flex items-center gap-1 cursor-pointer">
              <ArrowLeft className="w-3.5 h-3.5" /> Kembali ke Daftar Biro
            </button>
            <div className="border-b border-slate-800 pb-2">
              <h2 className="text-base font-bold text-white">{selectedBiroName}</h2>
              <span className="text-xs text-slate-400">{currentBiroMembers.length} Anggota terdaftar di biro ini</span>
            </div>

            <div className="space-y-2.5">
              {currentBiroMembers.length > 0 ? (
                currentBiroMembers.map((person) => {
                  const personTasks = currentBiroTasks.filter(t => cleanText(t.pic) === cleanText(person.nama));
                  const isExpanded = !!expandedCards[person.nama];

                  return (
                    <div key={`${person.nama}-${person.nip}`} className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                      <div 
                        onClick={() => toggleAccordion(person.nama)}
                        className="p-3.5 flex items-center justify-between cursor-pointer hover:bg-slate-800/40 transition"
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-lg bg-purple-500/10 border border-purple-500/20 text-purple-400 flex items-center justify-center font-bold text-xs">
                            {person.nama.charAt(0)}
                          </div>
                          <div>
                            <div className="font-semibold text-white text-sm">
                              {person.nama} <span className="text-xs font-mono text-cyan-400 font-normal">({person.status})</span>
                            </div>
                            <div className="text-xs text-slate-400 font-mono">
                              NIP: {person.nip || '-'} • <span className="text-slate-400">{person.jabatan}</span>
                            </div>
                          </div>
                        </div>

                        <div className="flex items-center gap-3">
                          <span className="text-xs font-mono text-slate-400">{personTasks.length} Tugas</span>
                          <div className={`p-1 rounded bg-slate-950 border border-slate-800 text-slate-400 ${isExpanded ? 'text-purple-400' : ''}`}>
                            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                          </div>
                        </div>
                      </div>

                      {isExpanded && (
                        <div className="p-3.5 border-t border-slate-800 bg-slate-950/60">
                          {personTasks.length > 0 ? (
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs">
                                <thead>
                                  <tr className="text-slate-400 border-b border-slate-800">
                                    <th className="py-2 px-2.5">Jobcard</th>
                                    <th className="py-2 px-2.5">Proyek</th>
                                    <th className="py-2 px-2.5">Deskripsi</th>
                                    <th className="py-2 px-2.5 text-center">Rev</th>
                                    <th className="py-2 px-2.5 font-mono">Jadwal</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-800 text-slate-300">
                                  {personTasks.map(task => (
                                    <tr key={task.id} className="hover:bg-slate-900/40">
                                      <td className="py-2 px-2.5 font-mono font-bold text-amber-300">{task.kodeJc || 'Pending'}</td>
                                      <td className="py-2 px-2.5 text-emerald-400">{task.project}</td>
                                      <td className="py-2 px-2.5 text-slate-200">{task.taskName}</td>
                                      <td className="py-2 px-2.5 text-center font-mono">{task.rev || '0'}</td>
                                      <td className="py-2 px-2.5 font-mono text-[11px] text-slate-400">
                                        {formatDisplayDate(task.startDate)} s/d {formatDisplayDate(task.endDate)}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : (
                            <div className="text-xs text-slate-500 py-3 text-center italic">Belum ada tugas untuk personel ini.</div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })
              ) : (
                <div className="py-8 text-center text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-xl">
                  Tidak ada anggota di biro ini. Pastikan Master IM4 sudah di-update.
                </div>
              )}
            </div>
          </div>
        )}

        {/* Modal / Panel Planner (Tanpa Password untuk Admin Rendal) */}
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
                {Object.values(manualTasks).flatMap(tasks => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').length > 0 ? (
                  Object.values(manualTasks).flatMap(tasks => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').map((task, idx) => (
                    <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                      <div>
                        <div className="font-semibold text-white">{task.pic} <span className="font-mono text-slate-500">#{idx + 1}</span></div>
                        <div className="text-slate-400 text-[11px]">{task.taskName}</div>
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
                  <div className="py-8 text-center text-xs text-slate-500">Tidak ada pengajuan tugas yang menunggu approval.</div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}