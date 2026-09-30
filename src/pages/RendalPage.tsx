import React, { useState, useEffect, useCallback, useMemo, useRef, ChangeEvent } from 'react';
import { departmentsData, Department } from '../data';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';
import { 
  Building2, Briefcase, HardHat, ArrowLeft, ChevronRight, ChevronDown, ChevronUp, 
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

interface SelectedFormPage {
  biroName: string;
  deptName: string;
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
  const b1 = (biro1 || '').toLowerCase().replace('&', ' dan ').trim();
  const b2 = (biro2 || '').toLowerCase().replace('&', ' dan ').trim();
  if (!b1 || !b2) return false;
  if (b1 === b2) return true;
  return cleanText(b1) === cleanText(b2);
}

function getBiroPrefix(biroName: string): string {
  const b = (biroName || '').toLowerCase().replace('&', ' dan ').trim();
  if (b.includes('dokumen')) return 'DP';
  if (b.includes('logistik')) return 'DL';
  if (b.includes('administrasi')) return 'DA';
  if (b.includes('pengembangan')) return 'PD';
  if (b.includes('non kapal')) return 'NK';
  if (b.includes('kapal selam')) return 'KS';
  if (b.includes('kapal permukaan')) return 'KP';
  return 'WO';
}

function cleanProjectString(raw: any): string {
  if (!raw) return '';
  return String(raw).replace(/[\u00A0\u200B\uFEFF\t\r\n\s]+/g, ' ').trim().toUpperCase();
}

function getProjectNormKey(s: string): string {
  return (s || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase().replace(/0/g, 'o');
}

async function fetchSafeWorkbook(paths: (string | undefined)[]): Promise<XLSX.WorkBook | null> {
  for (const p of paths) {
    if (!p) continue;
    try {
      const res = await fetch(p);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const bytes = new Uint8Array(buf.slice(0, 4));
        if (bytes[0] === 80 && bytes[1] === 75 && bytes[2] === 3 && bytes[3] === 4) {
          return XLSX.read(buf, { type: 'array' });
        }
      }
    } catch {}
  }
  return null;
}

export default function RendalPage({ user, onLogout }: RendalPageProps) {
  const [accessMode, setAccessMode] = useState<'landing' | 'organik' | 'subkon'>('landing');
  const [subconSelectedDept, setSubconSelectedDept] = useState<Department | null>(null);
  const [subconSelectedBiro, setSubconSelectedBiro] = useState<string | null>(null);
  const [subconPageMode, setSubconPageMode] = useState<'members' | 'form' | 'release'>('members');
  const [subconSearch, setSubconSearch] = useState('');

  const [selectedDept, setSelectedDept] = useState<Department | null>(null);
  const [selectedFormBiro, setSelectedFormBiro] = useState<SelectedFormPage | null>(null);
  const [formPageMode, setFormPageMode] = useState<'members' | 'form'>('members');
  const [searchQuery, setSearchQuery] = useState('');

  const [formData, setFormData] = useState({
    nama: '',
    kodeProyek: '',
    taskName: '',
    startDate: '',
    endDate: '',
    jo: '',
    rev: '0',
    release: '',
  });

  const [editingTask, setEditingTask] = useState<TaskItem | null>(null);
  const [editFormData, setEditFormData] = useState({
    project: '',
    taskName: '',
    startDate: '',
    endDate: '',
    jo: '',
    kodeJc: '',
    rev: '0',
    release: '',
  });

  const [editingReleaseId, setEditingReleaseId] = useState<string | null>(null);
  const [editingReleaseVal, setEditingReleaseVal] = useState<string>('');
  const [manualTasks, setManualTasks] = useState<{ [biroKey: string]: TaskItem[] }>({});

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

  const [isPlannerOpen, setIsPlannerOpen] = useState(false);
  const [editingTaskKode, setEditingTaskKode] = useState<{ [taskId: string]: string }>({});
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});

  const toggleAccordion = (cardKey: string) => {
    setExpandedCards(prev => ({ ...prev, [cardKey]: !prev[cardKey] }));
  };

  const [jobcardWorkbook, setJobcardWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [im4Workbook, setIm4Workbook] = useState<XLSX.WorkBook | null>(null);
  const [realisasiWorkbook, setRealisasiWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [drawingControlMap, setDrawingControlMap] = useState<Record<string, DrawingControlRow[]>>({});
  const [activeDrawingSheetTitle, setActiveDrawingSheetTitle] = useState<string>('');
  const [isFetchingDrawing, setIsFetchingDrawing] = useState<boolean>(false);

  useEffect(() => {
    async function loadAllExcelFiles() {
      try {
        const [wbJc, wbIm4, wbRealisasi] = await Promise.all([
          fetchSafeWorkbook(['/JOBCARD_DESAIN.xlsx', './JOBCARD_DESAIN.xlsx']),
          fetchSafeWorkbook(['/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx', './AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx']),
          fetchSafeWorkbook(['/Realisasi JO.xlsx', './Realisasi JO.xlsx'])
        ]);
        if (wbJc) setJobcardWorkbook(wbJc);
        if (wbIm4) setIm4Workbook(wbIm4);
        if (wbRealisasi) setRealisasiWorkbook(wbRealisasi);
      } catch {}
    }
    loadAllExcelFiles();
    loadAllJobCards();
  }, [loadAllJobCards]);

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
        alert('Gagal membaca file Excel Personel IM4.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleUpdateJobcardExcel = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        setJobcardWorkbook(wb);
        alert(`Katalog Jobcard "${file.name}" berhasil diperbarui!`);
      } catch {
        alert('Gagal membaca file JOBCARD_DESAIN.xlsx.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const fetchDrawingControlForProject = useCallback(async (projectCode: string) => {
    const cleanProj = cleanText(projectCode);
    if (!cleanProj) return;
    setIsFetchingDrawing(true);

    try {
      const url = `${GAS_DRAWING_API_URL}?project=${encodeURIComponent(projectCode.trim())}`;
      const res = await fetch(url);
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data) && json.data.length > 0) {
          const rowsList: DrawingControlRow[] = json.data.map((item: any) => ({
            noDwg: String(item.noDwg || '').trim(),
            drawingName: String(item.drawingName || '').trim(),
            fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
            rev: String(item.rev || '0').trim(),
            finishDate: String(item.finishDate || '').trim(),
          }));
          setDrawingControlMap(prev => ({ ...prev, [cleanProj]: rowsList }));
          setActiveDrawingSheetTitle(json.fileName || `Drawing Control ${projectCode}`);
          setIsFetchingDrawing(false);
          return;
        }
      }
    } catch {}
    setIsFetchingDrawing(false);
  }, []);

  const handleProjectChange = (newProject: string) => {
    setFormData(prev => ({ ...prev, kodeProyek: newProject, taskName: '', release: '' }));
    fetchDrawingControlForProject(newProject);
  };

  const handleDeskripsiChange = (selectedDesc: string) => {
    const cleanProj = cleanText(formData.kodeProyek || '');
    const rows = drawingControlMap[cleanProj] || [];
    let autoRelease = '';
    let autoRev = formData.rev || '0';

    if (rows.length > 0 && selectedDesc) {
      const cleanTarget = cleanText(selectedDesc);
      const currentRev = String(formData.rev || '0').trim();

      const exactMatch = rows.find(r => {
        const cFull = cleanText(r.fullDeskripsi);
        const cDwg = cleanText(r.noDwg);
        const cName = cleanText(r.drawingName);
        const rowRev = String(r.rev || '0').trim();
        return (cleanTarget.includes(cDwg) || cleanTarget.includes(cName) || cFull.includes(cleanTarget)) && rowRev === currentRev;
      });

      if (exactMatch && exactMatch.finishDate) {
        autoRelease = parseToStandardDate(exactMatch.finishDate);
        if (exactMatch.rev) autoRev = exactMatch.rev;
      }
    }
    setFormData(prev => ({ ...prev, taskName: selectedDesc, release: autoRelease, rev: autoRev }));
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

        const jabatan = String(row[jabatanCol] || '').trim();
        if (jabatan.toLowerCase().includes('kepala divisi')) {
          currentDept = 'Div. Desain';
          currentBiro = 'Div. Desain';
        } else if (jabatan.toLowerCase().includes('kepala departemen')) {
          currentDept = String(row[unitCol] || '').trim();
          currentBiro = `Staf ${currentDept}`;
        } else if (jabatan.toLowerCase().includes('kepala biro')) {
          currentBiro = jabatan.replace(/Kepala Biro/gi, 'Biro').trim();
        }

        results.push({
          nama,
          nip: String(row[nipCol] || '').trim(),
          status: String(row[statusCol] || 'PKWTT').trim(),
          jabatan,
          biro: currentBiro || 'Biro Umum',
          dept: currentDept || 'Div. Desain'
        });
      }
      return results;
    } catch {
      return [];
    }
  }, [im4Workbook]);

  const dynamicOutsourcingList = useMemo(() => {
    return allParsedFromExcel.filter(p => p.status.toLowerCase().includes('outsourcing'));
  }, [allParsedFromExcel]);

  const getSubconMembersForBiro = useCallback((biroName: string) => {
    return dynamicOutsourcingList.filter(os => isBiroMatch(os.biro, biroName));
  }, [dynamicOutsourcingList]);

  const getSubconCountForDept = useCallback((deptName: string): number => {
    try {
      const dept = departmentsData.find(d => d.name === deptName);
      if (!dept) return 0;
      const uniq = new Set<string>();
      dept.biros.forEach(biro => {
        getSubconMembersForBiro(biro.name).forEach(p => uniq.add(p.nip || p.nama));
      });
      return uniq.size;
    } catch {
      return 0;
    }
  }, [getSubconMembersForBiro]);

  const getBiroMembers = useCallback((biroName: string) => {
    return allParsedFromExcel.filter(p => isBiroMatch(p.biro, biroName) && !p.status.toLowerCase().includes('outsourcing'));
  }, [allParsedFromExcel]);

  const currentActiveBiroName = useMemo(() => {
    if (accessMode === 'subkon') return subconSelectedBiro || '';
    return selectedFormBiro?.biroName || '';
  }, [accessMode, subconSelectedBiro, selectedFormBiro]);

  const currentActiveBiroKey = cleanText(currentActiveBiroName);
  const currentActiveBiroTasks = manualTasks[currentActiveBiroKey] || [];

  const activeSubconMembers = useMemo(() => {
    if (!subconSelectedBiro) return [];
    return getSubconMembersForBiro(subconSelectedBiro);
  }, [subconSelectedBiro, getSubconMembersForBiro]);

  const subconWorkOrders = useMemo(() => {
    if (!subconSelectedBiro) return [];
    const prefix = getBiroPrefix(subconSelectedBiro);
    const subconNames = getSubconMembersForBiro(subconSelectedBiro).map(m => cleanText(m.nama));
    const filteredTasks = currentActiveBiroTasks.filter(t => subconNames.includes(cleanText(t.pic)));
    return filteredTasks.map((task, idx) => ({ ...task, packageTitle: `${prefix}${idx + 1}` }));
  }, [subconSelectedBiro, currentActiveBiroTasks, getSubconMembersForBiro]);

  // Hitung jumlah tugas yang menumpuk di planner
  const pendingTasksCount = useMemo(() => {
    let count = 0;
    Object.values(manualTasks).forEach(tasks => {
      tasks.forEach(t => {
        if (!t.kodeJc || t.kodeJc.trim() === '') count++;
      });
    });
    return count;
  }, [manualTasks]);

  const projectOptions = useMemo((): string[] => {
    const projectMap = new Map<string, string>();
    GOOGLE_DRIVE_SHEETS.forEach(s => projectMap.set(getProjectNormKey(s.projectKey), cleanProjectString(s.projectKey)));
    return Array.from(projectMap.values());
  }, []);

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    const activeBiro = currentActiveBiroName;
    if (!activeBiro) return;

    try {
      const { data: biroList } = await supabase.from('biros').select('id, name');
      let validBiroId = biroList?.[0]?.id || null;
      if (biroList) {
        const found = biroList.find(b => isBiroMatch(b.name, activeBiro));
        if (found) validBiroId = found.id;
      }

      const autoKode = accessMode === 'subkon' ? `${getBiroPrefix(activeBiro)}${(manualTasks[cleanText(activeBiro)] || []).length + 1}` : '';
      const revVal = formData.rev || '0';

      const { data: insertedRow, error } = await supabase.from('job_cards').insert({
        biro_id: validBiroId,
        biro_name: activeBiro,
        personil_name: formData.nama,
        project_code: formData.kodeProyek,
        project: formData.kodeProyek,
        task_name: formData.taskName,
        start_date: formData.startDate,
        end_date: formData.endDate,
        pic: formData.nama,
        jo: formData.jo,
        kode_jc: autoKode,
        status: accessMode === 'subkon' ? 'approved' : 'pending',
        rev: revVal,
        release: formData.release,
      }).select().single();

      if (error) {
        alert('Gagal menyimpan tugas: ' + error.message);
        return;
      }

      alert('Tugas berhasil disimpan!');
      setFormData({ nama: '', kodeProyek: '', taskName: '', startDate: '', endDate: '', jo: '', rev: '0', release: '' });
      loadAllJobCards();
    } catch {
      alert('Koneksi database bermasalah.');
    }
  };

  const handleSaveKodeJcForTask = async (taskId: string) => {
    const inputVal = (editingTaskKode[taskId] || '').trim().toUpperCase();
    if (!inputVal) return;
    const { error } = await supabase.from('job_cards').update({ kode_jc: inputVal, status: 'approved' }).eq('id', taskId);
    if (error) {
      alert('Gagal: ' + error.message);
      return;
    }
    alert('Jobcard berhasil disetujui!');
    loadAllJobCards();
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen font-sans">
      <header className="sticky top-0 z-30 bg-slate-900/90 backdrop-blur border-b border-slate-800">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2.5 cursor-pointer" onClick={() => { setAccessMode('landing'); setSelectedFormBiro(null); setSelectedDept(null); setSubconSelectedBiro(null); setSubconSelectedDept(null); }}>
            <div className="p-1.5 bg-blue-600 rounded-lg text-white"><Building2 className="w-4 h-4" /></div>
            <span className="font-bold text-sm text-white">DIVISI DESAIN — <span className="text-cyan-400">{user.nama} ({user.role.toUpperCase()})</span></span>
          </div>

          <div className="flex items-center gap-2">
            {user.role === 'admin' && (
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
            )}

            {accessMode !== 'landing' && (
              <button onClick={() => { setAccessMode('landing'); setSelectedFormBiro(null); setSelectedDept(null); setSubconSelectedBiro(null); setSubconSelectedDept(null); }} className="px-3 py-1 text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg border border-slate-700 cursor-pointer">
                Ganti Portal
              </button>
            )}

            <button onClick={onLogout} className="px-3 py-1 text-xs bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 rounded-lg border border-rose-500/30 cursor-pointer">
              Keluar
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        {/* ================= 1. LANDING PORTAL ================= */}
        {accessMode === 'landing' && (
          <div className="max-w-2xl mx-auto text-center space-y-8 pt-8">
            <div className="space-y-2">
              <h1 className="text-2xl sm:text-3xl font-bold text-white">Sistem Penugasan Job Card</h1>
              <p className="text-slate-400 text-sm">Pilih portal akses kerja Anda</p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div onClick={() => setAccessMode('organik')} className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-2xl p-6 cursor-pointer text-left transition group">
                <div className="p-3 bg-blue-500/10 text-blue-400 rounded-xl w-fit mb-4 group-hover:bg-blue-600 group-hover:text-white transition"><Briefcase className="w-6 h-6" /></div>
                <h3 className="text-lg font-bold text-white">Pegawai Organik</h3>
                <span className="text-xs text-slate-400">Pegawai PKWTT & PKWT Divisi Desain</span>
              </div>

              <div onClick={() => setAccessMode('subkon')} className="bg-slate-900 border border-slate-800 hover:border-amber-500 rounded-2xl p-6 cursor-pointer text-left transition group">
                <div className="p-3 bg-amber-500/10 text-amber-400 rounded-xl w-fit mb-4 group-hover:bg-amber-600 group-hover:text-white transition"><HardHat className="w-6 h-6" /></div>
                <h3 className="text-lg font-bold text-white">Mitra / Subkon</h3>
                <span className="text-xs text-slate-400">Personil Outsourcing</span>
              </div>
            </div>
          </div>
        )}

        {/* ================= 2. PORTAL SUBKON ================= */}
        {accessMode === 'subkon' && (
          <div className="space-y-6">
            {!subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
                  <h2 className="text-lg font-bold text-white">Pilih Departemen</h2>
                  {user.role === 'admin' && (
                    <div className="flex gap-2">
                      <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                        <Users className="w-3.5 h-3.5 text-amber-400" /><span>Update IM4</span>
                        <input type="file" accept=".xlsx" onChange={handleUpdateIm4Excel} className="hidden" />
                      </label>
                      <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                        <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-400" /><span>Update Jobcard</span>
                        <input type="file" accept=".xlsx" onChange={handleUpdateJobcardExcel} className="hidden" />
                      </label>
                    </div>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                  {departmentsData.map((dept) => (
                    <div key={dept.id} onClick={() => setSubconSelectedDept(dept)} className="bg-slate-900 border border-slate-800 hover:border-amber-500 rounded-xl p-4 cursor-pointer flex items-center justify-between">
                      <div><h4 className="font-semibold text-white text-sm">{dept.name}</h4><span className="text-xs text-slate-400">{dept.biros.length} Biro</span></div>
                      <span className="px-2 py-0.5 bg-amber-500/10 text-amber-400 rounded font-bold text-xs">{getSubconCountForDept(dept.name)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {subconSelectedDept && !subconSelectedBiro && (
              <div className="space-y-4">
                <button onClick={() => setSubconSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 rounded border border-slate-700 flex items-center gap-1"><ArrowLeft className="w-3 h-3" /> Kembali</button>
                <div className="space-y-2">
                  {subconSelectedDept.biros.map((biro) => (
                    <div key={biro.id} className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex items-center justify-between">
                      <span className="font-semibold text-white text-sm">{biro.name}</span>
                      <div className="flex gap-1.5">
                        <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('members'); }} className="px-2.5 py-1 bg-slate-800 text-xs rounded border border-slate-700">Anggota</button>
                        <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('form'); }} className="px-2.5 py-1 bg-emerald-600 text-white text-xs rounded">Form</button>
                        <button onClick={() => { setSubconSelectedBiro(biro.name); setSubconPageMode('release'); }} className="px-2.5 py-1 bg-purple-600 text-white text-xs rounded">Work Order</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {subconSelectedBiro && (
              <div className="space-y-4">
                <button onClick={() => setSubconSelectedBiro(null)} className="px-3 py-1 text-xs bg-slate-800 rounded border border-slate-700 flex items-center gap-1"><ArrowLeft className="w-3 h-3" /> Kembali</button>
                {subconPageMode === 'form' && (
                  <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                    <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Nama Personel Outsourcing</label>
                        <SearchableSelect options={activeSubconMembers.map(p => p.nama)} value={formData.nama} onChange={(val) => setFormData(p => ({ ...p, nama: val }))} required />
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Kode Proyek</label>
                        <SearchableSelect options={projectOptions} value={formData.kodeProyek} onChange={(val) => handleProjectChange(val)} required />
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Nomor JO</label>
                        <input type="text" value={formData.jo} onChange={(e) => setFormData(p => ({ ...p, jo: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Deskripsi</label>
                        <SearchableSelect options={dynamicTaskOptions} value={formData.taskName} onChange={(val) => handleDeskripsiChange(val)} required />
                      </div>
                      <div className="grid grid-cols-2 gap-3">
                        <div><label className="block text-slate-400 mb-1">Tanggal Mulai</label><input type="date" value={formData.startDate} onChange={(e) => setFormData(p => ({ ...p, startDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" /></div>
                        <div><label className="block text-slate-400 mb-1">Tanggal Selesai</label><input type="date" value={formData.endDate} onChange={(e) => setFormData(p => ({ ...p, endDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" /></div>
                      </div>
                      <button type="submit" className="px-5 py-2 bg-amber-600 text-white font-bold rounded-lg cursor-pointer">Simpan Tugas</button>
                    </form>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ================= 3. PORTAL ORGANIK ================= */}
        {accessMode === 'organik' && (
          <div className="space-y-6">
            {!selectedDept && !selectedFormBiro && (
              <div className="space-y-5">
                <div className="flex items-center justify-between border-b border-slate-800 pb-3 flex-wrap gap-2">
                  <h2 className="text-lg font-bold text-white">Departemen Desain</h2>
                  {user.role === 'admin' && (
                    <label className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-lg border border-slate-700 flex items-center gap-1 cursor-pointer">
                      <Clock className="w-3.5 h-3.5 text-emerald-400" /><span>Upload Realisasi JO</span>
                      <input type="file" accept=".xlsx" onChange={handleManualUploadRealisasi} className="hidden" />
                    </label>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                  {departmentsData.map((dept) => (
                    <div key={dept.id} onClick={() => setSelectedDept(dept)} className="bg-slate-900 border border-slate-800 hover:border-blue-500 rounded-xl p-4 cursor-pointer flex items-center justify-between">
                      <div><h4 className="font-semibold text-white text-sm">{dept.name}</h4><span className="text-xs text-slate-400">{dept.biros.length} Biro</span></div>
                      <ChevronRight className="w-4 h-4 text-slate-500" />
                    </div>
                  ))}
                </div>
              </div>
            )}

            {selectedDept && !selectedFormBiro && (
              <div className="space-y-4">
                <button onClick={() => setSelectedDept(null)} className="px-3 py-1 text-xs bg-slate-800 rounded border border-slate-700 flex items-center gap-1"><ArrowLeft className="w-3 h-3" /> Kembali</button>
                <div className="space-y-2">
                  {selectedDept.biros.map((biro) => (
                    <div key={biro.id} className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 flex items-center justify-between">
                      <span className="font-semibold text-white text-sm">{biro.name}</span>
                      <div className="flex gap-1.5">
                        <button onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('members'); }} className="px-2.5 py-1 bg-blue-600 text-white text-xs rounded">Anggota</button>
                        <button onClick={() => { setSelectedFormBiro({ biroName: biro.name, deptName: selectedDept.name }); setFormPageMode('form'); }} className="px-2.5 py-1 bg-emerald-600 text-white text-xs rounded">Form</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {selectedFormBiro && formPageMode === 'form' && (
              <div className="space-y-4">
                <button onClick={() => setSelectedFormBiro(null)} className="px-3 py-1 text-xs bg-slate-800 rounded border border-slate-700 flex items-center gap-1"><ArrowLeft className="w-3 h-3" /> Kembali</button>
                <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                  <form onSubmit={handleSubmitForm} className="space-y-4 text-xs">
                    <div>
                      <label className="block text-slate-400 mb-1">Nama Personel Organik</label>
                      <SearchableSelect options={getBiroMembers(selectedFormBiro.biroName).map(p => p.nama)} value={formData.nama} onChange={(val) => setFormData(p => ({ ...p, nama: val }))} required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Kode Proyek</label>
                      <SearchableSelect options={projectOptions} value={formData.kodeProyek} onChange={(val) => handleProjectChange(val)} required />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Plan JO</label>
                      <input type="text" value={formData.jo} onChange={(e) => setFormData(p => ({ ...p, jo: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono" />
                    </div>
                    <div>
                      <label className="block text-slate-400 mb-1">Deskripsi</label>
                      <SearchableSelect options={dynamicTaskOptions} value={formData.taskName} onChange={(val) => handleDeskripsiChange(val)} required />
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div><label className="block text-slate-400 mb-1">Plan Start</label><input type="date" value={formData.startDate} onChange={(e) => setFormData(p => ({ ...p, startDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" /></div>
                      <div><label className="block text-slate-400 mb-1">Plan Finish</label><input type="date" value={formData.endDate} onChange={(e) => setFormData(p => ({ ...p, endDate: e.target.value }))} required className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-white" /></div>
                    </div>
                    <button type="submit" className="px-5 py-2 bg-emerald-600 text-white font-bold rounded-lg cursor-pointer">Simpan Tugas</button>
                  </form>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= MODAL PLANNER (ADMIN RENDAL) ================= */}
        {isPlannerOpen && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl">
              <div className="p-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <span className="font-bold text-xs text-white flex items-center gap-1.5">
                  <Lock className="w-3.5 h-3.5 text-amber-400" /> Planner Panel — Approval Jobcard ({pendingTasksCount} Menunggu)
                </span>
                <button onClick={() => setIsPlannerOpen(false)} className="text-slate-400 hover:text-white cursor-pointer"><X className="w-4 h-4" /></button>
              </div>

              <div className="p-4 max-h-96 overflow-y-auto space-y-2">
                {Object.entries(manualTasks).flatMap(([_, tasks]) => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').length > 0 ? (
                  Object.entries(manualTasks).flatMap(([_, tasks]) => tasks).filter(t => !t.kodeJc || t.kodeJc.trim() === '').map((task) => (
                    <div key={task.id} className="p-3 bg-slate-950 border border-slate-800 rounded-lg flex items-center justify-between gap-3 text-xs">
                      <div>
                        <div className="font-semibold text-white">{task.pic}</div>
                        <div className="text-slate-400 text-[11px]">{task.taskName}</div>
                        <div className="text-emerald-400 font-mono text-[10px]">{task.project} • {task.biroName}</div>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="text"
                          value={editingTaskKode[task.id] ?? task.kodeJc ?? ''}
                          onChange={(e) => setEditingTaskKode(prev => ({ ...prev, [task.id]: e.target.value.toUpperCase() }))}
                          placeholder="No Jobcard..."
                          className="w-36 px-2 py-1 bg-slate-900 border border-slate-700 rounded text-xs font-mono text-white uppercase"
                        />
                        <button onClick={() => handleSaveKodeJcForTask(task.id)} className="px-2.5 py-1 bg-amber-600 text-white font-bold rounded text-xs">Simpan</button>
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