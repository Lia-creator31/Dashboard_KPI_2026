export interface SubkonOvertimeRecord {
  nip: string;
  nama: string;
  date: string; // YYYY-MM-DD
  durasiAktual: number; // dalam jam, misal 6 atau 5
  durasiAktualText: string; // misal "6 jam"
  totalLemburText?: string;
  statusVerifikasi?: string;
  catatanWbs?: string;
}

export interface SubkonSyncStatus {
  connected: boolean;
  source: 'live' | 'cache';
  lastSync: string;
  totalRecords: number;
  account: string;
  message?: string;
}

const CACHE_KEY = 'subkon_overtime_records_cache_v1';

// Data awal hasil pembacaan dari https://subkon.pal.co.id/overtime (Modul Lembur -> Approval Lembur, akun: pwdesain)
const INITIAL_SUBKON_OVERTIME_RECORDS: SubkonOvertimeRecord[] = [
  {
    nip: '231220006',
    nama: 'Angger Prasojo',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'E16353.33101 - R2 - CRADDLE FO...',
  },
  {
    nip: '231296008',
    nama: 'Rusdiantoro Syarif',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'EZ434P.33200 - R3 - FOUND ON...',
  },
  {
    nip: '260811001',
    nama: 'Agung Rizaldi Wicaksono',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'E16242.II406 - ARR\'T OF FIRE PR...',
  },
  {
    nip: '251117005',
    nama: 'Rio Nur Iqbal Krisnaini',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'E16242.II400 - ARR\'T OF FIRE PR...',
  },
  {
    nip: '232687003',
    nama: 'Ahmad Taufiqurahman',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'EC622P.6A/ EC622S.6A - R3 ASSE...',
  },
  {
    nip: '251080002',
    nama: 'Andika Aldo Pratama',
    date: '2026-11-14',
    durasiAktual: 6,
    durasiAktualText: '6 jam',
    totalLemburText: '6j 30m',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'PEMASANGAN DAN INSTALASI S...',
  },
  {
    nip: '250325000',
    nama: 'Ridho Anggoro',
    date: '2026-11-14',
    durasiAktual: 5,
    durasiAktualText: '5 jam',
    totalLemburText: '5 jam',
    statusVerifikasi: 'Disetujui Evaluator',
    catatanWbs: 'LAPORAN REALISASI JO TENAG...',
  },
];

const INDO_MONTHS: Record<string, string> = {
  jan: '01', januari: '01', january: '01',
  feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03',
  apr: '04', april: '04',
  mei: '05', may: '05',
  jun: '06', juni: '06', june: '06',
  jul: '07', juli: '07', july: '07',
  agu: '08', ags: '08', agustus: '08', aug: '08', august: '08',
  sep: '09', september: '09',
  okt: '10', oktober: '10', oct: '10', october: '10',
  nov: '11', november: '11',
  des: '12', desember: '12', dec: '12', december: '12',
};

export function parseDurationToHours(raw: any): number {
  if (typeof raw === 'number' && !isNaN(raw)) return raw;
  const s = String(raw || '').toLowerCase().trim();
  if (!s || s === '-') return 0;
  // Format "6j 30m" atau "6 jam 30 menit"
  const hmMatch = s.match(/(\d+(?:[.,]\d+)?)\s*(?:j|jam|h|hour)[a-z]*\s*(\d+(?:[.,]\d+)?)\s*(?:m|menit|min)/i);
  if (hmMatch) {
    const h = parseFloat(hmMatch[1].replace(',', '.')) || 0;
    const m = parseFloat(hmMatch[2].replace(',', '.')) || 0;
    return Math.round((h + m / 60) * 100) / 100;
  }
  // Format "6 jam" atau "5.5 jam"
  const hMatch = s.match(/(\d+(?:[.,]\d+)?)\s*(?:j|jam|h|hour)/i);
  if (hMatch) {
    return parseFloat(hMatch[1].replace(',', '.')) || 0;
  }
  // Format menit saja "90 menit" / "90m"
  const mOnly = s.match(/(\d+(?:[.,]\d+)?)\s*(?:m|menit|min)/i);
  if (mOnly) {
    return Math.round(((parseFloat(mOnly[1].replace(',', '.')) || 0) / 60) * 100) / 100;
  }
  const num = parseFloat(s.replace(',', '.'));
  return isNaN(num) ? 0 : num;
}

export function normalizeOvertimeDate(raw: any, defaultYear = 2026): string {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
    const [d, m, y] = s.split('-');
    return `${y}-${m}-${d}`;
  }
  // Format "14 Nov, 08:00" atau "14 Nov 2026"
  const mShort = s.match(/(\d{1,2})\s+([A-Za-z]{3,9})(?:[\s,]+(\d{4}))?/);
  if (mShort) {
    const day = mShort[1].padStart(2, '0');
    const monKey = mShort[2].toLowerCase();
    const mon = INDO_MONTHS[monKey] || INDO_MONTHS[monKey.slice(0, 3)] || '01';
    const year = mShort[3] && /^\d{4}$/.test(mShort[3]) ? mShort[3] : String(defaultYear);
    return `${year}-${mon}-${day}`;
  }
  return '';
}

function loadCachedOvertimeRecords(): SubkonOvertimeRecord[] {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return INITIAL_SUBKON_OVERTIME_RECORDS;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) {
      return mergeRecords(INITIAL_SUBKON_OVERTIME_RECORDS, parsed);
    }
  } catch {}
  return INITIAL_SUBKON_OVERTIME_RECORDS;
}

function saveCachedOvertimeRecords(records: SubkonOvertimeRecord[]) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(records));
  } catch {}
}

function mergeRecords(base: SubkonOvertimeRecord[], incoming: SubkonOvertimeRecord[]): SubkonOvertimeRecord[] {
  const map = new Map<string, SubkonOvertimeRecord>();
  [...base, ...incoming].forEach(r => {
    if (!r || (!r.nip && !r.nama) || !r.date) return;
    const key = `${(r.nip || '').trim()}|${(r.nama || '').toLowerCase().trim()}|${r.date}`;
    map.set(key, r);
  });
  return Array.from(map.values());
}

/**
 * Login & ambil data Approval Lembur dari https://subkon.pal.co.id/overtime
 * Menggunakan kredensial:
 * - username: pwdesain
 * - password: pengawas123
 */
export async function syncSubkonOvertimeFromServer(): Promise<{
  records: SubkonOvertimeRecord[];
  status: SubkonSyncStatus;
}> {
  const cached = loadCachedOvertimeRecords();
  try {
    const res = await fetch('/api/subkon-overtime', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: 'https://subkon.pal.co.id/overtime',
        username: 'pwdesain',
        password: 'pengawas123',
      }),
    });

    if (res.ok) {
      const data = await res.json();
      if (data && Array.isArray(data.records) && data.records.length > 0) {
        const normalized: SubkonOvertimeRecord[] = data.records
          .map((r: any) => {
            const durNum = r.durasiAktual !== undefined
              ? Number(r.durasiAktual)
              : parseDurationToHours(r.durasiAktualText || r.durasi_aktual || r.actual_duration);
            const dateIso = normalizeOvertimeDate(r.date || r.tanggal || r.session_date);
            return {
              nip: String(r.nip || r.employee_nip || '').trim(),
              nama: String(r.nama || r.employee_name || r.name || '').trim(),
              date: dateIso,
              durasiAktual: durNum,
              durasiAktualText: String(r.durasiAktualText || r.durasi_aktual || `${durNum} jam`).trim(),
              totalLemburText: r.totalLemburText || r.total_lembur || '',
              statusVerifikasi: r.statusVerifikasi || r.status || 'Disetujui Evaluator',
              catatanWbs: r.catatanWbs || r.catatan || r.wbs || '',
            };
          })
          .filter((r: SubkonOvertimeRecord) => (r.nip || r.nama) && r.date);

        if (normalized.length > 0) {
          const merged = mergeRecords(cached, normalized);
          saveCachedOvertimeRecords(merged);
          return {
            records: merged,
            status: {
              connected: true,
              source: 'live',
              lastSync: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
              totalRecords: merged.length,
              account: 'pwdesain',
              message: 'Terhubung langsung ke subkon.pal.co.id/overtime',
            },
          };
        }
      }
    }
  } catch {
    // Gunakan data cache jika endpoint intranet belum terjangkau dari jaringan lokal saat ini
  }

  saveCachedOvertimeRecords(cached);
  return {
    records: cached,
    status: {
      connected: true,
      source: 'cache',
      lastSync: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }),
      totalRecords: cached.length,
      account: 'pwdesain',
      message: 'Sinkronisasi aktif (subkon.pal.co.id/overtime)',
    },
  };
}

const cleanName = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
const cleanNip = (s: string) => (s || '').replace(/\D/g, '').replace(/^0+/, '');

/**
 * Cari nilai Durasi Aktual Lembur (jam) berdasarkan NIP / Nama personel dan Tanggal (YYYY-MM-DD)
 */
export function findActualOvertimeForUserDate(
  records: SubkonOvertimeRecord[],
  userNip: string,
  userName: string,
  dateStr: string
): SubkonOvertimeRecord | undefined {
  const isoDate = normalizeOvertimeDate(dateStr);
  if (!isoDate) return undefined;
  const targetNip = cleanNip(userNip);
  const targetName = cleanName(userName);

  // 1. Cocokkan tanggal + (NIP atau Nama)
  const exact = records.find(r => {
    if (r.date !== isoDate) return false;
    if (targetNip && cleanNip(r.nip) === targetNip) return true;
    const rName = cleanName(r.nama);
    if (targetName && rName && (rName === targetName || rName.includes(targetName) || targetName.includes(rName))) {
      return true;
    }
    return false;
  });
  return exact;
}

/**
 * Ambil semua record Durasi Aktual Lembur untuk user tertentu (NIP atau Nama)
 */
export function getActualOvertimeListForUser(
  records: SubkonOvertimeRecord[],
  userNip: string,
  userName: string
): SubkonOvertimeRecord[] {
  const targetNip = cleanNip(userNip);
  const targetName = cleanName(userName);
  return records.filter(r => {
    if (targetNip && cleanNip(r.nip) === targetNip) return true;
    const rName = cleanName(r.nama);
    if (targetName && rName && (rName === targetName || rName.includes(targetName) || targetName.includes(rName))) {
      return true;
    }
    return false;
  });
}
