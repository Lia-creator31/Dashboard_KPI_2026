import * as XLSX from 'xlsx';

// Membuat daftar nama kandidat dari satu teks proyek, supaya Drawing Control di Drive
// ditemukan walau namanya hanya cocok sebagian.
//
// "WFFBNW - W000304" -> ["WFFBNW", "FFBNW", "W000304"]
// "WFFBNW2"          -> ["WFFBNW2", "WFFBNW", "FFBNW"]
// "W0000304 - 305"   -> ["W0000304", "W000304", "W000305"]
export function projectCandidates(raw: string): string[] {
  const out = new Set<string>();
  const add = (s: string) => { const v = s.trim(); if (v) out.add(v); };

  // Jika ada kode berstrip seperti "R41-PCC" atau "R41-FPB", tambahkan juga gabungan utuhnya ("R41PCC", "R41-PCC")
  const hyphenCodes = (raw || '').toUpperCase().match(/\b[A-Z]\d{1,3}-[A-Z0-9]{2,6}\b/g) || [];
  hyphenCodes.forEach(hc => {
    add(hc);
    add(hc.replace(/-/g, ''));
  });

  const tokens = (raw || '').toUpperCase()
    .split(/\s*(?:-|–|—|\/|,|;|&|\+|\(|\)|\bDAN\b)\s*/)
    .map(t => t.replace(/\.XLSX$|\.XLS$/i, '').replace(/[^A-Z0-9]/g, ''))
    .filter(Boolean);

  let prefix = ''; // huruf/awalan kode sebelumnya, untuk "W000304 - 305" atau "JW1 - JW2"
  for (const t of tokens) {
    // angka saja (mis. "305") -> pakai huruf dari kode sebelumnya
    if (/^\d+$/.test(t)) {
      if (prefix) add(prefix + t.padStart(6, '0'));
      continue;
    }
    add(t);

    // kode A26AOP01 / A26A0P01 -> tambahkan kedua variasi (O dan 0)
    if (/^A\d{2}A[O0]P\d{2}$/.test(t)) {
      add(t.replace(/A0P/, 'AOP'));
      add(t.replace(/AOP/, 'A0P'));
      add('KSOT');
    }

    // JW1 / JW2 -> berkaitan dengan LCU
    if (/^JW\d+$/.test(t)) {
      add('LCU');
    }

    // kode standar: 1 huruf + angka -> W000304
    const code = t.match(/^([A-Z])0*(\d+)$/);
    if (code) {
      prefix = code[1];
      if (code[2].length <= 6) add(code[1] + code[2].padStart(6, '0'));
      continue;
    }

    // nama huruf + 1-2 angka di belakang (WFFBNW1, WFFBNW2) -> WFFBNW
    const fam = t.match(/^([A-Z]{3,})(\d{1,2})$/);
    const base = fam ? fam[1] : (/^[A-Z]{3,}$/.test(t) ? t : '');
    if (base) {
      add(base);
      // WFFBNW -> FFBNW (atau sebaliknya FFBNW -> WFFBNW)
      if (base.startsWith('W') && base.length >= 5) add(base.slice(1));
      if (base === 'FFBNW') add('WFFBNW');
    }
  }
  return Array.from(out);
}

export interface DrawingSheetConfig {
  id: string;
  projectKey: string;
  title: string;
  aliases?: string[]; // kode proyek di dalam kurung atau nama alternatif pada judul file di Drive
}

// SATU-SATUNYA tempat daftar Link Google Sheet Drawing Control untuk seluruh aplikasi.
// Sesuai isi folder Google Drive "DRAWING CONTROL ONLINE":
// - Drawing Control KSR 20 M ALU (W000314)
// - Drawing Control FRIGATE 140 M (W000304-305)
// - Drawing Control KSOT (A26AOP01)
// - Drawing Control LPD UAE USED (W000308)
// - Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312)
// - Drawing Control LD PN 124 M. (W000306-307)
// - Drawing Control GEOMARIN V (M000313)
// - Drawing Control KSSR 18 M (W000311)
// - Drawing Control LCU LD PN 124 M (JW1 - JW2)
// - Drawing Control FFBNW FRIGATE 140 M (WFFBNW)
// - Drawing Control R41-PCC.xlsx
// - Drawing Control R41-FPB.xlsx
export const GOOGLE_DRIVE_SHEETS: DrawingSheetConfig[] = [
  { id: '19qiMBb1b7qmeRXkFNTbh-eKtVEKDUTgshloa116HQ48', projectKey: 'W000308', title: 'Drawing Control LPD UAE USED (W000308)', aliases: ['LPD UAE'] },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000304', title: 'Drawing Control FRIGATE 140 M (W000304-305)', aliases: ['W000304-305'] },
  { id: '1OLxSphh-jiqmUGIHE9QSXkRCVfwoVhkRsyP3InRzAgk', projectKey: 'W000305', title: 'Drawing Control FRIGATE 140 M (W000304-305)', aliases: ['W000304-305'] },
  { id: '1CBL96MejQnswfg_sLIkwkK9K594S9hDWoSGDBIeIBNo', projectKey: 'M000313', title: 'Drawing Control GEOMARIN V (M000313)', aliases: ['GEOMARIN V', 'GEOMARIN'] },
  { id: '16rrjyU60BmzRIYFL3I1WRLkw_AfJPUwi8vZiJx2LYCg', projectKey: 'W000314', title: 'Drawing Control KSR 20 M ALU (W000314)', aliases: ['KSR 20 M ALU', 'KSR 20 M'] },
  { id: '1l0ZoTCctHrmkruIkC4Gk1NtoqsugMVGABrsf3yb6PHo', projectKey: 'KSOT', title: 'Drawing Control KSOT (A26AOP01)', aliases: ['A26AOP01', 'A26A0P01'] },
  { id: '1uv2D4ECJ0z-NWP8byzXTkkM17MKD6f0r-RCtKjpG5Hs', projectKey: 'M000312', title: 'Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312)' },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000306', title: 'Drawing Control LD PN 124 M. (W000306-307)', aliases: ['W000306-307'] },
  { id: '1ea2z7gEfavWoxZp9nBTc19A2PSq_GjkXP11tKvpAMfo', projectKey: 'W000307', title: 'Drawing Control LD PN 124 M. (W000306-307)', aliases: ['W000306-307'] },
  { id: '1eTO-dvVd3psdP4mFQc2ZQW_Z-lyooCb2TVjR4mYTbDw', projectKey: 'W000311', title: 'Drawing Control KSSR 18 M (W000311)', aliases: ['KSSR 18 M'] },
  { id: '1jGNHh9HhSrPlEv7JZvtW8ToMU2ADrT0yE87tk9E53zY', projectKey: 'LCU', title: 'Drawing Control LCU LD PN 124 M (JW1 - JW2)', aliases: ['JW1', 'JW2', 'JW1 - JW2', 'LCU LD PN 124 M'] },
  { id: '1XsrklFhcUkjgfEjuT4jDRWx6Uufsj8Jt', projectKey: 'FFBNW', title: 'Drawing Control FFBNW FRIGATE 140 M (WFFBNW)', aliases: ['WFFBNW', 'WFFBNW1', 'WFFBNW2'] },
  { id: '', projectKey: 'R41-PCC', title: 'Drawing Control R41-PCC.xlsx', aliases: ['R41PCC', 'PCC'] },
  { id: '', projectKey: 'R41-FPB', title: 'Drawing Control R41-FPB.xlsx', aliases: ['R41FPB', 'FPB'] },
];

// Ekstrak seluruh kode proyek & kode di dalam tanda kurung (...) dari judul file di Drive
function extractKeysFromDriveSheets(): string[] {
  const keys = new Set<string>();
  GOOGLE_DRIVE_SHEETS.forEach(s => {
    const pk = s.projectKey.toUpperCase().trim();
    if (pk) {
      keys.add(pk);
      keys.add(pk.replace(/[^A-Z0-9]/g, ''));
    }
    (s.aliases || []).forEach(a => {
      const up = a.toUpperCase().trim();
      if (up) {
        keys.add(up);
        keys.add(up.replace(/[^A-Z0-9]/g, ''));
      }
    });
    // Ambil juga apa pun yang tertulis di dalam tanda kurung pada title, misal "(W000304-305)", "(A26AOP01)", "(JW1 - JW2)", "(WFFBNW)"
    const inParens = s.title.match(/\(([^)]+)\)/g) || [];
    inParens.forEach(p => {
      projectCandidates(p.replace(/[()]/g, '')).forEach(c => keys.add(c));
    });
  });
  return Array.from(keys);
}

// Otomatis terbentuk dari GOOGLE_DRIVE_SHEETS beserta seluruh kode di dalam kurung & aliasnya
export const KNOWN_DRAWING_PROJECT_KEYS: string[] = extractKeysFromDriveSheets();

// Jika proyek adalah CORP2026 (atau proyek non-link lainnya), cek di dalam deskripsi
// apakah ada kode proyek yang terdaftar di link Drawing Control (mis. "W000314 - E25291.01100...")
// ATAU format nama file Drawing Control di Drive:
//   - Kode proyek standar: W000314, M000313, W000304-305
//   - Kode alfanumerik / kontrak di dalam kurung: A26AOP01, JW1, JW2, WFFBNW
//   - Kode berstrip: R41-PCC, R41-FPB
export function findDrawingProjectInDescription(desc: string): string {
  if (!desc) return '';
  const upper = desc.toUpperCase().trim();

  // 0) Cek kode berstrip seperti "R41-PCC" atau "R41-FPB" di dalam deskripsi
  const mHyphen = upper.match(/\b([A-Z]\d{1,3}-[A-Z0-9]{2,6})\b/);
  if (mHyphen) return mHyphen[1];

  // 1) Cek kandidat proyek dari token-token deskripsi terhadap daftar link Drawing Control
  const cands = projectCandidates(upper);
  for (const c of cands) {
    const match = KNOWN_DRAWING_PROJECT_KEYS.find(k => k === c || k.replace(/[^A-Z0-9]/g, '') === c);
    if (match) return match;
  }
  // 2) Cek kecocokan kata utuh untuk tiap projectKey/alias di link Drawing Control
  for (const key of KNOWN_DRAWING_PROJECT_KEYS) {
    if (key.length < 3) continue;
    const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`(?:^|[^A-Z0-9])${escaped}(?:$|[^A-Z0-9])`, 'i');
    if (re.test(upper)) return key;
  }
  // 3) Dinamis (otomatis untuk kapal/proyek baru yang ditambahkan di folder Google Drive "DRAWING CONTROL ONLINE"):
  //    a) Format W000xxx / M000xxx di awal atau di tengah deskripsi
  const mLeadCode = upper.match(/^([WM]\d{4,6})\s*[-–—_:/]/);
  if (mLeadCode) {
    const norm = projectCandidates(mLeadCode[1]);
    return norm.find(x => /^[WM]\d{6}$/.test(x)) || mLeadCode[1];
  }
  const mAnyCode = upper.match(/(?:^|[^A-Z0-9.])([WM]0{2,4}\d{2,4})(?![A-Z0-9.])/);
  if (mAnyCode) {
    const norm = projectCandidates(mAnyCode[1]);
    return norm.find(x => /^[WM]\d{6}$/.test(x)) || mAnyCode[1];
  }
  //    b) Format kode proyek di dalam kurung pada judul Drive, seperti A26AOP01, JW1, JW2, R41-PCC, atau kode proyek di depan sebelum " - NOMOR_GAMBAR"
  const mLeadPrefix = upper.match(/^([A-Z0-9]{2,12}(?:-[A-Z0-9]{2,8})?)\s*[-–—_:/]\s*([A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+)/);
  if (mLeadPrefix && !/^R(?:EV)?\d*$/i.test(mLeadPrefix[1])) {
    return mLeadPrefix[1];
  }
  //    c) Jika ada teks "Drawing Control <NAMA PROYEK> (<KODE>)" langsung di deskripsi
  const mDcTitle = upper.match(/DRAWING\s+CONTROL\s+([A-Z0-9\s.&-]+?)(?:\s*\(([A-Z0-9\s-]+)\)|\.XLSX?|$)/i);
  if (mDcTitle) {
    if (mDcTitle[2]) {
      const c = projectCandidates(mDcTitle[2]);
      if (c.length > 0) return c[0];
    }
    return mDcTitle[1].trim();
  }
  return '';
}

export function resolveDrawingProject(project: string, desc: string): string {
  const p = (project || '').trim();
  const pClean = p.toUpperCase().replace(/[^A-Z0-9]/g, '');
  // Jika proyek adalah CORP (mis. CORP2026) atau tidak ada di link Drawing Control,
  // cek apakah ada nama/kode proyek Drawing Control di dalam deskripsi.
  const isKnown = projectCandidates(p).some(c => KNOWN_DRAWING_PROJECT_KEYS.includes(c));
  if (pClean.startsWith('CORP') || !isKnown) {
    const fromDesc = findDrawingProjectInDescription(desc);
    if (fromDesc) return fromDesc;
  }
  return p;
}

// Pilihan Kode Proyek untuk Form:
// Persis mengikuti urutan dan teks di dalam tanda kurung (...) pada folder Google Drive "DRAWING CONTROL ONLINE":
// 1. Drawing Control LPD UAE USED (W000308)                 -> W000308
// 2. Drawing Control FRIGATE 140 M (W000304-305)            -> W000304-305
// 3. Drawing Control GEOMARIN V (M000313)                   -> M000313
// 4. Drawing Control KSR 20 M ALU (W000314)                 -> W000314
// 5. Drawing Control KSOT (A26AOP01)                        -> A26AOP01
// 6. Drawing Control PENGAWASAN & PELAYANAN PULAU (M000312) -> M000312
// 7. Drawing Control LD PN 124 M. (W000306-307)             -> W000306-307
// 8. Drawing Control KSSR 18 M (W000311)                    -> W000311
// 9. Drawing Control LCU LD PN 124 M (JW1 - JW2)            -> JW1 - JW2
// 10. Drawing Control FFBNW FRIGATE 140 M (WFFBNW)          -> WFFBNW
// (File R41-PCC.xlsx & R41-FPB.xlsx tidak memiliki tanda kurung kode proyek)
export function getFormDriveProjectOptions(): string[] {
  const out: string[] = [];
  const add = (v: string) => {
    const clean = v.trim();
    if (clean && !out.includes(clean)) out.push(clean);
  };
  GOOGLE_DRIVE_SHEETS.forEach(s => {
    const m = s.title.match(/\(([^)]+)\)/);
    if (m) {
      add(m[1].trim());
    }
  });
  return out;
}

// Daftar projectKey unik dari seluruh proyek Drawing Control di Google Drive,
// dipakai untuk mencari nomor gambar satu per satu di semua proyek ketika proyeknya CORP2026
// tetapi di dalam deskripsi tidak tertulis kode proyeknya (mis. "E90961.92009 - LIGHTWEIGHT ESTIMATION").
export const ALL_DRIVE_PROJECT_KEYS: string[] = Array.from(
  new Set(GOOGLE_DRIVE_SHEETS.map(s => s.projectKey.toUpperCase().trim()).filter(Boolean))
);

export function isCorpProject(project: string): boolean {
  return (project || '').toUpperCase().replace(/[^A-Z0-9]/g, '').startsWith('CORP');
}

const cleanAlnum = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();

// Cari konfigurasi sheet Google Drive berdasarkan projectKey, alias (mis. A26AOP01, JW1, JW2, WFFBNW), atau judul
export function findMatchingDriveSheet(projectCode: string): DrawingSheetConfig | undefined {
  const cp = cleanAlnum(projectCode);
  if (!cp) return undefined;
  return GOOGLE_DRIVE_SHEETS.find(s => {
    if (!s.id) return false;
    const k = cleanAlnum(s.projectKey);
    if (k === cp || cp.includes(k)) return true;
    if ((s.aliases || []).some(a => { const ca = cleanAlnum(a); return ca && (ca === cp || cp.includes(ca)); })) return true;
    return false;
  });
}

export const dedupeRows = <T extends { noDwg: string; drawingName?: string; rev: string; finishDate: string; planStart?: string; planFinish?: string; jo?: string }>(rows: T[]): T[] => {
  const map = new Map<string, T>();
  const scheduleByDwg = new Map<string, { planStart?: string; planFinish?: string; jo?: string }>();

  for (const r of rows) {
    const dwgKey = cleanAlnum(r.noDwg) || cleanAlnum(r.drawingName || '');
    if (dwgKey) {
      const prev = scheduleByDwg.get(dwgKey) || {};
      if (!prev.planStart && r.planStart) prev.planStart = r.planStart;
      if (!prev.planFinish && r.planFinish) prev.planFinish = r.planFinish;
      if (!prev.jo && r.jo) prev.jo = r.jo;
      scheduleByDwg.set(dwgKey, prev);
    }

    const k = `${cleanAlnum(r.noDwg)}|${cleanAlnum(r.drawingName || '')}|${String(r.rev ?? '0').trim()}`;
    const existing = map.get(k);
    if (!existing) {
      map.set(k, { ...r });
    } else {
      // Gabungkan nilai planStart, planFinish, jo, dan finishDate bila gambar ditulis dalam 2 baris (Plan & Actual)
      if (!existing.finishDate && r.finishDate) existing.finishDate = r.finishDate;
      if (!existing.planStart && r.planStart) existing.planStart = r.planStart;
      if (!existing.planFinish && r.planFinish) existing.planFinish = r.planFinish;
      if (!existing.jo && r.jo) existing.jo = r.jo;
    }
  }

  // Wariskan nilai jo, planStart, planFinish ke baris gambar yang sama (mis. Rev 1 mewarisi JO & Plan Finish dari Rev 0)
  map.forEach((item) => {
    const dwgKey = cleanAlnum(item.noDwg) || cleanAlnum(item.drawingName || '');
    const sch = dwgKey ? scheduleByDwg.get(dwgKey) : undefined;
    if (sch) {
      if (!item.planStart && sch.planStart) item.planStart = sch.planStart;
      if (!item.planFinish && sch.planFinish) item.planFinish = sch.planFinish;
      if (!item.jo && sch.jo) item.jo = sch.jo;
    }
  });

  return Array.from(map.values());
};

export interface DrawingControlRow {
  noDwg: string;
  drawingName: string;
  fullDeskripsi: string;
  rev: string;
  finishDate: string;
  planStart?: string;
  planFinish?: string;
  jo?: string;
}

const GAS_DRAWING_API_URL = 'https://script.google.com/macros/s/AKfycbx7bLS2vj_oeW4xDFp3a98A19pN347TuQHRceeFVxZZVC84E398vb4rqEK2SQ0JxMpD/exec';

const DC_MONTHS: Record<string, string> = {
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

export function parseDrawingControlDate(val: any): string {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'number') {
    if (val > 30000 && val < 65000) {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      return d.toISOString().split('T')[0];
    }
    return '';
  }
  let str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'nan') return '';
  str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
  const mon = (k: string) => DC_MONTHS[k.toLowerCase()] || DC_MONTHS[k.toLowerCase().slice(0, 3)] || '';
  const mJs = str.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (mJs && mon(mJs[1])) return `${mJs[3]}-${mon(mJs[1])}-${mJs[2].padStart(2, '0')}`;
  const mDdMon = str.match(/^(\d{1,2})[-/\s.]+([A-Za-z]+)[-/\s.]+(\d{2,4})/);
  if (mDdMon && mon(mDdMon[2])) {
    let y = mDdMon[3];
    if (y.length === 2) y = '20' + y;
    return `${y}-${mon(mDdMon[2])}-${mDdMon[1].padStart(2, '0')}`;
  }
  const mYMon = str.match(/^(\d{4})[-/\s.]+([A-Za-z]+)[-/\s.]+(\d{1,2})/);
  if (mYMon && mon(mYMon[2])) return `${mYMon[1]}-${mon(mYMon[2])}-${mYMon[3].padStart(2, '0')}`;
  const mDmy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mDmy) return `${mDmy[3]}-${mDmy[2].padStart(2, '0')}-${mDmy[1].padStart(2, '0')}`;
  const mYmd = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (mYmd) return `${mYmd[1]}-${mYmd[2].padStart(2, '0')}-${mYmd[3].padStart(2, '0')}`;
  const num = Number(str);
  if (!isNaN(num) && num > 30000 && num < 65000) {
    return new Date(Math.round((num - 25569) * 86400 * 1000)).toISOString().split('T')[0];
  }
  return '';
}

function normalizeHeaderToken(raw: any): string {
  return String(raw ?? '')
    .replace(/[\r\n\t]+/g, ' ')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isNoDwgHeaderCell(raw: any): boolean {
  const s = normalizeHeaderToken(raw);
  const compact = s.replace(/\s+/g, '');
  if (!s) return false;
  return (
    s.includes('no dwg') ||
    compact.includes('nodwg') ||
    s.includes('drawing number') ||
    s.includes('drawing no') ||
    s.includes('dwg number') ||
    s.includes('dwg no') ||
    s.includes('no drawing') ||
    s.includes('nomor gambar') ||
    s.includes('no gambar')
  );
}

function isDrawingNameHeaderCell(raw: any): boolean {
  const s = normalizeHeaderToken(raw);
  if (!s) return false;
  return (
    s.includes('drawing name') ||
    s.includes('dwg name') ||
    s.includes('nama gambar') ||
    s.includes('judul gambar') ||
    s.includes('drawing title') ||
    s.includes('nama dokumen')
  );
}

function isLikelyDrawingNumberCell(val: any): boolean {
  const s = String(val ?? '').trim();
  if (!s || s.length < 4 || s.length > 35) return false;
  if (/^\d+$/.test(s)) return false; // bukan nomor urut 1, 2, 3...
  // Format umum: E00001.51001, E1312C.18100, W000304-01, 100-001, dll.
  if (/^[A-Z]{1,4}[A-Z0-9]*\.[A-Z0-9]+$/i.test(s) && /\d/.test(s)) return true;
  if (/^[A-Z]{1,4}\d{3,}[A-Z0-9._-]*$/i.test(s)) return true;
  if (/^[A-Z0-9]{2,}[-._][A-Z0-9]{2,}(?:[-._][A-Z0-9]+)*$/i.test(s) && /\d{2,}/.test(s) && /[A-Z]/i.test(s)) return true;
  return false;
}

// Parse 2D array dari tab Drawing Control menjadi daftar DrawingControlRow
export function parseDrawingControlSheetRows(rawRows: any[][]): DrawingControlRow[] {
  if (!rawRows || rawRows.length < 2) return [];

  // 1) Cari baris header di 80 baris pertama
  let headerIdx = -1;
  const maxScan = Math.min(80, rawRows.length);
  for (let r = 0; r < maxScan; r++) {
    const row = rawRows[r] || [];
    const hasNoDwg = row.some(isNoDwgHeaderCell);
    const hasName = row.some(isDrawingNameHeaderCell);
    const tokens = row.map(normalizeHeaderToken);
    const hasScheduleCols = tokens.filter(t =>
      t.includes('plan start') ||
      t.includes('plan finish') ||
      t.includes('finish date') ||
      t === 'durasi'
    ).length >= 2;
    if (hasNoDwg || hasName || hasScheduleCols) {
      headerIdx = r;
      break;
    }
  }

  // Jika tidak ada teks header, cek apakah ada minimal 3 baris berisi nomor gambar (mis. E00001.51001)
  if (headerIdx === -1) {
    let firstDwgRow = -1;
    let dwgRowCount = 0;
    for (let r = 0; r < Math.min(100, rawRows.length); r++) {
      const row = rawRows[r] || [];
      if (row.slice(0, 6).some(isLikelyDrawingNumberCell)) {
        if (firstDwgRow === -1) firstDwgRow = r;
        dwgRowCount++;
      }
    }
    if (dwgRowCount >= 3 && firstDwgRow >= 0) {
      headerIdx = Math.max(0, firstDwgRow - 1);
    } else {
      // Bukan tab Drawing Control (mis. tab Dashboard / Rekap / Cover)
      return [];
    }
  }

  // 2) Gabungkan teks header dari baris sekitar headerIdx (untuk menangani merged header 2-3 baris di Excel/Sheets)
  const headerGroupRows: number[] = [];
  for (const cand of [headerIdx - 1, headerIdx, headerIdx + 1, headerIdx + 2]) {
    if (cand < 0 || cand >= rawRows.length) continue;
    const row = rawRows[cand] || [];
    const isDataRow = cand > headerIdx && row.slice(0, 6).some(isLikelyDrawingNumberCell);
    if (!isDataRow) headerGroupRows.push(cand);
  }
  if (!headerGroupRows.includes(headerIdx)) headerGroupRows.push(headerIdx);

  let maxCols = 0;
  for (let r = 0; r < Math.min(100, rawRows.length); r++) {
    const len = (rawRows[r] || []).length;
    if (len > maxCols) maxCols = len;
  }
  if (maxCols === 0) return [];

  const combinedHdr: string[] = [];
  for (let c = 0; c < maxCols; c++) {
    const parts: string[] = [];
    for (const hr of headerGroupRows) {
      const tok = normalizeHeaderToken(rawRows[hr]?.[c]);
      if (tok && !parts.includes(tok)) parts.push(tok);
    }
    combinedHdr[c] = parts.join(' ').trim();
  }

  // 3) Deteksi kolom NO DWG & DRAWING NAME dari header
  let cNoDwg = combinedHdr.findIndex(isNoDwgHeaderCell);
  let cName = combinedHdr.findIndex(isDrawingNameHeaderCell);

  // 4) Verifikasi / deteksi cNoDwg dari baris data (menangani kasus sel header di atas NO DWG kosong/filter seperti pada screenshot)
  const sampleStart = headerIdx + 1;
  const sampleEnd = Math.min(rawRows.length, sampleStart + 80);
  const countDwgInCol = (colIdx: number): number => {
    if (colIdx < 0 || colIdx >= maxCols) return 0;
    let cnt = 0;
    for (let r = sampleStart; r < sampleEnd; r++) {
      if (isLikelyDrawingNumberCell(rawRows[r]?.[colIdx])) cnt++;
    }
    return cnt;
  };

  if (cNoDwg === -1 || countDwgInCol(cNoDwg) === 0) {
    // Cek kolom tepat di sebelah kiri DRAWING NAME terlebih dahulu
    if (cName > 0 && countDwgInCol(cName - 1) > 0) {
      cNoDwg = cName - 1;
    } else {
      let bestCol = -1;
      let bestCnt = 0;
      for (let c = 0; c < Math.min(8, maxCols); c++) {
        if (c === cName) continue;
        const cnt = countDwgInCol(c);
        if (cnt > bestCnt) {
          bestCnt = cnt;
          bestCol = c;
        }
      }
      if (bestCol !== -1 && bestCnt > 0) {
        cNoDwg = bestCol;
      } else if (cNoDwg === -1) {
        cNoDwg = cName > 0 ? cName - 1 : 0;
      }
    }
  }

  if (cName === -1) {
    cName = cNoDwg + 1;
  }

  // 5) Deteksi kolom JO, Plan Start, Plan Finish, Rev, dan Finish Date
  const offset = Math.max(0, cNoDwg);
  const findHdrCol = (pred: (h: string, idx: number) => boolean, fallback: number): number => {
    const idx = combinedHdr.findIndex(pred);
    return idx !== -1 ? idx : fallback;
  };

  const cPlanStart = findHdrCol(
    (h) => h.includes('plan start') || h.includes('start date') || (h.includes('plan') && h.includes('start')),
    findHdrCol((h, i) => i > cName && (h === 'start' || h.endsWith(' start')), 9 + offset)
  );
  const cPlanFinish = findHdrCol(
    (h) => h.includes('plan finish') || h.includes('plan end') || (h.includes('plan') && (h.includes('finish') || h.includes('end'))),
    cPlanStart !== -1 && combinedHdr[cPlanStart + 1]?.includes('finish') ? cPlanStart + 1 : 10 + offset
  );
  const cRev = findHdrCol(
    (h, i) => i > cName && (h === 'rev' || /\brev(?:isi)?\b/.test(h)),
    13 + offset
  );
  const cFinish = findHdrCol(
    (h, i) => i > cName && (h.includes('finish date') || h.includes('release date') || h.includes('actual finish') || h.includes('tgl release')),
    cRev !== -1 && (combinedHdr[cRev + 1]?.includes('finish') || combinedHdr[cRev + 1]?.includes('date')) ? cRev + 1 : 14 + offset
  );
  const cJo = combinedHdr.findIndex(
    (h, i) => i > cName && (/\bjo\b/.test(h) || h.includes('plan jo') || h.includes('jam orang') || h.includes('manhour') || h.includes('man hour') || /\bmh\b/.test(h))
  );

  const SKIP_LABELS = new Set([
    'no dwg', 'no. dwg', 'no.dwg', 'nodwg', 'drawing number', 'drawing no', 'dwg no',
    'drawing name', 'nama gambar', 'plan start', 'plan finish', 'start', 'finish',
    'finish date', 'rev', 'jo', 'durasi', 'total', 'sub total', 'subtotal',
  ]);

  const rows: DrawingControlRow[] = [];
  for (let r = headerIdx + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row) continue;
    const noDwg = String(row[cNoDwg] ?? '').replace(/[\r\n\t]+/g, ' ').trim();
    const dwgName = String(row[cName] ?? '').replace(/[\r\n\t]+/g, ' ').trim();
    if (!noDwg && !dwgName) continue;

    const lowNo = normalizeHeaderToken(noDwg);
    const lowName = normalizeHeaderToken(dwgName);
    if (SKIP_LABELS.has(lowNo) || SKIP_LABELS.has(lowName)) continue;
    if (isNoDwgHeaderCell(noDwg) || isDrawingNameHeaderCell(dwgName)) continue;
    // Abaikan baris yang hanya berisi nomor urut 1, 2, 3 tanpa nama gambar
    if (/^\d{1,4}$/.test(noDwg) && !dwgName) continue;

    let revRaw = String(row[cRev] ?? '').trim();
    let finishRaw = row[cFinish];
    if (!parseDrawingControlDate(finishRaw)) {
      if (row[cRev + 1] && parseDrawingControlDate(row[cRev + 1])) {
        finishRaw = row[cRev + 1];
      } else if (row[15 + offset] && parseDrawingControlDate(row[15 + offset])) {
        revRaw = String(row[14 + offset] ?? revRaw).trim();
        finishRaw = row[15 + offset];
      } else if (row[14 + offset] && parseDrawingControlDate(row[14 + offset])) {
        revRaw = String(row[13 + offset] ?? revRaw).trim();
        finishRaw = row[14 + offset];
      }
    }

    const planStart = parseDrawingControlDate(row[cPlanStart]);
    const planFinish = parseDrawingControlDate(row[cPlanFinish]);
    const finishDate = parseDrawingControlDate(finishRaw) || String(finishRaw ?? '').trim();
    const joRaw = cJo !== -1 ? String(row[cJo] ?? '').trim() : '';
    const joClean = joRaw && joRaw !== '-' && joRaw.toLowerCase() !== 'nan'
      ? joRaw.replace(/[^0-9.,]/g, '').replace(',', '.')
      : '';

    rows.push({
      noDwg,
      drawingName: dwgName,
      fullDeskripsi: (noDwg && dwgName) ? `${noDwg}-${dwgName}` : (dwgName || noDwg),
      rev: revRaw && revRaw !== '-' && revRaw.toLowerCase() !== 'nan' ? revRaw : '0',
      finishDate,
      planStart,
      planFinish,
      jo: joClean,
    });
  }

  return dedupeRows(rows);
}

// Cek apakah nama sheet bertuliskan "drawing control" atau "drawin control" (bebas huruf besar/kecil)
function isDrawingControlSheetName(sheetName: string): boolean {
  const s = (sheetName || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  const compact = s.replace(/\s+/g, '');
  return /drawin[g]?\s*control/i.test(s) || compact.includes('drawingcontrol') || compact.includes('drawincontrol');
}

// Variasi penulisan nama tab Drawing Control di Google Sheets (untuk endpoint gviz/tq yang case-sensitive)
const GVIZ_DRAWING_CONTROL_TAB_NAMES: string[] = [
  'Drawing Control',
  'DRAWING CONTROL',
  'drawing control',
  'Drawing control',
  'Drawin Control',
  'DRAWIN CONTROL',
  'drawin control',
  'Drawin control',
  'Drawing Control (2)',
  'DRAWING CONTROL (2)',
  'drawing control (2)',
  'Drawing control (2)',
  'Drawin Control (2)',
  'DRAWIN CONTROL (2)',
  'drawin control (2)',
  'Drawing Control (1)',
  'DRAWING CONTROL (1)',
  'Drawing Control ',
  'DRAWING CONTROL ',
  'Drawin Control ',
  'DRAWIN CONTROL ',
];

const sheetIdRowsCache = new Map<string, DrawingControlRow[]>();
const projectRowsCache = new Map<string, DrawingControlRow[]>();

export function clearDrawingControlCache(): void {
  sheetIdRowsCache.clear();
  projectRowsCache.clear();
}

function parseCsvTextToRows(csvText: string): DrawingControlRow[] {
  if (!csvText || !csvText.trim() || csvText.trim().startsWith('<!DOCTYPE') || csvText.trim().startsWith('<html')) {
    return [];
  }
  try {
    const wb = XLSX.read(csvText, { type: 'string' });
    const firstSheet = wb.Sheets[wb.SheetNames[0]];
    if (!firstSheet) return [];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    return parseDrawingControlSheetRows(rawRows);
  } catch {
    return [];
  }
}

async function fetchDrawingRowsFromSheetId(sheetId: string): Promise<DrawingControlRow[]> {
  if (!sheetId) return [];
  if (sheetIdRowsCache.has(sheetId)) {
    const cached = sheetIdRowsCache.get(sheetId)!;
    if (cached.length > 0) return cached;
  }

  // 1) Coba ambil workbook penuh (.xlsx) sehingga seluruh daftar nama tab (SheetNames) bisa dicocokkan
  //    secara case-insensitive ("drawin control", "Drawing Control", "DRAWING CONTROL", "Drawing Control (2)", dll.)
  try {
    const xlsxRes = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`);
    if (xlsxRes.ok) {
      const buf = await xlsxRes.arrayBuffer();
      const bytes = new Uint8Array(buf.slice(0, 4));
      if (bytes[0] === 0x50 && bytes[1] === 0x4B) {
        const wb = XLSX.read(buf, { type: 'array' });
        const dcTabs = wb.SheetNames.filter(isDrawingControlSheetName);
        const otherTabs = wb.SheetNames.filter(n => !isDrawingControlSheetName(n));
        let bestRows: DrawingControlRow[] = [];

        for (const tabName of dcTabs) {
          const rawRows: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[tabName], { header: 1, defval: '' });
          const parsed = parseDrawingControlSheetRows(rawRows);
          if (parsed.length > bestRows.length) {
            bestRows = parsed;
          }
        }
        if (bestRows.length === 0) {
          for (const tabName of otherTabs) {
            const rawRows: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[tabName], { header: 1, defval: '' });
            const parsed = parseDrawingControlSheetRows(rawRows);
            if (parsed.length > bestRows.length) {
              bestRows = parsed;
            }
          }
        }
        if (bestRows.length > 0) {
          sheetIdRowsCache.set(sheetId, bestRows);
          return bestRows;
        }
      }
    }
  } catch {
    // Lanjut ke metode gviz/tq CSV
  }

  // 2) Metode gviz/tq CSV:
  //    Cek juga apakah htmlview menyediakan nama tab persisnya, lalu uji semua variasi huruf kapital
  //    ("Drawing Control", "DRAWING CONTROL", "drawin control", "Drawin Control", "Drawing Control (2)", dll.)
  const extraTabNames: string[] = [];
  try {
    const htmlRes = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/htmlview`);
    if (htmlRes.ok) {
      const html = await htmlRes.text();
      const matches = html.match(/drawin[g]?\s*control(?:\s*\(\d+\))?/gi) || [];
      matches.forEach(m => {
        const clean = m.trim();
        if (clean && !extraTabNames.includes(clean)) extraTabNames.push(clean);
      });
    }
  } catch {
    // Abaikan jika htmlview diblokir CORS
  }

  const allTabCandidates = Array.from(new Set([...extraTabNames, ...GVIZ_DRAWING_CONTROL_TAB_NAMES]));
  const baseUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv`;

  // Ambil tab pertama (default) sebagai pembanding, karena gviz/tq mengembalikan tab pertama bila nama tab tidak cocok
  let defaultCsv = '';
  try {
    const defRes = await fetch(baseUrl);
    if (defRes.ok) defaultCsv = await defRes.text();
  } catch {}

  // Uji kandidat nama tab secara paralel per batch
  const BATCH_SIZE = 6;
  for (let i = 0; i < allTabCandidates.length; i += BATCH_SIZE) {
    const batch = allTabCandidates.slice(i, i + BATCH_SIZE);
    const csvList = await Promise.all(
      batch.map(async (tabName) => {
        try {
          const res = await fetch(`${baseUrl}&sheet=${encodeURIComponent(tabName)}`);
          if (!res.ok) return '';
          return await res.text();
        } catch {
          return '';
        }
      })
    );

    // Prioritaskan respons yang berbeda dari defaultCsv (artinya tab dengan nama tersebut benar-benar ditemukan!)
    for (const csvText of csvList) {
      if (!csvText || (defaultCsv && csvText === defaultCsv)) continue;
      const rows = parseCsvTextToRows(csvText);
      if (rows.length > 0) {
        sheetIdRowsCache.set(sheetId, rows);
        return rows;
      }
    }
  }

  // Jika semua kandidat sama dengan defaultCsv, berarti tab Drawing Control memang berada di tab pertama
  if (defaultCsv) {
    const defRows = parseCsvTextToRows(defaultCsv);
    if (defRows.length > 0) {
      sheetIdRowsCache.set(sheetId, defRows);
      return defRows;
    }
  }

  return [];
}

async function fetchDrawingRowsForSingleCandidate(projectCode: string): Promise<DrawingControlRow[]> {
  const cleanProj = cleanAlnum(projectCode);
  if (!cleanProj) return [];

  // 1) Utamakan membaca langsung dari file Google Sheet proyek yang dipilih (tab Drawing Control / DRAWING CONTROL / drawin control)
  const match = findMatchingDriveSheet(projectCode);
  if (match && match.id) {
    const sheetRows = await fetchDrawingRowsFromSheetId(match.id);
    if (sheetRows.length > 0) return sheetRows;
  }

  // 2) Cadangan ke Google Apps Script bila file bukan Google Sheet standar (mis. file .xlsx di Drive)
  const gasQueries = Array.from(new Set([projectCode.trim(), match?.projectKey || ''].filter(Boolean)));
  for (const q of gasQueries) {
    try {
      const res = await fetch(`${GAS_DRAWING_API_URL}?project=${encodeURIComponent(q)}`);
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data) && json.data.length > 0) {
          const gasRows: DrawingControlRow[] = json.data.map((item: any) => ({
            noDwg: String(item.noDwg || '').trim(),
            drawingName: String(item.drawingName || '').trim(),
            fullDeskripsi: String(item.deskripsi || item.fullDeskripsi || `${item.noDwg || ''}-${item.drawingName || ''}`).trim(),
            rev: String(item.rev ?? '0').trim(),
            finishDate: parseDrawingControlDate(item.finishDate) || String(item.finishDate || '').trim(),
            planStart: parseDrawingControlDate(item.planStart || item.startDate || item.plan_start || ''),
            planFinish: parseDrawingControlDate(item.planFinish || item.endDate || item.plan_finish || ''),
            jo: String(item.jo || item.planJo || item.durasi || '').replace(/[^0-9.,]/g, '').replace(',', '.'),
          }));
          if (gasRows.length > 0) return dedupeRows(gasRows);
        }
      }
    } catch {}
  }

  return [];
}

// Fungsi utama untuk mengambil seluruh baris Drawing Control berdasarkan kode proyek apa pun
export async function fetchDrawingControlRowsForProject(projectCode: string): Promise<DrawingControlRow[]> {
  const cacheKey = cleanAlnum(projectCode);
  if (!cacheKey) return [];
  if (projectRowsCache.has(cacheKey)) {
    const cached = projectRowsCache.get(cacheKey)!;
    if (cached.length > 0) return cached;
  }

  // Jika projectCode langsung cocok dengan salah satu file di GOOGLE_DRIVE_SHEETS, langsung ambil dari file tersebut
  const directMatch = findMatchingDriveSheet(projectCode);
  if (directMatch && directMatch.id) {
    const directRows = await fetchDrawingRowsForSingleCandidate(projectCode);
    if (directRows.length > 0) {
      projectRowsCache.set(cacheKey, directRows);
      return directRows;
    }
  }

  const cands = projectCandidates(projectCode);
  if (!cands.length) cands.push(projectCode);
  const results = await Promise.all(
    cands.map(c => fetchDrawingRowsForSingleCandidate(c).catch(() => [] as DrawingControlRow[]))
  );
  const deduped = dedupeRows(results.flat());
  if (deduped.length > 0) {
    projectRowsCache.set(cacheKey, deduped);
  }
  return deduped;
}

// Pecah teks deskripsi gabungan (mis. "E25291.01100 - ESTIMATION MATERIAL..." atau "R0 - E25291.01100 - ...")
// menjadi 2 bagian terpisah: Drawing Number (noDwg) dan Drawing Name (drawingName).
export function splitTaskToDrawingParts(raw: string): { drawingNumber: string; drawingName: string } {
  const s = (raw || '')
    .replace(/(\b[A-Z]{1,4}\d{2,}[A-Z0-9]*)\s*\.\s*([A-Z0-9]+)/gi, '$1.$2')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return { drawingNumber: '', drawingName: '' };

  // Hilangkan awalan R0 / REV 1 / Kode Proyek di depan bila ada
  let rest = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-–—_:]\s*/i, '');
  rest = rest.replace(/^[WM]\d{4,6}\s*[-–—_:/]\s*/i, '');
  rest = rest.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-–—_:]\s*/i, '');

  // 1) Pola "NO_DWG - DRAWING_NAME" atau "NO_DWG_R0_DRAWING_NAME"
  const mLead = rest.match(/^([A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+|[A-Z]{1,4}\d{3,}[A-Z0-9]*)\s*(?:[-–—_:]\s*(?:R(?:EV)?\.?\s*\d{1,2}\s*[-–—_:]\s*)?|\s+-\s+)(.+)$/i);
  if (mLead) {
    return {
      drawingNumber: mLead[1].trim().toUpperCase(),
      drawingName: mLead[2].trim(),
    };
  }

  // 2) Jika seluruh teks hanya berisi nomor gambar (mis. "E25291.01100")
  if (/^([A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+|[A-Z]{1,4}\d{3,}[A-Z0-9]*)$/i.test(rest)) {
    return { drawingNumber: rest.toUpperCase(), drawingName: '' };
  }

  // 3) Cari nomor gambar berformat titik di tengah teks
  const mDot = rest.match(/\b([A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+)\b/i);
  if (mDot) {
    const num = mDot[1].toUpperCase();
    const name = rest
      .replace(mDot[0], '')
      .replace(/^[\s\-–—_:]+|[\s\-–—_:]+$/g, '')
      .replace(/\s*[-–—_:]\s*R(?:EV)?\.?\s*\d{1,2}\s*[-–—_:]/i, ' - ')
      .trim();
    return { drawingNumber: num, drawingName: name };
  }

  return { drawingNumber: '', drawingName: s };
}

function extractSimpleDrawingNumber(desc: string): string {
  let s = (desc || '')
    .replace(/(\b[A-Z]{1,4}\d{2,}[A-Z0-9]*)\s*\.\s*([A-Z0-9]+)/gi, '$1.$2')
    .replace(/\s+/g, ' ')
    .trim();
  s = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-\s_:]\s*/i, '');
  s = s.replace(/^[WM]\d{4,6}\s*[-–—_:/]\s*/i, '');
  s = s.replace(/^(?:KSOT|LCU|W?FFBNW\d*)\s*[-–—_:/]\s*/i, '');
  s = s.replace(/^[A-Z]{2,12}\d{0,6}\s*[-–—_:/]\s*(?=[A-Z]{1,4}\d{2,}[A-Z0-9]*\.)/i, '');
  s = s.replace(/^R(?:EV)?\.?\s*\d{0,2}\s*[-\s_:]\s*/i, '');
  const mDot = s.match(/\b[A-Z]{1,4}\d{2,}[A-Z0-9]*(?:\.[A-Z0-9]+)+/i);
  if (mDot) return mDot[0].toUpperCase();
  const m = s.match(/^[A-Z]{1,4}\d{3,}[A-Z0-9]*(?:\.[A-Z0-9]+)*/i);
  if (m) return m[0].toUpperCase();
  const first = s.split(/\s+-\s+|\s+|_/)[0] || '';
  return /\d/.test(first) && first.length >= 4 ? first.toUpperCase() : '';
}

// Cari baris Drawing Control yang cocok dengan taskName dan rev (prioritaskan rev yang sama, fallback ke gambar yang sama untuk jadwal/JO)
export function findDrawingControlMatch(
  rows: DrawingControlRow[],
  taskName: string,
  rev: string | number = '0'
): { release: string; planStart: string; planFinish: string; jo: string } {
  const empty = { release: '', planStart: '', planFinish: '', jo: '' };
  if (!rows || rows.length === 0 || !taskName) return empty;

  const wanted = cleanAlnum(extractSimpleDrawingNumber(taskName) || taskName);
  const strippedTask = (taskName || '')
    .replace(/^[WM]\d{4,6}\s*[-–—_:/]\s*/i, '')
    .replace(/^[A-Z0-9]{2,12}(?:-[A-Z0-9]{2,8})?\s*[-–—_:/]\s*/i, '')
    .trim();
  const cleanTarget = cleanAlnum(taskName);
  const cleanStripped = cleanAlnum(strippedTask);
  if (!wanted && !cleanTarget) return empty;

  const toNum = (v: any) => {
    const n = parseInt(String(v ?? '').replace(/[^0-9]/g, ''), 10);
    return isNaN(n) ? 0 : n;
  };
  const wantRev = toNum(rev);

  const isRowDrawingMatch = (r: DrawingControlRow): boolean => {
    const rKey = cleanAlnum(extractSimpleDrawingNumber(r.noDwg) || r.noDwg);
    if (wanted && rKey && rKey === wanted) return true;
    const cDwg = cleanAlnum(r.noDwg);
    if (cDwg && cDwg.length >= 5 && cleanTarget.includes(cDwg)) return true;
    const cName = cleanAlnum(r.drawingName);
    if (
      cName &&
      cName.length >= 5 &&
      (cName === cleanStripped ||
        cleanTarget.includes(cName) ||
        (cleanStripped.length >= 5 && cName.includes(cleanStripped)))
    ) {
      return true;
    }
    return false;
  };

  const matchedRows = rows.filter(isRowDrawingMatch);
  if (matchedRows.length === 0) return empty;

  const revMatches = matchedRows.filter(r => toNum(r.rev) === wantRev);
  const releaseHit = revMatches.find(r => /^\d{4}-\d{2}-\d{2}$/.test(parseDrawingControlDate(r.finishDate)));
  const release = releaseHit ? parseDrawingControlDate(releaseHit.finishDate) : '';

  const pickVal = (getter: (r: DrawingControlRow) => string | undefined): string => {
    for (const r of revMatches) {
      const v = getter(r);
      if (v) return v;
    }
    for (const r of matchedRows) {
      const v = getter(r);
      if (v) return v;
    }
    return '';
  };

  const planStart = pickVal(r => (/^\d{4}-\d{2}-\d{2}$/.test(parseDrawingControlDate(r.planStart)) ? parseDrawingControlDate(r.planStart) : ''));
  const planFinish = pickVal(r => (/^\d{4}-\d{2}-\d{2}$/.test(parseDrawingControlDate(r.planFinish)) ? parseDrawingControlDate(r.planFinish) : ''));
  const jo = pickVal(r => (r.jo && r.jo !== '-' && r.jo !== '0' ? r.jo : ''));

  return { release, planStart, planFinish, jo };
}

// Gabungkan kembali Drawing Number dan Drawing Name menjadi teks task_name untuk disimpan di database
export function combineDrawingParts(drawingNumber: string, drawingName: string): string {
  const num = (drawingNumber || '').trim();
  const name = (drawingName || '').trim();
  if (num && name) {
    // Hindari duplikasi jika di dalam name sudah diawali num
    if (name.toUpperCase().startsWith(num.toUpperCase())) return name;
    return `${num} - ${name}`;
  }
  return num || name;
}