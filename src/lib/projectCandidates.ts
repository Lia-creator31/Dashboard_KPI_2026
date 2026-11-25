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

export const dedupeRows = <T extends { noDwg: string; rev: string; finishDate: string }>(rows: T[]): T[] => {
  const seen = new Set<string>();
  return rows.filter(r => {
    const k = `${r.noDwg}|${r.rev}|${r.finishDate}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
};

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