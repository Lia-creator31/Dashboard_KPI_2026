// Membuat daftar nama kandidat dari satu teks proyek, supaya Drawing Control di Drive
// ditemukan walau namanya hanya cocok sebagian.
//
// "WFFBNW - W000304" -> ["WFFBNW", "FFBNW", "W000304"]
// "WFFBNW2"          -> ["WFFBNW2", "WFFBNW", "FFBNW"]
// "W0000304 - 305"   -> ["W0000304", "W000304", "W000305"]
export function projectCandidates(raw: string): string[] {
  const out = new Set<string>();
  const add = (s: string) => { const v = s.trim(); if (v) out.add(v); };

  const tokens = (raw || '').toUpperCase()
    .split(/\s*(?:-|–|—|\/|,|;|&|\+|\bDAN\b)\s*/)
    .map(t => t.replace(/[^A-Z0-9]/g, ''))
    .filter(Boolean);

  let prefix = ''; // huruf kode sebelumnya, untuk "W000304 - 305"
  for (const t of tokens) {
    // angka saja (mis. "305") -> pakai huruf dari kode sebelumnya
    if (/^\d+$/.test(t)) {
      if (prefix) add(prefix + t.padStart(6, '0'));
      continue;
    }
    add(t);

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
      // WFFBNW -> FFBNW (nama di daftar sheet tanpa W)
      if (base.startsWith('W') && base.length >= 5) add(base.slice(1));
    }
  }
  return Array.from(out);
}

export interface DrawingSheetConfig {
  id: string;
  projectKey: string;
  title: string;
}

// SATU-SATUNYA tempat daftar Link Google Sheet Drawing Control untuk seluruh aplikasi.
// Jika ada link Google Sheet baru, cukup tambahkan 1 baris di sini -> otomatis berlaku di
// RendalPage, KabiroPage, KpiDashboardPage, SyncReleasePage, dan ImportJobcardPage.
export const GOOGLE_DRIVE_SHEETS: DrawingSheetConfig[] = [
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

// Otomatis terbentuk dari GOOGLE_DRIVE_SHEETS (tidak perlu ketik 2 kali)
export const KNOWN_DRAWING_PROJECT_KEYS: string[] = Array.from(
  new Set(GOOGLE_DRIVE_SHEETS.map(s => s.projectKey.toUpperCase().trim()))
);

// Jika proyek adalah CORP2026 (atau proyek non-link lainnya), cek di dalam deskripsi
// apakah ada kode proyek yang terdaftar di link Drawing Control (mis. "W000314 - E25291.01100...")
// ATAU pola kode proyek baru di awal deskripsi (mis. "W000315 - ...", "M000314 - ...").
export function findDrawingProjectInDescription(desc: string): string {
  if (!desc) return '';
  const upper = desc.toUpperCase().trim();
  // 1) Cek kandidat proyek dari token-token deskripsi terhadap daftar link Drawing Control
  const cands = projectCandidates(upper);
  for (const c of cands) {
    const match = KNOWN_DRAWING_PROJECT_KEYS.find(k => k === c);
    if (match) return match;
  }
  // 2) Cek kecocokan kata utuh untuk tiap projectKey di link Drawing Control
  for (const key of KNOWN_DRAWING_PROJECT_KEYS) {
    const re = new RegExp(`(?:^|[^A-Z0-9])${key}(?:$|[^A-Z0-9])`, 'i');
    if (re.test(upper)) return key;
  }
  // 3) Dinamis (otomatis untuk proyek baru yang ditambahkan di folder Google Drive / GAS):
  //    Jika di awal deskripsi ada pola kode proyek (mis. "W000315 - E25..." atau "M000316 - ..."),
  //    ambil kode proyek tersebut meskipun belum ditulis di KNOWN_DRAWING_PROJECT_KEYS.
  const mLeadCode = upper.match(/^([WM]\d{4,6})\s*[-–—_:/]/);
  if (mLeadCode) {
    const norm = projectCandidates(mLeadCode[1]);
    return norm.find(x => /^[WM]\d{6}$/.test(x)) || mLeadCode[1];
  }
  // 4) Atau jika di dalam deskripsi ada kode proyek standar W000xxx / M000xxx (tanpa titik di belakangnya, agar tidak tertukar dengan nomor gambar)
  const mAnyCode = upper.match(/(?:^|[^A-Z0-9.])([WM]0{2,4}\d{2,4})(?![A-Z0-9.])/);
  if (mAnyCode) {
    const norm = projectCandidates(mAnyCode[1]);
    return norm.find(x => /^[WM]\d{6}$/.test(x)) || mAnyCode[1];
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

export const dedupeRows = <T extends { noDwg: string; rev: string; finishDate: string }>(rows: T[]): T[] => {
  const seen = new Set<string>();
  return rows.filter(r => {
    const k = `${r.noDwg}|${r.rev}|${r.finishDate}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
};