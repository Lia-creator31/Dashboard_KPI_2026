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

export const KNOWN_DRAWING_PROJECT_KEYS = [
  'M000313',
  'W000304',
  'W000305',
  'W000308',
  'W000314',
  'M000312',
  'W000311',
  'W000306',
  'W000307',
  'KSOT',
  'LCU',
  'FFBNW',
];

// Jika proyek adalah CORP2026 (atau proyek non-link lainnya), cek di dalam deskripsi
// apakah ada kode proyek yang terdaftar di link Drawing Control (mis. "W000314 - E25291.01100...").
export function findDrawingProjectInDescription(desc: string): string {
  if (!desc) return '';
  const upper = desc.toUpperCase();
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