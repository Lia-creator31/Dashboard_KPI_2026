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

export const dedupeRows = <T extends { noDwg: string; rev: string; finishDate: string }>(rows: T[]): T[] => {
  const seen = new Set<string>();
  return rows.filter(r => {
    const k = `${r.noDwg}|${r.rev}|${r.finishDate}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
};