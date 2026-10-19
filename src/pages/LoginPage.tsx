import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { KeyRound } from 'lucide-react';
import { supabase } from '../lib/supabase';
import { UserSession } from '../App';

interface LoginPageProps {
  onLogin: (userData: UserSession) => void;
}

// Admin Rendal dikenali langsung dari NIP
const ADMIN_ACCOUNTS = [
  { nama: 'Dyan Asih Purwanti', nip: '104134283' },
  { nama: 'Hashfi Moch Adam', nip: '105264937' },
];

const MASTER_BUCKET = 'master-files';

interface Im4Member {
  nama: string;
  nip: string;
  status: string;
  jabatan: string;
}

// Samakan format NIP (hilangkan spasi dan ".0" dari angka Excel)
const normalizeNip = (v: any): string =>
  String(v ?? '').trim().replace(/\s+/g, '').replace(/\.0+$/, '');

async function loadIm4Workbook(): Promise<XLSX.WorkBook | null> {
  try {
    const { data, error } = await supabase.storage.from(MASTER_BUCKET).download('im4.xlsx');
    if (!error && data) return XLSX.read(await data.arrayBuffer(), { type: 'array' });
  } catch { /* lanjut ke cadangan */ }
  for (const p of ['/AKSES AKUN IM4 UNTUK MENU DRAWING CONTROL (1).xlsx']) {
    try {
      const res = await fetch(p);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      const b = new Uint8Array(buf.slice(0, 4));
      if (b[0] === 80 && b[1] === 75 && b[2] === 3 && b[3] === 4) return XLSX.read(buf, { type: 'array' });
    } catch { /* abaikan */ }
  }
  return null;
}

// Baca master IM4: nama, NIP, status, jabatan
function parseMembers(wb: XLSX.WorkBook | null): Im4Member[] {
  const out: Im4Member[] = [];
  if (!wb) return out;
  try {
    const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes('education')) || wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName]; if (!sheet) return out;
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    let headerIdx = -1;
    for (let r = 0; r < Math.min(15, rows.length); r++) {
      const vals = (rows[r] || []).map(v => String(v).trim().toLowerCase());
      if (vals.includes('nama') && (vals.includes('nip') || vals.includes('status') || vals.includes('jabatan'))) { headerIdx = r; break; }
    }
    if (headerIdx === -1) return out;
    const headers = rows[headerIdx].map(v => String(v).trim().toLowerCase());
    const namaCol = headers.findIndex(h => h === 'nama');
    const nipCol = headers.findIndex(h => h === 'nip');
    const statusCol = headers.findIndex(h => h === 'status');
    const jabatanCol = headers.findIndex(h => h.includes('jabatan'));
    if (namaCol === -1 || nipCol === -1) return out;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const nama = String(row[namaCol] || '').trim();
      const nip = normalizeNip(row[nipCol]);
      if (!nama || nama.toLowerCase() === 'nan' || !nip) continue;
      out.push({
        nama,
        nip,
        status: statusCol !== -1 ? String(row[statusCol] || '').trim() : '',
        jabatan: jabatanCol !== -1 ? String(row[jabatanCol] || '').trim() : '',
      });
    }
  } catch { /* abaikan */ }
  return out;
}

// Aturan peran dari data IM4 (ubah di sini bila aturan akses berubah):
// - status Outsourcing  -> outsourcing
// - jabatan Kepala Biro -> kabiro
// - selain itu          -> tidak punya akses
function roleFromMember(m: Im4Member): 'kabiro' | 'outsourcing' | null {
  const status = m.status.toLowerCase();
  const jabatan = m.jabatan.toLowerCase();
  if (status.includes('outsourcing')) return 'outsourcing';
  if (jabatan.includes('kepala biro') || jabatan.includes('kabiro')) return 'kabiro';
  return null;
}

export default function LoginPage({ onLogin }: LoginPageProps) {
  const [nip, setNip] = useState('');
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleLoginSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const cleanNip = normalizeNip(nip);
    if (!cleanNip) {
      setError('NIP wajib diisi!');
      return;
    }

    // Admin: langsung dikenali dari NIP
    const admin = ADMIN_ACCOUNTS.find(a => a.nip === cleanNip);
    if (admin) {
      onLogin({ nama: admin.nama, nip: cleanNip, role: 'admin' });
      return;
    }

    // Kabiro / Outsourcing: dicari di master IM4 lewat NIP
    setIsLoading(true);
    try {
      const members = parseMembers(await loadIm4Workbook());
      if (members.length === 0) {
        setError('Master personel (IM4) belum termuat. Hubungi admin Rendal.');
        return;
      }
      const member = members.find(m => m.nip === cleanNip);
      if (!member) {
        setError('Login gagal: NIP tidak terdaftar.');
        return;
      }
      const role = roleFromMember(member);
      if (!role) {
        setError('Login gagal: NIP ini tidak memiliki akses (hanya Kabiro, Outsourcing, dan Admin).');
        return;
      }
      onLogin({ nama: member.nama, nip: cleanNip, role });
    } catch {
      setError('Terjadi kesalahan saat memuat data. Coba lagi.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex items-center justify-center p-4 font-sans">
      <div className="max-w-md w-full bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
        <div className="text-center space-y-2">
          <h1 className="text-xl font-bold text-white">Sistem Penugasan Job Card</h1>
          <p className="text-xs text-slate-400">Silakan masukkan NIP Anda untuk masuk</p>
        </div>

        {error && (
          <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg text-rose-400 text-xs text-center">
            {error}
          </div>
        )}

        <form onSubmit={handleLoginSubmit} className="space-y-4 text-xs">
          <div>
            <label className="block text-slate-400 mb-1">NIP</label>
            <div className="relative">
              <input
                type="text"
                inputMode="numeric"
                value={nip}
                onChange={(e) => { setNip(e.target.value); setError(''); }}
                placeholder="Ketik NIP Anda..."
                required
                autoFocus
                className="w-full px-3 py-2.5 pl-9 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono focus:outline-none focus:border-blue-500"
              />
              <KeyRound className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
            </div>
          </div>

          <button
            type="submit"
            disabled={isLoading}
            className="w-full py-3 bg-blue-600 hover:bg-blue-500 disabled:opacity-60 disabled:cursor-wait text-white font-bold rounded-xl transition cursor-pointer shadow-lg"
          >
            {isLoading ? 'Memeriksa NIP...' : 'Masuk Sistem'}
          </button>
        </form>
      </div>
    </div>
  );
}