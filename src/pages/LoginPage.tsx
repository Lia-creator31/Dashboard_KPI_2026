import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { KeyRound, ShieldCheck, ArrowRight, Compass } from 'lucide-react';
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

// Akun yang langsung masuk ke Rekap Absensi Divisi
const REKAP_ACCOUNTS = [
  { nama: 'Pengguna Rekap Absensi', nip: '105265074' },
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

    // Rekap: langsung dikenali dari NIP, tidak perlu cek master IM4
    const rekap = REKAP_ACCOUNTS.find(a => a.nip === cleanNip);
    if (rekap) {
      onLogin({ nama: rekap.nama, nip: cleanNip, role: 'rekap' });
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
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col justify-between p-4 font-sans">
      {/* Top bar identitas korporat */}
      <div className="max-w-5xl w-full mx-auto pt-3 px-2 flex items-center justify-between text-xs text-slate-400">
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded-md bg-blue-600/15 border border-blue-500/30 text-blue-400 flex items-center justify-center">
            <Compass className="w-4 h-4" />
          </div>
          <div>
            <span className="font-bold text-slate-200 tracking-tight">DIVISI DESAIN</span>
            <span className="mx-2 text-slate-700">|</span>
            <span className="text-slate-400">Sistem Penugasan &amp; Monitoring KPI</span>
          </div>
        </div>
        <span className="font-mono text-[11px] text-slate-500 hidden sm:inline">TA 2026</span>
      </div>

      {/* Card Login Utama */}
      <div className="max-w-sm w-full mx-auto my-8">
        <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 sm:p-7 space-y-6 shadow-xl">
          <div className="space-y-1.5 border-b border-slate-800 pb-4">
            <div className="text-[11px] font-semibold uppercase tracking-wider text-blue-400 flex items-center gap-1.5">
              <ShieldCheck className="w-3.5 h-3.5" /> Otentikasi Personel
            </div>
            <h1 className="text-lg font-bold text-white">Masuk ke Portal Kerja</h1>
            <p className="text-xs text-slate-400 leading-relaxed">
              Gunakan Nomor Induk Pegawai (NIP) untuk mengakses penugasan Jobcard, Work Order, atau Dashboard KPI.
            </p>
          </div>

          {error && (
            <div className="p-3 bg-rose-500/10 border border-rose-500/25 rounded-lg text-rose-300 text-xs leading-relaxed">
              {error}
            </div>
          )}

          <form onSubmit={handleLoginSubmit} className="space-y-4">
            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1.5">
                Nomor Induk Pegawai (NIP)
              </label>
              <div className="relative">
                <input
                  type="text"
                  inputMode="numeric"
                  value={nip}
                  onChange={(e) => { setNip(e.target.value); setError(''); }}
                  placeholder="Masukkan NIP Anda"
                  required
                  autoFocus
                  className="w-full px-3.5 py-2.5 pl-9 bg-slate-950 border border-slate-700 rounded-lg text-sm text-white font-mono placeholder:font-sans placeholder:text-slate-500 focus:outline-none focus:border-blue-500 transition"
                />
                <KeyRound className="w-4 h-4 text-slate-500 absolute left-3 top-3" />
              </div>
            </div>

            <button
              type="submit"
              disabled={isLoading}
              className="w-full py-2.5 px-4 bg-blue-600 hover:bg-blue-500 disabled:opacity-60 disabled:cursor-wait text-white text-xs font-semibold rounded-lg transition cursor-pointer flex items-center justify-center gap-2"
            >
              <span>{isLoading ? 'Memverifikasi NIP...' : 'Masuk ke Sistem'}</span>
              {!isLoading && <ArrowRight className="w-3.5 h-3.5" />}
            </button>
          </form>

          <div className="pt-3 border-t border-slate-800/80 text-[11px] text-slate-500 flex items-center justify-between">
            <span>Akses: Rendal · Kabiro · Mitra</span>
            <span className="font-mono">IM4 &amp; Drawing Control</span>
          </div>
        </div>
      </div>

      {/* Footer sederhana */}
      <div className="max-w-5xl w-full mx-auto pb-3 px-2 text-center sm:text-left text-[11px] text-slate-600">
        Divisi Desain — Sistem Informasi Penugasan, Realisasi Jam Orang (JO), dan Evaluasi KPI
      </div>
    </div>
  );
}