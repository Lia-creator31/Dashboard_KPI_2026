import React, { useState } from 'react';
import { User, KeyRound, Briefcase, HardHat } from 'lucide-react';
import { UserSession } from '../App';

interface LoginPageProps {
  onLogin: (userData: UserSession) => void;
}

const ADMIN_ACCOUNTS = [
  { nama: 'dyan asih purwanti', nip: '104134283' },
  { nama: 'hashfi moch adam', nip: '105264937' },
];

export default function LoginPage({ onLogin }: LoginPageProps) {
  const [nama, setNama] = useState('');
  const [nip, setNip] = useState('');
  const [role, setRole] = useState<'kabiro' | 'outsourcing'>('kabiro');
  const [error, setError] = useState('');

const cleanNama = nama.trim().toLowerCase();
const cleanNip = nip.trim();

// Nama yang termasuk daftar admin (baru nama, belum cek NIP)
const matchedAdmin = ADMIN_ACCOUNTS.find(a => cleanNama.includes(a.nama));
const isAdmin = !!matchedAdmin;

// Status hanya muncul jika kolom nama sudah diisi dan bukan admin
const showRoleSelector = nama.trim().length > 0 && !isAdmin;

const handleLoginSubmit = (e: React.FormEvent) => {
  e.preventDefault();
  if (!nama.trim() || !nip.trim()) {
    setError('Nama dan NIP wajib diisi!');
    return;
  }

  // Jika nama adalah admin, NIP harus cocok
  if (matchedAdmin && cleanNip !== matchedAdmin.nip) {
    setError('Login gagal: NIP tidak sesuai.');
    return;
  }

  onLogin({
    nama: nama.trim(),
    nip: cleanNip,
    role: matchedAdmin ? 'admin' : role,
  });
};

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex items-center justify-center p-4 font-sans">
      <div className="max-w-md w-full bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
        <div className="text-center space-y-2">
          <h1 className="text-xl font-bold text-white">Sistem Penugasan Job Card</h1>
          <p className="text-xs text-slate-400">Silakan masukkan identitas Anda untuk masuk</p>
        </div>

        {error && (
          <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg text-rose-400 text-xs text-center">
            {error}
          </div>
        )}

        <form onSubmit={handleLoginSubmit} className="space-y-4 text-xs">
          <div>
            <label className="block text-slate-400 mb-1">Nama Lengkap</label>
            <div className="relative">
              <input
                type="text"
                value={nama}
                onChange={(e) => { setNama(e.target.value); setError(''); }}
                placeholder="Ketik nama Anda..."
                required
                className="w-full px-3 py-2.5 pl-9 bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-blue-500"
              />
              <User className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
            </div>
          </div>

          <div>
            <label className="block text-slate-400 mb-1">NIP</label>
            <div className="relative">
              <input
                type="text"
                value={nip}
                onChange={(e) => { setNip(e.target.value); setError(''); }}
                placeholder="Ketik NIP Anda..."
                required
                className="w-full px-3 py-2.5 pl-9 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono focus:outline-none focus:border-blue-500"
              />
              <KeyRound className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
            </div>
          </div>

          {/* Bagian Pilih Status yang tadinya langsung muncul, sekarang dibungkus kondisi showRoleSelector */}
          {showRoleSelector && (
            <div className="transition-all duration-300 space-y-1">
              <label className="block text-slate-400 mb-1">Pilih Status</label>
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => setRole('kabiro')}
                  className={`p-3 rounded-xl border flex flex-col items-center justify-center gap-1.5 transition cursor-pointer ${
                    role === 'kabiro' 
                      ? 'bg-blue-600/10 border-blue-500 text-blue-400 font-bold' 
                      : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <Briefcase className="w-4 h-4" />
                  <span>Kabiro</span>
                </button>

                <button
                  type="button"
                  onClick={() => setRole('outsourcing')}
                  className={`p-3 rounded-xl border flex flex-col items-center justify-center gap-1.5 transition cursor-pointer ${
                    role === 'outsourcing' 
                      ? 'bg-amber-600/10 border-amber-500 text-amber-400 font-bold' 
                      : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <HardHat className="w-4 h-4" />
                  <span>Outsourcing</span>
                </button>
              </div>
            </div>
          )}

          {isAdmin && (
            <div className="p-3 bg-purple-500/10 border border-purple-500/20 rounded-xl text-purple-300 text-center font-medium">
              ✨ Akses Admin Rendal Terdeteksi
            </div>
          )}

          <button
            type="submit"
            className="w-full py-3 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl transition cursor-pointer shadow-lg"
          >
            Masuk Sistem
          </button>
        </form>

        {/* <div className="text-[11px] text-slate-500 text-center border-t border-slate-800 pt-3">
          *Masukkan nama <b>Dyan Asih Purwanti</b> atau <b>Hashfi Moch Adam</b> untuk login sebagai Admin Rendal.
        </div> */}
      </div>
    </div>
  );
}