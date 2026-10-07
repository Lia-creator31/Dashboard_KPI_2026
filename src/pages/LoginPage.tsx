```tsx
import React, { useState } from 'react';
import { KeyRound } from 'lucide-react';
import { UserSession } from '../App';

interface LoginPageProps {
  onLogin: (userData: UserSession) => void;
}

// Admin hanya berdasarkan NIP
const ADMIN_ACCOUNTS = [
  { nama: 'dyan asih purwanti', nip: '104134283' },
  { nama: 'hashfi moch adam', nip: '105264937' },
];

export default function LoginPage({ onLogin }: LoginPageProps) {
  const [nip, setNip] = useState('');
  const [error, setError] = useState('');

  const cleanNip = nip.trim();

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!cleanNip) {
      setError('NIP wajib diisi!');
      return;
    }

    // Cek apakah NIP merupakan admin
    const matchedAdmin = ADMIN_ACCOUNTS.find(
      (admin) => admin.nip === cleanNip
    );

    // Jika admin
    if (matchedAdmin) {
      onLogin({
        nama: matchedAdmin.nama,
        nip: cleanNip,
        role: 'admin',
      });

      return;
    }

    // Untuk user biasa:
    // Nanti NIP akan dicari ke database Supabase
    // untuk mendapatkan nama dan role.
    
    // SEMENTARA:
    // Jika belum menggunakan Supabase untuk validasi user,
    // bisa diarahkan ke logic user biasa di sini.

    setError('NIP tidak ditemukan. Silakan periksa kembali NIP Anda.');
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex items-center justify-center p-4 font-sans">
      <div className="max-w-md w-full bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">

        {/* Header */}
        <div className="text-center space-y-2">
          <h1 className="text-xl font-bold text-white">
            Sistem Penugasan Job Card
          </h1>

          <p className="text-xs text-slate-400">
            Masukkan NIP Anda untuk masuk ke sistem
          </p>
        </div>

        {/* Error */}
        {error && (
          <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg text-rose-400 text-xs text-center">
            {error}
          </div>
        )}

        {/* Form */}
        <form
          onSubmit={handleLoginSubmit}
          className="space-y-4 text-xs"
        >

          {/* NIP */}
          <div>
            <label className="block text-slate-400 mb-1">
              NIP
            </label>

            <div className="relative">
              <input
                type="text"
                value={nip}
                onChange={(e) => {
                  setNip(e.target.value);
                  setError('');
                }}
                placeholder="Masukkan NIP Anda..."
                required
                autoFocus
                className="w-full px-3 py-3 pl-9 bg-slate-950 border border-slate-700 rounded-lg text-white font-mono focus:outline-none focus:border-blue-500"
              />

              <KeyRound className="w-4 h-4 text-slate-400 absolute left-3 top-3.5" />
            </div>
          </div>

          {/* Admin Info */}
          {ADMIN_ACCOUNTS.some(
            (admin) => admin.nip === cleanNip
          ) && (
            <div className="p-3 bg-purple-500/10 border border-purple-500/20 rounded-xl text-purple-300 text-center font-medium">
              ✨ Akses Admin Rendal Terdeteksi
            </div>
          )}

          {/* Login */}
          <button
            type="submit"
            className="w-full py-3 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl transition cursor-pointer shadow-lg"
          >
            Masuk Sistem
          </button>

        </form>

      </div>
    </div>
  );
}
```