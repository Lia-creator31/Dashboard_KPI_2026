import React, { useState } from 'react';
import { Briefcase, HardHat, ShieldAlert } from 'lucide-react';

interface LoginPageProps {
  onLogin: (role: 'organik' | 'subkon' | 'rendal', meta?: any) => void;
}

export default function LoginPage({ onLogin }: LoginPageProps) {
  const [rendalPinModal, setRendalPinModal] = useState(false);
  const [pin, setPin] = useState('');

  const handleRendalSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    // PIN bebas atau langsung masuk karena proteksi password planner dihilangkan sesuai permintaan
    onLogin('rendal');
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex items-center justify-center p-4">
      <div className="max-w-md w-full bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-2xl">
        <div className="text-center space-y-2">
          <h1 className="text-xl font-bold text-white">Sistem Penugasan Job Card</h1>
          <p className="text-xs text-slate-400">Silakan pilih akses portal masuk Anda</p>
        </div>

        <div className="space-y-3">
          <button
            onClick={() => onLogin('organik')}
            className="w-full bg-slate-800 hover:bg-slate-750 border border-slate-700 hover:border-blue-500 rounded-xl p-4 text-left transition flex items-center justify-between group cursor-pointer"
          >
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-blue-500/10 text-blue-400 rounded-lg group-hover:bg-blue-600 group-hover:text-white transition">
                <Briefcase className="w-5 h-5" />
              </div>
              <div>
                <div className="font-semibold text-sm text-white">Portal Kabiro & Organik</div>
                <div className="text-[11px] text-slate-400">Pengisian form penugasan & pegawai organik</div>
              </div>
            </div>
          </button>

          <button
            onClick={() => onLogin('subkon')}
            className="w-full bg-slate-800 hover:bg-slate-750 border border-slate-700 hover:border-amber-500 rounded-xl p-4 text-left transition flex items-center justify-between group cursor-pointer"
          >
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-amber-500/10 text-amber-400 rounded-lg group-hover:bg-amber-600 group-hover:text-white transition">
                <HardHat className="w-5 h-5" />
              </div>
              <div>
                <div className="font-semibold text-sm text-white">Portal Outsourcing</div>
                <div className="text-[11px] text-slate-400">Akses mandiri personil mitra/subkon</div>
              </div>
            </div>
          </button>

          <button
            onClick={() => onLogin('rendal')}
            className="w-full bg-slate-800 hover:bg-slate-750 border border-slate-700 hover:border-purple-500 rounded-xl p-4 text-left transition flex items-center justify-between group cursor-pointer"
          >
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-purple-500/10 text-purple-400 rounded-lg group-hover:bg-purple-600 group-hover:text-white transition">
                <ShieldAlert className="w-5 h-5" />
              </div>
              <div>
                <div className="font-semibold text-sm text-white">Portal Rendal</div>
                <div className="text-[11px] text-slate-400">Master file, Planner & Release (Tanpa Password)</div>
              </div>
            </div>
          </button>
        </div>
      </div>
    </div>
  );
}