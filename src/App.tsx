import { useState } from 'react';
import LoginPage from './pages/LoginPage';
import RendalPage from './pages/RendalPage';
import KabiroPage from './pages/KabiroPage';
import OutsourcingPage from './pages/OutsourcingPage';

export interface UserSession {
  nama: string;
  nip: string;
  role: 'kabiro' | 'outsourcing' | 'admin' | 'rekap';
}

const SESSION_STORAGE_KEY = 'kpi_active_session_v1';

export default function App() {
  const [session, setSession] = useState<UserSession | null>(() => {
    try {
      const saved = sessionStorage.getItem(SESSION_STORAGE_KEY);
      return saved ? (JSON.parse(saved) as UserSession) : null;
    } catch {
      return null;
    }
  });

  const handleSetSession = (userData: UserSession | null) => {
    setSession(userData);
    try {
      if (userData) sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(userData));
      else sessionStorage.removeItem(SESSION_STORAGE_KEY);
    } catch {}
  };

  // 1. Jika belum login, arahkan ke halaman login
  if (!session) {
    return <LoginPage onLogin={(userData) => handleSetSession(userData)} />;
  }

  // 2. Jika login sebagai Admin (Rendal), buka RendalPage (Akses Penuh: Upload Excel, Planner, dll)
  if (session.role === 'admin') {
    return <RendalPage user={session} onLogout={() => handleSetSession(null)} />;
  }

  // 3. Jika login sebagai Kabiro, buka KabiroPage (Hanya bisa lihat anggota & isi form, TIDAK bisa upload atau approve planner)
  if (session.role === 'kabiro') {
    return <KabiroPage user={session} onLogout={() => handleSetSession(null)} />;
  }

  // 4. Jika login sebagai Outsourcing, buka OutsourcingPage (Khusus untuk mengisi Timesheet dari Work Order)
  if (session.role === 'outsourcing') {
    return <OutsourcingPage user={session} onLogout={() => handleSetSession(null)} />; 
  }

    // 2b. Role rekap: langsung ke Rekap Absensi Divisi (ditangani di RendalPage)
  if (session.role === 'rekap') {
    return <RendalPage user={session} onLogout={() => handleSetSession(null)} />;
  }

  // Fallback jika role tidak dikenali
  return null;
}