import { useState } from 'react';
import LoginPage from './pages/LoginPage';
import RendalPage from './pages/RendalPage';
import KabiroPage from './pages/KabiroPage';
import OutsourcingPage from './pages/OutsourcingPage';
import RekapAbsensiDivisiPage from './pages/RekapAbsensiDivisiPage'; // 1. Impor komponen Rekap Absensi

export interface UserSession {
  nama: string;
  nip: string;
  role: 'kabiro' | 'outsourcing' | 'admin' | 'rekap_absensi'; // 2. Tambahkan 'rekap_absensi' ke tipe role
}

export default function App() {
  const [session, setSession] = useState<UserSession | null>(null);

  // 1. Jika belum login, arahkan ke halaman login
  if (!session) {
    return <LoginPage onLogin={(userData) => setSession(userData)} />;
  }

  // 2. Jika login sebagai Admin (Rendal), buka RendalPage
  if (session.role === 'admin') {
    return <RendalPage user={session} onLogout={() => setSession(null)} />;
  }

  // 3. Jika login sebagai Kabiro, buka KabiroPage
  if (session.role === 'kabiro') {
    return <KabiroPage user={session} onLogout={() => setSession(null)} />;
  }

  // 4. Jika login sebagai Outsourcing, buka OutsourcingPage
  if (session.role === 'outsourcing') {
    return <OutsourcingPage user={session} onLogout={() => setSession(null)} />; 
  }

  // 5. Jika login dengan NIP Rekap Absensi, buka RekapAbsensiDivisiPage
  if (session.role === 'rekap_absensi') {
    return <RekapAbsensiDivisiPage user={session} onLogout={() => setSession(null)} />;
  }

  // Fallback jika role tidak dikenali
  return null;
}