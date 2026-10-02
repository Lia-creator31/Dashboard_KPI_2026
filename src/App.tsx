import { useState } from 'react';
import LoginPage from './pages/LoginPage';
import RendalPage from './pages/RendalPage';
import KabiroPage from './pages/KabiroPage'; // <-- TAMBAHKAN IMPORT INI
// import OutsourcingPage from './pages/OutsourcingPage'; // <-- Nanti tambahkan jika file Outsourcing sudah dibuat

export interface UserSession {
  nama: string;
  nip: string;
  role: 'kabiro' | 'outsourcing' | 'admin';
}

export default function App() {
  const [session, setSession] = useState<UserSession | null>(null);

  // 1. Jika belum login, arahkan ke halaman login
  if (!session) {
    return <LoginPage onLogin={(userData) => setSession(userData)} />;
  }

  // 2. Jika login sebagai Admin (Rendal), buka RendalPage (Akses Penuh: Upload Excel, Planner, dll)
  if (session.role === 'admin') {
    return <RendalPage user={session} onLogout={() => setSession(null)} />;
  }

  // 3. Jika login sebagai Kabiro, buka KabiroPage (Hanya bisa lihat anggota & isi form, TIDAK bisa upload atau approve planner)
  if (session.role === 'kabiro') {
    return <KabiroPage user={session} onLogout={() => setSession(null)} />;
  }

  // 4. Jika login sebagai Outsourcing
  // (Sementara diarahkan ke KabiroPage atau buat file OutsourcingPage.tsx terpisah nanti)
  if (session.role === 'outsourcing') {
    // return <OutsourcingPage user={session} onLogout={() => setSession(null)} />; 
    return <KabiroPage user={session} onLogout={() => setSession(null)} />; 
  }

  // Fallback jika role tidak dikenali
  return null;
}