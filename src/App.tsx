import { useState } from 'react';
import LoginPage from './pages/LoginPage';
import RendalPage from './pages/RendalPage';

export interface UserSession {
  nama: string;
  nip: string;
  role: 'kabiro' | 'outsourcing' | 'admin';
}

export default function App() {
  const [session, setSession] = useState<UserSession | null>(null);

  // Jika belum login, arahkan ke halaman login
  if (!session) {
    return <LoginPage onLogin={(userData) => setSession(userData)} />;
  }

  // Jika login sebagai Admin (Dyan Asih Purwanti / Hashfi Moch Adam), buka RendalPage dengan akses penuh
  if (session.role === 'admin') {
    return <RendalPage user={session} onLogout={() => setSession(null)} />;
  }

  // Jika login sebagai Kabiro atau Outsourcing (tetap masuk ke RendalPage dengan pembatasan hak akses sesuai peran)
  return <RendalPage user={session} onLogout={() => setSession(null)} />;
}