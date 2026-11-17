import { defineConfig, Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath, URL } from 'node:url';

function subkonOvertimePlugin(): Plugin {
  return {
    name: 'subkon-overtime-connector',
    configureServer(server) {
      server.middlewares.use('/api/subkon-overtime', async (req, res) => {
        if (req.method !== 'POST' && req.method !== 'GET') {
          res.statusCode = 405;
          res.end(JSON.stringify({ success: false, error: 'Method Not Allowed' }));
          return;
        }

        let bodyStr = '';
        req.on('data', chunk => { bodyStr += chunk; });
        req.on('end', async () => {
          try {
            let payload: any = {};
            try { payload = bodyStr ? JSON.parse(bodyStr) : {}; } catch {}
            const username = payload.username || 'pwdesain';
            const password = payload.password || 'pengawas123';
            const baseUrl = 'https://subkon.pal.co.id';

            const cookieJar: Record<string, string> = {};
            const updateCookies = (headers: Headers) => {
              const setCookie = (headers as any).getSetCookie
                ? (headers as any).getSetCookie()
                : (headers.get('set-cookie') ? [headers.get('set-cookie')!] : []);
              for (const sc of setCookie) {
                const first = String(sc).split(';')[0];
                const eqIdx = first.indexOf('=');
                if (eqIdx > 0) {
                  cookieJar[first.slice(0, eqIdx).trim()] = first.slice(eqIdx + 1).trim();
                }
              }
            };
            const cookieHeader = () =>
              Object.entries(cookieJar).map(([k, v]) => `${k}=${v}`).join('; ');

            // 1. Ambil halaman login untuk mendapatkan CSRF token / session cookie
            let csrfToken = '';
            try {
              const loginPageRes = await fetch(`${baseUrl}/login`, {
                method: 'GET',
                headers: {
                  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36',
                  'Accept': 'text/html,application/xhtml+xml,application/json',
                },
              });
              updateCookies(loginPageRes.headers);
              const html = await loginPageRes.text();
              const tokenMatch =
                html.match(/name=["']_token["']\s+value=["']([^"']+)["']/i) ||
                html.match(/meta\s+name=["']csrf-token["']\s+content=["']([^"']+)["']/i);
              if (tokenMatch) csrfToken = tokenMatch[1];
            } catch {}

            // 2. Kirim request login dengan username: pwdesain & password: pengawas123
            const loginEndpoints = ['/login', '/api/login', '/api/auth/login', '/auth/login'];
            let bearerToken = '';
            for (const ep of loginEndpoints) {
              try {
                const headers: Record<string, string> = {
                  'Content-Type': 'application/json',
                  'Accept': 'application/json, text/html, */*',
                  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36',
                };
                if (cookieHeader()) headers['Cookie'] = cookieHeader();
                if (csrfToken) headers['X-CSRF-TOKEN'] = csrfToken;
                if (cookieJar['XSRF-TOKEN']) {
                  headers['X-XSRF-TOKEN'] = decodeURIComponent(cookieJar['XSRF-TOKEN']);
                }

                const lRes = await fetch(`${baseUrl}${ep}`, {
                  method: 'POST',
                  headers,
                  redirect: 'manual',
                  body: JSON.stringify({
                    username,
                    email: username,
                    password,
                    ...(csrfToken ? { _token: csrfToken } : {}),
                  }),
                });
                updateCookies(lRes.headers);
                const text = await lRes.text();
                try {
                  const j = JSON.parse(text);
                  if (j.token || j.access_token || j.data?.token) {
                    bearerToken = j.token || j.access_token || j.data?.token;
                  }
                } catch {}
                if (lRes.status === 200 || lRes.status === 302) break;
              } catch {}
            }

            // 3. Ambil data dari fitur Modul Lembur -> Approval Lembur (/overtime)
            const records: any[] = [];
            const dataEndpoints = [
              '/api/overtime',
              '/api/overtime/approval',
              '/overtime?json=1',
              '/overtime',
            ];

            for (const ep of dataEndpoints) {
              try {
                const headers: Record<string, string> = {
                  'Accept': 'application/json, text/html, */*',
                  'X-Requested-With': 'XMLHttpRequest',
                  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36',
                };
                if (cookieHeader()) headers['Cookie'] = cookieHeader();
                if (bearerToken) headers['Authorization'] = `Bearer ${bearerToken}`;

                const dRes = await fetch(`${baseUrl}${ep}`, { headers });
                if (!dRes.ok) continue;
                const body = await dRes.text();

                // Coba parse sebagai JSON atau Inertia data-page
                let parsedJson: any = null;
                try {
                  parsedJson = JSON.parse(body);
                } catch {
                  const inertiaMatch = body.match(/data-page=["']([^"']+)["']/i);
                  if (inertiaMatch) {
                    const decoded = inertiaMatch[1]
                      .replace(/&quot;/g, '"')
                      .replace(/&#039;/g, "'")
                      .replace(/&amp;/g, '&');
                    try { parsedJson = JSON.parse(decoded); } catch {}
                  }
                }

                if (parsedJson) {
                  const candidateArrays = [
                    parsedJson.data,
                    parsedJson.records,
                    parsedJson.overtimes,
                    parsedJson.props?.overtimes?.data,
                    parsedJson.props?.overtimes,
                    parsedJson.props?.records?.data,
                    parsedJson.props?.records,
                    parsedJson.props?.data,
                  ];
                  for (const arr of candidateArrays) {
                    if (Array.isArray(arr) && arr.length > 0) {
                      arr.forEach((item: any) => {
                        records.push({
                          nip: item.nip || item.employee_nip || item.karyawan?.nip || '',
                          nama: item.nama || item.employee_name || item.name || item.karyawan?.nama || '',
                          date: item.date || item.tanggal || item.start_date || item.session_date || '',
                          durasiAktual: item.durasi_aktual ?? item.actual_duration ?? item.durasiAktual,
                          durasiAktualText: item.durasi_aktual_text || item.durasi_aktual || item.actual_duration || '',
                          totalLemburText: item.total_lembur || item.total_overtime || '',
                          statusVerifikasi: item.status_verifikasi || item.status || '',
                          catatanWbs: item.catatan || item.wbs || '',
                        });
                      });
                      break;
                    }
                  }
                }
                if (records.length > 0) break;
              } catch {}
            }

            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({
              success: true,
              account: username,
              records,
            }));
          } catch (err: any) {
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({
              success: false,
              error: err?.message || 'Connection error',
              records: [],
            }));
          }
        });
      });
    },
  };
}

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), subkonOvertimePlugin()],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  optimizeDeps: {
    exclude: ['lucide-react'],
  },
});
