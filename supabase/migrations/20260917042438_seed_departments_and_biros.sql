/*
# Seed departments and biros from existing application data

This migration inserts the 6 departments and 21 biros that are currently
hardcoded in the application's data.ts file. Using ON CONFLICT DO NOTHING so
re-running won't duplicate data.

## Data inserted:
- 6 departments (Perencanaan Desain, Desain Dasar, Struktur Lambung, 
  Struktur Permesinan, Perlengkapan Listrik, HPS)
- 21 biros across all departments, with sort_order matching their list position
*/

-- Departments
INSERT INTO departments (id, name, icon, sort_order) VALUES
  ('dept-1', 'Departemen Perencanaan Desain', 'Wrench', 1),
  ('dept-2', 'Departemen Desain Dasar', 'CircleDollarSign', 2),
  ('dept-3', 'Departemen Struktur dan Perlengkapan Lambung', 'Users', 3),
  ('dept-4', 'Departemen Struktur & Perlengkapan Permesinan', 'Truck', 4),
  ('dept-5', 'Departemen Perlengkapan Listrik & Elektronika', 'ShieldCheck', 5),
  ('dept-6', 'Departemen HPS', 'Laptop', 6)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-1 (Perencanaan Desain)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-101', 'dept-1', 'Biro Dokumen dan Perencanaan', 1),
  ('b-102', 'dept-1', 'Biro Dukungan Logistik Terpadu', 2),
  ('b-103', 'dept-1', 'Biro Dukungan & Administrasi', 3)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-2 (Desain Dasar)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-201', 'dept-2', 'Biro Desain Dasar Pengembangan Desain', 1),
  ('b-202', 'dept-2', 'Biro Desain Dasar Kapal Selam', 2),
  ('b-203', 'dept-2', 'Biro Desain Dasar Non Kapal', 3),
  ('b-204', 'dept-2', 'Biro Desain Dasar Kapal Permukaan', 4)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-3 (Struktur dan Perlengkapan Lambung)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-301', 'dept-3', 'Biro Desain Struktur Lambung', 1),
  ('b-302', 'dept-3', 'Biro Desain Akomodasi', 2),
  ('b-303', 'dept-3', 'Biro Desain Perlengkapan Lambung', 3),
  ('b-304', 'dept-3', 'Biro Desain Produksi Lambung', 4)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-4 (Struktur & Perlengkapan Permesinan)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-401', 'dept-4', 'Biro Sistem Propulsi', 1),
  ('b-402', 'dept-4', 'Biro Pengaturan Permesinan', 2),
  ('b-403', 'dept-4', 'Biro Sistem HVAC dan Permesinan Geladak', 3)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-5 (Perlengkapan Listrik & Elektronika)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-501', 'dept-5', 'Biro Sistem Kelistrikan', 1),
  ('b-502', 'dept-5', 'Biro Sistem Kontrol dan Otomasi', 2),
  ('b-503', 'dept-5', 'Biro Elektronika', 3)
ON CONFLICT (id) DO NOTHING;

-- Biros for dept-6 (HPS)
INSERT INTO biros (id, department_id, name, sort_order) VALUES
  ('b-601', 'dept-6', 'Biro HPS Material', 1),
  ('b-602', 'dept-6', 'Biro HPS Jasa & Investasi', 2)
ON CONFLICT (id) DO NOTHING;
