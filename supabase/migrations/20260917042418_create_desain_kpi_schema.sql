/*
# Create database schema for Rekapitulasi DESAIN 2026

This migration creates the complete database schema for the KPI recapitulation
dashboard ("Rekapitulasi DESAIN 2026"). The app manages departments, biros
(sub-units), job cards (task assignments), and monthly KPI records for each
employee. There is no sign-in screen — the app is a single-tenant internal tool,
so all policies allow anon + authenticated access.

## Tables created:
1. `departments` — 6 engineering design departments (Perencanaan, Desain Dasar, etc.)
2. `biros` — sub-units within each department (e.g. Biro Desain Struktur Lambung)
3. `job_cards` — task assignments linked to a biro, with project code, task name,
   start/end dates, PIC, and JO number. This replaces the current localStorage storage.
4. `kpi_records` — monthly KPI data per employee per biro, including effective hours,
   overtime, idle, timesheet regular/overtime, late count, sick days, and IPM score.

## Security:
- RLS enabled on all tables.
- All policies use `TO anon, authenticated` with `USING (true)` / `WITH CHECK (true)`
  because this is a single-tenant internal tool with no sign-in screen. The data is
  intentionally shared across all users of the application.
*/

-- ============================================================
-- 1. DEPARTMENTS
-- ============================================================
CREATE TABLE IF NOT EXISTS departments (
  id text PRIMARY KEY,
  name text NOT NULL,
  code text,
  icon text NOT NULL DEFAULT 'Building2',
  description text,
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE departments ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_departments" ON departments;
CREATE POLICY "anon_select_departments" ON departments FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_departments" ON departments;
CREATE POLICY "anon_insert_departments" ON departments FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_departments" ON departments;
CREATE POLICY "anon_update_departments" ON departments FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_departments" ON departments;
CREATE POLICY "anon_delete_departments" ON departments FOR DELETE
  TO anon, authenticated USING (true);

-- ============================================================
-- 2. BIROS
-- ============================================================
CREATE TABLE IF NOT EXISTS biros (
  id text PRIMARY KEY,
  department_id text NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
  name text NOT NULL,
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE biros ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_biros" ON biros;
CREATE POLICY "anon_select_biros" ON biros FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_biros" ON biros;
CREATE POLICY "anon_insert_biros" ON biros FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_biros" ON biros;
CREATE POLICY "anon_update_biros" ON biros FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_biros" ON biros;
CREATE POLICY "anon_delete_biros" ON biros FOR DELETE
  TO anon, authenticated USING (true);

CREATE INDEX IF NOT EXISTS idx_biros_department_id ON biros(department_id);

-- ============================================================
-- 3. JOB CARDS
-- ============================================================
CREATE TABLE IF NOT EXISTS job_cards (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  biro_id text NOT NULL REFERENCES biros(id) ON DELETE CASCADE,
  personil_name text NOT NULL,
  project_code text NOT NULL,
  task_name text NOT NULL,
  start_date date,
  end_date date,
  pic text,
  jo text,
  created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE job_cards ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_job_cards" ON job_cards;
CREATE POLICY "anon_select_job_cards" ON job_cards FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_job_cards" ON job_cards;
CREATE POLICY "anon_insert_job_cards" ON job_cards FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_job_cards" ON job_cards;
CREATE POLICY "anon_update_job_cards" ON job_cards FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_job_cards" ON job_cards;
CREATE POLICY "anon_delete_job_cards" ON job_cards FOR DELETE
  TO anon, authenticated USING (true);

CREATE INDEX IF NOT EXISTS idx_job_cards_biro_id ON job_cards(biro_id);
CREATE INDEX IF NOT EXISTS idx_job_cards_personil_name ON job_cards(personil_name);

-- ============================================================
-- 4. KPI RECORDS (monthly recapitulation per employee)
-- ============================================================
CREATE TABLE IF NOT EXISTS kpi_records (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  biro_id text NOT NULL REFERENCES biros(id) ON DELETE CASCADE,
  nip text NOT NULL,
  nama text NOT NULL,
  month text NOT NULL,
  effective_hour numeric DEFAULT 0,
  overtime_hour numeric DEFAULT 0,
  idle_hour numeric DEFAULT 0,
  timesheet_reguler numeric DEFAULT 0,
  timesheet_overtime numeric DEFAULT 0,
  terlambat numeric DEFAULT 0,
  sakit numeric DEFAULT 0,
  ipm numeric DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE(biro_id, nip, month)
);

ALTER TABLE kpi_records ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_kpi_records" ON kpi_records;
CREATE POLICY "anon_select_kpi_records" ON kpi_records FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_kpi_records" ON kpi_records;
CREATE POLICY "anon_insert_kpi_records" ON kpi_records FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_kpi_records" ON kpi_records;
CREATE POLICY "anon_update_kpi_records" ON kpi_records FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_kpi_records" ON kpi_records;
CREATE POLICY "anon_delete_kpi_records" ON kpi_records FOR DELETE
  TO anon, authenticated USING (true);

CREATE INDEX IF NOT EXISTS idx_kpi_records_biro_id ON kpi_records(biro_id);
CREATE INDEX IF NOT EXISTS idx_kpi_records_month ON kpi_records(month);
CREATE INDEX IF NOT EXISTS idx_kpi_records_biro_month ON kpi_records(biro_id, month);
