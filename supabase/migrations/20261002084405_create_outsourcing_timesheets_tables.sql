/*
# Create outsourcing timesheet tables

1. New Tables
- `outsourcing_timesheets` stores the header of each outsourcing timesheet.
  - `id` is the UUID primary key.
  - `code` is a required unique timesheet code.
  - `user_id` and `user_name` identify the person creating the timesheet.
  - `date_start` and `date_end` define the timesheet period.
  - `unit_kerja`, `divisi`, and `status` use the requested defaults.
  - `created_at` records creation time in UTC.
- `outsourcing_timesheet_lines` stores work entries belonging to a timesheet.
  - `timesheet_id` references `outsourcing_timesheets` and cascades on parent deletion.
  - Work order, description, date, day, and hour fields store each line item.
  - `created_at` records creation time in UTC.

2. Security
- Row Level Security is enabled on both tables.
- The current application does not use Supabase authentication for these records,
  so anon and authenticated roles receive shared CRUD access, matching the
  existing single-tenant application pattern.

3. Important Notes
- Tables and policies use idempotent statements so this migration can be safely re-run.
- No existing tables, columns, or application files are modified.
*/

CREATE TABLE IF NOT EXISTS public.outsourcing_timesheets (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  code text UNIQUE NOT NULL,
  user_id text NOT NULL,
  user_name text NOT NULL,
  date_start date NOT NULL,
  date_end date NOT NULL,
  unit_kerja text DEFAULT 'Biro Dukungan & Administrasi',
  divisi text DEFAULT '71000 - Divisi Desain',
  status text DEFAULT 'Draft',
  created_at timestamp with time zone DEFAULT timezone('utc'::text, now())
);

CREATE TABLE IF NOT EXISTS public.outsourcing_timesheet_lines (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  timesheet_id uuid REFERENCES public.outsourcing_timesheets(id) ON DELETE CASCADE,
  date date NOT NULL,
  work_order_id text,
  work_order_code text,
  description text,
  effective_hours numeric DEFAULT 0,
  overtime_hours numeric DEFAULT 0,
  day text,
  created_at timestamp with time zone DEFAULT timezone('utc'::text, now())
);

ALTER TABLE public.outsourcing_timesheets ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.outsourcing_timesheet_lines ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_outsourcing_timesheets" ON public.outsourcing_timesheets;
CREATE POLICY "anon_select_outsourcing_timesheets"
ON public.outsourcing_timesheets FOR SELECT
TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_outsourcing_timesheets" ON public.outsourcing_timesheets;
CREATE POLICY "anon_insert_outsourcing_timesheets"
ON public.outsourcing_timesheets FOR INSERT
TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_outsourcing_timesheets" ON public.outsourcing_timesheets;
CREATE POLICY "anon_update_outsourcing_timesheets"
ON public.outsourcing_timesheets FOR UPDATE
TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_outsourcing_timesheets" ON public.outsourcing_timesheets;
CREATE POLICY "anon_delete_outsourcing_timesheets"
ON public.outsourcing_timesheets FOR DELETE
TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_select_outsourcing_timesheet_lines" ON public.outsourcing_timesheet_lines;
CREATE POLICY "anon_select_outsourcing_timesheet_lines"
ON public.outsourcing_timesheet_lines FOR SELECT
TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_outsourcing_timesheet_lines" ON public.outsourcing_timesheet_lines;
CREATE POLICY "anon_insert_outsourcing_timesheet_lines"
ON public.outsourcing_timesheet_lines FOR INSERT
TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_outsourcing_timesheet_lines" ON public.outsourcing_timesheet_lines;
CREATE POLICY "anon_update_outsourcing_timesheet_lines"
ON public.outsourcing_timesheet_lines FOR UPDATE
TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_outsourcing_timesheet_lines" ON public.outsourcing_timesheet_lines;
CREATE POLICY "anon_delete_outsourcing_timesheet_lines"
ON public.outsourcing_timesheet_lines FOR DELETE
TO anon, authenticated USING (true);

CREATE INDEX IF NOT EXISTS outsourcing_timesheet_lines_timesheet_id_idx
ON public.outsourcing_timesheet_lines (timesheet_id);

CREATE INDEX IF NOT EXISTS outsourcing_timesheets_user_id_idx
ON public.outsourcing_timesheets (user_id);