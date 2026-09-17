-- 1. Tambahkan kolom yang dibutuhkan
ALTER TABLE job_cards 
  ADD COLUMN IF NOT EXISTS biro_name text,
  ADD COLUMN IF NOT EXISTS project text,
  ADD COLUMN IF NOT EXISTS task_name text,
  ADD COLUMN IF NOT EXISTS start_date text,
  ADD COLUMN IF NOT EXISTS end_date text,
  ADD COLUMN IF NOT EXISTS pic text,
  ADD COLUMN IF NOT EXISTS jo text,
  ADD COLUMN IF NOT EXISTS kode_jc text,
  ADD COLUMN IF NOT EXISTS status text DEFAULT 'pending',
  ADD COLUMN IF NOT EXISTS created_at timestamptz DEFAULT now();

-- 2. Atur izin akses (RLS) agar form web bisa membaca dan menyimpan data
ALTER TABLE job_cards ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Public access to job_cards" ON job_cards;

CREATE POLICY "Public access to job_cards" 
ON job_cards 
FOR ALL 
TO anon, authenticated 
USING (true) 
WITH CHECK (true);