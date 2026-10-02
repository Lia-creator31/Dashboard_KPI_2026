import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

if (!supabaseUrl || !supabaseAnonKey) {
  console.error(
    'VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diisi di file .env'
  );
}

export const supabase = createClient(
  supabaseUrl ?? 'https://placeholder.supabase.co',
  supabaseAnonKey ?? 'placeholder-key'
);

export interface JobCard {
  id: string;
  biro_id: string;
  biro_name: string;
  personil_name: string;
  project: string;
  task_name: string;
  start_date: string;
  end_date: string;
  pic: string;
  jo: string;
  kode_jc: string;
  status: string;
  created_at: string;
  rev?: string;
  release?: string;
}