import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

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
}
