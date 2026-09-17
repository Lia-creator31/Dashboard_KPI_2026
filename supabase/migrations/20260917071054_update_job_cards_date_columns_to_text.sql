/*
# Update job_cards column types for start_date and end_date

The job_cards table already has all requested columns (biro_name, project,
task_name, start_date, end_date, pic, jo, kode_jc, status, created_at) from
a prior migration. However, start_date and end_date were created as `date`
type. This migration changes them to `text` to match the requested schema,
so the app can store flexible date string formats.

## Changes:
1. `start_date`: changed from `date` to `text`
2. `end_date`: changed from `date` to `text`

## Safety:
- The table currently has 0 rows, so no data is lost during the type conversion.
- RLS policies for SELECT and INSERT (anon + authenticated) already exist
  from the original migration and remain unchanged.

## RLS status (unchanged, already in place):
- `anon_select_job_cards` — SELECT TO anon, authenticated
- `anon_insert_job_cards` — INSERT TO anon, authenticated
- `anon_update_job_cards` — UPDATE TO anon, authenticated
- `anon_delete_job_cards` — DELETE TO anon, authenticated
*/

ALTER TABLE job_cards ALTER COLUMN start_date TYPE text USING start_date::text;
ALTER TABLE job_cards ALTER COLUMN end_date TYPE text USING end_date::text;
