/*
# Create private master-files Storage bucket and anon policies

1. Storage Bucket
- Create the private `master-files` bucket.
- Files in this bucket are not publicly accessible by URL.

2. Storage Policies
- Allow the `anon` role to SELECT objects in `master-files`.
- Allow the `anon` role to INSERT objects in `master-files`.
- Allow the `anon` role to UPDATE objects in `master-files`.
- Policies are limited to this bucket and do not affect other Storage buckets.

3. Important Notes
- This is an intentionally shared bucket for the application's anonymous client.
- No application source code is changed by this migration.
- The migration is safe to re-run: the bucket is inserted only if missing and
  policies are replaced by name before being recreated.
*/

INSERT INTO storage.buckets (id, name, public)
VALUES ('master-files', 'master-files', false)
ON CONFLICT (id) DO UPDATE SET public = false;

DROP POLICY IF EXISTS "anon_select_master_files" ON storage.objects;
CREATE POLICY "anon_select_master_files"
ON storage.objects
FOR SELECT
TO anon
USING (bucket_id = 'master-files');

DROP POLICY IF EXISTS "anon_insert_master_files" ON storage.objects;
CREATE POLICY "anon_insert_master_files"
ON storage.objects
FOR INSERT
TO anon
WITH CHECK (bucket_id = 'master-files');

DROP POLICY IF EXISTS "anon_update_master_files" ON storage.objects;
CREATE POLICY "anon_update_master_files"
ON storage.objects
FOR UPDATE
TO anon
USING (bucket_id = 'master-files')
WITH CHECK (bucket_id = 'master-files');