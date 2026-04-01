-- schema-admin-rls.sql
-- Run this once in the Supabase SQL editor.
-- Grants authenticated (logged-in) users write access to the sales table,
-- which is what admin.html needs for the flagged-sale edit, approve, and
-- add-sale flows. The anon key still only allows reads.

-- Allow logged-in admin users to insert new sales
create policy "authenticated insert sales"
  on sales
  for insert
  to authenticated
  with check (true);

-- Allow logged-in admin users to update existing sales
-- (covers edit + approve: sets flagged=false, status='verified')
create policy "authenticated update sales"
  on sales
  for update
  to authenticated
  using (true)
  with check (true);
