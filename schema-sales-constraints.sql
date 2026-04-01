-- schema-sales-constraints.sql
--
-- Adds the UNIQUE constraint on sales.sheets_row_index that the GAS upsert
-- depends on (ON CONFLICT sheets_row_index).
--
-- Run once in Supabase SQL editor.
-- ─────────────────────────────────────────────────────────────────────────────

-- Step 1: Remove any duplicate sheets_row_index values before adding the constraint.
-- Keeps the most recently created row for each duplicate index; deletes the rest.
-- NULL values are ignored (UNIQUE allows multiple NULLs in Postgres).
DELETE FROM sales a
USING  sales b
WHERE  a.sheets_row_index IS NOT NULL
  AND  a.sheets_row_index = b.sheets_row_index
  AND  a.created_at < b.created_at;

-- Step 2: Add the unique constraint.
ALTER TABLE sales
  ADD CONSTRAINT sales_sheets_row_index_key UNIQUE (sheets_row_index);
