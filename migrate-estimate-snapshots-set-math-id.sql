-- migrate-estimate-snapshots-set-math-id.sql
--
-- Step 1: Add set_math_id column to estimate_snapshots
-- Step 2: Backfill from set_math JOIN sets
-- Step 3: Re-run schema-price-intelligence-view.sql (updated separately)
--
-- Run this once in the Supabase SQL editor.
-- ─────────────────────────────────────────────────────────────────────────────

-- 1. Add the column (no-op if it already exists)
ALTER TABLE estimate_snapshots
  ADD COLUMN IF NOT EXISTS set_math_id uuid REFERENCES set_math(id);

-- 2. Backfill: match on set_name → sets.name → sets.id → set_math
--    Joins: estimate_snapshots.set_name = sets.name
--           set_math.set_id = sets.id AND set_math.player = es.player AND set_math.variant = es.variant
UPDATE estimate_snapshots es
SET    set_math_id = sm.id
FROM   set_math sm
JOIN   sets     s  ON s.id = sm.set_id
WHERE  es.set_name = s.name
  AND  es.player   = sm.player
  AND  es.variant  = sm.variant
  AND  es.set_math_id IS NULL;   -- only rows not yet populated

-- 3. Verify (optional — paste into a second query)
--    SELECT COUNT(*) AS total,
--           COUNT(set_math_id) AS with_id,
--           COUNT(*) - COUNT(set_math_id) AS missing
--    FROM   estimate_snapshots;
