-- schema-price-intelligence-view.sql
--
-- Replaces the old static price_intelligence table (which was populated by the
-- now-retired listings pipeline) with a live VIEW that computes aggregated sales
-- stats directly from the sales table.
--
-- Benefits:
--   • Always up-to-date — no rebuild step, no GAS timeout risk
--   • Both the browser (via get_price_intelligence RPC) and GAS query the same data
--   • No code change needed in GAS: sbGet_('price_intelligence', ...) still works
--
-- Run this once in Supabase SQL editor.
-- ─────────────────────────────────────────────────────────────────────────────

-- 1. Drop price_intelligence (was a static table on first run; now a view on re-runs).
--    CASCADE also drops the get_price_intelligence() RPC that depends on it.
DROP VIEW  IF EXISTS price_intelligence CASCADE;
DROP TABLE IF EXISTS price_intelligence CASCADE;

-- 2. Create the live view.
--    Columns match exactly what loadAllData_() in GAS and get_price_intelligence()
--    in the browser expect.
CREATE VIEW price_intelligence AS
SELECT
  s.set_id,
  st.name                                                                        AS set_name,
  s.player,
  s.variant,
  sm.id                                                                          AS set_math_id,

  -- All-time stats
  COUNT(*)                                                                       AS total_sales,
  ROUND(AVG(s.price)::numeric, 2)                                               AS avg_price,

  -- 30-day window
  ROUND(AVG(s.price) FILTER (WHERE s.date >= CURRENT_DATE - 30)::numeric, 2)   AS avg_30d,
  COUNT(*)          FILTER (WHERE s.date >= CURRENT_DATE - 30)                  AS sales_30d,

  -- 90-day window
  ROUND(AVG(s.price) FILTER (WHERE s.date >= CURRENT_DATE - 90)::numeric, 2)   AS avg_90d,
  COUNT(*)          FILTER (WHERE s.date >= CURRENT_DATE - 90)                  AS sales_90d,

  -- Most recent sale
  MAX(s.date)                                                                    AS last_sale_date,
  (
    SELECT  s2.price
    FROM    sales s2
    WHERE   s2.set_id  = s.set_id
      AND   s2.player  = s.player
      AND   s2.variant = s.variant
      AND   (s2.status IS NULL OR s2.status NOT IN ('exclude', 'invalid'))
      AND   s2.flagged IS NOT TRUE
      AND   s2.price > 0
    ORDER BY s2.date DESC, s2.created_at DESC
    LIMIT 1
  )                                                                              AS last_sale_price

FROM  sales  s
JOIN  sets   st ON st.id = s.set_id AND st.active = true
LEFT JOIN set_math sm
       ON sm.set_id  = s.set_id
      AND sm.player  = s.player
      AND sm.variant = s.variant

WHERE s.player  IS NOT NULL
  AND s.player  != 'Unknown'
  AND s.variant IS NOT NULL
  AND (s.status IS NULL OR s.status NOT IN ('exclude', 'invalid'))
  AND s.flagged IS NOT TRUE
  AND s.price   IS NOT NULL
  AND s.price   > 0

GROUP BY s.set_id, st.name, s.player, s.variant, sm.id;


-- 3. Recreate the get_price_intelligence() RPC so the browser call keeps working.
--    Must DROP first — Postgres won't let you change return type with CREATE OR REPLACE.
DROP FUNCTION IF EXISTS get_price_intelligence();

CREATE FUNCTION get_price_intelligence()
RETURNS TABLE (
  set_id          uuid,
  set_name        text,
  player          text,
  variant         text,
  set_math_id     uuid,
  total_sales     bigint,
  avg_price       numeric,
  avg_30d         numeric,
  sales_30d       bigint,
  avg_90d         numeric,
  sales_90d       bigint,
  last_sale_date  date,
  last_sale_price numeric
)
LANGUAGE sql STABLE
AS $$
  SELECT * FROM price_intelligence;
$$;
