-- ════════════════════════════════════════════════════════════════════════════
-- schema-live-estimates.sql — Pre-computed estimation table
--
-- Run once in the Supabase SQL editor. Creates the live_estimates table
-- that replaces browser-side estimate-engine.js computation.
--
-- Refreshed daily by GAS refreshLiveEstimates() at 4 AM.
-- Each run replaces all rows with fresh estimates from buildEstimates().
-- ════════════════════════════════════════════════════════════════════════════

CREATE TABLE IF NOT EXISTS live_estimates (
  -- ── Primary key ──────────────────────────────────────────────────────
  id              uuid DEFAULT gen_random_uuid() PRIMARY KEY,

  -- ── Identity ─────────────────────────────────────────────────────────
  set_name        text    NOT NULL,
  player          text    NOT NULL,
  variant         text    NOT NULL,
  set_math_id     uuid    REFERENCES set_math(id),

  -- ── Core estimate ────────────────────────────────────────────────────
  estimated_price numeric NOT NULL,
  confidence      text,                  -- 'High', 'Medium-High', 'Medium', etc.
  method          text,                  -- 'direct-comp', 'ref-dist+ratio', etc.
  is_refreshed    boolean DEFAULT false, -- true = blended with real sales

  -- ── Refreshed-only fields (NULL when is_refreshed = false) ───────────
  own_price       numeric,               -- real sales price (blended source)
  freshness       numeric,               -- 0.0–1.0 (exponential decay)
  model_estimate  numeric,               -- pure model price before blending

  -- ── Engine reasoning metadata ────────────────────────────────────────
  evidence_weight numeric,               -- 0.0–1.0 (set-level non-base evidence)

  -- PVS (Player Value Scale)
  pvs_percentile  numeric,               -- 0.0–1.0
  pvs_rank        integer,
  pvs_n_players   integer,
  pvs_implied_base numeric,
  pvs_source      text,                  -- 'base', 'inferred', 'cross-set-base', etc.

  -- Reference distribution
  ref_dist_info   jsonb,                 -- { refSets, sameLeague, nPoints, rangeMin,
                                         --   rangeMax, extrapolated, extrapolatedUp,
                                         --   targetPop, exactPopMatch, scaleFactor }
  range_estimate  numeric,               -- ref-dist interpolated price

  -- Ratio anchors
  ratio_estimate  numeric,               -- weighted ratio price
  anchors         jsonb,                 -- array of anchor objects

  -- Direct comp
  comp_used       boolean DEFAULT false,
  comp_data       jsonb,                 -- { price, n, comps: [{price, set, sales, ...}] }

  -- ── Engine version / metadata ────────────────────────────────────────
  engine_version  integer,
  computed_at     timestamptz DEFAULT now(),

  -- ── Uniqueness: one row per set/player/variant ───────────────────────
  CONSTRAINT live_estimates_unique UNIQUE (set_name, player, variant)
);

-- ── Indexes ─────────────────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_live_estimates_set
  ON live_estimates (set_name);

CREATE INDEX IF NOT EXISTS idx_live_estimates_lookup
  ON live_estimates (set_name, player, variant);

CREATE INDEX IF NOT EXISTS idx_live_estimates_player
  ON live_estimates (player, variant);

-- ── Row-Level Security ──────────────────────────────────────────────────
-- Public read access (anon key), write only via service role (GAS).
ALTER TABLE live_estimates ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Public read access"
  ON live_estimates FOR SELECT
  USING (true);

CREATE POLICY "Service role full access"
  ON live_estimates FOR ALL
  USING (auth.role() = 'service_role');
