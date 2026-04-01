-- ═══════════════════════════════════════════════════════════════════════════
-- GHOSTLADDER — Estimation Engine v2 Schema Additions
-- Run in Supabase SQL editor. Safe to run against existing DB:
-- no existing tables are modified.
--
-- New tables:
--   engine_params          — tunable parameters for estimate-engine.js
--   estimate_snapshots     — point-in-time predictions, resolved vs actuals
--   engine_calibration_log — tracks every parameter change + MAE impact
--
-- Existing tables left unchanged:
--   model_weights, prediction_log, calibration_log (legacy model history)
--   price_signals (still used; listing columns already there for future phase)
-- ═══════════════════════════════════════════════════════════════════════════


-- ─────────────────────────────────────────────────────────────────────────
-- 1. ENGINE_PARAMS
--    All constants currently hardcoded in estimate-engine.js, made explicit
--    and queryable. One row per version; exactly one row is_active = true.
--    model.html reads from this table instead of model_weights.
--
--    Future Listings integration: listing_blend, listing_floor_discount,
--    and supply_pressure columns are reserved here (NULL = feature off).
--    When ready, set their values on a new version row and flip is_active.
-- ─────────────────────────────────────────────────────────────────────────

CREATE TABLE public.engine_params (
  id         uuid    NOT NULL DEFAULT gen_random_uuid(),
  version    integer NOT NULL UNIQUE,           -- sequential: 1, 2, 3…
  tag        text,                              -- optional label e.g. "tau-90-to-60"
  is_active  boolean NOT NULL DEFAULT false,
  notes      text,
  applied_at timestamptz DEFAULT now(),
  created_at timestamptz DEFAULT now(),

  -- ── Freshness / staleness decay ───────────────────────────────────────
  -- freshness = e^(−days_since_last_sale / staleness_tau)
  -- At staleness_tau days, own-data weight ≈ 37% of the blend.
  -- Lower = model takes over faster. τ=60 means 63% model weight at 60d.
  staleness_tau         numeric NOT NULL DEFAULT 90,

  -- ── Evidence accumulation ─────────────────────────────────────────────
  -- evidenceWeight = 1 − e^(−evidence_k × recency_weighted_sales)
  -- Default 0.0231 → ~30 weighted non-base sales reaches ~50% confidence.
  -- Lower = needs more data before trusting the model; higher = trusts sparse data sooner.
  evidence_k            numeric NOT NULL DEFAULT 0.0231,

  -- ── salesStrength() recency weight scalars ────────────────────────────
  -- Applied when accumulating evidence from sales at different time horizons.
  -- 30d sales count most; all-time sales count least.
  weight_30d            numeric NOT NULL DEFAULT 3.0,
  weight_90d            numeric NOT NULL DEFAULT 1.5,
  weight_alltime        numeric NOT NULL DEFAULT 1.0,

  -- ── recencyWeight() confidence factors ───────────────────────────────
  -- How confident to be in a price point depending on recency of the data.
  -- 1.0 = have 30d sales; 0.7 = only 90d data; 0.4 = only all-time average.
  recency_conf_30d      numeric NOT NULL DEFAULT 1.0,
  recency_conf_90d      numeric NOT NULL DEFAULT 0.7,
  recency_conf_alltime  numeric NOT NULL DEFAULT 0.4,

  -- ── Rarity enforcement ────────────────────────────────────────────────
  -- Minimum multiplier enforced between adjacent variant rarity tiers.
  -- Prevents monotonicity violations (rarer card pricing below common card).
  rarity_premium        numeric NOT NULL DEFAULT 1.20,

  -- ── Reference distribution matching ──────────────────────────────────
  -- pop_proximity_floor: cross-set comps with match quality below this are discarded.
  -- Computed as: min(targetPop, matchPop) / max(targetPop, matchPop)
  pop_proximity_floor   numeric NOT NULL DEFAULT 0.60,
  -- League weights: same-league comps are trusted more than cross-league.
  league_same_weight    numeric NOT NULL DEFAULT 0.90,
  league_diff_weight    numeric NOT NULL DEFAULT 0.70,

  -- ── Evidence threshold for ratio blending ─────────────────────────────
  -- Ratio estimate is only included in the final blend when evidenceWeight
  -- exceeds this threshold. Prevents ratio noise on thin data sets.
  evidence_threshold    numeric NOT NULL DEFAULT 0.15,

  -- ── Confidence-tier quality gates (freshness blend) ───────────────────
  -- When a real-data cell is stale, these control how aggressively the model
  -- overrides the real price based on the model's confidence tier.
  -- effective_model_weight = (1 − freshness) × quality_[tier]
  quality_medium        numeric NOT NULL DEFAULT 0.85,
  quality_medium_low    numeric NOT NULL DEFAULT 0.65,
  quality_low           numeric NOT NULL DEFAULT 0.40,
  quality_very_low      numeric NOT NULL DEFAULT 0.20,

  -- ── Interpolation extrapolation caps ──────────────────────────────────
  -- Prevents log-linear interpolation from producing absurd prices when
  -- a player's percentile falls outside the reference distribution range.
  extrap_cap_low        numeric NOT NULL DEFAULT 0.10,   -- floor: 10% of dist minimum
  extrap_cap_high       numeric NOT NULL DEFAULT 10.0,   -- ceiling: 10× dist maximum

  -- ── Computed accuracy (updated as estimate_snapshots resolve) ─────────
  mae_all_time          numeric,      -- mean abs error across all resolved snapshots
  mae_30d               numeric,      -- MAE on snapshots resolved in last 30 days
  mae_by_confidence     jsonb,        -- e.g. {"Medium": 0.07, "Medium-Low": 0.14, "Low": 0.28, "Very Low": 0.45}
  mae_by_method         jsonb,        -- e.g. {"direct-comp": 0.06, "ref-dist": 0.12, "ratio": 0.22}
  total_snapshots       integer DEFAULT 0,
  resolved_snapshots    integer DEFAULT 0,

  -- ── Future: Listings signal integration (Phase 2) ─────────────────────
  -- These are NULL = feature off. Activate by setting values on a new version.
  -- listing_blend: weight for listing ask signal vs model estimate (0.0–1.0)
  listing_blend          numeric,
  -- listing_floor_discount: asks tend to run above realized prices.
  -- A discount of 0.92 means: treat a $100 ask as an ~$92 signal.
  listing_floor_discount numeric,
  -- supply_pressure: downward price adjustment per unit of excess supply.
  supply_pressure        numeric,

  CONSTRAINT engine_params_pkey PRIMARY KEY (id)
);

-- Enforce single active version at the database level
CREATE UNIQUE INDEX engine_params_one_active_idx
  ON public.engine_params (is_active)
  WHERE is_active = true;


-- ─────────────────────────────────────────────────────────────────────────
-- 2. ESTIMATE_SNAPSHOTS
--    Point-in-time predictions from estimate-engine.js, with enough metadata
--    to explain the estimate and measure accuracy when real sales come in.
--
--    One row per player+variant per snapshot run.
--    snapshot_run_id groups all cells produced in the same job execution,
--    so you can reconstruct a full "what did the whole set look like on date X".
--
--    Populated by: periodic Apps Script job (recommended: weekly).
--    Resolved by:  a separate job that checks price_intelligence for new
--                  sales after snapshot_at and writes actual_price + error_pct.
--
--    Forward-compatible: listing_price / listing_supply / listing_component
--    columns are NULL until Listings integration is live.
-- ─────────────────────────────────────────────────────────────────────────

CREATE TABLE public.estimate_snapshots (
  id              uuid NOT NULL DEFAULT gen_random_uuid(),
  snapshot_at     timestamptz NOT NULL DEFAULT now(),
  snapshot_run_id uuid NOT NULL,    -- shared UUID across all rows from one job run

  -- ── Subject ───────────────────────────────────────────────────────────
  set_name        text NOT NULL,
  player          text NOT NULL,
  variant         text NOT NULL,

  -- ── Prediction ────────────────────────────────────────────────────────
  estimated_price numeric NOT NULL,
  -- Confidence tier: 'Medium' | 'Medium-Low' | 'Low' | 'Very Low'
  confidence      text,
  -- How the estimate was produced: 'direct-comp' | 'ref-dist' | 'ratio' | 'fallback'
  method          text,

  -- ── Engine state at snapshot time ─────────────────────────────────────
  engine_version  integer NOT NULL,   -- FK → engine_params.version
  evidence_weight numeric,            -- 0–1, aggregate data confidence for the set
  n_ref_points    integer,            -- reference distribution point count
  ref_sets        text[],             -- which sets contributed reference points
  extrapolated    boolean DEFAULT false,
  extrap_dir      text,               -- 'up' | 'down' (if extrapolated = true)

  -- ── Real-data context at snapshot time ───────────────────────────────
  -- own_price: best available real sale price at snapshot time.
  -- NULL = pure gap-fill (no real data existed).
  own_price       numeric,
  freshness       numeric,            -- e^(−days_since_sale / tau) at snapshot time
  last_sale_date  date,
  -- is_refreshed = true means this cell had real data but model was blended in
  is_refreshed    boolean DEFAULT false,

  -- ── Future: Listings data at snapshot time ────────────────────────────
  -- Populated once listing_blend > 0 in engine_params.
  listing_price    numeric,           -- lowest_ask from price_signals at snapshot time
  listing_supply   integer,           -- supply_count from price_signals at snapshot time
  listing_component numeric,          -- fraction of estimated_price from listing signal

  -- ── Resolution ────────────────────────────────────────────────────────
  -- Filled in by resolve job when a new real sale is detected after snapshot_at.
  actual_price    numeric,
  actual_date     date,
  -- error_pct = ABS(actual_price − estimated_price) / actual_price
  error_pct       numeric,
  resolved_at     timestamptz,

  CONSTRAINT estimate_snapshots_pkey PRIMARY KEY (id),
  CONSTRAINT estimate_snapshots_engine_fkey
    FOREIGN KEY (engine_version) REFERENCES public.engine_params(version)
);

-- Indexes for the most common query patterns
CREATE INDEX est_snap_set_idx         ON public.estimate_snapshots (set_name, snapshot_at DESC);
CREATE INDEX est_snap_player_var_idx  ON public.estimate_snapshots (player, variant, snapshot_at DESC);
CREATE INDEX est_snap_run_idx         ON public.estimate_snapshots (snapshot_run_id);
CREATE INDEX est_snap_confidence_idx  ON public.estimate_snapshots (confidence, snapshot_at DESC);
CREATE INDEX est_snap_method_idx      ON public.estimate_snapshots (method, snapshot_at DESC);
CREATE INDEX est_snap_engine_idx      ON public.estimate_snapshots (engine_version);
-- Partial index: fast lookup of unresolved rows for the resolve job
CREATE INDEX est_snap_unresolved_idx  ON public.estimate_snapshots (snapshot_at)
  WHERE actual_price IS NULL;


-- ─────────────────────────────────────────────────────────────────────────
-- 3. ENGINE_CALIBRATION_LOG
--    Tracks every parameter change made to engine_params, why it was made,
--    and what the measured MAE impact was before and after.
--
--    Mirrors calibration_log (which tracked old model_weights changes)
--    but uses engine_params.version as the FK reference.
-- ─────────────────────────────────────────────────────────────────────────

CREATE TABLE public.engine_calibration_log (
  id                 uuid NOT NULL DEFAULT gen_random_uuid(),
  triggered_at       timestamptz DEFAULT now(),

  -- Version transition
  old_version        integer NOT NULL,   -- FK → engine_params.version
  new_version        integer NOT NULL,   -- FK → engine_params.version

  -- What changed
  parameter_changed  text NOT NULL,      -- e.g. 'staleness_tau', 'quality_medium_low'
  old_value          numeric,
  new_value          numeric,

  -- Measured accuracy impact (computed from estimate_snapshots at decision time)
  mae_before         numeric,
  mae_after          numeric,
  improvement_pct    numeric,            -- (mae_before − mae_after) / mae_before × 100
  observations_used  integer,            -- resolved snapshot count used to decide

  -- Context
  reasoning          text,               -- human or AI-generated explanation
  was_applied        boolean DEFAULT true,
  -- 'manual' = human decision; 'auto' = threshold-triggered; 'ai' = AI-suggested
  triggered_by       text DEFAULT 'manual',

  CONSTRAINT engine_calibration_log_pkey PRIMARY KEY (id),
  CONSTRAINT eng_cal_log_old_fkey
    FOREIGN KEY (old_version) REFERENCES public.engine_params(version),
  CONSTRAINT eng_cal_log_new_fkey
    FOREIGN KEY (new_version) REFERENCES public.engine_params(version)
);


-- ─────────────────────────────────────────────────────────────────────────
-- 4. SEED: Engine v1 — the current hardcoded values in estimate-engine.js
--    This makes the baseline explicit and queryable from day one.
--    model.html can immediately read from this row.
-- ─────────────────────────────────────────────────────────────────────────

INSERT INTO public.engine_params (
  version, tag, is_active, notes,
  staleness_tau, evidence_k,
  weight_30d, weight_90d, weight_alltime,
  recency_conf_30d, recency_conf_90d, recency_conf_alltime,
  rarity_premium, pop_proximity_floor,
  league_same_weight, league_diff_weight,
  evidence_threshold,
  quality_medium, quality_medium_low, quality_low, quality_very_low,
  extrap_cap_low, extrap_cap_high,
  -- Listings phase: all NULL (off)
  listing_blend, listing_floor_discount, supply_pressure
) VALUES (
  1,
  'v2.0-baseline',
  true,
  'Baseline — mirrors constants hardcoded in estimate-engine.js at launch. '
  'No Listings signal yet (listing_blend NULL = off).',
  -- Freshness
  90, 0.0231,
  -- salesStrength weights
  3.0, 1.5, 1.0,
  -- recencyWeight confidence factors
  1.0, 0.7, 0.4,
  -- Structure
  1.20, 0.60,
  -- League weights
  0.90, 0.70,
  -- Evidence threshold
  0.15,
  -- Confidence quality gates
  0.85, 0.65, 0.40, 0.20,
  -- Extrapolation caps
  0.10, 10.0,
  -- Listings (off)
  NULL, NULL, NULL
);
