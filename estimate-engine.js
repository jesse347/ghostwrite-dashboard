// ════════════════════════════════════════════════════════════════════════════
// estimate-engine.js — Shared estimation engine for ghostladder dashboard
// Produces model-estimated prices for player/variant combos with no real sales.
// Consumed by: heatmap.html, case-sim.html, sets.html
//
// Global helpers: recencyWeight, bestSalesPrice, salesStrength, weightedMedianRatio
// New helpers:    buildPlayerValueScale, interpolateFromDistribution,
//                 buildReferenceDistribution, buildPvsInfo
// Main function:  buildEstimates(setName, setProducts, weights)
//   → { estimates, byPlayer, evidenceWeight, setRatios, pvs }
//
// Methodology (v2 — Player Value Scale):
//  1. Build Player Value Scale (PVS): rank every player by implied Base price.
//     Source priority: direct Base sales → inferred from variant comparisons
//     → set median for players with no data at all.
//  2. Pre-compute reference set data: for each other set, build byPlayer + PVS.
//  3. Cross-variant ratios from THIS set's real sales (ratio pairs).
//  4. Fill ratio gaps with cross-set ratios. No population-math fallback.
//  5. Enforce monotonicity (rarer variant always costs more).
//  6. Build direct player+variant comps from other sets (strongest signal).
//  7. Evidence weight = 1 − e^(−k × nonBaseWeightedSales), k = 0.0231
//     → 30 recency-weighted non-base sales gives 50% weight.
//  8. Estimate missing cells (priority order):
//     a. Direct comp (same player + same variant in another set)
//     b. Reference distribution: map player's PVS percentile to variant price
//        using (percentile → price) pairs from the nearest comparable set(s).
//        Same-league sets used exclusively when available.
//     c. Blend in ratio estimate when evidenceWeight is significant (> 15%).
//     d. Ratio-only if no reference distribution exists yet.
//  9. (removed — implied_price fallback was dependent on price_signals)
// ════════════════════════════════════════════════════════════════════════════

// ── Staleness / freshness ─────────────────────────────────────────────────────
//
// Freshness decays exponentially from 1.0 (just sold) toward 0 (very stale).
// τ = 90 days → at 90 days, own-data weight = e^(-1) ≈ 37%.
// Used by the Full Model mode to blend real sales prices with model estimates.

const STALENESS_TAU = 90; // days

function computeFreshness(lastSaleDate) {
  if (!lastSaleDate) return 0; // no date = treat as maximally stale
  const days = (Date.now() - new Date(lastSaleDate).getTime()) / 86400000;
  return Math.exp(-days / STALENESS_TAU);
}

// ── Comp price: recency-weighted blend across all sales windows ───────────────
//
// Unlike bestSalesPrice (which returns only the most recent window), this
// blends all observed periods so a single recent low sale doesn't dominate.
// Weights: 0–30d × 3,  30–90d × 1.5,  90d+ × 1.
// Each window's marginal average is derived from the non-overlapping slice.
//
function blendedCompPrice(p) {
  if (!p) return 0;
  const s30  = p.sales30d  || 0;
  const s90  = p.sales90d  || 0;
  const sAll = p.totalSales || 0;
  let wtSum = 0, wtTotal = 0;

  if (s30 > 0 && p.avg30d > 0) {
    wtSum   += p.avg30d * s30 * 3;
    wtTotal += s30 * 3;
  }
  const s30_90 = Math.max(0, s90 - s30);
  if (s30_90 > 0 && p.avg90d > 0) {
    const midAvg = (p.avg90d * s90 - (p.avg30d || 0) * s30) / s30_90;
    if (midAvg > 0) { wtSum += midAvg * s30_90 * 1.5; wtTotal += s30_90 * 1.5; }
  }
  const sOld = Math.max(0, sAll - s90);
  if (sOld > 0 && p.avgPrice > 0) {
    const oldAvg = (p.avgPrice * sAll - (p.avg90d || 0) * s90) / sOld;
    if (oldAvg > 0) { wtSum += oldAvg * sOld; wtTotal += sOld; }
  }
  return wtTotal > 0 ? wtSum / wtTotal : (p.avgPrice || p.avg90d || p.avg30d || 0);
}

// ── Proximity weight between two populations ──────────────────────────────────
//
// Closer rarities are better ratio anchors.  Distance is measured in log space
// so e.g. Chrome(/5) ↔ Gold(/10) (log-dist 0.69) → 0.81 weight and
//          Chrome(/5) ↔ Victory(/50) (log-dist 2.30) → 0.50 weight.
//
function computeProximityWeight(targetPop, sourcePop) {
  if (!targetPop || !sourcePop) return 0.5;
  if (targetPop === sourcePop) return 1.0;
  return Math.exp(-0.3 * Math.abs(Math.log(targetPop / sourcePop)));
}

// ── AFA Grading Premium helpers ───────────────────────────────────────────────
//
// AFA-graded cards sell for a significant premium over ungraded (e.g. 2×).
// price_intelligence averages (avg_price, avg_30d, avg_90d) can be inflated
// when graded sales are mixed in. These helpers strip that contamination so
// the engine uses accurate ungraded-equivalent prices as comps.
//
// Usage:
//   const gradedNormMap = buildGradedNormMap(rawGradedSales);
//   applyGradingNorm(allData.products, gradedNormMap);
//
// Both functions are safe to call in GAS (apps-script-functions.js) and in
// every browser page boot sequence (heatmap, sets, case-sim, guide).
//
// The grading multiplier stored in engine_params.afa_grade_multiplier is used
// by computeGradingMultiplier() (GAS) and resolveSnapshots() for MAE adjustment.

/**
 * Build a per-combo normalization map from raw graded sale records.
 * Key format: "player|||variant"  (set-agnostic — premium is universal across sets)
 * Value:      { sum, count, sum30d, count30d, sum90d, count90d }
 *
 * @param {Array} gradedSales  Raw rows with { player, variant, price, date }
 *                             Works with Supabase REST rows (GAS) and browser fetch.
 * @returns {Object} gradedNormMap
 */
function buildGradedNormMap(gradedSales) {
  if (!Array.isArray(gradedSales) || !gradedSales.length) return {};
  const now = Date.now();
  const d30 = new Date(now - 30 * 86400000).toISOString().slice(0, 10);
  const d90 = new Date(now - 90 * 86400000).toISOString().slice(0, 10);
  const map = {};
  gradedSales.forEach(s => {
    if (!s.player || !s.variant) return;
    const k = `${s.player}|||${s.variant}`;
    if (!map[k]) map[k] = { sum: 0, count: 0, sum30d: 0, count30d: 0, sum90d: 0, count90d: 0 };
    const g     = map[k];
    const price = +(s.price || 0);
    if (price <= 0) return;
    const dateStr = (s.date || '').slice(0, 10);
    g.sum   += price;
    g.count += 1;
    if (dateStr >= d30) { g.sum30d += price; g.count30d += 1; }
    if (dateStr >= d90) { g.sum90d += price; g.count90d += 1; }
  });
  return map;
}

// AFA grading multiplier — matches engine_params.afa_grade_multiplier (set to 1.5 via SQL).
// A graded sale at $1 000 is worth $1 000 / 1.5 = ~$667 as an ungraded-equivalent comp.
const AFA_MULT = 1.5;

/**
 * Normalize product price averages so that graded-sale inflation is removed.
 * Graded sales are NOT discarded — they are kept at their ungraded-equivalent
 * price (actual ÷ AFA_MULT) so they still contribute data to the engine.
 *
 * Modifies products IN-PLACE:
 *   • All-graded combo  → avgPrice / avg30d / avg90d set to normalised graded avg
 *   • Mixed combo       → graded actuals replaced by normalised equivalents in avg
 *   • Sale counts       → always preserved (never reduced), since graded sales count
 *
 * @param {Array}  products       Product objects (modified in-place)
 * @param {Object} gradedNormMap  Output of buildGradedNormMap()
 */
function applyGradingNorm(products, gradedNormMap) {
  if (!gradedNormMap || !products) return;
  products.forEach(prod => {
    const k = `${prod.player}|||${prod.variant}`;
    const g = gradedNormMap[k];
    if (!g || g.count === 0) return;

    // Stash unadjusted (raw) averages so display code can show actual sale prices
    prod._rawAvgPrice = prod.avgPrice;
    prod._rawAvg30d   = prod.avg30d;
    prod._rawAvg90d   = prod.avg90d;

    // ── All-time average ──────────────────────────────────────────────────
    if (prod.totalSales > 0) {
      if (g.count >= prod.totalSales) {
        // All sales graded → use normalised graded average (actual ÷ AFA_MULT)
        prod.avgPrice = Math.round(g.sum / g.count / AFA_MULT * 100) / 100;
        // totalSales unchanged — graded sales still count
      } else if (prod.avgPrice > 0) {
        // Mixed → replace graded actuals with normalised equivalents in the sum
        const blended = prod.avgPrice * prod.totalSales - g.sum + g.sum / AFA_MULT;
        prod.avgPrice = Math.round(Math.max(0, blended) / prod.totalSales * 100) / 100;
      }
    }

    // ── 30-day average ────────────────────────────────────────────────────
    if (g.count30d > 0 && prod.sales30d > 0) {
      if (g.count30d >= prod.sales30d) {
        prod.avg30d = Math.round(g.sum30d / g.count30d / AFA_MULT * 100) / 100;
      } else if (prod.avg30d > 0) {
        const blended = prod.avg30d * prod.sales30d - g.sum30d + g.sum30d / AFA_MULT;
        prod.avg30d = Math.round(Math.max(0, blended) / prod.sales30d * 100) / 100;
      }
    }

    // ── 90-day average ────────────────────────────────────────────────────
    if (g.count90d > 0 && prod.sales90d > 0) {
      if (g.count90d >= prod.sales90d) {
        prod.avg90d = Math.round(g.sum90d / g.count90d / AFA_MULT * 100) / 100;
      } else if (prod.avg90d > 0) {
        const blended = prod.avg90d * prod.sales90d - g.sum90d + g.sum90d / AFA_MULT;
        prod.avg90d = Math.round(Math.max(0, blended) / prod.sales90d * 100) / 100;
      }
    }
  });
}

// ── Global helpers ────────────────────────────────────────────────────────────

function recencyWeight(product) {
  if (!product) return 0;
  if (product.avg30d  && product.sales30d  >= 1) return 1.0;
  if (product.avg90d  && product.sales90d  >= 1) return 0.7;
  if (product.avgPrice && product.totalSales >= 1) return 0.4;
  return 0;
}

function bestSalesPrice(product) {
  if (!product) return 0;
  if (product.avg30d  && product.sales30d  >= 1) return product.avg30d;
  if (product.avg90d  && product.sales90d  >= 1) return product.avg90d;
  if (product.avgPrice && product.totalSales >= 1) return product.avgPrice;
  return 0;
}

// Recency-weighted sales count — more recent = more weight
function salesStrength(product) {
  if (!product) return 0;
  const s30  = (product.sales30d  || 0) * 3.0;
  const s90  = Math.max(0, (product.sales90d  || 0) - (product.sales30d  || 0)) * 1.5;
  const sAll = Math.max(0, (product.totalSales || 0) - (product.sales90d  || 0)) * 1.0;
  return s30 + s90 + sAll;
}

// Weighted median ratio across players
function weightedMedianRatio(pairs) {
  if (!pairs.length) return null;
  const sorted = [...pairs].sort((a, b) => a.ratio - b.ratio);
  const totalW = sorted.reduce((s, p) => s + (p.weight || 1), 0);
  let cumW = 0;
  for (const p of sorted) {
    cumW += (p.weight || 1);
    if (cumW >= totalW / 2) return { ratio: p.ratio, n: pairs.length, totalWeight: totalW };
  }
  return { ratio: sorted[Math.floor(sorted.length / 2)].ratio, n: pairs.length, totalWeight: totalW };
}

// ── Player Value Scale (PVS) ──────────────────────────────────────────────────
//
// Ranks all players in a set by their implied Base card price.
//
// Pass 1 — players with direct Base sales (most reliable)
// Pass 2 — players without Base: infer from variant price comparisons
//          against players who DO have Base data
// Pass 3 — players with no data at all: assign set median
//
// Returns:
//   playerImpliedBase[player] = { price, source, inferredFrom }
//   playerPercentile[player]  = 0.0–1.0  (0 = least valuable, 1 = most)
//   setMedianBase             = number
//   pvsSorted                 = [{player, price, source}] sorted ascending

function buildPlayerValueScale(byPlayer) {
  const playerImpliedBase = {};

  // ── Pass 1: direct Base sales ───────────────────────────────────────────
  Object.entries(byPlayer).forEach(([player, pvars]) => {
    const bp = bestSalesPrice(pvars['Base']);
    if (bp > 0) {
      playerImpliedBase[player] = { price: bp, source: 'base', inferredFrom: null };
    }
  });

  // ── Pass 2: infer Base from variant comparisons ─────────────────────────
  // For each player without Base, look at their variant prices and compare
  // to players who DO have direct Base data via the same variant.
  // Formula: impliedBase = (thisVariantPrice / otherVariantPrice) × otherBasePrice
  Object.entries(byPlayer).forEach(([player, pvars]) => {
    if (playerImpliedBase[player]) return; // already placed

    const inferences = [];
    Object.entries(pvars).forEach(([variant, prod]) => {
      if (variant === 'Base') return;
      const thisPrice = bestSalesPrice(prod);
      if (thisPrice <= 0) return;
      const thisStr = salesStrength(prod);

      Object.entries(byPlayer).forEach(([otherPlayer, otherPvars]) => {
        const otherBase = playerImpliedBase[otherPlayer];
        if (!otherBase || otherBase.source !== 'base') return; // only compare to confirmed base players
        const otherVariantPrice = bestSalesPrice(otherPvars[variant]);
        if (otherVariantPrice <= 0) return;

        const implied = (thisPrice / otherVariantPrice) * otherBase.price;
        const otherStr = salesStrength(otherPvars[variant]);
        const weight   = Math.sqrt(Math.max(0.1, thisStr) * Math.max(0.1, otherStr));
        inferences.push({ price: implied, weight, inferredFrom: { variant, otherPlayer } });
      });
    });

    if (inferences.length > 0) {
      // Weighted median of all inferences
      const sorted = [...inferences].sort((a, b) => a.price - b.price);
      const totalW = sorted.reduce((s, i) => s + i.weight, 0);
      let cumW = 0, chosen = sorted[0];
      for (const inf of sorted) {
        cumW += inf.weight;
        if (cumW >= totalW / 2) { chosen = inf; break; }
      }
      playerImpliedBase[player] = { price: chosen.price, source: 'inferred', inferredFrom: chosen.inferredFrom };
    }
  });

  // ── Compute set median from known players ────────────────────────────────
  const knownPrices = Object.values(playerImpliedBase).map(b => b.price).sort((a, b) => a - b);
  const setMedianBase = knownPrices.length ? knownPrices[Math.floor(knownPrices.length / 2)] : 1;

  // ── Pass 3: assign set median to players with no data ───────────────────
  Object.keys(byPlayer).forEach(player => {
    if (!playerImpliedBase[player]) {
      playerImpliedBase[player] = { price: setMedianBase, source: 'median', inferredFrom: null };
    }
  });

  // ── Compute percentiles ──────────────────────────────────────────────────
  const pvsSorted = Object.entries(playerImpliedBase)
    .map(([player, data]) => ({ player, price: data.price, source: data.source }))
    .sort((a, b) => a.price - b.price);

  const playerPercentile = {};
  pvsSorted.forEach((entry, i) => {
    playerPercentile[entry.player] = pvsSorted.length > 1 ? i / (pvsSorted.length - 1) : 0.5;
  });

  return { playerImpliedBase, playerPercentile, setMedianBase, pvsSorted };
}

// ── Package PVS info for the drawer ──────────────────────────────────────────

function buildPvsInfo(player, playerImpliedBase, playerPercentile, pvsSorted) {
  const data = playerImpliedBase[player];
  if (!data) return null;
  const pct  = playerPercentile[player] ?? 0.5;
  const rank = pvsSorted.findIndex(e => e.player === player) + 1; // 1 = least valuable
  return {
    impliedBase:   Math.round(data.price),
    source:        data.source,
    inferredFrom:  data.inferredFrom || null,
    percentile:    Math.round(pct * 100) / 100,
    percentilePct: Math.round(pct * 100),
    rank,
    nPlayers:      pvsSorted.length,
  };
}

// ── Reference Distribution ────────────────────────────────────────────────────
//
// Builds a sorted array of (percentile, price) pairs for a given variant
// using player data from other sets.
//
// Matching strategy: population-based (not name-based).
//   For a target variant with population=25, we look for any variant in the
//   reference set that also has population=25 — regardless of its name.
//   If no exact population match exists, we use the nearest-population variant
//   and scale prices by (refPop / targetPop):
//     ref rarer  (smaller pop) → ref prices are HIGHER → scale DOWN  (factor < 1)
//     ref common (larger  pop) → ref prices are LOWER  → scale UP    (factor > 1)
//   Matches with population proximity < 0.4 are discarded as too distant.
//
// League priority:
//   - Same-league sets used exclusively when available
//   - Cross-league blend when no same-league match exists
//
// Returns: { points, sameLeague, refSets, nPoints, rangeMin, rangeMax,
//            targetPop, exactPopMatch, scaleFactor } | null

function buildReferenceDistribution(variant, targetPop, currentSetName, refSetData) {
  const thisSetLeague  = (allData.setLeagueMap || {})[currentSetName] || '';
  const allPoints      = [];
  const sameLeagueSets = new Set();
  const diffLeagueSets = new Set();
  let   anyExact       = false;
  let   anyApprox      = false;

  Object.entries(refSetData).forEach(([refSet, { byPlayer: refByPlayer, pvs: refPVS }]) => {
    const otherLeague   = (allData.setLeagueMap || {})[refSet] || '';
    const isLeagueMatch = !!(thisSetLeague && otherLeague && thisSetLeague === otherLeague);

    // ── Find best-matching variant in this reference set by population ──────
    const refPops = {}; // variant → population for this reference set
    ((allData.setMath || {})[refSet] || []).forEach(r => {
      if (r.variant && r.population != null) refPops[r.variant] = +r.population;
    });

    // Helper: does this variant have at least one player with real sales?
    const hasSales = v => Object.values(refByPlayer).some(p => bestSalesPrice(p[v]) > 0);

    let matchV = null, matchPop = null, scaleFactor = 1.0, isExact = false;

    if (targetPop != null) {
      // ── Exact population match (prefer same variant name on tie) ──────────
      const exactCandidates = Object.entries(refPops)
        .filter(([v, pop]) => pop === targetPop && hasSales(v));
      if (exactCandidates.length > 0) {
        const nameMatch = exactCandidates.find(([v]) => v === variant);
        [matchV, matchPop] = nameMatch
          ? [nameMatch[0], targetPop]
          : [exactCandidates[0][0], targetPop];
        scaleFactor = 1.0;
        isExact     = true;
      } else {
        // ── Nearest population match ──────────────────────────────────────────
        let bestDiff = Infinity;
        Object.entries(refPops).forEach(([v, pop]) => {
          if (!hasSales(v)) return;
          const diff = Math.abs(pop - targetPop);
          if (diff < bestDiff) { bestDiff = diff; matchV = v; matchPop = pop; }
        });
        if (matchV) {
          scaleFactor = matchPop / targetPop; // adjust price to targetPop level
          isExact     = false;
        }
      }
    } else {
      // No population data for target variant → fall back to name match
      if (hasSales(variant)) { matchV = variant; matchPop = null; scaleFactor = 1.0; isExact = true; }
    }

    if (!matchV) return;

    // Population proximity: how close is the match?  (1.0 = exact, < 1 = approximate)
    // Distant matches (proximity < 0.4, e.g. pop=10 vs pop=100) are discarded.
    const popProximity = (targetPop && matchPop)
      ? Math.min(targetPop, matchPop) / Math.max(targetPop, matchPop)
      : 1.0;
    if (popProximity < 0.4) return;

    // ── Collect (price, metadata) pairs — percentile assigned after all points collected ──
    let hasVariantData = false;
    Object.entries(refByPlayer).forEach(([player, pvars]) => {
      const rawPrice = bestSalesPrice(pvars[matchV]);
      if (rawPrice <= 0) return;

      hasVariantData = true;
      const str = Math.min(3.0, salesStrength(pvars[matchV]) / 5);
      allPoints.push({
        // percentile assigned below via price-ranking (not PVS-based)
        price:      rawPrice * scaleFactor, // price adjusted for population difference
        player,
        set:        refSet,
        sameLeague: isLeagueMatch,
        weight:     (isLeagueMatch ? 1.0 : 0.7) * Math.max(0.2, str) * popProximity,
        scaleFactor,
        isExact,
        matchPop,
      });
    });

    if (hasVariantData) {
      if (isExact) anyExact = true; else anyApprox = true;
      if (isLeagueMatch) sameLeagueSets.add(refSet);
      else diffLeagueSets.add(refSet);
    }
  });

  if (!allPoints.length) return null;

  const hasSameLeague = sameLeagueSets.size > 0;
  const points = hasSameLeague ? allPoints.filter(p => p.sameLeague) : allPoints;
  if (!points.length) return null;

  // Assign percentiles by price rank (cheapest = 0, most expensive = 1).
  // This keeps the distribution monotonically increasing and avoids distortion
  // from non-athlete entities (e.g. Trophy, Eminem) that have high PVS scores
  // but lower variant-specific prices than star athletes.
  points.sort((a, b) => a.price - b.price);
  const nPts = points.length;
  points.forEach((p, i) => { p.percentile = nPts === 1 ? 0.5 : i / (nPts - 1); });

  const avgScaleFactor = Math.round(
    (points.reduce((s, p) => s + p.scaleFactor, 0) / points.length) * 100
  ) / 100;

  return {
    points,
    sameLeague:    hasSameLeague,
    refSets:       hasSameLeague ? [...sameLeagueSets] : [...diffLeagueSets],
    nPoints:       points.length,
    rangeMin:      points[0],
    rangeMax:      points[points.length - 1],
    targetPop,
    exactPopMatch: !anyApprox,    // true only if every reference set had an exact pop match
    scaleFactor:   avgScaleFactor,
  };
}

// ── Log-linear interpolation / extrapolation ──────────────────────────────────
//
// Given a target percentile and sorted (percentile, price) reference points,
// estimates price using log-linear interpolation within the range and
// log-linear extrapolation (using the nearest segment's slope) outside it.
//
// Prices are log-normally distributed so interpolating in log space is more
// accurate than linear interpolation.
//
// Returns: { price, extrapolated, extrapolatedUp } | null

function interpolateFromDistribution(targetPercentile, points) {
  if (!points || !points.length) return null;
  if (points.length === 1) {
    return {
      price:         points[0].price,
      extrapolated:  targetPercentile !== points[0].percentile,
      extrapolatedUp: targetPercentile > points[0].percentile,
    };
  }

  // Find the boundary indices around the target percentile
  let loIdx = -1, hiIdx = -1;
  for (let i = 0; i < points.length; i++) {
    if (points[i].percentile <= targetPercentile) loIdx = i;
    if (hiIdx < 0 && points[i].percentile >= targetPercentile) hiIdx = i;
  }

  // ── Extrapolate below range ──────────────────────────────────────────────
  if (loIdx < 0) {
    const p0 = points[0], p1 = points[1];
    if (p0.percentile === p1.percentile) return { price: p0.price, extrapolated: true, extrapolatedUp: false };
    const slope    = (Math.log(p1.price) - Math.log(p0.price)) / (p1.percentile - p0.percentile);
    const logPrice = Math.log(p0.price) + slope * (targetPercentile - p0.percentile);
    // Floor at 10% of the range minimum to prevent unreasonably small values
    return { price: Math.max(Math.exp(logPrice), p0.price * 0.1), extrapolated: true, extrapolatedUp: false };
  }

  // ── Extrapolate above range ──────────────────────────────────────────────
  if (hiIdx < 0) {
    const rangeTop = points[points.length - 1]; // highest-percentile point (end of range)

    // Anchor from the highest-PRICED point, not the highest-percentile point.
    // A celebrity anchor (e.g. Eminem) can sit at the top of the PVS ranking
    // with a lower price than star athletes below it, creating a misleading
    // negative slope. Using the price-max point as the anchor ensures we
    // extrapolate upward from the most valuable known reference price.
    const byPrice   = [...points].sort((a, b) => b.price - a.price);
    const priceHigh = byPrice[0]; // highest-priced reference point
    const priceLow  = byPrice[1]; // second-highest, for slope

    // Slope between the top two price points (always positive by construction).
    const pctGap = Math.abs(priceHigh.percentile - priceLow.percentile) || 0.1;
    const slope  = Math.max(0, (Math.log(priceHigh.price) - Math.log(priceLow.price)) / pctGap);

    // Extrapolate from priceHigh over the distance from rangeTop to target.
    const dist     = targetPercentile - rangeTop.percentile;
    const logPrice = Math.log(priceHigh.price) + slope * dist;
    // Cap at 10× the range price-max to prevent runaway extrapolation
    return { price: Math.min(Math.exp(logPrice), priceHigh.price * 10), extrapolated: true, extrapolatedUp: true };
  }

  // ── Interpolate ──────────────────────────────────────────────────────────
  const lo = points[loIdx], hi = points[hiIdx];
  if (lo.percentile === hi.percentile) {
    return { price: (lo.price + hi.price) / 2, extrapolated: false, extrapolatedUp: false };
  }
  const t        = (targetPercentile - lo.percentile) / (hi.percentile - lo.percentile);
  const logPrice = Math.log(lo.price) * (1 - t) + Math.log(hi.price) * t;
  return { price: Math.exp(logPrice), extrapolated: false, extrapolatedUp: false };
}

// ── Global population-tier ratio table ───────────────────────────────────────
//
// Computes the universal variant→variant price ratio for every population pair
// using ALL data across ALL sets.  Keyed by "popA|||popB" (e.g. "5|||10").
//
// Why population-keyed, not name-keyed?
//   Fire is always 1/1, Chrome /5, Gold /10 regardless of set.  "Victory" and
//   "Sapphire" are just brand names for the same /50 rarity tier.  Collectors
//   price by scarcity, so the ratio between any two rarity tiers should be the
//   same across all sets.  Using population as the key unifies all those names.
//
// Base variants are excluded: Base has a variable population by set (WNBA /800,
// NBA /1200, etc.) and is a commodity card whose ratio to rare variants isn't
// meaningful for pricing purposes.
//
// Stored lazily on allData._globalRatioTable and reused for every set estimated
// in the same page session.
//
function buildGlobalRatioTable() {
  // Build set×variant → population lookup
  const variantSetPop = {};
  Object.entries(allData.setMath || {}).forEach(([sn, rows]) => {
    rows.forEach(r => {
      if (r.variant && r.population != null)
        variantSetPop[`${sn}|||${r.variant}`] = +r.population;
    });
  });

  // player → population → best { price, strength }
  // "Best" = highest salesStrength (most data-rich) across all sets
  const playerPopPrices = {};
  (allData.products || []).forEach(p => {
    if (!p.player || p.player === 'Unknown' || p.variant === 'Base') return;
    const pop = variantSetPop[`${p.set}|||${p.variant}`];
    if (!pop) return;
    const price = bestSalesPrice(p);
    if (price <= 0) return;
    const str = salesStrength(p);
    if (!playerPopPrices[p.player]) playerPopPrices[p.player] = {};
    const cur = playerPopPrices[p.player][pop];
    if (!cur || str > cur.strength) playerPopPrices[p.player][pop] = { price, strength: str };
  });

  const popSet = new Set();
  Object.values(playerPopPrices).forEach(pm => Object.keys(pm).forEach(p => popSet.add(+p)));
  const pops = [...popSet];

  const table = {};
  pops.forEach(popA => {
    pops.forEach(popB => {
      if (popA >= popB) return; // compute one direction; derive inverse below
      const pairs = [];
      Object.values(playerPopPrices).forEach(pm => {
        const a = pm[popA], b = pm[popB];
        if (!a || !b) return;
        const wt = Math.sqrt(Math.max(0.1, a.strength) * Math.max(0.1, b.strength));
        pairs.push({ ratio: a.price / b.price, weight: wt });
      });
      if (pairs.length < 2) return;
      const med = weightedMedianRatio(pairs);
      if (!med) return;
      table[`${popA}|||${popB}`] = { ratio: med.ratio, n: med.n, source: 'global-pop' };
      table[`${popB}|||${popA}`] = { ratio: 1 / med.ratio, n: med.n, source: 'global-pop' };
    });
  });
  return table;
}

// ── Main estimation engine ────────────────────────────────────────────────────

function buildEstimates(setName, setProducts, weights) {
  // Local alias — avoids conflict with sets.html's recency-aware page-level bestPrice
  const bestPrice = bestSalesPrice;

  // Tunable params
  const w          = weights || {};
  const RARITY_PREMIUM = +w.rarity_premium || 1.20;
  const EVIDENCE_K     = +w.evidence_k     || 0.0231; // 30 recency-weighted sales → 50%

  // ── 0. Index products by player ─────────────────────────────────────────
  const byPlayer = {};
  setProducts.forEach(p => {
    if (!byPlayer[p.player]) byPlayer[p.player] = {};
    byPlayer[p.player][p.variant] = p;
  });

  const mathRows    = (allData.setMath || {})[setName] || [];
  const allVariants = mathRows.length
    ? [...new Set(mathRows.map(r => r.variant).filter(Boolean))]
    : [...new Set(setProducts.map(p => p.variant))];

  // Include all mathRows players (even those with no products yet)
  mathRows.forEach(({ player }) => {
    if (player && player !== 'Unknown' && !byPlayer[player]) byPlayer[player] = {};
  });

  // Per-player valid variants — set_math is the source of truth for which
  // player×variant combos actually exist. Only estimate these combos; never
  // extrapolate across a variant the player doesn't have in the catalog.
  const validVariantsByPlayer = {};
  mathRows.forEach(({ player, variant }) => {
    if (!player || player === 'Unknown' || !variant) return;
    if (!validVariantsByPlayer[player]) validVariantsByPlayer[player] = new Set();
    validVariantsByPlayer[player].add(variant);
  });

  // Variant populations — used for monotonicity enforcement only
  const variantPop = {};
  mathRows.forEach(r => {
    if (r.variant && r.population != null && !variantPop[r.variant])
      variantPop[r.variant] = +r.population;
  });

  // ── 1. Build Player Value Scale for current set ──────────────────────────
  const pvs = buildPlayerValueScale(byPlayer);
  const { playerImpliedBase, playerPercentile, setMedianBase, pvsSorted } = pvs;

  // ── 2. Pre-compute reference set data ───────────────────────────────────
  // Build byPlayer and PVS for every other set (reused for every variant).
  const refSetData = {};
  (allData.products || []).forEach(p => {
    if (p.set === setName || !p.player || p.player === 'Unknown') return;
    if (!refSetData[p.set]) refSetData[p.set] = { products: [] };
    refSetData[p.set].products.push(p);
  });
  Object.entries(refSetData).forEach(([, data]) => {
    const refByPlayer = {};
    data.products.forEach(p => {
      if (!refByPlayer[p.player]) refByPlayer[p.player] = {};
      refByPlayer[p.player][p.variant] = p;
    });
    data.byPlayer = refByPlayer;
    data.pvs      = buildPlayerValueScale(refByPlayer);
  });

  // Include the current set itself as a reference source.
  // Players in this set who already have real sales for a variant are the most
  // directly relevant data points when estimating other players in the same set.
  // bestSalesPrice() returns 0 for players with no real data, so no circularity.
  refSetData[setName] = { byPlayer, pvs };

  // Cache reference distribution per variant (expensive to rebuild per player)
  const refDistCache = {};

  // ── 2.5. Lazy-initialise global population-tier ratio table ─────────────
  if (!allData._globalRatioTable) allData._globalRatioTable = buildGlobalRatioTable();
  const globalRatios = allData._globalRatioTable;

  // ── 3. Cross-variant ratios from THIS set's real sales ───────────────────
  // Base excluded: Base has variable population across sets and isn't a useful
  // anchor for pricing rare variants (collectors think in rarity tiers, not Base).
  const setRatios = {};
  allVariants.forEach(vA => {
    allVariants.forEach(vB => {
      if (vA === vB || vA === 'Base' || vB === 'Base') return;
      const key   = `${vA}|||${vB}`;
      const pairs = [];
      Object.values(byPlayer).forEach(pvars => {
        const pA = bestSalesPrice(pvars[vA]);
        const pB = bestSalesPrice(pvars[vB]);
        if (pA > 0 && pB > 0) {
          const wA     = salesStrength(pvars[vA]);
          const wB     = salesStrength(pvars[vB]);
          const weight = Math.sqrt(Math.max(0.1, wA) * Math.max(0.1, wB));
          pairs.push({ ratio: pA / pB, weight });
        }
      });
      const med = weightedMedianRatio(pairs);
      if (med) setRatios[key] = { ...med, source: 'sales' };
    });
  });

  // ── 4. Fill ratio gaps ────────────────────────────────────────────────────
  // Priority: global population-tier table → per-player cross-set observations.
  // Base excluded from both paths (variable population; not a valid rarity anchor).
  //
  // Global table: pre-computed weighted-median ratios between every population
  // pair across all sets.  This is the primary source — it pools all available
  // data and is rarity-name-agnostic ("Victory" and "Sapphire" are both /50).
  //
  // Per-player fallback: kept for unusual populations not yet in the global table.

  const allOtherProducts = (allData.products || []).filter(p => p.set !== setName);

  const variantPopMap = {};
  Object.entries(allData.setMath || {}).forEach(([sn, rows]) => {
    rows.forEach(r => {
      const k = `${sn}|||${r.variant}`;
      if (!variantPopMap[k] && r.population) variantPopMap[k] = +r.population;
    });
  });

  const otherByPlayer = {};
  allOtherProducts.forEach(p => {
    const k = `${p.set}|||${p.player}`;
    if (!otherByPlayer[k]) otherByPlayer[k] = {};
    otherByPlayer[k][p.variant] = p;
  });

  const POP_PROXIMITY_FLOOR = 0.60;
  const NAME_MATCH_BONUS    = 1.00;
  const POP_ONLY_BONUS      = 0.80;

  // N_GLOBAL: how strongly the global table acts as a prior when blending with
  // within-set observations.  At 10, a single within-set player contributes
  // only 1/(1+10) ≈ 9% of the final ratio — the global baseline dominates until
  // the set has accumulated meaningful within-set evidence (≥5 players → >33%).
  const N_GLOBAL = 10;

  allVariants.forEach(vA => {
    allVariants.forEach(vB => {
      if (vA === vB || vA === 'Base' || vB === 'Base') return;
      const key  = `${vA}|||${vB}`;
      const popA = variantPop[vA];
      const popB = variantPop[vB];

      // ── 4a. Global population-tier ratio — always blend in ──────────────
      // A within-set ratio from 1-2 players is noisy; the global table pools
      // all sets and should dominate until we have strong within-set evidence.
      if (popA && popB) {
        const gr = globalRatios[`${popA}|||${popB}`];
        if (gr && gr.n >= 2) {
          const within = setRatios[key]; // step-3 result, may be undefined
          if (!within) {
            // No within-set data — use global directly
            setRatios[key] = gr;
          } else {
            // Blend: within-set n players + N_GLOBAL-equivalent global prior
            const withinN = within.n || 1;
            setRatios[key] = {
              ratio:  (withinN * within.ratio + N_GLOBAL * gr.ratio) / (withinN + N_GLOBAL),
              n:      withinN,
              // Label reflects which source is dominant
              source: withinN >= 5 ? 'sales' : 'global-pop',
            };
          }
          return;
        }
      }

      // ── 4b. Per-player cross-set fallback ──────────────────────────────
      if (setRatios[key]) return; // within-set data already present — skip cross-set
      const pairs = [];
      Object.entries(otherByPlayer).forEach(([setPlayerKey, pvars]) => {
        const [otherSet] = setPlayerKey.split('|||');

        let priceA = bestPrice(pvars[vA]), weightA = NAME_MATCH_BONUS;
        if (!priceA && popA) {
          let bestProx = 0, bestV = null;
          Object.keys(pvars).forEach(v => {
            const vPop = variantPopMap[`${otherSet}|||${v}`];
            if (!vPop) return;
            const prox = Math.min(vPop, popA) / Math.max(vPop, popA);
            if (prox >= POP_PROXIMITY_FLOOR && prox > bestProx && bestPrice(pvars[v]) > 0)
              { bestProx = prox; bestV = v; }
          });
          if (bestV) { priceA = bestPrice(pvars[bestV]); weightA = (bestV === vA ? NAME_MATCH_BONUS : POP_ONLY_BONUS) * bestProx; }
        }
        if (!priceA) return;

        let priceB = bestPrice(pvars[vB]), weightB = NAME_MATCH_BONUS;
        if (!priceB && popB) {
          let bestProx = 0, bestV = null;
          Object.keys(pvars).forEach(v => {
            const vPop = variantPopMap[`${otherSet}|||${v}`];
            if (!vPop) return;
            const prox = Math.min(vPop, popB) / Math.max(vPop, popB);
            if (prox >= POP_PROXIMITY_FLOOR && prox > bestProx && bestPrice(pvars[v]) > 0)
              { bestProx = prox; bestV = v; }
          });
          if (bestV) { priceB = bestPrice(pvars[bestV]); weightB = (bestV === vB ? NAME_MATCH_BONUS : POP_ONLY_BONUS) * bestProx; }
        }
        if (!priceB) return;

        const ratio   = priceA / priceB;
        const matchWt = Math.sqrt(weightA * weightB);
        const salesW  = Math.sqrt(Math.max(0.1, salesStrength(pvars[vA] || pvars[vB])));
        pairs.push({ ratio, weight: matchWt * Math.min(1.0, salesW / 5) });
      });

      if (pairs.length) {
        const med = weightedMedianRatio(pairs);
        if (med) setRatios[key] = { ...med, source: 'cross-set' };
      }
    });
  });

  // ── 5. Monotonicity enforcement ──────────────────────────────────────────
  allVariants.forEach(vA => {
    allVariants.forEach(vB => {
      if (vA === vB || vA === 'Base' || vB === 'Base') return;
      const key  = `${vA}|||${vB}`;
      const popA = variantPop[vA], popB = variantPop[vB];
      if (!setRatios[key] || !popA || !popB || popA === popB) return;
      if (popA < popB && setRatios[key].ratio < 1)
        setRatios[key] = { ratio: (popB / popA) * RARITY_PREMIUM, n: 0, source: 'corrected' };
    });
  });

  // ── 6. Direct player+variant comps from other sets ───────────────────────
  const thisSetLeague = (allData.setLeagueMap || {})[setName] || '';
  const playerCompMap = {};
  const compByPlayerVariant = {};

  allOtherProducts.forEach(p => {
    if (!p.player || p.player === 'Unknown') return;
    // blendedCompPrice considers all sales windows (recency-weighted) instead of
    // just the most recent window, preventing a single low sale from dominating.
    // Graded sales already normalised to ungraded-equivalent by applyGradingNorm.
    const compPrice = blendedCompPrice(p);
    if (compPrice <= 0) return;
    const k = `${p.player}|||${p.variant}`;
    if (!compByPlayerVariant[k]) compByPlayerVariant[k] = [];
    const otherLeague = (allData.setLeagueMap || {})[p.set] || '';
    const leagueMatch = otherLeague && thisSetLeague && otherLeague === thisSetLeague ? 0.90 : 0.70;
    const recency     = p.avg30d && p.sales30d >= 1 ? 1.0 : p.avg90d && p.sales90d >= 1 ? 0.7 : 0.4;
    const strength    = Math.min(1.0, salesStrength(p) / 10);
    compByPlayerVariant[k].push({ price: compPrice, weight: leagueMatch * recency * Math.max(0.2, strength), set: p.set, sales: p.totalSales || 0 });
  });

  Object.entries(compByPlayerVariant).forEach(([k, comps]) => {
    const totalW = comps.reduce((s, c) => s + c.weight, 0);
    if (totalW <= 0) return;
    playerCompMap[k] = {
      price: Math.round(comps.reduce((s, c) => s + c.price * c.weight, 0) / totalW),
      n:     comps.length,
      comps,
    };
  });

  // ── 7. Evidence weight for this set ─────────────────────────────────────
  // Recency-weighted count of non-base variant sales in this set.
  // k = 0.0231 → 30 recency-weighted parallel sales gives ~50% weight.
  const nonBaseWeightedSales = Object.values(byPlayer).reduce((total, pvars) => {
    return total + Object.entries(pvars)
      .filter(([v]) => v !== 'Base')
      .reduce((s, [, p]) => s + salesStrength(p), 0);
  }, 0);
  const evidenceWeight = 1 - Math.exp(-EVIDENCE_K * nonBaseWeightedSales);

  // ── 8. Estimate cells ────────────────────────────────────────────────────
  // estimates{}         — gap-fill estimates (cells with NO real sales data)
  // refreshedEstimates{} — freshness-weighted blends for cells WITH real sales
  //   Full Model mode uses refreshedEstimates to surface a current-market price
  //   even when the card's own sales history exists but may be stale.
  const estimates          = {};
  const refreshedEstimates = {};

  Object.entries(byPlayer).forEach(([player, pvars]) => {
    // Restrict to variants this player actually has in set_math.
    // Fall back to allVariants only for players not in the catalog at all.
    const playerVariants = validVariantsByPlayer[player]
      ? [...validVariantsByPlayer[player]]
      : allVariants;

    playerVariants.forEach(tv => {
      // avgPrice is the normalized ungraded average — the only reliable signal.
      // lastSalePrice is intentionally excluded: it may be from a graded sale
      // and is preserved on the product for display purposes only.
      const hasRealData = pvars[tv]?.avgPrice > 0;

      // ── 8a. Build ratio anchors from player's own real prices ─────────────
      // Base excluded: collectors anchor rare-variant prices to other rare variants,
      // not to Base.  Proximity weighting: variants closer in rarity carry more
      // weight (Chrome /5 → Gold /10 is a better anchor than Chrome /5 → Victory /50).
      //
      // Cross-set fallback: if the player has no within-set sales for sv, check
      // playerCompMap[player|||sv] — their comp price for that variant from other
      // sets.  This lets Angel Reese's 2024 Gold sale anchor her 2025 Chrome
      // estimate via the universal Gold→Chrome ratio, instead of relying solely
      // on the reference distribution.  Cross-set anchors carry a 0.8× weight
      // penalty to reflect the additional uncertainty.
      const anchors = [];
      allVariants.forEach(sv => {
        if (sv === 'Base') return;
        let srcPrice    = bestSalesPrice(pvars[sv]);
        let srcCrossSet = false;
        let srcSetLabel = '';
        if (srcPrice <= 0) {
          const crossComp = playerCompMap[`${player}|||${sv}`];
          if (crossComp && crossComp.price > 0) {
            srcPrice    = crossComp.price;
            srcCrossSet = true;
            srcSetLabel = crossComp.comps?.length === 1
              ? (crossComp.comps[0].set || 'another set')
              : 'other sets';
          }
        }
        if (srcPrice <= 0) return;
        const ratioData = setRatios[`${tv}|||${sv}`];
        if (!ratioData) return;
        const impliedPrice    = srcPrice * ratioData.ratio;
        // Cross-set comps don't have recency data for this set — use a fixed 0.7
        const anchorRecency   = srcCrossSet ? 0.7 : recencyWeight(pvars[sv]);
        const ratioQuality    = ratioData.source === 'sales'      ? 1.0
                              : ratioData.source === 'global-pop' ? 0.9
                              : ratioData.source === 'cross-set'  ? 0.8
                              : ratioData.source === 'corrected'  ? 0.6 : 0.4;
        const anchorStrength  = srcCrossSet
          ? Math.min(1.0, (playerCompMap[`${player}|||${sv}`]?.n || 1) / 5)
          : Math.min(1.0, salesStrength(pvars[sv]) / 10);
        const crossSetPenalty = srcCrossSet ? 0.8 : 1.0;
        const proximity       = computeProximityWeight(variantPop[tv], variantPop[sv]);
        anchors.push({
          sourceVariant: sv,
          sourcePrice:   Math.round(srcPrice),
          ratio:         Math.round(ratioData.ratio * 100) / 100,
          ratioN:        ratioData.n,
          ratioSource:   ratioData.source,
          impliedPrice:  Math.round(impliedPrice),
          weight:        anchorRecency * ratioQuality * Math.max(0.2, anchorStrength) * proximity * crossSetPenalty,
          ...(srcCrossSet ? { crossSet: true, srcSetLabel } : {}),
        });
      });

      // ── 8b. Direct comp: same player + same variant in another set ────────
      const compData = playerCompMap[`${player}|||${tv}`];
      if (compData && compData.price > 0) {
        const sig        = (allData.signalMap || {})[`${setName}|||${player}|||${tv}`];
        const compConf   = compData.n >= 3 ? 'High' : compData.n >= 2 ? 'Medium-High' : 'Medium';
        const compFields = {
          price:          compData.price,
          confidence:     compConf,
          pvs:            buildPvsInfo(player, playerImpliedBase, playerPercentile, pvsSorted),
          refDist:        null,
          rangeEstimate:  null,
          evidenceWeight: Math.round(evidenceWeight * 100),
          ratioEstimate:  null,
          anchors,
          signal:         sig || null,
          compUsed:       true,
          compData,
        };
        if (hasRealData) {
          // Real-data cell: comp is the model estimate — blend with own stale price
          const ownPrice   = bestSalesPrice(pvars[tv]);
          const freshness  = computeFreshness(pvars[tv]?.lastSaleDate);
          const confScores = { 'High': 0.97, 'Medium-High': 0.92, 'Medium': 0.80 };
          const modelQ     = confScores[compConf] || 0.65;
          const effModelWt = (1 - freshness) * modelQ;
          refreshedEstimates[`${player}|||${tv}`] = {
            ...compFields,
            price:         Math.round((1 - effModelWt) * ownPrice + effModelWt * compData.price),
            ownPrice:      Math.round(ownPrice),
            freshness:     Math.round(freshness * 100),
            modelEstimate: compData.price,
            isRefreshed:   true,
          };
        } else {
          estimates[`${player}|||${tv}`] = compFields;
        }
        return;
      }

      // ── 8c. Reference distribution (primary fallback) ─────────────────────
      if (!refDistCache.hasOwnProperty(tv)) {
        refDistCache[tv] = buildReferenceDistribution(tv, variantPop[tv] || null, setName, refSetData);
      }
      const refDistRaw = refDistCache[tv];
      const pct        = playerPercentile[player] ?? 0.5;

      let rangeEstimate = null, extrapolated = false, extrapolatedUp = false;
      if (refDistRaw) {
        const interp = interpolateFromDistribution(pct, refDistRaw.points);
        if (interp) {
          rangeEstimate  = interp.price;
          extrapolated   = interp.extrapolated;
          extrapolatedUp = interp.extrapolatedUp;
        }
      }

      // ── 8d. Ratio estimate (secondary, blended at high evidence) ──────────
      let ratioEstimate = null;
      if (anchors.length > 0) {
        const totalW = anchors.reduce((s, a) => s + a.weight, 0);
        if (totalW > 0) ratioEstimate = anchors.reduce((s, a) => s + a.impliedPrice * a.weight, 0) / totalW;
      }

      // ── 8e. Combine ────────────────────────────────────────────────────────
      let finalEstimate = null;
      if (rangeEstimate !== null) {
        if (ratioEstimate !== null && evidenceWeight > 0.15) {
          // Range is primary; ratios blend in as evidence accumulates.
          // At 30 weighted non-base sales (evidenceWeight ≈ 0.5), it's 50/50.
          finalEstimate = (1 - evidenceWeight) * rangeEstimate + evidenceWeight * ratioEstimate;
        } else {
          finalEstimate = rangeEstimate;
        }
      } else if (ratioEstimate !== null) {
        // No reference distribution yet — use ratio-only as fallback
        finalEstimate = ratioEstimate;
      } else {
        return; // no estimate possible
      }

      finalEstimate = Math.round(finalEstimate);

      // ── Confidence ─────────────────────────────────────────────────────────
      const hasSalesRatio = anchors.some(a => a.ratioSource === 'sales');
      // ratioEstimate is required for any confidence above Very Low.
      // ref-dist without a ratio produced 96.7% MAE in backfill — it provides
      // shape context but no reliable scale, so it can't be trusted on its own.
      const confidence =
        refDistRaw && refDistRaw.sameLeague && !extrapolated && refDistRaw.nPoints >= 3 && ratioEstimate && evidenceWeight > 0.3 ? 'Medium'
      : refDistRaw && refDistRaw.sameLeague && !extrapolated && refDistRaw.nPoints >= 3 && ratioEstimate ? 'Medium-Low'
      : refDistRaw && !extrapolated && ratioEstimate                                                      ? 'Low'
      : refDistRaw                                                                                        ? 'Very Low' // no ratio OR extrapolated
      : hasSalesRatio && anchors.length >= 2                                                              ? 'Low'
      : 'No Data'; // zero usable distribution — no estimate should be trusted

      const sig       = (allData.signalMap || {})[`${setName}|||${player}|||${tv}`];
      const pvsInfo   = buildPvsInfo(player, playerImpliedBase, playerPercentile, pvsSorted);
      const refDistOut = refDistRaw ? {
        refSets:       refDistRaw.refSets,
        sameLeague:    refDistRaw.sameLeague,
        nPoints:       refDistRaw.nPoints,
        rangeMin:      refDistRaw.rangeMin,
        rangeMax:      refDistRaw.rangeMax,
        extrapolated,
        extrapolatedUp,
        targetPop:     refDistRaw.targetPop,
        exactPopMatch: refDistRaw.exactPopMatch,
        scaleFactor:   refDistRaw.scaleFactor,
      } : null;
      const sharedFields = {
        confidence,
        pvs:           pvsInfo,
        refDist:       refDistOut,
        rangeEstimate: rangeEstimate !== null ? Math.round(rangeEstimate) : null,
        evidenceWeight: Math.round(evidenceWeight * 100),
        ratioEstimate: ratioEstimate !== null ? Math.round(ratioEstimate) : null,
        anchors,
        signal:        sig || null,
        compUsed:      false,
        compData:      null,
      };

      if (hasRealData) {
        // ── Refreshed estimate: blend own stale price with current model estimate ──
        // Confidence quality gates the model's effective contribution so a Very Low
        // model estimate doesn't fully override a moderately stale real price.
        const ownPrice     = bestSalesPrice(pvars[tv]);
        const freshness    = computeFreshness(pvars[tv]?.lastSaleDate);
        const confScores   = { 'High': 0.97, 'Medium-High': 0.92, 'Medium': 0.80, 'Medium-Low': 0.60, 'Low': 0.35, 'Very Low': 0.15, 'No Data': 0.05 };
        const modelQuality = confScores[confidence] || 0.20;
        const effModelWt   = (1 - freshness) * modelQuality;
        const effOwnWt     = 1 - effModelWt;
        refreshedEstimates[`${player}|||${tv}`] = {
          ...sharedFields,
          price:         Math.round(effOwnWt * ownPrice + effModelWt * finalEstimate),
          ownPrice:      Math.round(ownPrice),
          freshness:     Math.round(freshness * 100), // 0–100 percent
          modelEstimate: finalEstimate,
          isRefreshed:   true,
        };
        // Real-data cells stay out of estimates{} — gap-fill dict is unchanged.
      } else {
        // ── Gap-fill estimate: no real data, model estimate is authoritative ───────
        estimates[`${player}|||${tv}`] = {
          ...sharedFields,
          price: finalEstimate,
        };
      }
    });
  });

  return { estimates, refreshedEstimates, byPlayer, evidenceWeight, setRatios, pvs };
}
