// ════════════════════════════════════════════════════════════════════════════
// apps-script-snapshots.js — Ghostladder Estimation Engine v2 Jobs
// Google Apps Script (V8 runtime)
//
// ── Setup ────────────────────────────────────────────────────────────────────
// 1. In your GAS project, create a new file called "estimate-engine.gs".
//    Copy the full contents of estimate-engine.js into it verbatim.
//    GAS V8 handles const, let, arrow functions, spread, optional chaining, etc.
//
// 2. Add this file to the same GAS project alongside apps-script-functions.js.
//
// 3. All files in a GAS project share global scope.  The `allData` variable
//    declared below is visible to every function in estimate-engine.gs, so
//    buildEstimates() and its helpers work exactly as they do in the browser.
//    We populate `allData` from live Supabase data before calling buildEstimates.
//
// 4. Run createSnapshotTriggers() once to register weekly + daily schedules.
//
// ── Jobs ─────────────────────────────────────────────────────────────────────
//   takeEstimateSnapshots()  — weekly (Mon 3 AM): estimates every active
//                              player×variant cell, writes to estimate_snapshots
//   resolveSnapshots()       — daily (6 AM): matches unresolved snapshot rows
//                              against new price_intelligence sales, writes actuals
//   updateEngineAccuracy()   — called automatically by resolveSnapshots, or run
//                              manually: recomputes MAE metrics on engine_params
//   createSnapshotTriggers() — run once to register time-based triggers
//
// ── Prerequisites ────────────────────────────────────────────────────────────
//   SB_URL_ and SB_KEY_ are already defined in apps-script-functions.js.
//   The schema-v2-engine.sql file must have been run in Supabase first
//   (creates engine_params, estimate_snapshots, engine_calibration_log).
// ════════════════════════════════════════════════════════════════════════════


// ── Shared global — mirrors the allData object used by the browser app ────────
// estimate-engine.gs reads this global (allData.products, allData.setMath,
// allData.setLeagueMap, allData.signalMap). We assign it fresh before each job.
var allData = {}; // eslint-disable-line no-var

// ── Constants ─────────────────────────────────────────────────────────────────
const SNAP_CHUNK_SIZE_    = 200;  // rows per Supabase INSERT batch
const RESOLVE_CHUNK_SIZE_ = 100;  // parallel PATCH requests per fetchAll call


// ════════════════════════════════════════════════════════════════════════════
// 1.  takeEstimateSnapshots()
//     Run weekly (e.g. Monday 3 AM via createSnapshotTriggers).
//
//     For every active set, runs the full estimation engine and writes one
//     estimate_snapshots row per player×variant cell — both gap-fill estimates
//     (no real sales data) and refreshed estimates (real-data cells with a
//     freshness-weighted model blend).
//
//     All rows in a single job share a snapshot_run_id so you can reconstruct
//     "what did the whole set look like on date X" in a single query.
// ════════════════════════════════════════════════════════════════════════════

function takeEstimateSnapshots() {
  console.log('▶ takeEstimateSnapshots started');

  // ── Load all reference data from Supabase ──────────────────────────────
  const loaded = loadAllData_();
  allData = loaded.allData; // assign global so estimate-engine.gs functions see it

  // ── Get active engine_params ───────────────────────────────────────────
  const engineParams = loadActiveEngineParams_();
  console.log(`  Engine version: ${engineParams.version} (${engineParams.tag || 'no tag'})`);

  // ── Shared run ID for all rows in this execution ───────────────────────
  const runId      = Utilities.getUuid();
  const snapshotAt = new Date().toISOString();
  const allRows    = [];

  for (const set of loaded.activeSets) {
    const setProducts = loaded.allData.productsBySet[set.name] || [];
    console.log(`  Processing: ${set.name} (${setProducts.length} products)`);

    let result;
    try {
      result = buildEstimates(set.name, setProducts, engineParams);
    } catch (e) {
      console.error(`  ✗ buildEstimates failed for "${set.name}": ${e.message}`);
      continue;
    }

    const { estimates, refreshedEstimates } = result;
    const evidenceWeightFraction = result.evidenceWeight; // already 0–1

    // ── Guard: only write combos that exist in set_math ──────────────────
    // set_math is the single source of truth for which player×variant combos
    // are real. price_intelligence can contain orphaned rows (retired combos,
    // import artefacts) that would otherwise generate phantom snapshot rows
    // and corrupt any population-weighted calculation downstream.
    const mathRows    = allData.setMath[set.name] || [];
    const validCombos = new Set(mathRows.map(r => `${r.player}|||${r.variant}`));
    let phantomCount  = 0;

    // ── Gap-fill estimates (cells with no real sales data) ───────────────
    Object.entries(estimates).forEach(([key, est]) => {
      if (!validCombos.has(key)) { phantomCount++; return; }
      const [player, variant] = key.split('|||');
      allRows.push(buildSnapshotRow_({
        runId, snapshotAt,
        setName: set.name, player, variant, est,
        engineVersion:  engineParams.version,
        evidenceWeight: evidenceWeightFraction,
        isRefreshed:    false,
      }));
    });

    // ── Refreshed estimates (real-data cells with freshness blend) ────────
    Object.entries(refreshedEstimates).forEach(([key, est]) => {
      if (!validCombos.has(key)) { phantomCount++; return; }
      const [player, variant] = key.split('|||');
      allRows.push(buildSnapshotRow_({
        runId, snapshotAt,
        setName: set.name, player, variant, est,
        engineVersion:  engineParams.version,
        evidenceWeight: evidenceWeightFraction,
        isRefreshed:    true,
      }));
    });

    if (phantomCount > 0) {
      console.warn(`  ⚠ ${set.name}: skipped ${phantomCount} combos not in set_math`);
    }
  }

  console.log(`  Total snapshot rows to write: ${allRows.length}`);
  sbInsertChunked_('estimate_snapshots', allRows, SNAP_CHUNK_SIZE_);

  console.log(`✓ takeEstimateSnapshots complete — ${allRows.length} rows (run ${runId})`);
}


// ── Build one estimate_snapshots row ──────────────────────────────────────────

function buildSnapshotRow_({ runId, snapshotAt, setName, player, variant, est,
                              engineVersion, evidenceWeight, isRefreshed }) {
  const refDist = est.refDist || null;

  return {
    snapshot_run_id:    runId,
    snapshot_at:        snapshotAt,
    set_name:           setName,
    player:             player,
    variant:            variant,

    // ── Prediction ──────────────────────────────────────────────────────
    estimated_price:    est.price,
    confidence:         est.confidence  || null,
    method:             deriveMethod_(est),

    // ── Engine state at snapshot time ───────────────────────────────────
    engine_version:     engineVersion,
    evidence_weight:    typeof evidenceWeight === 'number' ? evidenceWeight : null,
    n_ref_points:       refDist?.nPoints    || null,
    ref_sets:           refDist?.refSets    || null,
    extrapolated:       refDist?.extrapolated ?? false,
    extrap_dir:         (refDist?.extrapolated)
                          ? (refDist.extrapolatedUp ? 'up' : 'down')
                          : null,

    // ── Real-data context (refreshed cells only) ─────────────────────────
    own_price:          isRefreshed ? (est.ownPrice || null) : null,
    freshness:          isRefreshed ? ((est.freshness || 0) / 100) : null, // store 0–1
    last_sale_date:     null,   // populated by resolveSnapshots when actual comes in
    is_refreshed:       isRefreshed,

    // ── Future: Listings signal (NULL until listing_blend > 0 in engine_params) ─
    listing_price:      null,
    listing_supply:     null,
    listing_component:  null,

    // ── Resolution fields — filled later by resolveSnapshots() ──────────
    actual_price:       null,
    actual_date:        null,
    error_pct:          null,
    resolved_at:        null,
  };
}


// Derive method string from the estimate object produced by buildEstimates()
function deriveMethod_(est) {
  if (est.compUsed)                   return 'direct-comp';
  if (est.refDist && est.ratioEstimate !== null && est.ratioEstimate > 0)
                                      return 'ref-dist+ratio';
  if (est.refDist)                    return 'ref-dist';
  if (est.ratioEstimate !== null)     return 'ratio';
  if (est.signal?.implied_price)      return 'fallback';
  return 'unknown';
}


// ════════════════════════════════════════════════════════════════════════════
// 2.  resolveSnapshots()
//     Run daily (e.g. 6 AM via createSnapshotTriggers).
//
//     Looks for estimate_snapshots rows where actual_price IS NULL.
//     For each, checks whether price_intelligence has a last_sale_date
//     AFTER the snapshot was taken. If so, writes actual_price, actual_date,
//     error_pct, and resolved_at. Then calls updateEngineAccuracy().
// ════════════════════════════════════════════════════════════════════════════

function resolveSnapshots() {
  console.log('▶ resolveSnapshots started');

  // ── Fetch unresolved snapshots ─────────────────────────────────────────
  // Supabase REST: actual_price=is.null  →  WHERE actual_price IS NULL
  // Use a large limit; if you have >5000 unresolved rows, paginate.
  const unresolved = sbGet_('estimate_snapshots',
    'select=id,snapshot_at,set_name,player,variant,estimated_price' +
    '&actual_price=is.null' +
    '&order=snapshot_at.asc' +
    '&limit=5000'
  );

  if (!unresolved.length) {
    console.log('  No unresolved snapshots — nothing to do.');
    return;
  }
  console.log(`  Found ${unresolved.length} unresolved snapshot rows`);

  // ── Fetch current price_intelligence ──────────────────────────────────
  // Only rows that have a last_sale_date (i.e. at least one real sale recorded).
  const piRaw = sbGet_('price_intelligence',
    'select=set_id,player,variant,last_sale_date,last_sale_price' +
    '&last_sale_date=not.is.null' +
    '&last_sale_price=not.is.null'
  );

  // ── We also need the set_id → set_name map ─────────────────────────────
  const setsRaw = sbGet_('sets', 'select=id,name&active=eq.true');
  const setIdToName_ = {};
  setsRaw.forEach(s => { setIdToName_[s.id] = s.name; });

  // Build lookup: "setName|||player|||variant" → { last_sale_date, last_sale_price }
  const piMap = {};
  piRaw.forEach(row => {
    const setName = setIdToName_[row.set_id];
    if (!setName || !row.player || !row.variant) return;
    piMap[`${setName}|||${row.player}|||${row.variant}`] = {
      last_sale_date:  row.last_sale_date,   // 'YYYY-MM-DD'
      last_sale_price: row.last_sale_price,
    };
  });

  // ── Match unresolved snapshots against price_intelligence ──────────────
  const toResolve = [];
  const now       = new Date().toISOString();

  unresolved.forEach(snap => {
    const k  = `${snap.set_name}|||${snap.player}|||${snap.variant}`;
    const pi = piMap[k];
    if (!pi) return;

    // Compare dates: use YYYY-MM-DD portions for a day-level comparison.
    // A sale on the same day as the snapshot is eligible (conservative, safe).
    const snapshotDate = snap.snapshot_at.split('T')[0];  // 'YYYY-MM-DD'
    if (pi.last_sale_date <= snapshotDate) return;        // sale is not newer than snapshot

    const actual    = +pi.last_sale_price;
    const estimated = +snap.estimated_price;
    if (actual <= 0) return;

    // error_pct = |actual − estimated| / actual  (stored as a fraction, e.g. 0.12 = 12%)
    const errorPct = Math.abs(actual - estimated) / actual;

    toResolve.push({
      id:           snap.id,
      actual_price: actual,
      actual_date:  pi.last_sale_date,
      error_pct:    Math.round(errorPct * 100000) / 100000, // 5 dp
      resolved_at:  now,
    });
  });

  if (!toResolve.length) {
    console.log('  No new resolutions found.');
    updateEngineAccuracy(); // still refresh MAE stats even if nothing new
    return;
  }
  console.log(`  Resolving ${toResolve.length} snapshot rows`);

  // ── PATCH resolved rows in parallel batches ────────────────────────────
  // UrlFetchApp.fetchAll() sends up to 100 requests in parallel — far faster
  // than sequential sbPatch_ calls when resolving hundreds of rows.
  let errorCount = 0;
  for (let i = 0; i < toResolve.length; i += RESOLVE_CHUNK_SIZE_) {
    const chunk    = toResolve.slice(i, i + RESOLVE_CHUNK_SIZE_);
    const requests = chunk.map(row => ({
      url: `${SB_URL_}/rest/v1/estimate_snapshots?id=eq.${row.id}`,
      method: 'patch',
      headers: {
        apikey:          SB_KEY_,
        Authorization:   `Bearer ${SB_KEY_}`,
        'Content-Type':  'application/json',
        Prefer:          'return=minimal',
      },
      payload: JSON.stringify({
        actual_price: row.actual_price,
        actual_date:  row.actual_date,
        error_pct:    row.error_pct,
        resolved_at:  row.resolved_at,
      }),
      muteHttpExceptions: true,
    }));

    const responses = UrlFetchApp.fetchAll(requests);
    responses.forEach((res, idx) => {
      if (res.getResponseCode() >= 300) {
        console.error(`  ✗ Patch failed for ${chunk[idx].id}: ${res.getContentText()}`);
        errorCount++;
      }
    });
    console.log(`  Patched rows ${i + 1}–${Math.min(i + RESOLVE_CHUNK_SIZE_, toResolve.length)}`);
  }

  console.log(`✓ resolveSnapshots complete — ${toResolve.length - errorCount} resolved` +
              (errorCount ? `, ${errorCount} errors` : ''));

  // ── Refresh engine accuracy metrics after new resolutions ──────────────
  updateEngineAccuracy();
}


// ════════════════════════════════════════════════════════════════════════════
// 3.  updateEngineAccuracy()
//     Recomputes MAE fields on engine_params from all resolved snapshots.
//     Called automatically at the end of resolveSnapshots(); also safe to
//     run standalone at any time (e.g. after a manual data backfill).
//
//     Updates per engine version:
//       mae_all_time       — mean absolute error across all resolved snapshots
//       mae_30d            — MAE on snapshots resolved in the last 30 days
//       mae_by_confidence  — MAE per confidence tier (JSON)
//       mae_by_method      — MAE per estimation method (JSON)
//       total_snapshots    — total rows in estimate_snapshots for this version
//       resolved_snapshots — rows with actual_price NOT NULL
// ════════════════════════════════════════════════════════════════════════════

function updateEngineAccuracy() {
  console.log('▶ updateEngineAccuracy started');

  // ── Resolved snapshot error data ───────────────────────────────────────
  const resolved = sbGet_('estimate_snapshots',
    'select=engine_version,confidence,method,error_pct,resolved_at' +
    '&actual_price=not.is.null' +
    '&error_pct=not.is.null'
  );

  if (!resolved.length) {
    console.log('  No resolved snapshots — skipping MAE update.');
    return;
  }

  // ── Total snapshot count per engine version ────────────────────────────
  // Fetch all rows (just the version column) to get total_snapshots correctly.
  const allSnaps = sbGet_('estimate_snapshots', 'select=engine_version');
  const totalByVersion = {};
  allSnaps.forEach(row => {
    const v = row.engine_version;
    totalByVersion[v] = (totalByVersion[v] || 0) + 1;
  });

  // ── Group resolved snapshots by engine version ─────────────────────────
  const cutoff30d = new Date(Date.now() - 30 * 86400000).toISOString();

  const byVersion = {};
  resolved.forEach(row => {
    const v = String(row.engine_version);
    if (!byVersion[v]) {
      byVersion[v] = {
        allErrors:    [],
        errors30d:    [],
        byConfidence: {},
        byMethod:     {},
        resolvedCount: 0,
      };
    }
    const g = byVersion[v];
    const e = +row.error_pct;
    g.allErrors.push(e);
    g.resolvedCount++;

    if (row.resolved_at >= cutoff30d) g.errors30d.push(e);

    const conf = row.confidence || 'Unknown';
    if (!g.byConfidence[conf]) g.byConfidence[conf] = [];
    g.byConfidence[conf].push(e);

    const meth = row.method || 'unknown';
    if (!g.byMethod[meth]) g.byMethod[meth] = [];
    g.byMethod[meth].push(e);
  });

  const mean_ = arr => arr.length
    ? Math.round((arr.reduce((s, v) => s + v, 0) / arr.length) * 100000) / 100000
    : null;

  // ── PATCH each engine_params version ──────────────────────────────────
  Object.entries(byVersion).forEach(([version, data]) => {
    const maeByConf   = {};
    const maeByMethod = {};

    Object.entries(data.byConfidence).forEach(([k, arr]) => { maeByConf[k]   = mean_(arr); });
    Object.entries(data.byMethod)    .forEach(([k, arr]) => { maeByMethod[k] = mean_(arr); });

    const maeAll  = mean_(data.allErrors);
    const mae30d  = mean_(data.errors30d);
    const total   = totalByVersion[version] || data.resolvedCount;

    try {
      sbPatch_('engine_params', `version=eq.${version}`, {
        mae_all_time:       maeAll,
        mae_30d:            mae30d,
        mae_by_confidence:  maeByConf,
        mae_by_method:      maeByMethod,
        total_snapshots:    total,
        resolved_snapshots: data.resolvedCount,
      });
      const s = `v${version}: MAE=${maeAll?.toFixed(4) ?? 'n/a'} (${data.resolvedCount} resolved / ${total} total)`;
      console.log(`  ✓ engine_params ${s}`);
    } catch (e) {
      console.error(`  ✗ Failed to update engine_params v${version}: ${e.message}`);
    }
  });

  console.log('✓ updateEngineAccuracy complete');
}


// ════════════════════════════════════════════════════════════════════════════
// 4.  Data loading helpers
// ════════════════════════════════════════════════════════════════════════════

/**
 * Loads all data from Supabase and returns:
 *   activeSets  — array of { id, name, league, figures_per_case, hit_ratio }
 *   allData     — object matching the shape that estimate-engine.gs expects:
 *     .products      — all price_intelligence rows (camelCase, with .set property)
 *     .productsBySet — { [setName]: product[] }  (used to pass setProducts arg)
 *     .setMath       — { [setName]: { player, variant, population }[] }
 *     .setLeagueMap  — { [setName]: leagueString }
 *     .signalMap     — { ["setName|||player|||variant"]: price_signals row }
 */
function loadAllData_() {
  console.log('  Loading data from Supabase...');

  // ── Sets ───────────────────────────────────────────────────────────────
  const setsRaw = sbGet_('sets',
    'select=id,name,league,figures_per_case,hit_ratio' +
    '&active=eq.true' +
    '&order=name.asc'
  );

  const setLeagueMap = {};
  const setIdToName_ = {};
  setsRaw.forEach(s => {
    setLeagueMap[s.name]  = s.league || '';
    setIdToName_[s.id]    = s.name;
  });

  // ── price_intelligence — all active sets ───────────────────────────────
  // Supabase REST: select columns we need, no join — we resolve set_id in JS.
  const piRaw = sbGet_('price_intelligence',
    'select=set_id,player,variant,' +
    'avg_price,avg_30d,avg_90d,total_sales,sales_30d,sales_90d,' +
    'last_sale_date,last_sale_price'
  );

  // Map to camelCase product objects (the shape estimate-engine.js expects)
  // and build both the flat array (allData.products) and per-set index.
  const allProducts   = [];
  const productsBySet = {};

  piRaw.forEach(row => {
    const setName = setIdToName_[row.set_id];
    if (!setName) return; // orphaned row — set may be inactive

    const product = {
      set:           setName,
      player:        row.player,
      variant:       row.variant,
      avgPrice:      row.avg_price      || 0,
      avg30d:        row.avg_30d        || 0,
      avg90d:        row.avg_90d        || 0,
      totalSales:    row.total_sales    || 0,
      sales30d:      row.sales_30d      || 0,
      sales90d:      row.sales_90d      || 0,
      lastSaleDate:  row.last_sale_date  || null,
      lastSalePrice: row.last_sale_price || 0,
    };

    allProducts.push(product);
    if (!productsBySet[setName]) productsBySet[setName] = [];
    productsBySet[setName].push(product);
  });

  // ── set_math — variant populations ────────────────────────────────────
  const mathRaw = sbGet_('set_math',
    'select=set_id,player,variant,population,real_population' +
    '&order=set_id.asc'
  );

  const setMath = {};
  mathRaw.forEach(row => {
    const setName = setIdToName_[row.set_id];
    if (!setName) return;
    if (!setMath[setName]) setMath[setName] = [];
    setMath[setName].push({
      player:          row.player,
      variant:         row.variant,
      population:      row.population,
      real_population: row.real_population,
    });
  });

  // ── price_signals — implied_price fallback + listing signals ──────────
  const signalsRaw = sbGet_('price_signals',
    'select=set_name,player,variant,' +
    'implied_price,implied_confidence,' +
    'lowest_ask,median_ask,supply_count,' +
    'avg_30d,sales_30d,last_sale_price,last_sale_date'
  );

  const signalMap = {};
  signalsRaw.forEach(row => {
    if (!row.set_name || !row.player || !row.variant) return;
    signalMap[`${row.set_name}|||${row.player}|||${row.variant}`] = row;
  });

  console.log(
    `  Loaded: ${setsRaw.length} sets, ${allProducts.length} products, ` +
    `${mathRaw.length} set_math rows, ${signalsRaw.length} signals`
  );

  return {
    activeSets: setsRaw,
    allData: {
      products:     allProducts,   // all sets combined — for cross-set reference
      productsBySet,               // per-set slice — for setProducts arg to buildEstimates
      setMath,
      setLeagueMap,
      signalMap,
    },
  };
}


/**
 * Returns the single is_active=true row from engine_params.
 * Throws if none is found (means schema-v2-engine.sql hasn't been run).
 */
function loadActiveEngineParams_() {
  const rows = sbGet_('engine_params', 'select=*&is_active=eq.true&limit=1');
  if (!rows.length) {
    throw new Error(
      'No active engine_params row found. ' +
      'Run schema-v2-engine.sql in Supabase to create the tables and seed v1.'
    );
  }
  return rows[0];
}


// ── Chunked INSERT (no upsert — snapshots are append-only) ──────────────────
function sbInsertChunked_(table, rows, chunkSize) {
  if (!rows.length) return;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);
    const res = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/${table}`, {
      method: 'post',
      headers: {
        apikey:          SB_KEY_,
        Authorization:   `Bearer ${SB_KEY_}`,
        'Content-Type':  'application/json',
        Prefer:          'return=minimal',
      },
      payload:           JSON.stringify(chunk),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() >= 300) {
      throw new Error(
        `sbInsertChunked_ ${table} chunk ${i / chunkSize + 1}: ${res.getContentText()}`
      );
    }
    console.log(`  Inserted rows ${i + 1}–${Math.min(i + chunkSize, rows.length)} into ${table}`);
  }
}


// ════════════════════════════════════════════════════════════════════════════
// 5.  Trigger management
// ════════════════════════════════════════════════════════════════════════════

/**
 * Run once to register time-based triggers.
 * Safe to run again — removes existing snapshot triggers first to avoid duplicates.
 *
 *   takeEstimateSnapshots  →  every Monday 3–4 AM
 *   resolveSnapshots       →  every day 6–7 AM
 */
function createSnapshotTriggers() {
  const MANAGED = ['takeEstimateSnapshots', 'resolveSnapshots'];

  // Remove existing triggers for managed functions
  ScriptApp.getProjectTriggers()
    .filter(t => MANAGED.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Weekly snapshot: Monday 3–4 AM
  ScriptApp.newTrigger('takeEstimateSnapshots')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();

  // Daily resolve: 6–7 AM
  ScriptApp.newTrigger('resolveSnapshots')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  console.log('✓ Triggers registered:');
  console.log('  takeEstimateSnapshots — every Monday 3–4 AM');
  console.log('  resolveSnapshots      — every day 6–7 AM');
}


/** List active snapshot-related triggers. */
function listSnapshotTriggers() {
  const fns = ['takeEstimateSnapshots', 'resolveSnapshots', 'updateEngineAccuracy'];
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => fns.includes(t.getHandlerFunction()));
  if (!triggers.length) {
    console.log('No snapshot triggers found. Run createSnapshotTriggers() to set them up.');
    return;
  }
  triggers.forEach(t => console.log(`  ${t.getHandlerFunction()} — ${t.getTriggerSource()}`));
  console.log(`Total: ${triggers.length} trigger(s)`);
}


// ════════════════════════════════════════════════════════════════════════════
// 6.  Debugging / manual utilities
// ════════════════════════════════════════════════════════════════════════════

/**
 * Dry-run estimate for a single set — does NOT write to Supabase.
 * Useful for testing the estimation engine in GAS before running the full job.
 *
 * Usage: change SET_NAME below and run testEstimateForSet_()
 */
function testEstimateForSet_() {
  const SET_NAME = 'MLB 2025 Game Face'; // ← change to any active set name

  const loaded = loadAllData_();
  allData = loaded.allData;

  const engineParams  = loadActiveEngineParams_();
  const setProducts   = loaded.allData.productsBySet[SET_NAME];

  if (!setProducts) {
    console.log(`Set "${SET_NAME}" not found. Active sets:`);
    loaded.activeSets.forEach(s => console.log('  ' + s.name));
    return;
  }

  console.log(`Running estimation for "${SET_NAME}" (${setProducts.length} products)…`);
  const result = buildEstimates(SET_NAME, setProducts, engineParams);

  const gapCount   = Object.keys(result.estimates).length;
  const freshCount = Object.keys(result.refreshedEstimates).length;
  console.log(`  Gap-fill estimates:   ${gapCount}`);
  console.log(`  Refreshed estimates:  ${freshCount}`);
  console.log(`  Evidence weight:      ${(result.evidenceWeight * 100).toFixed(1)}%`);

  // Sample a few gap-fill estimates
  const sample = Object.entries(result.estimates).slice(0, 5);
  sample.forEach(([key, est]) => {
    console.log(`  ${key}: $${est.price} (${est.confidence}, ${deriveMethod_(est)})`);
  });

  console.log('✓ Test complete — nothing was written to Supabase.');
}


/**
 * Show a summary of the most recent snapshot run: how many rows, per-set breakdown.
 */
function lastSnapshotSummary() {
  // Get the most recent snapshot_run_id
  const latest = sbGet_('estimate_snapshots',
    'select=snapshot_run_id,snapshot_at' +
    '&order=snapshot_at.desc' +
    '&limit=1'
  );
  if (!latest.length) { console.log('No snapshots found.'); return; }

  const { snapshot_run_id, snapshot_at } = latest[0];
  console.log(`Most recent run: ${snapshot_run_id} at ${snapshot_at}`);

  const rows = sbGet_('estimate_snapshots',
    `select=set_name,is_refreshed,confidence` +
    `&snapshot_run_id=eq.${snapshot_run_id}`
  );

  // Group by set
  const bySet = {};
  rows.forEach(r => {
    if (!bySet[r.set_name]) bySet[r.set_name] = { gap: 0, refreshed: 0, byConf: {} };
    const s = bySet[r.set_name];
    if (r.is_refreshed) s.refreshed++; else s.gap++;
    s.byConf[r.confidence] = (s.byConf[r.confidence] || 0) + 1;
  });

  Object.entries(bySet).sort().forEach(([set, d]) => {
    console.log(`  ${set}: ${d.gap} gap-fill, ${d.refreshed} refreshed`);
  });
  console.log(`  Total: ${rows.length} rows`);
}
