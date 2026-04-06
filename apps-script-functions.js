// ── Custom formula: eBay search URL with dynamic startDate ───────────────────
// Usage in SEARCH URLS tab: =EBAY_URL("your base keywords here")
// Or with a custom base URL: =EBAY_URL_CUSTOM(A2)
//
// Returns full eBay URL with startDate set to (most recent sale date - 1 day)

function EBAY_URL(keywords) {
  const startMs = getEbayStartDate_();
  const base = "https://www.ebay.com/sh/research?marketplace=EBAY-US"
    + "&keywords=" + encodeURIComponent(keywords || "ghostwrite")
    + "&dayRange=CUSTOM"
    + "&startDate=" + startMs
    + "&categoryId=0&offset=0&limit=50&sorting=-datelastsold&tabName=SOLD&tz=America%2FNew_York";
  return base;
}

// Use this if you want to pass in a full existing URL and just update the startDate
function EBAY_UPDATE_URL(existingUrl) {
  if (!existingUrl) return "No URL provided";
  const startMs = getEbayStartDate_();
  // Replace existing startDate param
  const updated = String(existingUrl).replace(
    /startDate=\d+/,
    "startDate=" + startMs
  );
  // If no startDate param existed, append it
  if (updated === String(existingUrl)) {
    return existingUrl + "&startDate=" + startMs;
  }
  return updated;
}

function getEbayStartDate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  // Read col J (index 10, 1-based) — Date Sold; SALES data starts row 2
  const dates = sh.getRange(2, 10, lastRow - 1, 1).getValues();

  let maxDate = null;
  for (const row of dates) {
    const d = row[0];
    if (d instanceof Date && !isNaN(d)) {
      if (!maxDate || d > maxDate) maxDate = d;
    }
  }

  if (!maxDate) return 0;

  // Subtract 1 day and set to midnight ET
  const oneDayMs = 24 * 60 * 60 * 1000;
  const startDate = new Date(maxDate.getTime() - oneDayMs);
  startDate.setHours(0, 0, 0, 0);

  return startDate.getTime();
}


function showEbayStartDate() {
  const ms = getEbayStartDate_();
  if (!ms) { SpreadsheetApp.getUi().alert("No dates found in SALES sheet col J."); return; }
  const d = new Date(ms);
  SpreadsheetApp.getUi().alert(
    "Most recent sale minus 1 day:\n\n" +
    d.toLocaleDateString('en-US') + "\n\n" +
    "Unix ms: " + ms + "\n\n" +
    "Paste into your eBay URL as:\nstartDate=" + ms
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ghostwrite")
    .addItem("Backfill Supabase (run once)", "backfillSupabase")
    .addItem("Sync selected rows to Supabase", "syncSelectedRowsToSupabase")
    .addItem("Sync set_math names from Sales", "syncSetMathNamesFromSales")
    .addItem("Re-clean selected rows (SALES)", "cleanSelectedRowsSales")
    .addItem("Re-clean selected rows (LISTINGS) + sync to SB", "cleanAndSyncSelectedListings")
    .addSeparator()
    .addItem("Run live listings ingestion", "runLiveIngestion")
    .addItem("Schedule auto-ingestion (every 3 days)", "setupLiveIngestionTrigger")
    .addItem("Stop auto-ingestion", "stopLiveIngestionTrigger")
    .addItem("Emergency stop ingestion", "stopLiveIngestion")
    .addSeparator()
    .addItem("Import today's eBay scrape → SALES", "importEbayScrapeCSV")
    .addItem("Import historical scrape → SALES", "importHistoricalScrape")
    .addItem("Backfill Set Math IDs (col W)", "backfillSetMathIds")
    .addItem("Schedule daily CSV import (1pm)", "setupDailyImportTrigger")
    .addItem("Stop daily CSV import", "stopDailyImportTrigger")
        .addItem("Sync SALES → Supabase", "syncSalesToSupabase")
    .addSeparator()
    .addItem("Compute AFA grading multiplier", "computeGradingMultiplier")
    .addSeparator()
    .addItem("Pull Price Intelligence → sheet", "pullPriceIntelligence")
    .addItem("Pull Estimate Snapshots → sheet", "pullEstimateSnapshots")
    .addItem("Push Inventory → Supabase", "pushInventoryToSupabase")
    .addToUi();
}

// ── API key storage ───────────────────────────────────────────────────────────
function setupApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "Claude API Key",
    "Paste your Anthropic API key (starts with sk-ant-).",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (!key.startsWith("sk-ant-")) {
      ui.alert("That doesn't look right — key should start with sk-ant-");
      return;
    }
    PropertiesService.getUserProperties().setProperty("CLAUDE_API_KEY", key);
    ui.alert("API key saved.");
  }
}

function getApiKey() {
  const key = PropertiesService.getUserProperties().getProperty("CLAUDE_API_KEY");
  if (!key) throw new Error("No API key found. Run ghostwrite > Set API key first.");
  return key;
}

function testApiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const setCatalog = loadSetCatalogFromSupabase();
    const result = classifyBatch(
      ["Ghostwrite NBA Game Face Luka Doncic Victory /50 100%"],
      setCatalog
    );
    ui.alert(
      "Connection works!\n\n" +
      "Test title: 'Ghostwrite NBA Game Face Luka Doncic Victory /50 100%'\n\n" +
      "Set: "     + result[0].set     + "\n" +
      "Player: "  + result[0].player  + "\n" +
      "Variant: " + result[0].variant + "\n" +
      "Exclude: " + result[0].exclude +
      (result[0].reason ? " (" + result[0].reason + ")" : "") + "\n" +
      "Inferred: " + (result[0].inferred ? "yes — needs review" : "no")
    );
  } catch(e) {
    ui.alert("Connection failed: " + e.message);
  }
}


// ── Normalize name to Title Case ─────────────────────────────────────────────
function toTitleCase(str) {
  if (!str) return str;
  // Preserve Jr. Sr. II III IV suffixes
  return String(str).toLowerCase().replace(/\b\w+/g, word => {
    // Keep suffixes as-is after normalizing
    if (['jr.','sr.','ii','iii','iv','v'].includes(word)) {
      return word === 'jr.' ? 'Jr.' : word === 'sr.' ? 'Sr.' : word.toUpperCase();
    }
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
}

// ── Load PLAYER ROSTER from Supabase set_math (canonical source) ──────────────
// Format: { "Set Name": [{player, basePop, variants}] }
function loadPlayerRosterFromSupabase() {
  try {
    const { url, key } = getSupabaseConfig_();
    const headers = { 'apikey': key, 'Authorization': 'Bearer ' + key };

    // Fetch sets (id → name, figures_per_case)
    const setsResp = UrlFetchApp.fetch(
      url + '/rest/v1/sets?select=id,name,figures_per_case,active&active=eq.true',
      { method: 'GET', headers, muteHttpExceptions: true }
    );
    if (setsResp.getResponseCode() !== 200) throw new Error('sets fetch failed');
    const sets = JSON.parse(setsResp.getContentText());
    const setIdToName = {};
    sets.forEach(s => { setIdToName[s.id] = s.name; });

    // Fetch set_math (all rows)
    const mathResp = UrlFetchApp.fetch(
      url + '/rest/v1/set_math?select=set_id,player,variant,population,real_population&order=set_id.asc,player.asc,population.desc',
      { method: 'GET', headers, muteHttpExceptions: true }
    );
    if (mathResp.getResponseCode() !== 200) throw new Error('set_math fetch failed');
    const mathRows = JSON.parse(mathResp.getContentText());

    // Group by set, then by player — collect variants and base pop
    // roster shape: { setName: [{player, basePop, variants}] }
    const roster = {};
    const playerVariants = {}; // "setName|||player" -> [{variant, population}]

    mathRows.forEach(row => {
      const setName = setIdToName[row.set_id];
      if (!setName || !row.player || row.player === 'Unknown') return;
      const key = setName + '|||' + row.player;
      if (!playerVariants[key]) playerVariants[key] = { setName, player: row.player, base: null, variants: [] };
      // Use ONLY real_population — this is the number printed on the figure and used in eBay listings.
      // If real_population is missing, omit the population from the roster line entirely
      // rather than falling back to the internal population count (which sellers never use).
      const displayPop = row.real_population || null;
      if (row.variant === 'Base') {
        playerVariants[key].base = displayPop;
      } else {
        // Only include variant in roster if we have a real population to match against
        if (displayPop) playerVariants[key].variants.push(row.variant + ' /' + displayPop);
        else            playerVariants[key].variants.push(row.variant); // name only, no pop
      }
    });

    Object.values(playerVariants).forEach(({ setName, player, base, variants }) => {
      if (!roster[setName]) roster[setName] = [];
      roster[setName].push({
        player,
        basePop: base ? String(base) : '',
        variants: variants.length ? variants.join(', ') : '',
      });
    });

    const totalPlayers = Object.values(roster).reduce((s, arr) => s + arr.length, 0);
    Logger.log('Loaded roster from Supabase: ' + totalPlayers + ' players across ' + Object.keys(roster).length + ' sets');
    return roster;

  } catch(e) {
    Logger.log('Supabase roster load failed: ' + e.message);
    return {};
  }
}


// Format roster for a specific set for inclusion in AI prompt
function formatRosterForPrompt(roster, setName) {
  const players = roster[setName];
  if (!players || !players.length) return null;

  return players.map(p =>
    `  - ${p.player} (base: ${p.basePop}${p.variants ? ' | variants: ' + p.variants : ''})`
  ).join("\n");
}

// ── Format set catalog for AI prompt ─────────────────────────────────────────
function formatCatalogForPrompt(catalog) {
  if (!catalog.length) return "No sets defined yet.";

  return catalog.map((s, i) =>
    `${i + 1}. SET NAME: "${s.setName}"
   League: ${s.league} | Year: ${s.year} | Base population: ${s.basePop}
   Keywords to look for: ${s.keywords}`
  ).join("\n\n");
}

// ── Main: clean all unclassified rows ─────────────────────────────────────────
// ── Auto-loop cleaning ───────────────────────────────────────────────────────
// Starts the auto-loop. Sets a trigger to re-run cleanNewRows every minute
// until all rows are done, then emails + writes status to sheet.


function setCleanStatus(msg) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'ghostwrite', 5);
  } catch(e) {}
}

// ─────────────────────────────────────────────────────────────────────────────
// loadEstimateMap_
// Returns a map of "setName|||player|||variant" → estimated_price (number)
// built from the most recent snapshot run in estimate_snapshots.
// Used by cleanNewSalesRows to auto-flag prices that deviate wildly from
// the current model estimate.
// ─────────────────────────────────────────────────────────────────────────────
function loadEstimateMap_() {
  try {
    // Fetch the latest snapshot_run_id so we only read the freshest estimates.
    const latest = sbGet_('estimate_snapshots',
      'select=snapshot_run_id&order=snapshot_at.desc&limit=1'
    );
    if (!latest.length) return {};

    const runId = latest[0].snapshot_run_id;
    const rows  = sbGet_('estimate_snapshots',
      'select=set_name,player,variant,estimated_price' +
      '&snapshot_run_id=eq.' + encodeURIComponent(runId)
    );

    const map = {};
    rows.forEach(function(r) {
      if (!r.set_name || !r.player || !r.variant || !r.estimated_price) return;
      const key = r.set_name + '|||' + r.player + '|||' + r.variant;
      // Prefer refreshed estimates (is_refreshed=true) if both exist —
      // but since we deduplicate by first seen, run takeEstimateSnapshots so
      // refreshed rows come first. In practice one row per combo per run.
      if (!map[key]) map[key] = +r.estimated_price;
    });
    Logger.log('Loaded ' + Object.keys(map).length + ' estimates for auto-flag check');
    return map;
  } catch(e) {
    Logger.log('loadEstimateMap_ skipped: ' + e.message);
    return {};
  }
}

// Thresholds for auto-flagging suspicious import prices.
// A row is flagged when its price is MORE than AUTO_FLAG_RATIO_HIGH × the estimate
// OR LESS than AUTO_FLAG_RATIO_LOW × the estimate, and the estimate is at least
// AUTO_FLAG_MIN_ESTIMATE dollars (avoids noise on very cheap cards).
var AUTO_FLAG_RATIO_HIGH   = 4.0;   // price > 4× estimate  → suspicious
var AUTO_FLAG_RATIO_LOW    = 0.20;  // price < 20% estimate → suspicious
var AUTO_FLAG_MIN_ESTIMATE = 5;     // skip check if estimate is under $5

// ─────────────────────────────────────────────────────────────────────────────
// cleanNewSalesRows  — same as cleanNewRows but targets the SALES sheet.
// Called by importEbayScrapeCSV after new rows are appended; also runs as a
// 1-minute time-based trigger until all unclassified rows are processed.
// ─────────────────────────────────────────────────────────────────────────────
function cleanNewSalesRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sheet) { Logger.log('SALES sheet not found'); return; }

  // Load catalogs.
  // setCatalog is optional fallback context for the AI prompt — classification works
  // fine without it (player roster is the primary signal). Never bail on catalog failure.
  // playerRoster is essential — bail if it fails.
  let setCatalog = [], playerRoster;
  try { setCatalog = loadSetCatalogFromSupabase(); } catch(e) {
    Logger.log('SET CATALOG load skipped (tab missing?): ' + e.message);
  }
  try {
    playerRoster = loadPlayerRosterFromSupabase();
  } catch(e) {
    setCleanStatus('❌ Error loading player roster: ' + e.message);
    stopSalesCleanTriggers_();
    return;
  }

  // Build set_math UUID cache for col W backfill — non-fatal if unavailable
  let setMathCache = {};
  try { setMathCache = buildSetMathCache_(); } catch(e) {
    Logger.log('Set math cache skipped: ' + e.message);
  }

  // Load most-recent estimates for auto-flagging suspicious prices — non-fatal
  let estimateMap = {};
  try { estimateMap = loadEstimateMap_(); } catch(e) {
    Logger.log('Estimate map skipped: ' + e.message);
  }

  const lastRow   = sheet.getLastRow();
  const props     = PropertiesService.getScriptProperties();
  const resumeRow = Math.max(2, Number(props.getProperty('CLEAN_RESUME_ROW') || 2));

  // Collect up to 200 unclassified rows starting from resumeRow
  const toProcess = [];
  if (lastRow >= resumeRow) {
    const rangeCount = lastRow - resumeRow + 1;
    const batchData  = sheet.getRange(resumeRow, 1, rangeCount, 21).getValues(); // A-U
    for (let i = 0; i < batchData.length; i++) {
      if (toProcess.length >= 200) break;
      const title   = batchData[i][1];  // col B
      const player  = batchData[i][13]; // col N
      const variant = batchData[i][14]; // col O
      const set_    = batchData[i][15]; // col P
      if (title && String(title).trim() && !player && !variant && !set_) {
        toProcess.push({
          row:   resumeRow + i,
          title: String(title).trim(),
          price: Number(batchData[i][12]) || 0,  // col M — for auto-flag check
        });
      }
    }
  }

  if (!toProcess.length) {
    stopSalesCleanTriggers_();
    props.deleteProperty('CLEAN_RESUME_ROW');
    setCleanStatus('✅ All rows cleaned + synced to Supabase!');
    return;
  }

  // Process in batches of 8
  const BATCH = 8;
  for (let i = 0; i < toProcess.length; i += BATCH) {
    const batch = toProcess.slice(i, i + BATCH);
    try {
      const results = classifyBatch(batch.map(function(r) { return r.title; }), setCatalog, playerRoster);
      const syncRows = [];
      for (let j = 0; j < batch.length; j++) {
        const row    = batch[j].row;
        const result = results[j] || { set: '', player: '', variant: '', exclude: false, inferred: false, afa: '' };
        sheet.getRange(row, 14).setValue(toTitleCase(result.player  || ''));
        sheet.getRange(row, 15).setValue(result.variant || '');
        sheet.getRange(row, 16).setValue(result.set     || '');
        sheet.getRange(row, 21).setValue(result.afa     || '');
        // Col R (18): notes — collect warnings before writing
        const notesParts = [];
        if (result.note)    notesParts.push(result.note);
        if (result.inferred) notesParts.push('⚠ set inferred — please verify');
        if (result.exclude)  sheet.getRange(row, 17).setValue(result.reason || 'exclude');

        // ── Auto-flag suspicious prices ──────────────────────────────────────
        // Compare this row's price to the most recent model estimate.
        // Flag if price is more than 4× or less than 20% of the estimate
        // (and the estimate is at least $5 — avoids noise on very cheap cards).
        if (!result.exclude && result.set && result.player && result.variant) {
          const salePrice  = Number(batch[j].price) || 0;
          const estKey     = result.set + '|||' + toTitleCase(result.player || '') + '|||' + (result.variant || '');
          const estimate   = estimateMap[estKey] || 0;
          if (salePrice > 0 && estimate >= AUTO_FLAG_MIN_ESTIMATE) {
            const ratio = salePrice / estimate;
            if (ratio > AUTO_FLAG_RATIO_HIGH || ratio < AUTO_FLAG_RATIO_LOW) {
              sheet.getRange(row, 19).setValue('FLAG'); // col S
              notesParts.push(
                '⚠ Auto-flagged: $' + salePrice +
                ' vs est. $' + Math.round(estimate) +
                ' (' + Math.round(ratio * 100) + '%)'
              );
              Logger.log('Auto-flagged row ' + row + ': ' + estKey +
                         ' price=$' + salePrice + ' est=$' + Math.round(estimate));
            }
          }
        }

        if (notesParts.length) sheet.getRange(row, 18).setValue(notesParts.join(' | '));

        // Write set_math UUID to col W (23)
        const smKey = (result.set || '') + '|||' + toTitleCase(result.player || '') + '|||' + (result.variant || '');
        const smId  = setMathCache[smKey] || '';
        if (smId) sheet.getRange(row, 23).setValue(smId);
        syncRows.push(row);
      }
      try { syncSalesRowsToSupabase_(syncRows); } catch(e) { Logger.log('Supabase sync skipped: ' + e.message); }
    } catch(e) {
      for (let j = 0; j < batch.length; j++) {
        sheet.getRange(batch[j].row, 18).setValue('AI error: ' + e.message);
      }
    }
    // Only sleep between batches, not after the last one
    if (i + BATCH < toProcess.length) Utilities.sleep(15000);
  }

  // Advance resume pointer
  if (toProcess.length > 0) {
    props.setProperty('CLEAN_RESUME_ROW', String(toProcess[toProcess.length - 1].row + 1));
  }

  const remaining = countUnclassifiedSales_(Number(props.getProperty('CLEAN_RESUME_ROW') || 2));
  setCleanStatus('🔄 ' + remaining + ' rows remaining...');

  // Schedule next trigger run if still work to do
  const hasTrigger = ScriptApp.getProjectTriggers()
    .some(function(t) { return t.getHandlerFunction() === 'cleanNewSalesRows'; });
  if (!hasTrigger && remaining > 0) {
    ScriptApp.newTrigger('cleanNewSalesRows').timeBased().everyMinutes(1).create();
  }

  if (remaining === 0) {
    stopSalesCleanTriggers_();
    props.deleteProperty('CLEAN_RESUME_ROW');
    setCleanStatus('✅ All rows cleaned + synced to Supabase!');
  }
}

// Deletes any active cleanNewSalesRows time-based triggers.
function stopSalesCleanTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'cleanNewSalesRows') ScriptApp.deleteTrigger(t);
  });
}

// Syncs a list of 1-based SALES row numbers to Supabase.
function syncSalesRowsToSupabase_(rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_SHEET_NAME);
    if (!sheet) return;
    const setIdCache = getSetIdCache_();
    const records = [];
    rowNumbers.forEach(function(rowNum) {
      const row    = sheet.getRange(rowNum, 1, 1, 23).getValues()[0]; // A-W (incl set_math_id)
      const record = sheetRowToSalesRecord_(row, rowNum, setIdCache);
      if (record) records.push(record);
    });
    if (records.length) {
      const result = upsertSalesToSupabase_(records);
      Logger.log('Synced ' + result.inserted + ' SALES rows to Supabase (' + result.errors + ' errors)');
    }
  } catch(e) {
    Logger.log('Supabase sync error (non-fatal): ' + e.message);
  }
}

// Counts unclassified SALES rows (has title, no player) starting from fromRow.
function countUnclassifiedSales_(fromRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_SHEET_NAME);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  const start   = Math.max(2, fromRow || 2);
  if (lastRow < start) return 0;
  const data = sheet.getRange(start, 1, lastRow - start + 1, 14).getValues();
  let count = 0;
  for (const row of data) {
    const title  = row[1];  // col B
    const player = row[13]; // col N
    if (title && String(title).trim() && !player) count++;
  }
  return count;
}

// ── Core AI batch classification ──────────────────────────────────────────────
function classifyBatch(titles, setCatalog, playerRoster) {
  const apiKey = getApiKey();
  const catalogText = formatCatalogForPrompt(setCatalog);
  const numberedTitles = titles.map((t, i) => (i + 1) + ". " + t).join("\n");

  // Build filtered roster — only include sets plausibly relevant to this batch
  // This keeps prompt size small and avoids input token rate limits
  const allRosterLines = [];
  if (playerRoster) {
    // Extract candidate player names and pop numbers from titles for pre-filtering
    const titlesLower = titles.map(t => t.toLowerCase());

    for (const [setName, players] of Object.entries(playerRoster)) {
      const setLower = setName.toLowerCase();
      // Include this set if any title contains a word from the set name,
      // OR if any player last name appears in any title
      const setWords = setLower.split(/\s+/).filter(w => w.length > 3);
      const setMatch = titlesLower.some(t => setWords.some(w => t.includes(w)));
      const playerMatch = players.some(p => {
        const lastName = p.player.split(' ').pop().toLowerCase().replace('.','');
        return titlesLower.some(t => t.includes(lastName));
      });

      if (setMatch || playerMatch) {
        for (const p of players) {
          allRosterLines.push(
            `${p.player} | base: ${p.basePop} | set: "${setName}"${p.variants ? ' | variants: ' + p.variants : ''}`
          );
        }
      }
    }

    // Fallback: if nothing matched, include the full roster
    if (!allRosterLines.length) {
      for (const [setName, players] of Object.entries(playerRoster)) {
        for (const p of players) {
          allRosterLines.push(
            `${p.player} | base: ${p.basePop} | set: "${setName}"${p.variants ? ' | variants: ' + p.variants : ''}`
          );
        }
      }
    }
  }
  const rosterText = allRosterLines.length
    ? allRosterLines.join("\n")
    : "No roster loaded.";

  // SET CATALOG as fallback text
  const fallbackText = catalogText || "No set catalog.";

  const prompt = `You are a data cleaning assistant for ghostwrite, a collectible toy company.
Classify each eBay listing title below into Set, Player, and Variant.

━━━ PLAYER ROSTER (primary lookup — use this first) ━━━
Each line: Player Name | base population | set name | known variants
NOTE: All population numbers in this roster are the REAL print-run numbers that sellers print on the figure
and write in their eBay listings (e.g. /2000, /50, 1/1). These will match exactly what you see in titles.
${rosterText}

━━━ CLASSIFICATION METHOD ━━━
Use this order of priority:

1. ROSTER MATCH (primary): Extract the player name and population number from the title.
   Match the player name against the roster using fuzzy/semantic matching.
   Extract the population number from the title — it can appear in many formats:
   "/1000", "limited to 1000", "of 1000", "numbered to 1000", "1000 made", "#/1000", "out of 1000"
   Also recognize parenthesized slash formats: "(1/1)", "(5/10)", "(23/50)" — the SECOND number is the population.
   The NUMBER must match exactly (1000 ≠ 2000), but the format can be anything.
   ⚠ IMPORTANT: "100%" in a title is a brand authenticity indicator ("100% authentic"), NOT a population. Never extract 100 as a population from "100%".
   Match that number against the base pop (or variant pop) in the roster — THIS IS THE MOST IMPORTANT SIGNAL.
   The roster populations are real_population values — the exact numbers that appear on the physical figure
   and in eBay listings. A title saying "/2000" will match a roster entry showing "base: 2000" exactly.
   Do NOT let keywords in the title override the population number match.
   Example: "Paul Skenes limited to 1000" → /1000 → MLB 2025 Game Face Base
   Example: "Eminem Detroit Pistons NBA Figure numbered to 2000" → /2000 → Eminem Home (NOT NBA 2024-25 which is /1000)
   ⚠ MULTI-SET DISAMBIGUATION (critical): When a player appears in more than one set in the roster,
   use the population number to identify the CORRECT SET — do NOT pick a set based on title keywords
   first and then match variants within it.
   Procedure: scan EVERY set this player appears in, compare the title's population number against
   each set's base pop AND variant pops. The set where a number matches is the correct set.
   Example: Cooper Flagg appears in "2025-26 Ghostwrite NBA" (base: 2000) and "NBA 2024-25 Game Face" (base: 1000).
   Title says "/2000" → correct set is "2025-26 Ghostwrite NBA", variant = Base.
   Even if the title says "NBA" or "Game Face", /2000 wins over all keyword matching.
   Only fall back to keyword-based set matching if NO population number is present in the title.

2. VARIANT DETECTION: Once set and player are matched, identify the variant:
   STEP A — Check the player's variants list from the roster FIRST. It overrides all defaults below.
   The variants field lists exactly what exists for that player, e.g. "SP /50, Gold /10, Chrome /5, Fire /1"
   Match the population number from the title to the variant name in that list.
   ⚠ CRITICAL: If the population number from the title matches the BASE population ("base: N" in the roster),
   the variant is "Base" — even though "Base" does not appear in the variants list.
   Examples:
   - Roster shows base: 2000 + title has /2000 → variant is "Base" (population matches base, not any variant)
   - Roster shows "SP /50" + title has /50 → variant is "SP" (NOT Victory)
   - Roster shows "Victory /50" + title has /50 → variant is "Victory"
   - Roster shows "Vides /25" + title has /25 → variant is "Vides" (NOT Emerald)

   STEP B — If no roster variants list, use these defaults:
   - /1 or "numbered to 1" or "1/1" or "(1/1)" = Fire
   - /5 or "numbered to 5" = Chrome
   - /10 or "numbered to 10" or "gold" = Gold
   - /25 or "numbered to 25" = Emerald
   - /50 or "numbered to 50" = Victory
   - Matches base pop from roster, or no number = Base

   STEP C — Handle descriptive phrases:
   ⚠ CRITICAL: "SP", "SSP", "short print", "super short print", "case hit", "Plakata" in the title:
     → These are SELLER MARKETING TERMS, not variant names. Treat them as noise UNLESS:
       (a) The player's roster explicitly lists an SP variant, AND
       (b) A population number appears in the title that matches the SP population in the roster
     → If either condition is missing: completely IGNORE these terms and classify by population number only
     → Example: "Freddie Freeman SSP Plakata Case Hit /1000" with no SP in roster → Base (population /1000 = base pop)
     → Example: "Cooper Flagg Rookie /2000" → Base (population /2000 = base pop; "Rookie" is noise)
   - "gold short print", "gold /10", "gold numbered to 10" = Gold ONLY if /10 appears in title
   - Ignore always: "First Edition", "1st Edition", "RC", "Rookie", "Figure", "Figurine", "100%", "Plakata", "Case Hit" — not variant indicators
   - CRITICAL: Never guess a variant. If you cannot determine variant from population number or unambiguous explicit text, return "Base".

   STEP D — Emoji vs explicit text rule (CRITICAL):
   - EMOJIS (🔥, ⭐, ✨, 💎, 🏆, etc.) are PURELY DECORATIVE and must be completely ignored
   - A 🔥 emoji does NOT indicate the Fire variant
   - HOWEVER: the WORD "FIRE", "CHROME", "GOLD", "EMERALD", "VICTORY" as explicit text in the title IS a valid variant signal
   - "FIRE CHROME" or "FIRE chrome" in the title → output variant = "Fire" (Chrome is a surface finish, not a separate variant — never output "Fire Chrome", always just "Fire")
   - "GOLD CHROME" → output variant = "Gold"; "CHROME" alone → output variant = "Chrome"
   - Always cross-check explicit variant text against the population number — if both agree, use that variant; if they conflict, trust the population number
   - Strip all emojis mentally before classifying, but keep explicit variant words

3. FALLBACK (if no roster match): Use set keywords to detect set:
${fallbackText}

━━━ PLAYER MATCHING RULES ━━━
- Use fuzzy/semantic matching — "Acuna", "Ronald Acuna", "Ronald Acuña Jr." all match "Ronald Acuna Jr."
- Last name only is valid if unique in the roster
- Always output the EXACT canonical name from the roster (Title Case, plain ASCII, Jr. with period)
- If genuinely no match, return "Unknown" and set inferred: true
- LOGO / LOGOMAN RULE: If the title contains "logo", "logoman", or "logo man" (case-insensitive)
  and no other player name from the roster is present in the title, set player = "Trophy".
  These are special league-logo figures listed by sellers under various names.

━━━ EXCLUSION RULES ━━━
IMPORTANT: Strip all emojis from the title before applying any rules below.
Emojis are decorative only and should never influence classification or exclusion decisions.
Set exclude=true if (checked BEFORE any other classification):
- Title contains "break" (any form: "basketball break", "box break", "group break") → reason: "break listing" — HIGHEST PRIORITY
- Title ends with a bare number (1, 2, 3 etc.) with NO "/" anywhere → reason: "break listing"
- Title contains "400%" → reason: "400%"
- Title contains "spot" or "random slot" → reason: "break listing"
- Multi-pack: "(2)", "(3)", "lot of" → reason: "multi-pack"
- Clearly not a ghostwrite product → reason: "not ghostwrite"

━━━ AFA GRADING RULES ━━━
AFA is a third-party grading service. Detect it as follows:
- If the title contains "AFA", "A.F.A.", or "Afa" (case-insensitive) → the item is AFA graded
- Extract the numeric grade which appears after "AFA", "AFA Grade", or "Afa Grade":
  e.g. "AFA 9.0" → "AFA 9.0", "AFA Grade 9.5" → "AFA 9.5", "Afa Grade 9" → "AFA 9", "AFA Grade 9.0🔥" → "AFA 9.0"
- Always normalize to format "AFA [grade]" (e.g. "AFA 9.0", "AFA 8.5")
- If AFA is present but no grade found → return "AFA"
- "Population 10" or "Pop 10" refers to variant population, NOT the AFA grade — ignore it for grading
- AFA grading does NOT affect set, player, or variant detection — classify those normally

━━━ RESPONSE FORMAT ━━━
Respond ONLY with a JSON array, one object per title, same order as input:
[
  {
    "set": "Exact set name from roster",
    "player": "Exact canonical name from roster",
    "variant": "Exact variant name from roster (e.g. Base, Fire, Chrome, Gold, Emerald, Victory, Vides, Bet On Women, SP, or any other variant name from the roster)",
    "exclude": false,
    "reason": "",
    "inferred": false,
    "afa": "AFA 9.0 or empty string if not graded",
    "note": "One-line explanation of how variant was determined, e.g. 'Base — /2000 matched base pop' or 'Fire — FIRE CHROME text + (1/1) population' or 'SP — /50 matched SP in roster' or 'Base — no population found, defaulted'"
  }
]

Titles to classify:
${numberedTitles}`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 4096,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("API HTTP " + response.getResponseCode() + ": " + response.getContentText().substring(0, 200));
  }

  const data = JSON.parse(response.getContentText());
  if (data.stop_reason === "max_tokens") {
    Logger.log("Phase 2 Haiku: response truncated (hit max_tokens). Batch had %s titles.", titles.length);
    throw new Error("AI response truncated — max_tokens reached for " + titles.length + " titles");
  }
  const text = data.content[0].text.trim().replace(/```json|```/g, "").trim();

  try {
    return JSON.parse(text);
  } catch(e) {
    // Model sometimes appends explanation text after the JSON array.
    // Extract just the [...] array portion and retry.
    const arrayMatch = text.match(/\[\s*\{[\s\S]*\}\s*\]/);
    if (arrayMatch) {
      try {
        return JSON.parse(arrayMatch[0]);
      } catch(e2) {
        throw new Error("Could not parse AI response: " + text.substring(0, 200));
      }
    }
    throw new Error("Could not parse AI response: " + text.substring(0, 200));
  }
}


// ── Fix PRICE INTELLIGENCE range limits ──────────────────────────────────────
// Replaces hardcoded $2001 row limits with $5000 in all PI formulas
function fixPriceIntelligenceRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pi = ss.getSheetByName("PRICE INTELLIGENCE");
  if (!pi) { SpreadsheetApp.getUi().alert("PRICE INTELLIGENCE sheet not found."); return; }

  const lastRow = pi.getLastRow();
  const lastCol = pi.getLastColumn();
  if (lastRow < 4) { SpreadsheetApp.getUi().alert("No data in PRICE INTELLIGENCE."); return; }

  let fixed = 0;
  const data = pi.getRange(4, 1, lastRow - 3, lastCol).getFormulas();

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const formula = data[r][c];
      if (formula && formula.includes('$2001')) {
        data[r][c] = formula.replace(/\$2001/g, '$5000');
        fixed++;
      }
    }
  }

  pi.getRange(4, 1, lastRow - 3, lastCol).setFormulas(data);
  SpreadsheetApp.getUi().alert("Done! Fixed " + fixed + " formula(s) — all ranges now extend to row 5000.");
}


// ── Web app (for dashboard) ───────────────────────────────────────────────────
function doGet(e) {
  const output = getMarketData();
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Flag/unflag a listing ─────────────────────────────────────────────────────
function doPost(e) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Content-Type": "application/json",
  };
  try {
    const params = JSON.parse(e.postData.contents);
    const rowIndex = Number(params.rowIndex);
    const flagged  = params.flagged; // true or false

    if (!rowIndex || rowIndex < 4) throw new Error("Invalid row index");

    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName("RAW PASTE");
    if (!raw) throw new Error("RAW PASTE sheet not found");

    raw.getRange(rowIndex, 19).setValue(flagged ? "FLAG" : ""); // col S (1-based col 19 = S)

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, rowIndex, flagged }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getMarketData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName("RAW PASTE");
  const pi  = ss.getSheetByName("PRICE INTELLIGENCE");

  const rawLastRow = raw.getLastRow();
  const rawData = rawLastRow > 3
    ? raw.getRange(4, 1, rawLastRow - 3, 21).getValues()
    : [];

  const recentSales = [];
  const now = new Date();
  const thirtyDaysAgo = new Date(now - 30 * 24 * 60 * 60 * 1000);
  let weeklySalesCount = 0;
  const oneWeekAgo = new Date(now - 7 * 24 * 60 * 60 * 1000);

  const allListings = []; // all valid sales for price guide

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const title   = row[1];
    const price   = row[12];
    const player  = row[13];
    const variant = row[14];
    const set_    = row[15];
    const status  = row[16];
    const dupe    = String(row[11] || "").trim();
    const date    = row[9];
    const image   = row[0];
    const url     = row[2];
    const flagged = String(row[18] || "").trim();
    const source  = String(row[19] || "eBay").trim(); // col T
    const afa     = String(row[20] || "").trim();     // col U
    const rowIndex = i + 4; // actual sheet row number (data starts at row 4)

    if (!title || status !== "valid" || dupe === "dupe") continue;

    const saleDate = date ? new Date(date) : null;
    if (saleDate && saleDate >= oneWeekAgo) weeklySalesCount++;

    const listing = {
      rowIndex,
      title:   String(title).substring(0, 100),
      price:   Number(price),
      player:  String(player  || "Unknown"),
      variant: String(variant || "Base"),
      set:     String(set_    || ""),
      date:    saleDate ? saleDate.toISOString() : null,
      url:     String(url   || ""),
      image:   String(image || ""),
      flagged: flagged === "FLAG",
      source:  source,
      afa:     afa,
    };

    allListings.push(listing);

    if (price > 0 && saleDate && saleDate >= thirtyDaysAgo) {
      recentSales.push({...listing, title: listing.title.substring(0, 80)});
    }
  }

  recentSales.sort((a, b) => b.price - a.price);

  const piLastRow = pi.getLastRow();
  const piData = piLastRow > 3
    ? pi.getRange(4, 1, piLastRow - 3, 16).getValues()
    : [];

  const products = [];
  const variantTotals = {};
  const setVariantTotals = {};

  for (const row of piData) {
    const set_    = row[0]; const player  = row[1]; const variant = row[2];
    const edition = row[3]; const retail  = row[4]; const nSales  = row[5];
    const avg     = row[6]; const low     = row[7]; const high    = row[8];
    const mult    = row[9]; const d30n    = row[10]; const d30avg = row[11];
    const d90n    = row[12]; const d90avg = row[13]; const lastDate= row[14];
    const lastPx  = row[15];

    if (!set_ || !player) continue;

    const product = {
      set: String(set_), player: String(player), variant: String(variant || "Base"),
      edition: String(edition || ""), retail: Number(retail) || 0,
      totalSales: Number(nSales) || 0, avgPrice: Number(avg) || 0,
      lowPrice: Number(low) || 0, highPrice: Number(high) || 0,
      multiplier: Number(mult) || 0, sales30d: Number(d30n) || 0,
      avg30d: Number(d30avg) || 0, sales90d: Number(d90n) || 0,
      avg90d: Number(d90avg) || 0,
      lastSaleDate: lastDate ? new Date(lastDate).toISOString() : null,
      lastSalePrice: Number(lastPx) || 0,
    };

    products.push(product);

    if (product.avgPrice > 0 && product.totalSales > 0) {
      const vKey = product.variant;
      if (!variantTotals[vKey]) variantTotals[vKey] = { sum: 0, salesSum: 0 };
      variantTotals[vKey].sum += product.avgPrice * product.totalSales;
      variantTotals[vKey].salesSum += product.totalSales;

      const svKey = product.set + "|||" + product.variant;
      if (!setVariantTotals[svKey]) setVariantTotals[svKey] = { set: product.set, variant: product.variant, sum: 0, salesSum: 0 };
      setVariantTotals[svKey].sum += product.avgPrice * product.totalSales;
      setVariantTotals[svKey].salesSum += product.totalSales;
    }
  }

  const variantAverages = Object.entries(variantTotals)
    .map(([variant, d]) => ({ variant, avgPrice: d.salesSum > 0 ? Math.round(d.sum / d.salesSum) : 0, totalSales: d.salesSum }))
    .sort((a, b) => b.avgPrice - a.avgPrice);

  const setVariantAverages = Object.values(setVariantTotals)
    .map(d => ({ set: d.set, variant: d.variant, avgPrice: d.salesSum > 0 ? Math.round(d.sum / d.salesSum) : 0, totalSales: d.salesSum }))
    .sort((a, b) => b.avgPrice - a.avgPrice);

  const sets    = [...new Set(products.map(p => p.set))].filter(Boolean).sort();
  const players = [...new Set(products.map(p => p.player))].filter(Boolean).sort();

  // Load AI model output from script properties (set by Layer 2)
  let aiModelOutput = null;
  try {
    const stored = PropertiesService.getScriptProperties().getProperty("AI_MODEL_OUTPUT");
    if (stored) aiModelOutput = JSON.parse(stored);
  } catch(e) {}

  // Build set-level variant multipliers from MODEL INPUTS tab (Layer 1)
  const modelMultipliers = [];
  try {
    const modelSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MODEL INPUTS");
    if (modelSheet && modelSheet.getLastRow() > 1) {
      const mData = modelSheet.getRange(2, 1, modelSheet.getLastRow() - 1, 9).getValues();
      for (const row of mData) {
        if (!row[0] || !row[1]) continue;
        modelMultipliers.push({
          set: String(row[0]), variant: String(row[1]),
          avgMult: Number(row[2]) || 0, weightedMult: Number(row[3]) || 0,
          sampleSize: Number(row[4]) || 0, confidence: String(row[5] || ""),
          minMult: Number(row[6]) || 0, maxMult: Number(row[7]) || 0,
          stdDev: Number(row[8]) || 0,
        });
      }
    }
  } catch(e) {}

  return {
    fetchedAt: new Date().toISOString(),
    weeklySalesCount,
    totalProducts: products.length,
    products,
    recentSales,
    variantAverages,
    setVariantAverages,
    sets,
    players,
    allListings,
    modelMultipliers,
    aiModelOutput,
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// eBay OAuth + Browse API helpers
// ═══════════════════════════════════════════════════════════════════════════════

const EBAY_APP_ID    = "JesseEin-ghostwri-PRD-395507e11-01c7cb05";
const EBAY_CERT_ID   = "PRD-95507e11d482-59a0-4b4b-98bb-a219";
const EBAY_TOKEN_KEY = "EBAY_OAUTH_TOKEN";       // ScriptProperties cache key
const EBAY_TOKEN_EXP = "EBAY_OAUTH_TOKEN_EXP";  // expiry epoch ms

// ── OAuth: Client Credentials flow ───────────────────────────────────────────
function getEbayAccessToken_() {
  const props = PropertiesService.getScriptProperties();

  // Return cached token if still valid (with 60s buffer)
  const cached    = props.getProperty(EBAY_TOKEN_KEY);
  const cachedExp = parseInt(props.getProperty(EBAY_TOKEN_EXP) || "0", 10);
  if (cached && Date.now() < cachedExp - 60000) {
    Logger.log("Using cached eBay token (expires in %ss)", Math.round((cachedExp - Date.now()) / 1000));
    return cached;
  }

  const credentials = Utilities.base64Encode(`${EBAY_APP_ID}:${EBAY_CERT_ID}`);

  const res = UrlFetchApp.fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "post",
    headers: {
      "Authorization": `Basic ${credentials}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    payload: "grant_type=client_credentials&scope=https%3A%2F%2Fapi.ebay.com%2Foauth%2Fapi_scope",
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("eBay auth failed: " + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());
  const expMs = Date.now() + (data.expires_in * 1000);

  props.setProperties({ [EBAY_TOKEN_KEY]: data.access_token, [EBAY_TOKEN_EXP]: String(expMs) });
  Logger.log("✅ New eBay token acquired. Expires in %ss", data.expires_in);
  return data.access_token;
}

// ── Browse API: search item summaries ────────────────────────────────────────
function ebayBrowseSearch_(query, {limit=10, sort="newlyListed", filter=""} = {}) {
  const token  = getEbayAccessToken_();
  const params = {
    q:     query,
    limit: String(limit),
    sort,
    ...(filter ? { filter } : {}),
  };
  const qs  = Object.entries(params).map(([k,v]) => `${k}=${encodeURIComponent(v)}`).join("&");
  const url = `https://api.ebay.com/buy/browse/v1/item_summary/search?${qs}`;

  const res = UrlFetchApp.fetch(url, {
    headers: {
      "Authorization": `Bearer ${token}`,
      "X-EBAY-C-MARKETPLACE-ID": "EBAY-US",
    },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("eBay Browse search failed: " + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

// ── Browse API: full item detail ──────────────────────────────────────────────
function ebayBrowseItemDetail_(itemId) {
  const token = getEbayAccessToken_();
  const url   = `https://api.ebay.com/buy/browse/v1/item/${encodeURIComponent(itemId)}`;

  const res = UrlFetchApp.fetch(url, {
    headers: {
      "Authorization": `Bearer ${token}`,
      "X-EBAY-C-MARKETPLACE-ID": "EBAY-US",
    },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("eBay item detail failed: " + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

// ═══════════════════════════════════════════════════════════════════════════════
// eBay LIVE LISTINGS INGESTION ENGINE  v2 — trigger-loop pattern
// ─────────────────────────────────────────────────────────────────────────────
// Designed around Google Apps Script's 6-minute execution limit.
// Three separate phases, each resumable via trigger:
//
//  PHASE 1 — startLiveIngestion()
//    Fetches all live listings from eBay API (~1-2 min), writes raw unclassified
//    rows to LIVE LISTINGS tab, stores fetch metadata in ScriptProperties,
//    then kicks off Phase 2 via a 1-minute trigger.
//
//  PHASE 2 — classifyLiveListingsLoop()  [trigger-driven, repeats every minute]
//    Picks up where it left off (LIVE_CLASSIFY_RESUME_ROW), classifies 8 rows
//    at a time with Claude Haiku (same pattern as cleanNewRows), sleeps 15s
//    between batches, saves progress. When all rows are classified, deletes
//    trigger and kicks off Phase 3.
//
//  PHASE 3 — finalizeLiveIngestion()
//    Reads all classified rows from current scrape run, builds LIVE SNAPSHOTS
//    aggregate, updates PRICE INTELLIGENCE live columns Q-V, cleans up.
//
// Menu entry points:
//   startLiveIngestion()          — starts the full pipeline (Phase 1)
//   stopLiveIngestion()           — cancels any running pipeline
//   setupLiveIngestionTrigger()   — schedules automatic runs every 3 days
//   stopLiveIngestionTrigger()    — removes the schedule
// ═══════════════════════════════════════════════════════════════════════════════

// ── Constants ─────────────────────────────────────────────────────────────────
const LIVE_QUERY     = 'ghostwrite -novel -paperback -DVD -vhs -hardcover -"complete set" -sealed -tenyo -atari -"box break" -break -spot';
const LIVE_PAGE_SIZE = 200;  // max per Browse API call
const LIVE_MAX_PAGES = 6;    // up to 1200 listings total
const LIVE_CLASSIFY_BATCH = 100; // rows to classify per JS pass (no API cost)

// ScriptProperty keys
const PROP_LIVE_RUN_ID       = "LIVE_RUN_ID";        // ISO timestamp of current run
const PROP_LIVE_RESUME_ROW   = "LIVE_CLASSIFY_RESUME_ROW";
const PROP_LIVE_STATUS       = "LIVE_STATUS";         // human-readable status

// LIVE LISTINGS tab — column indices (1-based)
const LL = {
  SNAPSHOT_ID:    1,  // A
  ITEM_ID:        2,  // B
  LEGACY_ID:      3,  // C
  SCRAPED_AT:     4,  // D — ISO timestamp of the run that fetched this
  TITLE:          5,  // E
  ASK_PRICE:      6,  // F
  CURRENCY:       7,  // G
  LISTING_TYPE:   8,  // H — "Fixed Price", "Auction", "Auction+BIN"
  HAS_BEST_OFFER: 9,  // I
  BID_COUNT:      10, // J
  CONDITION:      11, // K
  LISTED_DATE:    12, // L
  DAYS_LISTED:    13, // M
  IMAGE_URL:      14, // N
  EBAY_URL:       15, // O
  SELLER:         16, // P
  SELLER_SCORE:   17, // Q
  PLAYER:         18, // R — filled by Phase 2
  VARIANT:        19, // S — filled by Phase 2
  SET_NAME:       20, // T — filled by Phase 2
  AFA:            21, // U — filled by Phase 2
  EXCLUDE:        22, // V — filled by Phase 2
  AMBIGUOUS:      24, // X — filled by Phase 2: variant or set ambiguity
  SET_MATH_ID:    25, // Y — filled by Phase 2: UUID from set_math table
  STATUS:         23, // W
};

// LIVE SNAPSHOTS tab — column indices (1-based)
const LS = {
  SNAPSHOT_DATE:  1,  // A
  SET_NAME:       2,  // B
  PLAYER:         3,  // C
  VARIANT:        4,  // D
  ACTIVE_COUNT:   5,  // E
  MEDIAN_ASK:     6,  // F
  LOW_ASK:        7,  // G
  HIGH_ASK:       8,  // H
  AVG_ASK:        9,  // I
  BO_COUNT:       10, // J
};

// PRICE INTELLIGENCE live columns (appended after col P = 16)
const PI_LIVE = {
  ACTIVE_LISTINGS: 17, // Q
  MEDIAN_ASK:      18, // R
  LOW_ASK:         19, // S
  ASK_TREND:       20, // T
  DAYS_ON_MARKET:  21, // U
  LAST_SCRAPED:    22, // V
};


// ═══════════════════════════════════════════════════════════════════════════════
// SHEET SETUP
// ═══════════════════════════════════════════════════════════════════════════════

function ensureLiveListingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("LIVE LISTINGS");
  if (!sheet) {
    sheet = ss.insertSheet("LIVE LISTINGS");
    const headers = [
      "Snapshot ID", "Item ID", "Legacy Item ID", "Scraped At",
      "Title", "Ask Price", "Currency", "Listing Type", "Has Best Offer",
      "Bid Count", "Condition", "Listed Date", "Days Listed", "Image URL",
      "eBay URL", "Seller", "Seller Score",
      "Player", "Variant", "Set", "AFA", "Exclude Reason", "Status", "Ambiguous", "Set Math ID"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureLiveSnapshotsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("LIVE SNAPSHOTS");
  if (!sheet) {
    sheet = ss.insertSheet("LIVE SNAPSHOTS");
    const headers = [
      "Snapshot Date", "Set", "Player", "Variant",
      "Active Listings", "Median Ask", "Low Ask", "High Ask", "Avg Ask",
      "Best Offer Count"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensurePiLiveColumns_() {
  const pi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRICE INTELLIGENCE");
  if (!pi) return;
  const headerRow = 3;
  const cols = [
    { col: PI_LIVE.ACTIVE_LISTINGS, label: "Active Listings" },
    { col: PI_LIVE.MEDIAN_ASK,      label: "Median Ask"      },
    { col: PI_LIVE.LOW_ASK,         label: "Low Ask"         },
    { col: PI_LIVE.ASK_TREND,       label: "Ask Trend"       },
    { col: PI_LIVE.DAYS_ON_MARKET,  label: "Days on Market"  },
    { col: PI_LIVE.LAST_SCRAPED,    label: "Last Scraped"    },
  ];
  for (const h of cols) {
    const cell = pi.getRange(headerRow, h.col);
    if (!cell.getValue()) { cell.setValue(h.label).setFontWeight("bold"); }
  }
}

function setLiveStatus_(msg) {
  PropertiesService.getScriptProperties().setProperty(PROP_LIVE_STATUS, msg);
  Logger.log("[LIVE] " + msg);
  // Also write to LIVE LISTINGS cell A1 area for visibility
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIVE LISTINGS");
    if (sheet) sheet.getRange(1, 22).setValue("Status: " + msg + " [" + new Date().toLocaleTimeString() + "]");
  } catch(e) {}
}


// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 1 — FETCH FROM EBAY + WRITE RAW ROWS
// ═══════════════════════════════════════════════════════════════════════════════

function startLiveIngestion() {
  const props = PropertiesService.getScriptProperties();

  // ── Ensure tabs exist ──────────────────────────────────────────────────────
  const llSheet = ensureLiveListingsSheet_();
  ensureLiveSnapshotsSheet_();
  ensurePiLiveColumns_();

  // ── Generate run ID ────────────────────────────────────────────────────────
  const runId = new Date().toISOString();
  props.setProperty(PROP_LIVE_RUN_ID, runId);
  props.deleteProperty(PROP_LIVE_RESUME_ROW);

  // ── Load catalogs for JS classifier (Phase 2) ────────────────────────────
  let setCatalog, playerRoster;
  try {
    setCatalog   = loadSetCatalogFromSupabase(); // uses Supabase sets table as canonical
    playerRoster = loadPlayerRosterFromSupabase(); // uses Supabase set_math as canonical
  } catch(e) {
    setLiveStatus_("❌ Error loading catalogs: " + e.message);
    SpreadsheetApp.getUi().alert("Could not load SET CATALOG or PLAYER ROSTER: " + e.message);
    return;
  }

  setLiveStatus_("Phase 1: fetching listings from eBay...");

  // ── Fetch all pages from eBay Browse API ───────────────────────────────────
  const allListings = [];
  let offset = 0;
  let totalAvailable = null;

  for (let page = 0; page < LIVE_MAX_PAGES; page++) {
    const qs = "q=" + encodeURIComponent(LIVE_QUERY) +
               "&limit=" + LIVE_PAGE_SIZE +
               "&offset=" + offset +
               "&sort=newlyListed";
    const url = "https://api.ebay.com/buy/browse/v1/item_summary/search?" + qs;

    const res = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Bearer " + getEbayAccessToken_(),
        "X-EBAY-C-MARKETPLACE-ID": "EBAY-US",
      },
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      Logger.log("eBay API error page %s: %s", page, res.getContentText().substring(0, 200));
      break;
    }

    const data = JSON.parse(res.getContentText());
    if (totalAvailable === null) totalAvailable = data.total || 0;
    const items = data.itemSummaries || [];
    if (!items.length) break;

    allListings.push(...items);
    offset += items.length;
    Logger.log("Page %s: %s items (total so far: %s / %s)", page + 1, items.length, allListings.length, totalAvailable);

    if (allListings.length >= totalAvailable) break;
    Utilities.sleep(500);
  }

  if (!allListings.length) {
    setLiveStatus_("❌ No listings returned from eBay");
    SpreadsheetApp.getUi().alert("No listings returned from eBay. Check your credentials.");
    return;
  }

  // ── Second pass: fetch pure auction listings (not returned by default) ────
  setLiveStatus_("Phase 1: fetching auction listings...");
  const auctionListings = [];
  offset = 0;
  totalAvailable = null;
  for (let page = 0; page < LIVE_MAX_PAGES; page++) {
    const qs = "q=" + encodeURIComponent(LIVE_QUERY) +
               "&limit=" + LIVE_PAGE_SIZE +
               "&offset=" + offset +
               "&sort=newlyListed" +
               "&filter=buyingOptions:%7BAUCTION%7D";
    const url = "https://api.ebay.com/buy/browse/v1/item_summary/search?" + qs;
    const res = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Bearer " + getEbayAccessToken_(),
        "X-EBAY-C-MARKETPLACE-ID": "EBAY-US",
      },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      Logger.log("eBay auction fetch error page %s: %s", page, res.getContentText().substring(0, 200));
      break;
    }
    const data = JSON.parse(res.getContentText());
    if (totalAvailable === null) totalAvailable = data.total || 0;
    const items = data.itemSummaries || [];
    if (!items.length) break;
    auctionListings.push(...items);
    offset += items.length;
    Logger.log("Auction page %s: %s items (total: %s / %s)", page + 1, items.length, auctionListings.length, totalAvailable);
    if (auctionListings.length >= totalAvailable) break;
    Utilities.sleep(500);
  }
  Logger.log("Auction listings fetched: %s", auctionListings.length);

  // Merge — dedupe by itemId
  const seenIds = new Set(allListings.map(i => i.itemId));
  for (const item of auctionListings) {
    if (!seenIds.has(item.itemId)) {
      allListings.push(item);
      seenIds.add(item.itemId);
    }
  }
  Logger.log("Total listings after merge: %s", allListings.length);

  // ── Refresh scraped_at for all currently-visible listings ────────────────────
  // Keeps is_active deactivation accurate: scraped_at advances each run for
  // listings still on eBay, and stalls for sold/ended ones → deactivated after
  // LISTING_STALE_DAYS_ days by markStaleListingsInactive_ in runPostIngestion.
  setLiveStatus_("Phase 1: refreshing listing timestamps...");
  refreshListingsScrapedAt_(allListings.map(item => item.itemId), runId);

  // ── Dedup: skip itemIds already in sheet from the last 3 days ───────────────
  const recentIds = new Set();
  const llLastRow = llSheet.getLastRow();
  if (llLastRow > 1) {
    const threeDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    
    const existing = llSheet.getRange(2, LL.ITEM_ID, llLastRow - 1, 3).getValues();
    for (const row of existing) {
      const itemId  = String(row[0] || "").trim();
      const scraped = row[2] ? new Date(row[2]) : null;
      if (itemId && scraped && scraped >= threeDaysAgo) recentIds.add(itemId);
    }
  }
  Logger.log("Dedup: %s item IDs in sheet from last 3 days", recentIds.size);

  // ── Log buying option breakdown for diagnostics ──────────────────────────
  const boCounts = {};
  allListings.forEach(item => {
    const key = (item.buyingOptions || ["UNKNOWN"]).join("+");
    boCounts[key] = (boCounts[key] || 0) + 1;
  });
  Logger.log("Buying options breakdown: %s", JSON.stringify(boCounts));

  const newListings = allListings.filter(item => !recentIds.has(item.itemId));
  Logger.log("New listings to write: %s (skipped %s dupes)", newListings.length, allListings.length - newListings.length);

  // ── Write raw unclassified rows to LIVE LISTINGS ───────────────────────────
  if (newListings.length) {
    const now = new Date(runId);
    const rows = newListings.map(item => {
      const listedDate = item.itemCreationDate ? new Date(item.itemCreationDate) : null;
      const daysListed = listedDate ? Math.round((now - listedDate) / 86400000) : "";
      const buyingOptions = item.buyingOptions || [];
      const isAuction  = buyingOptions.includes("AUCTION") && !buyingOptions.includes("FIXED_PRICE");
      const priceObj   = isAuction
        ? (item.currentBidPrice || item.price || {})
        : (item.price || {});
      const askPrice   = parseFloat(priceObj.value || "0") || 0;
      const hasBo      = buyingOptions.includes("BEST_OFFER");
      const listingType = isAuction
        ? (buyingOptions.includes("FIXED_PRICE") ? "Auction+BIN" : "Auction")
        : "Fixed Price";
      const bidCount = isAuction ? (item.bidCount || 0) : "";
      return [
        item.itemId + "_" + runId.substring(0, 10), // A snapshot ID
        item.itemId,                                  // B
        item.legacyItemId || "",                      // C
        runId,                                        // D scraped at
        (item.title || "").substring(0, 200),         // E
        askPrice,                                     // F
        (item.price || {}).currency || "USD",         // G
        listingType,                                  // H
        hasBo,                                        // I
        bidCount,                                     // J
        item.condition || "",                         // K
        listedDate ? listedDate.toISOString() : "",   // L
        daysListed,                                   // M
        ((item.image || {}).imageUrl) || "",          // N
        item.itemWebUrl || "",                        // O
        ((item.seller || {}).username) || "",         // P
        ((item.seller || {}).feedbackScore) || "",    // Q
        "", "", "", "", "",                           // R-V: classified in Phase 2
        isAuction ? "auction" : "active",             // W
      ];
    });

    const startRow = llSheet.getLastRow() + 1;
    llSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log("Wrote %s raw rows starting at row %s", rows.length, startRow);

  } else {
    setLiveStatus_("No new listings — running Phase 3 aggregate...");
    finalizeLiveIngestion_();
    return;
  }

  // ── Phase 2: classify all new rows via Claude Haiku ───────────────────────
  setLiveStatus_("Phase 2: classifying " + newListings.length + " listings...");
  const setMathCache = buildSetMathCache_();
  classifyLiveListings_(llSheet, runId, setCatalog, playerRoster, setMathCache);

  // ── Phase 3: aggregate + update PI, then push to Supabase ─────────────────
  finalizeLiveIngestion_();
}


// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 2 (DEPRECATED) — PURE JS CLASSIFIER  (no AI API calls, zero cost)
// ─────────────────────────────────────────────────────────────────────────────
// @deprecated  Renamed to classifyLiveListings_JS_ on 2026-03-24.
//              The active Phase 2 classifier is now classifyLiveListings_()
//              which uses Claude Haiku for the same quality as the sales
//              pipeline.  This JS version is kept intact in case volume ever
//              makes Haiku costs prohibitive (currently ~12 new listings/day →
//              negligible cost).  To reactivate: swap the call in
//              startLiveIngestion from classifyLiveListings_ back to
//              classifyLiveListings_JS_.
// ═══════════════════════════════════════════════════════════════════════════════

/*
function classifyLiveListings_JS_(llSheet, runId, setCatalog, playerRoster, setMathCache) {
  const lastRow = llSheet.getLastRow();
  if (lastRow < 2) return;

  // Load all rows for this run
  const data = llSheet.getRange(2, 1, lastRow - 1, LL.STATUS).getValues();
  const toClassify = [];
  for (let i = 0; i < data.length; i++) {
    const scrapedAt = String(data[i][LL.SCRAPED_AT - 1] || "").trim();
    const player    = String(data[i][LL.PLAYER - 1]     || "").trim();
    const title     = String(data[i][LL.TITLE - 1]      || "").trim();
    if (scrapedAt === runId && title && !player) {
      toClassify.push({ row: i + 2, title });
    }
  }

  if (!toClassify.length) return;
  Logger.log("Phase 2 JS: classifying %s rows", toClassify.length);

  // ── Build player lookup index ──────────────────────────────────────────────
  // Index by EVERY name token (not just last name) so hyphenated names,
  // first-name-only titles, and multi-word names all resolve correctly.
  const playerIndex = {};

  function addToIndex(token, entry) {
    const key = token.toLowerCase().replace(/[^a-z]/g, "");
    if (key.length < 2) return; // allow 2-char tokens (e.g. "ja" from "A'Ja")
    if (!playerIndex[key]) playerIndex[key] = [];
    playerIndex[key].push(entry);
  }

  for (const [setName, players] of Object.entries(playerRoster)) {
    for (const p of players) {
      const variantMap = {};
      const varStr = p.variants || "Victory /50, Emerald /25, Gold /10, Chrome /5, Fire /1";
      varStr.split(",").forEach(part => {
        const m = part.match(/([A-Za-z ]+?)\s*\/\s*(\d+)/);
        if (m) variantMap[parseInt(m[2], 10)] = m[1].trim();
      });

      const entry = {
        player:   p.player,
        setName,
        basePop:  parseInt(String(p.basePop).replace(/[^0-9]/g, ""), 10) || 0,
        variantMap,
      };

      // Index by every part of the name, including hyphen-split parts
      // Also index the full stripped name as one token (catches A'Ja → "aja")
      p.player.split(/[\s\-]+/).forEach(part => addToIndex(part, entry));
      // Full stripped name (e.g. "ajawilson") — catches cases where apostrophe
      // fragments the name into sub-3-char pieces
      const fullStripped = p.player.toLowerCase().replace(/[^a-z]/g, "");
      if (fullStripped.length >= 3) addToIndex(fullStripped, entry);
      // Last segment of full name
      const lastSegment = p.player.split(" ").pop();
      addToIndex(lastSegment, entry);
    }
  }

  const DEFAULT_POP_VARIANT = { 1: "Fire", 5: "Chrome", 10: "Gold", 25: "Emerald", 50: "Victory" };

  // Variant keywords — if these appear explicitly in the title, use as a signal
  // Variant keywords — all use word boundaries to avoid false matches
  // e.g. "gold" must not match "golden" (Golden State Warriors)
  // SP/SSP are intentionally excluded here — handled separately with variantMap guard
  const VARIANT_KEYWORDS = {
    "\\bfire\\b":    "Fire",
    "\\bchrome\\b":  "Chrome",
    "\\bgold\\b":    "Gold",
    "\\bemerald\\b": "Emerald",
    "\\bvictory\\b": "Victory",
    "\\bvides\\b":   "Vides",
    "bet on women":     "Bet On Women",
  };

  // Set year signals — used to disambiguate sets with same league
  // Maps year tokens found in title to preferred set name fragments
  const SET_YEAR_SIGNALS = [
    { pattern: /202[56]\/26|2025.26|2026/i,  fragment: "2025-26" },
    { pattern: /202[45]\/25|2024.25/i,        fragment: "2024-25" },
    { pattern: /\b2025\b/,                    fragment: "2025"    },
    { pattern: /\b2024\b/,                    fragment: "2024"    },
  ];

  // Build set keyword scorer with year-awareness
  const setKeywords = setCatalog.map(s => ({
    setName: s.setName,
    words:   (s.keywords || "").toLowerCase().split(/[\s,]+/).filter(w => w.length > 2),
    nameLow: s.setName.toLowerCase(),
  }));

  // ── Fix 1: extractPop — add POP: format ──────────────────────────────────
  function extractPop(title) {
    // Strip year-like patterns (YYYY/YY or YYYY-YY) before parsing
    // so "2025/26" doesn't get matched as a population slash
    const t = title
      .replace(/,/g, "")
      .replace(/\b20\d{2}[\/\-]\d{2,4}\b/g, " ") // remove e.g. 2025/26, 2024-25
      .replace(/#\d{1,2}\b/g, " ");               // remove jersey numbers e.g. #32, #7

    const patterns = [
      /\bpop\s*:\s*(\d+)/i,       // POP: 50, POP:50
      /\bpop\s+(\d+)/i,            // POP 50
      /[#\/](\d+)/,                // /50, #50  (year patterns already stripped above)
      /numbered\s+to\s+(\d+)/i,
      /limited\s+to\s+(\d+)/i,
      /out\s+of\s+(\d+)/i,
      /of\s+(\d+)/i,
      /(\d+)\s+made/i,
      /(\d+)\s+copies/i,
    ];
    for (const re of patterns) {
      const m = t.match(re);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > 0 && n <= 5000) return n;
      }
    }
    return null;
  }

  function isExcluded(title) {
    const t = title.toLowerCase();
    if (/\bbreak\b/.test(t))                            return "break listing";
    if (/\bspot\b/.test(t))                             return "break listing";
    if (/\brandom\s+slot\b/.test(t))                   return "break listing";
    if (/400%/.test(t))                                  return "400%";
    // Multi-pack: (3x), (10x), (3), lot of N, complete/full set
    if (/\(\s*\d+\s*x?\s*\)/.test(t))               return "multi-pack";
    if (/\blot\s+of\s+\d/.test(t))                    return "multi-pack";
    if (/full\s+set|complete\s+set/.test(t))            return "multi-pack";
    if (/\bnovel\b|\bpaperback\b|\bdvd\b/.test(t))   return "not ghostwrite";
    return null;
  }

  function detectAfa(title) {
    const m = title.match(/\bA\.?F\.?A\.?\s*(?:Grade\s*)?(\d+(?:\.\d+)?)/i);
    return m ? "AFA " + m[1] : "";
  }

  // ── Fix 2: detectVariantKeyword — read explicit variant from title text ───
  function detectVariantKeyword(title) {
    const t = title.toLowerCase();
    for (const [kw, variant] of Object.entries(VARIANT_KEYWORDS)) {
      const re = new RegExp(kw, 'i');
      if (re.test(t)) return variant;
    }
    return null;
  }

  // ── Fix 3: detectSetFromTitle — year-aware set scoring ───────────────────
  function detectSetFromTitle(tLow) {
    // Find which year signal matches the title
    let yearFragment = null;
    for (const sig of SET_YEAR_SIGNALS) {
      if (sig.pattern.test(tLow)) { yearFragment = sig.fragment; break; }
    }

    let bestSet = "", bestScore = 0;
    for (const s of setKeywords) {
      let score = s.words.filter(w => tLow.includes(w)).length;
      // Boost sets whose name contains the matched year fragment
      if (yearFragment && s.nameLow.includes(yearFragment.toLowerCase())) score += 3;
      // Boost by league signal
      if (/\bwnba\b/.test(tLow) && s.nameLow.includes("wnba")) score += 2;
      else if (/\bnba\b/.test(tLow) && s.nameLow.includes("nba") && !s.nameLow.includes("wnba")) score += 1;
      else if (/\bmlb\b/.test(tLow) && s.nameLow.includes("mlb")) score += 2;
      if (score > bestScore) { bestScore = score; bestSet = s.setName; }
    }
    return bestScore > 0 ? bestSet : "";
  }

  function classifyTitle(title) {
    const excl = isExcluded(title);
    if (excl) return { player: "", variant: "", set: "", afa: "", exclude: true, reason: excl };

    const afa        = detectAfa(title);
    const tLow       = title.toLowerCase();
    const pop        = extractPop(title);
    const varKw      = detectVariantKeyword(title); // explicit variant word in title

    const NOISE = new Set(["ghostwrite","figure","figurine","ghost","write","nba","wnba","mlb",
                           "game","face","series","limited","numbered","edition","rookie",
                           "print","afa","grade","graded","sealed","case","hit","ssp","super",
                           "first","edition","parallel",
                           // NBA team names
                           "hornets","lakers","nuggets","thunder","mavericks","celtics",
                           "warriors","suns","bucks","heat","nets","knicks","clippers",
                           "rockets","spurs","raptors","pelicans","grizzlies","jazz",
                           "blazers","kings","wolves","magic","wizards","pistons","bulls",
                           "cavaliers","pacers","hawks","sixers",
                           // MLB team names  
                           "braves","yankees","dodgers","astros","redsox","cubs","cardinals",
                           "giants","mets","phillies","padres","brewers","reds","pirates",
                           "marlins","nationals","rockies","diamondbacks","athletics","rays",
                           "rangers","angels","mariners","royals","tigers","white","sox",
                           "guardians","twins","orioles","bluejays",
                           // WNBA team names
                           "sky","aces","sun","sparks","storm","mystics","wings","mercury",
                           "liberty","fever","lynx","dream","valkyries",
                           // Common filler words in titles
                           "qty","avail","available","lot","new","release","season","the",
                           "las","vegas","los","angeles","san","antonio","connecticut",
                           "chicago","york","seattle","dallas","phoenix","indiana",
                           "washington","atlanta","minnesota","star","base","rookie",
                           "famous","fan","legend","collectible","box","and","with",
                           "limited","edition","figurine","mini","single","serial"]);
    const tokens = tLow.split(/\s+/).map(w => w.replace(/[^a-z]/g, "")).filter(w => w.length >= 2 && !NOISE.has(w));

    // ── Player matching ─────────────────────────────────────────────────────
    // Collect ALL candidates from all tokens, score each, pick best
    const seen = new Set();
    const scored = [];

    for (const token of tokens) {
      const candidates = playerIndex[token] || [];
      for (const c of candidates) {
        const uid = c.player + "|||" + c.setName;
        if (seen.has(uid)) continue;
        seen.add(uid);

        // Pop filter — skip if pop doesn't match any known pop for this player
        if (pop !== null) {
          const variantPops = new Set(Object.keys(c.variantMap).map(Number));
          // Hard-filter on pop only when it's a parallel number (<1000)
          // and doesn't match any known pop for this player (variant OR base).
          // Must include basePop in the check — WNBA base is /800 which is <1000
          // but is a valid base pop, not a parallel.
          // Numbers >=1000: skip filter entirely (absolute pop problem).
          if (pop < 1000 && !variantPops.has(pop) && pop !== c.basePop) continue;
        }

        // Name score — how many name parts appear in the title
        const nameParts = c.player.toLowerCase().replace(/[.']/g, "").split(/[\s\-]+/);
        const nameScore = nameParts.filter(p => tLow.replace(/[^a-z0-9 ]/g, " ").includes(p)).length;

        // Set score — does the title contain year/league signals for this set?
        const detectedSet = detectSetFromTitle(tLow);
        const setScore = detectedSet === c.setName ? 2 : 0;

        scored.push({ ...c, nameScore, setScore, totalScore: nameScore + setScore });
      }
    }

    // Sort: most name parts matched + set bonus first
    scored.sort((a, b) => b.totalScore - a.totalScore);
    const bestMatch = scored[0] || null;

    if (bestMatch) {
      let variant = "Base";
      let ambiguous = "";

     // Extract explicit variant keyword from title first (word-boundary — won't match "golden")
// SP excluded here; Step 3 handles it separately with its own variantMap guard.
const VARIANT_KW_RE = /\b(base|victory|emerald|gold|chrome|fire|vides)\b/i;
const kwMatch = title.match(VARIANT_KW_RE);
const varKw = kwMatch ? kwMatch[1].charAt(0).toUpperCase() + kwMatch[1].slice(1).toLowerCase() : null;

// ── Step 1: derive variant from pop number (highest confidence signal) ───
let popDerivedVariant = null;
if (pop !== null) {
  if (bestMatch.variantMap[pop])       popDerivedVariant = bestMatch.variantMap[pop];
  else if (pop >= 1000)                popDerivedVariant = "Base";
  else if (pop === bestMatch.basePop)  popDerivedVariant = "Base";
  else if (DEFAULT_POP_VARIANT[pop])   popDerivedVariant = DEFAULT_POP_VARIANT[pop];
  else                                 popDerivedVariant = "Base";
  variant = popDerivedVariant;
}

// ── Step 2: keyword override when pop is ambiguous ───────────────────────
if (varKw && varKw !== "Base") {
  const kwIsValid = bestMatch.validVariants
    ? bestMatch.validVariants.has(varKw.toLowerCase())
    : Object.values(bestMatch.variantMap).some(v => v.toLowerCase() === varKw.toLowerCase());
  const popIsAmbiguous = pop !== null && bestMatch.ambiguousPops && bestMatch.ambiguousPops.has(pop);
  if (kwIsValid && (popDerivedVariant === null || popIsAmbiguous)) {
    variant = varKw;
  } else if (popDerivedVariant === null && !kwIsValid) {
    variant = "";
  }
}

// ── Step 3: SP/SSP — only when no pop AND player has SP in set_math ────
if (popDerivedVariant === null && variant !== "SP") {
  if (/\bsp\b|\bssp\b|short\s*print/i.test(title)) {
    const spEntry = Object.entries(bestMatch.variantMap).find(([, v]) => v === "SP");
    if (spEntry) variant = "SP";
  }
}

      // ── Ambiguity detection ──────────────────────────────────────────────
      const ambiguousReasons = [];

      // Case A: variant ambiguous — SAME player in SAME set has multiple variants at this pop.
      // Only flag if the ambiguity is within the winning set, not across sets.
      // e.g. NBA 2025-26 Nikola Jokic has both Vides /25 and Emerald /25 → truly ambiguous.
      // But Trophy Chrome /5 existing in 5 different sets is NOT ambiguous if the set is known.
      if (pop !== null && pop < 1000) {
        const sameSetSamePlayerAtPop = scored
          .filter(c => c.player === bestMatch.player && c.setName === bestMatch.setName && c.variantMap[pop])
          .map(c => c.variantMap[pop]);
        const uniqueVariantsInWinningSet = [...new Set(sameSetSamePlayerAtPop)];
        // Truly ambiguous only if the same player in the same set maps to different variants at this pop
        // Use case-insensitive comparison for varKw (detectVariantKeyword returns lowercase)
        const varKwLow = (varKw || "").toLowerCase();
        const variantLow = (variant || "").toLowerCase();
        if (uniqueVariantsInWinningSet.length > 1 ||
            (varKwLow && varKwLow !== variantLow &&
             !uniqueVariantsInWinningSet.some(v => v.toLowerCase() === varKwLow))) {
          ambiguousReasons.push("variant: " + uniqueVariantsInWinningSet.join(" or "));
        }
      }

      // Case B: set ambiguous — SAME player exists in multiple sets AND both score equally.
      // Only flag when the set score difference is zero (no year/league signal broke the tie).
      if (scored.length > 1 &&
          scored[1].totalScore === bestMatch.totalScore &&
          scored[1].setName !== bestMatch.setName &&
          scored[1].player === bestMatch.player) {
        ambiguousReasons.push("set: " + bestMatch.setName + " or " + scored[1].setName);
      }

      if (ambiguousReasons.length) ambiguous = ambiguousReasons.join("; ");

      return { player: bestMatch.player, variant, set: bestMatch.setName,
               afa, exclude: false, reason: "", ambiguous };
    }

    // ── No player match — try set detection only ────────────────────────────
    const detectedSet = detectSetFromTitle(tLow);
    if (detectedSet) {
      return { player: "Unknown", variant: "Base", set: detectedSet, afa, exclude: false, reason: "" };
    }

    return { player: "Unknown", variant: "", set: "", afa, exclude: true, reason: "no match" };
  }

  // ── Classify all rows and write back in batches ─
  const WRITE_BATCH = 50;
  for (let i = 0; i < toClassify.length; i += WRITE_BATCH) {
    const chunk = toClassify.slice(i, i + WRITE_BATCH);
    const writeData = chunk.map(r => {
      const res = classifyTitle(r.title);
      const smKey = (res.set || "") + "|||" + (res.player || "") + "|||" + (res.variant || "");
      const setMathId = (!res.exclude && setMathCache) ? (setMathCache[smKey] || "") : "";
      return [
        toTitleCase(res.player || ""),
        res.variant   || "",
        res.set       || "",
        res.afa       || "",
        res.exclude ? (res.reason || "exclude") : "",
        "",              // STATUS (col W) — skip, not written here
        res.ambiguous || "",  // AMBIGUOUS (col X)
        setMathId,       // SET_MATH_ID (col Y)
      ];
    });
    // Write cols R–V (player/variant/set/afa/exclude) then skip W, write X and Y
    llSheet.getRange(chunk[0].row, LL.PLAYER, chunk.length, 5).setValues(
      writeData.map(r => r.slice(0, 5))
    );
    llSheet.getRange(chunk[0].row, LL.AMBIGUOUS, chunk.length, 1).setValues(
      writeData.map(r => [r[6]])
    );
    llSheet.getRange(chunk[0].row, LL.SET_MATH_ID, chunk.length, 1).setValues(
      writeData.map(r => [r[7]])
    );
  }

  Logger.log("Phase 2 JS: classified %s rows", toClassify.length);
}
*/  // END classifyLiveListings_JS_ (deprecated — see comment block above)

// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 2 — HAIKU CLASSIFIER  (same pipeline as completed-sales classifier)
// ─────────────────────────────────────────────────────────────────────────────
// Uses classifyBatch() (Claude Haiku) for all new listings in this run.
// Identical output columns as the old JS classifier: player, variant, set, afa,
// exclude, ambiguous, set_math_id.  At ~12 new listings/day the API cost is
// negligible; quality matches the sales pipeline.
// ═══════════════════════════════════════════════════════════════════════════════

function classifyLiveListings_(llSheet, runId, setCatalog, playerRoster, setMathCache) {
  const lastRow = llSheet.getLastRow();
  if (lastRow < 2) return;

  // ── Identify rows for this run that still need classification ──────────────
  const data = llSheet.getRange(2, 1, lastRow - 1, LL.STATUS).getValues();
  const toClassify = [];
  for (let i = 0; i < data.length; i++) {
    const scrapedAt = String(data[i][LL.SCRAPED_AT - 1] || "").trim();
    const player    = String(data[i][LL.PLAYER - 1]     || "").trim();
    const title     = String(data[i][LL.TITLE - 1]      || "").trim();
    if (scrapedAt === runId && title && !player) {
      toClassify.push({ row: i + 2, title });
    }
  }

  if (!toClassify.length) {
    Logger.log("Phase 2 Haiku: no new listings to classify this run.");
    return;
  }
  Logger.log("Phase 2 Haiku: classifying %s new listings via Claude Haiku...", toClassify.length);

  // ── Classify in batches (classifyBatch handles batching internally) ────────
  // Send all titles at once — at ~12 listings/day this is a single API call.
  const titles = toClassify.map(r => r.title);
  let results;
  try {
    results = classifyBatch(titles, setCatalog, playerRoster);
  } catch (e) {
    Logger.log("Phase 2 Haiku: classifyBatch error — %s. Skipping classification.", e.message);
    return;
  }

  if (!results || results.length !== toClassify.length) {
    Logger.log("Phase 2 Haiku: unexpected result count (%s vs %s) — skipping.", (results || []).length, toClassify.length);
    return;
  }

  // ── Write results back — same columns as JS classifier ────────────────────
  const WRITE_BATCH = 50;
  for (let i = 0; i < toClassify.length; i += WRITE_BATCH) {
    const chunk = toClassify.slice(i, i + WRITE_BATCH);
    const writeData = chunk.map((r, idx) => {
      const res = results[i + idx] || {};
      const smKey = (res.set || "") + "|||" + (res.player || "") + "|||" + (res.variant || "");
      const setMathId = (!res.exclude && setMathCache) ? (setMathCache[smKey] || "") : "";
      return [
        toTitleCase(res.player || ""),
        res.variant  || "",
        res.set      || "",
        res.afa      || "",
        res.exclude ? (res.reason || "exclude") : "",
        "",           // STATUS (col W) — not written here
        "",           // AMBIGUOUS (col X) — Haiku prompt doesn't return ambiguity signal; leave blank
        setMathId,    // SET_MATH_ID (col Y)
      ];
    });
    // Write cols R–V (player/variant/set/afa/exclude)
    llSheet.getRange(chunk[0].row, LL.PLAYER, chunk.length, 5).setValues(
      writeData.map(r => r.slice(0, 5))
    );
    // Write col X (AMBIGUOUS) and col Y (SET_MATH_ID)
    llSheet.getRange(chunk[0].row, LL.AMBIGUOUS, chunk.length, 1).setValues(
      writeData.map(r => [r[6]])
    );
    llSheet.getRange(chunk[0].row, LL.SET_MATH_ID, chunk.length, 1).setValues(
      writeData.map(r => [r[7]])
    );
  }

  Logger.log("Phase 2 Haiku: classified %s listings.", toClassify.length);
}

// classifyLiveListingsLoop kept as a no-op stub so any lingering triggers don't error
function classifyLiveListingsLoop() {
  Logger.log("classifyLiveListingsLoop called — now a no-op (classification is synchronous)");
  stopClassifyTrigger_();
}

function stopClassifyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "classifyLiveListingsLoop") ScriptApp.deleteTrigger(t);
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 3 — AGGREGATE + UPDATE PI
// ═══════════════════════════════════════════════════════════════════════════════

function finalizeLiveIngestion_() {
  const props   = PropertiesService.getScriptProperties();
  const runId   = props.getProperty(PROP_LIVE_RUN_ID);
  const llSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIVE LISTINGS");
  const lsSheet = ensureLiveSnapshotsSheet_();

  if (!llSheet || !runId) return;

  const lastRow = llSheet.getLastRow();
  if (lastRow < 2) return;

  const data = llSheet.getRange(2, 1, lastRow - 1, LL.STATUS).getValues();
  const thisRun = data.filter(row => String(row[LL.SCRAPED_AT - 1] || "").trim() === runId);

  Logger.log("Phase 3: aggregating %s rows from run %s", thisRun.length, runId);

  const groups = {};
  const snapshotDate = runId.substring(0, 10);
  const now = new Date(runId);

  for (const row of thisRun) {
    const exclude = String(row[LL.EXCLUDE - 1] || "").trim();
    if (exclude) continue;

    const player  = String(row[LL.PLAYER - 1]   || "").trim();
    const variant = String(row[LL.VARIANT - 1]   || "").trim() || "Base";
    const set_    = String(row[LL.SET_NAME - 1]  || "").trim();
    const price   = Number(row[LL.ASK_PRICE - 1])|| 0;
    const hasBo   = row[LL.HAS_BEST_OFFER - 1];
    const listed  = row[LL.LISTED_DATE - 1] ? new Date(row[LL.LISTED_DATE - 1]) : null;

    if (!set_ || !player || player === "Unknown") continue;

    const key = set_ + "|||" + player + "|||" + variant;
    if (!groups[key]) groups[key] = { set: set_, player, variant, prices: [], boCount: 0, oldestListed: null };

    if (price > 0) groups[key].prices.push(price);
    if (hasBo) groups[key].boCount++;
    if (listed && (!groups[key].oldestListed || listed < groups[key].oldestListed)) {
      groups[key].oldestListed = listed;
    }
  }

  // ── Write LIVE SNAPSHOTS ───────────────────────────────────────────────────
  const snapRows = [];
  for (const g of Object.values(groups)) {
    const prices = g.prices.slice().sort((a, b) => a - b);
    const n = prices.length;
    if (!n) continue;
    const median = n % 2 === 0 ? (prices[n/2-1] + prices[n/2]) / 2 : prices[Math.floor(n/2)];
    snapRows.push([
      snapshotDate, g.set, g.player, g.variant,
      n,
      Math.round(median),
      Math.round(prices[0]),
      Math.round(prices[n-1]),
      Math.round(prices.reduce((s,p)=>s+p,0)/n),
      g.boCount,
    ]);
  }
  if (snapRows.length) {
    const sr = lsSheet.getLastRow() + 1;
    lsSheet.getRange(sr, 1, snapRows.length, snapRows[0].length).setValues(snapRows);
    Logger.log("Wrote %s rows to LIVE SNAPSHOTS", snapRows.length);
  }

  // ── Update PRICE INTELLIGENCE columns Q-V ─────────────────────────────────
  const pi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRICE INTELLIGENCE");
  if (pi && pi.getLastRow() >= 4) {
    const piLastRow = pi.getLastRow();
    const piData    = pi.getRange(4, 1, piLastRow - 3, 3).getValues();

    const piRowMap = {};
    for (let i = 0; i < piData.length; i++) {
      const key = piData[i][0] + "|||" + piData[i][1] + "|||" + piData[i][2];
      piRowMap[key] = i + 4;
    }

    const prevMedians = loadPreviousMedians_();

    let updated = 0;
    for (const [key, g] of Object.entries(groups)) {
      const piRow = piRowMap[key];
      if (!piRow) continue;

      const prices = g.prices.slice().sort((a, b) => a - b);
      const n = prices.length;
      const median = n
        ? (n % 2 === 0 ? (prices[n/2-1] + prices[n/2]) / 2 : prices[Math.floor(n/2)])
        : 0;
      const low = n ? prices[0] : 0;

      const daysOnMarket = g.oldestListed
        ? Math.round((now - g.oldestListed) / 86400000) : "";

      let askTrend = "";
      if (prevMedians[key] && prevMedians[key] > 0 && median > 0) {
        const pct = Math.round(((median - prevMedians[key]) / prevMedians[key]) * 100);
        askTrend = (pct >= 0 ? "+" : "") + pct + "%";
      }

      pi.getRange(piRow, PI_LIVE.ACTIVE_LISTINGS).setValue(n);
      pi.getRange(piRow, PI_LIVE.MEDIAN_ASK).setValue(n ? Math.round(median) : "");
      pi.getRange(piRow, PI_LIVE.LOW_ASK).setValue(n ? Math.round(low) : "");
      pi.getRange(piRow, PI_LIVE.ASK_TREND).setValue(askTrend);
      pi.getRange(piRow, PI_LIVE.DAYS_ON_MARKET).setValue(daysOnMarket);
      pi.getRange(piRow, PI_LIVE.LAST_SCRAPED).setValue(snapshotDate);
      updated++;
    }

    for (const [key, piRow] of Object.entries(piRowMap)) {
      if (!groups[key]) {
        const cur = pi.getRange(piRow, PI_LIVE.ACTIVE_LISTINGS).getValue();
        if (cur !== "" && cur !== 0) {
          pi.getRange(piRow, PI_LIVE.ACTIVE_LISTINGS).setValue(0);
          pi.getRange(piRow, PI_LIVE.LAST_SCRAPED).setValue(snapshotDate);
        }
      }
    }
    Logger.log("Updated PI live columns for %s combos", updated);
  }

  // ── Push classified listings to Supabase ─────────────────────────────────
  setLiveStatus_("Phase 3: syncing listings to Supabase...");
  const syncResult = syncListingsToSupabase_(llSheet, runId);

  props.setProperty("LIVE_INGESTION_LAST_RUN", runId);
  props.deleteProperty(PROP_LIVE_RUN_ID);
  props.deleteProperty(PROP_LIVE_RESUME_ROW);

  const doneMsg = "✅ Live ingestion complete!\n\n" +
    "Rows this run: " + thisRun.length + "\n" +
    "Valid classified: " + Object.keys(groups).length + " combos\n" +
    "LIVE SNAPSHOTS rows written: " + snapRows.length + "\n" +
    "Listings synced to Supabase: " + syncResult.inserted + "\n" +
    (syncResult.errors ? "⚠ Sync errors: " + syncResult.errors + "\n" : "") +
    "PRICE INTELLIGENCE updated: Yes\n" +
    "Run ID: " + snapshotDate;

  setLiveStatus_(doneMsg);
  try { SpreadsheetApp.getUi().alert(doneMsg); } catch(e) {}
}

// Load most recent median ask per combo from LIVE SNAPSHOTS (for trend calc)
function loadPreviousMedians_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIVE SNAPSHOTS");
  if (!sheet || sheet.getLastRow() < 2) return {};
  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(2, 1, lastRow - 1, LS.MEDIAN_ASK).getValues();
  const latest  = {};
  for (const row of data) {
    const date    = String(row[LS.SNAPSHOT_DATE - 1] || "").trim();
    const set_    = String(row[LS.SET_NAME - 1]      || "").trim();
    const player  = String(row[LS.PLAYER - 1]        || "").trim();
    const variant = String(row[LS.VARIANT - 1]       || "").trim();
    const median  = Number(row[LS.MEDIAN_ASK - 1])   || 0;
    if (!set_ || !player || !date) continue;
    const key = set_ + "|||" + player + "|||" + variant;
    if (!latest[key] || date > latest[key].date) latest[key] = { date, median };
  }
  const result = {};
  for (const [key, v] of Object.entries(latest)) result[key] = v.median;
  return result;
}


// ═══════════════════════════════════════════════════════════════════════════════
// MENU ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Starts the full 3-phase ingestion pipeline.
 */
function runLiveIngestion() {
  startLiveIngestion();
}

/**
 * Emergency stop — cancels any running pipeline.
 */
function stopLiveIngestion() {
  stopClassifyTrigger_();
  const props = PropertiesService.getScriptProperties();
  const runId = props.getProperty(PROP_LIVE_RUN_ID);

  if (runId) {
    try {
      const llSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIVE LISTINGS");
      if (llSheet && llSheet.getLastRow() > 1) {
        const lastRow = llSheet.getLastRow();
        const data = llSheet.getRange(2, LL.SCRAPED_AT, lastRow - 1, 1).getValues();
        const toDelete = [];
        for (let i = data.length - 1; i >= 0; i--) {
          if (String(data[i][0] || "").trim() === runId) toDelete.push(i + 2);
        }
        for (const row of toDelete) llSheet.deleteRow(row);
        Logger.log("Emergency stop: deleted %s incomplete rows for run %s", toDelete.length, runId);
      }
    } catch(e) {
      Logger.log("Could not clean up incomplete rows: " + e.message);
    }
  }

  props.deleteProperty(PROP_LIVE_RUN_ID);
  props.deleteProperty(PROP_LIVE_RESUME_ROW);
  setLiveStatus_("⏹ Stopped manually.");
  try { SpreadsheetApp.getUi().alert("Live ingestion stopped. Incomplete rows cleaned up."); } catch(e) {}
}

/**
 * Schedule automatic ingestion every 3 days at 6am.
 */
function setupLiveIngestionTrigger() {
  stopLiveIngestionTrigger(true);
  ScriptApp.newTrigger("runLiveIngestion")
    .timeBased().everyDays(3).atHour(6).create();
  SpreadsheetApp.getUi().alert(
    "✅ Live ingestion scheduled!\n\nWill run automatically every 3 days at 6am."
  );
}

/**
 * Remove the scheduled trigger.
 */
function stopLiveIngestionTrigger(silent) {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runLiveIngestion") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (!silent) SpreadsheetApp.getUi().alert("Removed " + removed + " ingestion trigger(s).");
}


// ════════════════════════════════════════════════════════════════════════════
// SUPABASE SYNC — keeps Supabase sales table in sync with SALES sheet
// ─────────────────────────────────────────────────────────────────────────
// cleanSelectedRowsSales() auto-syncs each batch after re-classifying.
// Run backfillSupabase() once to sync all existing rows.
//
// Setup (one-time):
//   1. In Apps Script editor → Project Settings → Script Properties
//   2. Add two properties:
//      SUPABASE_URL  →  https://sabzbyuqrondoayhwoth.supabase.co
//      SUPABASE_KEY  →  your anon/public key
//   3. Run backfillSupabase() once to sync existing rows
//   After that, syncing happens automatically at the end of each clean run.
//
// ════════════════════════════════════════════════════════════════════════════

// ── Config helpers ────────────────────────────────────────────────────────────

function getSupabaseConfig_() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('SUPABASE_URL');
  const key = props.getProperty('SUPABASE_KEY');
  if (!url || !key) throw new Error('SUPABASE_URL and SUPABASE_KEY must be set in Script Properties.');
  return { url, key };
}

// ── Set name → UUID cache ─────────────────────────────────────────────────────
// Fetches all sets from Supabase once per script run and caches in memory.
// Maps set name (e.g. "MLB 2025 Game Face") → set UUID

let _setIdCache = null;

function getSetIdCache_() {
  if (_setIdCache) return _setIdCache;
  const { url, key } = getSupabaseConfig_();
  const resp = UrlFetchApp.fetch(url + '/rest/v1/sets?select=id,name', {
    method: 'GET',
    headers: {
      'apikey': key,
      'Authorization': 'Bearer ' + key,
    },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to fetch sets from Supabase: ' + resp.getContentText());
  }
  const sets = JSON.parse(resp.getContentText());
  _setIdCache = {};
  sets.forEach(s => { _setIdCache[s.name] = s.id; });
  return _setIdCache;
}

// ── Row builder ───────────────────────────────────────────────────────────────
// Converts a raw sheet row array (0-indexed, cols A-U) into a Supabase sales record.
// Returns null if the row should be skipped (not valid, is a dupe, or missing title).

function sheetRowToSalesRecord_(row, rowIndex, setIdCache) {
  const title   = row[1];   // B
  const price   = row[12];  // M - clean price
  const player  = row[13];  // N
  const variant = row[14];  // O
  const setName = row[15];  // P
  const status  = String(row[16] || '').trim(); // Q
  const dupe    = String(row[11] || '').trim(); // L
  const date    = row[9];   // J
  const imageUrl = row[0];  // A
  const ebayUrl  = row[2];  // C
  const notes    = row[17]; // R
  const flagged  = String(row[18] || '').trim(); // S
  const source   = String(row[19] || 'eBay').trim(); // T
  const afa        = String(row[20] || '').trim(); // U
  const setMathId  = String(row[22] || '').trim(); // W

  // Skip invalid, dupes, or unclassified rows
  if (!title || status !== 'valid' || dupe === 'dupe') return null;

  // Resolve set_id — null if set name not found (will still insert, just without set_id)
  const setId = setName ? (setIdCache[String(setName).trim()] || null) : null;

  // Parse date safely
  let saleDate = null;
  if (date) {
    const d = new Date(date);
    if (!isNaN(d.getTime())) saleDate = d.toISOString().split('T')[0]; // YYYY-MM-DD
  }

  return {
    set_id:           setId,
    player:           String(player  || '').trim() || null,
    variant:          String(variant || '').trim() || null,
    title:            String(title).trim().substring(0, 500),
    price:            Number(price)  || null,
    date:             saleDate,
    image_url:        String(imageUrl || '').trim() || null,
    ebay_url:         String(ebayUrl  || '').trim() || null,
    afa_grade:        afa  || null,
    source:           source || 'eBay',
    status:           status,
    notes:            notes ? String(notes).trim() : null,
    flagged:          flagged === 'FLAG',
    sheets_row_index: rowIndex,
    set_math_id:      setMathId || null,
  };
}

// ── Core upsert ───────────────────────────────────────────────────────────────
// Sends an array of sales records to Supabase via upsert on sheets_row_index.
// Works in chunks of 200 to stay well under URL fetch limits.

function upsertSalesToSupabase_(records) {
  if (!records || !records.length) return { inserted: 0, errors: 0 };
  const { url, key } = getSupabaseConfig_();

  const CHUNK = 200;
  let inserted = 0, errors = 0;

  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(url + '/rest/v1/sales?on_conflict=sheets_row_index', {
        method: 'POST',
        headers: {
          'apikey': key,
          'Authorization': 'Bearer ' + key,
          'Content-Type':  'application/json',
          'Prefer':        'resolution=merge-duplicates',
        },
        payload: JSON.stringify(chunk),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code === 200 || code === 201) {
        inserted += chunk.length;
      } else {
        Logger.log('Supabase upsert error ' + code + ': ' + resp.getContentText());
        errors += chunk.length;
      }
    } catch(e) {
      Logger.log('Supabase upsert exception: ' + e.message);
      errors += chunk.length;
    }
  }

  return { inserted, errors };
}

// ── One-time backfill ─────────────────────────────────────────────────────────
// Run this ONCE manually to sync all existing cleaned rows to Supabase.
// Safe to re-run — upserts on sheets_row_index so no duplicates.

function backfillSupabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) { SpreadsheetApp.getUi().alert(SALES_SHEET_NAME + ' sheet not found.'); return; }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows found.'); return; }

  setCleanStatus('🔄 Backfilling Supabase...');

  let setIdCache;
  try {
    setIdCache = getSetIdCache_();
  } catch(e) {
    SpreadsheetApp.getUi().alert('Failed to connect to Supabase:\n\n' + e.message);
    return;
  }

  const rowCount = lastRow - 1; // rows 2 to lastRow
  const allData  = sh.getRange(2, 1, rowCount, 23).getValues(); // cols A-W

  const records = [];
  for (let i = 0; i < allData.length; i++) {
    const rowNum = i + 2; // 1-based sheet row number
    const record = sheetRowToSalesRecord_(allData[i], rowNum, setIdCache);
    if (record) records.push(record);
  }

  if (!records.length) {
    SpreadsheetApp.getUi().alert('No valid rows found to sync.');
    return;
  }

  const result = upsertSalesToSupabase_(records);
  setCleanStatus('✅ Backfill complete');
  SpreadsheetApp.getUi().alert(
    '✅ Supabase backfill complete!\n\n' +
    'Rows synced: ' + result.inserted + '\n' +
    'Errors: '      + result.errors   + '\n\n' +
    (result.errors ? 'Check Apps Script logs for error details.' : 'All good!')
  );
}

// ── Sync selected rows ────────────────────────────────────────────────────────
// Select any rows in SALES and run this to push them to Supabase.
// Use this after manually correcting a row that was already cleaned.

function syncSelectedRowsToSupabase() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SALES_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select rows in the ' + SALES_SHEET_NAME + ' sheet first.');
    return;
  }

  const selection = sheet.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert('No rows selected. Highlight one or more rows first.');
    return;
  }

  let setIdCache;
  try {
    setIdCache = getSetIdCache_();
  } catch(e) {
    SpreadsheetApp.getUi().alert('Failed to connect to Supabase:\n\n' + e.message);
    return;
  }

  const startRow = selection.getRow();
  const numRows  = selection.getNumRows();

  // Enforce data rows only (row 2+)
  const firstDataRow = Math.max(startRow, 2);
  const lastDataRow  = startRow + numRows - 1;
  if (firstDataRow > lastDataRow) {
    SpreadsheetApp.getUi().alert('No data rows selected (rows must be row 2 or below).');
    return;
  }

  const count   = lastDataRow - firstDataRow + 1;
  const data    = sheet.getRange(firstDataRow, 1, count, 23).getValues(); // cols A-W
  const records = [];

  for (let i = 0; i < data.length; i++) {
    const rowNum = firstDataRow + i;
    const record = sheetRowToSalesRecord_(data[i], rowNum, setIdCache);
    if (record) records.push(record);
  }

  if (!records.length) {
    SpreadsheetApp.getUi().alert(
      'No valid rows found in selection.\n\n' +
      'Rows must have Status = "valid" and not be marked as duplicates.'
    );
    return;
  }

  const result = upsertSalesToSupabase_(records);
  SpreadsheetApp.getUi().alert(
    '✅ Sync complete!\n\n' +
    'Rows selected: ' + count + '\n' +
    'Rows synced:   ' + result.inserted + '\n' +
    'Errors:        ' + result.errors + '\n\n' +
    (result.errors ? 'Check Apps Script logs for details.' : 'All good!')
  );
}




// ══════════════════════════════════════════════════════════════════════════════
// SYNC SET_MATH NAMES FROM SALES
// ──────────────────────────────────────────────────────────────────────────────
// Reads all unique player names from SALES col N (already AI-cleaned),
// then updates any set_math rows in Supabase whose player name doesn't match
// exactly — using case-insensitive fuzzy matching to find the right canonical.
//
// Run this after editing set_math manually or adding new sets.
// ══════════════════════════════════════════════════════════════════════════════

function syncSetMathNamesFromSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sh) { ui.alert(SALES_SHEET_NAME + ' sheet not found.'); return; }

  // ── 1. Get canonical player names from SALES col N ───────────────────────
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('No data in ' + SALES_SHEET_NAME + '.'); return; }

  const playerCol = sh.getRange(2, 14, lastRow - 1, 1).getValues(); // col N = 14
  const canonical = new Set(
    playerCol
      .map(r => String(r[0] || '').trim())
      .filter(n => n && n !== 'Unknown')
  );

  if (!canonical.size) { ui.alert('No player names found in ' + SALES_SHEET_NAME + ' column N.'); return; }
  Logger.log('Canonical players from ' + SALES_SHEET_NAME + ': ' + canonical.size);

  // ── 2. Fetch all set_math rows from Supabase ──────────────────────────────
  const { url, key } = getSupabaseConfig_();
  const sbHeaders = { 'apikey': key, 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' };

  setCleanStatus('🔄 Syncing set_math names from Supabase...');

  let mathRows;
  try {
    const resp = UrlFetchApp.fetch(url + '/rest/v1/set_math?select=id,set_id,player,variant,population', {
      method: 'GET',
      headers: sbHeaders,
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) {
      ui.alert('Failed to fetch set_math: ' + resp.getContentText().substring(0, 200));
      return;
    }
    mathRows = JSON.parse(resp.getContentText());
  } catch(e) {
    ui.alert('Error fetching set_math: ' + e.message);
    return;
  }

  Logger.log('set_math rows fetched: ' + mathRows.length);

  // ── 3. Build normalizer ───────────────────────────────────────────────────
  // Normalize: lowercase, strip punctuation except spaces, collapse spaces
  function norm(s) {
    return String(s || '').toLowerCase().trim()
      .replace(/[^a-z0-9 ]/g, '')
      .replace(/\s+/g, ' ');
  }

  // Build norm→canonical lookup from SALES names
  const normToCanonical = {};
  canonical.forEach(name => { normToCanonical[norm(name)] = name; });

  // ── 4. Find mismatches ────────────────────────────────────────────────────
  const toUpdate = []; // {id, sbPlayer, canonicalPlayer}
  const noMatch  = []; // {sbPlayer} — couldn't find a match

  mathRows.forEach(row => {
    const sbPlayer = String(row.player || '').trim();
    if (!sbPlayer || sbPlayer === 'Unknown') return;

    // Exact match — already correct
    if (canonical.has(sbPlayer)) return;

    // Normalized match
    const normSb = norm(sbPlayer);
    if (normToCanonical[normSb]) {
      toUpdate.push({ id: row.id, sbPlayer, canonicalPlayer: normToCanonical[normSb] });
      return;
    }

    // No match found
    noMatch.push(sbPlayer);
  });

  Logger.log('To update: ' + toUpdate.length + ', No match: ' + noMatch.length);

  if (!toUpdate.length && !noMatch.length) {
    setCleanStatus('✅ set_math names already in sync');
    ui.alert('✅ All set_math player names already match ' + SALES_SHEET_NAME + '.\n\nNothing to update.');
    return;
  }

  // ── 5. Confirm with user ──────────────────────────────────────────────────
  let msg = `Found ${toUpdate.length} name(s) to fix`;
  if (toUpdate.length > 0) {
    const preview = toUpdate.slice(0, 8).map(r => `  "${r.sbPlayer}" → "${r.canonicalPlayer}"`).join('\n');
    msg += ':\n' + preview;
    if (toUpdate.length > 8) msg += `\n  ...and ${toUpdate.length - 8} more`;
  }
  if (noMatch.length > 0) {
    msg += `\n\n⚠ ${noMatch.length} name(s) in set_math have NO match in ${SALES_SHEET_NAME}:\n`;
    msg += noMatch.slice(0, 6).map(n => '  "' + n + '"').join('\n');
    if (noMatch.length > 6) msg += '\n  ...and ' + (noMatch.length - 6) + ' more';
    msg += '\n\nThese will NOT be updated — check manually.';
  }
  msg += '\n\nProceed with the ' + toUpdate.length + ' update(s)?';

  const confirm = ui.alert('Sync set_math names from Sales', msg, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) { setCleanStatus(''); return; }

  // ── 6. PATCH each mismatched row by id ───────────────────────────────────
  let ok = 0, errors = 0;

  toUpdate.forEach(r => {
    try {
      const resp = UrlFetchApp.fetch(
        `${url}/rest/v1/set_math?id=eq.${encodeURIComponent(r.id)}`,
        {
          method: 'PATCH',
          headers: { ...sbHeaders, 'Prefer': 'return=minimal' },
          payload: JSON.stringify({ player: r.canonicalPlayer }),
          muteHttpExceptions: true,
        }
      );
      const code = resp.getResponseCode();
      if (code === 200 || code === 204) {
        ok++;
      } else {
        errors++;
        Logger.log('PATCH failed ' + code + ' for "' + r.sbPlayer + '": ' + resp.getContentText());
      }
    } catch(e) {
      errors++;
      Logger.log('PATCH exception for "' + r.sbPlayer + '": ' + e.message);
    }
  });

  setCleanStatus(errors ? '⚠ set_math sync done with errors' : '✅ set_math names synced');
  ui.alert(
    '✅ Sync complete!\n\n' +
    `Updated: ${ok}\nErrors: ${errors}\n` +
    (errors ? '\nCheck Apps Script logs (View → Logs) for error details.' : '') +
    (noMatch.length ? `\n\n⚠ ${noMatch.length} unmatched name(s) still need manual review.` : '')
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// SUPABASE → SHEET SYNC  +  SUPABASE-SOURCED CATALOG LOADER
// ══════════════════════════════════════════════════════════════════════════════

// ── Load SET CATALOG from Supabase sets table ─────────────────────────────────
// Returns same shape as loadSetCatalogFromSupabase() so classifyLiveListings_ needs no changes.
// Format: [{setName, league, year, basePop, keywords}]
function loadSetCatalogFromSupabase() {
  try {
    const { url, key } = getSupabaseConfig_();
    const headers = { 'apikey': key, 'Authorization': 'Bearer ' + key };

    const resp = UrlFetchApp.fetch(
      url + '/rest/v1/sets?select=name,league,year,case_price,figures_per_case,keywords,active&active=eq.true&order=name.asc',
      { method: 'GET', headers, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) throw new Error('sets fetch: ' + resp.getResponseCode());

    const sets = JSON.parse(resp.getContentText());
    return sets.map(s => ({
      setName:  s.name,
      league:   s.league  || '',
      year:     s.year    ? String(s.year) : '',
      basePop:  '',  // not directly on sets table; classifier uses set_math for this
      // keywords is a Postgres text[] — join into comma-separated string for compatibility
      keywords: Array.isArray(s.keywords) ? s.keywords.join(', ') : (s.keywords || ''),
    }));
  } catch(e) {
    Logger.log('Supabase sets fetch failed: ' + e.message);
    return [];
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// LIVE LISTINGS → SUPABASE
// ═══════════════════════════════════════════════════════════════════════════════

// ── Build set_math lookup: "setName|||player|||variant" → set_math UUID ────────
function buildSetMathCache_() {
  try {
    const { url, key } = getSupabaseConfig_();
    const headers = { 'apikey': key, 'Authorization': 'Bearer ' + key };

    // Fetch sets for id→name map
    const setsResp = UrlFetchApp.fetch(
      url + '/rest/v1/sets?select=id,name',
      { method: 'GET', headers, muteHttpExceptions: true }
    );
    if (setsResp.getResponseCode() !== 200) throw new Error('sets fetch failed');
    const setIdToName = {};
    JSON.parse(setsResp.getContentText()).forEach(s => { setIdToName[s.id] = s.name; });

    // Fetch all set_math rows
    const mathResp = UrlFetchApp.fetch(
      url + '/rest/v1/set_math?select=id,set_id,player,variant',
      { method: 'GET', headers, muteHttpExceptions: true }
    );
    if (mathResp.getResponseCode() !== 200) throw new Error('set_math fetch failed');
    const mathRows = JSON.parse(mathResp.getContentText());

    const cache = {};
    mathRows.forEach(row => {
      const setName = setIdToName[row.set_id];
      if (!setName) return;
      const key_ = setName + '|||' + row.player + '|||' + row.variant;
      cache[key_] = row.id;
    });

    Logger.log('Built set_math cache: %s entries', Object.keys(cache).length);
    return cache;
  } catch(e) {
    Logger.log('buildSetMathCache_ failed: ' + e.message);
    return {};
  }
}

// ── Sync classified listings from current run to Supabase ─────────────────────
// Full history model — each run creates new rows, no upsert over old ones.
// Dedup key: item_id + scraped_at (unique per listing per run).
function syncListingsToSupabase_(llSheet, runId) {
  if (!llSheet || !runId) return { inserted: 0, errors: 0 };

  const lastRow = llSheet.getLastRow();
  if (lastRow < 2) return { inserted: 0, errors: 0 };

  // Read all columns for this run
  const numCols = LL.SET_MATH_ID; // read up to col Y
  const data = llSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const records = [];
  for (const row of data) {
    const scrapedAt  = String(row[LL.SCRAPED_AT  - 1] || '').trim();
    const exclude    = String(row[LL.EXCLUDE      - 1] || '').trim();
    const player     = String(row[LL.PLAYER       - 1] || '').trim();

    // Only sync rows from this run that are classified (have a player) and not excluded
    if (scrapedAt !== runId) continue;
    if (exclude) continue;
    if (!player || player === 'Unknown') continue;

    const itemId      = String(row[LL.ITEM_ID       - 1] || '');
    const legacyId    = String(row[LL.LEGACY_ID     - 1] || '');
    const title       = String(row[LL.TITLE         - 1] || '').substring(0, 300);
    const askPrice    = Number(row[LL.ASK_PRICE     - 1]) || null;
    const currency    = String(row[LL.CURRENCY      - 1] || 'USD');
    const listingType = String(row[LL.LISTING_TYPE  - 1] || '');
    const hasBo       = row[LL.HAS_BEST_OFFER - 1] === true || row[LL.HAS_BEST_OFFER - 1] === 'TRUE';
    const bidCount    = Number(row[LL.BID_COUNT     - 1]) || null;
    const condition   = String(row[LL.CONDITION     - 1] || '');
    const listedDate  = row[LL.LISTED_DATE - 1];
    const daysListed  = Number(row[LL.DAYS_LISTED   - 1]) || null;
    const imageUrl    = String(row[LL.IMAGE_URL     - 1] || '');
    const ebayUrl     = String(row[LL.EBAY_URL      - 1] || '');
    const seller      = String(row[LL.SELLER        - 1] || '');
    const sellerScore = Number(row[LL.SELLER_SCORE  - 1]) || null;
    const variant     = String(row[LL.VARIANT       - 1] || '');
    const setName     = String(row[LL.SET_NAME      - 1] || '');
    const afa         = String(row[LL.AFA           - 1] || '') || null;
    const ambiguous   = String(row[LL.AMBIGUOUS     - 1] || '') || null;
    const setMathId   = String(row[LL.SET_MATH_ID   - 1] || '') || null;
    const status      = String(row[LL.STATUS        - 1] || '');

    // Parse listed date safely
    let listedDateStr = null;
    if (listedDate) {
      const d = new Date(listedDate);
      if (!isNaN(d.getTime())) listedDateStr = d.toISOString().split('T')[0];
    }

    records.push({
      item_id:       itemId,
      legacy_item_id: legacyId || null,
      scraped_at:    scrapedAt,
      title,
      ask_price:     askPrice,
      currency,
      listing_type:  listingType,
      has_best_offer: hasBo,
      bid_count:     bidCount,
      condition,
      listed_date:   listedDateStr,
      days_listed:   daysListed,
      image_url:     imageUrl || null,
      ebay_url:      ebayUrl  || null,
      seller:        seller   || null,
      seller_score:  sellerScore,
      player,
      variant,
      set_name:      setName,
      afa_grade:     afa,
      ambiguous:     ambiguous,
      set_math_id:   setMathId,
      listing_status: status || null,
    });
  }

  if (!records.length) {
    Logger.log('syncListingsToSupabase_: no classified records to sync');
    return { inserted: 0, errors: 0 };
  }

  Logger.log('syncListingsToSupabase_: syncing %s records', records.length);

  const { url, key } = getSupabaseConfig_();
  const headers = {
    'apikey': key,
    'Authorization': 'Bearer ' + key,
    'Content-Type': 'application/json',
    // Use ignore on conflict so full history is preserved — never overwrite
    'Prefer': 'resolution=merge-duplicates,return=minimal',
  };

  const CHUNK = 200;
  let inserted = 0, errors = 0;

  // Log first record for diagnostics (helps spot column mismatches)
  Logger.log('syncListingsToSupabase_: first record sample: %s',
    JSON.stringify(records[0]).substring(0, 300));

  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(
        url + '/rest/v1/listings?on_conflict=item_id',
        {
          method: 'POST',
          headers,
          payload: JSON.stringify(chunk),
          muteHttpExceptions: true,
        }
      );
      const code = resp.getResponseCode();
      // 200 = OK (with representation), 201 = Created, 204 = No Content (return=minimal upsert)
      if (code === 200 || code === 201 || code === 204) {
        inserted += chunk.length;
      } else {
        // Log the full error body to make schema/constraint mismatches visible
        const body = resp.getContentText();
        Logger.log('syncListingsToSupabase_: HTTP %s on chunk %s–%s: %s',
          code, i, i + chunk.length - 1, body.substring(0, 500));
        errors += chunk.length;
      }
    } catch(e) {
      Logger.log('syncListingsToSupabase_: exception on chunk %s–%s: %s',
        i, i + CHUNK - 1, e.message);
      errors += chunk.length;
    }
  }

  Logger.log('syncListingsToSupabase_: inserted=%s errors=%s', inserted, errors);
  return { inserted, errors };
}

// ─────────────────────────────────────────────────────────────────────────────
// MANUAL TOOL: Backfill all classified listings to Supabase
// ─────────────────────────────────────────────────────────────────────────────
// Run this manually when listings were classified but never synced (e.g. after
// a re-clean, or when the runId expired before finalize ran).
// Syncs ALL classified, non-excluded rows regardless of scraped_at / runId.
// Safe to re-run — upserts on item_id so no duplicates.

function backfillListingsToSupabase() {
  const llSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIVE LISTINGS");
  if (!llSheet) { SpreadsheetApp.getUi().alert('LIVE LISTINGS sheet not found.'); return; }

  const lastRow = llSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows.'); return; }

  const numCols = LL.SET_MATH_ID;
  const data = llSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const records = [];
  for (const row of data) {
    const exclude = String(row[LL.EXCLUDE - 1] || '').trim();
    const player  = String(row[LL.PLAYER  - 1] || '').trim();
    if (exclude || !player || player === 'Unknown') continue;

    const itemId      = String(row[LL.ITEM_ID       - 1] || '');
    if (!itemId) continue; // need item_id for upsert key

    const scrapedAt   = String(row[LL.SCRAPED_AT    - 1] || '');
    const legacyId    = String(row[LL.LEGACY_ID     - 1] || '');
    const title       = String(row[LL.TITLE         - 1] || '').substring(0, 300);
    const askPrice    = Number(row[LL.ASK_PRICE     - 1]) || null;
    const currency    = String(row[LL.CURRENCY      - 1] || 'USD');
    const listingType = String(row[LL.LISTING_TYPE  - 1] || '');
    const hasBo       = row[LL.HAS_BEST_OFFER - 1] === true || row[LL.HAS_BEST_OFFER - 1] === 'TRUE';
    const bidCount    = Number(row[LL.BID_COUNT     - 1]) || null;
    const condition   = String(row[LL.CONDITION     - 1] || '');
    const listedDate  = row[LL.LISTED_DATE - 1];
    const daysListed  = Number(row[LL.DAYS_LISTED   - 1]) || null;
    const imageUrl    = String(row[LL.IMAGE_URL     - 1] || '');
    const ebayUrl     = String(row[LL.EBAY_URL      - 1] || '');
    const seller      = String(row[LL.SELLER        - 1] || '');
    const sellerScore = Number(row[LL.SELLER_SCORE  - 1]) || null;
    const variant     = String(row[LL.VARIANT       - 1] || '');
    const setName     = String(row[LL.SET_NAME      - 1] || '');
    const afa         = String(row[LL.AFA           - 1] || '') || null;
    const ambiguous   = String(row[LL.AMBIGUOUS     - 1] || '') || null;
    const setMathId   = String(row[LL.SET_MATH_ID   - 1] || '') || null;
    const status      = String(row[LL.STATUS        - 1] || '');

    let listedDateStr = null;
    if (listedDate) {
      const d = new Date(listedDate);
      if (!isNaN(d.getTime())) listedDateStr = d.toISOString().split('T')[0];
    }

    records.push({
      item_id:        itemId,
      legacy_item_id: legacyId || null,
      scraped_at:     scrapedAt,
      title,
      ask_price:      askPrice,
      currency,
      listing_type:   listingType,
      has_best_offer: hasBo,
      bid_count:      bidCount,
      condition,
      listed_date:    listedDateStr,
      days_listed:    daysListed,
      image_url:      imageUrl || null,
      ebay_url:       ebayUrl  || null,
      seller:         seller   || null,
      seller_score:   sellerScore,
      player,
      variant,
      set_name:       setName,
      afa_grade:      afa,
      ambiguous:      ambiguous,
      set_math_id:    setMathId,
      listing_status: status || null,
    });
  }

  if (!records.length) {
    SpreadsheetApp.getUi().alert('No classified listings found to sync.');
    return;
  }

  Logger.log('backfillListingsToSupabase: syncing %s records', records.length);

  const { url, key } = getSupabaseConfig_();
  const headers = {
    'apikey': key,
    'Authorization': 'Bearer ' + key,
    'Content-Type': 'application/json',
    'Prefer': 'resolution=merge-duplicates,return=minimal',
  };

  const CHUNK = 200;
  let inserted = 0, errors = 0;

  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(
        url + '/rest/v1/listings?on_conflict=item_id',
        { method: 'POST', headers, payload: JSON.stringify(chunk), muteHttpExceptions: true }
      );
      const code = resp.getResponseCode();
      if (code === 200 || code === 201 || code === 204) {
        inserted += chunk.length;
      } else {
        Logger.log('backfillListingsToSupabase: HTTP %s: %s', code, resp.getContentText().substring(0, 500));
        errors += chunk.length;
      }
    } catch(e) {
      Logger.log('backfillListingsToSupabase: exception: %s', e.message);
      errors += chunk.length;
    }
  }

  const msg = '✅ Listings backfill complete!\n\n' +
    'Records synced: ' + inserted + '\n' +
    'Errors: ' + errors;
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ─────────────────────────────────────────────────────────────────────────────
// MANUAL TOOL: Re-classify selected rows and sync to Supabase
// ─────────────────────────────────────────────────────────────────────────────
// HOW TO USE:
//   1. Open the LIVE LISTINGS sheet.
//   2. Select one or more rows whose classification needs correcting (click
//      row numbers to select whole rows, or highlight any cells in those rows).
//   3. Run this function from Extensions → Apps Script → Run, or bind it to a
//      custom menu item (see onOpen below — add it there once happy with it).
//
// WHAT IT DOES:
//   • Re-classifies every selected row via Claude Haiku (same model as the
//     sales pipeline).
//   • Writes updated player / variant / set / afa / exclude / set_math_id back
//     to the sheet in-place.
//   • Upserts the corrected records to the Supabase `listings` table using
//     item_id as the conflict key, overwriting the old classification columns.
//   • Rows classified as excluded or "Unknown" are skipped for SB sync but
//     still get their sheet columns updated.
// ─────────────────────────────────────────────────────────────────────────────
function cleanAndSyncSelectedListings() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ui      = SpreadsheetApp.getUi();
  const llSheet = ss.getSheetByName("LIVE LISTINGS");

  if (!llSheet) {
    ui.alert("LIVE LISTINGS sheet not found.");
    return;
  }
  if (ss.getActiveSheet().getName() !== "LIVE LISTINGS") {
    ui.alert("Please navigate to the LIVE LISTINGS sheet first, then select the rows you want to clean.");
    return;
  }

  // ── Get selected rows ──────────────────────────────────────────────────────
  const selection = llSheet.getActiveRange();
  if (!selection) {
    ui.alert("No cells selected. Select one or more rows in LIVE LISTINGS first.");
    return;
  }

  const startRow = selection.getRow();
  const numRows  = selection.getNumRows();

  if (startRow < 2) {
    ui.alert("Row 1 is the header. Please select data rows only (row 2+).");
    return;
  }

  const numCols = LL.SET_MATH_ID;
  const data    = llSheet.getRange(startRow, 1, numRows, numCols).getValues();

  // Build list of rows that have a title (skip blank/header rows in selection)
  const toClassify = [];
  for (let i = 0; i < data.length; i++) {
    const title = String(data[i][LL.TITLE - 1] || "").trim();
    if (title) toClassify.push({ sheetRow: startRow + i, dataIdx: i, title });
  }

  if (!toClassify.length) {
    ui.alert("No listing titles found in selected rows.");
    return;
  }

  // ── Confirm ────────────────────────────────────────────────────────────────
  const confirm = ui.alert(
    "Re-classify Listings",
    `Re-classify ${toClassify.length} listing(s) via Claude Haiku and sync to Supabase?\n\nThis will overwrite the classification columns (player / variant / set / etc.) for these rows in both the sheet and the database.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // ── Load catalogs ──────────────────────────────────────────────────────────
  let setCatalog = [], playerRoster = {}, setMathCache = {};
  try { setCatalog   = loadSetCatalogFromSupabase(); }   catch(e) { Logger.log("setCatalog load failed: " + e.message); }
  try { playerRoster = loadPlayerRosterFromSupabase(); } catch(e) { Logger.log("playerRoster load failed: " + e.message); }
  try { setMathCache = buildSetMathCache_(); }            catch(e) { Logger.log("setMathCache load failed: " + e.message); }

  // ── Classify via Haiku ─────────────────────────────────────────────────────
  let results;
  try {
    results = classifyBatch(toClassify.map(r => r.title), setCatalog, playerRoster);
  } catch(e) {
    ui.alert("Classification error:\n" + e.message);
    return;
  }

  if (!results || results.length !== toClassify.length) {
    ui.alert(`Unexpected result count from Haiku (${(results || []).length} vs ${toClassify.length}). Aborting.`);
    return;
  }

  // ── Write classification back to sheet ─────────────────────────────────────
  for (let idx = 0; idx < toClassify.length; idx++) {
    const r   = toClassify[idx];
    const res = results[idx] || {};
    const smKey    = (res.set || "") + "|||" + (res.player || "") + "|||" + (res.variant || "");
    const setMathId = (!res.exclude && setMathCache) ? (setMathCache[smKey] || "") : "";

    // Cols R–V: player, variant, set, afa, exclude
    llSheet.getRange(r.sheetRow, LL.PLAYER, 1, 5).setValues([[
      toTitleCase(res.player || ""),
      res.variant || "",
      res.set     || "",
      res.afa     || "",
      res.exclude ? (res.reason || "exclude") : "",
    ]]);
    // Col X: ambiguous (Haiku doesn't return this field — clear it)
    llSheet.getRange(r.sheetRow, LL.AMBIGUOUS,   1, 1).setValue("");
    // Col Y: set_math_id
    llSheet.getRange(r.sheetRow, LL.SET_MATH_ID, 1, 1).setValue(setMathId);
  }

  // ── Build Supabase records ─────────────────────────────────────────────────
  const { url, key } = getSupabaseConfig_();
  const records = [];

  for (let idx = 0; idx < toClassify.length; idx++) {
    const res = results[idx] || {};
    if (res.exclude || !res.player || res.player === "Unknown") continue;

    const row         = data[toClassify[idx].dataIdx];
    const itemId      = String(row[LL.ITEM_ID      - 1] || "");
    if (!itemId) continue;

    const legacyId    = String(row[LL.LEGACY_ID    - 1] || "") || null;
    const scrapedAt   = String(row[LL.SCRAPED_AT   - 1] || "").trim();
    const title       = String(row[LL.TITLE        - 1] || "").substring(0, 300);
    const askPrice    = Number(row[LL.ASK_PRICE    - 1]) || null;
    const currency    = String(row[LL.CURRENCY     - 1] || "USD");
    const listingType = String(row[LL.LISTING_TYPE - 1] || "");
    const hasBo       = row[LL.HAS_BEST_OFFER - 1] === true || row[LL.HAS_BEST_OFFER - 1] === "TRUE";
    const bidCount    = Number(row[LL.BID_COUNT    - 1]) || null;
    const condition   = String(row[LL.CONDITION    - 1] || "");
    const listedDate  = row[LL.LISTED_DATE - 1];
    const daysListed  = Number(row[LL.DAYS_LISTED  - 1]) || null;
    const imageUrl    = String(row[LL.IMAGE_URL    - 1] || "") || null;
    const ebayUrl     = String(row[LL.EBAY_URL     - 1] || "") || null;
    const seller      = String(row[LL.SELLER       - 1] || "") || null;
    const sellerScore = Number(row[LL.SELLER_SCORE - 1]) || null;
    const smKey       = (res.set || "") + "|||" + (toTitleCase(res.player || "")) + "|||" + (res.variant || "");
    const setMathId   = setMathCache[smKey] || null;

    let listedDateStr = null;
    if (listedDate) {
      const d = new Date(listedDate);
      if (!isNaN(d.getTime())) listedDateStr = d.toISOString().split("T")[0];
    }

    records.push({
      item_id:         itemId,
      legacy_item_id:  legacyId,
      scraped_at:      scrapedAt,
      title,
      ask_price:       askPrice,
      currency,
      listing_type:    listingType,
      has_best_offer:  hasBo,
      bid_count:       bidCount,
      condition,
      listed_date:     listedDateStr,
      days_listed:     daysListed,
      image_url:       imageUrl,
      ebay_url:        ebayUrl,
      seller,
      seller_score:    sellerScore,
      player:          toTitleCase(res.player || ""),
      variant:         res.variant  || "",
      set_name:        res.set      || "",
      afa_grade:       res.afa      || null,
      ambiguous:       null,
      set_math_id:     setMathId,
      is_active:       true,   // re-cleaning → treat as active
    });
  }

  // ── Upsert to Supabase ─────────────────────────────────────────────────────
  const sbHeaders = {
    "apikey":        key,
    "Authorization": "Bearer " + key,
    "Content-Type":  "application/json",
    "Prefer":        "resolution=merge-duplicates,return=minimal",
  };

  const CHUNK = 50;
  let synced = 0, errors = 0;

  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(url + "/rest/v1/listings?on_conflict=item_id", {
        method: "POST",
        headers: sbHeaders,
        payload: JSON.stringify(chunk),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code === 200 || code === 201 || code === 204) {
        synced += chunk.length;
      } else {
        Logger.log("cleanAndSync: HTTP %s: %s", code, resp.getContentText().substring(0, 400));
        errors += chunk.length;
      }
    } catch(e) {
      Logger.log("cleanAndSync: exception: %s", e.message);
      errors += chunk.length;
    }
  }

  const skipped = toClassify.length - records.length;
  ui.alert(
    "Done!",
    `Classified: ${toClassify.length} row(s)\n` +
    `Synced to Supabase: ${synced}\n` +
    (skipped  ? `Skipped (excluded/unknown): ${skipped}\n` : "") +
    (errors   ? `⚠ Errors: ${errors} — check Apps Script logs\n` : ""),
    ui.ButtonSet.OK
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// Ghostladder — Self-Calibrating Price Engine: Apps Script Functions
//
// Paste these functions into your existing Google Sheets bound Apps Script.
// Each function is standalone and can be run/triggered independently.
// Call runPostIngestion() from your existing ingestion trigger.
// ─────────────────────────────────────────────────────────────────────────────

const SB_URL_ = 'https://sabzbyuqrondoayhwoth.supabase.co';
const SB_KEY_ = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNhYnpieXVxcm9uZG9heWh3b3RoIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MzQwNjgyOSwiZXhwIjoyMDg4OTgyODI5fQ.TJdOISx5wkK7jHvC1PjhrZQVJ7blfdU0k1iqCqEuFPI';

// Listings not seen within this many days are deactivated (is_active → false).
// Raise this if your scraper runs less frequently than daily.
const LISTING_STALE_DAYS_ = 3;

// ── Supabase REST helpers ─────────────────────────────────────────────────────

function sbGet_(table, qs, opts) {
  opts = opts || {};
  var headers = { apikey: SB_KEY_, Authorization: 'Bearer ' + SB_KEY_ };
  // PostgREST defaults to 1000 rows. For large tables, set a higher Range.
  if (opts.limit) {
    headers['Range'] = '0-' + (opts.limit - 1);
    headers['Range-Unit'] = 'items';
  }
  const res = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/${table}?${qs}`, {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200 && res.getResponseCode() !== 206)
    throw new Error(`sbGet_ ${table}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}

// Paginated fetch — works around Supabase 1000-row limit
function sbGetAll_(table, qs) {
  var all = [], off = 0;
  while (true) {
    var batch = sbGet_(table, qs + '&offset=' + off + '&limit=1000');
    all = all.concat(batch);
    if (batch.length < 1000) break;
    off += 1000;
  }
  return all;
}

function sbUpsert_(table, rows, onConflict) {
  if (!rows.length) return;
  // onConflict: comma-separated column names for the conflict target.
  // Required when the table's PK is a generated uuid but the logical
  // unique key is a different column (or set of columns). Without this,
  // PostgREST defaults to the PK, fails to match any existing row, and
  // then hits the unique constraint on INSERT.
  const qs = onConflict ? `?on_conflict=${encodeURIComponent(onConflict)}` : '';
  const res = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/${table}${qs}`, {
    method: 'post',
    headers: {
      apikey: SB_KEY_,
      Authorization: `Bearer ${SB_KEY_}`,
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates,return=minimal',
    },
    payload: JSON.stringify(rows),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300)
    throw new Error(`sbUpsert_ ${table}: ${res.getContentText()}`);
}

function sbPatch_(table, filter, data) {
  const res = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/${table}?${filter}`, {
    method: 'patch',
    headers: {
      apikey: SB_KEY_,
      Authorization: `Bearer ${SB_KEY_}`,
      'Content-Type': 'application/json',
      Prefer: 'return=minimal',
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300)
    throw new Error(`sbPatch_ ${table}: ${res.getContentText()}`);
}

function sbRpc_(fn, params) {
  const res = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/rpc/${fn}`, {
    method: 'post',
    headers: {
      apikey: SB_KEY_,
      Authorization: `Bearer ${SB_KEY_}`,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(params),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200)
    throw new Error(`sbRpc_ ${fn}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}

// ─────────────────────────────────────────────────────────────────────────────
// markStaleListingsInactive
// Sets is_active = false on any listing whose scraped_at is older than
// LISTING_STALE_DAYS_. Called from runPostIngestion trigger.
// ─────────────────────────────────────────────────────────────────────────────
function markStaleListingsInactive_() {
  const cutoff = new Date(Date.now() - LISTING_STALE_DAYS_ * 24 * 3600 * 1000).toISOString();
  Logger.log(`markStaleListingsInactive_: deactivating rows with scraped_at < ${cutoff}`);
  sbPatch_('listings',
    `is_active=eq.true&scraped_at=lt.${encodeURIComponent(cutoff)}`,
    { is_active: false });
  Logger.log('markStaleListingsInactive_: done.');
}

// ─────────────────────────────────────────────────────────────────────────────
// Refreshes scraped_at to the current runId for every itemId still visible on
// eBay. This makes scraped_at behave as "last-seen-at" rather than "first-seen-
// at", enabling markStaleListingsInactive_ to correctly detect sold/ended
// listings after LISTING_STALE_DAYS_ days of absence.
// Chunked at 200 ids/request to stay safely under URL length limits.
// ─────────────────────────────────────────────────────────────────────────────
function refreshListingsScrapedAt_(itemIds, runId) {
  if (!itemIds || !itemIds.length) return;
  // Chunk size kept small: IDs like v1%7C358372033719%7C0 are ~22 chars each.
  // 50 × 22 + URL overhead ≈ 1,170 chars — safely under GAS's 2,048-char URL limit.
  const CHUNK = 50;
  let updated = 0;
  for (let i = 0; i < itemIds.length; i += CHUNK) {
    const chunk = itemIds.slice(i, i + CHUNK);
    // encodeURIComponent encodes pipe chars (|) which UrlFetchApp rejects in URLs.
    // PostgREST decodes %7C → | when matching, so the filter still works correctly.
    const inList = chunk.map(id => encodeURIComponent(String(id))).join(',');
    sbPatch_('listings',
      `item_id=in.(${inList})`,
      { scraped_at: runId });
    updated += chunk.length;
  }
  Logger.log(`refreshListingsScrapedAt_: updated scraped_at for ${updated} listings → runId=${runId}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Master hook — call this from your existing ingestion trigger
// ─────────────────────────────────────────────────────────────────────────────
function runPostIngestion() {
  // NOTE: ensureActiveModelWeights, populatePriceSignals, updateSellerProfiles,
  // logPredictionOutcomes, runCalibration have all been removed — those belong
  // to the retired listing-blend ML pipeline and no longer exist.
  markStaleListingsInactive_();  // deactivate listings not seen in LISTING_STALE_DAYS_
}

// ═══════════════════════════════════════════════════════════════════════════════
// CSV IMPORT — eBay Scrape → SALES sheet
// ─────────────────────────────────────────────────────────────────────────────
// Reads today's ebay_data_YYYY-MM-DD.csv from Google Drive,
// expands multi-sale rows with numeric Multi-Sale ID,
// deduplicates, appends to SALES sheet, triggers AI classification.
// ═══════════════════════════════════════════════════════════════════════════════

const SCRAPE_FOLDER_ID   = '1Bo7JWUYwi31wH4UYkLnoZuGGXu6EQBHZ';
const SCRAPE_FILE_PREFIX = 'ebay_data_';
const SALES_SHEET_NAME   = 'SALES';
const HISTORICAL_SHEET_NAME = 'HISTORICAL SCRAPE';
const SALES_HEADERS = [
  'Image URL','Title','eBay URL','','','','','','','Date Sold','',
  'Dupe','Price','Player','Variant','Set','Status','Notes','Flag','Source','AFA',
  'Multi-Sale ID','Set Math ID'
];

function ensureSalesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SALES_SHEET_NAME);
    sheet.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS]);
    sheet.getRange(1, 1, 1, SALES_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findLatestScrapeFile_() {
  const folder = DriveApp.getFolderById(SCRAPE_FOLDER_ID);
  const dates = [new Date(), new Date(Date.now() - 86400000)];
  for (const d of dates) {
    const dateStr  = Utilities.formatDate(d, 'America/New_York', 'yyyy-MM-dd');
    const filename = SCRAPE_FILE_PREFIX + dateStr + '.csv';
    const files    = folder.getFilesByName(filename);
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('Found scrape file: %s', filename);
      return { file, dateStr };
    }
  }
  return null;
}

function parsePrice_(str) {
  const n = parseFloat(String(str || '').replace(/[^0-9.]/g, ''));
  return isNaN(n) ? null : n;
}

function parseScrapeDate_(str) {
  if (!str || str === '-') return null;
  try { const d = new Date(str); return isNaN(d.getTime()) ? null : d; } catch(e) { return null; }
}

function buildSalesRow_(imageUrl, title, url, price, saleDate, multiSaleId) {
  const fullImage = (imageUrl || '').startsWith('//') ? 'https:' + imageUrl : (imageUrl || '');
  const row = new Array(23).fill('');
  row[0]  = fullImage;
  row[1]  = title;
  row[2]  = url;
  row[9]  = saleDate || '';
  row[12] = price    || '';
  row[16] = 'valid';
  row[19] = 'eBay';
  row[21] = multiSaleId;
  // row[22] = set_math_id — filled by backfillSetMathIds() after classification
  return row;
}

function importEbayScrapeCSV() {
  // NOTE: SpreadsheetApp.getUi() is NOT available in time-based trigger context.
  // All user feedback uses setCleanStatus() + Logger.log() instead.

  const found = findLatestScrapeFile_();
  if (!found) {
    const msg = 'No scrape file found. Expected: ebay_data_YYYY-MM-DD.csv in the ebay_scrape folder.';
    setCleanStatus('❌ ' + msg);
    Logger.log(msg);
    return;
  }

  const { file } = found;
  const salesSheet = ensureSalesSheet_();
  setCleanStatus('🔄 Importing ' + file.getName() + '...');

  const csvText = file.getBlob().getDataAsString();
  const parsed  = Utilities.parseCsv(csvText);
  if (parsed.length < 2) {
    setCleanStatus('❌ CSV file is empty: ' + file.getName());
    Logger.log('CSV file is empty: ' + file.getName());
    return;
  }

  const headers = parsed[0].map(function(h) { return h.trim(); });
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const required = ['Listing', 'Listing URL', 'Avg sold price', 'Total sold', 'Date last sold'];
  const missing = required.filter(function(h) { return col[h] === undefined; });
  if (missing.length) {
    const msg = 'CSV missing columns: ' + missing.join(', ');
    setCleanStatus('❌ ' + msg);
    Logger.log(msg);
    return;
  }

  // Dedup against existing SALES rows
  const salesLastRow = salesSheet.getLastRow();
  const existingUrls = new Set();
  if (salesLastRow > 1) {
    salesSheet.getRange(2, 3, salesLastRow - 1, 1).getValues().forEach(function(r) {
      const u = String(r[0] || '').trim();
      if (u) existingUrls.add(u.split('?')[0]);
    });
  }

  const newRows = [];
  let skippedDupes = 0, skippedInvalid = 0;
  const seenThisRun = new Set();

  for (let i = 1; i < parsed.length; i++) {
    const r         = parsed[i];
    const title     = String(r[col['Listing']]        || '').trim();
    const url       = String(r[col['Listing URL']]     || '').trim();
    const imageUrl  = String(r[col['Image URL']]       || '').trim();
    const priceStr  = String(r[col['Avg sold price']]  || '').trim();
    const totalSold = parseInt(r[col['Total sold']]    || '1', 10) || 1;
    const dateStr   = String(r[col['Date last sold']]  || '').trim();

    if (!title || !url) { skippedInvalid++; continue; }

    const baseUrl = url.split('?')[0];
    if (existingUrls.has(baseUrl) || seenThisRun.has(baseUrl)) { skippedDupes++; continue; }

    const price    = parsePrice_(priceStr);
    const saleDate = parseScrapeDate_(dateStr);

    for (let n = 1; n <= totalSold; n++) {
      newRows.push(buildSalesRow_(imageUrl, title, url, price, saleDate, n));
    }
    seenThisRun.add(baseUrl);
  }

  if (!newRows.length) {
    const msg = 'Import complete — no new rows (all ' + skippedDupes + ' listings already in SALES)';
    setCleanStatus('✅ ' + msg);
    Logger.log(msg);
    return;
  }

  const appendRow = salesSheet.getLastRow() + 1;
  salesSheet.getRange(appendRow, 1, newRows.length, 23).setValues(newRows);

  // Set CLEAN_RESUME_ROW so AI cleaner only processes new rows
  PropertiesService.getScriptProperties().deleteProperty('CLEAN_RESUME_ROW');
  PropertiesService.getScriptProperties().setProperty('CLEAN_RESUME_ROW', String(appendRow));

  setCleanStatus('🔄 Import done — ' + newRows.length + ' rows added. Starting AI cleaning...');

  // Non-blocking toast — no user click required, cleaning starts immediately
  SpreadsheetApp.getActiveSpreadsheet().toast(
    newRows.length + ' rows imported from ' + file.getName() + '. ' +
    skippedDupes + ' dupes skipped. Cleaning in progress — watch cell A2 for status.',
    '✅ Import complete',
    12
  );

  cleanNewSalesRows();
}

// ─────────────────────────────────────────────────────────────────────────────
// importHistoricalScrape
// One-off import of HISTORICAL SCRAPE sheet → SALES sheet.
// Does NOT auto-classify — run AI cleaning manually after spot-checking.
// ─────────────────────────────────────────────────────────────────────────────
function importHistoricalScrape() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(HISTORICAL_SHEET_NAME);

  if (!src) { ui.alert('Sheet "' + HISTORICAL_SHEET_NAME + '" not found.'); return; }

  const lastRow = src.getLastRow();
  if (lastRow < 2) { ui.alert('No data in ' + HISTORICAL_SHEET_NAME + '.'); return; }

  const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const required = ['Listing', 'Avg sold price', 'Total sold', 'Date last sold'];
  const missing = required.filter(function(h) { return col[h] === undefined; });
  if (missing.length) { ui.alert('Missing columns: ' + missing.join(', ')); return; }

  const data = src.getRange(2, 1, lastRow - 1, src.getLastColumn()).getValues();

  const salesSheet = ensureSalesSheet_();
  const newRows = [];
  let skippedInvalid = 0;

  // Track seen title+date+price combos for dupe flagging
  // We don't skip duplicates — we flag them for manual review
  const seenCombos = {};

  data.forEach(function(r) {
    const title     = String(r[col['Listing']]        || '').trim();
    const url       = String(r[col['Listing URL']]     || '').trim();
    const imageUrl  = String(r[col['Image URL']]       || '').trim();
    const priceStr  = String(r[col['Avg sold price']]  || '').trim();
    const totalSold = parseInt(r[col['Total sold']]    || '1', 10) || 1;
    const dateStr   = String(r[col['Date last sold']]  || '').trim();

    if (!title) { skippedInvalid++; return; }

    const price    = parsePrice_(priceStr);
    const saleDate = parseScrapeDate_(dateStr);

    // Dupe detection key: title + date + price + image URL
    // Image URL included since many historical rows have no listing URL,
    // but image URL is a reliable per-listing identifier
    const dupeKey = title + '|||'
      + (saleDate ? saleDate.toISOString().substring(0, 10) : dateStr)
      + '|||' + (price || '')
      + '|||' + imageUrl;
    const isDupe = !!seenCombos[dupeKey];
    seenCombos[dupeKey] = true;

    for (var n = 1; n <= totalSold; n++) {
      var salesRow = buildSalesRow_(imageUrl, title, url, price, saleDate, n);
      // Col L (index 11) = dupe flag
      if (isDupe) salesRow[11] = 'dupe';
      newRows.push(salesRow);
    }
  });

  if (!newRows.length) {
    ui.alert('No rows to import. Check that HISTORICAL SCRAPE has data.');
    return;
  }

  const appendRow = salesSheet.getLastRow() + 1;
  salesSheet.getRange(appendRow, 1, newRows.length, 23).setValues(newRows);

  const dupeCount = newRows.filter(function(r) { return r[11] === 'dupe'; }).length;
  Logger.log('importHistoricalScrape: wrote %s rows (%s flagged as potential dupes, %s invalid skipped)',
    newRows.length, dupeCount, skippedInvalid);

  ui.alert(
    '✅ Historical import complete!\n\n' +
    'Rows added to SALES: ' + newRows.length + '\n' +
    'Potential dupes flagged (col L): ' + dupeCount + '\n' +
    'Invalid skipped: ' + skippedInvalid + '\n\n' +
    'Next steps:\n' +
    '1. Review flagged dupes in col L (filter for "dupe")\n' +
    '2. Run "Clean new rows with AI" to classify\n' +
    '3. Run "Backfill Set Math IDs" to populate column W'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// backfillSetMathIds
// Fills col W (Set Math ID) for SALES rows where player/variant/set are filled
// but set_math_id is empty. Safe to run multiple times.
// ─────────────────────────────────────────────────────────────────────────────
function backfillSetMathIds() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(SALES_SHEET_NAME);

  if (!sheet) { ui.alert('SALES sheet not found.'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows in SALES sheet.'); return; }

  let setMathCache;
  try {
    setMathCache = buildSetMathCache_();
  } catch(e) {
    ui.alert('Failed to load set_math from Supabase: ' + e.message);
    return;
  }

  // Read cols N(14), O(15), P(16), W(23) in one batch
  // getRange(row, col, numRows, numCols) — cols N-W = cols 14-23 = 10 cols
  const data = sheet.getRange(2, 14, lastRow - 1, 10).getValues();
  let filled = 0, alreadyFilled = 0, noMatch = 0;
  const updates = []; // { rowNum, uuid }

  data.forEach(function(row, i) {
    const player   = String(row[0] || '').trim(); // N
    const variant  = String(row[1] || '').trim(); // O
    const setName  = String(row[2] || '').trim(); // P
    const existing = String(row[9] || '').trim(); // W (index 9 = col 23 - col 14)

    if (existing) { alreadyFilled++; return; }
    if (!player || !variant || !setName) return;

    const uuid = setMathCache[setName + '|||' + player + '|||' + variant];
    if (uuid) {
      updates.push({ rowNum: i + 2, uuid: uuid });
      filled++;
    } else {
      noMatch++;
    }
  });

  // Write all updates
  updates.forEach(function(u) {
    sheet.getRange(u.rowNum, 23).setValue(u.uuid);
  });

  Logger.log('backfillSetMathIds: filled=%s alreadyFilled=%s noMatch=%s', filled, alreadyFilled, noMatch);

  ui.alert(
    '✅ Set Math ID backfill complete!\n\n' +
    'IDs filled: ' + filled + '\n' +
    'Already filled: ' + alreadyFilled + '\n' +
    'No match: ' + noMatch +
    (noMatch > 0 ? '\n\nRows with no match may have unclassified or misclassified player/variant/set.' : '')
  );
}

// ── Schedule daily import at 1pm ─────────────────────────────────────────────
function setupDailyImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'importEbayScrapeCSV') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importEbayScrapeCSV')
    .timeBased().everyDays(1).atHour(13).create();
  SpreadsheetApp.getUi().alert('✅ Daily import scheduled!\n\nWill automatically import the eBay scrape CSV every day at 1pm and trigger AI cleaning.');
}

function stopDailyImportTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'importEbayScrapeCSV') { ScriptApp.deleteTrigger(t); removed++; }
  });
  SpreadsheetApp.getUi().alert('Removed ' + removed + ' import trigger(s).');
}

// ─────────────────────────────────────────────────────────────────────────────
// cleanSelectedRowsSales
// Re-classifies selected rows in the SALES sheet using Claude Haiku.
// Select rows first, then run. Re-classified rows are auto-synced to Supabase.
//
// Column mapping for SALES sheet:
//   B (2)  = Title (input)
//   N (14) = Player (output)
//   O (15) = Variant (output)
//   P (16) = Set (output)
//   Q (17) = Status / exclude reason (output)
//   U (21) = AFA grade (output)
//   R (18) = Notes / inferred warning (output)
// After classification, run backfillSetMathIds() to update col W.
// ─────────────────────────────────────────────────────────────────────────────
function cleanSelectedRowsSales() {
  const sheet     = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (sheet.getName() !== SALES_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select rows in the ' + SALES_SHEET_NAME + ' sheet.');
    return;
  }

  let setCatalog, playerRoster;
  try {
    setCatalog   = loadSetCatalogFromSupabase();
    playerRoster = loadPlayerRosterFromSupabase();
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error loading catalogs: ' + e.message);
    return;
  }

  const startRow = selection.getRow();
  const numRows  = selection.getNumRows();

  // Collect rows to process — skip header row (row 1)
  const toProcess = [];
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row < 2) continue; // skip header
    const title = sheet.getRange(row, 2).getValue(); // col B
    if (title) toProcess.push({ row, title: String(title) });
  }

  if (!toProcess.length) {
    SpreadsheetApp.getUi().alert('No rows with titles found in selection.');
    return;
  }

  let cleaned = 0;
  let synced  = 0;
  const BATCH = 20;

  for (let i = 0; i < toProcess.length; i += BATCH) {
    const batch = toProcess.slice(i, i + BATCH);
    const syncRowNums = [];
    try {
      const results = classifyBatch(batch.map(r => r.title), setCatalog, playerRoster);
      for (let j = 0; j < batch.length; j++) {
        const row    = batch[j].row;
        const result = results[j] || { set: '', player: '', variant: '', exclude: false, inferred: false };

        // Validate returned variant exists in roster for this player/set
        // If not, fall back to Base to avoid hallucinated variants (e.g. Victory from 🔥 emoji)
        const validatedResult = validateVariant_(result, playerRoster);

        sheet.getRange(row, 14).setValue(toTitleCase(validatedResult.player  || ''));  // N player
        sheet.getRange(row, 15).setValue(validatedResult.variant || '');               // O variant
        sheet.getRange(row, 16).setValue(validatedResult.set     || '');               // P set
        sheet.getRange(row, 17).setValue(validatedResult.exclude ? (validatedResult.reason || 'exclude') : 'valid'); // Q status
        // R notes: AI classification note + inferred warning + validateVariant_ correction (in priority order)
        const rNotes = [validatedResult.note, validatedResult.inferred ? '⚠ set inferred — please verify' : null, validatedResult.variantNote || null].filter(Boolean).join(' | ');
        sheet.getRange(row, 18).setValue(rNotes); // R notes
        if (validatedResult.afa) sheet.getRange(row, 21).setValue(validatedResult.afa); // U AFA

        cleaned++;
        syncRowNums.push(row);
      }
      // Sync this batch to Supabase immediately after writing to sheet
      try {
        syncSalesRowsToSupabase_(syncRowNums);
        synced += syncRowNums.length;
      } catch(e) {
        Logger.log('Supabase sync skipped: ' + e.message);
      }
    } catch(e) {
      for (let j = 0; j < batch.length; j++) {
        sheet.getRange(batch[j].row, 18).setValue('AI error: ' + e.message);
      }
    }
    Utilities.sleep(15000);
  }

  SpreadsheetApp.getUi().alert(
    'Done! ' + cleaned + ' rows re-classified and ' + synced + ' synced to Supabase.\n\n' +
    'Run "Backfill Set Math IDs" to update column W if needed.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// validateVariant_
// Post-processes AI classifier output to catch hallucinated variants.
// If the returned variant doesn't exist in the player's roster for that set,
// falls back to Base and notes the correction.
// This catches cases where Haiku misreads emojis, marketing terms (SSP, Plakata)
// or other noise as variant signals.
// ─────────────────────────────────────────────────────────────────────────────
function validateVariant_(result, playerRoster) {
  // If excluded, unknown player, or no set — pass through unchanged
  if (result.exclude || !result.player || result.player === 'Unknown' || !result.set) {
    return result;
  }

  // Normalize compound variant names the AI sometimes returns when a finish
  // descriptor ("Chrome") bleeds into the variant string.
  // e.g. "Fire Chrome" → "Fire", "Gold Chrome" → "Gold"
  const COMPOUND_NORM = {
    'Fire Chrome': 'Fire', 'fire chrome': 'Fire',
    'Gold Chrome': 'Gold', 'gold chrome': 'Gold',
    'Chrome Gold':  'Gold',
  };
  if (result.variant && COMPOUND_NORM[result.variant]) {
    result = Object.assign({}, result, { variant: COMPOUND_NORM[result.variant] });
  }

  const variant = result.variant || 'Base';

  // Always valid variants — present in every set
  const ALWAYS_VALID = new Set(['Base', 'Victory', 'Emerald', 'Gold', 'Chrome', 'Fire']);
  if (ALWAYS_VALID.has(variant)) return result;

  // For non-standard variants (SP, Vides, Bet On Women, etc.),
  // check the roster to confirm this variant exists for this player/set
  const rosterPlayers = (playerRoster && playerRoster[result.set]) || [];
  const rosterEntry   = rosterPlayers.find(function(p) {
    return p.player.toLowerCase() === result.player.toLowerCase();
  });

  if (!rosterEntry) return result; // no roster entry — can't validate, pass through

  // Check if the variant appears in this player's known variants string
  const variantsStr = rosterEntry.variants || '';
  // Extract variant name = everything before the ' /' population suffix
  // Must preserve multi-word names like "Bet On Women"
  const knownVariants = variantsStr.split(',').map(function(v) {
    return v.trim().replace(/\s*\/\d+.*$/, '').trim();
  });

  if (knownVariants.indexOf(variant) >= 0) {
    return result; // variant confirmed in roster — pass through
  }

  // Variant not in roster — fall back to Base and flag it
  Logger.log('validateVariant_: corrected %s %s %s → Base (not in roster)',
    result.player, result.set, variant);

  return Object.assign({}, result, {
    variant:     'Base',
    variantNote: '⚠ variant "' + variant + '" not in roster — defaulted to Base',
  });
}

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
//  9. Fall back to implied_price for players with zero data anywhere.
// ════════════════════════════════════════════════════════════════════════════

// ── Staleness / freshness ─────────────────────────────────────────────────────
//
// Freshness decays exponentially from 1.0 (just sold) toward 0 (very stale).
// τ = 90 days → at 90 days, own-data weight = e^(-1) ≈ 37%.
// Used by the Full Model mode to blend real sales prices with model estimates.

let STALENESS_TAU = 55; // days — default, overridden by engine_params
// AFA grading multiplier — population-aware curve.
// High-pop cards get the full premium (maxPremium); rare cards approach 1.0×.
// Curve: mult = 1 + (maxPremium - 1) × log(pop) / log(maxPop)
// Defaults overridden by engine_params (afa_grade_multiplier, afa_grade_max_pop).
let AFA_GRADE_MAX_PREMIUM = 1.5;
let AFA_GRADE_MAX_POP     = 100;

function afaMultiplier(pop) {
  if (!pop || pop <= 1) return 1.0;
  const maxPop = AFA_GRADE_MAX_POP || 100;
  const maxPrem = AFA_GRADE_MAX_PREMIUM || 1.5;
  if (pop >= maxPop) return maxPrem;
  return 1 + (maxPrem - 1) * Math.log(pop) / Math.log(maxPop);
}

// Backtester sets this to the snapshot date so freshness is computed relative
// to the simulated "now", not the real current date.
let FRESHNESS_REFERENCE_DATE = null;

function computeFreshness(lastSaleDate) {
  if (!lastSaleDate) return 0; // no date = treat as maximally stale
  const now = FRESHNESS_REFERENCE_DATE ? new Date(FRESHNESS_REFERENCE_DATE).getTime() : Date.now();
  const days = (now - new Date(lastSaleDate).getTime()) / 86400000;
  return Math.exp(-days / STALENESS_TAU);
}

// ── AFA Grading helpers ───────────────────────────────────────────────────────

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
    const g = map[k];
    const price = +(s.price || 0);
    if (price <= 0) return;
    const dateStr = (s.date || '').slice(0, 10);
    g.sum += price; g.count += 1;
    if (dateStr >= d30) { g.sum30d += price; g.count30d += 1; }
    if (dateStr >= d90) { g.sum90d += price; g.count90d += 1; }
  });
  return map;
}

function applyGradingNorm(products, gradedNormMap, popLookup) {
  if (!gradedNormMap || !products) return;
  products.forEach(prod => {
    const k = `${prod.player}|||${prod.variant}`;
    const g = gradedNormMap[k];
    if (!g || g.count === 0) return;
    // Population-aware multiplier: high-pop → full premium, rare → ~1.0×
    const pop  = (popLookup && prod.set && prod.variant) ? (popLookup[`${prod.set}|||${prod.variant}`] || null) : null;
    const mult = afaMultiplier(pop);
    if (prod.totalSales > 0) {
      if (g.count >= prod.totalSales) {
        prod.avgPrice = Math.round(g.sum / g.count / mult * 100) / 100;
      } else if (prod.avgPrice > 0) {
        const blended = prod.avgPrice * prod.totalSales - g.sum + g.sum / mult;
        prod.avgPrice = Math.round(Math.max(0, blended) / prod.totalSales * 100) / 100;
      }
    }
    if (g.count30d > 0 && prod.sales30d > 0) {
      if (g.count30d >= prod.sales30d) {
        prod.avg30d = Math.round(g.sum30d / g.count30d / mult * 100) / 100;
      } else if (prod.avg30d > 0) {
        const blended = prod.avg30d * prod.sales30d - g.sum30d + g.sum30d / mult;
        prod.avg30d = Math.round(Math.max(0, blended) / prod.sales30d * 100) / 100;
      }
    }
    if (g.count90d > 0 && prod.sales90d > 0) {
      if (g.count90d >= prod.sales90d) {
        prod.avg90d = Math.round(g.sum90d / g.count90d / mult * 100) / 100;
      } else if (prod.avg90d > 0) {
        const blended = prod.avg90d * prod.sales90d - g.sum90d + g.sum90d / mult;
        prod.avg90d = Math.round(Math.max(0, blended) / prod.sales90d * 100) / 100;
      }
    }
  });
}

// ── Comp price: recency-weighted blend across all sales windows ───────────────

function blendedCompPrice(p) {
  if (!p) return 0;
  const s30 = p.sales30d || 0, s90 = p.sales90d || 0, sAll = p.totalSales || 0;
  let wtSum = 0, wtTotal = 0;
  if (s30 > 0 && p.avg30d > 0) { wtSum += p.avg30d * s30 * 3; wtTotal += s30 * 3; }
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

function computeProximityWeight(targetPop, sourcePop) {
  if (!targetPop || !sourcePop) return 0.5;
  if (targetPop === sourcePop) return 1.0;
  return Math.exp(-0.3 * Math.abs(Math.log(targetPop / sourcePop)));
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

// Recency-weighted price — blends all available windows so recent prices
// dominate when available, but stale data still contributes when it's all we have.
// Used by ref-dist to track market trends (e.g. Victory price crash).
function recencyWeightedPrice(product) {
  if (!product) return 0;
  var num = 0, den = 0;
  if (product.avg30d  && product.sales30d  >= 1) { num += product.avg30d  * 3.0; den += 3.0; }
  if (product.avg90d  && product.sales90d  >= 1) { num += product.avg90d  * 1.5; den += 1.5; }
  if (product.avgPrice && product.totalSales >= 1) { num += product.avgPrice * 0.5; den += 0.5; }
  return den > 0 ? num / den : 0;
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

// ── Debug tracing (GAS version) ─────────────────────────────────────────────
// Set ENGINE_DEBUG_TARGET before calling buildEstimates to trace a specific card.
// var ENGINE_DEBUG_TARGET = { player: 'Zaccharie Risacher', variant: 'Gold' };
var ENGINE_DEBUG_TARGET = null;
function _dbg(player, variant) {
  if (!ENGINE_DEBUG_TARGET) return;
  if (ENGINE_DEBUG_TARGET.player && ENGINE_DEBUG_TARGET.player !== player) return;
  if (ENGINE_DEBUG_TARGET.variant && ENGINE_DEBUG_TARGET.variant !== variant) return;
  var args = Array.prototype.slice.call(arguments, 2);
  console.log('[ENGINE-DEBUG] ' + args.join(' '));
}
function _dbgActive(player, variant) {
  if (!ENGINE_DEBUG_TARGET) return false;
  if (ENGINE_DEBUG_TARGET.player && ENGINE_DEBUG_TARGET.player !== player) return false;
  if (ENGINE_DEBUG_TARGET.variant && ENGINE_DEBUG_TARGET.variant !== variant) return false;
  return true;
}

function buildPlayerValueScale(byPlayer, opts) {
  opts = opts || {};
  var crossSetProducts = opts.crossSetProducts || [];
  var playerPriors = opts.playerPriors || {};
  var currentSetLeague = opts.currentSetLeague || '';
  var setLeagueMap = opts.setLeagueMap || {};

  const playerImpliedBase = {};

  // ── Pass 1: direct within-set Base sales ──────────────────────────────────
  Object.entries(byPlayer).forEach(([player, pvars]) => {
    const bp = bestSalesPrice(pvars['Base']);
    if (bp > 0) {
      playerImpliedBase[player] = { price: bp, source: 'base', inferredFrom: null };
    }
  });

  // ── Pass 1.5: cross-set Base price enrichment ─────────────────────────────
  var crossSetByPlayer = {};
  crossSetProducts.forEach(function(p) {
    if (!p.player || p.player === 'Unknown' || p.variant !== 'Base') return;
    var bp = bestSalesPrice(p);
    if (bp <= 0) return;
    if (!crossSetByPlayer[p.player]) crossSetByPlayer[p.player] = [];
    var otherLeague = setLeagueMap[p.set] || '';
    var sameLeague = !!(currentSetLeague && otherLeague && currentSetLeague === otherLeague);
    crossSetByPlayer[p.player].push({
      price: bp, set: p.set, sameLeague: sameLeague,
      weight: (sameLeague ? 1.0 : 0.5) * Math.min(1.0, salesStrength(p) / 5),
    });
  });

  Object.entries(byPlayer).forEach(([player]) => {
    var crossPrices = crossSetByPlayer[player];
    if (!crossPrices || !crossPrices.length) return;
    var sameLeaguePrices = crossPrices.filter(function(c) { return c.sameLeague; });
    var usePrices = sameLeaguePrices.length ? sameLeaguePrices : crossPrices;
    var totalW = usePrices.reduce(function(s, c) { return s + c.weight; }, 0);
    if (totalW <= 0) return;
    var crossAvg = usePrices.reduce(function(s, c) { return s + c.price * c.weight; }, 0) / totalW;

    var existing = playerImpliedBase[player];
    if (existing && existing.source === 'base') {
      var withinSales = bestSalesPrice(byPlayer[player] && byPlayer[player]['Base']) > 0
        ? (byPlayer[player]['Base'].totalSales || 1)
        : 1;
      var crossSales = usePrices.reduce(function(s, c) { return s + (c.weight * 10); }, 0);
      var withinWeight = Math.min(withinSales, 10);
      var crossWeight  = Math.min(crossSales, 20);
      var blended = (existing.price * withinWeight + crossAvg * crossWeight) / (withinWeight + crossWeight);
      playerImpliedBase[player] = {
        price: blended,
        source: 'base+cross',
        inferredFrom: { withinPrice: existing.price, crossPrice: crossAvg, withinW: withinWeight, crossW: crossWeight, sets: usePrices.map(function(c) { return c.set; }) },
      };
    } else if (!existing) {
      playerImpliedBase[player] = { price: crossAvg, source: 'cross-set-base', inferredFrom: { sets: usePrices.map(function(c) { return c.set; }) } };
    }
  });

  // ── Pass 2: infer Base from variant comparisons ─────────────────────────
  Object.entries(byPlayer).forEach(([player, pvars]) => {
    if (playerImpliedBase[player]) return;

    const inferences = [];
    Object.entries(pvars).forEach(([variant, prod]) => {
      if (variant === 'Base') return;
      const thisPrice = bestSalesPrice(prod);
      if (thisPrice <= 0) return;
      const thisStr = salesStrength(prod);

      Object.entries(byPlayer).forEach(([otherPlayer, otherPvars]) => {
        const otherBase = playerImpliedBase[otherPlayer];
        if (!otherBase || (otherBase.source !== 'base' && otherBase.source !== 'cross-set-base')) return;
        const otherVariantPrice = bestSalesPrice(otherPvars[variant]);
        if (otherVariantPrice <= 0) return;

        const implied = (thisPrice / otherVariantPrice) * otherBase.price;
        const otherStr = salesStrength(otherPvars[variant]);
        const weight   = Math.sqrt(Math.max(0.1, thisStr) * Math.max(0.1, otherStr));
        inferences.push({ price: implied, weight, inferredFrom: { variant, otherPlayer } });
      });
    });

    if (inferences.length > 0) {
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

  // ── Pass 2.5: prior tier anchoring ────────────────────────────────────────
  if (knownPrices.length > 0) {
    var p95 = knownPrices[Math.min(knownPrices.length - 1, Math.floor(knownPrices.length * 0.95))];
    var p75 = knownPrices[Math.min(knownPrices.length - 1, Math.floor(knownPrices.length * 0.75))];
    var tierPrices = { 1: Math.max(p95, setMedianBase * 2), 2: Math.max(p75, setMedianBase * 1.3) };

    Object.keys(byPlayer).forEach(function(player) {
      if (playerImpliedBase[player]) return;
      var prior = playerPriors[player];
      if (!prior || !prior.tier || prior.tier > 2) return;
      playerImpliedBase[player] = { price: tierPrices[prior.tier], source: 'prior-tier-' + prior.tier, inferredFrom: null };
    });
  }

  // ── Pass 3: assign set median to players with no data ───────────────────
  Object.keys(byPlayer).forEach(player => {
    if (!playerImpliedBase[player]) {
      playerImpliedBase[player] = { price: setMedianBase, source: 'median', inferredFrom: null };
    }
  });

  // ── Compute percentiles (with tie-handling) ───────────────────────────────
  const pvsSorted = Object.entries(playerImpliedBase)
    .map(([player, data]) => ({ player, price: data.price, source: data.source }))
    .sort((a, b) => a.price - b.price);

  const playerPercentile = {};
  if (pvsSorted.length <= 1) {
    pvsSorted.forEach(function(e) { playerPercentile[e.player] = 0.5; });
  } else {
    var i = 0;
    while (i < pvsSorted.length) {
      var j = i;
      while (j < pvsSorted.length && pvsSorted[j].price === pvsSorted[i].price) j++;
      var midPct = ((i + j - 1) / 2) / (pvsSorted.length - 1);
      for (var k = i; k < j; k++) playerPercentile[pvsSorted[k].player] = midPct;
      i = j;
    }
  }

  // ── Prior-blended percentile ─────────────────────────────────────────────
  var PRIOR_FADE_RATE = 3;
  var TIER_PRIOR_PCT = { 1: 0.95, 2: 0.80 };
  var DEFAULT_PRIOR_PCT = 0.50;

  Object.keys(playerPercentile).forEach(function(player) {
    var pib = playerImpliedBase[player];
    if (!pib) return;
    var baseProd = byPlayer[player] && byPlayer[player]['Base'];
    var baseSales = (baseProd && baseProd.totalSales) ? baseProd.totalSales : 0;

    var priorTier = (playerPriors[player] && playerPriors[player].tier) || null;
    var priorPct = (priorTier && TIER_PRIOR_PCT[priorTier]) || DEFAULT_PRIOR_PCT;

    var priorWeight = 1 / (1 + baseSales / PRIOR_FADE_RATE);
    var rawPct = playerPercentile[player];
    var blended = priorPct * priorWeight + rawPct * (1 - priorWeight);
    playerPercentile[player] = blended;

    if (_dbgActive(player)) {
      _dbg(player, null, 'Prior blend: raw=' + (rawPct*100).toFixed(1) + '% prior=' + (priorPct*100).toFixed(0) + '% (tier=' + (priorTier||'none') + ') weight=' + (priorWeight*100).toFixed(0) + '% → blended=' + (blended*100).toFixed(1) + '% (baseSales=' + baseSales + ')');
    }
  });

  // Debug: log PVS for target player
  if (ENGINE_DEBUG_TARGET && ENGINE_DEBUG_TARGET.player) {
    var d = playerImpliedBase[ENGINE_DEBUG_TARGET.player];
    if (d) {
      console.log('[ENGINE-DEBUG] PVS for "' + ENGINE_DEBUG_TARGET.player + '": impliedBase=$' + d.price.toFixed(0) + ', source=' + d.source + ', percentile=' + (playerPercentile[ENGINE_DEBUG_TARGET.player]*100).toFixed(1) + '%');
      console.log('[ENGINE-DEBUG] PVS full ranking (' + pvsSorted.length + ' players):');
      pvsSorted.forEach(function(e, idx) {
        var pct = (playerPercentile[e.player] * 100).toFixed(1);
        var marker = e.player === ENGINE_DEBUG_TARGET.player ? ' <<<' : '';
        console.log('[ENGINE-DEBUG]   ' + pct + '%  $' + e.price.toFixed(0) + '  ' + e.player + ' (' + e.source + ')' + marker);
      });
    }
  }

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
// ── League Similarity Index ──────────────────────────────────────────────────
//
// Compares leagues by their variant price levels at each shared population tier.
// Uses variant sales (not Base) because variant populations are consistent across
// sets, so prices are directly comparable regardless of Base population differences.
//
// Returns: leagueSim[leagueA][leagueB] = 0.0–1.0 similarity score
//          (1.0 = identical price levels, lower = cheaper league)

function buildLeagueSimilarityIndex() {
  const leagueMap   = allData.setLeagueMap || {};
  const products    = allData.products || [];
  const setMathAll  = allData.setMath || {};

  // Build population lookup: "setName|||variant" → population
  const popLookup = {};
  Object.entries(setMathAll).forEach(([sn, rows]) => {
    rows.forEach(r => {
      if (r.variant && r.population != null)
        popLookup[`${sn}|||${r.variant}`] = +r.population;
    });
  });

  // Collect prices grouped by league + pop tier (skip Base — pop varies by set)
  // Key: "league|||pop" → array of recency-weighted prices
  const tierPrices = {};
  products.forEach(p => {
    if (!p.player || p.player === 'Unknown' || p.variant === 'Base') return;
    const league = leagueMap[p.set] || '';
    if (!league) return;
    const pop = popLookup[`${p.set}|||${p.variant}`];
    if (!pop) return;
    const price = recencyWeightedPrice(p);
    if (price <= 0) return;
    const key = `${league}|||${pop}`;
    if (!tierPrices[key]) tierPrices[key] = [];
    tierPrices[key].push(price);
  });

  // Compute median price per league+tier
  const tierMedians = {};
  Object.entries(tierPrices).forEach(([key, prices]) => {
    prices.sort((a, b) => a - b);
    const mid = Math.floor(prices.length / 2);
    tierMedians[key] = prices.length % 2 === 0
      ? (prices[mid - 1] + prices[mid]) / 2
      : prices[mid];
  });

  // Extract unique leagues and pop tiers
  const leagues = [...new Set(Object.keys(tierMedians).map(k => k.split('|||')[0]))];
  const pops    = [...new Set(Object.keys(tierMedians).map(k => +k.split('|||')[1]))];

  // Compare every league pair across shared pop tiers
  const sim = {};
  leagues.forEach(a => {
    sim[a] = {};
    leagues.forEach(b => {
      if (a === b) { sim[a][b] = 1.0; return; }
      let wtSum = 0, wtTotal = 0;
      pops.forEach(pop => {
        const medA = tierMedians[`${a}|||${pop}`];
        const medB = tierMedians[`${b}|||${pop}`];
        if (!medA || !medB) return;
        const nA = (tierPrices[`${a}|||${pop}`] || []).length;
        const nB = (tierPrices[`${b}|||${pop}`] || []).length;
        const sampleWeight = Math.min(nA, nB); // weight by smaller sample
        const ratio = Math.min(medA, medB) / Math.max(medA, medB);
        wtSum   += ratio * sampleWeight;
        wtTotal += sampleWeight;
      });
      sim[a][b] = wtTotal > 0 ? wtSum / wtTotal : 0.5;
    });
  });

  return sim;
}

// Returns: { points, sameLeague, refSets, nPoints, rangeMin, rangeMax,
//            targetPop, exactPopMatch, scaleFactor } | null

function buildReferenceDistribution(variant, targetPop, currentSetName, refSetData, leagueSim) {
  const thisSetLeague  = (allData.setLeagueMap || {})[currentSetName] || '';
  const allPoints      = [];
  const refSetsUsed    = new Set();
  let   anyExact       = false;
  let   anyApprox      = false;

  Object.entries(refSetData).forEach(([refSet, { byPlayer: refByPlayer, pvs: refPVS }]) => {
    const otherLeague   = (allData.setLeagueMap || {})[refSet] || '';
    const isLeagueMatch = !!(thisSetLeague && otherLeague && thisSetLeague === otherLeague);

    // League similarity weight: 1.0 for same league, data-driven for cross-league
    const leagueWeight = isLeagueMatch ? 1.0
      : (leagueSim && thisSetLeague && otherLeague && leagueSim[thisSetLeague])
        ? (leagueSim[thisSetLeague][otherLeague] ?? 0.5)
        : 0.5;

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
      // Use recency-weighted price so ref-dist tracks market trends
      // (recent sales dominate over stale all-time averages).
      const rawPrice = recencyWeightedPrice(pvars[matchV]);
      if (rawPrice <= 0) return;

      hasVariantData = true;
      const str = Math.min(3.0, salesStrength(pvars[matchV]) / 5);
      // Freshness decay: stale ref-dist points carry less weight so the
      // distribution tracks market trends (e.g. Victory price crash).
      // Floor of 0.15 ensures very stale points still provide shape context.
      const freshness = Math.max(0.15, computeFreshness(pvars[matchV]?.lastSaleDate));
      allPoints.push({
        // percentile assigned below via price-ranking (not PVS-based)
        price:      rawPrice * scaleFactor, // price adjusted for population difference
        player,
        set:        refSet,
        sameLeague: isLeagueMatch,
        weight:     leagueWeight * Math.max(0.2, str) * popProximity * freshness,
        scaleFactor,
        isExact,
        matchPop,
      });
    });

    if (hasVariantData) {
      if (isExact) anyExact = true; else anyApprox = true;
      refSetsUsed.add(refSet);
    }
  });

  if (!allPoints.length) return null;

  // All leagues contribute — similarity weights handle the blending.
  // sameLeague flag is still set for confidence scoring downstream.
  const hasSameLeague = allPoints.some(p => p.sameLeague);
  const points = allPoints;
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
    refSets:       [...refSetsUsed],
    nPoints:       points.length,
    rangeMin:      points[0],
    rangeMax:      points[points.length - 1],
    targetPop,
    exactPopMatch: !anyApprox,    // true only if every reference set had an exact pop match
    scaleFactor:   avgScaleFactor,
  };
}

// ── Variant-level trend adjustment ───────────────────────────────────────────
function computeVariantTrendFactor(matchVariant, refSetData, sameLeagueSets) {
  // Detect price trends by comparing recent windows (30d, 90d) to all-time.
  // Uses a cascade: try 30d first, fall back to 90d, then lastSale vs avg.
  // Even a single recent sale signal can trigger if the drop is large enough
  // and corroborated by multiple players.

  var refSetsToUse = sameLeagueSets && sameLeagueSets.size > 0
    ? [...sameLeagueSets]
    : Object.keys(refSetData);

  var d30  = [];  // players with 30d data: { recent, allTime }
  var d90  = [];  // players with 90d data: { recent, allTime }
  var last = [];  // all players: { recent: lastSalePrice, allTime: avgPrice }

  refSetsToUse.forEach(function(refSet) {
    var refData = refSetData[refSet];
    if (!refData || !refData.byPlayer) return;

    Object.entries(refData.byPlayer).forEach(function([player, pvars]) {
      var prod = pvars[matchVariant];
      if (!prod || !prod.avgPrice || prod.totalSales < 1) return;

      if (prod.avg30d && prod.sales30d >= 1) {
        d30.push({ recent: prod.avg30d, allTime: prod.avgPrice });
      }
      if (prod.avg90d && prod.sales90d >= 1) {
        d90.push({ recent: prod.avg90d, allTime: prod.avgPrice });
      }
      if (prod.lastSalePrice > 0 && prod.totalSales >= 2) {
        last.push({ recent: prod.lastSalePrice, allTime: prod.avgPrice });
      }
    });
  });

  var median = function(arr) {
    var s = arr.slice().sort(function(a, b) { return a - b; });
    var mid = Math.floor(s.length / 2);
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  };

  var computeRatio = function(pairs) {
    var ratios = pairs.map(function(p) { return p.allTime > 0 ? p.recent / p.allTime : 1; });
    return median(ratios);
  };

  // Cascade: prefer shorter windows (more recent signal)
  var factor = 1.0;
  var source = 'none';

  if (d30.length >= 2) {
    factor = computeRatio(d30);
    source = '30d (n=' + d30.length + ')';
  } else if (d90.length >= 2) {
    factor = computeRatio(d90);
    source = '90d (n=' + d90.length + ')';
  } else if (last.length >= 3) {
    factor = computeRatio(last);
    source = 'last-vs-avg (n=' + last.length + ')';
  }

  console.log('[TREND] variant=' + matchVariant + ' source=' + source + ' d30=' + d30.length + ' d90=' + d90.length + ' last=' + last.length);
  if (d30.length > 0) console.log('[TREND] 30d prices: ' + d30.map(function(p){return '$'+Math.round(p.recent)+'/$'+Math.round(p.allTime)}).join(', '));
  if (d90.length > 0) console.log('[TREND] 90d prices: ' + d90.map(function(p){return '$'+Math.round(p.recent)+'/$'+Math.round(p.allTime)}).join(', '));
  console.log('[TREND] trendFactor=' + factor.toFixed(2));

  // Clamp between 0.3 and 1.0
  return Math.max(0.3, Math.min(1.0, factor));
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

function buildGlobalRatioTable() {
  const variantSetPop = {};
  Object.entries(allData.setMath || {}).forEach(([sn, rows]) => {
    rows.forEach(r => {
      if (r.variant && r.population != null)
        variantSetPop[`${sn}|||${r.variant}`] = +r.population;
    });
  });
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
      if (popA >= popB) return;
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
  const RARITY_PREMIUM        = +w.rarity_premium        || 1.20;
  const EVIDENCE_K            = +w.evidence_k            || 0.0231;
  const ANCHOR_STALE_DISCOUNT = +(w.anchor_stale_discount ?? 1.0);
  STALENESS_TAU               = +w.staleness_tau          || 55;
  AFA_GRADE_MAX_PREMIUM       = +w.afa_grade_multiplier   || 1.5;
  AFA_GRADE_MAX_POP           = +w.afa_grade_max_pop      || 100; // 30 recency-weighted sales → 50%

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

  // Variant populations — used for monotonicity enforcement only
  const variantPop = {};
  mathRows.forEach(r => {
    if (r.variant && r.population != null && !variantPop[r.variant])
      variantPop[r.variant] = +r.population;
  });

  // ── 1. Build Player Value Scale for current set ──────────────────────────
  const crossSetProducts = (allData.products || []).filter(p => p.set !== setName);
  const playerPriors = allData.playerPriors || {};
  const currentSetLeague = (allData.setLeagueMap || {})[setName] || '';

  const pvs = buildPlayerValueScale(byPlayer, {
    crossSetProducts,
    playerPriors,
    currentSetLeague,
    setLeagueMap: allData.setLeagueMap || {},
  });
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

  // ── 2.6. Lazy-initialise league similarity index ───────────────────────
  if (!allData._leagueSimilarity) allData._leagueSimilarity = buildLeagueSimilarityIndex();
  const leagueSim = allData._leagueSimilarity;

  // ── 3. Cross-variant ratios from THIS set's real sales ───────────────────
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

      // 4a. Global population-tier ratio — always blend in
      if (popA && popB) {
        const gr = globalRatios[`${popA}|||${popB}`];
        if (gr && gr.n >= 2) {
          const within = setRatios[key];
          if (!within) {
            setRatios[key] = gr;
          } else {
            const withinN = within.n || 1;
            setRatios[key] = {
              ratio:  (withinN * within.ratio + N_GLOBAL * gr.ratio) / (withinN + N_GLOBAL),
              n:      withinN,
              source: withinN >= 5 ? 'sales' : 'global-pop',
            };
          }
          return;
        }
      }

      // 4b. Per-player cross-set fallback
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
    const compPrice = blendedCompPrice(p);
    if (compPrice <= 0) return;
    const k = `${p.player}|||${p.variant}`;
    if (!compByPlayerVariant[k]) compByPlayerVariant[k] = [];
    const otherLeague = (allData.setLeagueMap || {})[p.set] || '';
    const leagueMatch = otherLeague && thisSetLeague && otherLeague === thisSetLeague ? 0.90 : 0.70;
    const recency     = p.avg30d && p.sales30d >= 1 ? 1.0 : p.avg90d && p.sales90d >= 1 ? 0.7 : 0.4;
    const strength    = Math.min(1.0, salesStrength(p) / 10);
    compByPlayerVariant[k].push({ price: compPrice, weight: leagueMatch * recency * Math.max(0.2, strength), recency, set: p.set, sales: p.totalSales || 0 });
  });

  Object.entries(compByPlayerVariant).forEach(([k, comps]) => {
    const totalW = comps.reduce((s, c) => s + c.weight, 0);
    if (totalW <= 0) return;
    playerCompMap[k] = {
      price:       Math.round(comps.reduce((s, c) => s + c.price * c.weight, 0) / totalW),
      n:           comps.length,
      bestRecency: Math.max(...comps.map(c => c.recency)),
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
    allVariants.forEach(tv => {
      // avgPrice is the normalized ungraded average — the only reliable signal.
      // lastSalePrice is intentionally excluded: it may be from a graded sale
      // and is preserved on the product for display purposes only.
      const hasRealData = pvars[tv]?.avgPrice > 0;

      // ── 8a. Build ratio anchors from player's own real prices ─────────────
      // Cross-set fallback: if the player has no within-set sales for sv, check
      // playerCompMap[player|||sv] — their comp price for that variant from other
      // sets.  Cross-set anchors carry a 0.8× weight penalty.
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
      // Fresh comps (bestRecency = 1.0, i.e. 30d sales) are authoritative.
      // Stale comps blend with 8c/8d/8e estimates proportional to staleness.
      const compData = playerCompMap[`${player}|||${tv}`];
      const compRecency = compData ? (compData.bestRecency || 0.4) : 0;
      const compIsFresh = compData && compData.price > 0 && compRecency >= 1.0;

      if (compIsFresh) {
        // Fresh comp — authoritative, no need for 8c/8d/8e
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
          compRecency:    compRecency,
          anchorFreshness: 1.0,
        };
        if (hasRealData) {
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
        refDistCache[tv] = buildReferenceDistribution(tv, variantPop[tv] || null, setName, refSetData, leagueSim);
      }
      const refDistRaw = refDistCache[tv];
      const pct        = playerPercentile[player] ?? 0.5;

      let rangeEstimate = null, extrapolated = false, extrapolatedUp = false;
      if (refDistRaw) {
        // ── Variant-level trend adjustment ──────────────────────────────────
        if (!refDistCache['__trended__' + tv]) {
          refDistCache['__trended__' + tv] = true;
          var sameLeagueSets = new Set(refDistRaw.refSets);
          var trendFactorVal = computeVariantTrendFactor(tv, refSetData, sameLeagueSets);
          refDistCache['__trendVal__' + tv] = trendFactorVal;
          if (trendFactorVal < 1.0) {
            refDistRaw.points.forEach(function(p) { p.price = p.price * trendFactorVal; });
          }
        }
        var trendFactor = refDistCache['__trendVal__' + tv] || 1.0;
        if (trendFactor < 1.0) {
          _dbg(player, tv, 'Variant trend factor: ' + trendFactor.toFixed(2) + ' (recent prices ' + ((1-trendFactor)*100).toFixed(0) + '% below all-time)');
        }

        if (_dbgActive(player, tv)) {
          _dbg(player, tv, 'Ref-dist for variant="' + tv + '": ' + refDistRaw.nPoints + ' points from [' + refDistRaw.refSets.join(', ') + ']');
          _dbg(player, tv, 'Player PVS percentile: ' + (pct*100).toFixed(1) + '%');
          _dbg(player, tv, 'Ref-dist points (price-ranked):');
          refDistRaw.points.forEach(function(p) {
            _dbg(player, tv, '  ' + (p.percentile*100).toFixed(1) + '%  $' + p.price.toFixed(0) + '  ' + p.player + ' (' + p.set + ') w=' + p.weight.toFixed(2));
          });
        }
        const interp = interpolateFromDistribution(pct, refDistRaw.points);
        if (interp) {
          rangeEstimate  = interp.price;
          extrapolated   = interp.extrapolated;
          extrapolatedUp = interp.extrapolatedUp;
          _dbg(player, tv, 'Interpolation result: $' + interp.price.toFixed(0) + ' (extrapolated=' + interp.extrapolated + ', up=' + interp.extrapolatedUp + ')');
        }
      }

      // ── 8d. Ratio estimate (secondary, blended at high evidence) ──────────
      // Only compute ratio blending for rare variants (pop ≤ 10: Gold, Chrome, Fire).
      // For common variants (Victory /50, Emerald /25, etc.), ratio blending inflates
      // estimates because global pop ratios overestimate the value of common variants.
      // Exception: always allow ratio blending when an anchor shares the same
      // population as the target (e.g. Vides /25 → Emerald /25).
      var RATIO_BLEND_MAX_POP = 10;
      var targetPopForRatio = variantPop[tv] || null;
      var hasSamePopAnchor = targetPopForRatio != null && anchors.some(a => variantPop[a.sourceVariant] === targetPopForRatio);
      var hasOwnSaleAnchor = anchors.some(a => !a.crossSet && bestSalesPrice(pvars[a.sourceVariant]) > 0);
      var allowRatioBlend = (targetPopForRatio != null && targetPopForRatio <= RATIO_BLEND_MAX_POP) || hasSamePopAnchor || hasOwnSaleAnchor;

      let ratioEstimate = null;
      if (allowRatioBlend && anchors.length > 0) {
        const totalW = anchors.reduce((s, a) => s + a.weight, 0);
        if (totalW > 0) ratioEstimate = anchors.reduce((s, a) => s + a.impliedPrice * a.weight, 0) / totalW;
      }

      // ── 8e. Combine ────────────────────────────────────────────────────────
      // Anchor freshness: weighted average of each anchor's recency score.
      // Discounts the ratio method's share when its source sales are stale.
      let anchorFreshness = 1.0;
      if (ANCHOR_STALE_DISCOUNT > 0 && anchors.length > 0) {
        const aTotalW = anchors.reduce((s, a) => s + a.weight, 0);
        if (aTotalW > 0) {
          const rawFreshness = anchors.reduce((s, a) => s + (a.weight / aTotalW) * (a.crossSet ? 0.7 : recencyWeight(pvars[a.sourceVariant])), 0);
          anchorFreshness = 1 - ANCHOR_STALE_DISCOUNT * (1 - rawFreshness);
        }
      }

      let finalEstimate = null;
      if (rangeEstimate !== null) {
        const ratioShare = evidenceWeight * anchorFreshness;
        if (ratioEstimate !== null && ratioShare > 0.15) {
          finalEstimate = (1 - ratioShare) * rangeEstimate + ratioShare * ratioEstimate;
          _dbg(player, tv, 'Combine: range=$' + Math.round(rangeEstimate) + ' x ' + ((1-ratioShare)*100).toFixed(0) + '% + ratio=$' + Math.round(ratioEstimate) + ' x ' + (ratioShare*100).toFixed(0) + '% (anchorFreshness=' + anchorFreshness.toFixed(2) + ') = $' + Math.round(finalEstimate));
        } else {
          finalEstimate = rangeEstimate;
          var skipReason = !allowRatioBlend ? 'pop=' + targetPopForRatio + '>=' + (RATIO_BLEND_MAX_POP+1) : 'ratioShare=' + (ratioShare*100).toFixed(0) + '%';
          _dbg(player, tv, 'Combine: range-only=$' + Math.round(rangeEstimate) + ' (' + skipReason + ', ratioEstimate=' + (ratioEstimate !== null ? '$'+Math.round(ratioEstimate) : 'null') + ')');
        }
      } else if (ratioEstimate !== null) {
        finalEstimate = ratioEstimate;
        _dbg(player, tv, 'Combine: ratio-only=$' + Math.round(finalEstimate));
      } else {
        _dbg(player, tv, 'Combine: no estimate possible');
        return;
      }

      // ── 8f. Blend stale comp with model estimate ──────────────────────────
      // If a direct comp exists but is stale (bestRecency < 1.0), blend it in
      // proportional to its recency: fresh-ish comps still carry weight, very
      // stale comps defer to the model.
      if (compData && compData.price > 0 && !compIsFresh) {
        const compShare = compRecency; // 0.7 for 90d, 0.4 for all-time
        finalEstimate = compShare * compData.price + (1 - compShare) * finalEstimate;
        _dbg(player, tv, 'Stale comp blend: comp=$' + compData.price + ' x ' + (compShare*100).toFixed(0) + '% + model=$' + Math.round(finalEstimate) + ' x ' + ((1-compShare)*100).toFixed(0) + '% = $' + Math.round(finalEstimate));
      }

      finalEstimate = Math.round(finalEstimate);
      _dbg(player, tv, 'FINAL: $' + finalEstimate);

      // ── Confidence ─────────────────────────────────────────────────────────
      const hasSalesRatio = anchors.some(a => a.ratioSource === 'sales');
      // For rare variants (pop ≤ 10), ratioEstimate is required for Medium+.
      // For common variants (pop > 10), ratio blending is disabled — confidence
      // is based on ref-dist quality alone (same-league, non-extrapolated, n≥3).
      const hasRatioOrCommon = ratioEstimate || !allowRatioBlend;
      const confidence =
        refDistRaw && refDistRaw.sameLeague && !extrapolated && refDistRaw.nPoints >= 3 && hasRatioOrCommon && evidenceWeight > 0.3 ? 'Medium'
      : refDistRaw && refDistRaw.sameLeague && !extrapolated && refDistRaw.nPoints >= 3 && hasRatioOrCommon ? 'Medium-Low'
      : refDistRaw && !extrapolated && hasRatioOrCommon                                                     ? 'Low'
      : refDistRaw                                                                                          ? 'Very Low'
      : hasSalesRatio && anchors.length >= 2                                                                ? 'Low'
      : 'No Data';

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
        compUsed:       !!(compData && compData.price > 0 && !compIsFresh),
        compData:       (compData && compData.price > 0 && !compIsFresh) ? compData : null,
        compRecency:    (compData && compData.price > 0) ? compRecency : null,
        anchorFreshness: Math.round(anchorFreshness * 100),
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

  // ── 9. Fallback for players with zero real data anywhere ─────────────────
  // Use price_signals.implied_price as last resort
  mathRows.forEach(({ player, variant }) => {
    if (!player || player === 'Unknown') return;
    if (estimates[`${player}|||${variant}`]) return;
    if ((allData.products || []).some(p =>
        p.set === setName && p.player === player && p.variant === variant && p.avgPrice > 0)) return;
    const sig = (allData.signalMap || {})[`${setName}|||${player}|||${variant}`];
    if (!sig?.implied_price || +sig.implied_price <= 0) return;
    estimates[`${player}|||${variant}`] = {
      price:          Math.round(+sig.implied_price),
      confidence:     sig.implied_confidence || 'Very Low',
      pvs:            buildPvsInfo(player, playerImpliedBase, playerPercentile, pvsSorted),
      refDist:        null,
      rangeEstimate:  null,
      evidenceWeight: Math.round(evidenceWeight * 100),
      ratioEstimate:  null,
      anchors:        [],
      signal:         sig,
      compUsed:       false,
      compData:       null,
    };
  });

  return { estimates, refreshedEstimates, byPlayer, evidenceWeight, setRatios, pvs };
}

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
    // Build set_math_id lookup: "player|||variant" → uuid
    const smIdMap     = {};
    mathRows.forEach(r => { smIdMap[`${r.player}|||${r.variant}`] = r.id; });
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
        setMathId:      smIdMap[key] || null,
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
        setMathId:      smIdMap[key] || null,
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


// ════════════════════════════════════════════════════════════════════════════
// 1b. refreshLiveEstimates()
//     Run daily (e.g. 4 AM via createSnapshotTriggers).
//
//     For every active set, runs the full estimation engine and writes rows
//     into live_estimates — the pre-computed table that powers all browser pages.
//     Unlike estimate_snapshots (append-only weekly log), live_estimates always
//     holds exactly one row per set×player×variant and is replaced on each run.
//
//     This eliminates browser-side estimate-engine.js computation entirely.
//     Pages fetch live_estimates via simple REST queries instead.
// ════════════════════════════════════════════════════════════════════════════

function refreshLiveEstimates() {
  console.log('▶ refreshLiveEstimates started');

  // ── Load all reference data from Supabase ──────────────────────────────
  // loadAllData_() already handles grading normalization via
  // buildGradedNormMap + applyGradingNorm on the products array.
  const loaded = loadAllData_();
  allData = loaded.allData; // assign global so estimate-engine.gs functions see it

  // ── Load player priors (tier anchoring) ────────────────────────────────
  // loadAllData_() does not fetch these; the engine uses them for PVS
  // tier-based percentile anchoring (Tier 1 → 95th, Tier 2 → 80th).
  const playersRaw = sbGet_('players', 'select=name,prior_tier,league');
  allData.playerPriors = {};
  playersRaw.forEach(p => {
    allData.playerPriors[p.name] = { tier: p.prior_tier, league: p.league };
  });
  console.log(`  Player priors: ${playersRaw.length} players loaded`);

  // ── Get active engine_params ───────────────────────────────────────────
  const engineParams = loadActiveEngineParams_();
  console.log(`  Engine version: ${engineParams.version} (${engineParams.tag || 'no tag'})`);

  const allRows    = [];
  const computedAt = new Date().toISOString();

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
    const mathRows    = allData.setMath[set.name] || [];
    const validCombos = new Set(mathRows.map(r => `${r.player}|||${r.variant}`));
    const smIdMap     = {};
    mathRows.forEach(r => { smIdMap[`${r.player}|||${r.variant}`] = r.id; });
    let phantomCount  = 0;

    // ── Gap-fill estimates (cells with no real sales data) ───────────────
    Object.entries(estimates).forEach(([key, est]) => {
      if (!validCombos.has(key)) { phantomCount++; return; }
      const [player, variant] = key.split('|||');
      allRows.push(buildLiveEstimateRow_({
        setName: set.name, player, variant, est,
        engineVersion:  engineParams.version,
        evidenceWeight: evidenceWeightFraction,
        isRefreshed:    false,
        setMathId:      smIdMap[key] || null,
        computedAt,
      }));
    });

    // ── Refreshed estimates (real-data cells with freshness blend) ────────
    Object.entries(refreshedEstimates).forEach(([key, est]) => {
      if (!validCombos.has(key)) { phantomCount++; return; }
      const [player, variant] = key.split('|||');
      allRows.push(buildLiveEstimateRow_({
        setName: set.name, player, variant, est,
        engineVersion:  engineParams.version,
        evidenceWeight: evidenceWeightFraction,
        isRefreshed:    true,
        setMathId:      smIdMap[key] || null,
        computedAt,
      }));
    });

    if (phantomCount > 0) {
      console.warn(`  ⚠ ${set.name}: skipped ${phantomCount} combos not in set_math`);
    }
  }

  console.log(`  Total live_estimates rows to write: ${allRows.length}`);

  // ── Delete all existing rows, then insert fresh ────────────────────────
  // Using delete-then-insert (not upsert) for clean replacement.
  const delRes = UrlFetchApp.fetch(`${SB_URL_}/rest/v1/live_estimates?id=not.is.null`, {
    method: 'delete',
    headers: {
      apikey:          SB_KEY_,
      Authorization:   `Bearer ${SB_KEY_}`,
      'Content-Type':  'application/json',
      Prefer:          'return=minimal',
    },
    muteHttpExceptions: true,
  });
  if (delRes.getResponseCode() >= 300) {
    console.error(`  ✗ DELETE live_estimates failed: ${delRes.getContentText()}`);
    return;
  }
  console.log('  Cleared live_estimates table');

  // Insert in chunks
  sbInsertChunked_('live_estimates', allRows, SNAP_CHUNK_SIZE_);

  console.log(`✓ refreshLiveEstimates complete — ${allRows.length} rows written`);
}


// ── Build one live_estimates row ─────────────────────────────────────────────

function buildLiveEstimateRow_({ setName, player, variant, est,
                                  engineVersion, evidenceWeight, isRefreshed,
                                  setMathId, computedAt }) {
  const pvs     = est.pvs     || {};
  const refDist = est.refDist || null;

  return {
    set_name:        setName,
    player:          player,
    variant:         variant,
    set_math_id:     setMathId || null,

    // ── Core estimate ──────────────────────────────────────────────────
    estimated_price: est.price,
    confidence:      est.confidence  || null,
    method:          deriveMethod_(est),
    is_refreshed:    isRefreshed,

    // ── Refreshed-only fields ──────────────────────────────────────────
    own_price:       isRefreshed ? (est.ownPrice     || null) : null,
    freshness:       isRefreshed ? ((est.freshness || 0) / 100) : null, // engine returns 0–100, store 0–1
    model_estimate:  isRefreshed ? (est.modelEstimate || null) : null,

    // ── Engine reasoning ───────────────────────────────────────────────
    evidence_weight: typeof evidenceWeight === 'number' ? evidenceWeight : null,

    // PVS
    pvs_percentile:  pvs.percentile    ?? null,
    pvs_rank:        pvs.rank          ?? null,
    pvs_n_players:   pvs.nPlayers      ?? null,
    pvs_implied_base: pvs.impliedBase  ?? null,
    pvs_source:      pvs.source        ?? null,

    // Reference distribution
    ref_dist_info:   refDist ? JSON.stringify({
      refSets:          refDist.refSets,
      sameLeague:       refDist.sameLeague,
      nPoints:          refDist.nPoints,
      rangeMin:         refDist.rangeMin,
      rangeMax:         refDist.rangeMax,
      extrapolated:     refDist.extrapolated,
      extrapolatedUp:   refDist.extrapolatedUp,
      targetPop:        refDist.targetPop,
      exactPopMatch:    refDist.exactPopMatch,
      scaleFactor:      refDist.scaleFactor,
    }) : null,
    range_estimate:  est.rangeEstimate ?? null,

    // Ratio anchors
    ratio_estimate:  est.ratioEstimate ?? null,
    anchors:         est.anchors?.length ? JSON.stringify(est.anchors) : null,

    // Direct comp
    comp_used:       est.compUsed || false,
    comp_data:       est.compData ? JSON.stringify(est.compData) : null,
    comp_recency:    est.compRecency ?? null,

    // Anchor freshness (0–100, how fresh the ratio anchors are)
    anchor_freshness: est.anchorFreshness ?? null,

    // ── Metadata ───────────────────────────────────────────────────────
    engine_version:  engineVersion,
    computed_at:     computedAt,
  };
}


// ── Diagnostic dry-run — run this INSTEAD of takeEstimateSnapshots to learn ───
// which set_math products are being skipped and why.
// Nothing is written to Supabase. Check Apps Script > Executions > Logs.
//
// Output per set:
//   covered  = products that produced an estimate (gap-fill or refreshed)
//   skipped  = set_math rows with no estimate
//   reason A = has price_intelligence (real sales) but still no estimate
//              → rangeEstimate AND ratioEstimate both failed for a real-data cell
//   reason B = no price_intelligence, price_signals HAS implied_price
//              → Step 9 should have caught this — unexpected skip
//   reason C = no price_intelligence, no implied_price anywhere
//              → pure blind spot; needs a fallback
// ─────────────────────────────────────────────────────────────────────────────
function diagnoseEstimateSnapshots() {
  console.log('▶ diagnoseEstimateSnapshots — dry run, no writes');

  const loaded = loadAllData_();
  allData = loaded.allData;

  const engineParams = loadActiveEngineParams_();
  console.log(`  Engine v${engineParams.version} (${engineParams.tag || 'no tag'})`);

  let totalMath = 0, totalCovered = 0, totalSkipped = 0;
  const skippedByVariant = {};  // variant → count
  const skippedBySet     = {};  // setName → count
  const reasonCounts     = { A: 0, B: 0, C: 0 };

  for (const set of loaded.activeSets) {
    const setName     = set.name;
    const setProducts = loaded.allData.productsBySet[setName] || [];
    const mathRows    = (loaded.allData.setMath || {})[setName] || [];

    if (!mathRows.length) continue;
    totalMath += mathRows.length;

    let result;
    try {
      result = buildEstimates(setName, setProducts, engineParams);
    } catch (e) {
      console.error(`  ✗ buildEstimates crashed for "${setName}": ${e.message}`);
      totalSkipped += mathRows.length;
      skippedBySet[setName] = mathRows.length;
      continue;
    }

    const { estimates, refreshedEstimates } = result;
    const coveredKeys = new Set([
      ...Object.keys(estimates),
      ...Object.keys(refreshedEstimates),
    ]);

    let setCovered = 0, setSkipped = 0;
    const setSkipDetails = [];

    mathRows.forEach(({ player, variant }) => {
      if (!player || player === 'Unknown') { totalMath--; return; } // don't count Unknown
      const key = `${player}|||${variant}`;

      if (coveredKeys.has(key)) {
        setCovered++;
      } else {
        setSkipped++;
        // Categorise the skip
        const hasPriceIntel = setProducts.some(
          p => p.player === player && p.variant === variant && p.avgPrice > 0
        );
        const sig = (loaded.allData.signalMap || {})[`${setName}|||${player}|||${variant}`];
        const hasImplied = !!(sig?.implied_price && +sig.implied_price > 0);

        let reason;
        if (hasPriceIntel)       { reason = 'A'; } // real data, still skipped
        else if (hasImplied)     { reason = 'B'; } // implied_price existed, Step 9 missed it
        else                     { reason = 'C'; } // complete blind spot

        reasonCounts[reason]++;
        skippedByVariant[variant] = (skippedByVariant[variant] || 0) + 1;
        setSkipDetails.push(`${player} / ${variant} [${reason}]`);
      }
    });

    totalCovered += setCovered;
    totalSkipped += setSkipped;

    if (setSkipped > 0) {
      skippedBySet[setName] = setSkipped;
      console.log(
        `  ${setName}: ${setCovered}/${mathRows.length} covered, ${setSkipped} skipped`
      );
      // Log up to 10 skipped rows per set so the output stays readable
      setSkipDetails.slice(0, 10).forEach(d => console.log(`    ↳ ${d}`));
      if (setSkipDetails.length > 10)
        console.log(`    ↳ ... and ${setSkipDetails.length - 10} more`);
    } else {
      console.log(`  ${setName}: ✓ fully covered (${setCovered})`);
    }
  }

  // ── Summary ──────────────────────────────────────────────────────────────
  console.log('');
  console.log('═══ SUMMARY ═══════════════════════════════════════');
  console.log(`  Total set_math rows : ${totalMath}`);
  console.log(`  Covered             : ${totalCovered}`);
  console.log(`  Skipped             : ${totalSkipped}`);
  console.log('');
  console.log('  Skip reasons:');
  console.log(`    [A] Real data, estimation still failed : ${reasonCounts.A}`);
  console.log(`    [B] Implied price existed, Step 9 miss : ${reasonCounts.B}`);
  console.log(`    [C] Complete blind spot (no data)      : ${reasonCounts.C}`);
  console.log('');
  console.log('  Skipped by variant (top 10):');
  Object.entries(skippedByVariant)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .forEach(([v, n]) => console.log(`    ${v.padEnd(20)} ${n}`));
  console.log('');
  console.log('  Sets with most skipped:');
  Object.entries(skippedBySet)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .forEach(([s, n]) => console.log(`    ${s.padEnd(30)} ${n}`));
  console.log('═══════════════════════════════════════════════════');
}

// ── Build one estimate_snapshots row ──────────────────────────────────────────

function buildSnapshotRow_({ runId, snapshotAt, setName, player, variant, est,
                              engineVersion, evidenceWeight, isRefreshed, setMathId }) {
  const refDist = est.refDist || null;

  return {
    snapshot_run_id:    runId,
    snapshot_at:        snapshotAt,
    set_name:           setName,
    player:             player,
    variant:            variant,
    set_math_id:        setMathId || null,

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
    freshness:          isRefreshed ? ((est.freshness || 0) / 100) : null, // store 0–1
    is_refreshed:       isRefreshed,

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

  // ── Fetch unresolved snapshots that are at least 7 days old ───────────
  // We wait for the full 7-day measurement window to close before resolving.
  // Snapshots are taken weekly, so the "actual" we compare against is the
  // average market price during the 7 days following the snapshot — not a
  // single last_sale_price which may be a one-off outlier.
  const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();

  const unresolved = sbGet_('estimate_snapshots',
    'select=id,snapshot_at,set_name,player,variant,estimated_price' +
    '&actual_price=is.null' +
    '&snapshot_at=lt.' + sevenDaysAgo +
    '&order=snapshot_at.asc' +
    '&limit=5000'
  );

  if (!unresolved.length) {
    console.log('  No unresolved snapshots older than 7 days — nothing to do.');
    return;
  }
  console.log(`  Found ${unresolved.length} unresolved snapshot rows (≥7 days old)`);

  // ── We need the set_id → set_name map ─────────────────────────────────
  const setsRaw = sbGet_('sets', 'select=id,name&active=eq.true');
  const setIdToName_ = {};
  setsRaw.forEach(s => { setIdToName_[s.id] = s.name; });

  // ── Fetch all sales since the earliest unresolved snapshot ────────────
  // One bulk query; we'll slice each combo's 7-day window in memory.
  // We need price + afa_grade so we can compute the window average and
  // decide whether grading adjustment is needed.
  const earliestSnapshotDate = unresolved[0].snapshot_at.slice(0, 10);
  const salesRaw = sbGet_('sales',
    'select=set_id,player,variant,date,price,afa_grade' +
    '&date=gte.' + earliestSnapshotDate +
    '&status=neq.invalid',
    { limit: 10000 }
  );
  console.log(`  Raw sales loaded since ${earliestSnapshotDate}: ${salesRaw.length}`);

  // Build lookup: "setName|||player|||variant" → [{date, price, graded}]
  const salesByCombo = {};
  salesRaw.forEach(s => {
    const setName = setIdToName_[s.set_id];
    if (!setName || !s.player || !s.variant || !s.date || !(s.price > 0)) return;
    const k = `${setName}|||${s.player}|||${s.variant}`;
    if (!salesByCombo[k]) salesByCombo[k] = [];
    salesByCombo[k].push({ date: s.date.slice(0, 10), price: +s.price, graded: !!s.afa_grade });
  });

  // ── Load grading multiplier ────────────────────────────────────────────
  // When ALL window sales are AFA-graded, we divide the average by the
  // multiplier before computing error_pct. If any ungraded sales exist in
  // the window we use their average directly — no adjustment needed.
  const gradingParamsRow = sbGet_('engine_params',
    'select=afa_grade_multiplier&is_active=eq.true&limit=1'
  )[0];
  const gradingMult = +(gradingParamsRow?.afa_grade_multiplier) || 2.0;
  console.log(`  Grading multiplier: ${gradingMult}×`);

  // ── Match unresolved snapshots against 7-day post-snapshot window ──────
  const toResolve = [];
  const now       = new Date().toISOString();

  unresolved.forEach(snap => {
    const k        = `${snap.set_name}|||${snap.player}|||${snap.variant}`;
    const allSales = salesByCombo[k] || [];

    // Window: day after snapshot through 7 days later (inclusive).
    const snapDate  = snap.snapshot_at.slice(0, 10);           // 'YYYY-MM-DD'
    const windowEnd = new Date(
      new Date(snapDate + 'T00:00:00Z').getTime() + 7 * 24 * 60 * 60 * 1000
    ).toISOString().slice(0, 10);

    const windowSales = allSales.filter(s => s.date > snapDate && s.date <= windowEnd);
    if (!windowSales.length) return; // no sales occurred in the measurement window

    // Prefer ungraded sales for the average — the model estimates ungraded prices.
    // Only fall back to graded sales (with multiplier adjustment) if every sale
    // in the window was graded.
    const ungradedSales = windowSales.filter(s => !s.graded);
    const useSales      = ungradedSales.length > 0 ? ungradedSales : windowSales;
    const isAllGraded   = ungradedSales.length === 0;

    const avgActual = Math.round(
      useSales.reduce((sum, s) => sum + s.price, 0) / useSales.length
    );

    // Normalise graded-only averages so the engine isn't penalised for correctly
    // predicting the ungraded price of a card that only traded graded that week.
    const effectiveActual = isAllGraded
      ? Math.max(1, Math.round(avgActual / gradingMult))
      : avgActual;

    const estimated = +snap.estimated_price;
    if (!(effectiveActual > 0) || !(estimated > 0)) return;

    // error_pct = |effectiveActual − estimated| / effectiveActual
    // Stored as a fraction (e.g. 0.12 = 12%) with 5 decimal places.
    const errorPct = Math.abs(effectiveActual - estimated) / effectiveActual;

    // actual_date: the date of the last sale in the window (for reference).
    const actualDate = windowSales.reduce((latest, s) => s.date > latest ? s.date : latest, '');

    toResolve.push({
      id:                snap.id,
      actual_price:      avgActual,   // raw window average (pre-grading-adjustment)
      actual_date:       actualDate,
      error_pct:         Math.round(errorPct * 100000) / 100000,
      actual_afa_graded: isAllGraded,
      resolved_at:       now,
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
        actual_price:      row.actual_price,
        actual_date:       row.actual_date,
        error_pct:         row.error_pct,
        actual_afa_graded: row.actual_afa_graded,
        resolved_at:       row.resolved_at,
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

  // ── Expire snapshots older than 90 days that never resolved ────────────
  // These are cards that simply didn't trade in the measurement window.
  // Keeping them inflates total_snapshots and will never contribute MAE signal.
  const expiryDate = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString();
  try {
    const expireResp = UrlFetchApp.fetch(
      `${SB_URL_}/rest/v1/estimate_snapshots?actual_price=is.null&snapshot_at=lt.${expiryDate}`,
      {
        method: 'delete',
        headers: {
          apikey:        SB_KEY_,
          Authorization: `Bearer ${SB_KEY_}`,
          Prefer:        'return=representation',
        },
        muteHttpExceptions: true,
      }
    );
    const expired = JSON.parse(expireResp.getContentText() || '[]');
    if (expired.length) console.log(`  Expired ${expired.length} unresolved snapshots older than 90 days`);
  } catch(e) {
    console.warn(`  ⚠ Expiry step failed (non-fatal): ${e.message}`);
  }

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
    'select=engine_version,confidence,method,variant,error_pct,resolved_at' +
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
        byVariant:    {},
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

    const varnt = row.variant || 'Unknown';
    if (!g.byVariant[varnt]) g.byVariant[varnt] = [];
    g.byVariant[varnt].push(e);
  });

  const mean_ = arr => arr.length
    ? Math.round((arr.reduce((s, v) => s + v, 0) / arr.length) * 100000) / 100000
    : null;

  // ── PATCH each engine_params version ──────────────────────────────────
  Object.entries(byVersion).forEach(([version, data]) => {
    const maeByConf    = {};
    const maeByMethod  = {};
    const maeByVariant = {};

    Object.entries(data.byConfidence).forEach(([k, arr]) => { maeByConf[k]    = mean_(arr); });
    Object.entries(data.byMethod)    .forEach(([k, arr]) => { maeByMethod[k]  = mean_(arr); });
    Object.entries(data.byVariant)   .forEach(([k, arr]) => { maeByVariant[k] = mean_(arr); });

    const maeAll  = mean_(data.allErrors);
    const mae30d  = mean_(data.errors30d);
    const total   = totalByVersion[version] || data.resolvedCount;

    try {
      sbPatch_('engine_params', `version=eq.${version}`, {
        mae_all_time:       maeAll,
        mae_30d:            mae30d,
        mae_by_confidence:  maeByConf,
        mae_by_method:      maeByMethod,
        mae_by_variant:     maeByVariant,
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
// 3b.  computeGradingMultiplier()
//      Run on-demand (or scheduled monthly) to recalculate the AFA grading
//      premium from observed data and store it in engine_params.
//
//      Method:
//        1. Fetch all graded sales (afa_grade IS NOT NULL) from the sales table.
//        2. For each graded sale, find the ungraded avg_price for the same
//           player×variant combo in price_intelligence.
//        3. Compute ratio = graded_price / ungraded_avg for each pair.
//        4. Filter outliers (ratio < 0.5 or > 20).
//        5. Take the median ratio and store as afa_grade_multiplier on the
//           active engine_params row.
//
//      Prerequisite: Run this SQL in Supabase first:
//        ALTER TABLE engine_params
//          ADD COLUMN IF NOT EXISTS afa_grade_multiplier NUMERIC DEFAULT 2.0;
//        ALTER TABLE estimate_snapshots
//          ADD COLUMN IF NOT EXISTS actual_afa_graded BOOLEAN DEFAULT FALSE;
// ════════════════════════════════════════════════════════════════════════════

function computeGradingMultiplier() {
  console.log('▶ computeGradingMultiplier started');

  // ── Fetch all graded sales ─────────────────────────────────────────────
  const gradedSales = sbGet_('sales',
    'select=set_id,player,variant,price' +
    '&afa_grade=not.is.null&status=neq.invalid'
  );
  console.log(`  Found ${gradedSales.length} graded sale records`);

  if (!gradedSales.length) {
    console.log('  No graded sales found — multiplier unchanged.');
    return;
  }

  // ── Fetch price_intelligence for ungraded baselines ────────────────────
  // We use avg_price as the denominator — it represents the all-time ungraded
  // average (we'll normalize it below if needed, but even the mixed avg is a
  // reasonable denominator since graded cards are rare relative to total sales).
  const piRaw = sbGet_('price_intelligence',
    'select=set_id,player,variant,avg_price&avg_price=gt.0'
  );
  const piMap = {};
  piRaw.forEach(r => {
    piMap[`${r.set_id}|||${r.player}|||${r.variant}`] = +r.avg_price;
  });

  // ── Compute ratio for each graded sale ─────────────────────────────────
  const ratios = [];
  gradedSales.forEach(s => {
    if (!s.player || !s.variant || !(+s.price > 0)) return;
    const ungradedAvg = piMap[`${s.set_id}|||${s.player}|||${s.variant}`];
    if (!ungradedAvg || ungradedAvg <= 0) return;
    const ratio = +s.price / ungradedAvg;
    if (ratio < 0.5 || ratio > 20) return; // filter obvious outliers
    ratios.push(ratio);
  });
  console.log(`  Valid ratio pairs: ${ratios.length}`);

  if (!ratios.length) {
    console.log('  No valid ratio pairs found — multiplier unchanged.');
    return;
  }

  // ── Median ratio ───────────────────────────────────────────────────────
  ratios.sort((a, b) => a - b);
  const mid    = Math.floor(ratios.length / 2);
  const median = ratios.length % 2 === 1
    ? ratios[mid]
    : (ratios[mid - 1] + ratios[mid]) / 2;
  const multiplier = Math.round(median * 100) / 100;

  console.log(
    `  Ratios: min=${ratios[0].toFixed(2)}, ` +
    `p25=${ratios[Math.floor(ratios.length * 0.25)].toFixed(2)}, ` +
    `median=${multiplier.toFixed(2)}, ` +
    `p75=${ratios[Math.floor(ratios.length * 0.75)].toFixed(2)}, ` +
    `max=${ratios[ratios.length - 1].toFixed(2)}`
  );

  // ── Store on active engine_params row ─────────────────────────────────
  sbPatch_('engine_params', 'is_active=eq.true', {
    afa_grade_multiplier: multiplier,
  });
  console.log(`✓ computeGradingMultiplier complete — stored ${multiplier}× grading premium`);
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
    'select=id,set_id,player,variant,population,real_population' +
    '&order=set_id.asc'
  );

  const setMath = {};
  mathRaw.forEach(row => {
    const setName = setIdToName_[row.set_id];
    if (!setName) return;
    if (!setMath[setName]) setMath[setName] = [];
    setMath[setName].push({
      id:              row.id,
      player:          row.player,
      variant:         row.variant,
      population:      row.population,
      real_population: row.real_population,
    });
  });

  // ── Grading premium normalization ──────────────────────────────────────
  // AFA-graded sales inflate price_intelligence averages. Strip their
  // contribution using buildGradedNormMap + applyGradingNorm (estimate-engine.gs).
  // allProducts and productsBySet share the same object references, so
  // normalising allProducts in-place also normalises productsBySet automatically.
  // Runs after setMath is built so we can look up variant populations for the
  // population-aware grading curve.
  try {
    // Build "setName|||variant" → population lookup from setMath
    const gradingPopLookup = {};
    Object.entries(setMath).forEach(([sn, rows]) => {
      rows.forEach(r => {
        if (r.variant && r.population != null)
          gradingPopLookup[`${sn}|||${r.variant}`] = +r.population;
      });
    });

    const gradedSalesNorm = sbGet_('sales',
      'select=player,variant,price,date' +
      '&afa_grade=not.is.null&status=neq.invalid'
    );
    if (gradedSalesNorm.length > 0) {
      const gradedNormMap = buildGradedNormMap(gradedSalesNorm);
      applyGradingNorm(allProducts, gradedNormMap, gradingPopLookup);
      console.log(
        `  Grading norm applied — ${gradedSalesNorm.length} graded sales across ` +
        `${Object.keys(gradedNormMap).length} player×variant combos`
      );
    }
  } catch (normErr) {
    console.log('  Grading norm skipped (non-fatal): ' + normErr.message);
  }

  // ── price_signals — implied_price fallback + listing signals ──────────
  // Non-fatal: signals are fallback context only. If Supabase is unreachable
  // (e.g. free-tier waking from sleep) estimation still runs with empty signalMap.
  let signalMap = {};
  try {
    const signalsRaw = sbGet_('price_signals',
      'select=set_name,player,variant,' +
      'implied_price,implied_confidence,' +
      'lowest_ask,median_ask,supply_count,' +
      'avg_30d,sales_30d,last_sale_price,last_sale_date'
    );
    signalsRaw.forEach(row => {
      if (!row.set_name || !row.player || !row.variant) return;
      signalMap[`${row.set_name}|||${row.player}|||${row.variant}`] = row;
    });
    console.log(`  Loaded ${Object.keys(signalMap).length} price_signals rows`);
  } catch(e) {
    console.log('  price_signals load failed (non-fatal): ' + e.message);
  }

  console.log(
    `  Loaded: ${setsRaw.length} sets, ${allProducts.length} products, ` +
    `${mathRaw.length} set_math rows, ${Object.keys(signalMap).length} signals`
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
 *   refreshLiveEstimates   →  every day 4–5 AM
 *   resolveSnapshots       →  every day 6–7 AM
 */
function createSnapshotTriggers() {
  const MANAGED = ['takeEstimateSnapshots', 'refreshLiveEstimates', 'resolveSnapshots'];

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

  // Daily live estimates refresh: 4–5 AM
  ScriptApp.newTrigger('refreshLiveEstimates')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  // Daily resolve: 6–7 AM
  ScriptApp.newTrigger('resolveSnapshots')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  console.log('✓ Triggers registered:');
  console.log('  takeEstimateSnapshots — every Monday 3–4 AM');
  console.log('  refreshLiveEstimates  — every day 4–5 AM');
  console.log('  resolveSnapshots      — every day 6–7 AM');
}


/** List active snapshot-related triggers. */
function listSnapshotTriggers() {
  const fns = ['takeEstimateSnapshots', 'refreshLiveEstimates', 'resolveSnapshots', 'updateEngineAccuracy'];
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

// ═══════════════════════════════════════════════════════════════════════════════
// SYNC SALES SHEET → SUPABASE
// ─────────────────────────────────────────────────────────────────────────────
// Reads ALL valid rows from the SALES sheet and inserts into SB sales table.
// Assumes the table has been truncated first (TRUNCATE TABLE sales).
//
// SALES sheet column layout (1-based):
//   A(1)=Image   B(2)=Title    C(3)=URL      J(10)=Date    L(12)=Dupe
//   M(13)=Price  N(14)=Player  O(15)=Variant  P(16)=Set    Q(17)=Status
//   R(18)=Notes  S(19)=Flag    T(20)=Source   U(21)=AFA
//   V(22)=Multi-Sale ID        W(23)=Set Math ID
// ═══════════════════════════════════════════════════════════════════════════════
function syncSalesToSupabase() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('SALES');

  if (!sheet) { ui.alert('SALES sheet not found.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows in SALES sheet.'); return; }

  const confirm = ui.alert(
    'Sync SALES → Supabase',
    'This will insert all valid SALES rows into the SB sales table.\n\n' +
    'Make sure you have already run TRUNCATE TABLE sales in Supabase.\n\n' +
    'Proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Build set name → UUID cache
  let setIdCache;
  try {
    setIdCache = getSetIdCache_();
  } catch(e) {
    ui.alert('Failed to connect to Supabase:\n\n' + e.message);
    return;
  }

  // Read all 23 columns in one batch
  const data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();

  const records = [];
  let skippedNoTitle = 0, skippedDupe = 0, skippedInvalid = 0;

  data.forEach(function(row, i) {
    const rowNum  = i + 2;
    const image   = String(row[0]  || '').trim();  // A
    const title   = String(row[1]  || '').trim();  // B
    const url     = String(row[2]  || '').trim();  // C
    const date    = row[9];                         // J
    const dupe    = String(row[11] || '').trim();  // L
    const price   = row[12];                        // M
    const player  = String(row[13] || '').trim();  // N
    const variant = String(row[14] || '').trim();  // O
    const setName = String(row[15] || '').trim();  // P
    const status  = String(row[16] || '').trim();  // Q
    const notes   = String(row[17] || '').trim();  // R
    const flagged = String(row[18] || '').trim();  // S
    const source  = String(row[19] || 'eBay').trim(); // T
    const afa     = String(row[20] || '').trim();  // U
    const multiId = row[21];                        // V
    const smId    = String(row[22] || '').trim();  // W

    if (!title)             { skippedNoTitle++; return; }
    if (dupe === 'dupe')    { skippedDupe++;    return; }
    if (status !== 'valid') { skippedInvalid++; return; }

    const setId = setName ? (setIdCache[setName] || null) : null;

    let saleDate = null;
    if (date) {
      const d = new Date(date);
      if (!isNaN(d.getTime())) saleDate = d.toISOString().split('T')[0];
    }

    const multiSaleId = (multiId && !isNaN(+multiId)) ? +multiId : null;

    records.push({
      set_id:           setId,
      player:           player  || null,
      variant:          variant || null,
      title:            title.substring(0, 500),
      price:            Number(price) || null,
      date:             saleDate,
      image_url:        image  || null,
      ebay_url:         url    || null,
      afa_grade:        afa    || null,
      source:           source || 'eBay',
      status:           status,
      notes:            notes  || null,
      flagged:          flagged === 'FLAG',
      sheets_row_index: rowNum,
      multi_sale_id:    multiSaleId,
      set_math_id:      smId   || null,
    });
  });

  if (!records.length) {
    ui.alert(
      'No valid rows found to sync.\n\n' +
      'No title: ' + skippedNoTitle + '\n' +
      'Dupe flag: ' + skippedDupe + '\n' +
      'Invalid: '   + skippedInvalid
    );
    return;
  }

  Logger.log('syncSalesToSupabase: syncing %s records', records.length);

  const { url: sbUrl, key: sbKey } = getSupabaseConfig_();
  const headers = {
    'apikey':        sbKey,
    'Authorization': 'Bearer ' + sbKey,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };

  const CHUNK = 200;
  let inserted = 0, errors = 0;

  for (let i = 0; i < records.length; i += CHUNK) {
    const chunk = records.slice(i, i + CHUNK);
    try {
      const resp = UrlFetchApp.fetch(sbUrl + '/rest/v1/sales', {
        method: 'POST',
        headers,
        payload: JSON.stringify(chunk),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code === 200 || code === 201) {
        inserted += chunk.length;
        Logger.log('Inserted rows %s–%s of %s', i + 1, Math.min(i + CHUNK, records.length), records.length);
      } else {
        Logger.log('Insert error %s: %s', code, resp.getContentText().substring(0, 300));
        errors += chunk.length;
      }
    } catch(e) {
      Logger.log('Insert exception: ' + e.message);
      errors += chunk.length;
    }
  }

  ui.alert(
    '✅ Sync complete!\n\n' +
    'Rows inserted: '       + inserted + '\n' +
    'Errors: '              + errors   + '\n' +
    'Skipped (no title): '  + skippedNoTitle + '\n' +
    'Skipped (dupe flag): ' + skippedDupe    + '\n' +
    'Skipped (invalid): '   + skippedInvalid +
    (errors ? '\n\nCheck Apps Script Logs for error details.' : '')
  );
}
// ── Config ────────────────────────────────────────────────────────────────
const SB_URL  = 'https://sabzbyuqrondoayhwoth.supabase.co';
const SB_ANON = 'sb_publishable_KAbUO5YvOapq_kWkHqt7BA_UTwrP60s';
const SHEET_NAME = 'estimate_snapshots';   // tab name to write into
const TABLE      = 'estimate_snapshots';   // Supabase table name
const PAGE_SIZE  = 1000;                   // Supabase max per request

// ── Pull price_intelligence into a sheet ─────────────────────────────────
function pullPriceIntelligence() {
  const SHEET = 'price_intelligence';
  const TBL   = 'price_intelligence';

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET);
  if (!sheet) sheet = ss.insertSheet(SHEET);

  const rows = fetchAllRows_(TBL, 'set_id');
  if (!rows.length) {
    SpreadsheetApp.getUi().alert('No rows returned from ' + TBL);
    return;
  }

  const headers = Object.keys(rows[0]);
  const data    = rows.map(r => headers.map(h => r[h] ?? ''));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(`Done — ${rows.length.toLocaleString()} rows written to "${SHEET}"`);
}

// Paginated fetch with configurable order column
function fetchAllRows_(table, orderCol) {
  const headers = {
    'apikey':        SB_ANON,
    'Authorization': 'Bearer ' + SB_ANON,
  };
  let all    = [];
  let offset = 0;
  while (true) {
    const url = `${SB_URL}/rest/v1/${table}?select=*&order=${orderCol}.asc&limit=${PAGE_SIZE}&offset=${offset}`;
    const res = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      throw new Error(`Supabase error ${res.getResponseCode()}: ${res.getContentText()}`);
    }
    const page = JSON.parse(res.getContentText());
    all = all.concat(page);
    if (page.length < PAGE_SIZE) break;
    offset += PAGE_SIZE;
  }
  return all;
}

// ── Main ──────────────────────────────────────────────────────────────────
function pullEstimateSnapshots() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const rows = fetchAllRows(TABLE);
  if (!rows.length) {
    SpreadsheetApp.getUi().alert('No rows returned from ' + TABLE);
    return;
  }

  // Write headers from first row's keys, then data
  const headers = Object.keys(rows[0]);
  const data    = rows.map(r => headers.map(h => r[h] ?? ''));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  // Light formatting: freeze header row, bold it
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(`Done — ${rows.length.toLocaleString()} rows written to "${SHEET_NAME}"`);
}

// ── Paginated fetch (handles tables > 1,000 rows) ─────────────────────────
function fetchAllRows(table) {
  const headers = {
    'apikey':        SB_ANON,
    'Authorization': 'Bearer ' + SB_ANON,
  };

  let all    = [];
  let offset = 0;

  while (true) {
    const url = `${SB_URL}/rest/v1/${table}?select=*&order=id.asc&limit=${PAGE_SIZE}&offset=${offset}`;
    const res = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });

    if (res.getResponseCode() !== 200) {
      throw new Error(`Supabase error ${res.getResponseCode()}: ${res.getContentText()}`);
    }

    const page = JSON.parse(res.getContentText());
    all = all.concat(page);

    if (page.length < PAGE_SIZE) break;   // last page
    offset += PAGE_SIZE;
  }

  return all;
}


// ════════════════════════════════════════════════════════════════════════════
// WALK-FORWARD BACKTESTER
// ════════════════════════════════════════════════════════════════════════════
//
// Completely isolated from production. Writes to backtest_snapshots only.
// Does NOT touch estimate_snapshots, engine_params, or any prod table.
//
// Usage (run from Apps Script editor):
//   runWalkForward('2025-09-22', '2026-03-15')
//
// After it completes, run the SQL resolution query in Supabase to populate
// actual_price / error_pct, then query backtest_snapshots for MAE analysis.
//
// If you hit the 6-min GAS execution limit, split into date ranges:
//   runWalkForward('2025-09-22', '2025-12-01')
//   runWalkForward('2025-12-01', '2026-03-15')
// Each call gets its own run_id so results are separately queryable.
// ════════════════════════════════════════════════════════════════════════════

/**
 * Main entry point. Iterates weekly from startDateStr to endDateStr,
 * building synthetic price_intelligence at each cutoff and running
 * buildEstimates() against it. Results go to backtest_snapshots.
 */
function runWalkForward(startDateStr, endDateStr) {
  startDateStr = startDateStr || '2025-09-22';
  // Default end: 7 days ago so the last snapshot has a full resolution window
  endDateStr   = endDateStr   || new Date(Date.now() - 7 * 86400000).toISOString().slice(0, 10);

  const runId = 'wf_' + startDateStr.replace(/-/g, '') + '_' + endDateStr.replace(/-/g, '');
  console.log(`▶ runWalkForward  run_id=${runId}  ${startDateStr} → ${endDateStr}`);

  // ── Load time-invariant data once ────────────────────────────────────────
  const setsRaw = sbGet_('sets',
    'select=id,name,league&active=eq.true&order=name.asc'
  );
  const setIdToName  = {};
  const setLeagueMap = {};
  setsRaw.forEach(s => {
    setIdToName[s.id]    = s.name;
    setLeagueMap[s.name] = s.league || '';
  });

  // set_math — populations (current values; stable enough for backtesting)
  const mathRaw = sbGet_('set_math',
    'select=set_id,player,variant,population,real_population&order=set_id.asc'
  );
  const setMath = {};
  mathRaw.forEach(row => {
    const setName = setIdToName[row.set_id];
    if (!setName) return;
    if (!setMath[setName]) setMath[setName] = [];
    setMath[setName].push({
      player:          row.player,
      variant:         row.variant,
      population:      row.population,
      real_population: row.real_population,
    });
  });

  // price_signals — use current values (historical not available; minor approximation)
  let signalMap = {};
  try {
    const signalsRaw = sbGet_('price_signals',
      'select=set_name,player,variant,' +
      'implied_price,implied_confidence,' +
      'lowest_ask,median_ask,supply_count,' +
      'avg_30d,sales_30d,last_sale_price,last_sale_date'
    );
    signalsRaw.forEach(row => {
      if (!row.set_name || !row.player || !row.variant) return;
      signalMap[`${row.set_name}|||${row.player}|||${row.variant}`] = row;
    });
    console.log(`  Loaded ${Object.keys(signalMap).length} price_signals rows`);
  } catch(e) {
    console.log('  price_signals load skipped (non-fatal): ' + e.message);
  }

  const engineParams = loadActiveEngineParams_();
  console.log(`  Engine v${engineParams.version}`);

  // Player priors
  const playerPriors = {};
  try {
    const playersRaw = sbGet_('players', 'select=name,prior_tier');
    (playersRaw || []).forEach(p => {
      if (p.name && p.prior_tier) playerPriors[p.name] = { tier: +p.prior_tier };
    });
    console.log(`  Loaded ${Object.keys(playerPriors).length} player priors`);
  } catch(e) {
    console.log('  players load skipped (non-fatal): ' + e.message);
  }

  // ── Fetch ALL sales once — filter per cutoff in memory ──────────────────
  // One bulk fetch is far faster than N weekly queries.
  console.log('  Loading all sales...');
  var allSalesRaw = [], _off = 0;
  while (true) {
    var _batch = sbGet_('sales',
      'select=set_id,player,variant,date,price,afa_grade' +
      '&status=neq.invalid' +
      '&order=date.asc&offset=' + _off + '&limit=1000');
    allSalesRaw = allSalesRaw.concat(_batch);
    if (_batch.length < 1000) break;
    _off += 1000;
  }
  // Pre-resolve set_id → set_name and drop rows we can't use
  const allSales = allSalesRaw
    .map(s => ({ ...s, setName: setIdToName[s.set_id] }))
    .filter(s => s.setName && s.player && s.variant && s.date && +s.price > 0);
  console.log(`  ${allSales.length} usable sales loaded`);

  // ── Walk forward weekly ──────────────────────────────────────────────────
  const start   = new Date(startDateStr + 'T00:00:00Z');
  const end     = new Date(endDateStr   + 'T00:00:00Z');
  let   current = new Date(start);
  let   weekNum = 0;
  const allRows = [];

  while (current <= end) {
    weekNum++;
    const cutoff = current.toISOString().slice(0, 10); // 'YYYY-MM-DD'

    // Sales strictly before this cutoff date
    const salesBefore = allSales.filter(s => s.date < cutoff);
    if (!salesBefore.length) {
      console.log(`  Week ${weekNum} (${cutoff}): no sales yet — skipping`);
      current.setUTCDate(current.getUTCDate() + 7);
      continue;
    }

    // Build synthetic PI from filtered sales
    const { products, productsBySet } = buildBacktestPI_(salesBefore, cutoff);

    // Apply grading normalization using only graded sales before cutoff.
    // Uses inline backtest versions (cutoff-relative windows, self-contained).
    const gradedBefore = salesBefore.filter(s => s.afa_grade);
    if (gradedBefore.length > 0) {
      const normMap = buildBacktestGradedNormMap_(gradedBefore, cutoff);
      applyBacktestGradingNorm_(products, normMap);
    }

    // Swap the global allData so buildEstimates reads synthetic data.
    // allData is a var-declared global in this file — safe to reassign.
    allData = { products, productsBySet, setMath, setLeagueMap, signalMap, playerPriors };

    // Set freshness reference to snapshot date so computeFreshness measures
    // staleness relative to the simulated "now", not the real current date.
    FRESHNESS_REFERENCE_DATE = cutoff;

    let weekRows = 0;
    for (const set of setsRaw) {
      const setName  = set.name;
      const setProds = productsBySet[setName] || [];
      const mathRows = setMath[setName]       || [];
      if (!mathRows.length) continue;

      let result;
      try {
        result = buildEstimates(setName, setProds, engineParams);
      } catch(e) {
        console.error(`  ✗ buildEstimates: ${setName} @ ${cutoff} — ${e.message}`);
        continue;
      }

      const { estimates, refreshedEstimates } = result;
      const validCombos = new Set(mathRows.map(r => `${r.player}|||${r.variant}`));

      const push_ = (key, est, isRefreshed) => {
        if (!validCombos.has(key) || !(est.price > 0)) return;
        const [player, variant] = key.split('|||');
        allRows.push(buildBacktestRow_({
          runId, cutoff, setName, player, variant, est,
          engineVersion: engineParams.version,
          isRefreshed,
        }));
        weekRows++;
      };

      Object.entries(estimates).forEach(([k, e])          => push_(k, e, false));
      Object.entries(refreshedEstimates).forEach(([k, e]) => push_(k, e, true));
    }

    console.log(`  Week ${weekNum} (${cutoff}): ${weekRows} rows`);
    current.setUTCDate(current.getUTCDate() + 7);
  }

  console.log(`  Total: ${allRows.length} rows across ${weekNum} weeks — writing...`);
  sbInsertChunked_('backtest_snapshots', allRows, SNAP_CHUNK_SIZE_);

  FRESHNESS_REFERENCE_DATE = null; // Reset to real time
  console.log(`✓ runWalkForward complete — run_id: ${runId}`);
  return runId;
}


/**
 * Debug a single player+variant at a specific cutoff date.
 * Runs one week of the backtester with full engine debug logging.
 * Does NOT write to Supabase — just logs to the GAS execution log.
 *
 * Usage:  debugEstimate('Zaccharie Risacher', 'Gold', '2025-09-26')
 */
function debugEstimate(playerName, variantName, cutoffDate) {
  // Enable debug tracing
  ENGINE_DEBUG_TARGET = { player: playerName, variant: variantName };

  const setsRaw = sbGet_('sets', 'select=id,name,league&active=eq.true&order=name.asc');
  const setIdToName = {}, setLeagueMap = {};
  setsRaw.forEach(s => { setIdToName[s.id] = s.name; setLeagueMap[s.name] = s.league || ''; });

  const mathRaw = sbGet_('set_math', 'select=set_id,player,variant,population,real_population&order=set_id.asc');
  const setMath = {};
  mathRaw.forEach(row => {
    const setName = setIdToName[row.set_id];
    if (!setName) return;
    if (!setMath[setName]) setMath[setName] = [];
    setMath[setName].push({ player: row.player, variant: row.variant, population: row.population, real_population: row.real_population });
  });

  let signalMap = {};
  const engineParams = loadActiveEngineParams_();

  console.log(`▶ debugEstimate: ${playerName} ${variantName} @ ${cutoffDate}`);
  console.log(`  Engine v${engineParams.version}`);

  // Load all sales before cutoff (paginate past 1000-row limit)
  var allSalesRaw = [], _off2 = 0;
  while (true) {
    var _b2 = sbGet_('sales', 'select=set_id,player,variant,date,price,afa_grade&status=neq.invalid&order=date.asc&offset=' + _off2 + '&limit=1000');
    allSalesRaw = allSalesRaw.concat(_b2);
    if (_b2.length < 1000) break;
    _off2 += 1000;
  }
  const allSales = allSalesRaw
    .map(s => ({ ...s, setName: setIdToName[s.set_id] }))
    .filter(s => s.setName && s.player && s.variant && s.date && +s.price > 0);

  const salesBefore = allSales.filter(s => s.date < cutoffDate);
  console.log(`  ${salesBefore.length} sales before ${cutoffDate}`);

  const { products, productsBySet } = buildBacktestPI_(salesBefore, cutoffDate);

  // Grading normalization
  const gradedBefore = salesBefore.filter(s => s.afa_grade);
  if (gradedBefore.length > 0) {
    const normMap = buildBacktestGradedNormMap_(gradedBefore, cutoffDate);
    applyBacktestGradingNorm_(products, normMap);
  }

  // Player priors
  const playerPriors = {};
  try {
    const playersRaw = sbGet_('players', 'select=name,prior_tier');
    (playersRaw || []).forEach(p => {
      if (p.name && p.prior_tier) playerPriors[p.name] = { tier: +p.prior_tier };
    });
  } catch(e) { /* non-fatal */ }

  allData = { products, productsBySet, setMath, setLeagueMap, signalMap, playerPriors };

  // Find which set this player is in
  let targetSet = null;
  for (const [setName, rows] of Object.entries(setMath)) {
    if (rows.some(r => r.player === playerName && r.variant === variantName)) {
      targetSet = setName;
      break;
    }
  }
  if (!targetSet) {
    console.log(`  ✗ ${playerName} ${variantName} not found in any set_math`);
    ENGINE_DEBUG_TARGET = null;
    return;
  }

  console.log(`  Set: ${targetSet}`);
  const setProds = productsBySet[targetSet] || [];
  console.log(`  Products in set at cutoff: ${setProds.length}`);

  try {
    const result = buildEstimates(targetSet, setProds, engineParams);
    const est = result.estimates[`${playerName}|||${variantName}`];
    const ref = result.refreshedEstimates[`${playerName}|||${variantName}`];
    if (est) console.log(`  ★ ESTIMATE: $${est.price} (${est.confidence}, ${est.rangeEstimate ? 'range=$'+est.rangeEstimate : 'no range'}, ${est.ratioEstimate ? 'ratio=$'+est.ratioEstimate : 'no ratio'})`);
    else if (ref) console.log(`  ★ REFRESHED: $${ref.price} (${ref.confidence})`);
    else console.log(`  ★ No estimate produced`);
  } catch(e) {
    console.error(`  ✗ buildEstimates failed: ${e.message}`);
  }

  ENGINE_DEBUG_TARGET = null;
  console.log(`✓ debugEstimate complete`);
}


/**
 * Reconstructs price_intelligence-style product objects from raw sales
 * filtered to a cutoff date. Mirrors the shape loadAllData_() produces.
 */
function buildBacktestPI_(sales, cutoff) {
  const d30 = new Date(cutoff + 'T00:00:00Z');
  const d90 = new Date(cutoff + 'T00:00:00Z');
  d30.setUTCDate(d30.getUTCDate() - 30);
  d90.setUTCDate(d90.getUTCDate() - 90);
  const cutoff30 = d30.toISOString().slice(0, 10);
  const cutoff90 = d90.toISOString().slice(0, 10);

  // Group by setName|||player|||variant
  const groups = {};
  sales.forEach(s => {
    const k = `${s.setName}|||${s.player}|||${s.variant}`;
    if (!groups[k]) groups[k] = [];
    groups[k].push(s);
  });

  const products     = [];
  const productsBySet = {};

  Object.entries(groups).forEach(([k, rows]) => {
    const [setName, player, variant] = k.split('|||');

    // Newest first for lastSaleDate / lastSalePrice
    rows.sort((a, b) => b.date.localeCompare(a.date));

    const prices   = rows.map(r => +r.price);
    const rows30   = rows.filter(r => r.date >= cutoff30);
    const rows90   = rows.filter(r => r.date >= cutoff90);

    const avg = arr => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;

    const product = {
      set:           setName,
      player,
      variant,
      avgPrice:      Math.round(avg(prices)                      * 100) / 100,
      avg30d:        Math.round(avg(rows30.map(r => +r.price))   * 100) / 100,
      avg90d:        Math.round(avg(rows90.map(r => +r.price))   * 100) / 100,
      totalSales:    rows.length,
      sales30d:      rows30.length,
      sales90d:      rows90.length,
      lastSaleDate:  rows[0].date,
      lastSalePrice: +rows[0].price,
    };

    products.push(product);
    if (!productsBySet[setName]) productsBySet[setName] = [];
    productsBySet[setName].push(product);
  });

  return { products, productsBySet };
}


/**
 * Builds one backtest_snapshots row from an estimate result.
 * Mirrors buildSnapshotRow_() but targets the backtest table.
 */
function buildBacktestRow_({ runId, cutoff, setName, player, variant,
                              est, engineVersion, isRefreshed }) {
  const refDist = est.refDist || null;
  return {
    backtest_run_id: runId,
    snapshot_date:   cutoff,
    set_name:        setName,
    player,
    variant,
    estimated_price: est.price,
    confidence:      est.confidence || null,
    method:          deriveMethod_(est),
    engine_version:  engineVersion,
    evidence_weight: typeof est.evidenceWeight === 'number' ? est.evidenceWeight / 100 : null,
    n_ref_points:    refDist?.nPoints    || null,
    ref_sets:        refDist?.refSets    ? JSON.stringify(refDist.refSets) : null,
    extrapolated:    refDist?.extrapolated  || false,
    extrap_dir:      refDist?.extrapolatedUp ? 'up'
                   : refDist?.extrapolated   ? 'down' : null,
    is_refreshed:    isRefreshed,
  };
}


/**
 * Backtest-specific version of buildGradedNormMap.
 * Uses cutoff date (not Date.now()) for 30d/90d windows — critical for
 * historical backtesting where the cutoff may be months in the past.
 */
function buildBacktestGradedNormMap_(gradedSales, cutoff) {
  if (!gradedSales || !gradedSales.length) return {};
  const base = new Date(cutoff + 'T00:00:00Z');
  const d30  = new Date(base); d30.setUTCDate(d30.getUTCDate() - 30);
  const d90  = new Date(base); d90.setUTCDate(d90.getUTCDate() - 90);
  const c30  = d30.toISOString().slice(0, 10);
  const c90  = d90.toISOString().slice(0, 10);

  const map = {};
  gradedSales.forEach(s => {
    if (!s.player || !s.variant) return;
    const k = `${s.player}|||${s.variant}`;
    if (!map[k]) map[k] = { sum: 0, count: 0, sum30d: 0, count30d: 0, sum90d: 0, count90d: 0 };
    const g     = map[k];
    const price = +(s.price || 0);
    if (price <= 0) return;
    const dateStr = (s.date || '').slice(0, 10);
    g.sum   += price; g.count   += 1;
    if (dateStr >= c30) { g.sum30d += price; g.count30d += 1; }
    if (dateStr >= c90) { g.sum90d += price; g.count90d += 1; }
  });
  return map;
}


/**
 * Backtest-specific version of applyGradingNorm.
 * Logic is identical to the production version — inlined here so the
 * backtester has no dependency on estimate-engine.gs being in scope.
 */
function applyBacktestGradingNorm_(products, gradedNormMap) {
  if (!gradedNormMap || !products) return;
  products.forEach(prod => {
    const k = `${prod.player}|||${prod.variant}`;
    const g = gradedNormMap[k];
    if (!g || g.count === 0) return;

    if (prod.totalSales > 0) {
      if (g.count >= prod.totalSales) {
        prod.avgPrice = 0;
      } else if (prod.avgPrice > 0) {
        const uc = prod.totalSales - g.count;
        const us = prod.avgPrice * prod.totalSales - g.sum;
        prod.avgPrice   = Math.round(Math.max(0, us) / uc * 100) / 100;
        prod.totalSales = uc;
      }
    }
    if (g.count30d > 0 && prod.sales30d > 0) {
      if (g.count30d >= prod.sales30d) {
        prod.avg30d = 0;
      } else if (prod.avg30d > 0) {
        const uc = prod.sales30d - g.count30d;
        const us = prod.avg30d * prod.sales30d - g.sum30d;
        prod.avg30d   = Math.round(Math.max(0, us) / uc * 100) / 100;
        prod.sales30d = uc;
      }
    }
    if (g.count90d > 0 && prod.sales90d > 0) {
      if (g.count90d >= prod.sales90d) {
        prod.avg90d = 0;
      } else if (prod.avg90d > 0) {
        const uc = prod.sales90d - g.count90d;
        const us = prod.avg90d * prod.sales90d - g.sum90d;
        prod.avg90d   = Math.round(Math.max(0, us) / uc * 100) / 100;
        prod.sales90d = uc;
      }
    }
  });
}
function runBacktestPart1() {
  runWalkForward('2025-09-26', '2025-12-20');
}

function runBacktestPart2() {
  runWalkForward('2025-12-20', '2026-03-19');
}

function testDebugSkenes() {
  var playerName = 'Paul Skenes';
  var variantName = 'Victory';
  var cutoffDate = '2026-02-02';

  ENGINE_DEBUG_TARGET = { player: playerName, variant: variantName };

  var setsRaw = sbGet_('sets', 'select=id,name,league&active=eq.true&order=name.asc');
  var setIdToName = {}, setLeagueMap = {};
  setsRaw.forEach(function(s) { setIdToName[s.id] = s.name; setLeagueMap[s.name] = s.league || ''; });

  var mathRaw = sbGet_('set_math', 'select=set_id,player,variant,population,real_population&order=set_id.asc');
  var setMath = {};
  mathRaw.forEach(function(row) {
    var setName = setIdToName[row.set_id];
    if (!setName) return;
    if (!setMath[setName]) setMath[setName] = [];
    setMath[setName].push({ player: row.player, variant: row.variant, population: row.population, real_population: row.real_population });
  });

  var engineParams = loadActiveEngineParams_();
  console.log('Engine v' + engineParams.version);

  var allSalesRaw = [];
  var offset = 0;
  while (true) {
    var batch = sbGet_('sales', 'select=set_id,player,variant,date,price,afa_grade&status=neq.invalid&order=date.asc&offset=' + offset + '&limit=1000');
    allSalesRaw = allSalesRaw.concat(batch);
    if (batch.length < 1000) break;
    offset += 1000;
  }
  console.log('SALES LOADED: ' + allSalesRaw.length);
  var allSales = allSalesRaw
    .map(function(s) { return Object.assign({}, s, { setName: setIdToName[s.set_id] }); })
    .filter(function(s) { return s.setName && s.player && s.variant && s.date && +s.price > 0; });

  var salesBefore = allSales.filter(function(s) { return s.date < cutoffDate; });
  console.log(salesBefore.length + ' sales before ' + cutoffDate);

  var result = buildBacktestPI_(salesBefore, cutoffDate);
  var products = result.products;
  var productsBySet = result.productsBySet;

  var gradedBefore = salesBefore.filter(function(s) { return s.afa_grade; });
  if (gradedBefore.length > 0) {
    var normMap = buildBacktestGradedNormMap_(gradedBefore, cutoffDate);
    applyBacktestGradingNorm_(products, normMap);
  }

  var playerPriors = {};
  try {
    var playersRaw = sbGet_('players', 'select=name,prior_tier');
    (playersRaw || []).forEach(function(p) {
      if (p.name && p.prior_tier) playerPriors[p.name] = { tier: +p.prior_tier };
    });
  } catch(e) {}

  allData = { products: products, productsBySet: productsBySet, setMath: setMath, setLeagueMap: setLeagueMap, signalMap: {}, playerPriors: playerPriors };

  var targetSet = 'MLB 2025 Game Face';
  var setProds = productsBySet[targetSet] || [];
  console.log('Products in set at cutoff: ' + setProds.length);

  var est = buildEstimates(targetSet, setProds, engineParams);
  var gapFill = est.estimates[playerName + '|||' + variantName];
  var refreshed = est.refreshedEstimates[playerName + '|||' + variantName];

  if (gapFill) console.log('GAP-FILL ESTIMATE: $' + gapFill.price + ' (' + gapFill.confidence + ')');
  if (refreshed) console.log('REFRESHED ESTIMATE: $' + refreshed.price + ' (own=$' + refreshed.ownPrice + ' freshness=' + refreshed.freshness + '% model=$' + refreshed.modelEstimate + ' conf=' + refreshed.confidence + ')');
  if (!gapFill && !refreshed) console.log('No estimate produced');

  ENGINE_DEBUG_TARGET = null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// pushInventoryToSupabase
//
// Reads the inventory sheet, matches each row to a set_math_id in Supabase,
// then inserts all rows into ghostwrite_inventory.
//
// Matching logic:
//   Set   → sets where league = col G (Brand), year = col E, name contains col H
//   Player → set_math where player ≈ col F (normalized, fuzzy)
//   Variant → first word of col J  (e.g. "Victory /50" → "Victory")
//   Population → number after "/" in col J  (e.g. "Victory /50" → 50)
//
// Run once to seed the table. Truncate ghostwrite_inventory in Supabase and
// re-run to refresh.
// ═══════════════════════════════════════════════════════════════════════════════

const INVENTORY_SHEET_NAME_ = 'Inventory'; // ← set to your exact tab name

// Column indices (0-based) — matches the header row you provided
const INV_COL = {
  BUY_DATE:   0,   // A
  CHANNEL:    1,   // B
  BUY_PRICE:  2,   // C
  // D = ghost (skip)
  YEAR:       4,   // E
  PLAYER:     5,   // F
  BRAND:      6,   // G
  SET_LABEL:  7,   // H
  SIZE:       8,   // I
  VAR_RAW:    9,   // J
  // K-M = GCO/GRD/RG (skip)
  CNT:        14,  // O
  NOTE:       15,  // P
  GRADE_DATE: 16,  // Q
  GR_FEE:     17,  // R
  TOTAL_COST: 18,  // S
  GRADED:     19,  // T  (grading company, e.g. "AFA")
  GRADE:      20,  // U  (numeric grade)
  INV_NAME:   22,  // W  (main inventory label)
  SELL_DATE:  26,  // AA
  SELL_CHAN:   27,  // AB
  SELL_PRICE: 28,  // AC
  PL:         29,  // AD
};

function pushInventoryToSupabase() {
  const ui = SpreadsheetApp.getUi();

  // ── 1. Load reference data from Supabase ──────────────────────────────────
  ui.alert('Loading reference data from Supabase…\n(click OK to begin, then check logs for progress)');

  const sets    = sbGet_('sets',     'select=id,name,year,league');
  const setMath = fetchAllSetMath_();

  // Index set_math by set_id for fast lookup
  const smBySet = {};
  for (const row of setMath) {
    if (!smBySet[row.set_id]) smBySet[row.set_id] = [];
    smBySet[row.set_id].push(row);
  }

  // ── 2. Read inventory sheet ───────────────────────────────────────────────
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME_);
  if (!sheet) { ui.alert('Sheet "' + INVENTORY_SHEET_NAME_ + '" not found.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows found.'); return; }

  // Read enough columns to cover col AD (index 29 = col 30)
  const data = sheet.getRange(2, 1, lastRow - 1, 30).getValues();

  // ── 3. Process each row ───────────────────────────────────────────────────
  const toInsert   = [];
  const unmatched  = [];
  let   skipped    = 0;

  for (let i = 0; i < data.length; i++) {
    const r = data[i];

    // Skip blank rows
    const player = String(r[INV_COL.PLAYER] || '').trim();
    if (!player) { skipped++; continue; }

    const varRaw   = String(r[INV_COL.VAR_RAW] || '').trim();
    const brand    = String(r[INV_COL.BRAND]    || '').trim();
    const setLabel = String(r[INV_COL.SET_LABEL]|| '').trim();
    const year     = parseInt(r[INV_COL.YEAR], 10) || null;

    const { variant, population } = parseInventoryVar_(varRaw);

    // Match set
    const matchedSet = matchInventorySet_(brand, year, setLabel, sets);
    let setMathId    = null;
    let matchStatus  = 'unmatched-set';

    if (matchedSet) {
      const candidates = smBySet[matchedSet.id] || [];
      const smRow = matchInventorySetMath_(player, variant, population, candidates);
      if (smRow) {
        setMathId   = smRow.id;
        matchStatus = 'matched';
      } else {
        matchStatus = 'unmatched-card';
        unmatched.push('Row ' + (i + 2) + ': ' + player + ' / ' + variant + ' pop/' + population + ' in ' + matchedSet.name);
      }
    } else {
      unmatched.push('Row ' + (i + 2) + ': No set found for ' + brand + ' ' + year + ' ' + setLabel);
    }

    toInsert.push({
      set_math_id:    setMathId,
      buy_date:       inventoryFormatDate_(r[INV_COL.BUY_DATE]),
      channel:        String(r[INV_COL.CHANNEL]    || '').trim() || null,
      buy_price:      inventoryToNum_(r[INV_COL.BUY_PRICE]),
      year:           year,
      player:         player || null,
      brand:          brand  || null,
      set_label:      setLabel || null,
      size:           parseInt(r[INV_COL.SIZE], 10) || null,
      var_raw:        varRaw || null,
      variant:        variant || null,
      population:     population,
      cnt:            parseInt(r[INV_COL.CNT], 10) || 1,
      note:           String(r[INV_COL.NOTE]       || '').trim() || null,
      grade_date:     inventoryFormatDate_(r[INV_COL.GRADE_DATE]),
      gr_fee:         inventoryToNum_(r[INV_COL.GR_FEE]),
      total_cost:     inventoryToNum_(r[INV_COL.TOTAL_COST]),
      graded:         String(r[INV_COL.GRADED]     || '').trim() || null,
      grade:          inventoryToNum_(r[INV_COL.GRADE]),
      inventory_name: String(r[INV_COL.INV_NAME]   || '').trim() || null,
      sell_date:      inventoryFormatDate_(r[INV_COL.SELL_DATE]),
      sell_channel:   String(r[INV_COL.SELL_CHAN]  || '').trim() || null,
      sell_price:     inventoryToNum_(r[INV_COL.SELL_PRICE]),
      pl:             inventoryToNum_(r[INV_COL.PL]),
      match_status:   matchStatus,
    });
  }

  if (!toInsert.length) { ui.alert('No rows to insert (all blank?).'); return; }

  // ── 4. Insert in chunks of 100 ────────────────────────────────────────────
  const CHUNK = 100;
  const sbHeaders = {
    'apikey':        SB_ANON,
    'Authorization': 'Bearer ' + SB_ANON,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };

  for (let i = 0; i < toInsert.length; i += CHUNK) {
    const chunk = toInsert.slice(i, i + CHUNK);
    const res = UrlFetchApp.fetch(SB_URL + '/rest/v1/ghostwrite_inventory', {
      method:             'POST',
      headers:            sbHeaders,
      payload:            JSON.stringify(chunk),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 201) {
      throw new Error('Insert failed at chunk starting row ' + (i + 2) + ':\n' + res.getContentText());
    }
  }

  // ── 5. Summary ────────────────────────────────────────────────────────────
  const matched = toInsert.filter(r => r.match_status === 'matched').length;
  let msg = 'Done!\n\n'
    + toInsert.length + ' rows inserted\n'
    + matched         + ' matched to set_math_id\n'
    + unmatched.length + ' unmatched\n'
    + skipped          + ' blank rows skipped';
  if (unmatched.length) {
    msg += '\n\nFirst 15 unmatched:\n' + unmatched.slice(0, 15).join('\n');
  }
  ui.alert(msg);
}

// ── Inventory matching helpers ────────────────────────────────────────────────

function parseInventoryVar_(varRaw) {
  const s = (varRaw || '').trim();
  const variant    = s.split(/[\s\/]/)[0] || null;           // first word (handles "Base/800" and "Base /800")
  const slashMatch = s.match(/\/\s*(\d+)/);
  const population = slashMatch ? parseInt(slashMatch[1], 10) : null;
  return { variant, population };
}

// Normalize player/variant names for matching: lowercase, strip accents + punctuation
function normInventoryStr_(s) {
  return (s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/['\.\-]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Match set: brand → sets.league, year → sets.year, setLabel → sets.name contains
function matchInventorySet_(brand, year, setLabel, sets) {
  const b = (brand    || '').toUpperCase().trim();
  const y = parseInt(year, 10);
  const s = (setLabel || '').toLowerCase().trim();
  for (const set of sets) {
    const league  = (set.league || '').toUpperCase().trim();
    const setYear = parseInt(set.year, 10);
    const name    = (set.name   || '').toLowerCase();
    if (league === b && setYear === y && name.includes(s)) return set;
  }
  return null;
}

// Match set_math: player (fuzzy), variant (case-insensitive), population (exact)
function matchInventorySetMath_(player, variant, population, candidates) {
  const np  = normInventoryStr_(player);
  const nv  = (variant || '').toLowerCase().trim();
  const pop = population;

  // Pass 1: exact normalized player + variant + population
  let m = candidates.find(r =>
    normInventoryStr_(r.player) === np &&
    (r.variant || '').toLowerCase() === nv &&
    r.population === pop
  );
  if (m) return m;

  // Pass 2: fuzzy player (one name contains the other) + variant + population
  m = candidates.find(r => {
    const rp = normInventoryStr_(r.player);
    return (rp.includes(np) || np.includes(rp)) &&
      (r.variant || '').toLowerCase() === nv &&
      r.population === pop;
  });
  return m || null;
}

function inventoryFormatDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof val === 'string' && val.trim()) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return null;
}

function inventoryToNum_(val) {
  if (val === null || val === undefined || val === '') return null;
  const n = parseFloat(String(val).replace(/[$,\s]/g, ''));
  return isNaN(n) ? null : n;
}

function fetchAllSetMath_() {
  const headers = { 'apikey': SB_ANON, 'Authorization': 'Bearer ' + SB_ANON };
  let all = [], offset = 0;
  while (true) {
    const url = SB_URL + '/rest/v1/set_math?select=id,set_id,player,variant,population'
              + '&order=id.asc&limit=1000&offset=' + offset;
    const res = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('set_math fetch failed: ' + res.getContentText());
    const page = JSON.parse(res.getContentText());
    all = all.concat(page);
    if (page.length < 1000) break;
    offset += 1000;
  }
  return all;
}

