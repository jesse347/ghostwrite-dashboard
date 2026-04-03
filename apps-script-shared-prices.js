// ════════════════════════════════════════════════════════════════════════════
// apps-script-shared-prices.js — Ghostladder Shared Price Sheet
// Google Apps Script (V8 runtime)
//
// This script is designed for a SEPARATE Google Sheet shared with teammates.
// It pulls pricing data from the Ghostladder Supabase database and writes it
// to a clean, formatted sheet.
//
// ── Setup ────────────────────────────────────────────────────────────────────
// 1. Open the shared Google Sheet
// 2. Extensions → Apps Script
// 3. Paste this entire file into Code.gs (replace any existing content)
// 4. Save, then reload the spreadsheet
// 5. A "Ghostladder" menu will appear in the menu bar
// 6. Click "Ghostladder → Pull Prices & Estimates" to run
//
// ── Functions ────────────────────────────────────────────────────────────────
//   pullPricesAndEstimates()  — fetches PI + live_estimates, writes to sheet
//   onOpen()                  — adds the custom Ghostladder menu
// ════════════════════════════════════════════════════════════════════════════


// ── Supabase config ──────────────────────────────────────────────────────────
// Using the publishable (anon) key — read-only access to public tables.
const SB_URL = 'https://sabzbyuqrondoayhwoth.supabase.co';
const SB_KEY = 'sb_publishable_KAbUO5YvOapq_kWkHqt7BA_UTwrP60s';
const SB_HEADERS = {
  'apikey':        SB_KEY,
  'Authorization': 'Bearer ' + SB_KEY,
  'Content-Type':  'application/json',
};


// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ghostladder')
    .addItem('Pull Prices & Estimates', 'pullPricesAndEstimates')
    .addItem('Match GL References', 'matchGLReferences')
    .addSeparator()
    .addItem('Build Repack Checklist', 'buildRepackChecklist')
    .addToUi();
}


// ── Main function ────────────────────────────────────────────────────────────

function pullPricesAndEstimates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Fetch data from Supabase ────────────────────────────────────────
  ui.alert('⏳ Pulling data…', 'Fetching prices and estimates from the database. This may take a moment.', ui.ButtonSet.OK);

  const [piRows, estRows, setsRows] = [
    sbGetAll_('price_intelligence',
      'select=set_math_id,set_name,player,variant,' +
      'total_sales,avg_price,sales_30d,avg_30d,sales_90d,avg_90d,' +
      'last_sale_date,last_sale_price'
    ),
    sbGetAll_('live_estimates',
      'select=set_math_id,set_name,player,variant,' +
      'estimated_price,confidence,method,is_refreshed,' +
      'own_price,freshness,model_estimate,evidence_weight,' +
      'computed_at'
    ),
    sbGetAll_('sets', 'select=name,league,active&active=eq.true', 'name'),
  ];

  // ── 2. Build lookups ───────────────────────────────────────────────────
  // Estimate map: "set_name|||player|||variant" → row
  const estMap = {};
  estRows.forEach(r => {
    estMap[`${r.set_name}|||${r.player}|||${r.variant}`] = r;
  });

  // Set → league
  const leagueMap = {};
  setsRows.forEach(s => { leagueMap[s.name] = s.league || ''; });

  // ── 3. Merge PI + estimates into unified rows ──────────────────────────
  // Start with all PI rows (real sales data), enrich with estimates
  const merged = [];
  const seen = new Set();

  // First pass: all PI rows (have real sales data)
  piRows.forEach(pi => {
    const k = `${pi.set_name}|||${pi.player}|||${pi.variant}`;
    seen.add(k);
    const est = estMap[k] || {};
    merged.push({
      set_math_id:      pi.set_math_id,
      set_name:         pi.set_name,
      league:           leagueMap[pi.set_name] || '',
      player:           pi.player,
      variant:          pi.variant,
      // Sales data
      total_sales:      pi.total_sales || 0,
      avg_price:        pi.avg_price,
      sales_30d:        pi.sales_30d || 0,
      avg_30d:          pi.avg_30d,
      sales_90d:        pi.sales_90d || 0,
      avg_90d:          pi.avg_90d,
      last_sale_date:   pi.last_sale_date,
      last_sale_price:  pi.last_sale_price,
      // Estimate data
      estimated_price:  est.estimated_price ?? null,
      confidence:       est.confidence || '',
      method:           est.method || '',
      is_refreshed:     est.is_refreshed ?? null,
      // Best price: refreshed estimate (blended) > own sales avg > gap-fill estimate
      best_price:       bestPrice_(pi, est),
    });
  });

  // Second pass: estimate-only rows (no real sales yet — gap-fill)
  estRows.forEach(est => {
    const k = `${est.set_name}|||${est.player}|||${est.variant}`;
    if (seen.has(k)) return;
    seen.add(k);
    merged.push({
      set_math_id:      est.set_math_id,
      set_name:         est.set_name,
      league:           leagueMap[est.set_name] || '',
      player:           est.player,
      variant:          est.variant,
      total_sales:      0,
      avg_price:        null,
      sales_30d:        0,
      avg_30d:          null,
      sales_90d:        0,
      avg_90d:          null,
      last_sale_date:   null,
      last_sale_price:  null,
      estimated_price:  est.estimated_price,
      confidence:       est.confidence || '',
      method:           est.method || '',
      is_refreshed:     est.is_refreshed ?? false,
      best_price:       est.estimated_price,
    });
  });

  // ── 4. Sort: Set → Variant rarity → Best Price desc ────────────────────
  const variantOrder = {
    'Base': 0, 'Victory': 1, 'Vides': 2, 'Emerald': 3,
    'Gold': 4, 'Chrome': 5, 'Fire': 6, 'SP': 7,
    'Sapphire': 8, 'Bet on Women': 9,
  };
  merged.sort((a, b) => {
    if (a.set_name !== b.set_name) return a.set_name.localeCompare(b.set_name);
    const va = variantOrder[a.variant] ?? 50;
    const vb = variantOrder[b.variant] ?? 50;
    if (va !== vb) return va - vb;
    return (b.best_price || 0) - (a.best_price || 0);
  });

  // ── 5. Write to sheet ──────────────────────────────────────────────────
  const SHEET_NAME = 'Prices & Estimates';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // Header row
  const headers = [
    'Set Math ID',
    'Set',
    'League',
    'Player',
    'Variant',
    'Best Price',
    'Total Sales',
    'Avg Price (All-Time)',
    '30D Sales',
    'Avg Price (30D)',
    '90D Sales',
    'Avg Price (90D)',
    'Last Sale Date',
    'Last Sale Price',
    'Estimated Value',
    'Confidence',
    'Method',
    'Refreshed?',
  ];

  // Data rows
  const dataRows = merged.map(r => [
    r.set_math_id || '',
    r.set_name,
    r.league,
    r.player,
    r.variant,
    r.best_price,
    r.total_sales,
    r.avg_price,
    r.sales_30d,
    r.avg_30d,
    r.sales_90d,
    r.avg_90d,
    r.last_sale_date || '',
    r.last_sale_price,
    r.estimated_price,
    r.confidence,
    r.method,
    r.is_refreshed === true ? 'Yes' : r.is_refreshed === false ? 'No' : '',
  ]);

  const allRows = [headers, ...dataRows];
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);

  // ── 6. Format ──────────────────────────────────────────────────────────
  // Header formatting
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1e1e1e');
  headerRange.setFontColor('#e8ff47');
  headerRange.setHorizontalAlignment('center');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Number formats (currency columns)
  const currencyCols = [6, 8, 10, 12, 14, 15]; // Best Price, Avg All-Time, Avg 30D, Avg 90D, Last Sale, Estimated
  currencyCols.forEach(col => {
    if (dataRows.length > 0) {
      sheet.getRange(2, col, dataRows.length, 1).setNumberFormat('$#,##0');
    }
  });

  // Date column
  if (dataRows.length > 0) {
    sheet.getRange(2, 13, dataRows.length, 1).setNumberFormat('yyyy-mm-dd');
  }

  // Integer columns (sales counts)
  const intCols = [7, 9, 11]; // Total Sales, 30D Sales, 90D Sales
  intCols.forEach(col => {
    if (dataRows.length > 0) {
      sheet.getRange(2, col, dataRows.length, 1).setNumberFormat('#,##0');
    }
  });

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Hide Set Math ID column (useful for lookups but not for reading)
  sheet.hideColumns(1);

  // ── 7. Add metadata row at bottom ──────────────────────────────────────
  const metaRow = dataRows.length + 3;
  sheet.getRange(metaRow, 1).setValue('Last updated');
  sheet.getRange(metaRow, 2).setValue(new Date().toLocaleString());
  sheet.getRange(metaRow, 3).setValue(`${merged.length} rows`);
  sheet.getRange(metaRow, 1, 1, 3).setFontColor('#999999').setFontStyle('italic');

  ui.alert('✅ Done!', `Pulled ${merged.length} rows into "${SHEET_NAME}".\n\n` +
    `• ${piRows.length} with real sales data\n` +
    `• ${merged.length - piRows.length} estimate-only (no sales yet)`,
    ui.ButtonSet.OK);
}


// ════════════════════════════════════════════════════════════════════════════
// Build Repack Checklist
// ════════════════════════════════════════════════════════════════════════════
//
// Pulls inventory (unsold) + live estimates from Supabase, writes a
// checklist sheet with editable Price and Qty columns + live summary formulas.
//
// Adding new cards:
//   - Player, Variant, Set columns have dropdown validation from Prices & Estimates
//   - Suggested Price auto-populates via VLOOKUP from Prices & Estimates
//   - Price defaults to Suggested Price (override freely)
//   - Formulas extend 200 rows past inventory data so new rows "just work"
//
// Sheet layout:
//   A-I:   Data table (Player, Variant, Set, Grade, Suggested Price, Price, Qty, Line Value, Source)
//   K-N:   Summary dashboard (all formulas — updates live as you edit Price/Qty)
//
// To connect this sheet to the Repack Simulator page:
//   1. File → Share → Publish to web
//   2. Select the "Repack Checklist" sheet, format = CSV
//   3. Copy the URL and paste it into repack-sim.html's SHEET_CSV_URL constant
// ────────────────────────────────────────────────────────────────────────────

function buildRepackChecklist() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Check for existing data ────────────────────────────────────────────
  const SHEET_NAME = 'Repack Checklist';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    const resp = ui.alert('⚠️ Existing checklist',
      'The "Repack Checklist" sheet already has data.\n\n' +
      'Replace it with fresh inventory data? (Your edits will be lost.)',
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  ui.alert('⏳ Building checklist…',
    'Pulling inventory and estimates from the database.',
    ui.ButtonSet.OK);

  // ── 1. Fetch data ──────────────────────────────────────────────────────
  const [invRows, estRows, setsRows] = [
    sbGetAll_('ghostwrite_inventory',
      'select=set_math_id,player,variant,brand,year,set_label,grade,graded,total_cost' +
      '&sell_date=is.null', 'player'),
    sbGetAll_('live_estimates',
      'select=set_math_id,set_name,player,variant,estimated_price,confidence,is_refreshed'),
    sbGetAll_('sets', 'select=id,name&active=eq.true', 'name'),
  ];

  // ── 2. Build estimate lookup by set_math_id ────────────────────────────
  const estBySmId = {};
  estRows.forEach(r => {
    if (r.set_math_id) estBySmId[r.set_math_id] = r;
  });

  // ── 3. Grade multiplier table ──────────────────────────────────────────
  const GRADE_MULT = { 10: 3.0, 9.75: 2.25, 9.5: 1.85, 9.25: 1.65, 9: 1.5, 8.5: 1.25, 8: 1.05 };
  function gradeMult(grade) {
    const g = Math.round(+grade * 4) / 4;
    return GRADE_MULT[g] || 1.5;
  }

  // ── 4. Process inventory into checklist rows ───────────────────────────
  const aggMap = {};
  invRows.forEach(row => {
    const est      = row.set_math_id ? estBySmId[row.set_math_id] : null;
    const setName  = est ? est.set_name : [row.brand, row.year, row.set_label].filter(Boolean).join(' ');
    const player   = est ? est.player   : row.player;
    const variant  = est ? est.variant  : row.variant;
    const isGraded = row.grade != null && row.grade > 0;
    const gradeLabel = (row.graded && row.grade != null) ? `${row.graded} ${row.grade}` : '';

    const baseEst = est ? +est.estimated_price : 0;
    const suggestedPrice = baseEst > 0
      ? (isGraded ? Math.round(baseEst * gradeMult(row.grade) + 15) : Math.round(baseEst))
      : 0;

    const k = `${row.set_math_id || ''}|||${gradeLabel}`;
    if (!aggMap[k]) {
      aggMap[k] = { player, variant, setName, gradeLabel, suggestedPrice, qty: 0, totalCost: 0 };
    }
    aggMap[k].qty++;
    if (row.total_cost != null) aggMap[k].totalCost += +row.total_cost;
  });

  const rows = Object.values(aggMap);

  // Sort: Set → Variant rarity → Price desc
  const varOrd = { Base:0, Victory:1, Vides:2, Emerald:3, Gold:4, Chrome:5, Fire:6, SP:7, Sapphire:8 };
  rows.sort((a, b) => {
    if (a.setName !== b.setName) return a.setName.localeCompare(b.setName);
    const va = varOrd[a.variant] ?? 50;
    const vb = varOrd[b.variant] ?? 50;
    if (va !== vb) return va - vb;
    return (b.suggestedPrice || 0) - (a.suggestedPrice || 0);
  });

  // ── 5. Ensure Prices & Estimates sheet exists ──────────────────────────
  const peSheet = ss.getSheetByName('Prices & Estimates');
  if (!peSheet) {
    ui.alert('❌ Missing "Prices & Estimates"',
      'Run "Pull Prices & Estimates" first so we have data for lookups and dropdowns.',
      ui.ButtonSet.OK);
    return;
  }

  // ── 6. Create / clear sheet ────────────────────────────────────────────
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    // Remove existing protections
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  }

  // ── 7. Write data table ────────────────────────────────────────────────
  const headers = ['Player', 'Variant', 'Set', 'Grade', 'Suggested Price', 'Price', 'Qty', 'Line Value', 'Source'];
  const DATA_START = 2;
  const EXTRA_ROWS = 200; // empty rows for adding new cards
  const FORMULA_END = DATA_START + rows.length + EXTRA_ROWS - 1;

  // Header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Inventory data rows (A-D: identity, F: price = suggested, G: qty, I: source)
  if (rows.length > 0) {
    const dataRows = rows.map(r => [
      r.player,                      // A
      r.variant,                     // B
      r.setName,                     // C
      r.gradeLabel,                  // D
      null,                          // E: Suggested Price (formula below)
      r.suggestedPrice || '',        // F: Price (editable — starts at suggested)
      r.qty,                         // G: Qty (editable)
      null,                          // H: Line Value (formula below)
      'Inventory',                   // I: Source
    ]);
    sheet.getRange(DATA_START, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // ── 8. Suggested Price formula (col E) — VLOOKUP from Prices & Estimates ──
  // Looks up Player+Variant+Set → Best Price from the P&E sheet.
  // Works for inventory rows AND new rows added by the user.
  // P&E columns: D=Player, E=Variant, B=Set, F=Best Price
  const sugFormulas = [];
  for (let i = DATA_START; i <= FORMULA_END; i++) {
    sugFormulas.push([
      `=IFERROR(INDEX('Prices & Estimates'!F:F,MATCH(A${i}&B${i}&C${i},'Prices & Estimates'!D:D&'Prices & Estimates'!E:E&'Prices & Estimates'!B:B,0)),"")`
    ]);
  }
  sheet.getRange(DATA_START, 5, sugFormulas.length, 1).setFormulas(sugFormulas);

  // ── 9. Line Value formula (col H) = Price × Qty ───────────────────────
  const lineFormulas = [];
  for (let i = DATA_START; i <= FORMULA_END; i++) {
    lineFormulas.push([
      `=IF(AND(F${i}<>"",G${i}<>""),F${i}*G${i},"")`
    ]);
  }
  sheet.getRange(DATA_START, 8, lineFormulas.length, 1).setFormulas(lineFormulas);

  // ── 10. Data validation dropdowns (A, B, C) from Prices & Estimates ───
  const peData = peSheet.getDataRange().getValues();
  const peHeaders = peData[0].map(h => String(h).trim().toLowerCase());
  const pePlayerCol = peHeaders.indexOf('player');
  const peVariantCol = peHeaders.indexOf('variant');
  const peSetCol = peHeaders.indexOf('set');

  if (pePlayerCol >= 0 && peVariantCol >= 0 && peSetCol >= 0) {
    // Extract unique sorted values
    const uniquePlayers  = [...new Set(peData.slice(1).map(r => String(r[pePlayerCol]).trim()).filter(Boolean))].sort();
    const uniqueVariants = [...new Set(peData.slice(1).map(r => String(r[peVariantCol]).trim()).filter(Boolean))].sort();
    const uniqueSets     = [...new Set(peData.slice(1).map(r => String(r[peSetCol]).trim()).filter(Boolean))].sort();

    // Write validation lists to a hidden helper sheet
    const HELPER_NAME = '_Validation';
    let helperSheet = ss.getSheetByName(HELPER_NAME);
    if (!helperSheet) {
      helperSheet = ss.insertSheet(HELPER_NAME);
      helperSheet.hideSheet();
    } else {
      helperSheet.clearContents();
    }

    // Write lists vertically: col A = Players, col B = Variants, col C = Sets
    if (uniquePlayers.length)  helperSheet.getRange(1, 1, uniquePlayers.length, 1).setValues(uniquePlayers.map(v => [v]));
    if (uniqueVariants.length) helperSheet.getRange(1, 2, uniqueVariants.length, 1).setValues(uniqueVariants.map(v => [v]));
    if (uniqueSets.length)     helperSheet.getRange(1, 3, uniqueSets.length, 1).setValues(uniqueSets.map(v => [v]));

    // Apply data validation to the full editable range (inventory rows + extra rows)
    const valRange = FORMULA_END;

    // Player dropdown (col A)
    const playerRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helperSheet.getRange(1, 1, uniquePlayers.length, 1), true)
      .setAllowInvalid(true) // allow typing values not in the list
      .build();
    sheet.getRange(DATA_START, 1, valRange - DATA_START + 1, 1).setDataValidation(playerRule);

    // Variant dropdown (col B)
    const variantRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helperSheet.getRange(1, 2, uniqueVariants.length, 1), true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(DATA_START, 2, valRange - DATA_START + 1, 1).setDataValidation(variantRule);

    // Set dropdown (col C)
    const setRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helperSheet.getRange(1, 3, uniqueSets.length, 1), true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(DATA_START, 3, valRange - DATA_START + 1, 1).setDataValidation(setRule);
  }

  // ── 11. Format ─────────────────────────────────────────────────────────
  // Header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1e1e1e');
  headerRange.setFontColor('#e8ff47');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Number formats for the full formula range
  const fmtRows = FORMULA_END - DATA_START + 1;
  [5, 6, 8].forEach(col => {
    sheet.getRange(DATA_START, col, fmtRows, 1).setNumberFormat('$#,##0');
  });
  sheet.getRange(DATA_START, 7, fmtRows, 1).setNumberFormat('#,##0');

  // Editable highlight on Price (F) and Qty (G) columns
  if (rows.length > 0) {
    sheet.getRange(DATA_START, 6, rows.length, 2).setBackground('#1a1a00');
  }

  // Protect formula columns with warning only
  const sugProtection = sheet.getRange(DATA_START, 5, fmtRows, 1).protect();
  sugProtection.setDescription('Suggested Price (auto-lookup)');
  sugProtection.setWarningOnly(true);

  const lineProtection = sheet.getRange(DATA_START, 8, fmtRows, 1).protect();
  lineProtection.setDescription('Line Value (auto-calculated)');
  lineProtection.setWarningOnly(true);

  // Auto-resize
  for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);

  // ── 12. Summary dashboard (columns K-N) ────────────────────────────────
  const S = 11; // column K
  const rng = (r, c, h, w) => sheet.getRange(r, S + c - 1, h || 1, w || 1);

  // Title
  rng(1, 1).setValue('REPACK DASHBOARD').setFontWeight('bold').setFontColor('#e8ff47').setFontSize(12);

  // Key Stats (formulas reference the full range including extra rows)
  rng(3, 1).setValue('Total Cards');
  rng(3, 2).setFormula(`=SUM(G${DATA_START}:G${FORMULA_END})`).setFontWeight('bold').setNumberFormat('#,##0');

  rng(4, 1).setValue('Pool Value');
  rng(4, 2).setFormula(`=SUM(H${DATA_START}:H${FORMULA_END})`).setFontWeight('bold').setNumberFormat('$#,##0');

  rng(5, 1).setValue('Pack Price (EV)');
  rng(5, 2).setFormula(`=IFERROR(L4/L3,0)`).setFontWeight('bold').setNumberFormat('$#,##0');

  rng(6, 1).setValue('Winning Packs');
  rng(6, 2).setFormula(`=IFERROR(SUMPRODUCT((F${DATA_START}:F${FORMULA_END}>L5)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})/SUM(G${DATA_START}:G${FORMULA_END}),0)`).setFontWeight('bold').setNumberFormat('0%');

  // Style stats labels
  sheet.getRange(3, S, 4, 1).setFontColor('#999999');
  sheet.getRange(3, S + 1, 4, 1).setFontSize(14);

  // Pool Odds
  rng(8, 1).setValue('POOL ODDS').setFontWeight('bold').setFontColor('#e8ff47');

  const tiers = [
    { label: '$500+',      formula: `=SUMPRODUCT((F${DATA_START}:F${FORMULA_END}>=500)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})` },
    { label: '$250–$499',  formula: `=SUMPRODUCT((F${DATA_START}:F${FORMULA_END}>=250)*(F${DATA_START}:F${FORMULA_END}<500)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})` },
    { label: '$100–$249',  formula: `=SUMPRODUCT((F${DATA_START}:F${FORMULA_END}>=100)*(F${DATA_START}:F${FORMULA_END}<250)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})` },
    { label: '$25–$99',    formula: `=SUMPRODUCT((F${DATA_START}:F${FORMULA_END}>=25)*(F${DATA_START}:F${FORMULA_END}<100)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})` },
    { label: 'Under $25',  formula: `=SUMPRODUCT((F${DATA_START}:F${FORMULA_END}<25)*(F${DATA_START}:F${FORMULA_END}>0)*(G${DATA_START}:G${FORMULA_END}>0)*G${DATA_START}:G${FORMULA_END})` },
  ];

  tiers.forEach((t, i) => {
    const row = 9 + i;
    rng(row, 1).setValue(t.label).setFontColor('#999999');
    rng(row, 2).setFormula(t.formula).setNumberFormat('#,##0').setFontWeight('bold');
    rng(row, 3).setFormula(`=IFERROR(L${row}/L3,0)`).setNumberFormat('0%').setFontColor('#999999');
  });

  // Metadata
  rng(15, 1).setValue('Last built').setFontColor('#666666').setFontStyle('italic');
  rng(15, 2).setValue(new Date().toLocaleString()).setFontColor('#666666').setFontStyle('italic');
  rng(16, 1).setValue(`${rows.length} inventory items`).setFontColor('#666666').setFontStyle('italic');

  // Protect summary columns
  const summaryProtection = sheet.getRange(1, S, 20, 4).protect();
  summaryProtection.setDescription('Summary dashboard (auto-calculated)');
  summaryProtection.setWarningOnly(true);

  // Column widths
  sheet.setColumnWidth(S, 130);
  sheet.setColumnWidth(S + 1, 100);
  sheet.setColumnWidth(S + 2, 60);
  sheet.setColumnWidth(10, 30);

  ui.alert('✅ Checklist built!',
    `${rows.length} inventory items loaded into "${SHEET_NAME}".\n\n` +
    '• Edit Price (col F) and Qty (col G) freely\n' +
    '• To add new cards: type in any empty row — Player, Variant, and Set have dropdowns\n' +
    '• Suggested Price auto-populates from Prices & Estimates\n' +
    '• Summary stats update automatically',
    ui.ButtonSet.OK);
}


// ── Match GL References ──────────────────────────────────────────────────────
//
// Reads the "Prices & Estimates" sheet and the currently active sheet,
// matches rows on Player + Set + Variant (Hit), and writes Best Price
// into the "GL Reference" column.
//
// The active sheet must have columns with these headers (in any order):
//   "Player Name"   — matches → "Player" in Prices & Estimates
//   "Set"           — matches → "Set" in Prices & Estimates
//   "Hit"           — matches → "Variant" in Prices & Estimates
//   "GL Reference"  — where the price gets written
// ─────────────────────────────────────────────────────────────────────────────

function matchGLReferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── 1. Read Prices & Estimates sheet ───────────────────────────────────
  const peSheet = ss.getSheetByName('Prices & Estimates');
  if (!peSheet) {
    ui.alert('❌ Error', '"Prices & Estimates" sheet not found.\n\nRun "Pull Prices & Estimates" first.', ui.ButtonSet.OK);
    return;
  }

  const peData = peSheet.getDataRange().getValues();
  if (peData.length < 2) {
    ui.alert('❌ Error', '"Prices & Estimates" sheet is empty.', ui.ButtonSet.OK);
    return;
  }

  // Find column indices by header name
  const peHeaders = peData[0].map(h => String(h).trim().toLowerCase());
  const pePlayerCol = peHeaders.indexOf('player');
  const peSetCol    = peHeaders.indexOf('set');
  const peVariantCol = peHeaders.indexOf('variant');
  const peBestCol   = peHeaders.indexOf('best price');

  if (pePlayerCol < 0 || peSetCol < 0 || peVariantCol < 0 || peBestCol < 0) {
    ui.alert('❌ Error', 'Cannot find required columns in "Prices & Estimates".\n\nExpected: Player, Set, Variant, Best Price', ui.ButtonSet.OK);
    return;
  }

  // Build lookup: "player|||set|||variant" → best price (all lowercase for fuzzy match)
  const priceMap = {};
  for (let i = 1; i < peData.length; i++) {
    const row = peData[i];
    const player  = String(row[pePlayerCol] || '').trim().toLowerCase();
    const set     = String(row[peSetCol]    || '').trim().toLowerCase();
    const variant = String(row[peVariantCol] || '').trim().toLowerCase();
    const price   = row[peBestCol];
    if (player && set && variant && price) {
      priceMap[`${player}|||${set}|||${variant}`] = price;
    }
  }

  // ── 2. Read the active sheet ───────────────────────────────────────────
  const targetSheet = ss.getActiveSheet();
  if (targetSheet.getName() === 'Prices & Estimates') {
    ui.alert('⚠️ Wrong sheet', 'Switch to your inventory/consignment sheet first, then run this.', ui.ButtonSet.OK);
    return;
  }

  const targetData = targetSheet.getDataRange().getValues();
  if (targetData.length < 2) {
    ui.alert('❌ Error', 'Active sheet is empty.', ui.ButtonSet.OK);
    return;
  }

  // Find the header row by scanning for a row that contains "Player Name"
  let tHeaderRow = -1;
  for (let i = 0; i < Math.min(targetData.length, 20); i++) {
    const rowVals = targetData[i].map(h => String(h).trim().toLowerCase());
    if (rowVals.includes('player name')) {
      tHeaderRow = i;
      break;
    }
  }
  if (tHeaderRow < 0) {
    ui.alert('❌ Error',
      'Cannot find a header row containing "Player Name" in "' + targetSheet.getName() + '".\n\n' +
      'Expected headers: "Player Name", "Set", "Hit", "GL Reference"',
      ui.ButtonSet.OK);
    return;
  }

  // Find column indices by header name
  const tHeaders = targetData[tHeaderRow].map(h => String(h).trim().toLowerCase());
  const tPlayerCol = tHeaders.indexOf('player name');
  const tSetCol    = tHeaders.indexOf('set');
  const tHitCol    = tHeaders.indexOf('hit');
  const tGLCol     = tHeaders.indexOf('gl reference');

  if (tPlayerCol < 0 || tSetCol < 0 || tHitCol < 0 || tGLCol < 0) {
    ui.alert('❌ Error',
      'Cannot find required columns in "' + targetSheet.getName() + '" (header row ' + (tHeaderRow + 1) + ').\n\n' +
      'Expected headers: "Player Name", "Set", "Hit", "GL Reference"',
      ui.ButtonSet.OK);
    return;
  }

  // ── 3. Match and write ─────────────────────────────────────────────────
  let matched = 0;
  let unmatched = 0;
  const unmatchedNames = [];

  for (let i = tHeaderRow + 1; i < targetData.length; i++) {
    const row = targetData[i];
    const player  = String(row[tPlayerCol] || '').trim();
    const set     = String(row[tSetCol]    || '').trim();
    const hit     = String(row[tHitCol]    || '').trim();

    if (!player || !set || !hit) continue; // skip empty rows

    const key = `${player.toLowerCase()}|||${set.toLowerCase()}|||${hit.toLowerCase()}`;
    const price = priceMap[key];

    if (price !== undefined) {
      // Write to GL Reference column (1-indexed: row i+1, col tGLCol+1)
      targetSheet.getRange(i + 1, tGLCol + 1).setValue(price);
      matched++;
    } else {
      unmatched++;
      if (unmatchedNames.length < 5) {
        unmatchedNames.push(`${player} / ${hit} / ${set}`);
      }
    }
  }

  // ── 4. Report ──────────────────────────────────────────────────────────
  let msg = `✅ Matched ${matched} rows.`;
  if (unmatched > 0) {
    msg += `\n\n⚠️ ${unmatched} rows had no match.`;
    if (unmatchedNames.length > 0) {
      msg += '\n\nExamples:\n• ' + unmatchedNames.join('\n• ');
      if (unmatched > 5) msg += `\n• ... and ${unmatched - 5} more`;
    }
  }
  ui.alert('Match GL References', msg, ui.ButtonSet.OK);
}


// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Determine the single "best" price for a row.
 * Priority: refreshed estimate (model-blended with real data) > 30D avg > all-time avg > gap-fill estimate
 */
function bestPrice_(pi, est) {
  // If we have a refreshed estimate (real data blended with model), use it
  if (est.is_refreshed && est.estimated_price > 0) return est.estimated_price;
  // If we have recent sales, prefer 30D average
  if (pi.avg_30d > 0) return pi.avg_30d;
  // Fall back to all-time average
  if (pi.avg_price > 0) return pi.avg_price;
  // Pure gap-fill estimate (no sales at all)
  if (est.estimated_price > 0) return est.estimated_price;
  return null;
}


/**
 * Paginated Supabase GET — fetches all rows (1000 per page).
 */
function sbGetAll_(table, qs, orderCol) {
  let all = [];
  let offset = 0;
  const limit = 1000;
  const order = orderCol ? `${orderCol}.asc` : 'set_name.asc,player.asc';

  while (true) {
    const url = `${SB_URL}/rest/v1/${table}?${qs}&order=${order}&limit=${limit}&offset=${offset}`;
    const res = UrlFetchApp.fetch(url, {
      headers: SB_HEADERS,
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() >= 300) {
      throw new Error(`sbGetAll_ ${table}: ${res.getContentText()}`);
    }

    const page = JSON.parse(res.getContentText());
    all = all.concat(page);

    if (page.length < limit) break;
    offset += limit;
  }

  return all;
}
