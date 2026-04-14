// ============================================================
// COMPETITOR BENCHMARK — Module 4
//
// Cross-joins Competitor Analysis sheet with internal SV/Closing
// data by Region + Period to produce benchmark delta columns.
//
// Called by:
//   - n8n webhook (after each new competitor row is appended)
//   - Manual: runCompetitorBenchmark()
//
// Adds/updates cols in "Competitor Analysis" sheet:
//   Our_SV | Our_Closing | Our_Revenue | Comp_SalesVol_Delta | Notes_Benchmark
// ============================================================

/**
 * Main entry — cross-joins latest period for all regions.
 * Call this after new competitor data is appended.
 */
function runCompetitorBenchmark() {
  const crmSS  = SpreadsheetApp.openById(CONFIG.SS_MAIN);
  const compSht = crmSS.getSheetByName('Competitor Analysis');
  if (!compSht) {
    Logger.log('Competitor Analysis sheet not found');
    return;
  }

  const compData = compSht.getDataRange().getValues();
  if (compData.length < 2) return;

  const ch = compData[0].map(h => String(h).trim());
  const CI = {};
  ch.forEach((h, i) => { CI[h] = i; });

  // Ensure benchmark columns exist — append if missing
  const BENCH_COLS = ['Our_SV', 'Our_Closing', 'Our_Revenue (Mio)', 'Comp vs Our Closing Delta', 'Benchmark Notes'];
  let lastCol = ch.length;
  BENCH_COLS.forEach((bc, idx) => {
    if (!ch.includes(bc)) {
      ch.push(bc);
      CI[bc] = lastCol + idx;
    }
  });

  // Write headers if extended
  compSht.getRange(1, 1, 1, ch.length).setValues([ch]);

  // ── Build internal data maps ──────────────────────────────
  // Group SV sheet by Region + Period (ISO year-month)
  const internalMap = buildInternalMap(crmSS);

  // ── Unique periods in competitor data ─────────────────────
  const periods = [...new Set(
    compData.slice(1).map(r => String(r[CI['Period']] || '').trim()).filter(Boolean)
  )];
  Logger.log(`Periods found: ${periods.join(', ')}`);

  // ── Enrich each row ────────────────────────────────────────
  const updatedRows = [ch]; // include header

  for (let i = 1; i < compData.length; i++) {
    const row    = compData[i].slice(); // copy
    const region = String(row[CI['Region']] || '').trim();
    const period = String(row[CI['Period']] || '').trim();

    const internal = internalMap[`${region}|${period}`] || {};

    // Extend row to full width
    while (row.length < ch.length) row.push('');

    row[CI['Our_SV']]                    = internal.sv       || '';
    row[CI['Our_Closing']]               = internal.closing  || '';
    row[CI['Our_Revenue (Mio)']]         = internal.revMio   || '';

    // Delta: Competitor total units sold (takeup * total) vs our closings
    const compTakeup = parseFloat(row[CI['Take Up % (Current)']] || 0) / 100;
    const compUnits  = parseInt(row[CI['Total Units']] || 0) || 0;
    const compSold   = Math.round(compTakeup * compUnits);
    const ourClose   = internal.closing || 0;

    if (compSold > 0 && ourClose > 0) {
      const delta = ourClose - compSold;
      row[CI['Comp vs Our Closing Delta']] = delta;
      row[CI['Benchmark Notes']] = delta >= 0
        ? `We closed ${delta} more units than ${row[CI['Cluster']]} this period`
        : `${row[CI['Cluster']]} sold ${Math.abs(delta)} more units than us`;
    } else {
      row[CI['Comp vs Our Closing Delta']] = '';
      row[CI['Benchmark Notes']] = '';
    }

    updatedRows.push(row);
  }

  compSht.clearContents();
  compSht.getRange(1, 1, updatedRows.length, ch.length).setValues(updatedRows);
  Logger.log(`Competitor Benchmark complete: ${updatedRows.length - 1} rows enriched`);
}

// ── Build internal SV/Closing map ────────────────────────────
// Returns { "Region|Period": { sv, closing, revMio } }
// Period format: "Month YYYY" e.g. "March 2026"

function buildInternalMap(crmSS) {
  const map = {};
  const svSht = crmSS.getSheetByName('SV');
  if (!svSht) return map;

  const svData = svSht.getDataRange().getValues();
  const sh = svData[0].map(h => String(h).trim());
  const SI = {};
  sh.forEach((h, i) => { SI[h] = i; });

  const MONTH_NAMES = ['January','February','March','April','May','June',
                       'July','August','September','October','November','December'];

  for (let i = 1; i < svData.length; i++) {
    const row  = svData[i];
    const team = String(row[SI['Team']] || '').trim();

    // Map team → region name
    const region = teamToRegion(team);
    if (!region) continue;

    // SV period from SV Date
    const svDate = row[SI['SV Date']];
    if (!svDate) continue;
    const d = new Date(svDate);
    if (isNaN(d.getTime())) continue;
    const period = `${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}`;
    const key    = `${region}|${period}`;

    if (!map[key]) map[key] = { sv: 0, closing: 0, revMio: 0 };
    map[key].sv++;

    // Closing
    const closingDate = row[SI['Closing Date']];
    if (closingDate) {
      const cd = new Date(closingDate);
      if (!isNaN(cd.getTime())) {
        const cPeriod = `${MONTH_NAMES[cd.getMonth()]} ${cd.getFullYear()}`;
        const cKey    = `${region}|${cPeriod}`;
        if (!map[cKey]) map[cKey] = { sv: 0, closing: 0, revMio: 0 };
        map[cKey].closing++;
        const rev = parseFloat(String(row[SI['Closing Revenue']] || '0').replace(/[^0-9.]/g,'')) || 0;
        map[cKey].revMio += Math.round(rev / 1000000) / 1000; // IDR to Mio, rounded to 3dp
      }
    }
  }

  return map;
}

// ── Team → Region mapping ─────────────────────────────────────
function teamToRegion(team) {
  const t = team.toLowerCase();
  if (t.includes('gmtd') || t.includes('makassar') || t.includes('tanjung bunga'))  return 'Makassar';
  if (t.includes('manado'))                                                           return 'Manado';
  if (t.includes('lcc') || t.includes('cikarang') || t.includes('lippo cikarang'))   return 'Cikarang';
  if (t.includes('karawang'))                                                         return 'Cikarang'; // Karawang is Cikarang region
  if (t.includes('lv') || t.includes('lippo village') || t.includes('park serpong')) return 'Tangerang';
  if (t.includes('pwk') || t.includes('purwakarta'))                                 return 'Tangerang';
  if (t.includes('ps') || t.includes('serpong'))                                     return 'Tangerang';
  return '';
}

/**
 * HTTP doPost endpoint — called by n8n after appending a new competitor row.
 * n8n sends: { action: "benchmark" }
 */
function doPost(e) {
  try {
    runCompetitorBenchmark();
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
