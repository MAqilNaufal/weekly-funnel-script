// ============================================================
// PBI_FUNNEL FLAT TABLE — Module 3
//
// Builds a denormalized flat table for Power BI with one row
// per Marketing Code + ISO Week, aggregating:
//   - Ad spend + platform leads (from Master Data)
//   - CRM leads (from omnichannelData)
//   - Site Visits + Closings + Revenue (from SV sheet)
//
// Output:  "PBI_Funnel" tab in SS_ADS spreadsheet
// Trigger: daily 7am WIB — set once with createPBITrigger()
// Window:  rolling 52 ISO weeks
// ============================================================

function refreshPBITable() {
  const now  = new Date();
  const curW = getISOWeek(now);
  const curY = now.getFullYear();

  // Rolling 52-week window: curWeek-51 → curWeek
  const weekKeys = [];
  for (let offset = 51; offset >= 0; offset--) {
    let w = curW - offset;
    let y = curY;
    if (w < 1) { w += 52; y -= 1; }
    weekKeys.push(`${y}-W${String(w).padStart(2,'0')}`);
  }
  const weekSet = new Set(weekKeys);

  Logger.log(`PBI Refresh: ${weekKeys[0]} → ${weekKeys[weekKeys.length-1]}`);

  // ── 1. Master Data: Spend + PlatformLeads per MktCode + Week + Region + Channel
  const adsSS      = SpreadsheetApp.openById(CONFIG.SS_ADS);
  const masterSheet = adsSS.getSheetByName(CONFIG.SHEET_MASTER_DATA);
  if (!masterSheet) throw new Error('Master Data (DO NOT TOUCH) not found');

  const masterVals = masterSheet.getDataRange().getValues();
  const mHdr = masterVals[0].map(h => String(h).trim());
  const MC = {};
  mHdr.forEach((h, i) => { MC[h] = i; });

  const adsMap = {}; // key: mktCode|week → { region, channel, spend, platformLeads }

  for (let i = 1; i < masterVals.length; i++) {
    const row  = masterVals[i];
    const yr   = String(row[MC['Year']]  || '').trim();
    const wkRaw = String(row[MC['Week']] || '').trim();
    const wk   = String(parseInt(wkRaw)).padStart(2,'0');
    const weekKey = `${yr}-W${wk}`;
    if (!weekSet.has(weekKey)) continue;

    let mktCode = String(row[MC['Marketing Code']] || '').trim();
    if (!mktCode || mktCode === '-') mktCode = extractMarketingCode(String(row[MC['Ad Name']] || ''));
    if (!mktCode) continue;

    const region  = String(row[MC['Region']]  || '').trim();
    const channel = String(row[MC['Channel']] || '').trim();
    const spend   = parseFloat(String(row[MC['Spends']] || '0').replace(/[^0-9.-]/g,'')) || 0;
    const leads   = parseFloat(String(row[MC['Result']] || '0').replace(/[^0-9.-]/g,'')) || 0;

    const key = `${mktCode}|${weekKey}`;
    if (!adsMap[key]) {
      adsMap[key] = { mktCode, weekKey, region, channel, spend: 0, platformLeads: 0 };
    }
    adsMap[key].spend         += spend;
    adsMap[key].platformLeads += leads;
    if (region && !adsMap[key].region) adsMap[key].region = region;
    if (channel && !adsMap[key].channel) adsMap[key].channel = channel;
  }

  // ── 2. omnichannelData: CRM Leads per MktCode + Week
  const crmSS   = SpreadsheetApp.openById(CONFIG.SS_MAIN);
  const omniSht = crmSS.getSheetByName(CONFIG.SHEET_OMNI);
  const crmMap  = {}; // key: mktCode|week → { crmLeads }

  if (omniSht) {
    const omniVals = omniSht.getDataRange().getValues();
    const oHdr = omniVals[0].map(h => String(h).trim());
    const OC = {};
    oHdr.forEach((h, i) => { OC[h] = i; });

    for (let i = 1; i < omniVals.length; i++) {
      const row   = omniVals[i];
      const yr    = String(row[OC['Year']]     || '').trim();
      const wkRaw = String(row[OC['ISO Week']] || '').trim();
      const wk    = String(parseInt(wkRaw)).padStart(2,'0');
      const weekKey = `${yr}-W${wk}`;
      if (!weekSet.has(weekKey)) continue;

      let mktCode = String(row[OC['Marketing Code']] || '').trim();
      if (!mktCode || mktCode === '-') continue;

      const key = `${mktCode}|${weekKey}`;
      crmMap[key] = (crmMap[key] || 0) + 1;
    }
  }

  // ── 3. SV sheet: Site Visits + Closings + Revenue per MktCode + Week
  const svSht = crmSS.getSheetByName('SV');
  const svMap = {};  // key: mktCode|week → { sv, closing, revenue }

  if (svSht) {
    const svVals = svSht.getDataRange().getValues();
    const sHdr = svVals[0].map(h => String(h).trim());
    const SC = {};
    sHdr.forEach((h, i) => { SC[h] = i; });

    for (let i = 1; i < svVals.length; i++) {
      const row     = svVals[i];
      const mktCode = String(row[SC['Marketing Code']] || '').trim();
      if (!mktCode || mktCode === '-') continue;

      const isoWeekRaw = row[SC['ISO Week']];
      let weekKey = '';

      if (isoWeekRaw && String(isoWeekRaw).trim()) {
        // If ISO Week column is present, use it
        const isoWk = parseInt(String(isoWeekRaw).trim());
        const svDate = row[SC['SV Date']];
        const yr = svDate ? new Date(svDate).getFullYear() : new Date().getFullYear();
        weekKey = `${yr}-W${String(isoWk).padStart(2,'0')}`;
      } else {
        // Fall back to computing from SV Date
        const svDate = row[SC['SV Date']];
        if (!svDate) continue;
        const d = new Date(svDate);
        if (isNaN(d)) continue;
        weekKey = `${d.getFullYear()}-W${String(getISOWeek(d)).padStart(2,'0')}`;
      }

      if (!weekSet.has(weekKey)) continue;

      const key = `${mktCode}|${weekKey}`;
      if (!svMap[key]) svMap[key] = { sv: 0, closing: 0, revenue: 0 };
      svMap[key].sv++;

      // Check if it also has a closing
      const closingDate = row[SC['Closing Date']];
      if (closingDate) {
        const cd = new Date(closingDate);
        if (!isNaN(cd)) {
          const closingWeek = `${cd.getFullYear()}-W${String(getISOWeek(cd)).padStart(2,'0')}`;
          const cKey = `${mktCode}|${closingWeek}`;
          if (!svMap[cKey]) svMap[cKey] = { sv: 0, closing: 0, revenue: 0 };
          svMap[cKey].closing++;
          const rev = parseFloat(String(row[SC['Closing Revenue']] || '0').replace(/[^0-9.-]/g,'')) || 0;
          svMap[cKey].revenue += rev;
        }
      }
    }
  }

  // ── 4. Merge all keys into one flat table ─────────────────
  const allKeys = new Set([
    ...Object.keys(adsMap),
    ...Object.keys(crmMap).filter(k => crmMap[k] > 0),
    ...Object.keys(svMap).filter(k => svMap[k].sv > 0 || svMap[k].closing > 0),
  ]);

  const flatRows = [];
  allKeys.forEach(key => {
    const [mktCode, weekKey] = key.split('|');
    const ads = adsMap[key] || {};
    const crm = crmMap[key] || 0;
    const sv  = svMap[key]  || { sv: 0, closing: 0, revenue: 0 };

    // Week start date (Monday of that ISO week)
    const [yr, wStr] = weekKey.split('-W');
    const weekStart = getWeekStartDate(parseInt(yr), parseInt(wStr));

    flatRows.push({
      weekKey,
      weekStart: Utilities.formatDate(weekStart, CONFIG.TZ, 'yyyy-MM-dd'),
      isoWeek:   parseInt(wStr),
      year:      parseInt(yr),
      region:    ads.region    || '',
      platform:  ads.channel   || '',
      mktCode,
      adSpend:        ads.spend          || 0,
      platformLeads:  ads.platformLeads  || 0,
      crmLeads:       crm,
      sv:             sv.sv,
      closing:        sv.closing,
      revenue:        sv.revenue,
    });
  });

  // Sort by weekKey desc, then mktCode asc
  flatRows.sort((a, b) => {
    if (b.weekKey !== a.weekKey) return b.weekKey.localeCompare(a.weekKey);
    return a.mktCode.localeCompare(b.mktCode);
  });

  // ── 5. Write PBI_Funnel sheet ─────────────────────────────
  let pbiSheet = adsSS.getSheetByName('PBI_Funnel');
  if (!pbiSheet) {
    pbiSheet = adsSS.insertSheet('PBI_Funnel');
  } else {
    pbiSheet.clearContents();
  }

  const headers = [
    'Week', 'Week Start', 'ISO Week', 'Year',
    'Region', 'Platform', 'Marketing Code',
    'Ad Spend', 'Platform Leads', 'CRM Leads',
    'Site Visits', 'Closings', 'Revenue',
  ];

  const outputRows = [headers];
  flatRows.forEach(r => {
    outputRows.push([
      r.weekKey, r.weekStart, r.isoWeek, r.year,
      r.region, r.platform, r.mktCode,
      r.adSpend, r.platformLeads, r.crmLeads,
      r.sv, r.closing, r.revenue,
    ]);
  });

  pbiSheet.getRange(1, 1, outputRows.length, headers.length).setValues(outputRows);

  // Freeze header row, bold it
  pbiSheet.setFrozenRows(1);
  pbiSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a5276').setFontColor('white');

  const generated = Utilities.formatDate(new Date(), CONFIG.TZ, 'dd MMM yyyy HH:mm');
  Logger.log(`PBI_Funnel refreshed: ${flatRows.length} rows · ${generated}`);
  return { rows: flatRows.length, generated };
}

// ── Helper: get Monday date of a given ISO year + week ───────

function getWeekStartDate(year, isoWeek) {
  // ISO week 1 is the week containing the first Thursday of the year.
  // Find Jan 4 of the year (always in week 1), then back to Monday of that week.
  const jan4 = new Date(year, 0, 4);
  const dayOfWeek = jan4.getDay() || 7; // Mon=1 … Sun=7
  const week1Mon = new Date(jan4.getTime() - (dayOfWeek - 1) * 86400000);
  return new Date(week1Mon.getTime() + (isoWeek - 1) * 7 * 86400000);
}

/**
 * Run once to create the daily 7am refresh trigger.
 */
function createPBITrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshPBITable')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshPBITable')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  Logger.log('PBI daily 7am trigger created');
}
