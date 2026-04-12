// ============================================================
// WEEKLY FUNNEL REPORT — Apps Script
// Spreadsheet: CRM Main (1L9S25h79CdH9QarN-_0dnkXogeVx1vJ5C4oXKt_8G-c)
// ============================================================

const CONFIG = {
  // Spreadsheet IDs
  SS_MAIN:      '1L9S25h79CdH9QarN-_0dnkXogeVx1vJ5C4oXKt_8G-c',
  SS_ADS:       '1M27qJ3VsBMd1VvYSOBblixZwWKx_1lJ9ONHGptyiFbY',
  SS_MKT_CODE:  '1aC7I0eDB1IA_0_T-VCbvxM78qfRKQcogScy-6495vvM',
  SS_CREATIVE:  '1DAGe0MrJNjDn06jm0ePHdKztGnx0XIxDdtQt1GiFtPc',

  // Sheet names
  SHEET_MASTER_DATA:  'Master Data (DO NOT TOUCH)',
  SHEET_MKT_CODE:     'AllMarketingCode',
  SHEET_CREATIVE_DB:  'centralized creative database',
  SHEET_OMNI:         'omnichannelData',
  SHEET_SV:           '[MAIN] SITE VISIT',
  SHEET_REPORT:       'Weekly Funnel Report',

  // Email
  RECIPIENTS: 'aqilnaufalb@gmail.com',

  // Timezone
  TZ: 'Asia/Jakarta',
};

// Marketing code regex: e.g. LV26M001, LC26M006, PS26M001
const MKT_CODE_REGEX = /\b([A-Z]{2,4}\d{2}M\d{3})\b/;

// ============================================================
// ISO WEEK UTILITIES
// ============================================================

function getISOWeek(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  const week1 = new Date(d.getFullYear(), 0, 4);
  return 1 + Math.round(((d - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

function getWeekDateRange(isoWeek, year) {
  const jan4 = new Date(year, 0, 4);
  const startOfWeek1 = new Date(jan4);
  startOfWeek1.setDate(jan4.getDate() - (jan4.getDay() + 6) % 7);
  const mon = new Date(startOfWeek1);
  mon.setDate(startOfWeek1.getDate() + (isoWeek - 1) * 7);
  const sun = new Date(mon);
  sun.setDate(mon.getDate() + 6);
  const fmt = d => Utilities.formatDate(d, CONFIG.TZ, 'dd MMM yyyy');
  return { mon, sun, label: `W${isoWeek}: ${fmt(mon)} – ${fmt(sun)}` };
}

// ============================================================
// HELPERS
// ============================================================

function extractMarketingCode(adName) {
  if (!adName) return '';
  const m = String(adName).match(MKT_CODE_REGEX);
  return m ? m[1] : '';
}

/**
 * Build a lookup map from a sheet.
 * keyColIdx: 0-based column index for the key
 * valueColIdxs: array of 0-based column indexes for values
 * Returns: { key: [val0, val1, ...] }
 */
function buildLookupMap(sheet, keyColIdx, valueColIdxs) {
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][keyColIdx]).trim();
    if (!key) continue;
    if (!map[key]) {
      map[key] = valueColIdxs.map(c => data[i][c]);
    }
  }
  return map;
}

function numFmt(n) {
  if (!n && n !== 0) return '-';
  return Number(n).toLocaleString('id-ID');
}

function pctFmt(num, den) {
  if (!den || den === 0) return '-';
  return (num / den * 100).toFixed(1) + '%';
}

function currencyFmt(n) {
  if (!n && n !== 0) return '-';
  return 'Rp ' + Number(n).toLocaleString('id-ID');
}

function delta(curr, prev) {
  if (!prev || prev === 0) return curr > 0 ? '+∞' : '-';
  const pct = ((curr - prev) / prev * 100).toFixed(1);
  return (pct > 0 ? '+' : '') + pct + '%';
}

// ============================================================
// MAIN REPORT GENERATOR
// ============================================================

function generateWeeklyFunnelReport() {
  const now = new Date();
  const thisWeek = getISOWeek(now);
  const lastWeek = thisWeek > 1 ? thisWeek - 1 : 52;
  const year = now.getFullYear();
  const weekRange = getWeekDateRange(thisWeek, year);

  Logger.log(`Generating report for ISO Week ${thisWeek} (${weekRange.label})`);

  // ── 1. Open all spreadsheets ──────────────────────────────
  const ssCRM      = SpreadsheetApp.openById(CONFIG.SS_MAIN);
  const ssAds      = SpreadsheetApp.openById(CONFIG.SS_ADS);
  const ssMktCode  = SpreadsheetApp.openById(CONFIG.SS_MKT_CODE);
  const ssCreative = SpreadsheetApp.openById(CONFIG.SS_CREATIVE);

  // ── 2. Read Master Data (DO NOT TOUCH) ───────────────────
  // Cols: A=Year, B=Week, C=Date, D=Channel, E=CampaignType, F=Campaign,
  //       G=AdSet, H=AdName, I=MktCode, J=Objective,
  //       K=Spends, L=Impressions, M=Clicks, N=Result, ..., T=Region
  const masterSheet = ssAds.getSheetByName(CONFIG.SHEET_MASTER_DATA);
  const masterData  = masterSheet.getDataRange().getValues();

  // Aggregate by marketing code for thisWeek and lastWeek
  const adsAgg = {}; // { mktCode: { region, channel, adName, spendTW, leadsTW, spendLW, leadsLW } }

  for (let i = 1; i < masterData.length; i++) {
    const row     = masterData[i];
    const rowWeek = parseInt(row[1]);
    if (rowWeek !== thisWeek && rowWeek !== lastWeek) continue;

    const adName  = String(row[7]).trim();
    const rawCode = String(row[8]).trim();
    const mktCode = rawCode.match(MKT_CODE_REGEX) ? rawCode : extractMarketingCode(adName);
    if (!mktCode) continue;

    const spend   = parseFloat(String(row[10]).replace(/,/g, '')) || 0;
    const leads   = parseFloat(String(row[13]).replace(/,/g, '')) || 0;
    const region  = String(row[19]).trim();
    const channel = String(row[3]).trim();

    if (!adsAgg[mktCode]) {
      adsAgg[mktCode] = { region, channel, adName, spendTW: 0, leadsTW: 0, spendLW: 0, leadsLW: 0 };
    }

    if (rowWeek === thisWeek) {
      adsAgg[mktCode].spendTW += spend;
      adsAgg[mktCode].leadsTW += leads;
    } else {
      adsAgg[mktCode].spendLW += spend;
      adsAgg[mktCode].leadsLW += leads;
    }
  }

  // ── 3. AllMarketingCode lookup ────────────────────────────
  // Cols: A=Region(0), B=Channel(1), C=MktCode(2), D=Campaign(3),
  //       E=AdSet(4), F=AdsLink(5), G=CreatedAt(6), H=Notes(7),
  //       I=WALink(8), J=PredText(9), K=CreativeLink(10)
  const mktCodeSheet = ssMktCode.getSheetByName(CONFIG.SHEET_MKT_CODE);
  const mktCodeMap   = buildLookupMap(mktCodeSheet, 2, [0, 1, 3, 10]); // key=MktCode → [region, channel, campaign, creativeLink]

  // ── 4. Creative Database lookup ───────────────────────────
  // Cols: A=cre(0), B=Region(1), C=Theme(2), D=Format(3), E=Version(4),
  //       F=Notes(5), G=Product(6), H=Caption(7), I=CreativeID(8),
  //       J=LinkCreative(9), K=LinkPost(10)
  const creativeSheet = ssCreative.getSheetByName(CONFIG.SHEET_CREATIVE_DB);
  const creativeMap   = buildLookupMap(creativeSheet, 8, [9, 10]); // key=CreativeID → [linkCreative, linkPost]

  // ── 5. omnichannelData — count CRM leads by mktCode + week ─
  // Cols: A=Team(0), ..., H=ISOWeek(7), ..., O=MktCode(14)
  const omniSheet = ssCRM.getSheetByName(CONFIG.SHEET_OMNI);
  const omniData  = omniSheet.getDataRange().getValues();

  const crmLeads = {}; // { mktCode: { tw: n, lw: n } }
  const crmByRegion = {}; // { region: { leadsTW, leadsLW } }

  for (let i = 1; i < omniData.length; i++) {
    const row     = omniData[i];
    const rowWeek = parseInt(row[7]);
    if (rowWeek !== thisWeek && rowWeek !== lastWeek) continue;

    const mktCode = String(row[14]).trim();
    const team    = String(row[0]).trim();

    if (mktCode) {
      if (!crmLeads[mktCode]) crmLeads[mktCode] = { tw: 0, lw: 0 };
      if (rowWeek === thisWeek) crmLeads[mktCode].tw++;
      else crmLeads[mktCode].lw++;
    }

    // Region from team name (extract region keyword)
    const region = team;
    if (!crmByRegion[region]) crmByRegion[region] = { leadsTW: 0, leadsLW: 0, svTW: 0, svLW: 0, closingTW: 0, closingLW: 0, revTW: 0, revLW: 0 };
    if (rowWeek === thisWeek) crmByRegion[region].leadsTW++;
    else crmByRegion[region].leadsLW++;
  }

  // ── 6. [MAIN] SITE VISIT — SV + Closing by mktCode + week ─
  // Cols: A=No(0), B=Team(1), ..., G=LeadSource(6), H=SVDate(7), I=SVUnit(8),
  //       J=ClosingDate(9), K=ClosingUnit(10), L=Revenue(11), ...,
  //       O=ISOWeek(14), P=MktCode(15)
  const svSheet = ssCRM.getSheetByName(CONFIG.SHEET_SV);
  const svData  = svSheet.getDataRange().getValues();

  const svAgg = {}; // { mktCode: { svTW, svLW, closingTW, closingLW, revTW, revLW } }

  for (let i = 1; i < svData.length; i++) {
    const row     = svData[i];
    const rowWeek = parseInt(row[14]);
    if (rowWeek !== thisWeek && rowWeek !== lastWeek) continue;

    const mktCode = String(row[15]).trim();
    const team    = String(row[1]).trim();
    const revenue = parseFloat(String(row[11]).replace(/[^0-9.-]/g, '')) || 0;
    const hasClosed = row[9] && String(row[9]).trim() !== '';

    if (mktCode) {
      if (!svAgg[mktCode]) svAgg[mktCode] = { svTW: 0, svLW: 0, closingTW: 0, closingLW: 0, revTW: 0, revLW: 0 };
      if (rowWeek === thisWeek) {
        svAgg[mktCode].svTW++;
        if (hasClosed) { svAgg[mktCode].closingTW++; svAgg[mktCode].revTW += revenue; }
      } else {
        svAgg[mktCode].svLW++;
        if (hasClosed) { svAgg[mktCode].closingLW++; svAgg[mktCode].revLW += revenue; }
      }
    }

    // Region breakdown
    if (!crmByRegion[team]) crmByRegion[team] = { leadsTW: 0, leadsLW: 0, svTW: 0, svLW: 0, closingTW: 0, closingLW: 0, revTW: 0, revLW: 0 };
    if (rowWeek === thisWeek) {
      crmByRegion[team].svTW++;
      if (hasClosed) { crmByRegion[team].closingTW++; crmByRegion[team].revTW += revenue; }
    } else {
      crmByRegion[team].svLW++;
      if (hasClosed) { crmByRegion[team].closingLW++; crmByRegion[team].revLW += revenue; }
    }
  }

  // ── 7. Compute summary totals ─────────────────────────────
  const totals = { platLeadsTW: 0, platLeadsLW: 0, crmLeadsTW: 0, crmLeadsLW: 0, svTW: 0, svLW: 0, closingTW: 0, closingLW: 0, revTW: 0, revLW: 0 };

  for (const c of Object.values(adsAgg)) {
    totals.platLeadsTW += c.leadsTW;
    totals.platLeadsLW += c.leadsLW;
  }
  for (const c of Object.values(crmLeads)) {
    totals.crmLeadsTW += c.tw;
    totals.crmLeadsLW += c.lw;
  }
  for (const c of Object.values(svAgg)) {
    totals.svTW       += c.svTW;
    totals.svLW       += c.svLW;
    totals.closingTW  += c.closingTW;
    totals.closingLW  += c.closingLW;
    totals.revTW      += c.revTW;
    totals.revLW      += c.revLW;
  }

  // ── 8. Build ad-level detail rows ────────────────────────
  const allCodes = new Set([...Object.keys(adsAgg), ...Object.keys(crmLeads), ...Object.keys(svAgg)]);
  const detailRows = [];

  for (const code of [...allCodes].sort()) {
    const ads    = adsAgg[code]    || { region: '', channel: '', adName: '', spendTW: 0, leadsTW: 0 };
    const crm    = crmLeads[code]  || { tw: 0, lw: 0 };
    const sv     = svAgg[code]     || { svTW: 0, closingTW: 0, revTW: 0 };

    // Lookup enrichment
    const mktInfo      = mktCodeMap[code]  || ['', '', '', ''];
    const regionLabel  = mktInfo[0] || ads.region;
    const channelLabel = mktInfo[1] || ads.channel;
    const creativeLink = mktInfo[3] || '';

    // Try to get post link from creative DB (creativeLink might be a Creative ID)
    const creativeInfo = creativeMap[creativeLink] || ['', ''];
    const linkCreative = creativeInfo[0] || creativeLink;
    const linkPost     = creativeInfo[1] || '';

    const cpl = ads.leadsTW > 0 ? (ads.spendTW / ads.leadsTW).toFixed(0) : '-';

    detailRows.push([
      code,
      regionLabel,
      channelLabel,
      ads.adName,
      ads.channel,
      ads.spendTW,
      cpl,
      ads.leadsTW,
      crm.tw,
      sv.svTW,
      sv.closingTW,
      sv.revTW,
      linkCreative,
      linkPost,
    ]);
  }

  // ── 9. Write to sheet ─────────────────────────────────────
  const reportSheet = ssCRM.getSheetByName(CONFIG.SHEET_REPORT);
  reportSheet.clearContents();

  let row = 1;

  // Config header
  reportSheet.getRange(row, 1, 1, 4).setValues([[
    `Weekly Funnel Report`, weekRange.label, `Generated: ${Utilities.formatDate(now, CONFIG.TZ, 'dd MMM yyyy HH:mm')} WIB`, ''
  ]]);
  row += 2;

  // Summary snapshot
  reportSheet.getRange(row, 1).setValue('SUMMARY SNAPSHOT');
  row++;
  const summaryHeaders = ['Metric', `W${thisWeek} (This Week)`, `W${lastWeek} (Last Week)`, 'WoW Δ'];
  reportSheet.getRange(row, 1, 1, 4).setValues([summaryHeaders]);
  row++;

  const summaryRows = [
    ['Total Platform Leads',  totals.platLeadsTW,  totals.platLeadsLW,  delta(totals.platLeadsTW, totals.platLeadsLW)],
    ['Total CRM Leads',       totals.crmLeadsTW,   totals.crmLeadsLW,   delta(totals.crmLeadsTW,  totals.crmLeadsLW)],
    ['Platform → CRM Rate',   pctFmt(totals.crmLeadsTW, totals.platLeadsTW), pctFmt(totals.crmLeadsLW, totals.platLeadsLW), '-'],
    ['Total Site Visits',     totals.svTW,         totals.svLW,         delta(totals.svTW,        totals.svLW)],
    ['Lead → SV Rate',        pctFmt(totals.svTW, totals.crmLeadsTW),   pctFmt(totals.svLW, totals.crmLeadsLW),  '-'],
    ['Total Closings',        totals.closingTW,    totals.closingLW,    delta(totals.closingTW,   totals.closingLW)],
    ['SV → Closing Rate',     pctFmt(totals.closingTW, totals.svTW),    pctFmt(totals.closingLW, totals.svLW),   '-'],
    ['Total Revenue',         currencyFmt(totals.revTW), currencyFmt(totals.revLW), delta(totals.revTW, totals.revLW)],
  ];
  reportSheet.getRange(row, 1, summaryRows.length, 4).setValues(summaryRows);
  row += summaryRows.length + 2;

  // Ad-level detail
  reportSheet.getRange(row, 1).setValue('AD / MARKETING CODE DETAIL');
  row++;
  const detailHeaders = ['Marketing Code', 'Region', 'Channel', 'Ad Name', 'Platform', 'Spend (TW)', 'CPL', 'Platform Leads', 'CRM Leads', 'Site Visits', 'Closings', 'Revenue', 'Creative Link', 'Post Link'];
  reportSheet.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  row++;
  if (detailRows.length > 0) {
    reportSheet.getRange(row, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
    row += detailRows.length;
  }
  row += 2;

  // Project/region breakdown
  reportSheet.getRange(row, 1).setValue('BREAKDOWN BY PROJECT / TEAM');
  row++;
  const regionHeaders = ['Team / Project', `Leads W${thisWeek}`, `Leads W${lastWeek}`, `SV W${thisWeek}`, `SV W${lastWeek}`, `Closing W${thisWeek}`, `Revenue W${thisWeek}`, 'L→SV%', 'SV→C%'];
  reportSheet.getRange(row, 1, 1, regionHeaders.length).setValues([regionHeaders]);
  row++;
  const regionRows = Object.entries(crmByRegion).sort().map(([team, d]) => [
    team, d.leadsTW, d.leadsLW, d.svTW, d.svLW, d.closingTW,
    currencyFmt(d.revTW),
    pctFmt(d.svTW, d.leadsTW),
    pctFmt(d.closingTW, d.svTW),
  ]);
  if (regionRows.length > 0) {
    reportSheet.getRange(row, 1, regionRows.length, regionHeaders.length).setValues(regionRows);
  }

  Logger.log(`Sheet updated. ${detailRows.length} marketing codes, ${Object.keys(crmByRegion).length} teams.`);

  // ── 10. Send email ────────────────────────────────────────
  const htmlBody = buildHtmlEmail(weekRange.label, totals, thisWeek, lastWeek, detailRows, regionRows);
  GmailApp.sendEmail(
    CONFIG.RECIPIENTS,
    `Weekly Funnel Report – ${weekRange.label}`,
    '',
    { htmlBody: htmlBody }
  );

  Logger.log('Email sent to: ' + CONFIG.RECIPIENTS);
}

// ============================================================
// EMAIL HTML BUILDER
// ============================================================

function buildHtmlEmail(weekLabel, totals, tw, lw, detailRows, regionRows) {
  const style = `
    body { font-family: Arial, sans-serif; font-size: 13px; color: #222; }
    h2 { color: #1a5276; }
    h3 { color: #2e86c1; margin-top: 24px; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
    th { background: #2e86c1; color: white; padding: 7px 10px; text-align: left; font-size: 12px; }
    td { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; font-size: 12px; }
    tr:nth-child(even) { background: #f4f8fb; }
    .pos { color: #1e8449; font-weight: bold; }
    .neg { color: #c0392b; font-weight: bold; }
    .section { margin-top: 28px; }
    a { color: #2e86c1; }
  `;

  const summaryTable = `
    <table>
      <tr><th>Metric</th><th>W${tw} (This Week)</th><th>W${lw} (Last Week)</th><th>WoW Δ</th></tr>
      <tr><td>Total Platform Leads</td><td>${numFmt(totals.platLeadsTW)}</td><td>${numFmt(totals.platLeadsLW)}</td><td>${fmtDelta(totals.platLeadsTW, totals.platLeadsLW)}</td></tr>
      <tr><td>Total CRM Leads</td><td>${numFmt(totals.crmLeadsTW)}</td><td>${numFmt(totals.crmLeadsLW)}</td><td>${fmtDelta(totals.crmLeadsTW, totals.crmLeadsLW)}</td></tr>
      <tr><td>Platform → CRM Rate</td><td>${pctFmt(totals.crmLeadsTW, totals.platLeadsTW)}</td><td>${pctFmt(totals.crmLeadsLW, totals.platLeadsLW)}</td><td>-</td></tr>
      <tr><td>Total Site Visits</td><td>${numFmt(totals.svTW)}</td><td>${numFmt(totals.svLW)}</td><td>${fmtDelta(totals.svTW, totals.svLW)}</td></tr>
      <tr><td>Lead → SV Rate</td><td>${pctFmt(totals.svTW, totals.crmLeadsTW)}</td><td>${pctFmt(totals.svLW, totals.crmLeadsLW)}</td><td>-</td></tr>
      <tr><td>Total Closings</td><td>${numFmt(totals.closingTW)}</td><td>${numFmt(totals.closingLW)}</td><td>${fmtDelta(totals.closingTW, totals.closingLW)}</td></tr>
      <tr><td>SV → Closing Rate</td><td>${pctFmt(totals.closingTW, totals.svTW)}</td><td>${pctFmt(totals.closingLW, totals.svLW)}</td><td>-</td></tr>
      <tr><td>Total Revenue</td><td>${currencyFmt(totals.revTW)}</td><td>${currencyFmt(totals.revLW)}</td><td>${fmtDelta(totals.revTW, totals.revLW)}</td></tr>
    </table>`;

  const detailHeadersHtml = ['Code', 'Region', 'Channel', 'Ad Name', 'Platform', 'Spend', 'CPL', 'Plat. Leads', 'CRM Leads', 'SV', 'Closing', 'Revenue', 'Creative', 'Post'].map(h => `<th>${h}</th>`).join('');

  const detailRowsHtml = detailRows.map(r => {
    const creativeLink = r[12] ? `<a href="${r[12]}">View</a>` : '-';
    const postLink     = r[13] ? `<a href="${r[13]}">View</a>` : '-';
    return `<tr>
      <td><b>${r[0]}</b></td><td>${r[1]}</td><td>${r[2]}</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;">${r[3]}</td>
      <td>${r[4]}</td><td>${numFmt(r[5])}</td><td>${numFmt(r[6])}</td>
      <td>${numFmt(r[7])}</td><td>${numFmt(r[8])}</td>
      <td>${numFmt(r[9])}</td><td>${numFmt(r[10])}</td><td>${currencyFmt(r[11])}</td>
      <td>${creativeLink}</td><td>${postLink}</td>
    </tr>`;
  }).join('');

  const regionHeadersHtml = ['Team / Project', `Leads TW`, `Leads LW`, `SV TW`, `SV LW`, `Closing TW`, `Revenue TW`, 'L→SV%', 'SV→C%'].map(h => `<th>${h}</th>`).join('');
  const regionRowsHtml = regionRows.map(r => `<tr>${r.map(v => `<td>${v}</td>`).join('')}</tr>`).join('');

  return `<!DOCTYPE html><html><head><style>${style}</style></head><body>
    <h2>Weekly Funnel Report</h2>
    <p><b>${weekLabel}</b></p>

    <div class="section">
      <h3>Summary Snapshot</h3>
      ${summaryTable}
    </div>

    <div class="section">
      <h3>Ad / Marketing Code Detail</h3>
      <table>
        <tr>${detailHeadersHtml}</tr>
        ${detailRowsHtml || '<tr><td colspan="14">No data for this week</td></tr>'}
      </table>
    </div>

    <div class="section">
      <h3>Breakdown by Project / Team</h3>
      <table>
        <tr>${regionHeadersHtml}</tr>
        ${regionRowsHtml || '<tr><td colspan="9">No data</td></tr>'}
      </table>
    </div>

    <p style="color:#888;font-size:11px;margin-top:32px;">
      Auto-generated by Weekly Funnel Report script · ${Utilities.formatDate(new Date(), CONFIG.TZ, 'dd MMM yyyy HH:mm')} WIB
    </p>
  </body></html>`;
}

function fmtDelta(curr, prev) {
  if (!prev || prev === 0) return curr > 0 ? '<span class="pos">+∞</span>' : '-';
  const pct = ((curr - prev) / prev * 100).toFixed(1);
  const cls = pct >= 0 ? 'pos' : 'neg';
  return `<span class="${cls}">${pct >= 0 ? '+' : ''}${pct}%</span>`;
}

// ============================================================
// TRIGGER SETUP — run once manually
// ============================================================

function createWeeklyTrigger() {
  // Delete existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'generateWeeklyFunnelReport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new Monday 8am trigger
  ScriptApp.newTrigger('generateWeeklyFunnelReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log('Weekly Monday 8am trigger created.');
}
