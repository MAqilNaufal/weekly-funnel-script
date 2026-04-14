// ============================================================
// DAILY ADS DIGEST — sends a daily email at 9am WIB with
// a color-coded summary of today's ads performance per region.
//
// Data source: Snapshot sheet in SS_ADS spreadsheet.
// Recipient:   CONFIG.RECIPIENTS (currently aqilnaufalb@gmail.com)
// Trigger:     Set once with createDailyDigestTrigger()
// ============================================================

/**
 * Main entry point — called by the daily time-based trigger.
 */
function sendDailyAdsDigest() {
  const ss     = SpreadsheetApp.openById(CONFIG.SS_ADS);
  const sheet  = ss.getSheetByName('Snapshot');
  if (!sheet) {
    Logger.log('Snapshot sheet not found in SS_ADS');
    return;
  }

  const rows = sheet.getDataRange().getValues();
  if (!rows || rows.length < 5) {
    Logger.log('Snapshot sheet has insufficient data');
    return;
  }

  const reportDate = rows[1][1] ? String(rows[1][1]) : Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy');
  const dataAsOf   = rows[2][1] ? String(rows[2][1]) : reportDate;
  const headers    = rows[3]; // [Region, Channel, Spend, Leads, CPL, CPL 7d Avg, CPL Δ vs 7d, Spend MTD, Leads MTD, CPL MTD, Flag]

  // Parse data rows — fill down Region (merged cells appear as blank)
  const dataRows = [];
  let currentRegion = '';
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    const regionCell  = String(row[0] || '').trim();
    const channelCell = String(row[1] || '').trim();
    if (regionCell) currentRegion = regionCell;
    if (!channelCell) continue;
    dataRows.push({
      region:    currentRegion,
      channel:   channelCell,
      spend:     row[2],
      leads:     row[3],
      cpl:       row[4],
      cpl7dAvg:  row[5],
      cplDelta:  row[6],
      spendMTD:  row[7],
      leadsMTD:  row[8],
      cplMTD:    row[9],
      flag:      String(row[10] || '').trim(),
    });
  }

  // Compute totals for summary strip
  const detailRows = dataRows.filter(r => r.channel !== 'Subtotal');
  let totSpend = 0, totLeads = 0, totSpendMTD = 0, totLeadsMTD = 0;
  detailRows.forEach(r => {
    totSpend    += parseNumSafe(r.spend);
    totLeads    += parseNumSafe(r.leads);
    totSpendMTD += parseNumSafe(r.spendMTD);
    totLeadsMTD += parseNumSafe(r.leadsMTD);
  });
  const totCPL    = totLeads    > 0 ? Math.round(totSpend    / totLeads)    : 0;
  const totCPLMTD = totLeadsMTD > 0 ? Math.round(totSpendMTD / totLeadsMTD) : 0;

  // Count warnings
  const warnings = dataRows.filter(r => r.flag && r.flag.includes('⚠️'));
  const flagSummary = warnings.length === 0
    ? '<span style="color:#1e8449;font-weight:600;">✅ All regions healthy</span>'
    : `<span style="color:#c0392b;font-weight:600;">⚠️ ${warnings.length} alert(s) flagged</span>`;

  const html = buildDigestEmail({
    reportDate, dataAsOf, dataRows,
    totSpend, totLeads, totCPL,
    totSpendMTD, totLeadsMTD, totCPLMTD,
    flagSummary,
  });

  GmailApp.sendEmail(
    CONFIG.RECIPIENTS,
    `Daily Ads Snapshot — ${reportDate}`,
    'Please view this email in an HTML-capable client.',
    { htmlBody: html, name: 'Ads Digest' }
  );

  Logger.log('Daily Ads Digest sent for ' + reportDate);
}

// ── Email builder ────────────────────────────────────────────

function buildDigestEmail({ reportDate, dataAsOf, dataRows, totSpend, totLeads, totCPL, totSpendMTD, totLeadsMTD, totCPLMTD, flagSummary }) {
  const subtotalRows = dataRows.filter(r => r.channel === 'Subtotal');
  const fmt = n => n ? Number(n).toLocaleString('id-ID') : '-';
  const fmtRaw = v => (v === null || v === undefined || v === '') ? '-' : v;

  const deltaStyle = val => {
    if (!val || val === '-') return 'color:#999;';
    const n = parseFloat(String(val).replace('%', ''));
    if (isNaN(n)) return 'color:#999;';
    // For CPL: negative delta = good (green), positive = bad (red)
    return n < 0 ? 'color:#1e8449;font-weight:600;' : n > 0 ? 'color:#c0392b;font-weight:600;' : 'color:#999;';
  };

  const flagStyle = flag => flag && flag.includes('⚠️') ? 'color:#c0392b;font-weight:600;' : 'color:#1e8449;font-weight:600;';

  // Summary cards HTML
  const summaryCards = `
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
      <tr>
        ${[
          ['Today Spend', 'IDR ' + fmt(totSpend)],
          ['Today Leads', fmt(totLeads)],
          ['Today CPL', 'IDR ' + fmt(totCPL)],
          ['MTD Spend', 'IDR ' + fmt(totSpendMTD)],
          ['MTD Leads', fmt(totLeadsMTD)],
          ['MTD CPL', 'IDR ' + fmt(totCPLMTD)],
        ].map(([label, value]) => `
          <td style="padding:0 6px 0 0;" width="16%">
            <div style="background:white;border-radius:8px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">${label}</div>
              <div style="font-size:18px;font-weight:700;color:#1a5276;">${value}</div>
            </div>
          </td>`).join('')}
      </tr>
    </table>`;

  // Region subtotal table rows
  const regionTableRows = subtotalRows.map(r => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:600;">${r.region}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${fmt(parseNumSafe(r.spend))}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${fmtRaw(r.leads)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${fmt(parseNumSafe(r.cpl))}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${fmt(parseNumSafe(r.cpl7dAvg))}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;${deltaStyle(r.cplDelta)}">${fmtRaw(r.cplDelta)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;${flagStyle(r.flag)}">${r.flag || '-'}</td>
    </tr>`).join('');

  // Channel detail table rows (non-subtotal, non-zero spend or non-zero leads)
  const channelRows = dataRows
    .filter(r => r.channel !== 'Subtotal' && (parseNumSafe(r.spend) > 0 || parseNumSafe(r.leads) > 0))
    .map(r => `
      <tr>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;color:#555;">${r.region}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;">${r.channel}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:right;">${fmt(parseNumSafe(r.spend))}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:right;">${fmtRaw(r.leads)}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:right;">${fmt(parseNumSafe(r.cpl))}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;text-align:right;${deltaStyle(r.cplDelta)}">${fmtRaw(r.cplDelta)}</td>
        <td style="padding:7px 12px;border-bottom:1px solid #eee;${flagStyle(r.flag)}">${r.flag || '-'}</td>
      </tr>`).join('');

  const thStyle = 'background:#1a5276;color:white;padding:9px 12px;text-align:left;font-size:11px;font-weight:600;white-space:nowrap;';
  const thStyleR = thStyle + 'text-align:right;';

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#222;font-size:14px;margin:0;padding:0;">
<div style="max-width:900px;margin:0 auto;padding:20px;">

  <!-- Header -->
  <div style="background:#1a5276;color:white;padding:20px 24px;border-radius:10px 10px 0 0;">
    <h1 style="margin:0;font-size:20px;font-weight:600;">Daily Ads Snapshot</h1>
    <div style="font-size:12px;opacity:0.8;margin-top:4px;">Report: ${reportDate} &nbsp;·&nbsp; Data as of: ${dataAsOf}</div>
  </div>

  <!-- Flag summary -->
  <div style="background:#d6eaf8;border-left:4px solid #2e86c1;padding:10px 16px;margin-bottom:20px;font-size:13px;">
    ${flagSummary}
  </div>

  <!-- Summary cards -->
  ${summaryCards}

  <!-- Region Subtotals -->
  <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,0.08);margin-bottom:20px;">
    <h2 style="font-size:13px;font-weight:600;color:#1a5276;margin-bottom:14px;text-transform:uppercase;letter-spacing:0.5px;">By Region (Subtotals)</h2>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
      <thead><tr>
        <th style="${thStyle}">Region</th>
        <th style="${thStyleR}">Spend</th>
        <th style="${thStyleR}">Leads</th>
        <th style="${thStyleR}">CPL</th>
        <th style="${thStyleR}">CPL 7d Avg</th>
        <th style="${thStyleR}">CPL Δ vs 7d</th>
        <th style="${thStyle}">Flag</th>
      </tr></thead>
      <tbody>${regionTableRows}</tbody>
    </table>
  </div>

  <!-- Channel Detail (active only) -->
  ${channelRows ? `
  <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,0.08);margin-bottom:20px;">
    <h2 style="font-size:13px;font-weight:600;color:#1a5276;margin-bottom:14px;text-transform:uppercase;letter-spacing:0.5px;">Active Channels</h2>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
      <thead><tr>
        <th style="${thStyle}">Region</th>
        <th style="${thStyle}">Channel</th>
        <th style="${thStyleR}">Spend</th>
        <th style="${thStyleR}">Leads</th>
        <th style="${thStyleR}">CPL</th>
        <th style="${thStyleR}">CPL Δ vs 7d</th>
        <th style="${thStyle}">Flag</th>
      </tr></thead>
      <tbody>${channelRows}</tbody>
    </table>
  </div>` : ''}

  <div style="font-size:11px;color:#aaa;text-align:center;padding:12px 0;">
    Auto-generated by Marketing Automation · <a href="https://report.maqilnaufal.my.id/ads.html" style="color:#2e86c1;">View live dashboard</a>
  </div>
</div>
</body>
</html>`;
}

// ── Helpers ──────────────────────────────────────────────────

function parseNumSafe(v) {
  if (v === null || v === undefined || v === '' || v === '-') return 0;
  return parseInt(String(v).replace(/[^0-9-]/g, '')) || 0;
}

/**
 * Run once manually in Apps Script editor to create the daily 9am WIB trigger.
 */
function createDailyDigestTrigger() {
  // Remove any existing digest triggers first
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyAdsDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendDailyAdsDigest')
    .timeBased()
    .atHour(9)   // 9am script timezone (Asia/Jakarta)
    .everyDays(1)
    .create();

  Logger.log('Daily digest trigger created: 9am daily');
}
