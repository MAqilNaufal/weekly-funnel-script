// ============================================================
// CREATIVES PERFORMANCE REPORT — Module 2
//
// Reads Master Data (DO NOT TOUCH) and aggregates per-ad
// spend/leads/CPL for current vs prior ISO week.
//
// Writes output to "Creatives Performance" tab in SS_ADS.
// Also joined into the Monday weekly email via generateWeeklyFunnelReport().
//
// Standalone call: generateCreativesReport()
// ============================================================

/**
 * Main entry point.
 * Returns structured data object (also writes to sheet).
 */
function generateCreativesReport() {
  const today     = new Date();
  const curWeek   = getISOWeek(today);
  const curYear   = today.getFullYear();

  // Prior week (handles year wrap)
  let priorWeek = curWeek - 1;
  let priorYear = curYear;
  if (priorWeek < 1) { priorWeek = 52; priorYear = curYear - 1; }

  Logger.log(`Creatives Report: Week ${curWeek}/${curYear} vs ${priorWeek}/${priorYear}`);

  // ── Read Master Data ──────────────────────────────────────
  const adsSS      = SpreadsheetApp.openById(CONFIG.SS_ADS);
  const masterSheet = adsSS.getSheetByName('Master Data (DO NOT TOUCH)');
  if (!masterSheet) throw new Error('Master Data (DO NOT TOUCH) sheet not found');

  const masterData = masterSheet.getDataRange().getValues();
  if (masterData.length < 2) return { rows: [], curWeek, priorWeek };

  const hdrs = masterData[0].map(h => String(h).trim());
  const COL  = {};
  hdrs.forEach((h, i) => { COL[h] = i; });

  // Required column indices
  const iYear    = COL['Year'];
  const iWeek    = COL['Week'];
  const iChannel = COL['Channel'];
  const iAdName  = COL['Ad Name'];
  const iMktCode = COL['Marketing Code'];
  const iSpend   = COL['Spends'];
  const iResult  = COL['Result'];
  const iCPL     = COL['CPL'];
  const iRegion  = COL['Region'];
  const iCTR     = COL['CTR'];
  const iImpr    = COL['Impressions'];
  const iClicks  = COL['Clicks'];

  // ── Read Creative Database for links ─────────────────────
  const creativeLinkMap = buildCreativeLinkMap();

  // ── Aggregate by AdName + Region + Channel + Week ────────
  const map = {};

  for (let i = 1; i < masterData.length; i++) {
    const row  = masterData[i];
    const yr   = Number(row[iYear]);
    const wk   = Number(row[iWeek]);

    const isCur   = (yr === curYear   && wk === curWeek);
    const isPrior = (yr === priorYear && wk === priorWeek);
    if (!isCur && !isPrior) continue;

    const adName  = String(row[iAdName]  || '').trim();
    const region  = String(row[iRegion]  || '').trim();
    const channel = String(row[iChannel] || '').trim();
    if (!adName || !region || !channel) continue;

    // Extract marketing code: use existing column first, else regex
    let mktCode = String(row[iMktCode] || '').trim();
    if (!mktCode) mktCode = extractMarketingCode(adName);

    const key = `${adName}||${region}||${channel}`;
    if (!map[key]) {
      map[key] = {
        adName, region, channel, mktCode,
        curSpend: 0, curLeads: 0, curImpr: 0, curClicks: 0,
        priorSpend: 0, priorLeads: 0,
        creativeLink: creativeLinkMap[adName] || '',
      };
    }

    const spend  = parseFloat(String(row[iSpend]  || '0').replace(/[^0-9.-]/g,'')) || 0;
    const leads  = parseFloat(String(row[iResult] || '0').replace(/[^0-9.-]/g,'')) || 0;
    const impr   = parseFloat(String(row[iImpr]   || '0').replace(/[^0-9.-]/g,'')) || 0;
    const clicks = parseFloat(String(row[iClicks] || '0').replace(/[^0-9.-]/g,'')) || 0;

    if (isCur) {
      map[key].curSpend  += spend;
      map[key].curLeads  += leads;
      map[key].curImpr   += impr;
      map[key].curClicks += clicks;
    } else {
      map[key].priorSpend += spend;
      map[key].priorLeads += leads;
    }
  }

  // ── Build sorted rows ─────────────────────────────────────
  const rows = Object.values(map).map(r => {
    const curCPL   = r.curLeads   > 0 ? Math.round(r.curSpend   / r.curLeads)   : (r.curSpend > 0 ? null : 0);
    const priorCPL = r.priorLeads > 0 ? Math.round(r.priorSpend / r.priorLeads) : null;
    const ctr      = r.curImpr > 0 ? (r.curClicks / r.curImpr * 100).toFixed(2) + '%' : '-';

    let cplDelta = '-';
    if (curCPL !== null && priorCPL !== null && priorCPL > 0) {
      const pct = Math.round((curCPL - priorCPL) / priorCPL * 100);
      cplDelta  = (pct >= 0 ? '+' : '') + pct + '%';
    }

    return { ...r, curCPL, priorCPL, cplDelta, ctr };
  });

  // Sort: active this week first, then by CPL ascending (best performers first)
  rows.sort((a, b) => {
    const aActive = a.curSpend > 0 ? 0 : 1;
    const bActive = b.curSpend > 0 ? 0 : 1;
    if (aActive !== bActive) return aActive - bActive;
    const aCPL = a.curCPL ?? 999999999;
    const bCPL = b.curCPL ?? 999999999;
    return aCPL - bCPL;
  });

  // ── Write to sheet ────────────────────────────────────────
  writeCreativesSheet(adsSS, rows, curWeek, curYear, priorWeek, priorYear);

  return { rows, curWeek, priorWeek, curYear, priorYear };
}

// ── Sheet writer ─────────────────────────────────────────────

function writeCreativesSheet(ss, rows, curWeek, curYear, priorWeek, priorYear) {
  let sheet = ss.getSheetByName('Creatives Performance');
  if (!sheet) {
    sheet = ss.insertSheet('Creatives Performance');
  } else {
    sheet.clearContents();
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy HH:mm');
  const curLabel   = `W${curWeek}/${curYear}`;
  const priorLabel = `W${priorWeek}/${priorYear}`;

  const headers = [
    'Ad Name', 'Region', 'Channel', 'Marketing Code',
    `Spend ${curLabel}`, `Leads ${curLabel}`, `CPL ${curLabel}`, 'CTR (TW)',
    `Spend ${priorLabel}`, `Leads ${priorLabel}`, `CPL ${priorLabel}`,
    'CPL Δ (WoW)', 'Creative Link',
    'Generated',
  ];

  const output = [
    [`Creatives Performance Report — ${curLabel} vs ${priorLabel}`, ...Array(headers.length - 1).fill('')],
    [`Generated: ${now}`, ...Array(headers.length - 1).fill('')],
    headers,
  ];

  rows.forEach(r => {
    output.push([
      r.adName, r.region, r.channel, r.mktCode,
      r.curSpend, r.curLeads, r.curCPL ?? '', r.ctr,
      r.priorSpend, r.priorLeads, r.priorCPL ?? '',
      r.cplDelta, r.creativeLink,
      now,
    ]);
  });

  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  // Formatting
  sheet.getRange(1, 1, 2, headers.length).merge();
  sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold').setBackground('#1a5276').setFontColor('white');
  sheet.getRange(3, 1, 1, headers.length).setBackground('#1a5276').setFontColor('white').setFontWeight('bold');

  Logger.log(`Creatives Performance sheet updated: ${rows.length} rows`);
}

// ── Build creative link lookup map ───────────────────────────
// Joins by matching Ad Name → Creative ID in the Creative Database.
// Creative IDs typically follow the format: theme_format_version
// Ad Names may contain the creative ID as a substring — we do a
// best-effort partial match.

function buildCreativeLinkMap() {
  const map = {};
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SS_CREATIVE);
    const sheet = ss.getSheetByName(CONFIG.SHEET_CREATIVE_DB) || ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) return map;

    const hdrs    = data[0].map(h => String(h).trim());
    const idIdx   = hdrs.indexOf('Creative ID');
    const linkIdx = hdrs.indexOf('Link Creative');
    if (idIdx < 0 || linkIdx < 0) return map;

    for (let i = 1; i < data.length; i++) {
      const creativeId = String(data[i][idIdx]  || '').trim().toLowerCase();
      const link       = String(data[i][linkIdx] || '').trim();
      if (creativeId && link) map[creativeId] = link;
    }
  } catch (e) {
    Logger.log('Creative link map error: ' + e.message);
  }
  return map;
}

// ── HTML section for weekly email ───────────────────────────

function buildCreativesEmailSection(rows, curWeek, priorWeek, curYear, priorYear) {
  if (!rows || rows.length === 0) return '';

  const curLabel   = `W${curWeek}/${curYear}`;
  const priorLabel = `W${priorWeek}/${priorYear}`;

  const thStyle  = 'background:#1a5276;color:white;padding:8px 10px;text-align:left;font-size:11px;font-weight:600;white-space:nowrap;';
  const thStyleR = thStyle + 'text-align:right;';

  const fmt = n => (n !== null && n !== undefined && n !== '') ? Number(n).toLocaleString('id-ID') : '-';

  const deltaStyle = val => {
    if (!val || val === '-') return 'color:#999;';
    const n = parseFloat(String(val).replace('%','').replace('+',''));
    if (isNaN(n)) return 'color:#999;';
    return n < 0 ? 'color:#1e8449;font-weight:600;' : n > 0 ? 'color:#c0392b;font-weight:600;' : 'color:#999;';
  };

  // Top 15 active creatives this week
  const active = rows.filter(r => r.curSpend > 0).slice(0, 15);

  const tbody = active.map(r => `
    <tr>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;font-size:11px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${r.adName}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;">${r.region}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;">${r.channel}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;">${fmt(r.curSpend)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;">${r.curLeads}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;">${fmt(r.curCPL)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;${deltaStyle(r.cplDelta)}">${r.cplDelta}</td>
    </tr>`).join('');

  return `
    <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,0.08);margin-bottom:20px;">
      <h2 style="font-size:13px;font-weight:600;color:#1a5276;margin-bottom:14px;text-transform:uppercase;letter-spacing:0.5px;">
        Top Active Creatives — ${curLabel} <span style="font-size:11px;color:#999;font-weight:400;">(vs ${priorLabel})</span>
      </h2>
      <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <thead><tr>
          <th style="${thStyle}">Ad Name</th>
          <th style="${thStyle}">Region</th>
          <th style="${thStyle}">Channel</th>
          <th style="${thStyleR}">Spend (TW)</th>
          <th style="${thStyleR}">Leads (TW)</th>
          <th style="${thStyleR}">CPL (TW)</th>
          <th style="${thStyleR}">CPL Δ</th>
        </tr></thead>
        <tbody>${tbody}</tbody>
      </table>
      ${rows.length > 15 ? `<p style="font-size:11px;color:#999;margin-top:10px;">Showing top 15 of ${rows.filter(r=>r.curSpend>0).length} active creatives. <a href="https://report.maqilnaufal.my.id/creatives.html" style="color:#2e86c1;">View all →</a></p>` : ''}
    </div>`;
}
