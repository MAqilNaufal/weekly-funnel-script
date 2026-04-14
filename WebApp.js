// ============================================================
// WEB APP — serves the Weekly Funnel Report as a dashboard
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Weekly Funnel Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Reads the already-populated Weekly Funnel Report sheet
 * and returns structured JSON for the dashboard.
 */
function getReportData() {
  const ss    = SpreadsheetApp.openById(CONFIG.SS_MAIN);
  const sheet = ss.getSheetByName(CONFIG.SHEET_REPORT);

  if (!sheet) return { error: 'Weekly Funnel Report sheet not found. Run the report first.' };

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return { error: 'No data. Run the report first.' };

  const result = {
    title:    String(data[0][0] || ''),
    weekLabel: String(data[0][1] || ''),
    generated: String(data[0][2] || ''),
    summary:  [],
    detail:   [],
    regions:  [],
  };

  let section = '';
  let detailHeaders = [];
  let regionHeaders = [];

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const first = String(row[0]).trim();

    if (first === 'SUMMARY SNAPSHOT')           { section = 'summary_header'; continue; }
    if (first === 'AD / MARKETING CODE DETAIL') { section = 'detail_header';  continue; }
    if (first === 'BREAKDOWN BY PROJECT / TEAM'){ section = 'region_header';  continue; }

    if (section === 'summary_header') {
      // Next row after label is the header row — skip it
      section = 'summary';
      continue;
    }
    if (section === 'detail_header') {
      detailHeaders = row.map(v => String(v));
      section = 'detail';
      continue;
    }
    if (section === 'region_header') {
      regionHeaders = row.map(v => String(v));
      section = 'region';
      continue;
    }

    if (section === 'summary' && first) {
      result.summary.push({ metric: row[0], tw: row[1], lw: row[2], delta: row[3] });
    }
    if (section === 'detail' && first) {
      const obj = {};
      detailHeaders.forEach((h, idx) => { obj[h] = row[idx]; });
      result.detail.push(obj);
    }
    if (section === 'region' && first) {
      const obj = {};
      regionHeaders.forEach((h, idx) => { obj[h] = row[idx]; });
      result.regions.push(obj);
    }
  }

  return result;
}

/**
 * Triggers report regeneration. Returns status message.
 */
function runReport() {
  try {
    generateWeeklyFunnelReport();
    return { success: true, message: 'Report generated successfully.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
