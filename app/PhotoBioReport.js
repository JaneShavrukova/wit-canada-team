// ============================================================
// WIT Canada — Photos & Bios Report
// ============================================================
// Region order and the PropertiesService key live in CONFIG
// (CONFIG.REGION_ORDER, CONFIG.PROPERTIES.SYNC_REPORT).


// ─────────────────────────────────────────────
// Save report to PropertiesService
// ─────────────────────────────────────────────

function saveSyncReport(byRegion) {
  const timestamp = reportTimestamp();

  PropertiesService
    .getScriptProperties()
    .setProperty(CONFIG.PROPERTIES.SYNC_REPORT, JSON.stringify({ timestamp, byRegion }));
}


// ─────────────────────────────────────────────
// Load report from PropertiesService
// ─────────────────────────────────────────────

function loadSyncReport() {
  const raw = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.SYNC_REPORT);

  if (!raw) return null;
  return JSON.parse(raw);
}


// ─────────────────────────────────────────────
// Report sheet — full view with formatting
// ─────────────────────────────────────────────

function buildPhotoBioReportSheet() {
  const data   = loadSyncReport();
  const report = getOrCreateReportSheet(CONFIG.SHEET.REPORTS.PHOTO_BIO);

  if (!data) {
    report.getRange(1, 1).setValue('No sync data yet. Run "Update photos & bios data" first.');
    return;
  }

  const { timestamp, byRegion } = data;
  const COL_COUNT = 4;

  // Title + timestamp banner (rows 1–2); content begins at row 3.
  let currentRow = writeReportTitleBlock(report, 'Photos & Bios Status Report', COL_COUNT, timestamp);

  CONFIG.REGION_ORDER.forEach(region => {
    const members = (byRegion[region] || []).filter(m => !m.hasPhoto || !m.hasBio);
    if (members.length === 0) return;

    // Region header
    report.getRange(currentRow, 1, 1, COL_COUNT).merge()
      .setValue(region.toUpperCase())
      .setFontSize(10)
      .setFontWeight('bold')
      .setBackground(THEME.report.SECTION_BG)
      .setFontColor(THEME.report.SECTION_FG);
    currentRow++;

    // Column headers
    report.getRange(currentRow, 1, 1, COL_COUNT).setBackground(THEME.surfaceMuted).setFontWeight('bold');
    report.getRange(currentRow, 1).setValue('Member');
    report.getRange(currentRow, 2).setValue('WIT Email');
    report.getRange(currentRow, 3).setValue('Photo');
    report.getRange(currentRow, 4).setValue('Bio');
    currentRow++;

    // Member rows
    members.forEach(m => {
      report.getRange(currentRow, 1).setValue(m.name);
      report.getRange(currentRow, 2).setValue(m.witEmail);
      report.getRange(currentRow, 3).setValue(m.hasPhoto ? '✅' : '❌');
      report.getRange(currentRow, 4).setValue(m.hasBio   ? '✅' : '❌');
      currentRow++;
    });

    currentRow++; // gap between regions
  });

  // Column widths
  setReportColumnWidths(report, [200, 220, 80, 80]);

  report.setFrozenRows(2);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(report);
}