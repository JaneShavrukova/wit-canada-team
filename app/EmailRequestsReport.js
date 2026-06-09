// ============================================================
// WIT Canada — Email Requests Daily Report Builder
// ============================================================
// Trigger: Time-driven → Day timer → 3:00–4:00 AM PT
// ============================================================


/**
 * Builds (or rebuilds) the Email_Requests sheet with all members
 * who currently need a WIT email (status = 'requested', 'sent', 'created', 'not activated').
 * Called daily by time-based trigger.
 * Also called internally by sendWeeklyEmailRequestsReport() before sending.
 *
 * @returns {Array} rows — used by sendWeeklyEmailRequestsReport()
 */
function buildEmailRequestsReport() {
  const sheet  = getMainSheet();
  const colMap = getColumnIndexMap(sheet);

  const firstCol  = getCol(colMap, CONFIG.HEADERS.FIRST_NAME);
  const lastCol   = getCol(colMap, CONFIG.HEADERS.LAST_NAME);
  const roleCol   = getCol(colMap, CONFIG.HEADERS.ROLE);
  const emailCol  = getCol(colMap, CONFIG.HEADERS.PERSONAL_EMAIL);
  const statusCol = getCol(colMap, CONFIG.HEADERS.EMAIL_STATUS);

  // ── Collect pending members ──────────────────────────────
  // WIT_External is intentionally excluded — external members
  // are tracked for Google Groups only, not WIT email creation.
  const PENDING_STATUSES = new Set([
    CONFIG.STATUS.REQUEST,
    CONFIG.STATUS.SENT,
    CONFIG.STATUS.CREATED,
    CONFIG.STATUS.NOT_ACTIVATED,
  ]);
  const rows = [];

  for (let row = CONFIG.SHEET.DATA_START_ROW; row <= sheet.getLastRow(); row++) {
    const first = safeString(sheet.getRange(row, firstCol).getValue());
    const last  = safeString(sheet.getRange(row, lastCol).getValue());
    if (!first && !last) continue;

    const requestStatus = safeString(sheet.getRange(row, statusCol).getValue()).toLowerCase();
    if (!PENDING_STATUSES.has(requestStatus)) continue;

    const role          = safeString(sheet.getRange(row, roleCol).getValue());
    const personalEmail = safeString(sheet.getRange(row, emailCol).getValue());
    const witEmail      = generateWITEmail(first, last);

    rows.push([
      `${first} ${last}`,
      role          || '—',
      personalEmail || '—',
      requestStatus,
      witEmail,
    ]);
  }

  // ── Write to sheet ───────────────────────────────────────
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const report   = getOrCreateReportSheet(CONFIG.SHEET.REPORTS.EMAIL_REQUESTS);

  const COL_COUNT = 5;

  // Title + timestamp banner (rows 1–2); data begins at row 3's headers.
  writeReportTitleBlock(report, 'WIT Email Creation Requests', COL_COUNT);

  // Column headers
  const headers = ['Full Name', 'Role', 'Personal Email', 'Request Status', 'WIT Email (suggested)'];
  report.getRange(3, 1, 1, COL_COUNT)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(THEME.report.SECTION_BG)
    .setFontColor(THEME.report.SECTION_FG);

  // Data rows or empty state
  if (rows.length > 0) {
    report.getRange(4, 1, rows.length, COL_COUNT).setValues(rows);
    _applyStatusColors(report, rows, 4);
  } else {
    report.getRange(4, 1, 1, COL_COUNT).merge()
      .setValue('✅ No pending email requests')
      .setFontColor(THEME.success).setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  // Column widths
  setReportColumnWidths(report, [200, 220, 220, 130, 240]);

  report.setFrozenRows(3);

  ss.setActiveSheet(report);

  Logger.log(`${CONFIG.SHEET.REPORTS.EMAIL_REQUESTS} report built: ${rows.length} pending member(s).`);
  return rows;
}


/**
 * Applies background color to the Request Status column based on value.
 * request → yellow, sent → blue, created → green, not activated → pink/red
 *
 * @param {Sheet}   sheet
 * @param {Array[]} rows
 * @param {number}  startRow — 1-indexed row where data begins
 */
function _applyStatusColors(sheet, rows, startRow) {
  rows.forEach((row, i) => {
    const status = row[3]; // index 3 = Request Status
    const color  = THEME.status[status]?.bg ?? THEME.surface;
    sheet.getRange(startRow + i, 4).setBackground(color);
  });
}
