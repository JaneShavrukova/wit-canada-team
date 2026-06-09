// ============================================================
// WIT Canada — Report Builder
// ============================================================
// Shared scaffolding for the generated report sheets
// (Email Requests, Photos & Bios). Keeps the title banner,
// timestamp row, sheet (re)creation, and column widths in one
// place so the per-report builders only describe their own data.
//
// GroupsView builds its sheet differently (array batch write +
// row groupings, no merged cells) and only reuses reportTimestamp().
// ============================================================

/**
 * Standard "Updated: …" timestamp string used across report sheets.
 * @returns {string}
 */
function reportTimestamp() {
  return new Date().toLocaleString('en-CA', {
    timeZone:  'America/Vancouver',
    dateStyle: 'medium',
    timeStyle: 'short',
  });
}

/**
 * Returns a cleared report sheet by name, creating it if missing.
 * @param {string} name
 * @returns {Sheet}
 */
function getOrCreateReportSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  return sheet;
}

/**
 * Writes the two-row banner (title + "Updated: …") merged across
 * `colCount` columns, and returns the next writable row (3).
 *
 * @param {Sheet}  sheet
 * @param {string} title
 * @param {number} colCount
 * @param {string} [timestamp]  defaults to reportTimestamp()
 * @returns {number} next free row (1-indexed)
 */
function writeReportTitleBlock(sheet, title, colCount, timestamp) {
  const ts = timestamp || reportTimestamp();
  const C  = THEME.report;

  sheet.getRange(1, 1, 1, colCount).merge()
    .setValue(title)
    .setFontSize(14).setFontWeight('bold')
    .setBackground(C.TITLE_BG).setFontColor(C.TITLE_FG);

  sheet.getRange(2, 1, 1, colCount).merge()
    .setValue(`Updated: ${ts}`)
    .setFontSize(10)
    .setFontColor(C.SUBTITLE_FG).setBackground(C.SUBTITLE_BG);

  return 3;
}

/**
 * Sets column widths from a left-to-right array of pixel widths.
 * @param {Sheet}    sheet
 * @param {number[]} widths
 */
function setReportColumnWidths(sheet, widths) {
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}
