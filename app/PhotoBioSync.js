function syncPhotosAndBios() {
  const sheet  = getMainSheet();
  const colMap = getColumnIndexMap(sheet);

  const firstCol  = getCol(colMap, CONFIG.HEADERS.FIRST_NAME);
  const lastCol   = getCol(colMap, CONFIG.HEADERS.LAST_NAME);
  const photoCol  = getCol(colMap, CONFIG.HEADERS.PHOTO);
  const bioCol    = getCol(colMap, CONFIG.HEADERS.BIO);
  const statusCol = getCol(colMap, CONFIG.HEADERS.MEMBER_STATUS);

  const regionCols = {
    'Alberta':   getCol(colMap, CONFIG.REGIONS.ALBERTA),
    'BC':        getCol(colMap, CONFIG.REGIONS.BC),
    'Maritimes': getCol(colMap, CONFIG.REGIONS.MARITIMES),
    'Ontario':   getCol(colMap, CONFIG.REGIONS.ONTARIO),
    'Quebec':    getCol(colMap, CONFIG.REGIONS.QUEBEC),
  };

  const photos = getFileNames(CONFIG.DRIVE.PHOTO_FOLDER_ID);
  const bios   = getFileNames(CONFIG.DRIVE.BIO_FOLDER_ID);

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - CONFIG.SHEET.DATA_START_ROW + 1;
  if (numRows <= 0) return;

  // ── Single batch read — all rows at once ─────────────────
  const data = sheet
    .getRange(CONFIG.SHEET.DATA_START_ROW, 1, numRows, sheet.getMaxColumns())
    .getValues();

  const byRegion = {};
  Object.keys(regionCols).forEach(r => byRegion[r] = []);
  byRegion['Unknown'] = [];

  // Write buffers: one entry per row, preserving existing value for skipped rows
  const photoValues = [];
  const bioValues  = [];

  data.forEach(row => {
    const first  = safeString(row[firstCol  - 1]);
    const last   = safeString(row[lastCol   - 1]);
    const status = safeString(row[statusCol - 1]).toLowerCase();

    if ((!first && !last) || (status !== CONFIG.MEMBER_STATUS.ACTIVE && status !== CONFIG.MEMBER_STATUS.ONBOARDING)) {
      // Not in scope — preserve whatever is already in the cell
      photoValues.push([row[photoCol - 1]]);
      bioValues.push([row[bioCol  - 1]]);
      return;
    }

    const hasPhoto = fileExistsForMember(photos, first, last);
    const hasBio   = fileExistsForMember(bios,   first, last);

    photoValues.push([hasPhoto]);
    bioValues.push([hasBio]);

    const region = Object.entries(regionCols).find(
      ([_, col]) => row[col - 1] === true
    )?.[0] ?? 'Unknown';

    const fullName = `${first} ${last}`.trim();
    byRegion[region].push({ name: fullName, witEmail: generateWITEmail(first, last), hasPhoto, hasBio, status });

    Logger.log({ first, last, hasPhoto, hasBio, region, status });
  });

  // ── Two batch writes — entire columns at once ─────────────
  sheet.getRange(CONFIG.SHEET.DATA_START_ROW, photoCol, numRows, 1).setValues(photoValues);
  sheet.getRange(CONFIG.SHEET.DATA_START_ROW, bioCol,  numRows, 1).setValues(bioValues);

  saveSyncReport(byRegion);

  // Always rebuild the report sheet — works in both trigger and manual context
  buildPhotoBioReportSheet();

  // Try to show a UI alert — will silently skip if running from a time-based trigger
  const totalUnmatched = Object.values(byRegion)
    .flat()
    .filter(m => !m.hasPhoto || !m.hasBio).length;

  tryShowAlert(totalUnmatched);
}


/**
 * Attempts to show a UI alert.
 * Silently does nothing if called from a time-based trigger (no UI context).
 */
function tryShowAlert(totalUnmatched) {
  try {
    const ui = SpreadsheetApp.getUi();

    if (totalUnmatched === 0) {
      ui.alert('✅ Sync Complete', 'All members have matching photos and bios.', ui.ButtonSet.OK);
    } else {
      ui.alert(
        '✅ Sync Complete',
        `${totalUnmatched} member(s) need review.\n\nOpen "WIT Operations → Check photos & bios status" for details.`,
        ui.ButtonSet.OK
      );
    }
  } catch (_) {
    // Running from time-based trigger — UI is not available. This is expected.
    Logger.log('syncPhotosAndBios: UI not available (trigger context). Alert skipped.');
  }
}