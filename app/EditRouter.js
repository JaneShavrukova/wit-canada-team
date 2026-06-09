// ============================================================
// WIT Canada — Edit Router (single onEdit dispatcher)
// ============================================================
// ONE installable On-Edit trigger (onEditInstallable) instead of
// four. Previously each handler was its own trigger, so every cell
// edit fired all four and each independently rebuilt the column map
// (a full header-row read). Now the column map is built once per
// edit and passed to the handlers via `ctx`.
//
// Registered by app/Triggers.js → setupTriggers().
//
// Handlers receive (e, ctx) where ctx = { sheet, sheetName, row,
// col, colMap }. They remain individually callable (e.g. from the
// editor): if ctx is omitted they rebuild it from `e`.
// ============================================================

/**
 * Builds the shared per-edit context from an onEdit event.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @returns {{sheet: Sheet, sheetName: string, row: number, col: number, colMap: Object}}
 */
function buildEditContext(e) {
  const sheet = e.range.getSheet();
  return {
    sheet,
    sheetName: sheet.getName(),
    row:       e.range.getRow(),
    col:       e.range.getColumn(),
    colMap:    getColumnIndexMap(sheet),
  };
}

/**
 * Installable On-Edit entry point. Builds the context once and
 * dispatches to each handler in turn. Handlers are isolated: one
 * throwing does not prevent the others from running, but the first
 * error is re-thrown afterwards so failures still surface in the
 * execution log / failure notifications.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditInstallable(e) {
  const sheet     = e.range.getSheet();
  const sheetName = sheet.getName();

  // Only the member sheets carry handler logic — skip everything
  // else (report sheets, etc.) before paying for a header read.
  if (sheetName !== CONFIG.SHEET.MAIN && sheetName !== CONFIG.SHEET.EXTERNAL) return;

  const ctx = {
    sheet,
    sheetName,
    row:    e.range.getRow(),
    col:    e.range.getColumn(),
    colMap: getColumnIndexMap(sheet),
  };

  // Order note: the first three react to distinct values; member-status
  // promotion runs last so it sees the row's settled state.
  const handlers = [
    processEmailRequestOnEdit,
    processGroupsRequestOnEdit,
    processIntroSentOnEdit,
    processMemberStatusOnEdit,
  ];

  let firstError = null;
  for (const handler of handlers) {
    try {
      handler(e, ctx);
    } catch (err) {
      Logger.log(`onEditInstallable: ${handler.name} failed — ${err.message}`);
      if (!firstError) firstError = err;
    }
  }

  if (firstError) throw firstError;
}
