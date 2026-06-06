// ============================================================
// WIT Canada — Member Status Auto-Promotion
// ============================================================
// Trigger: From spreadsheet → On edit (installable trigger)
//
// Watches: Contract Status, Email Status, Add to groups
// When all three conditions are met on the same row:
//   Contract Status  = 'signed'
//   Email Status     = 'active'
//   Add to groups  = 'added'
// → Automatically sets Member Status = 'active'
// ============================================================

function processMemberStatusOnEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET.MAIN) return;

  const colMap = getColumnIndexMap(sheet);

  const contractCol  = colMap[normalizeHeader(CONFIG.HEADERS.CONTRACT_STATUS)];
  const emailCol     = colMap[normalizeHeader(CONFIG.HEADERS.REQUEST_STATUS)];
  const groupsCol    = colMap[normalizeHeader(CONFIG.HEADERS.ADDED_TO_GROUPS)];
  const memberCol    = colMap[normalizeHeader(CONFIG.HEADERS.MEMBER_STATUS)];

  if (!contractCol || !emailCol || !groupsCol || !memberCol) return;

  // ── Guard: only react when one of the three watched columns is edited ──
  const editedCol = e.range.getColumn();
  const watchedCols = new Set([contractCol, emailCol, groupsCol]);
  if (!watchedCols.has(editedCol)) return;

  const row = e.range.getRow();
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Read current values ──────────────────────────────────
  const contract = safeString(sheet.getRange(row, contractCol).getValue()).toLowerCase();
  const email    = safeString(sheet.getRange(row, emailCol).getValue()).toLowerCase();
  const groups   = safeString(sheet.getRange(row, groupsCol).getValue()).toLowerCase();
  const current  = safeString(sheet.getRange(row, memberCol).getValue()).toLowerCase();

  // ── Skip if already active ───────────────────────────────
  if (current === CONFIG.MEMBER_STATUS.ACTIVE) return;

  // ── Check all three conditions ───────────────────────────
  const allDone =
    contract === CONFIG.CONTRACT.SIGNED &&
    email    === CONFIG.STATUS.ACTIVE &&
    groups   === CONFIG.GROUPS_STATUS.ADDED;

  if (!allDone) return;

  // ── Promote to active ────────────────────────────────────
  sheet.getRange(row, memberCol).setValue(CONFIG.MEMBER_STATUS.ACTIVE);

  Logger.log(`processMemberStatusOnEdit: row ${row} promoted to active.`);
}
