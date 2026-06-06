// ============================================================
// WIT Canada — Groups Sync
// ============================================================

function _syncAllMembers() {
  const sheet = getMainSheet();
  const colMap = getColumnIndexMap(sheet);

  for (let row = CONFIG.SHEET.DATA_START_ROW; row <= sheet.getLastRow(); row++) {
    syncMemberRow(sheet, row, colMap);
  }

  showAlert('Done', 'All members synced');
}

function _syncSelectedMember() {
  const sheet = getMainSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < CONFIG.SHEET.DATA_START_ROW) {
    showAlert('Error', 'Select a valid row');
    return;
  }

  const colMap = getColumnIndexMap(sheet);
  syncMemberRow(sheet, row, colMap);

  showAlert('Done', 'Member synced');
}


// ─────────────────────────────────────────────

function syncMemberRow(sheet, row, colMap) {
  const email = safeString(sheet.getRange(row, getCol(colMap, CONFIG.HEADERS.WIT_EMAIL)).getValue());
  if (!email) return;

  for (const [header, groupEmail] of Object.entries(CONFIG.GROUPS.HEADER_TO_EMAIL)) {
    const col = colMap[normalizeHeader(header)];
    if (!col) continue;

    const shouldBe = sheet.getRange(row, col).getValue() === true;

    if (shouldBe && !isMember(email, groupEmail)) {
      addToGroup(email, groupEmail);
    }
  }

  sheet.getRange(row, getCol(colMap, CONFIG.HEADERS.ADDED_TO_GROUPS)).setValue(CONFIG.GROUPS_STATUS.ADDED);
}


// ─────────────────────────────────────────────

function isMember(email, group) {
  try {
    AdminDirectory.Members.get(group, email);
    return true;
  } catch {
    return false;
  }
}

function addToGroup(email, group) {
  try {
    AdminDirectory.Members.insert({ email }, group);
  } catch (e) {
    Logger.log(e.message);
  }
}