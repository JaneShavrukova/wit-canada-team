// ============================================================
// WIT Canada — Groups Sync  [DISABLED — not deployed]
// ============================================================
// Automatically adds members to Google Groups via the Admin
// Directory API. This requires the **Groups Admin** role in Google
// Workspace, which the current operator does NOT have, so the whole
// feature is parked here and excluded from `clasp push` (see
// .claspignore → `disabled/`). It was never wired into the menu.
//
// The ACTIVE groups workflow lives elsewhere and needs no admin API:
//   • app/GroupsAccessRequest.js — emails the Ops Lead to assign groups
//   • app/GroupsView.js          — builds the Report_Groups snapshot
//
// To re-enable when admin access is available:
//   1. Move this file back to app/GroupsSync.js.
//   2. Re-add the AdminDirectory advanced service + the two
//      admin.directory.group / admin.directory.group.member scopes
//      to appsscript.json.
//   3. Add menu items for syncAllMembers / syncSelectedMember in
//      app/Menu.js.
// ============================================================


// ── Menu wrappers (permission-aware) ───────────────────────

function syncAllMembers() {
  runWithAlert(() => {
    try {
      _syncAllMembers();
    } catch (e) {
      if (isPermissionError(e)) {
        showPermissionAlert();
      } else {
        throw e;
      }
    }
  });
}

function syncSelectedMember() {
  runWithAlert(() => {
    try {
      _syncSelectedMember();
    } catch (e) {
      if (isPermissionError(e)) {
        showPermissionAlert();
      } else {
        throw e;
      }
    }
  });
}

function isPermissionError(e) {
  return e.message && e.message.toLowerCase().includes('permission');
}

function showPermissionAlert() {
  SpreadsheetApp.getUi().alert(
    'Permission required',
    'Assigning group access requires the Groups Admin role in Google Workspace.\n\n' +
    'Please contact the Ops Lead — she will run the sync for you:\n' +
    CONFIG.EMAIL.OPS_LEAD,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ── Sync ───────────────────────────────────────────────────

function _syncAllMembers() {
  const sheet = getMainSheet();
  const colMap = getColumnIndexMap(sheet);

  const failures = [];
  for (let row = CONFIG.SHEET.DATA_START_ROW; row <= sheet.getLastRow(); row++) {
    const result = syncMemberRow(sheet, row, colMap);
    if (result.failed.length > 0) {
      failures.push(`${result.email}: ${result.failed.join(', ')}`);
    }
  }

  if (failures.length > 0) {
    showAlert(
      '⚠️ Synced with errors',
      `These members could not be fully added (rows NOT marked "added"):\n\n` +
      `${failures.join('\n')}\n\nSee the execution log for details.`
    );
  } else {
    showAlert('Done', 'All members synced');
  }
}

function _syncSelectedMember() {
  const sheet = getMainSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < CONFIG.SHEET.DATA_START_ROW) {
    showAlert('Error', 'Select a valid row');
    return;
  }

  const colMap = getColumnIndexMap(sheet);
  const result = syncMemberRow(sheet, row, colMap);

  if (result.failed.length > 0) {
    showAlert(
      '⚠️ Partial sync',
      `Could not add ${result.email || 'this member'} to:\n` +
      `• ${result.failed.join('\n• ')}\n\n` +
      `Row was NOT marked "added". See the execution log for details.`
    );
  } else {
    showAlert('Done', 'Member synced');
  }
}


// ─────────────────────────────────────────────

/**
 * Adds the member to every checked group they're not already in, and marks
 * the row "added" ONLY if nothing failed. Returns a per-row report.
 *
 * @returns {{email: string, added: string[], failed: string[]}}
 */
function syncMemberRow(sheet, row, colMap) {
  const email = safeString(sheet.getRange(row, getCol(colMap, CONFIG.HEADERS.WIT_EMAIL)).getValue());
  if (!email) return { email: '', added: [], failed: [] };

  const added  = [];
  const failed = [];

  for (const [header, groupEmail] of Object.entries(CONFIG.GROUPS.HEADER_TO_EMAIL)) {
    const col = colMap[normalizeHeader(header)];
    if (!col) continue;

    const shouldBe = sheet.getRange(row, col).getValue() === true;
    if (!shouldBe) continue;
    if (isMember(email, groupEmail)) continue; // already a member — nothing to do

    if (addToGroup(email, groupEmail)) {
      added.push(header);
    } else {
      failed.push(header);
    }
  }

  // Only claim "added" when every required insert succeeded. A partial
  // failure leaves the cell in its prior state so the row stays visibly
  // incomplete instead of falsely reading "added".
  if (failed.length === 0) {
    sheet.getRange(row, getCol(colMap, CONFIG.HEADERS.ADDED_TO_GROUPS)).setValue(CONFIG.GROUPS_STATUS.ADDED);
  }

  return { email, added, failed };
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

/**
 * Adds a member to a group. Returns true on success, false on a per-group
 * failure (logged). Permission errors are global (no Groups Admin role), so
 * they are re-thrown — the menu wrapper turns them into the "contact the Ops
 * Lead" alert and stops, rather than silently failing every group.
 *
 * @returns {boolean}
 */
function addToGroup(email, group) {
  try {
    AdminDirectory.Members.insert({ email }, group);
    return true;
  } catch (e) {
    if (isPermissionError(e)) throw e;
    Logger.log(`addToGroup: failed to add ${email} to ${group} — ${e.message}`);
    return false;
  }
}
