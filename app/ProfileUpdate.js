// ============================================================
// WIT Canada — Profile Form Handler
// ============================================================
// Trigger: Google Form submit (installable, via createProfileFormTrigger)
//
// What it does:
//   When a member submits the profile update form, finds their row
//   in WIT_Members by WIT Email and fills in empty optional fields.
//   Overwrites existing sheet values if the member submits a new answer.
// ============================================================

/**
 * Resolves the submitter's WIT email from the manually entered form field.
 *
 * @param {Object} namedValues  e.namedValues from the form-submit event
 * @returns {string} lowercase email, or "" if the field was left blank
 */
function resolveWitEmail(namedValues) {
  return safeString(namedValues['Your WIT Email']?.[0]).toLowerCase();
}


/**
 * Main trigger handler for the profile update form.
 * Registered via createProfileFormTrigger().
 *
 * @param {GoogleAppsScript.Events.SpreadsheetsOnFormSubmit} e
 */
function handleProfileFormSubmit(e) {
  const namedValues = e.namedValues || {};

  // ── Extract form answers ─────────────────────────────────
  const getValue = (title) => {
    const arr = namedValues[title];
    return arr && arr.length ? safeString(arr[0]) : '';
  };

  // ── Resolve WIT email ────────────────────────────────────
  const witEmail = resolveWitEmail(namedValues);

  if (!witEmail) {
    Logger.log('handleProfileFormSubmit: submission has no WIT Email — skipping.');
    return;
  }

  if (!witEmail.endsWith('@women-in-tech.org')) {
    Logger.log(`handleProfileFormSubmit: "${witEmail}" is not a WIT address — skipping.`);
    return;
  }

  Logger.log(`handleProfileFormSubmit: processing submission for "${witEmail}"`);

  // ── Find matching row in WIT_Members ─────────────────────
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const emailCol  = getCol(colMap, CONFIG.HEADERS.WIT_EMAIL);
  const lastRow   = sheet.getLastRow();
  let   matchRow  = -1;

  for (let row = CONFIG.SHEET.DATA_START_ROW; row <= lastRow; row++) {
    const cell = safeString(sheet.getRange(row, emailCol).getValue()).toLowerCase();
    if (cell === witEmail) {
      matchRow = row;
      break;
    }
  }

  if (matchRow === -1) {
    Logger.log(`handleProfileFormSubmit: no row found for "${witEmail}" — sending alert to Ops.`);
    MailApp.sendEmail(
      CONFIG.EMAIL.OPS_LEAD,
      '⚠️ Profile Form: Member Not Found',
      `A profile update form was submitted for "${witEmail}" but no matching row ` +
      `was found in the ${CONFIG.SHEET.MAIN} sheet.\n\nPlease review manually.`
    );
    return;
  }

  Logger.log(`handleProfileFormSubmit: match found at row ${matchRow}`);

  // ── Field map: form question title → sheet column header ─
  const fieldMap = [
    { formTitle: 'Phone (WhatsApp)',  header: CONFIG.HEADERS.PHONE },
    { formTitle: 'LinkedIn URL',      header: CONFIG.HEADERS.LINKEDIN },
    { formTitle: 'Birthday', header: CONFIG.HEADERS.BIRTHDAY },
  ];

  // ── Write non-empty answers to sheet ─────────────────────
  let updatedCount = 0;

  fieldMap.forEach(({ formTitle, header }) => {
    const answer = getValue(formTitle);
    if (!answer) {
      Logger.log(`handleProfileFormSubmit: "${formTitle}" not answered — skipping.`);
      return;
    }

    let col;
    try {
      col = getCol(colMap, header);
    } catch (_) {
      Logger.log(`handleProfileFormSubmit: column "${header}" not found in sheet — skipping.`);
      return;
    }

    sheet.getRange(matchRow, col).setValue(answer);
    Logger.log(`handleProfileFormSubmit: wrote "${answer}" → "${header}" for ${witEmail}`);
    updatedCount++;
  });

  Logger.log(`handleProfileFormSubmit: done — ${updatedCount} field(s) updated for "${witEmail}"`);
}


// ─────────────────────────────────────────────
// Setup helpers (run once manually)
// ─────────────────────────────────────────────

/**
 * Registers the onFormSubmit trigger for handleProfileFormSubmit.
 * Run once from the Apps Script editor after linking the form.
 * Re-running is safe — removes any duplicate triggers first.
 */
function createProfileFormTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Remove duplicates before creating ────────────────────
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleProfileFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('handleProfileFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('createProfileFormTrigger: trigger registered for handleProfileFormSubmit.');
}

/**
 * Hides the "Form responses 2" sheet tab from view.
 * Run once after the form has been linked and the trigger created.
 */
function hideFormResponsesSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Form responses 2');

  if (!sheet) {
    Logger.log('hideFormResponsesSheet: "Form responses 2" tab not found — nothing to hide.');
    return;
  }

  sheet.hideSheet();
  Logger.log('hideFormResponsesSheet: "Form responses 2" tab is now hidden.');
}
