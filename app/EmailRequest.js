// ============================================================
// WIT Canada — Email Request OnEdit Handler
// ============================================================
// Trigger: From spreadsheet → On edit (installable trigger)
//
// What it does:
//   Watches the 'Email Status' column on the main sheet.
//   When a user manually sets status → 'requested', validates the row
//   and shows a confirmation dialog before proceeding.
//
// What it does NOT do anymore:
//   Send individual emails to Yevheniia. That is now handled
//   by sendWeeklyEmailRequestsReport() every Monday at 7 AM.
// ============================================================

function processEmailRequestOnEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET.MAIN) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  const colMap = getColumnIndexMap(sheet);

  // ── Guard: only watch the Request Status column ──────────
  const statusCol = colMap[normalizeHeader(CONFIG.HEADERS.REQUEST_STATUS)];
  if (!statusCol) throw new Error(`Column not found: ${CONFIG.HEADERS.REQUEST_STATUS}`);
  if (col !== statusCol) return;

  // ── Read new status value ────────────────────────────────
  const status = sheet.getRange(row, statusCol)
    .getValue()
    ?.toString()
    .trim()
    .toLowerCase();

  if (status !== CONFIG.STATUS.REQUEST) return; // only react to 'requested'

  // ── Resolve required columns ─────────────────────────────
  const firstNameCol   = colMap[normalizeHeader(CONFIG.HEADERS.FIRST_NAME)];
  const lastNameCol    = colMap[normalizeHeader(CONFIG.HEADERS.LAST_NAME)];
  const roleCol        = colMap[normalizeHeader(CONFIG.HEADERS.ROLE)];
  const emailCol       = colMap[normalizeHeader(CONFIG.HEADERS.PERSONAL_EMAIL)];

  if (!firstNameCol) throw new Error(`Column not found: ${CONFIG.HEADERS.FIRST_NAME}`);
  if (!lastNameCol)  throw new Error(`Column not found: ${CONFIG.HEADERS.LAST_NAME}`);
  if (!roleCol)      throw new Error(`Column not found: ${CONFIG.HEADERS.ROLE}`);
  if (!emailCol)     throw new Error(`Column not found: ${CONFIG.HEADERS.PERSONAL_EMAIL}`);

  // ── Read row data ────────────────────────────────────────
  const firstName      = sheet.getRange(row, firstNameCol).getValue();
  const lastName       = sheet.getRange(row, lastNameCol).getValue();
  const personalEmail  = sheet.getRange(row, emailCol).getValue();
  const role           = sheet.getRange(row, roleCol).getValue();

  // ── Validate required fields ─────────────────────────────
  if (!firstName || !lastName || !personalEmail) {
    if (CONFIG.UI.ALERTS) {
      showAlert(
        '⚠️ Missing Data',
        buildMissingFieldsMessage({ firstName, lastName, personalEmail })
      );
    }
    sheet.getRange(row, statusCol).setValue(e.oldValue ?? CONFIG.STATUS.NEW);
    return;
  }

  const WITEmail = generateWITEmail(firstName, lastName);
  const fullName = `${firstName} ${lastName}`;

  // ── Confirmation dialog ──────────────────────────────────
  if (CONFIG.UI.CONFIRMATION) {
    const ui       = SpreadsheetApp.getUi();
    const response = ui.alert(
      '📧 Request WIT Email',
      `Queue email creation request for:\n\n` +
      `Name:       ${fullName}\n` +
      `WIT Email:  ${WITEmail}\n\n` +
      `Ops Lead will be notified immediately.\n` +
      `Admin (Global Team) receives the batch email every Monday at 7 AM.`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      sheet.getRange(row, statusCol).setValue(e.oldValue ?? CONFIG.STATUS.NEW);
      return;
    }
  }

  // ── Instant notification to Ops Lead ────────────────────
  MailApp.sendEmail({
    to:      CONFIG.EMAIL.OPS_LEAD,
    subject: `[WIT Email Request] ${fullName}`,
    body:
      `Hi ${CONFIG.EMAIL.OPS_LEAD_NAME},\n\n` +
      `A WIT email has been requested for a new member.\n\n` +
      `Name:             ${fullName}\n` +
      `Role:             ${role}\n` +
      `Personal email:   ${personalEmail}\n` +
      `Suggested email:  ${WITEmail}\n\n` +
      `This member will also be included in the Monday batch report.\n\n` +
      `— WIT Automation`,
  });
  Logger.log(`processEmailRequestOnEdit: instant notification sent to Ops Lead for ${fullName}.`);

  // ── Confirm queued ───────────────────────────────────────
  if (CONFIG.UI.ALERTS) {
    showAlert(
      '✅ Queued',
      `${fullName} has been queued for email creation.\n\n` +
      `Ops Lead has been notified. Admin (Global Team) will receive the batch email on Monday at 7 AM.\n\n` +
      `Your next step: select required Google Groups and set Add to Groups = requested.`
    );
  }

  Logger.log(`processEmailRequestOnEdit: ${fullName} queued for weekly batch.`);
}


// generateWITEmail and buildMissingFieldsMessage live in Utils.gs.js