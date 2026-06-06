// ============================================================
// WIT Canada — Groups Access Request Notification
// ============================================================
// Trigger: From spreadsheet → On edit (installable trigger)
//
// Watches: 'Add to groups' column on both:
//   • WIT_Members  (internal members)
//   • WIT_External      (external members)
//
// When value changes to 'requested':
//   → Sends an email to the Ops Lead with member details
//      and the list of groups to assign.
// ============================================================

const WATCHED_SHEETS = new Set([CONFIG.SHEET.MAIN, "WIT_External"]);

function processGroupsRequestOnEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (!WATCHED_SHEETS.has(sheetName)) return;

  const colMap = getColumnIndexMap(sheet);
  const groupsCol = colMap[normalizeHeader(CONFIG.HEADERS.ADDED_TO_GROUPS)];
  if (!groupsCol) {
    Logger.log(`processGroupsRequestOnEdit: column "${CONFIG.HEADERS.ADDED_TO_GROUPS}" not found in sheet "${sheetName}". Trigger skipped.`);
    return;
  }

  // ── Guard: only react to the 'Add to groups' column ────
  if (e.range.getColumn() !== groupsCol) return;

  // ── Guard: only react when value transitions to 'requested'
  if (safeString(e.oldValue).toLowerCase() === CONFIG.GROUPS_STATUS.REQUESTED) return;
  const newValue = safeString(e.range.getValue()).toLowerCase();
  if (newValue !== CONFIG.GROUPS_STATUS.REQUESTED) return;

  const row = e.range.getRow();
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Read member info ──────────────────────────────────────
  const firstCol = colMap[normalizeHeader(CONFIG.HEADERS.FIRST_NAME)];
  const lastCol = colMap[normalizeHeader(CONFIG.HEADERS.LAST_NAME)];
  const roleCol = colMap[normalizeHeader(CONFIG.HEADERS.ROLE)];

  const firstName = firstCol
    ? safeString(sheet.getRange(row, firstCol).getValue())
    : "";
  const lastName = lastCol
    ? safeString(sheet.getRange(row, lastCol).getValue())
    : "";
  const role = roleCol
    ? safeString(sheet.getRange(row, roleCol).getValue())
    : "—";
  const fullName = `${firstName} ${lastName}`.trim() || "(unknown)";

  // ── Collect requested groups ──────────────────────────────
  const requestedGroups = [];

  for (const [groupName, groupEmail] of Object.entries(CONFIG.GROUPS.HEADER_TO_EMAIL)) {
    const col = colMap[normalizeHeader(groupName)];
    if (!col) {
      Logger.log(`processGroupsRequestOnEdit: group column "${groupName}" not found in sheet "${sheetName}".`);
      continue;
    }
    if (sheet.getRange(row, col).getValue() === true) {
      requestedGroups.push({ name: groupName, email: groupEmail });
    }
  }

  const isExternal = sheetName !== CONFIG.SHEET.MAIN;
  const memberType = isExternal ? "External member" : "Internal member";

  // ── Confirmation dialog ───────────────────────────────────
  const ui = SpreadsheetApp.getUi();
  const revert = () => e.range.setValue(e.oldValue ?? CONFIG.GROUPS_STATUS.DEFAULT);

  if (requestedGroups.length === 0) {
    ui.alert(
      '⚠️ No groups selected',
      `No groups are checked for ${fullName}.\n\n` +
      `Please tick the group checkboxes for this member first, ` +
      `then set "Add to groups" back to "requested".`,
      ui.ButtonSet.OK
    );
    revert();
    return;
  }

  const groupList = requestedGroups.map(g => `• ${g.name}`).join('\n');
  const response = ui.alert(
    '👥 Confirm groups request',
    `Send groups access request for ${fullName}?\n\n` +
    `Groups to assign:\n${groupList}\n\n` +
    `Ops Lead will be notified by email.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    revert();
    return;
  }

  // ── Send notification to Ops Lead ─────────────────────────
  try {
    MailApp.sendEmail({
      to: CONFIG.EMAIL.OPS_LEAD,
      subject: `Groups access requested — ${fullName}`,
      htmlBody: _buildGroupsRequestEmailHtml(
        fullName,
        role,
        memberType,
        sheetName,
        requestedGroups,
      ),
      name: "WIT Canada — Onboarding",
    });

    Logger.log(
      `processGroupsRequestOnEdit: groups request sent to ${CONFIG.EMAIL.OPS_LEAD} for ${fullName} (${sheetName}). Groups: ${requestedGroups.map(g => g.name).join(', ') || 'none'}.`,
    );
  } catch (err) {
    Logger.log(`processGroupsRequestOnEdit: failed to send email for ${fullName} — ${err.message}`);
    throw err;
  }
}

// ─────────────────────────────────────────────────────────────
// Email template
// ─────────────────────────────────────────────────────────────

function _buildGroupsRequestEmailHtml(
  fullName,
  role,
  memberType,
  sheetName,
  groups,
) {
  const groupRows =
    groups.length > 0
      ? groups
          .map(
            (g) => `
        <tr>
          <td style="padding:7px 10px;border-bottom:1px solid #e8ecf0;font-size:13px;color:#2c3e50;">${g.name}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #e8ecf0;font-size:13px;color:#1b4f8a;">${g.email}</td>
        </tr>`,
          )
          .join("")
      : `<tr><td colspan="2" style="padding:10px;font-size:13px;color:#adb5bd;font-style:italic;">No groups selected</td></tr>`;

  return `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr>
      <td align="center">
        <table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.07);">

          <!-- Header -->
          <tr>
            <td style="background:#1b4f8a;padding:24px 32px;">
              <p style="margin:0;font-size:11px;color:rgba(255,255,255,0.6);text-transform:uppercase;letter-spacing:0.08em;">WIT Canada · Operations</p>
              <h1 style="margin:6px 0 0;font-size:18px;color:#ffffff;font-weight:bold;">Groups Access Requested</h1>
            </td>
          </tr>

          <!-- Member info -->
          <tr>
            <td style="padding:24px 32px 16px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f8fd;border-radius:6px;border:1px solid #dce8f8;">
                <tr>
                  <td style="padding:14px 16px;">
                    <p style="margin:0 0 4px;font-size:11px;color:#7f8c8d;text-transform:uppercase;letter-spacing:0.06em;">${memberType} · ${sheetName}</p>
                    <p style="margin:0 0 2px;font-size:15px;font-weight:bold;color:#2c3e50;">${fullName}</p>
                    <p style="margin:0;font-size:13px;color:#7f8c8d;">${role}</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Groups table -->
          <tr>
            <td style="padding:0 32px 24px;">
              <p style="margin:0 0 10px;font-size:13px;font-weight:bold;color:#1b4f8a;">Groups to assign</p>
              <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e8ecf0;border-radius:6px;overflow:hidden;">
                <thead>
                  <tr style="background:#1b4f8a;">
                    <th style="padding:8px 10px;text-align:left;font-size:11px;color:#fff;font-weight:bold;">Group</th>
                    <th style="padding:8px 10px;text-align:left;font-size:11px;color:#fff;font-weight:bold;">Email</th>
                  </tr>
                </thead>
                <tbody>${groupRows}</tbody>
              </table>
            </td>
          </tr>

          <!-- Divider -->
          <tr>
            <td style="padding:0 32px;">
              <hr style="border:none;border-top:1px solid #e8ecf0;margin:0;">
            </td>
          </tr>

          <!-- Action note -->
          <tr>
            <td style="padding:20px 32px 28px;">
              <p style="margin:0;font-size:13px;color:#7f8c8d;line-height:1.6;">
                Please add this member to the listed groups in Google Admin, then update
                <strong>Add to Groups = added</strong> in the sheet.
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#f8f9fa;padding:14px 32px;border-top:1px solid #e8ecf0;">
              <p style="margin:0;font-size:11px;color:#adb5bd;text-align:center;">
                Women in Tech Canada · This is an automated message
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>`;
}
