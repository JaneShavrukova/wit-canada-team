// ============================================================
// WIT Canada — Weekly Email Report to Tamuna
// ============================================================
// Trigger: Time-driven → Week timer → Every Monday → 7:00–8:00 AM PT
// ============================================================

/**
 * Sends a weekly batch summary of all pending WIT email creation requests.
 * After successful send, updates all 'requested' rows → 'sent' in the main sheet.
 *
 * Flow:
 *   1. Rebuild Email_Requests sheet (fresh data)
 *   2. If nothing pending → log and exit, no email sent
 *   3. Build HTML email and send to RECIPIENT
 *   4. Mark all 'requested' rows → 'sent' in main sheet
 */
function sendWeeklyEmailRequestsReport() {
  const rows = buildEmailRequestsReport(); // always rebuild fresh before sending
  const T = THEME; // inline tokens (emails can't use CSS variables)

  // rows includes requested + sent + created + not activated (full tracker view)
  const newCount          = rows.filter(r => r[3] === CONFIG.STATUS.REQUEST).length;
  const notActivatedCount = rows.filter(r => r[3] === CONFIG.STATUS.NOT_ACTIVATED).length;
  const pendingCount      = newCount + notActivatedCount;

  if (pendingCount === 0) {
    Logger.log('Weekly report: no new or not-activated requests this week. Email not sent.');
    return;
  }

  const timestamp = new Date().toLocaleString('en-CA', {
    timeZone:  'America/Vancouver',
    dateStyle: 'full',
    timeStyle: 'short',
  });

  const introHtml =
    `<p>Here is the weekly summary of WIT email creation requests as of <strong>${timestamp}</strong>:</p>`;

  const afterTableHtml = `
      <p style="color:${T.textMuted}; font-size:13px;">
        New this week: <strong>${newCount}</strong> &nbsp;|&nbsp; Not activated: <strong>${notActivatedCount}</strong> &nbsp;|&nbsp; Total tracked: <strong>${rows.length}</strong>
      </p>

      <p>Please process at your convenience. Let me know if you have any questions.</p>

      <p>Best regards,<br>
      <strong>WIT Canada Operations</strong> (automated weekly report)</p>`;

  // Shared renderer → the single-member "send now" action reuses the same layout.
  const htmlBody = _renderEmailRequestsEmailHtml('Tamuna', introHtml, rows, afterTableHtml);

  MailApp.sendEmail({
    to:       CONFIG.EMAIL.OPS_LEAD,
    subject:  `WIT Email Requests — Weekly Summary (${newCount} new, ${notActivatedCount} not activated)`,
    htmlBody: htmlBody,
  });

  Logger.log(`Weekly email sent to ${CONFIG.EMAIL.OPS_LEAD}. ${rows.length} pending request(s).`);

  // ── Update status only AFTER successful send ─────────────
  _markRequestsAsSent();
}


/**
 * After the weekly email is sent to Yevheniia,
 * updates all 'requested' rows → 'sent' in the main sheet.
 * 'sent' means: Yevheniia has been notified, waiting for account creation.
 */
function _markRequestsAsSent() {
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const statusCol = getCol(colMap, CONFIG.HEADERS.EMAIL_STATUS);

  let count = 0;

  for (let row = CONFIG.SHEET.DATA_START_ROW; row <= sheet.getLastRow(); row++) {
    const cell   = sheet.getRange(row, statusCol);
    const status = safeString(cell.getValue()).toLowerCase();

    if (status === CONFIG.STATUS.REQUEST) {
      cell.setValue(CONFIG.STATUS.SENT); // 'sent'
      count++;
    }
  }

  Logger.log(`_markRequestsAsSent: ${count} row(s) updated requested → sent.`);
}


/**
 * Returns a styled HTML badge for the given request status.
 *
 * @param {string} status
 * @returns {string} HTML string
 */
function _statusBadge(status) {
  const c = THEME.status[status];
  const style = c
    ? `background:${c.bg}; color:${c.fg};`
    : `background:${THEME.surfaceMuted}; color:${THEME.textStrong};`;
  return `<span style="padding:2px 8px; border-radius:4px; font-size:12px; font-weight:600; ${style}">${status}</span>`;
}


/**
 * Renders the shared HTML body for email-creation-request emails — the exact
 * table/layout used by the Monday batch. Both the weekly summary and the
 * single-member "send now" action call this, so the two emails look identical.
 * Emails can't use CSS variables, so THEME tokens are inlined.
 *
 * @param {string}  greetingName    name after "Hi " (recipient of the request)
 * @param {string}  introHtml       HTML paragraph(s) shown above the table
 * @param {Array[]} rows            [name, role, personalEmail, status, witEmail][]
 * @param {string}  afterTableHtml  HTML shown below the table (summary + sign-off)
 * @returns {string} full HTML body
 */
function _renderEmailRequestsEmailHtml(greetingName, introHtml, rows, afterTableHtml) {
  const T = THEME;

  const tableRows = rows
    .map(([name, role, personalEmail, requestStatus, witEmail]) => `
      <tr>
        <td style="padding:8px 10px; border-bottom:1px solid ${T.border};">${name}</td>
        <td style="padding:8px 10px; border-bottom:1px solid ${T.border}; color:${T.textMuted};">${role}</td>
        <td style="padding:8px 10px; border-bottom:1px solid ${T.border};">${personalEmail}</td>
        <td style="padding:8px 10px; border-bottom:1px solid ${T.border};">${_statusBadge(requestStatus)}</td>
        <td style="padding:8px 10px; border-bottom:1px solid ${T.border}; color:${T.primary}; font-weight:bold;">${witEmail}</td>
      </tr>
    `)
    .join('');

  return `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: ${T.textStrong}; width: 100%;">

      <p>Hi ${greetingName},</p>
      ${introHtml}

      <table style="border-collapse: collapse; width: 100%; margin: 20px 0;">
        <thead>
          <tr style="background:${T.primary}; color:${T.onPrimary};">
            <th style="padding:10px; text-align:left;">Full Name</th>
            <th style="padding:10px; text-align:left;">Role</th>
            <th style="padding:10px; text-align:left;">Personal Email</th>
            <th style="padding:10px; text-align:left;">Status</th>
            <th style="padding:10px; text-align:left;">WIT Email (suggested)</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>

      ${afterTableHtml}

    </div>
  `;
}


/**
 * Sends a single-member WIT email-creation request in the shared batch layout.
 * Centralizes the date stamp and the report-row shape so callers (the on-edit
 * 'requested' notification and the ad-hoc menu action) supply only the
 * recipient, greeting, subject, and the wording around the table.
 *
 * @param {Object}   o
 * @param {string}   o.to             recipient email
 * @param {string}   o.greetingName   name after "Hi "
 * @param {string}   o.subject        email subject
 * @param {Object}   o.member         { fullName, role, personalEmail, status, witEmail }
 * @param {(date: string) => string} o.intro  builds the intro <p>, given the date
 * @param {string}   o.afterTableHtml HTML below the table
 */
function _sendMemberRequestEmail({ to, greetingName, subject, member, intro, afterTableHtml }) {
  const dateStr = new Date().toLocaleDateString('en-CA', {
    timeZone:  'America/Vancouver',
    dateStyle: 'full',
  });

  const row = [
    member.fullName,
    member.role          || '—',
    member.personalEmail || '—',
    member.status,
    member.witEmail,
  ];

  MailApp.sendEmail({
    to,
    subject,
    htmlBody: _renderEmailRequestsEmailHtml(greetingName, intro(dateStr), [row], afterTableHtml),
  });
}


/**
 * Menu action (👤 Member Onboarding → Send Email Request — this member).
 * Sends a WIT email-creation request for the single highlighted member using
 * the same layout as the Monday batch, then flips the row's Email Status →
 * 'sent'. Use to request one account ad-hoc without waiting for the weekly run.
 *
 * Recipient/greeting mirror the Monday batch (goes to CONFIG.EMAIL.OPS_LEAD).
 * Setting the status via setValue does not fire onEdit (script edits don't
 * trigger it), so there is no interaction with the edit handlers.
 */
function sendEmailRequestForSelectedMember() {
  const ctx = requireActiveMemberRow();
  if (!ctx) return;
  const { sheet, row, colMap } = ctx;

  const get = (header) => safeString(sheet.getRange(row, getCol(colMap, header)).getValue());

  const firstName     = get(CONFIG.HEADERS.FIRST_NAME);
  const lastName      = get(CONFIG.HEADERS.LAST_NAME);
  const role          = get(CONFIG.HEADERS.ROLE);
  const personalEmail = get(CONFIG.HEADERS.PERSONAL_EMAIL);

  // ── Validate required fields ─────────────────────────────
  if (!firstName || !lastName || !personalEmail) {
    showAlert('⚠️ Missing Data', buildMissingFieldsMessage({ firstName, lastName, personalEmail }));
    return;
  }

  const witEmail = generateWITEmail(firstName, lastName);
  const fullName = `${firstName} ${lastName}`;

  // ── Confirmation ─────────────────────────────────────────
  if (CONFIG.UI.CONFIRMATION) {
    const ui       = SpreadsheetApp.getUi();
    const response = ui.alert(
      '📧 Send Email Request',
      `Send a WIT email creation request for ${fullName}?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // ── Send (single-row table, Monday layout) ───────────────
  _sendMemberRequestEmail({
    to:           CONFIG.EMAIL.OPS_LEAD,
    greetingName: 'Tamuna',
    subject:      `WIT Email Request — ${fullName}`,
    member:       { fullName, role, personalEmail, status: CONFIG.STATUS.REQUEST, witEmail },
    intro: (d) =>
      `<p>Here is a WIT email creation request for the following member, as of <strong>${d}</strong></p>`,
    afterTableHtml:
      `<p>Please create this account at your convenience. Let me know if you have any questions.</p>`,
  });

  // ── Mark row requested → sent (mirrors the Monday flow) ──
  sheet.getRange(row, getCol(colMap, CONFIG.HEADERS.EMAIL_STATUS)).setValue(CONFIG.STATUS.SENT);

  Logger.log(`sendEmailRequestForSelectedMember: request sent for ${fullName}; status → sent.`);

  if (CONFIG.UI.ALERTS) {
    showAlert('✅ Sent', `Email request for ${fullName} sent.`);
  }
}