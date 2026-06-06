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

  const tableRows = rows
    .map(([name, role, personalEmail, requestStatus, witEmail]) => `
      <tr>
        <td style="padding:8px 10px; border-bottom:1px solid #e8eaed;">${name}</td>
        <td style="padding:8px 10px; border-bottom:1px solid #e8eaed; color:#5f6368;">${role}</td>
        <td style="padding:8px 10px; border-bottom:1px solid #e8eaed;">${personalEmail}</td>
        <td style="padding:8px 10px; border-bottom:1px solid #e8eaed;">${_statusBadge(requestStatus)}</td>
        <td style="padding:8px 10px; border-bottom:1px solid #e8eaed; color:#1a2fa3; font-weight:bold;">${witEmail}</td>
      </tr>
    `)
    .join('');

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #1a1a1a; max-width: 800px;">

      <p>Hi Tamuna,</p>
      <p>Here is the weekly summary of WIT email creation requests as of <strong>${timestamp}</strong>:</p>

      <table style="border-collapse: collapse; width: 100%; margin: 20px 0;">
        <thead>
          <tr style="background:#1a2fa3; color:#ffffff;">
            <th style="padding:10px; text-align:left;">Full Name</th>
            <th style="padding:10px; text-align:left;">Role</th>
            <th style="padding:10px; text-align:left;">Personal Email</th>
            <th style="padding:10px; text-align:left;">Status</th>
            <th style="padding:10px; text-align:left;">WIT Email (suggested)</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>

      <p style="color:#5f6368; font-size:13px;">
        New this week: <strong>${newCount}</strong> &nbsp;|&nbsp; Not activated: <strong>${notActivatedCount}</strong> &nbsp;|&nbsp; Total tracked: <strong>${rows.length}</strong>
      </p>

      <p>Please process at your convenience. Let me know if you have any questions.</p>

      <p>Best regards,<br>
      <strong>WIT Canada Operations</strong> (automated weekly report)</p>

    </div>
  `;

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
  const statusCol = getCol(colMap, CONFIG.HEADERS.REQUEST_STATUS);

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
  const styles = {
    'requested':     'background:#fff3cd; color:#856404;',
    'sent':          'background:#cfe2ff; color:#0a4a90;',
    'created':       'background:#d1e7dd; color:#0a5933;',
    'not activated': 'background:#fde8e8; color:#9c1c1c;',
  };
  const style = styles[status] ?? 'background:#f1f3f4; color:#3c4043;';
  return `<span style="padding:2px 8px; border-radius:4px; font-size:12px; font-weight:600; ${style}">${status}</span>`;
}