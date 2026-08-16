// ============================================================
// WIT Canada — Activation Follow-up (Email Status → 'not activated')
// ============================================================
// Driven by the Email Status column. On a transition INTO
// 'not activated' (the account exists but the member has not signed in):
//   1. branded Yes/No confirm modal,
//   2. YES → email the admin asking them to chase the activation,
//      NO  → revert Email Status to its previous value, send nothing.
//
// Only fires on a real transition into the value, so a re-save of the
// same cell does not re-prompt. A row with no WIT email is a no-op with
// a toast: without an address the admin has nothing to act on.
//
// Interaction with the other Email Status handlers: processEmailRequest
// reacts only to 'requested' and processOnboardingEmail only to
// 'requested'/'created', so none of them collide with this one.
//
// processMemberStatusOnEdit also runs on this edit and may set Member
// Status → 'onboarding'. That is fine even when the lead cancels: it
// treats 'not activated' and 'created' identically (neither is in its
// EMAIL_EXCLUDED set), so the reverted row lands on the same Member
// Status it would have had anyway.
//
// UX note: showModalDialog does NOT block server execution (unlike
// ui.alert), so the modal drives the outcome via google.script.run —
// _confirmNotActivatedNotice (send) or _cancelNotActivatedNotice
// (revert). The same _confirmNotActivatedNotice is called directly when
// confirmations are off. Reverting via setValue does NOT re-fire onEdit.
// ============================================================

function processNotActivatedOnEdit(e, ctx) {
  const { sheet, sheetName, row, col, colMap } = ctx || buildEditContext(e);
  if (sheetName !== CONFIG.SHEET.MAIN) return;

  // ── Guard: only the Email Status column drives this ──
  const statusCol = colMap[normalizeHeader(CONFIG.HEADERS.EMAIL_STATUS)];
  if (!statusCol) throw new Error(`Column not found: ${CONFIG.HEADERS.EMAIL_STATUS}`);
  if (col !== statusCol) return;
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Guard: only act on a real transition into 'not activated' ──
  const oldStatus = safeString(e.oldValue).toLowerCase();
  const newStatus = safeString(e.range.getValue()).toLowerCase();
  if (oldStatus === newStatus) return;
  if (newStatus !== CONFIG.STATUS.NOT_ACTIVATED) return;

  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const role      = safeString(get(CONFIG.HEADERS.ROLE));
  const witEmail  = safeString(get(CONFIG.HEADERS.WIT_EMAIL));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';

  // ── No WIT email → nothing for the admin to chase ──
  // The status itself is left alone; it may well be accurate.
  if (!witEmail) {
    Logger.log(`processNotActivatedOnEdit: no WIT email for row ${row} — nothing to follow up, skipped.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${fullName} has no WIT email on file, so no follow-up was sent to the admin.`,
      'ℹ️ Nothing to follow up',
      6,
    );
    return;
  }

  // Value to restore if the lead declines (defaults to 'created' — the
  // status that normally precedes it — when the edit carried no oldValue).
  const restoreTo = safeString(e.oldValue) || CONFIG.STATUS.CREATED;

  if (CONFIG.UI.CONFIRMATION) {
    _showNotActivatedConfirm({ row, fullName, role, witEmail, restoreTo });
  } else {
    _confirmNotActivatedNotice(row); // confirmations off → act immediately
  }
}


// ─────────────────────────────────────────────────────────────
// Branded Yes/No modal. Its buttons call back into
// _confirmNotActivatedNotice / _cancelNotActivatedNotice.
// ─────────────────────────────────────────────────────────────

function _showNotActivatedConfirm({ row, fullName, role, witEmail, restoreTo }) {
  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
${themeCss()}
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; }
  body {
    font-family: 'Google Sans', Roboto, Arial, sans-serif;
    background:
      radial-gradient(120% 90% at 100% 0%, rgba(232,184,75,0.14) 0%, rgba(232,184,75,0) 55%),
      var(--color-surface-muted);
    padding: 18px;
    color: var(--color-text-body);
  }
  .card {
    position: relative;
    background: var(--color-surface);
    border-radius: 16px;
    border: 1px solid var(--color-border);
    box-shadow: 0 10px 30px rgba(20,20,40,0.16), 0 2px 6px rgba(20,20,40,0.08);
    overflow: hidden;
    max-width: 460px;
    margin: 0 auto;
  }
  .banner {
    background: linear-gradient(135deg, #f0cd72 0%, #e8b84b 45%, #d3a02f 100%);
    padding: 20px 22px;
    display: flex;
    align-items: center;
    gap: 14px;
  }
  .badge {
    flex-shrink: 0;
    width: 44px; height: 44px;
    border-radius: 50%;
    background: rgba(255,255,255,0.28);
    box-shadow: inset 0 0 0 1px rgba(255,255,255,0.45);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
  }
  .banner h1 { color: var(--color-text-strong); font-size: 17px; font-weight: 600; line-height: 1.25; }
  .banner .sub { color: rgba(44,62,80,0.72); font-size: 12px; margin-top: 2px; }
  .body { padding: 20px 22px 22px; }
  .lead { font-size: 13.5px; line-height: 1.6; margin-bottom: 16px; }
  .lead strong { color: var(--color-text-strong); }
  .acct {
    background: var(--color-warning-tint);
    border: 1px solid rgba(232,184,75,0.45);
    border-left: 4px solid var(--color-warning);
    border-radius: 10px;
    padding: 13px 15px;
    margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(232,184,75,0.14);
  }
  .acct .label { font-size: 10.5px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: #9a7415; margin-bottom: 5px; }
  .acct .who { font-size: 14.5px; font-weight: 700; color: var(--color-text-strong); }
  .acct .role { font-size: 12px; color: var(--color-text-muted); margin: 1px 0 7px; }
  .acct .mail { font-size: 13px; font-weight: 700; color: #9a7415; word-break: break-all; display: flex; align-items: center; gap: 6px; }
  .note {
    font-size: 12px; line-height: 1.6; color: var(--color-text-muted);
    background: var(--color-surface-muted);
    border-radius: 8px; padding: 11px 13px; margin-bottom: 20px;
  }
  .note strong { color: var(--color-text-strong); }
  .buttons { display: flex; gap: 10px; }
  .btn {
    flex: 1; padding: 11px 16px; border-radius: 9px;
    font-size: 13px; font-weight: 600; cursor: pointer;
    border: none; font-family: inherit;
    transition: transform 0.08s ease, box-shadow 0.15s ease, background 0.15s ease, opacity 0.15s ease;
  }
  .btn:active { transform: translateY(1px); }
  .btn[disabled] { opacity: 0.6; cursor: default; transform: none; }
  .btn-primary {
    color: var(--color-on-primary);
    background: linear-gradient(135deg, #2560a0 0%, var(--color-primary) 55%, var(--color-primary-dark) 100%);
    box-shadow: 0 4px 12px rgba(27,79,138,0.30);
  }
  .btn-primary:hover:not([disabled]) { box-shadow: 0 6px 16px rgba(27,79,138,0.40); }
  .btn-secondary { background: var(--color-primary-tint); color: var(--color-primary); }
  .btn-secondary:hover:not([disabled]) { background: var(--color-primary-tint-hover); }
</style>
</head>
<body>
<div class="card">
  <div class="banner">
    <div class="badge">⏳</div>
    <div>
      <h1>Email not activated</h1>
      <div class="sub">Ask the admin to follow up?</div>
    </div>
  </div>
  <div class="body">
    <p class="lead">You're marking <strong>${fullName}</strong> as <strong>not activated</strong>. We can email the admin to chase the activation of the account below.</p>
    <div class="acct">
      <div class="label">Account awaiting activation</div>
      <div class="who">${fullName}</div>
      <div class="role">${role || '—'}</div>
      <div class="mail">✉️ ${witEmail}</div>
    </div>
    <p class="note">Choosing <strong>Yes</strong> emails the admin and keeps the status at <strong>not activated</strong>. Choosing <strong>No</strong> sends nothing and reverts the status to <strong>${restoreTo}</strong>.</p>
    <div class="buttons">
      <button id="ok" class="btn btn-primary" onclick="proceed()">Yes, email the admin</button>
      <button id="no" class="btn btn-secondary" onclick="cancel()">No, undo</button>
    </div>
  </div>
</div>
<script>
  var ROW = ${Number(row)};
  var RESTORE = ${JSON.stringify(restoreTo)};
  function busy(label){ var ok=document.getElementById('ok'), no=document.getElementById('no'); ok.disabled=true; no.disabled=true; ok.textContent=label; }
  function done(){ google.script.host.close(); }
  function proceed(){ busy('Sending…'); google.script.run.withSuccessHandler(done).withFailureHandler(done)._confirmNotActivatedNotice(ROW); }
  function cancel(){ busy('Reverting…'); google.script.run.withSuccessHandler(done).withFailureHandler(done)._cancelNotActivatedNotice(ROW, RESTORE); }
</script>
</body>
</html>`).setWidth(500).setHeight(430);

  SpreadsheetApp.getUi().showModalDialog(html, '⏳ Email not activated');
}


// ─────────────────────────────────────────────────────────────
// Server callbacks (invoked by the modal via google.script.run).
// Both re-read the row server-side — never trust client-passed data.
// ─────────────────────────────────────────────────────────────

/**
 * Emails the admin asking them to chase the account activation. Also the
 * direct path when confirmations are off. Feedback is a non-blocking toast.
 *
 * Unlike offboarding, a send failure does NOT revert the status: 'not
 * activated' describes the account's real state, which is true whether or
 * not the reminder went out. The lead is told so they can retry.
 *
 * @param {number} row
 */
function _confirmNotActivatedNotice(row) {
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const role      = safeString(get(CONFIG.HEADERS.ROLE));
  const witEmail  = safeString(get(CONFIG.HEADERS.WIT_EMAIL));
  const personal  = safeString(get(CONFIG.HEADERS.PERSONAL_EMAIL));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';
  const ss        = SpreadsheetApp.getActiveSpreadsheet();

  if (!witEmail) {
    Logger.log(`_confirmNotActivatedNotice: no WIT email for row ${row} — skipped.`);
    return;
  }

  try {
    MailApp.sendEmail({
      to:       CONFIG.EMAIL.OPS_LEAD,
      subject:  `[WIT Email Activation] ${fullName}`,
      htmlBody: _buildActivationReminderHtml(fullName, role, witEmail, personal),
      name:     `Women in Tech Canada — ${CONFIG.EMAIL.OPS_LEAD_NAME}`,
    });
    Logger.log(`_confirmNotActivatedNotice: activation follow-up sent to ${CONFIG.EMAIL.OPS_LEAD} for ${fullName} (${witEmail}).`);
  } catch (err) {
    Logger.log(`_confirmNotActivatedNotice ERROR: ${err.message}`);
    ss.toast(
      `Could not email the admin about ${fullName}. ${err.message} The status is unchanged — try again from the Email Status cell.`,
      '❌ Follow-up not sent',
      8,
    );
    return;
  }

  ss.toast(
    `Activation follow-up sent to the admin for ${fullName} (${witEmail}).`,
    '✅ Admin notified',
    7,
  );
}

/**
 * Reverts Email Status when the lead answers "No" in the confirm modal.
 * @param {number} row
 * @param {string} restoreTo  value to restore (its pre-edit Email Status)
 */
function _cancelNotActivatedNotice(row, restoreTo) {
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const statusCol = getCol(colMap, CONFIG.HEADERS.EMAIL_STATUS);
  sheet.getRange(row, statusCol).setValue(restoreTo || CONFIG.STATUS.CREATED);
  Logger.log(`_cancelNotActivatedNotice: row ${row} reverted to "${restoreTo}".`);
}


// ─────────────────────────────────────────────────────────────
// Email template — reuses the branded _emailShell (OnboardingEmail.js).
// Amber accent marks a follow-up (not a destructive action).
// ─────────────────────────────────────────────────────────────

function _buildActivationReminderHtml(fullName, role, witEmail, personalEmail) {
  const T   = THEME;
  const ops = CONFIG.EMAIL.OPS_LEAD;

  const contactRow = personalEmail
    ? `<p style="margin:6px 0 0;font-size:13px;color:${T.textMuted};">Personal email: ${personalEmail}</p>`
    : '';

  const inner = `
              <p style="margin:0 0 16px;font-size:15px;color:${T.textStrong};line-height:1.6;">Hi Tamuna,</p>
              <p style="margin:0 0 20px;font-size:14px;color:${T.textBody};line-height:1.6;">
                The WIT account below has been created but is still
                <strong>not activated</strong> — the member has not signed in yet.
                Could you check the account and re-send the activation invite if needed?
              </p>

              <!-- Member / account card -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background:${T.warningTint};border-radius:6px;border-left:3px solid ${T.warning};margin-bottom:24px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <p style="margin:0 0 6px;font-size:12px;font-weight:bold;color:${T.textStrong};text-transform:uppercase;letter-spacing:0.05em;">Awaiting activation</p>
                    <p style="margin:0 0 2px;font-size:15px;font-weight:bold;color:${T.textStrong};">${fullName}</p>
                    <p style="margin:0 0 8px;font-size:13px;color:${T.textMuted};">${role || '—'}</p>
                    <p style="margin:0;font-size:14px;color:${T.textStrong};font-weight:bold;">${witEmail}</p>
                    ${contactRow}
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 24px;font-size:14px;color:${T.textBody};line-height:1.6;">
                Once the member signs in, the Ops Lead will move their
                <strong>Email Status</strong> to <strong>active</strong> — no action needed
                from you at that point.
              </p>

              <hr style="border:none;border-top:1px solid ${T.border};margin:0 0 20px;">
              <p style="margin:0;font-size:13px;color:${T.textMuted};line-height:1.6;">
                Questions? Reach out to <a href="mailto:${ops}" style="color:${T.primary};text-decoration:none;">${ops}</a>.
              </p>`;

  return _emailShell('Follow-up — WIT email not activated', inner);
}
