// ============================================================
// WIT Canada — Offboarding (Member Status → 'inactive')
// ============================================================
// Mirror image of onboarding. Driven by the Member Status column:
//   • transition → 'inactive' (on a member who HAS a WIT email) :
//       1. branded confirm modal, 2. email the admin to DELETE the WIT
//       account, 3. flag Email Status → 'deleted'.
//
// 'inactive' is also the resting/default state, so this fires ONLY on
// a real transition into it AND only when a WIT email exists — a brand-
// new member sitting at 'inactive' with no account is a silent no-op.
//
// Member Status is otherwise script-written by processMemberStatusOnEdit,
// which watches Contract/Email/Groups (never Member Status itself), so a
// manual edit here does not collide with auto-promotion.
//
// UX note: showModalDialog does NOT block server execution (unlike
// ui.alert), so the confirm modal drives the outcome via google.script.run
// — _confirmOffboarding (send + flag) or _cancelOffboarding (revert). The
// same _confirmOffboarding is called directly when confirmations are off.
// Setting Email Status via setValue does NOT re-fire onEdit.
// ============================================================

function processOffboardingOnEdit(e, ctx) {
  const { sheet, sheetName, row, col, colMap } = ctx || buildEditContext(e);
  if (sheetName !== CONFIG.SHEET.MAIN) return;

  // ── Guard: only the Member Status column drives offboarding ──
  const memberCol = colMap[normalizeHeader(CONFIG.HEADERS.MEMBER_STATUS)];
  if (!memberCol) throw new Error(`Column not found: ${CONFIG.HEADERS.MEMBER_STATUS}`);
  if (col !== memberCol) return;
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Guard: only act on a real transition into 'inactive' ──
  const oldStatus = safeString(e.oldValue).toLowerCase();
  const newStatus = safeString(e.range.getValue()).toLowerCase();
  if (oldStatus === newStatus) return;
  if (newStatus !== CONFIG.MEMBER_STATUS.INACTIVE) return;

  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const role      = safeString(get(CONFIG.HEADERS.ROLE));
  const witEmail  = safeString(get(CONFIG.HEADERS.WIT_EMAIL));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';

  // ── No WIT email → nothing to offboard (e.g. a new default row) ──
  // Silent: 'inactive' is the resting state, so this isn't an error.
  if (!witEmail) {
    Logger.log(`processOffboardingOnEdit: no WIT email for row ${row} — no account to delete, skipped.`);
    return;
  }

  // Old value to restore if the lead cancels (defaults to 'active' — the
  // usual pre-offboarding state — when the edit carried no oldValue).
  const restoreTo = safeString(e.oldValue) || CONFIG.MEMBER_STATUS.ACTIVE;

  if (CONFIG.UI.CONFIRMATION) {
    _showOffboardConfirm({ row, fullName, role, witEmail, restoreTo });
  } else {
    _confirmOffboarding(row); // confirmations off → act immediately
  }
}


// ─────────────────────────────────────────────────────────────
// Branded confirmation modal (replaces the plain ui.alert).
// Its buttons call back into _confirmOffboarding / _cancelOffboarding.
// ─────────────────────────────────────────────────────────────

function _showOffboardConfirm({ row, fullName, role, witEmail, restoreTo }) {
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
      radial-gradient(120% 90% at 100% 0%, rgba(232,84,59,0.10) 0%, rgba(232,84,59,0) 55%),
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
  /* Gradient danger banner with icon badge */
  .banner {
    background: linear-gradient(135deg, #f0674d 0%, #e8543b 45%, #cf3f28 100%);
    padding: 20px 22px;
    display: flex;
    align-items: center;
    gap: 14px;
  }
  .badge {
    flex-shrink: 0;
    width: 44px; height: 44px;
    border-radius: 50%;
    background: rgba(255,255,255,0.18);
    box-shadow: inset 0 0 0 1px rgba(255,255,255,0.35);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
  }
  .banner h1 { color: #fff; font-size: 17px; font-weight: 600; line-height: 1.25; }
  .banner .sub { color: rgba(255,255,255,0.82); font-size: 12px; margin-top: 2px; }
  .body { padding: 20px 22px 22px; }
  .lead { font-size: 13.5px; line-height: 1.6; margin-bottom: 16px; }
  .lead strong { color: var(--color-text-strong); }
  /* Account chip with left accent + soft depth */
  .acct {
    background: var(--brand-coral-tint);
    border: 1px solid rgba(232,84,59,0.28);
    border-left: 4px solid var(--brand-coral);
    border-radius: 10px;
    padding: 13px 15px;
    margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(232,84,59,0.10);
  }
  .acct .label { font-size: 10.5px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: var(--brand-coral); margin-bottom: 5px; }
  .acct .who { font-size: 14.5px; font-weight: 700; color: var(--color-text-strong); }
  .acct .role { font-size: 12px; color: var(--color-text-muted); margin: 1px 0 7px; }
  .acct .mail { font-size: 13px; font-weight: 700; color: var(--brand-coral); word-break: break-all; display: flex; align-items: center; gap: 6px; }
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
  .btn-danger {
    color: #fff;
    background: linear-gradient(135deg, #ef6047 0%, #e8543b 55%, #d5432b 100%);
    box-shadow: 0 4px 12px rgba(232,84,59,0.34);
  }
  .btn-danger:hover:not([disabled]) { box-shadow: 0 6px 16px rgba(232,84,59,0.42); }
  .btn-secondary { background: var(--color-primary-tint); color: var(--color-primary); }
  .btn-secondary:hover:not([disabled]) { background: var(--color-primary-tint-hover); }
</style>
</head>
<body>
<div class="card">
  <div class="banner">
    <div class="badge">🗑️</div>
    <div>
      <h1>Offboard member</h1>
      <div class="sub">This deletes their WIT account</div>
    </div>
  </div>
  <div class="body">
    <p class="lead">You're marking <strong>${fullName}</strong> as <strong>inactive</strong>. The admin will be asked to delete the account below.</p>
    <div class="acct">
      <div class="label">Account to delete</div>
      <div class="who">${fullName}</div>
      <div class="role">${role || '—'}</div>
      <div class="mail">✉️ ${witEmail}</div>
    </div>
    <p class="note">On confirm, the deletion request is emailed to the admin and <strong>Email Status</strong> is set to <strong>deleted</strong>. Cancel reverts the status and sends nothing.</p>
    <div class="buttons">
      <button id="ok" class="btn btn-danger" onclick="proceed()">Send deletion request</button>
      <button id="no" class="btn btn-secondary" onclick="cancel()">Cancel</button>
    </div>
  </div>
</div>
<script>
  var ROW = ${Number(row)};
  var RESTORE = ${JSON.stringify(restoreTo)};
  function busy(label){ var ok=document.getElementById('ok'), no=document.getElementById('no'); ok.disabled=true; no.disabled=true; ok.textContent=label; }
  function done(){ google.script.host.close(); }
  function proceed(){ busy('Sending…'); google.script.run.withSuccessHandler(done).withFailureHandler(done)._confirmOffboarding(ROW); }
  function cancel(){ busy('Cancelling…'); google.script.run.withSuccessHandler(done).withFailureHandler(done)._cancelOffboarding(ROW, RESTORE); }
</script>
</body>
</html>`).setWidth(500).setHeight(430);

  SpreadsheetApp.getUi().showModalDialog(html, '🗑️ Offboard member');
}


// ─────────────────────────────────────────────────────────────
// Server callbacks (invoked by the modal via google.script.run).
// Both re-read the row server-side — never trust client-passed data.
// ─────────────────────────────────────────────────────────────

/**
 * Sends the deletion request to the admin and flags Email Status →
 * 'deleted'. Also the direct path when confirmations are off. Feedback is
 * a non-blocking toast. On send failure, reverts Member Status → 'active'.
 * @param {number} row
 */
function _confirmOffboarding(row) {
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const role      = safeString(get(CONFIG.HEADERS.ROLE));
  const witEmail  = safeString(get(CONFIG.HEADERS.WIT_EMAIL));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';
  const ss        = SpreadsheetApp.getActiveSpreadsheet();

  if (!witEmail) {
    Logger.log(`_confirmOffboarding: no WIT email for row ${row} — skipped.`);
    return;
  }

  try {
    MailApp.sendEmail({
      to:       CONFIG.EMAIL.OPS_LEAD,
      subject:  `[WIT Email Deletion] ${fullName}`,
      htmlBody: _buildDeleteRequestHtml(fullName, role, witEmail),
      name:     `Women in Tech Canada — ${CONFIG.EMAIL.OPS_LEAD_NAME}`,
    });
    Logger.log(`_confirmOffboarding: deletion request sent to ${CONFIG.EMAIL.OPS_LEAD} for ${fullName} (${witEmail}).`);
  } catch (err) {
    Logger.log(`_confirmOffboarding ERROR: ${err.message}`);
    ss.toast(`Could not send the deletion request for ${fullName}. ${err.message}`, '❌ Offboarding failed', 8);
    const memberCol = colMap[normalizeHeader(CONFIG.HEADERS.MEMBER_STATUS)];
    if (memberCol) sheet.getRange(row, memberCol).setValue(CONFIG.MEMBER_STATUS.ACTIVE);
    return;
  }

  // ── Flag Email Status → 'deleted' (script edit; no onEdit) ──
  const emailStatusCol = colMap[normalizeHeader(CONFIG.HEADERS.EMAIL_STATUS)];
  if (emailStatusCol) sheet.getRange(row, emailStatusCol).setValue(CONFIG.STATUS.DELETED);

  ss.toast(
    `Deletion request sent to the admin for ${fullName} (${witEmail}). Email Status set to "deleted".`,
    '✅ Offboarding requested',
    7,
  );
}

/**
 * Reverts Member Status when the lead cancels the confirm modal.
 * @param {number} row
 * @param {string} restoreTo  value to restore (its pre-edit Member Status)
 */
function _cancelOffboarding(row, restoreTo) {
  const sheet     = getMainSheet();
  const colMap    = getColumnIndexMap(sheet);
  const memberCol = getCol(colMap, CONFIG.HEADERS.MEMBER_STATUS);
  sheet.getRange(row, memberCol).setValue(restoreTo || CONFIG.MEMBER_STATUS.ACTIVE);
  Logger.log(`_cancelOffboarding: row ${row} reverted to "${restoreTo}".`);
}


// ─────────────────────────────────────────────────────────────
// Email template — reuses the branded _emailShell (OnboardingEmail.js).
// Coral accent (brand) marks the destructive action.
// ─────────────────────────────────────────────────────────────

function _buildDeleteRequestHtml(fullName, role, witEmail) {
  const T   = THEME;
  const ops = CONFIG.EMAIL.OPS_LEAD;

  const inner = `
              <p style="margin:0 0 16px;font-size:15px;color:${T.textStrong};line-height:1.6;">Hi Tamuna,</p>
              <p style="margin:0 0 20px;font-size:14px;color:${T.textBody};line-height:1.6;">
                The following member is being offboarded from Women in Tech Canada.
                Please delete their WIT account.
              </p>

              <!-- Member / account card -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background:${T.brand.coralTint};border-radius:6px;border-left:3px solid ${T.brand.coral};margin-bottom:24px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <p style="margin:0 0 6px;font-size:12px;font-weight:bold;color:${T.brand.coral};text-transform:uppercase;letter-spacing:0.05em;">Account to delete</p>
                    <p style="margin:0 0 2px;font-size:15px;font-weight:bold;color:${T.textStrong};">${fullName}</p>
                    <p style="margin:0 0 8px;font-size:13px;color:${T.textMuted};">${role || '—'}</p>
                    <p style="margin:0;font-size:14px;color:${T.brand.coral};font-weight:bold;">${witEmail}</p>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 24px;font-size:14px;color:${T.textBody};line-height:1.6;">
                Once the account has been removed, no further action is needed — the sheet already
                reflects this member as <strong>inactive</strong>.
              </p>

              <hr style="border:none;border-top:1px solid ${T.border};margin:0 0 20px;">
              <p style="margin:0;font-size:13px;color:${T.textMuted};line-height:1.6;">
                Questions? Reach out to <a href="mailto:${ops}" style="color:${T.primary};text-decoration:none;">${ops}</a>.
              </p>`;

  return _emailShell('Offboarding — delete WIT account', inner);
}
