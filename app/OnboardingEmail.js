// ============================================================
// WIT Canada — Onboarding Emails (Letter 1 & Letter 2)
// ============================================================
// Driven by the Email Status column — the single source of truth:
//   • transition → 'requested' : Letter 1 to the Personal Email
//   • transition → 'created'   : Letter 2 to the WIT Email (confirmed)
//
// Emails fire ONLY on a transition INTO the target value. Re-applying
// the same value, or editing other columns, sends nothing. There is
// no resend confirmation and no history tracking (by design).
//
// 'requested': processEmailRequestOnEdit already shows the confirmation
// and reverts the status on cancel, and it runs BEFORE this handler in
// the dispatcher — so Letter 1 is automatically gated by it (a cancel
// reverts the value, leaving "no transition" here).
//
// 'created': this handler shows its own confirmation (mirroring the
// 'requested' UX). Cancel reverts the status and sends nothing.
// ============================================================

function processOnboardingEmailOnEdit(e, ctx) {
  const { sheet, sheetName, row, col, colMap } = ctx || buildEditContext(e);
  if (sheetName !== CONFIG.SHEET.MAIN) return;

  // ── Guard: only the Email Status column drives onboarding emails ──
  const statusCol = colMap[normalizeHeader(CONFIG.HEADERS.EMAIL_STATUS)];
  if (!statusCol) throw new Error(`Column not found: ${CONFIG.HEADERS.EMAIL_STATUS}`);
  if (col !== statusCol) return;
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Guard: only act on a real transition into a target value ──
  const oldStatus = safeString(e.oldValue).toLowerCase();
  const newStatus = safeString(e.range.getValue()).toLowerCase();
  if (oldStatus === newStatus) return;

  if (newStatus === CONFIG.STATUS.REQUEST) {
    _sendLetter1(sheet, row, colMap);
  } else if (newStatus === CONFIG.STATUS.CREATED) {
    _sendLetter2WithConfirm(e, sheet, row, colMap, statusCol);
  }
}


// ─────────────────────────────────────────────────────────────
// Letter 1 — Email Status → 'requested' → Personal Email
// ─────────────────────────────────────────────────────────────

function _sendLetter1(sheet, row, colMap) {
  const get           = getRowValues(sheet, row, colMap);
  const firstName     = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const personalEmail = safeString(get(CONFIG.HEADERS.PERSONAL_EMAIL));

  // processEmailRequestOnEdit validates/reverts missing fields first, so
  // this is a defensive skip rather than a user-facing error.
  if (!personalEmail) {
    Logger.log(`_sendLetter1: no personal email for row ${row} — skipped.`);
    return;
  }

  try {
    MailApp.sendEmail({
      to:       personalEmail,
      subject:  'Welcome to Women in Tech Canada — Your Onboarding',
      htmlBody: _buildLetter1Html(firstName),
      name:     `Women in Tech Canada — ${CONFIG.EMAIL.OPS_LEAD_NAME}`,
    });
    Logger.log(`_sendLetter1: sent to ${personalEmail} (row ${row}).`);
  } catch (err) {
    Logger.log(`_sendLetter1 ERROR: ${err.message}`);
    showAlert('❌ Onboarding email failed', `Could not send Letter 1 to ${personalEmail}.\n\n${err.message}`);
  }
}


// ─────────────────────────────────────────────────────────────
// Letter 2 — Email Status → 'created' → WIT Email (with confirmation)
// ─────────────────────────────────────────────────────────────

function _sendLetter2WithConfirm(e, sheet, row, colMap, statusCol) {
  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const role      = safeString(get(CONFIG.HEADERS.ROLE));
  const witEmail  = safeString(get(CONFIG.HEADERS.WIT_EMAIL));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';

  const revert = () => sheet.getRange(row, statusCol).setValue(e.oldValue ?? CONFIG.STATUS.NEW);

  // ── Confirmation (mirrors the 'requested' UX) ────────────
  if (CONFIG.UI.CONFIRMATION) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '📧 Send onboarding email #2',
      `Email Status set to "created" for ${fullName}.\n\n` +
      `Send the onboarding checklist email to:\n${witEmail || '(no WIT email)'}\n\n` +
      `Continue? (Cancel reverts the status and sends nothing.)`,
      ui.ButtonSet.YES_NO,
    );

    if (response !== ui.Button.YES) {
      revert();
      return;
    }
  }

  // ── Guard: WIT email required to send ────────────────────
  if (!witEmail) {
    showAlert(
      '⚠️ Cannot send — no WIT email',
      `No WIT email found for ${fullName}.\n\nAdd it, then set Email Status to "created" again.`,
    );
    revert();
    return;
  }

  // ── Send ─────────────────────────────────────────────────
  try {
    MailApp.sendEmail({
      to:       witEmail,
      subject:  'Welcome to Women in Tech Canada — Complete Your Onboarding',
      htmlBody: _buildLetter2Html(firstName, lastName, role, witEmail),
      name:     `Women in Tech Canada — ${CONFIG.EMAIL.OPS_LEAD_NAME}`,
    });
    Logger.log(`_sendLetter2: sent to ${witEmail} (row ${row}).`);
    showAlert('✅ Email sent', `Onboarding checklist email sent to ${fullName} at ${witEmail}.`);
  } catch (err) {
    Logger.log(`_sendLetter2 ERROR: ${err.message}`);
    showAlert('❌ Onboarding email failed', `Could not send Letter 2 to ${witEmail}.\n\n${err.message}`);
  }
}


// ─────────────────────────────────────────────────────────────
// Email templates
// ─────────────────────────────────────────────────────────────

/**
 * Standard branded shell (header banner + body + footer) shared by both
 * letters. Emails can't use CSS variables, so THEME tokens are inlined.
 */
function _emailShell(heading, inner) {
  const T = THEME;
  return `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:${T.surfaceAlt};font-family:Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background:${T.surfaceAlt};padding:32px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:${T.surface};border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.07);">

          <!-- Header -->
          <tr>
            <td style="background:${T.primary};padding:28px 36px;">
              <p style="margin:0;font-size:11px;color:rgba(255,255,255,0.6);text-transform:uppercase;letter-spacing:0.08em;">Women in Tech Canada</p>
              <h1 style="margin:6px 0 0;font-size:20px;color:${T.onPrimary};font-weight:bold;">${heading}</h1>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:32px 36px;">${inner}
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:${T.surfaceMuted};padding:16px 36px;border-top:1px solid ${T.border};">
              <p style="margin:0;font-size:11px;color:${T.textFaint};text-align:center;">
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

/**
 * Letter 1 — sent to the Personal Email when Email Status becomes
 * 'requested'. Tells the member their WIT account is being set up and
 * gives two prefilled-mailto CTAs (activated / need help).
 */
function _buildLetter1Html(firstName) {
  const T       = THEME;
  const ops     = CONFIG.EMAIL.OPS_LEAD;
  const opsName = CONFIG.EMAIL.OPS_LEAD_NAME;
  const name    = firstName || 'there';

  const activatedMailto = buildUrl(`mailto:${ops}`, {
    subject: 'WIT Email Activated',
    body:    `Hi ${opsName},\n\nI have successfully activated my WIT email and am ready for the next onboarding steps.\n\nThanks,\n${name}`,
  });
  const helpMailto = buildUrl(`mailto:${ops}`, {
    subject: 'WIT Email Activation Issue',
    body:    `Hi ${opsName},\n\nI'm experiencing an issue activating my WIT email account.\n\nIssue details:\n[please describe the problem]\n\nThanks,\n${name}`,
  });

  const inner = `
              <p style="margin:0 0 16px;font-size:15px;color:${T.textStrong};line-height:1.6;">Hi ${name},</p>
              <p style="margin:0 0 16px;font-size:14px;color:${T.textBody};line-height:1.6;">Welcome to Women in Tech Canada — we're glad to have you with us!</p>
              <p style="margin:0 0 20px;font-size:14px;color:${T.textBody};line-height:1.6;">Your WIT Canada email is currently being set up by our admin.</p>

              <!-- Action required -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background:${T.brand.coralTint};border-radius:6px;border-left:3px solid ${T.brand.coral};margin-bottom:24px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <p style="margin:0 0 6px;font-size:12px;font-weight:bold;color:${T.brand.coral};text-transform:uppercase;letter-spacing:0.05em;">Action required</p>
                    <p style="margin:0;font-size:13px;color:${T.textBody};line-height:1.6;">
                      Watch for an activation email from <strong>Google Workspace</strong> in this inbox (check your spam folder too).
                      You'll need to activate your account within <strong>48 hours</strong> of receiving it.
                    </p>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 24px;font-size:14px;color:${T.textBody};line-height:1.6;">Once activated, you'll receive your onboarding checklist and next steps.</p>

              <!-- CTA buttons -->
              <table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto 10px;">
                <tr>
                  <td style="background:${T.primary};border-radius:6px;">
                    <a href="${activatedMailto}" style="display:inline-block;padding:12px 24px;color:${T.onPrimary};font-size:14px;font-weight:bold;text-decoration:none;">Confirm Activation</a>
                  </td>
                </tr>
              </table>
              <table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto 28px;">
                <tr>
                  <td style="background:${T.primaryTint};border-radius:6px;">
                    <a href="${helpMailto}" style="display:inline-block;padding:12px 24px;color:${T.primary};font-size:14px;font-weight:bold;text-decoration:none;">Get Activation Help</a>
                  </td>
                </tr>
              </table>

              <hr style="border:none;border-top:1px solid ${T.border};margin:0 0 20px;">
              <p style="margin:0;font-size:13px;color:${T.textMuted};line-height:1.6;">
                Questions? Reach out to <a href="mailto:${ops}" style="color:${T.primary};text-decoration:none;">${ops}</a>.
              </p>`;

  return _emailShell('Welcome to the team 👋', inner);
}

/**
 * Letter 2 — sent to the WIT Email when Email Status becomes 'created'.
 * Account is ready; drives the member to the onboarding checklist and
 * explains why it matters.
 */
function _buildLetter2Html(firstName, lastName, role, witEmail) {
  const T    = THEME;
  const ops  = CONFIG.EMAIL.OPS_LEAD;
  const name = firstName || 'there';

  const checklistUrl = buildUrl(CONFIG.URLS.MEMBER_GUIDE, {
    firstName: firstName || '',
    lastName:  lastName  || '',
    role:      role      || '',
    witEmail:  witEmail  || '',
  });

  const bullets = [
    'Access the communication channels relevant to your role',
    'Join the correct distribution lists, groups, and team spaces',
    'Complete your profile and introduction',
    'Ensure you receive updates, opportunities, and team communications',
  ].map((b) => `<li style="margin-bottom:6px;">${b}</li>`).join('');

  const inner = `
              <p style="margin:0 0 16px;font-size:15px;color:${T.textStrong};line-height:1.6;">Hi ${name},</p>
              <p style="margin:0 0 16px;font-size:14px;color:${T.textBody};line-height:1.6;">Your Women in Tech Canada account is ready. 🎉</p>
              <p style="margin:0 0 24px;font-size:14px;color:${T.textBody};line-height:1.6;">To get connected and fully onboarded, please complete your onboarding checklist:</p>

              <!-- CTA -->
              <table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto 28px;">
                <tr>
                  <td style="background:${T.primary};border-radius:6px;">
                    <a href="${checklistUrl}" target="_blank" style="display:inline-block;padding:12px 28px;color:${T.onPrimary};font-size:14px;font-weight:bold;text-decoration:none;">👉 Complete your Onboarding Checklist</a>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 10px;font-size:14px;color:${T.textBody};line-height:1.6;">The checklist will help you:</p>
              <ul style="margin:0 0 24px;padding-left:18px;font-size:14px;color:${T.textBody};line-height:1.7;">${bullets}</ul>

              <p style="margin:0 0 24px;font-size:14px;color:${T.textMuted};line-height:1.6;">Most members complete the checklist in just a few minutes.</p>

              <hr style="border:none;border-top:1px solid ${T.border};margin:0 0 20px;">
              <p style="margin:0;font-size:13px;color:${T.textMuted};line-height:1.6;">
                Questions? Reach out to <a href="mailto:${ops}" style="color:${T.primary};text-decoration:none;">${ops}</a>.
              </p>`;

  return _emailShell('Your account is ready 🎉', inner);
}
