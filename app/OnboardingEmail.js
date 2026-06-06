// ============================================================
// WIT Canada — Onboarding Welcome Email
// ============================================================
// Trigger: From spreadsheet → On edit (installable trigger)
//
// Fires when: 'Intro sent' checkbox is ticked on WIT_Members
// Guard:      Email Status must be 'created'
// Sends to:   Personal Email
// ============================================================

function processIntroSentOnEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET.MAIN) return;

  // ── Guard: only react to 'Intro sent' column ─────────────
  const colMap = getColumnIndexMap(sheet);
  const introCol = colMap[normalizeHeader(CONFIG.HEADERS.INTRO_SENT)];
  if (!introCol)
    throw new Error(`Column not found: ${CONFIG.HEADERS.INTRO_SENT}`);
  if (e.range.getColumn() !== introCol) return;

  // ── Guard: only react when box transitions false → true ──
  // e.oldValue === 'TRUE' means it was already checked — skip to prevent re-sends.
  if (e.oldValue === "TRUE") return;
  const checked = e.range.getValue();
  if (checked !== true) return;

  const row = e.range.getRow();
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Resolve required columns ─────────────────────────────
  const firstNameCol    = colMap[normalizeHeader(CONFIG.HEADERS.FIRST_NAME)];
  const lastNameCol     = colMap[normalizeHeader(CONFIG.HEADERS.LAST_NAME)];
  const personalEmailCol = colMap[normalizeHeader(CONFIG.HEADERS.PERSONAL_EMAIL)];
  const emailStatusCol  = colMap[normalizeHeader(CONFIG.HEADERS.REQUEST_STATUS)];
  const roleCol         = colMap[normalizeHeader(CONFIG.HEADERS.ROLE)];
  const witEmailCol     = colMap[normalizeHeader(CONFIG.HEADERS.WIT_EMAIL)];

  // ── Read row data ─────────────────────────────────────────
  const firstName = safeString(sheet.getRange(row, firstNameCol).getValue());
  const lastName  = safeString(sheet.getRange(row, lastNameCol).getValue());
  const personalEmail = safeString(
    sheet.getRange(row, personalEmailCol).getValue(),
  );
  const emailStatus = safeString(
    sheet.getRange(row, emailStatusCol).getValue(),
  ).toLowerCase();
  const role     = safeString(sheet.getRange(row, roleCol).getValue());
  const witEmail = safeString(sheet.getRange(row, witEmailCol).getValue());
  const fullName = `${firstName} ${lastName}`.trim();

  // ── Guard: email must be created or active ───────────────
  const emailReady = emailStatus === CONFIG.STATUS.CREATED || emailStatus === CONFIG.STATUS.ACTIVE;
  if (!emailReady) {
    sheet.getRange(row, introCol).setValue(false); // revert checkbox
    showAlert(
      "⚠️ Cannot send — email not ready",
      `${fullName}'s Email Status is "${emailStatus}".\n\nSet Email Status to "created" or "active" before marking Intro sent.`,
    );
    return;
  }

  // ── Guard: personal email required ───────────────────────
  if (!personalEmail) {
    sheet.getRange(row, introCol).setValue(false);
    showAlert(
      "⚠️ Cannot send — no personal email",
      `No personal email found for ${fullName}.\n\nAdd a personal email address and try again.`,
    );
    return;
  }

  // ── Confirmation ──────────────────────────────────────────
  if (CONFIG.UI.CONFIRMATION) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "📧 Send onboarding email",
      `Send the welcome email to:\n\n` +
        `Name:   ${fullName}\n` +
        `Email:  ${personalEmail}`,
      ui.ButtonSet.YES_NO,
    );

    if (response !== ui.Button.YES) {
      sheet.getRange(row, introCol).setValue(false);
      return;
    }
  }

  // ── Send ──────────────────────────────────────────────────
  try {
    MailApp.sendEmail({
      to: personalEmail,
      subject: "Welcome to Women in Tech Canada — Your Onboarding",
      htmlBody: _buildOnboardingEmailHtml(firstName, lastName, role, witEmail),
      name: `Women in Tech Canada — ${CONFIG.EMAIL.OPS_LEAD_NAME}`,
    });

    Logger.log(
      `processIntroSentOnEdit: onboarding email sent to ${personalEmail} (${fullName}).`,
    );

    showAlert(
      "✅ Email sent",
      `Onboarding email sent to ${fullName} at ${personalEmail}.`,
    );
  } catch (err) {
    sheet.getRange(row, introCol).setValue(false); // revert checkbox on failure
    Logger.log(`processIntroSentOnEdit ERROR: ${err.message}`);

    showAlert(
      "❌ Failed to send email",
      `Something went wrong while sending to ${personalEmail}.\n\nError: ${err.message}`,
    );
  }
}

// ─────────────────────────────────────────────────────────────
// Email template
// ─────────────────────────────────────────────────────────────

function _buildOnboardingEmailHtml(firstName, lastName, role, witEmail) {
  const params = '&firstName=' + encodeURIComponent(firstName || '')
    + '&lastName='  + encodeURIComponent(lastName  || '')
    + '&role='      + encodeURIComponent(role      || '')
    + '&witEmail='  + encodeURIComponent(witEmail  || '');
  const checklistUrl  = CONFIG.URLS.MEMBER_GUIDE + '?' + params.slice(1);
  const signatureUrl  = CONFIG.URLS.SIGNATURE_GENERATOR + params;

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.07);">

          <!-- Header -->
          <tr>
            <td style="background:#1b4f8a;padding:28px 36px;">
              <p style="margin:0;font-size:11px;color:rgba(255,255,255,0.6);text-transform:uppercase;letter-spacing:0.08em;">Women in Tech Canada</p>
              <h1 style="margin:6px 0 0;font-size:20px;color:#ffffff;font-weight:bold;">Welcome to the team 👋</h1>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:32px 36px;">

              <p style="margin:0 0 20px;font-size:15px;color:#2c3e50;line-height:1.6;">
                Hi ${firstName},
              </p>
              <p style="margin:0 0 20px;font-size:14px;color:#5d6d7e;line-height:1.6;">
                Welcome to Women in Tech Canada — we're glad to have you with us.
              </p>
              <p style="margin:0 0 28px;font-size:14px;color:#5d6d7e;line-height:1.6;">
                To get started, please complete your onboarding using the interactive checklist below:
              </p>

              <!-- CTA buttons -->
              <table cellpadding="0" cellspacing="0" style="margin-bottom:16px;">
                <tr>
                  <td style="background:#1b4f8a;border-radius:6px;">
                    <a href="${checklistUrl}" target="_blank"
                       style="display:inline-block;padding:12px 28px;color:#ffffff;font-size:14px;font-weight:bold;text-decoration:none;">
                      👉 Open Onboarding Checklist
                    </a>
                  </td>
                </tr>
              </table>
              <table cellpadding="0" cellspacing="0" style="margin-bottom:32px;">
                <tr>
                  <td style="background:#1b4f8a;border-radius:6px;">
                    <a href="${signatureUrl}" target="_blank"
                       style="display:inline-block;padding:12px 28px;color:#ffffff;font-size:14px;font-weight:bold;text-decoration:none;">
                      ✍️ Create your email signature →
                    </a>
                  </td>
                </tr>
              </table>

              <!-- Divider -->
              <hr style="border:none;border-top:1px solid #e8ecf0;margin:0 0 28px;">

              <!-- Steps -->
              <h2 style="margin:0 0 20px;font-size:15px;color:#1b4f8a;font-weight:bold;">Key Steps</h2>

              <!-- Step 1 -->
              <table cellpadding="0" cellspacing="0" style="margin-bottom:20px;width:100%;">
                <tr>
                  <td width="32" valign="top">
                    <div style="width:24px;height:24px;background:#e8f0fb;border-radius:50%;text-align:center;line-height:24px;font-size:12px;font-weight:bold;color:#1b4f8a;">1</div>
                  </td>
                  <td valign="top">
                    <p style="margin:0 0 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Email Activation</p>
                    <p style="margin:0;font-size:13px;color:#7f8c8d;line-height:1.5;">
                      You should have received an email with instructions to activate your WIT account.<br>
                      Please activate it within <strong>48 hours</strong>.
                    </p>
                  </td>
                </tr>
              </table>

              <!-- Step 2 -->
              <table cellpadding="0" cellspacing="0" style="margin-bottom:20px;width:100%;">
                <tr>
                  <td width="32" valign="top">
                    <div style="width:24px;height:24px;background:#e8f0fb;border-radius:50%;text-align:center;line-height:24px;font-size:12px;font-weight:bold;color:#1b4f8a;">2</div>
                  </td>
                  <td valign="top">
                    <p style="margin:0 0 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Access to Tools</p>
                    <p style="margin:0 0 6px;font-size:13px;color:#7f8c8d;line-height:1.5;">Once your email is activated, you will receive access to:</p>
                    <ul style="margin:0;padding-left:18px;font-size:13px;color:#7f8c8d;line-height:1.8;">
                      <li>Google Drive (WIT Canada shared drive)</li>
                      <li>Internal communication tools</li>
                      <li>Relevant working documents</li>
                    </ul>
                  </td>
                </tr>
              </table>

              <!-- Step 3 -->
              <table cellpadding="0" cellspacing="0" style="margin-bottom:20px;width:100%;">
                <tr>
                  <td width="32" valign="top">
                    <div style="width:24px;height:24px;background:#e8f0fb;border-radius:50%;text-align:center;line-height:24px;font-size:12px;font-weight:bold;color:#1b4f8a;">3</div>
                  </td>
                  <td valign="top">
                    <p style="margin:0 0 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Communication Channels</p>
                    <p style="margin:0;font-size:13px;color:#7f8c8d;line-height:1.5;">
                      You will be added to the appropriate groups based on your role.
                    </p>
                  </td>
                </tr>
              </table>

              <!-- Step 4 -->
              <table cellpadding="0" cellspacing="0" style="margin-bottom:28px;width:100%;">
                <tr>
                  <td width="32" valign="top">
                    <div style="width:24px;height:24px;background:#e8f0fb;border-radius:50%;text-align:center;line-height:24px;font-size:12px;font-weight:bold;color:#1b4f8a;">4</div>
                  </td>
                  <td valign="top">
                    <p style="margin:0 0 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Review Materials</p>
                    <p style="margin:0;font-size:13px;color:#7f8c8d;line-height:1.5;">
                      Please follow the onboarding checklist to complete all required steps and review key documents.
                    </p>
                  </td>
                </tr>
              </table>

              <!-- Divider -->
              <hr style="border:none;border-top:1px solid #e8ecf0;margin:0 0 24px;">

              <p style="margin:0;font-size:13px;color:#7f8c8d;line-height:1.6;">
                If you have any questions or run into any issues, feel free to reach out to
                <a href="mailto:${CONFIG.EMAIL.OPS_LEAD}" style="color:#1b4f8a;text-decoration:none;">${CONFIG.EMAIL.OPS_LEAD}</a>.
              </p>

            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#f8f9fa;padding:16px 36px;border-top:1px solid #e8ecf0;">
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
