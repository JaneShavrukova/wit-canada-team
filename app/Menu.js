// ============================================================
// WIT Canada — Main Menu
// This is the ONLY onOpen() in the project.
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('WIT Operations')

    // ── 👤 Member Onboarding ─────────────────────────────────
    .addSubMenu(ui.createMenu('👤 Member Onboarding')
      .addItem('📋 Onboarding Checklist',        'openMemberGuideForRow')
      .addItem('✍️ Email Signature Generator',   'openSignatureGeneratorForRow')
      .addItem('📇 Personal Info',               'openPersonalInfoForm'))

    .addSeparator()

    // ── 📘 Documentation ────────────────────────────────────
    .addSubMenu(ui.createMenu('📘 Documentation')
      .addItem('System Guide',                   'showFileGuide')
      .addItem('Onboarding SOP — Leads',         'showNewMemberGuideSidebar'))

    .addSeparator()

    // ── ⚙️ Actions ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ Actions')
      .addItem('Refresh Email Requests',         'buildEmailRequestsReport')
      .addItem('Refresh Photos & Bios',          'syncPhotosAndBios')
      .addItem('Refresh Groups',                 'buildGroupsView'))

    .addSeparator()

    // ── 🔍 Reports ───────────────────────────────────────────
    .addSubMenu(ui.createMenu('🔍 Reports')
      .addItem('Open Email Requests',            'openEmailRequestsReport')
      .addItem('Open Missing Bios',              'openPhotoBioReport')
      .addItem('Open Groups Status',             'openGroupsReport'))

    .addToUi();
}


// ============================================================
// Report navigation helpers
// ============================================================

function openEmailRequestsReport() {
  _openReportSheet(CONFIG.SHEET.REPORTS.EMAIL_REQUESTS);
}

function openPhotoBioReport() {
  _openReportSheet(CONFIG.SHEET.REPORTS.PHOTO_BIO);
}

function openGroupsReport() {
  _openReportSheet(CONFIG.SHEET.REPORTS.GROUPS);
}

function _openReportSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Report not found',
      `"${name}" hasn't been generated yet.\n\nRun the corresponding Refresh action first.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  ss.setActiveSheet(sheet);
}


// ============================================================
// Member Onboarding
// ============================================================

function openMemberGuideForRow() {
  const ctx = requireActiveMemberRow();
  if (!ctx) return;
  const { sheet, row, colMap } = ctx;

  const get       = (header) => safeString(sheet.getRange(row, colMap[normalizeHeader(header)]).getValue());

  const firstName = get(CONFIG.HEADERS.FIRST_NAME);
  const role      = get(CONFIG.HEADERS.ROLE);

  const url = buildUrl(CONFIG.URLS.MEMBER_GUIDE, { firstName, role });

  const displayName = [firstName].filter(Boolean).join(' ') || '';

  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
${themeCss()}
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Google Sans', Arial, sans-serif; background: var(--color-surface-muted); display: flex; align-items: flex-start; justify-content: center; padding: 16px; }
  .card { background: var(--color-surface); border-radius: 12px; box-shadow: 0 2px 12px rgba(27,79,138,0.10); padding: 28px 28px 24px; max-width: 420px; width: 100%; }
  .header { background: var(--color-primary); border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; display: flex; align-items: center; gap: 12px; }
  .header .emoji { font-size: 24px; line-height: 1; }
  .header h1 { color: var(--color-on-primary); font-size: 16px; font-weight: 600; line-height: 1.3; }
  .description { color: var(--color-text-body); font-size: 13px; line-height: 1.6; margin-bottom: 16px; }
  .covers-label { color: var(--color-primary); font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px; }
  ul { list-style: none; margin-bottom: 24px; }
  ul li { color: var(--color-text-body); font-size: 13px; padding: 4px 0 4px 20px; position: relative; line-height: 1.5; }
  ul li::before { content: '✓'; position: absolute; left: 0; color: var(--color-primary); font-weight: 700; font-size: 12px; }
  .buttons { display: flex; gap: 10px; }
  .btn { flex: 1; padding: 10px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; border: none; transition: background 0.15s, color 0.15s; font-family: inherit; }
  .btn-primary { background: var(--color-primary); color: var(--color-on-primary); text-decoration: none; display: flex; align-items: center; justify-content: center; gap: 6px; }
  .btn-primary:hover { background: var(--color-primary-dark); }
  .btn-secondary { background: var(--color-primary-tint); color: var(--color-primary); }
  .btn-secondary:hover { background: var(--color-primary-tint-hover); }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="emoji">📋</span>
    <h1>Onboarding Checklist</h1>
  </div>
  <p class="description">A personalised step-by-step checklist to complete during onboarding.</p>
  <p class="covers-label">What it covers</p>
  <ul>
    <li>Account &amp; email setup</li>
    <li>Communication &amp; access</li>
    <li>Community profile</li>
    <li>Say hello to the team</li>
  </ul>
  <div class="buttons">
    <a class="btn btn-primary" href="${url}" target="_blank" onclick="google.script.host.close()">Open Checklist ↗</a>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
  </div>
</div>
</body>
</html>`).setWidth(480).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '📋 Onboarding Checklist');
}

function openPersonalInfoForm() {
  const url = 'https://docs.google.com/forms/d/e/1FAIpQLSfw4HePeT3moc9RuFSx3wAgMfAFod7a4J5VVuHQM7aIePcuCA/viewform';
  const html = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '","_blank");google.script.host.close();</script>')
    .setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Personal Info Form…');
}


// ============================================================
// Signature Generator
// ============================================================

function openSignatureGeneratorForRow() {
  const ctx = requireActiveMemberRow();
  if (!ctx) return;
  const { sheet, row, colMap } = ctx;

  const get     = (header) => safeString(sheet.getRange(row, colMap[normalizeHeader(header)]).getValue());

  const firstName = get(CONFIG.HEADERS.FIRST_NAME);
  const lastName  = get(CONFIG.HEADERS.LAST_NAME);
  const role      = get(CONFIG.HEADERS.ROLE);
  const witEmail  = get(CONFIG.HEADERS.WIT_EMAIL);

  const url = buildUrl(CONFIG.URLS.SIGNATURE_GENERATOR, { firstName, lastName, role, witEmail });

  const html = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '","_blank");google.script.host.close();</script>')
    .setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Signature Generator…');
}
