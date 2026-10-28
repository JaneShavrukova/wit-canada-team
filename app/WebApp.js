// ============================================================
// WIT Canada — Web App Entry Point
// ============================================================
// Handles all doGet requests for the deployed web app.
// Route via ?page= parameter:
//   (default / no param) → OnboardingGuide member checklist
//   page=signature        → Email Signature Generator
// ============================================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'guide';

  if (page === 'signature') {
    const tpl = HtmlService.createTemplateFromFile('SignatureGenerator');
    tpl.firstName = (e.parameter.firstName || '').toString();
    tpl.lastName  = (e.parameter.lastName  || '').toString();
    tpl.role      = (e.parameter.role      || '').toString();
    tpl.witEmail  = (e.parameter.witEmail  || '').toString();
    tpl.phone     = (e.parameter.phone     || '').toString();
    return tpl.evaluate()
      .setTitle('Email Signature Generator — WIT Canada')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default: interactive member welcome checklist
  const fn = (e.parameter.firstName || '').toString();
  const ln = (e.parameter.lastName  || '').toString();
  const rl = (e.parameter.role      || '').toString();
  const we = (e.parameter.witEmail  || '').toString();
  const ph = (e.parameter.phone     || '').toString();
  const guideTpl = HtmlService.createTemplateFromFile('MemberGuide');
  guideTpl.firstName    = fn;
  guideTpl.lastName     = ln;
  guideTpl.role         = rl;
  guideTpl.witEmail     = we;
  // Carried through so the guide's "set up my signature" link hands the
  // phone number on to the generator (it pre-fills the optional row).
  guideTpl.signatureUrl = buildUrl(CONFIG.URLS.SIGNATURE_GENERATOR, {
    firstName: fn,
    lastName:  ln,
    role:      rl,
    witEmail:  we,
    phone:     ph,
  });
  return guideTpl.evaluate()
    .setTitle('WIT Canada — Welcome Guide')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
