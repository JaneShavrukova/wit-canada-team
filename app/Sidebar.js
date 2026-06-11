// ============================================================
// WIT Canada — Sidebar
// ============================================================

function showFileGuide() {
  showSidebar('FileGuide', 'About this file', 320);
}

/**
 * Installable onOpen handler: auto-opens the System Guide sidebar when the
 * spreadsheet is opened, so new users discover the workflow without hunting
 * through the menu.
 *
 * Must be an INSTALLABLE trigger (wired in app/Triggers.js → setupTriggers),
 * not the simple onOpen(): Ui.showSidebar() requires the script.container.ui
 * scope, which simple triggers (AuthMode.LIMITED) do not have — calling it
 * from the simple onOpen throws "Specified permissions are not sufficient".
 * As an installable trigger it runs with the installing user's full
 * authorization, so it works for every editor without each re-authorizing.
 */
function onOpenInstallable() {
  try {
    showFileGuide();
  } catch (err) {
    Logger.log('onOpenInstallable: could not open guide sidebar — ' + err);
  }
}

function showNewMemberGuideSidebar() {
  showSidebar('OnboardingGuide', 'Member onboarding guide', 280);
}

function showSidebar(templateName, title, width) {
  const html = HtmlService
    .createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .setWidth(width);

  SpreadsheetApp.getUi().showSidebar(html);
}