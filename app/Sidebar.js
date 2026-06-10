// ============================================================
// WIT Canada — Sidebar
// ============================================================

function showFileGuide() {
  showSidebar('ui/FileGuide', 'About this file', 320);
}

function showNewMemberGuideSidebar() {
  showSidebar('ui/OnboardingGuide', 'Member onboarding guide', 280);
}

function showSidebar(templateName, title, width) {
  const html = HtmlService
    .createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .setWidth(width);

  SpreadsheetApp.getUi().showSidebar(html);
}