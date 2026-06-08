// ============================================================
// WIT Canada — Trigger Registration (single source of truth)
// ============================================================
// Container-bound Apps Script has no version control over its
// installable triggers — they are normally wired by hand in the
// editor (Triggers panel), which means a project copy / re-deploy
// silently loses them. This file makes the wiring explicit and
// reproducible.
//
// Run setupTriggers() ONCE from the Apps Script editor after any
// fresh deploy (or after changing the wiring here). It is
// idempotent: it removes the triggers it owns before recreating
// them, so re-running never creates duplicates.
//
// NOT covered here:
//   • onOpen()  — simple trigger, fires automatically, no install.
//
// Time-based hours are interpreted in the script time zone
// (America/Vancouver, see appsscript.json), so atHour(3) == 3 AM PT.
// ============================================================

/**
 * Handlers this file owns. Used to dedupe before (re)creating,
 * so setupTriggers() is safe to run repeatedly.
 */
const MANAGED_TRIGGER_HANDLERS = [
  // On-edit (installable) — currently one trigger per handler.
  // NOTE: every edit fires ALL of these. Phase 3 of the refactor
  // collapses them behind a single onEdit dispatcher; until then
  // this list documents the real, current wiring.
  'processEmailRequestOnEdit',
  'processGroupsRequestOnEdit',
  'processMemberStatusOnEdit',
  'processIntroSentOnEdit',

  // Form submit (installable)
  'handleProfileFormSubmit',

  // Time-driven
  'syncPhotosAndBios',             // daily photo/bio match, ~6 AM PT
  'buildEmailRequestsReport',      // daily report rebuild, ~3 AM PT
  'sendWeeklyEmailRequestsReport', // Monday summary, ~7 AM PT
];

/**
 * Registers every installable trigger the project depends on.
 * Idempotent — run from the editor after a fresh deploy.
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Remove existing copies of the triggers we own ──────────
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (MANAGED_TRIGGER_HANDLERS.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  // ── On-edit (installable) ──────────────────────────────────
  // One trigger per handler — mirrors the current manual wiring.
  ['processEmailRequestOnEdit',
   'processGroupsRequestOnEdit',
   'processMemberStatusOnEdit',
   'processIntroSentOnEdit']
    .forEach((fn) => {
      ScriptApp.newTrigger(fn).forSpreadsheet(ss).onEdit().create();
    });

  // ── Form submit (installable) ──────────────────────────────
  ScriptApp.newTrigger('handleProfileFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // ── Time-driven: daily photo/bio match (~6 AM PT) ──────────
  ScriptApp.newTrigger('syncPhotosAndBios')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  // ── Time-driven: daily report rebuild (~3 AM PT) ───────────
  ScriptApp.newTrigger('buildEmailRequestsReport')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();

  // ── Time-driven: weekly summary (Monday ~7 AM PT) ──────────
  ScriptApp.newTrigger('sendWeeklyEmailRequestsReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();

  Logger.log(
    `setupTriggers: removed ${removed} existing, registered ` +
    `${MANAGED_TRIGGER_HANDLERS.length} trigger(s).`,
  );
}

/**
 * Diagnostic: logs all installable triggers currently registered
 * on the project. Run from the editor to audit live wiring against
 * MANAGED_TRIGGER_HANDLERS.
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    Logger.log('listTriggers: no installable triggers registered.');
    return;
  }

  triggers.forEach((t) => {
    Logger.log(
      `${t.getHandlerFunction()}  ·  ${t.getEventType()}  ·  ${t.getTriggerSource()}`,
    );
  });
}
