# WIT Canada — Operations Apps Script

Container-bound Google Apps Script that drives member onboarding ops from a
Google Sheet. Edits to the sheet fire trigger handlers (send emails, sync
Google Groups, promote member status); menu actions build report sheets and
open onboarding tools; a deployed web app serves the member guide and the
email-signature generator.

Pushed to Apps Script with `clasp push` (see `.clasp.json`). Files under
`.claspignore` stay in the repo but are not deployed.

## Layout

```
core/    Config.js   Single source of truth: sheet/column names, status enums,
                     group↔email map, Drive folder IDs, URLs.
         Utils.js    Stateless helpers: sheet access, column-index mapping,
                     string normalization, Drive lookups, UI alerts, email gen.

app/     Triggers.js            setupTriggers() — registers all installable triggers.
         Menu.js                onOpen() menu + onboarding-tool launchers.
         WebApp.js              doGet() — routes the web app by ?page=.
         EmailRequest.js        onEdit: Email Status → "requested".
         EmailRequestsReport.js Builds the Email_Requests report sheet (daily).
         EmailRequestWeeklySend.js  Monday batch summary email.
         GroupsAccessRequest.js onEdit: "Add to groups" → "requested" (notify).
         GroupsSync.js          Actually adds members to Google Groups (AdminDirectory).
         GroupsView.js          Builds the Report_Groups snapshot sheet.
         MemberStatusSync.js    onEdit: auto-promote Member Status → "active".
         OnboardingEmail.js     onEdit: "Intro sent" → send welcome email.
         PhotoBioSync.js        Matches Drive photos/bios to members.
         PhotoBioReport.js      Renders the photo/bio status sheet + sidebar.
         ProfileUpdate.js       Form-submit handler + its trigger installer.
         Sidebar.js             Doc sidebars and the member-guide launcher.

tools/   DriveStructure.js, SpreadSheetStructure.js — one-off export utilities.

*.html   FileGuide / MemberGuide / OnboardingGuide / SignatureGenerator —
         served via HtmlService. SiteTeamVisualPrototype.html is repo-only
         (excluded from push).
```

## How triggers are wired (read before changing handlers)

There is **no master `onEdit` dispatcher yet**. Each of these is registered as
its own *installable* On-Edit trigger, so **every cell edit fires all four**,
and each independently rebuilds the column map and reads cells:

| Handler | Sheet | Reacts to |
|---|---|---|
| `processEmailRequestOnEdit` | WIT_Members | Email Status → `requested` |
| `processGroupsRequestOnEdit` | WIT_Members, WIT_External | Add to groups → `requested` |
| `processMemberStatusOnEdit` | WIT_Members | Contract+Email+Groups all done → `active` |
| `processIntroSentOnEdit` | WIT_Members | Intro sent checkbox → true |

Other installable triggers:

| Handler | Type | Schedule |
|---|---|---|
| `handleProfileFormSubmit` | Form submit | on submit |
| `syncPhotosAndBios` | Time-driven | daily 6–7 AM PT |
| `buildEmailRequestsReport` | Time-driven | daily 3–4 AM PT |
| `sendWeeklyEmailRequestsReport` | Time-driven | Monday 7–8 AM PT |

`onOpen` is a simple trigger (no install needed).

**Wiring lives in `app/Triggers.js → setupTriggers()`.** It is idempotent (dedupes
by handler name). Run it once from the editor after any fresh deploy. Audit live
wiring with `listTriggers()`. (`ProfileUpdate.createProfileFormTrigger()` is the
legacy single-purpose installer, now subsumed by `setupTriggers`.)

> **Not reproducible in code:** the per-trigger *Failure notification* setting
> (the time-based jobs use "Notify me immediately" / "Notify me daily"). The
> `ScriptApp` trigger builder can't set this — re-apply it manually in the
> Triggers panel after running `setupTriggers()`. Time-based triggers also fire
> at a random minute within the configured hour, so exact run times drift.

## Conventions & constraints

- **`CONFIG` is the single source of truth.** Add new sheet names, headers,
  statuses, colors, folder IDs, and URLs there — not as scattered top-level
  consts. (A few legacy ones still live in app files; consolidating them is a
  planned refactor step.)
- **File evaluation order.** Apps Script has no modules; all files share one
  global scope and top-level statements run in load order. Do **not** reference
  `CONFIG` from another file's *top-level* code — only from inside function
  bodies (evaluated at call time). Some handlers build their `Set`s of watched
  sheets inside the function for exactly this reason.
- **Column access is header-driven**, never by fixed index. Resolve columns via
  `getColumnIndexMap(sheet)` + `normalizeHeader()` / `getCol()`, so reordering
  sheet columns doesn't break handlers. Header row = `CONFIG.SHEET.HEADER_ROW`,
  data starts at `CONFIG.SHEET.DATA_START_ROW`.
- **Minimize Spreadsheet API calls.** Each `getRange().getValue()` is a round
  trip. Prefer one batch `getValues()` / `setValues()` per range (see
  `PhotoBioSync` for the pattern).
- Helper functions intended as private are prefixed `_`; everything is global
  (GAS has no real namespaces), so keep names unambiguous.

## Local dev

```
clasp push            # deploy to Apps Script
clasp open            # open the project in the editor
```

There is no local test runner; handlers are verified by editing the live sheet
or running functions from the editor. After deploying trigger changes, run
`setupTriggers()` from the editor.
