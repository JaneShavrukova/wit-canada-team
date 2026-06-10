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
core/    Config.js   Data & identifiers: sheet/column names, status enums,
                     group↔email map, Drive folder IDs, URLs. NOT colors.
         Theme.js    Design tokens (THEME) — the one palette for sheets,
                     sidebars, web pages, and emails; themeCss() emits CSS vars.
         ReportBuilder.js  Shared report-sheet scaffolding (banner, timestamp,
                     get-or-create, column widths).
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
         MemberStatusSync.js    onEdit: Member Status → "onboarding" / "active".
         ContractUpload.js      onEdit: Contract → "signed" → prompt to upload.
         OnboardingEmail.js     onEdit: Email Status → Letter 1 (requested) / Letter 2 (created).
         PhotoBioSync.js        Matches Drive photos/bios to members.
         PhotoBioReport.js      Renders the photo/bio status sheet + sidebar.
         ProfileUpdate.js       Form-submit handler + its trigger installer.
         Sidebar.js             Doc sidebars and the member-guide launcher.

tools/   DriveStructure.js, SpreadsheetStructure.js — one-off export utilities.

*.html   FileGuide / MemberGuide / OnboardingGuide / SignatureGenerator —
         served via HtmlService. SiteTeamVisualPrototype.html is repo-only
         (excluded from push).
```

## How triggers are wired (read before changing handlers)

A single *installable* On-Edit trigger, **`onEditInstallable`** (app/EditRouter.js),
is the only edit trigger. It builds the column map **once** per edit and dispatches
to the four handlers, each of which no-ops unless the edit matches its watched
sheet + column/value. Handlers take `(e, ctx)` where `ctx = { sheet, sheetName,
row, col, colMap }`; called without `ctx` they rebuild it via `buildEditContext(e)`.
A handler that throws is isolated (logged) so the others still run; the first
error is re-thrown afterwards so failures still surface.

| Handler (dispatched, not trigger-bound) | Sheet | Reacts to |
|---|---|---|
| `processEmailRequestOnEdit` | WIT_Members | Email Status → `requested` (Ops Lead notification) |
| `processGroupsRequestOnEdit` | WIT_Members, WIT_External | Add to groups → `requested` |
| `processOnboardingEmailOnEdit` | WIT_Members | Email Status → `requested` (Letter 1, personal email) / `created` (Letter 2, WIT email) |
| `processContractSignedOnEdit` | WIT_Members | Contract Status → `signed` → prompt to upload signed contract |
| `processMemberStatusOnEdit` | WIT_Members | Contract/Email/Groups → Member Status `onboarding` (in progress) or `active` (complete) |

Onboarding emails are driven solely by **Email Status** (single source of truth);
they fire only on a transition *into* the target value. The `requested`
confirmation lives in `processEmailRequestOnEdit` (which runs first and reverts on
cancel, gating Letter 1); `created` gets its own confirmation in the onboarding
handler. The legacy "Intro sent" trigger has been removed.

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

> **Migrating to the dispatcher:** the live project still has the four legacy
> per-handler On-Edit triggers. Running `setupTriggers()` removes them and
> installs the single `onEditInstallable`. The interim is safe either way — if
> a legacy trigger fires a handler directly (no `ctx`), the handler rebuilds it
> via `buildEditContext(e)`.

> **Not reproducible in code:** the per-trigger *Failure notification* setting
> (the time-based jobs use "Notify me immediately" / "Notify me daily"). The
> `ScriptApp` trigger builder can't set this — re-apply it manually in the
> Triggers panel after running `setupTriggers()`. Time-based triggers also fire
> at a random minute within the configured hour, so exact run times drift.

## Conventions & constraints

- **`CONFIG` (core/Config.js) is the single source for data & identifiers** —
  sheet names, headers, statuses, folder IDs, URLs. Not visual styling.
- **`THEME` (core/Theme.js) is the single source for color/design tokens.**
  Never hard-code hex in feature files. Three consumption paths:
    - **Sheets / inline email styles** → `THEME.*` in JS.
    - **JS-built HTML** (modals, sidebars) → inject `${themeCss()}` into the
      string and use `var(--color-*)`.
    - **HtmlService templates** (`.html` guides) → `<?!= themeCss() ?>` in
      `<head>`, then `var(--color-*)`; for copy-out artifacts that can't use
      CSS vars (e.g. the email signature) print a token via `<?= THEME.x ?>`.
  `core/Theme.js` is the only `.js` file that may contain raw hex.
- **Primary UI is blue-driven (`THEME.primary` `#1b4f8a`).** The brand sheet's
  bright marketing accents live under **`THEME.brand`** (electric/cyan/navy/
  ivory/gold/coral) and are exposed as `--brand-*` CSS vars. Do **not** promote
  them to primary UI — use only on promotional/high-emphasis surfaces
  (signature template, campaign/event banners, key highlights), sparingly, to
  keep screens professional and not noisy.
- **Remaining inline hex** (planned final pass): the decorative long-tail in the
  three large `.html` guides (`FileGuide`, `OnboardingGuide`, `MemberGuide`) —
  sand neutrals and amber/green/red status tints. `themeCss()` is already wired
  into them, so finishing is a mechanical `var(--…)` swap once those semantic
  tints get tokens.
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
