# WIT Canada — Operations Apps Script

Container-bound Google Apps Script that automates Women in Tech Canada member
onboarding from a Google Sheet. Editing the sheet drives the workflow (sending
onboarding emails, prompting for the signed contract, syncing photos/bios,
promoting member status); menu actions build report sheets and open onboarding
tools; a deployed web app serves the member onboarding checklist and the email
signature generator.

## Documentation

- **[CLAUDE.md](CLAUDE.md)** — architecture: file layout, the onEdit dispatcher
  and trigger model, and the project's conventions (`CONFIG` for data, `THEME`
  for styling, header-driven column access, batch-read discipline).
- **[DEPLOY.md](DEPLOY.md)** — release checklist (`clasp push`, `setupTriggers()`,
  web-app redeploy, smoke tests).

## Layout

```
core/   Config.js (data & identifiers) · Theme.js (design tokens) ·
        ReportBuilder.js (report scaffolding) · Utils.js (shared helpers)
app/    onEdit dispatcher (EditRouter), trigger registration (Triggers),
        menu/sidebars, web-app entry point, and the feature handlers
tools/  one-off Drive / spreadsheet structure export utilities
*.html  HtmlService pages (member guide, signature generator, doc guides)
disabled/  parked code excluded from deploy (needs Workspace admin access)
```

## Develop

```
clasp push     # deploy code to the Apps Script project
clasp open     # open the project in the editor
```

`.claspignore` keeps repo-only files (docs, `disabled/`, the visual prototype's
test tooling) out of `clasp push`. There is no local test runner — handlers are
verified against the live sheet (see DEPLOY.md). After changing triggers, run
`setupTriggers()` from the editor.
