# Deploy checklist

Steps to push this Apps Script project and bring the live spreadsheet up to
date. Run top to bottom. Most of this is one-time housekeeping after the
refactor; for routine pushes only steps 1–2 (and 5 if the web app changed) apply.

> This file is repo-only — `.claspignore` excludes `**/*.md`, so it is not
> pushed to Apps Script.

## 1. Push the code

```
clasp push
```

Deploys every tracked file except those in `.claspignore`. Verify it reports
the expected file count and no errors.

## 2. Migrate triggers to the single dispatcher  *(one-time, after Phase 3)*

The live project still has the four legacy per-handler On-Edit triggers. Switch
to the single `onEditInstallable` dispatcher:

- [ ] Open the script: in the sheet, **Extensions → Apps Script**.
- [ ] In the editor, select the **`setupTriggers`** function and **Run** it.
      (Authorize if prompted.) It removes the legacy triggers and installs the
      current set.
- [ ] Select **`listTriggers`** and **Run** it → open **Executions** (or View →
      Logs) and confirm exactly these **5** handlers are registered:
      `onEditInstallable`, `handleProfileFormSubmit`, `syncPhotosAndBios`,
      `buildEmailRequestsReport`, `sendWeeklyEmailRequestsReport`.

Safe either way: until `setupTriggers()` runs, the legacy triggers keep working
(handlers rebuild their context via `buildEditContext`).

## 3. Re-apply failure-notification settings  *(after any `setupTriggers()` run)*

`setupTriggers()` recreates triggers via the API, which **cannot** set the
"notify me on failure" frequency — recreated triggers revert to the default
(daily). Re-set them by hand:

- [ ] Left sidebar → **Triggers** (⏰ icon).
- [ ] For each trigger below, hover the row → **pencil (edit)** → set
      **Failure notification settings** (top-right) → **Save**.

| Trigger | Failure notification |
|---|---|
| `syncPhotosAndBios` | Notify me **immediately** |
| `buildEmailRequestsReport` | Notify me **immediately** |
| `sendWeeklyEmailRequestsReport` | Notify me **daily** |

(Cosmetic only — controls how fast you're emailed when a scheduled job fails.)

## 4. Smoke-test the edit handlers

On the **WIT_Members** sheet, confirm the dispatcher routes correctly (use a
test row):

- [ ] Set **Email Status → `requested`** → confirmation dialog appears.
- [ ] Tick a group + set **Add to groups → `requested`** → groups-request dialog.
- [ ] Tick **Intro sent** on a `created`/`active` member → onboarding email path.
- [ ] Set Contract `signed` + Email `active` + Groups `added` → **Member Status**
      auto-flips to `active`.

## 5. Redeploy the web app  *(only if guides / signature / web app changed)*

The member guide and signature generator are served from a **versioned web-app
deployment** (`WEB_APP_URL` in `core/Config.js`). `clasp push` updates the code,
but the live `/exec` URL keeps serving the old version until you redeploy:

- [ ] **Deploy → Manage deployments** → edit the active deployment (pencil) →
      **Version: New version** → **Deploy**.
- [ ] Open `WEB_APP_URL` and confirm the guide/signature render with the new
      theme. (The signature's brand-blue comes from `THEME.brand.electric`.)

> Editor previews (`createTemplateFromFile`) always reflect the latest push, so
> menu modals/sidebars update immediately — only the public web-app URL needs a
> new version.

> **clasp gotchas (don't break the web app):** `clasp deploy --deploymentId
> AKfycbxh3-EQ7VUbbf1WMn9q9aSBBCnSimhRST5QwEjs6VDXij07JwJQMP0Md99DqpqrFNmU
> --description "…"` is now safe ONLY because `appsscript.json` has a `webapp`
> block — without it, clasp snapshots the web app as a plain version and the
> live URL serves "unable to open the file". Also: keep served `.html` at the
> repo **root** — `createTemplateFromFile` can't resolve folder-prefixed names
> (e.g. `ui/MemberGuide`), even though clasp "folders" work for `.gs` files.

## 6. Verify the photo/bio fix

`buildPhotoBioReportSheet` had a runtime bug (undefined `ss`) fixed in the theme
commit. Confirm it's healthy:

- [ ] Menu → **WIT Operations → Actions → Refresh Photos & Bios** → completes
      without error and the `Report_Photos&Bios` sheet rebuilds.

---

### Routine pushes (after the one-time migration is done)

1. `clasp push`
2. If guides/signature/web app changed → step 5 (new web-app version).
