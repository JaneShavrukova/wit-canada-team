# Deploy checklist

Steps to push this Apps Script project and bring the live spreadsheet up to
date. Run top to bottom. Most of this is one-time housekeeping after the
refactor; for routine pushes only steps 1–2 (and 5 if the web app changed) apply.

> This file is repo-only — `.claspignore` excludes `**/*.md`, so it is not
> pushed to Apps Script.

## Manual deploy: commit → push → clasp push

The everyday flow after a change. Commit the docs in the **same** commit as the
code (see CLAUDE.md → Standing instructions):

```
git add <files>
git commit -m "…"
git push origin main      # → GitHub
clasp push -f             # → Apps Script project (updates HEAD)
```

After `clasp push`:

- **Sheet automations** (the onEdit dispatcher, time-driven jobs, menu actions)
  run the new code immediately — nothing else needed.
- **Trigger wiring changed?** Run `setupTriggers()` from the editor (§2).
- **Served pages changed** (`MemberGuide`, `SignatureGenerator`, or `doGet`)?
  Cut a new web-app version (§5) — `clasp push` alone does not update the live
  `/exec` URL.

The numbered steps below are the fuller checklist (mostly one-time setup on a
fresh deploy).

## 1. Push the code

```
clasp push
```

Deploys every tracked file except those in `.claspignore`. Verify it reports
the expected file count and no errors.

## 2. Register triggers  *(after a fresh deploy or any trigger change)*

All installable triggers — the single `onEditInstallable` dispatcher, the form
submit, and the scheduled jobs — are installed by `setupTriggers()`. (The live
project has already been migrated off the legacy per-handler On-Edit triggers;
re-run this only on a fresh copy or after changing trigger wiring.)

- [ ] Open the script: in the sheet, **Extensions → Apps Script**.
- [ ] In the editor, select the **`setupTriggers`** function and **Run** it.
      (Authorize if prompted.) It is idempotent — it clears the triggers it owns
      and reinstalls the current set.
- [ ] Select **`listTriggers`** and **Run** it → open **Executions** (or View →
      Logs) and confirm exactly these **5** handlers are registered:
      `onEditInstallable`, `handleProfileFormSubmit`, `syncPhotosAndBios`,
      `buildEmailRequestsReport`, `sendWeeklyEmailRequestsReport`.

Handlers also rebuild their own context via `buildEditContext` if ever called
directly without a `ctx`, so the dispatcher swap is safe either way.

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

- [ ] Set **Contract Status → `signed`** → prompt to upload the signed contract.
- [ ] Set **Email Status → `requested`** → confirm dialog; the activation
      heads-up email goes to the member's **personal** inbox + Ops is notified.
- [ ] Set **Email Status → `created`** → confirm dialog; the onboarding checklist
      email goes to the member's **WIT** inbox.
- [ ] Tick a group + set **Add to groups → `requested`** → groups-request dialog.
- [ ] Set Contract `signed` + Email `active` + Groups `added` → **Member Status**
      auto-flips to `active` (it sits at `onboarding` while still in progress).

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

## 6. Verify the photo/bio sync

Confirm the report builder is healthy:

- [ ] Menu → **⚙️ Tools → Actions → Refresh Photos & Bios** → completes
      without error and the `Report_Photos&Bios` sheet rebuilds.

---

For the everyday push flow, see **Manual deploy: commit → push → clasp push** at
the top of this file. The numbered steps above are the fuller / one-time
checklist.
