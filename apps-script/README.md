# Apps Script auto-deploy

`apps-script/Code.gs` and `apps-script/appsscript.json` are the source of
truth for this app's Google Sheets backend (the Apps Script Web App). The
workflow at `.github/workflows/deploy-apps-script.yml` pushes this folder
to your Apps Script project and redeploys it automatically on every push
to `main` that touches `apps-script/**`.

Google requires an authorized identity to push code into an Apps Script
project — there's no API-key-only path for a personal account — so the
steps below are a one-time manual setup you run yourself. Nothing here
should be pasted into chat; secrets go straight into GitHub's UI.

## One-time setup

1. **Enable the Apps Script API** for your account:
   https://script.google.com/home/usersettings — flip the toggle on.

2. **Install clasp and log in** (opens a Google OAuth consent screen in
   your browser):
   ```
   npm install -g @google/clasp
   clasp login
   ```
   This creates `~/.clasprc.json` on your machine — that file is what
   lets CI push on your behalf later.

3. **Find your existing script's ID** — open your Google Sheet →
   Extensions → Apps Script → gear icon (Project Settings) → "Script ID".

4. **Pull down what's actually live right now.** From the repo root:
   ```
   clasp clone <SCRIPT_ID> --rootDir apps-script
   ```
   This overwrites `apps-script/Code.gs` and `apps-script/appsscript.json`
   with whatever is *currently deployed*, and writes the real `scriptId`
   into `.clasp.json` at the repo root (replacing the
   `PUT_YOUR_SCRIPT_ID_HERE` placeholder).

5. **Reconcile.** Run `git diff apps-script/` and compare against what
   was already committed here. If your live project doesn't yet have the
   hardened `doPost`/`doGet` (locking, cheap ID lookups, real
   success/error responses), decide which version you want live, then
   make the repo match it. Once `apps-script/` in the repo is the version
   you want deployed, push it up for real:
   ```
   clasp push -f
   ```
   `clasp push` overwrites the *live* project with whatever is in this
   folder — always edit here from now on, not in the Apps Script web
   editor, since the next auto-deploy will silently overwrite editor
   changes.

6. **Find your existing Web App deployment ID** — the one behind the
   URL already pasted into the app's Google Sheets Setup page:
   ```
   clasp deployments
   ```
   or Apps Script editor → Deploy → Manage deployments. Copy the
   deployment ID (starts with `AKfycb...`).

7. In GitHub → this repo → **Settings → Secrets and variables →
   Actions**:
   - **New repository secret** `CLASPRC_JSON` — paste the full contents
     of `~/.clasprc.json` from step 2. Treat it like a password: it can
     push code to your Apps Script project.
   - **New repository variable** `GAS_DEPLOYMENT_ID` — paste the
     deployment ID from step 6. Not secret, just needs to be somewhere
     the workflow can read it.

That's it. From now on, any push to `main` that changes `apps-script/**`
runs `clasp push` + `clasp deploy --deploymentId <id>` automatically,
updating the *same* deployment — so the Web App URL already configured in
the app never changes.

## Editing the backend

Edit `apps-script/Code.gs` (and `appsscript.json` if you need to change
manifest settings like `webapp.access`) directly in this repo. Commit and
push to `main` — the workflow takes it from there.
