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

Steps 1–2 are the same either way. After that, pick **Path A** if you're
starting from a brand-new blank Google Sheet (no data to preserve), or
**Path B** if you're adopting a script that's already deployed and live.

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

### Path A — starting from a blank Google Sheet

3A. Create a new blank spreadsheet (sheets.new), then in it: **Extensions
    → Apps Script**. This opens a new, empty, bound Apps Script project —
    ignore the default `myFunction(){}` stub it shows you.

4A. In the Apps Script editor: gear icon (**Project Settings**) → copy
    the **Script ID**. Paste it into `.clasp.json` at the repo root,
    replacing `PUT_YOUR_SCRIPT_ID_HERE`.

5A. From the repo root, push the real code up (overwriting the blank stub):
    ```
    clasp push -f
    ```
    You don't need to create any sheet tabs by hand — `doPost` creates
    each tab (Requisitions, RequisitionItems, Projects, TeamMembers, …)
    the first time the app writes to it. The sheet's original default
    tab (usually "Sheet1") is harmless; delete it whenever you like.

6A. Back in the Apps Script editor (refresh the page first): **Deploy →
    New deployment** → type **Web app** → *Execute as*: **Me**, *Who has
    access*: **Anyone** → Deploy. Copy the **Web app URL** — this is a
    brand-new deployment, so it's a *different* URL than any you used
    before. Tell me this URL and I'll update `CONFIG.gsScriptUrl` in
    `index.html` for you (or paste it into the app's Google Sheets Setup
    page yourself).

7A. Get this deployment's ID for CI to target on future auto-deploys:
    ```
    clasp deployments
    ```
    (or Apps Script editor → Manage deployments). Copy the deployment ID
    (starts with `AKfycb...`) — you'll need it in step 8.

### Path B — adopting an existing, already-deployed project

3B. Find the existing script's ID — open the Google Sheet →
    Extensions → Apps Script → gear icon (Project Settings) → "Script ID".

4B. Pull down what's actually live right now, from the repo root:
    ```
    clasp clone <SCRIPT_ID> --rootDir apps-script
    ```
    This overwrites `apps-script/Code.gs`/`appsscript.json` with whatever
    is *currently deployed*, and writes the real `scriptId` into
    `.clasp.json`.

5B. Reconcile: run `git diff apps-script/` and compare against what's
    already committed here. Once `apps-script/` in the repo is the
    version you want live, push it for real:
    ```
    clasp push -f
    ```

6B. Find the existing Web App deployment ID (the one behind the URL
    already pasted into the app's Google Sheets Setup page):
    ```
    clasp deployments
    ```
    or Apps Script editor → Deploy → Manage deployments.

### Finishing up (either path)

8. In GitHub → this repo → **Settings → Secrets and variables →
   Actions**:
   - **New repository secret** `CLASPRC_JSON` — paste the full contents
     of `~/.clasprc.json` from step 2. Treat it like a password: it can
     push code to your Apps Script project.
   - **New repository variable** `GAS_DEPLOYMENT_ID` — paste the
     deployment ID from step 7A/6B.

From then on: `clasp push` overwrites the *live* project with whatever is
in this repo — always edit `apps-script/Code.gs` here from now on, not in
the Apps Script web editor, since the next auto-deploy will silently
overwrite editor changes. Any push to `main` that changes `apps-script/**`
runs `clasp push` + `clasp deploy --deploymentId <id>` automatically,
updating the *same* deployment — so the Web App URL never changes again
after step 6A/6B.

## Editing the backend

Edit `apps-script/Code.gs` (and `appsscript.json` if you need to change
manifest settings like `webapp.access`) directly in this repo. Commit and
push to `main` — the workflow takes it from there.
