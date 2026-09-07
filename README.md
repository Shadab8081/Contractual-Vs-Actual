# Manpower Reconciliation App

A Streamlit app that reconciles the **Contractual vs Actual** workbook's
`QIC_ Actual_LP`, `QIC_ Actual_UP` and `QIC_ Actual_5&6` sheets against an
uploaded manpower file (which has `TOTAL MANPOWER` and `LEFT` sheets).

Every time you run it: upload the manpower file → click **Run
reconciliation** → check the summary → download the updated workbook →
optionally click **Confirm** to make it the new master for next time.

PGC and the Contractual sheets are never touched. Only employee-slot cells
are edited — Position labels, Contractual Requirement values, formulas and
formatting are all preserved exactly.

## What it does automatically vs. what it flags

- **Removes** a named employee if their Iqama shows up in the LEFT sheet, or
  if TOTAL MANPOWER now shows them active under a *different* project.
- **Adds** a new/moved-in employee into an already-vacant row whose Position
  label matches their SAP POSITION or ASSIGNED WORK — an exact match, or one
  is a whole-word subset of the other. It deliberately does **not** use
  loose/fuzzy text similarity, because that can misplace someone into a
  similar-looking but wrong role (e.g. "Cleaner" vs "Carpenter").
- **Flags for manual review** rather than guessing:
  - employees active in TOTAL MANPOWER with no matching vacant slot
    (genuine new/overflow headcount — you'll need to add a row yourself), and
  - employees currently listed who aren't found in TOTAL MANPOWER or LEFT at
    all (possible data entry mismatch worth checking by hand).

## One-time setup

### 1. Create a private GitHub repository

Create a new **private** repo (the workbook has employee names and
Iqama/passport numbers, so keep it private) and push these files to it:

```
app.py
reconcile.py
github_store.py
requirements.txt
.gitignore
data/master.xlsx      <- your current Contractual vs Actual workbook
```

```bash
cd this-folder
git init
git add app.py reconcile.py github_store.py requirements.txt .gitignore data/master.xlsx README.md
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

### 2. Create a GitHub Personal Access Token

The app needs permission to read and overwrite `data/master.xlsx` in that
repo on your behalf.

1. Go to **github.com → Settings → Developer settings → Personal access
   tokens → Fine-grained tokens → Generate new token**.
2. Restrict it to **only this one repository**.
3. Under **Repository permissions**, set **Contents: Read and write**.
4. Copy the generated token (`github_pat_...`) — you won't see it again.

### 3. Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with
   GitHub.
2. **New app** → pick your repo, branch `main`, main file `app.py`.
3. Before (or right after) deploying, open **App settings → Secrets** and
   paste:
   ```toml
   GITHUB_TOKEN = "github_pat_..."
   GITHUB_REPO = "your-username/your-repo-name"
   GITHUB_FILE_PATH = "data/master.xlsx"
   GITHUB_BRANCH = "main"
   ```
4. Deploy. Streamlit Cloud can access private repos you own once you've
   granted it repo access during sign-in.

That's it — from now on, using the app is just: upload the manpower file,
click **Run reconciliation**, download, and confirm to roll the master
forward.

## Running locally (optional, for testing changes)

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit .streamlit/secrets.toml with your real token/repo
streamlit run app.py
```

## Updating the reconciliation logic later

All the matching/editing logic lives in `reconcile.py`, independent of the
Streamlit UI in `app.py`. If your workbook's layout changes (new discipline
sections, renamed columns, etc.), that's the file to update.
