# Marketing Performance Dashboard

Auto-generated dashboard that rebuilds every time you push updated Excel files.

## Live Dashboard
👉 `https://YOUR-USERNAME.github.io/marketing-dashboard/`
*(replace YOUR-USERNAME with your GitHub username)*

---

## Repository Structure

```
├── index.html                    ← Generated automatically (don't edit manually)
├── dashboard_template.html       ← HTML shell with placeholders
├── generate_dashboard.py         ← Script that reads Excel → builds index.html
├── Marketing_Performance.xlsx    ← Quarterly BU data (edit this to update)
├── Enterprise_Comparison.xlsx    ← Enterprise model-level Apr vs May data
├── CV_Comparison.xlsx            ← CV model-level Apr vs May data
└── .github/
    └── workflows/
        └── generate.yml          ← GitHub Actions: auto-runs on every push
```

---

## How to Update Data

1. **Edit** any of the three Excel files on your computer
2. **Go to your GitHub repo** → click the file name → click the upload/replace button
3. **GitHub Actions automatically runs** `generate_dashboard.py`
4. **Within ~1 minute**, the live dashboard refreshes

That's it. No coding required.

---

## One-Time Setup (only do this once)

### Step 1 — Create the repo
1. Go to [github.com](https://github.com) → sign in
2. Click **+** (top right) → **New repository**
3. Name it `marketing-dashboard`
4. Set to **Public**
5. Check **"Add a README file"**
6. Click **Create repository**

### Step 2 — Upload all files
Click **Add file → Upload files** and upload **all files** from this folder:
- `index.html`
- `dashboard_template.html`
- `generate_dashboard.py`
- `Marketing_Performance.xlsx`
- `Enterprise_Comparison.xlsx`
- `CV_Comparison.xlsx`
- The `.github/workflows/generate.yml` file *(upload this separately — create the folder path manually)*

> **To upload the workflow file:**
> In your repo, click **Add file → Create new file**
> Type the name as: `.github/workflows/generate.yml`
> Paste the contents of the `generate.yml` file → Commit

### Step 3 — Enable GitHub Pages
1. Go to repo **Settings** → **Pages** (left sidebar)
2. Under **Branch** → select `main` and folder `/ (root)`
3. Click **Save**
4. Wait ~1 minute → your dashboard is live at `https://YOUR-USERNAME.github.io/marketing-dashboard/`

### Step 4 — Trigger first build
Go to **Actions** tab → click **Generate Dashboard** → click **Run workflow** → **Run workflow**

---

## Excel File Format

### Marketing_Performance.xlsx
Each sheet is named: `Enterprise`, `Enterprise_Cars`, `Enterprise_Bikes`, `CPS`, `NCBD`, `CV`, `NewAutoClassified`

Each sheet has this structure:
| | Q1 | Q2 | Q3 | Q4 | Apr'26 | May'26 | Jun'26 |
|---|---|---|---|---|---|---|---|
| Spends | ... | | | | | | |
| Revenue | ... | | | | | | |
| Leads | ... | | | | | | |
| TLeads | ... | | | | | | |

### Enterprise_Comparison.xlsx
Sheets named `Apr'26` and `May'26`
Columns: Brand, Model, Segment, Channel, Validation%, SoldCPL, Spends, Leads, TLeads, Revenue, Margin

### CV_Comparison.xlsx
Sheet named `Data`
Columns: Brand, Model, BU, then groups of 3 months (Apr, May, Jun) for: Spends, Leads, CPL, TLeads, TCPL, Margin, Validation
