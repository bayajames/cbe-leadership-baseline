# Baseline analysis: quick start (Windows + VS Code)

Run these once in the VS Code terminal (PowerShell), inside the *Data Analysis* folder:

```powershell
# 1. Check Python is installed (3.11 or newer)
py --version

# 2. Create and activate a virtual environment
py -m venv .venv
.\.venv\Scripts\Activate.ps1
#    If you see "running scripts is disabled", run this once, then activate again:
#    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# 3. Install the packages
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Then:

1. Install the **Python** and **Jupyter** extensions in VS Code (Ctrl + Shift + X).
2. Open `Baseline_Analysis.ipynb`.
3. Click **Select Kernel** (top right) → **Python Environments** → **.venv**.
4. Click **Run All**, or press Shift + Enter to go cell by cell.

Results are saved to:

- `outputs/Baseline_Analysis_Tables.xlsx`: all tables
- `outputs/figures/`: four charts (PNG)

| Step | What it does |
|---|---|
| 1–3 | Load libraries, find and read the two Kobo exports, inspect columns |
| 4–6 | Keep consenting respondents, convert Likert labels to 1–5, build scales |
| 7 | Data-quality checks (duplicates, speeders, straight-lining, school-code/location conflicts) |
| 8 | Reliability (Cronbach's alpha) |
| 9 | Respondent profile by location (cross-tabs, chi-square) |
| 10–11 | Scale and item descriptives; rural vs urban (Mann-Whitney U, rank-biserial r) |
| 12 | Leader vs teacher perception gaps |
| 13 | Leadership–preparedness correlations; PD dose-response |
| 14 | PLC readiness and multi-select items |
| 15 | Thematic coding of open-ended answers |
| 16 | Sensitivity checks |
| 17–18 | Charts and Excel export |
