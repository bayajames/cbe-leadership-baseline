"""
Respondent profile by school location: one combined, formatted table.

Reads the two Kobo baseline exports in this folder and saves:
  - Table_Respondent_Profile.docx  (Word table, report style)
  - Table_Respondent_Profile.xlsx  (same table in Excel)

Each characteristic shows n (%) for Rural, Urban and Total (column percentages)
and a chi-square test of association with location. Because several urban cells
are small, the p-value is a Monte Carlo (permutation) p, which stays valid with
small expected counts.

Requires: pip install pandas numpy scipy openpyxl python-docx
Run:      python profile_table.py
"""
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

# ---------------------------------------------------------------- settings
BASE = Path.cwd()
TABLE_NO = "Table 1.2"
TITLE = "Profile of Respondents by School Location"
OUT_DOCX = BASE / "Table_Respondent_Profile.docx"
OUT_XLSX = BASE / "Table_Respondent_Profile.xlsx"
N_PERMUTATIONS = 10000
LOC = "A1. Location of the Senior School"

def find_file(*keywords):
    for p in BASE.glob("*.xlsx"):
        if all(k.lower() in p.name.lower() for k in keywords) and not p.name.startswith("~$"):
            return p
    raise FileNotFoundError(f"No .xlsx containing {keywords} in {BASE}")

L = pd.read_excel(find_file("SCHOOL_LEADERS", "Baseline"))
T = pd.read_excel(find_file("TEACHERS", "Baseline"))

# Variables to include: (panel, label shown in table, dataframe, column, category order)
# Category order = order of rows in the table; categories not listed are added at the end.
YEARS5 = ["Less than 5 years", "5-10 years", "11-15 years", "More than 15 years"]
STAY = ["Less than 3 years", "3-5 years", "More than 5 years"]
ALL = pd.concat([L[[LOC, "A3. Gender"]], T[[LOC, "A3. Gender"]]], ignore_index=True)

VARIABLES = [
    ("All respondents", "Gender", ALL, "A3. Gender", ["Male", "Female", "Prefer not to say"]),
    ("School leaders", "Years in school leadership", L, "A4. Number of years in school leadership", YEARS5),
    ("School leaders", "Years in current school", L, "A5. Number of years in the current Senior School", STAY),
    ("School leaders", "Type of school", L, "A6. What type of Senior School is your school?",
     ["Day", "Boarding", "Day and Boarding"]),
    ("School leaders", "Learner enrolment", L,
     "A7. Approximate number of learners currently enrolled in the Senior School",
     ["Below 200", "200-399", "400-599", "600-799", "800 or more"]),
    ("School leaders", "Number of teachers", L,
     "A8. Approximately how many teachers currently teach in the Senior School?",
     ["Below 10", "10-19", "20-29", "30-39", "40 or more"]),
    ("Teachers", "Teaching experience", T, "A4. Number of years you have been teaching", YEARS5),
    ("Teachers", "Years in current school", T, "A5. Number of years you have taught in this Senior School", STAY),
    ("Teachers", "Received CBE-related training", T,
     "A6. Have you received professional development or training specifically related to CBE implementation?",
     ["Yes", "No"]),
    ("Teachers", "CBE training activities (past 3 years)ᵃ", T,
     "A7. If yes, how many CBE-related professional development activities have you participated in during the past three years?",
     ["1", "2-3", "4-5", "More than 5"]),
    ("Teachers", "Main subject area", T, "A8. Which area best describes your main teaching responsibility?",
     ["Sciences", "Mathematics", "Humanities/Languages", "Technical/Applied subjects"]),
]

# ---------------------------------------------------------------- statistics
rng = np.random.default_rng(2026)

def chi_square_mc(cat, loc):
    """Pearson chi-square with a Monte Carlo p-value (location labels shuffled)."""
    ct = pd.crosstab(cat, loc)
    if ct.shape[0] < 2 or ct.shape[1] < 2:
        return np.nan, 0, np.nan
    chi2, _, dof, _ = stats.chi2_contingency(ct, correction=False)
    # fast permutation: build tables with numpy instead of pandas
    ci, cats = pd.factorize(cat)
    li, locs = pd.factorize(loc)
    R, C = len(cats), len(locs)
    expected = np.outer(np.bincount(ci, minlength=R), np.bincount(li, minlength=C)) / len(ci)
    exceed = 0
    for _ in range(N_PERMUTATIONS):
        obs = np.bincount(ci * C + rng.permutation(li), minlength=R * C).reshape(R, C)
        if ((obs - expected) ** 2 / expected).sum() >= chi2 - 1e-9:
            exceed += 1
    return chi2, dof, (exceed + 1) / (N_PERMUTATIONS + 1)

def fmt_p(p):
    return "< .001" if p < .001 else f"{p:.3f}".lstrip("0")

rows = []          # (kind, label, rural, urban, total, test)  kind = panel | var | cat
current_panel = None
for panel, label, df, col, order in VARIABLES:
    d = df[[col, LOC]].dropna().copy()
    d[col] = d[col].astype(str).str.replace(r"\.0$", "", regex=True)
    if panel != current_panel:
        n_panel = len(df)
        rows.append(("panel", f"{panel} (n = {n_panel})", "", "", "", ""))
        current_panel = panel
    chi2, dof, p = chi_square_mc(d[col], d[LOC])
    test = f"χ²({dof}) = {chi2:.2f}, p = {fmt_p(p)}" if dof else ""
    rows.append(("var", label, "", "", "", test))
    ct = pd.crosstab(d[col], d[LOC])
    ct = ct.reindex([c for c in order if c in ct.index] + [c for c in ct.index if c not in order])
    ct["Total"] = ct.sum(axis=1)
    tot = ct.sum()
    for cat, r in ct.iterrows():
        cells = [f"{int(r[c])} ({r[c] / tot[c] * 100:.1f})" if c in r else "0 (0.0)"
                 for c in ["Rural", "Urban", "Total"]]
        rows.append(("cat", f"    {cat}", *cells, ""))

table = pd.DataFrame(rows, columns=["kind", "Characteristic", "Rural", "Urban", "Total", "Test"])
n_rural = int((L[LOC] == "Rural").sum() + (T[LOC] == "Rural").sum())
n_urban = int((L[LOC] == "Urban").sum() + (T[LOC] == "Urban").sum())
headers = ["Characteristic", f"Rural\nn (%)", f"Urban\nn (%)", f"Total\nn (%)", "χ² test"]
note = (f"N = {n_rural + n_urban} ({len(L)} school leaders, {len(T)} teachers; {n_rural} rural, {n_urban} urban). "
        "Percentages are column percentages within each characteristic. p-values are Monte Carlo estimates "
        f"({N_PERMUTATIONS:,} permutations) because several cells have small expected counts. "
        "ᵃ Teachers who received CBE-related training only.")

print(f"{TABLE_NO}. {TITLE}\n")
print(table.drop(columns="kind").to_string(index=False))
print("\nNote.", note)

# ---------------------------------------------------------------- Word
def set_border(cell, **edges):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge, size in edges.items():
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), "000000")
        borders.append(el)

def write(cell, text, bold=False, italic=False, left=False, size=10):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run(text)
    run.bold, run.italic = bold, italic
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"

doc = Document()
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Inches(0.9)
doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)

p = doc.add_paragraph()
p.add_run(TABLE_NO).bold = True
p.paragraph_format.space_after = Pt(0)
doc.add_paragraph().add_run(TITLE).italic = True

t = doc.add_table(rows=len(table) + 1, cols=len(headers))
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.autofit = False
widths = [Inches(2.6), Inches(0.95), Inches(0.95), Inches(0.95), Inches(1.45)]
for row in t.rows:
    for j, w in enumerate(widths):
        row.cells[j].width = w
for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
    gc.set(qn("w:w"), str(int(widths[j].twips)))

trPr = t.rows[0]._tr.get_or_add_trPr()          # repeat header row on each page
hdr = OxmlElement("w:tblHeader"); hdr.set(qn("w:val"), "true"); trPr.append(hdr)
for j, h in enumerate(headers):
    write(t.cell(0, j), h, bold=True, left=(j == 0))
    set_border(t.cell(0, j), top=12, bottom=8)

for i, r in enumerate(table.itertuples(index=False), start=1):
    if r.kind == "panel":
        merged = t.cell(i, 0).merge(t.cell(i, len(headers) - 1))
        write(merged, r.Characteristic, bold=True, italic=True, left=True)
        continue
    write(t.cell(i, 0), r.Characteristic.strip() if r.kind == "var" else "   " + r.Characteristic.strip(),
          bold=(r.kind == "var"), left=True)
    write(t.cell(i, 1), r.Rural)
    write(t.cell(i, 2), r.Urban)
    write(t.cell(i, 3), r.Total)
    write(t.cell(i, 4), r.Test, size=9)

for j in range(len(headers)):
    set_border(t.cell(len(table), j), bottom=12)

np_ = doc.add_paragraph()
np_.paragraph_format.space_before = Pt(4)
run = np_.add_run("Note. ")
run.italic = True
run.font.size = Pt(9)
np_.add_run(note).font.size = Pt(9)
doc.save(OUT_DOCX)

# ---------------------------------------------------------------- Excel
out = table.drop(columns="kind")
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
    out.to_excel(xw, sheet_name="Profile", index=False, header=False, startrow=2)
wb = load_workbook(OUT_XLSX)
ws = wb["Profile"]
ws["A1"] = f"{TABLE_NO}. {TITLE}"
ws["A1"].font = Font(bold=True, size=12)
thin, thick = Side(style="thin"), Side(style="medium")
for j, h in enumerate(headers, start=1):
    c = ws.cell(row=2, column=j, value=h.replace("\n", " "))
    c.font = Font(bold=True)
    c.border = Border(top=thick, bottom=thin)
    c.alignment = Alignment(horizontal="left" if j == 1 else "center")
for i, kind in enumerate(table["kind"], start=3):
    for j in range(1, 6):
        c = ws.cell(row=i, column=j)
        c.alignment = Alignment(horizontal="left" if j == 1 else "center")
        if kind == "panel":
            c.font = Font(bold=True, italic=True)
        elif kind == "var":
            c.font = Font(bold=True) if j == 1 else Font(size=9)
last = ws.max_row
for j in range(1, 6):
    ws.cell(row=last, column=j).border = Border(bottom=thick)
ws.cell(row=last + 1, column=1, value="Note. " + note).font = Font(italic=True, size=9)
for col, w in zip("ABCDE", [44, 13, 13, 13, 24]):
    ws.column_dimensions[col].width = w
wb.save(OUT_XLSX)

print(f"\nSaved: {OUT_DOCX.name} and {OUT_XLSX.name}")
