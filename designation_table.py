"""
Designation by school location: formatted n (%) table.

Reads Sheet1 of Book1.xlsx and saves the table in two forms:
  - Table_Designation_by_Location.docx  (Word table, APA-style borders)
  - Table_Designation_by_Location.xlsx  (formatted Excel sheet)

Requires: pip install pandas openpyxl python-docx scipy
Run:      python designation_table.py
"""
from pathlib import Path

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
INPUT_FILE = Path("Book1.xlsx")          # change if your file has another name
SHEET = "Sheet1"
ROW_VAR = "A2. Your current leadership position"
COL_VAR = "A1. Location of the Senior School"
TABLE_NO = "Table X"
TITLE = "School Location by Respondents' Current Designation"
ROW_HEADER = "Designation"
ROW_ORDER = ["Principal/Head of Senior School", "Deputy Principal/Deputy Head",
             "Senior Teacher/Departmental Leader", "Class/Year Coordinator", "Subject Teacher"]
COL_ORDER = ["Rural", "Urban"]
OUT_DOCX = Path("Table_Designation_by_Location.docx")
OUT_XLSX = Path("Table_Designation_by_Location.xlsx")

# ---------------------------------------------------------------- build the table
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET)
df = df.dropna(subset=[ROW_VAR, COL_VAR])

counts = pd.crosstab(df[ROW_VAR], df[COL_VAR])
rows = [r for r in ROW_ORDER if r in counts.index] + [r for r in counts.index if r not in ROW_ORDER]
cols = [c for c in COL_ORDER if c in counts.columns]
counts = counts.loc[rows, cols]
counts["Total"] = counts.sum(axis=1)
counts.loc["Total"] = counts.sum()

pct = counts / counts.loc["Total"] * 100            # column percentages
table = counts.astype(str) + " (" + pct.round(1).map("{:.1f}".format) + ")"
table.index.name = ROW_HEADER
headers = [ROW_HEADER] + [f"{c}\nn (%)" for c in table.columns]

chi2, p, dof, _ = stats.chi2_contingency(counts.iloc[:-1, :-1])
p_text = "< .001" if p < .001 else f"= {p:.3f}".replace("0.", ".")
note = (f"Note. N = {int(counts.loc['Total', 'Total'])}. Percentages are column percentages. "
        f"χ²({dof}) = {chi2:.2f}, p {p_text}.")

print(f"{TABLE_NO}. {TITLE}\n")
print(table.to_string())
print("\n" + note)

# ---------------------------------------------------------------- Word table
def set_border(cell, **edges):
    """Set top/bottom borders on a cell, e.g. set_border(c, top=8, bottom=8)."""
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

def write(cell, text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(text)
    run.bold, run.italic = bold, italic
    run.font.size = Pt(11)
    run.font.name = "Times New Roman"

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(11)

p1 = doc.add_paragraph()
p1.add_run(TABLE_NO).bold = True
p1.paragraph_format.space_after = Pt(0)
doc.add_paragraph().add_run(TITLE).italic = True

t = doc.add_table(rows=len(table) + 1, cols=len(headers))
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.autofit = False
widths = [Inches(2.9)] + [Inches(1.2)] * (len(headers) - 1)
for row in t.rows:
    for j, w in enumerate(widths):
        row.cells[j].width = w
for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):   # Word and LibreOffice read this
    gc.set(qn("w:w"), str(int(widths[j].twips)))

for j, h in enumerate(headers):
    c = t.cell(0, j)
    write(c, h, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    set_border(c, top=12, bottom=8)

for i, (label, row) in enumerate(table.iterrows(), start=1):
    is_total = label == "Total"
    write(t.cell(i, 0), label, bold=is_total, align=WD_ALIGN_PARAGRAPH.LEFT)
    for j, val in enumerate(row, start=1):
        write(t.cell(i, j), val, bold=is_total)
    if is_total:
        for j in range(len(headers)):
            set_border(t.cell(i, j), top=4, bottom=12)

note_p = doc.add_paragraph()
note_p.paragraph_format.space_before = Pt(4)
r = note_p.add_run("Note. ")
r.italic = True
r.font.size = Pt(10)
note_p.add_run(note.replace("Note. ", "")).font.size = Pt(10)

doc.save(OUT_DOCX)

# ---------------------------------------------------------------- Excel table
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
    table.reset_index().to_excel(xw, sheet_name="Table", index=False, startrow=2, header=False)

wb = load_workbook(OUT_XLSX)
ws = wb["Table"]
ws["A1"] = f"{TABLE_NO}. {TITLE}"
ws["A1"].font = Font(bold=True, size=12)
thin, thick = Side(style="thin"), Side(style="medium")
for j, h in enumerate(headers, start=1):
    c = ws.cell(row=2, column=j, value=h.replace("\n", " "))
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="left" if j == 1 else "center", wrap_text=True)
    c.border = Border(top=thick, bottom=thin)
last = ws.max_row
for row in ws.iter_rows(min_row=3, max_row=last):
    for c in row:
        c.alignment = Alignment(horizontal="left" if c.column == 1 else "center")
        if c.row == last:
            c.font = Font(bold=True)
            c.border = Border(top=thin, bottom=thick)
ws.cell(row=last + 1, column=1, value=note).font = Font(italic=True, size=9)
ws.column_dimensions["A"].width = 38
for col in "BCD":
    ws.column_dimensions[col].width = 16
wb.save(OUT_XLSX)

print(f"\nSaved: {OUT_DOCX.resolve()}\nSaved: {OUT_XLSX.resolve()}")

