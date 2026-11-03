"""
Table for Section 5.3: comparison of leaders' and teachers' assessments of teacher preparedness.

Leaders' items (G1-G8) are paired with the teacher self-rating item that asks the same thing (F).
For each pair: Leaders M (SD) | Teachers M (SD) | Difference | p (Mann-Whitney U) | r (effect size)
plus the overall scale means, overall and within rural and urban schools.

Saves Comparison_Table.docx (and prints the table).
Requires: pip install pandas numpy scipy openpyxl python-docx
Run:      python comparison_table.py
"""
import re
from pathlib import Path

import pandas as pd
from scipy import stats
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

BASE = Path.cwd()
OUT_DOCX = BASE / "Comparison_Table.docx"
TABLE_NO = "Table 5.3"
TITLE = "Comparison of School Leaders' and Teachers' Assessments of Teacher Preparedness"
LOC = "A1. Location of the Senior School"

# Leader item -> matching teacher item, with a short label for the table
PAIRS = [
    ("G1", "F1", "Understand the main requirements of CBE"),
    ("G2", "F4", "Adapt teaching approaches to learners' needs"),
    ("G3", "F3", "Confident in using learner-centred approaches"),
    ("G4", "F6", "Implement CBE-related assessment practices"),
    ("G5", "F7", "Use assessment information to support learner progress"),
    ("G6", "F10", "Collaborate with colleagues to address CBE challenges"),
    ("G7", "F9", "Confident in addressing CBE implementation challenges"),
    ("G8", "F11", "Professional learning has strengthened CBE practice"),
]

def find_file(*keywords):
    for p in BASE.glob("*.xlsx"):
        if all(k.lower() in p.name.lower() for k in keywords) and not p.name.startswith("~$"):
            return p
    raise FileNotFoundError(f"No .xlsx containing {keywords} in {BASE}")

L = pd.read_excel(find_file("SCHOOL_LEADERS", "Baseline"))
T = pd.read_excel(find_file("TEACHERS", "Baseline"))

def item(df, code):
    col = next(c for c in df.columns if c.startswith(code + ". "))
    return pd.to_numeric(df[col].astype(str).str.extract(r"^\s*(\d)\s*=")[0], errors="coerce")

def compare(a, b):
    """a = leaders, b = teachers. r > 0 means leaders rated higher."""
    a, b = a.dropna(), b.dropna()
    res = stats.mannwhitneyu(a, b, alternative="two-sided")
    r = 2 * res.statistic / (len(a) * len(b)) - 1
    p = "< .001" if res.pvalue < .001 else f"{res.pvalue:.3f}".lstrip("0")
    return [f"{a.mean():.2f} ({a.std():.2f})", f"{b.mean():.2f} ({b.std():.2f})",
            f"{a.mean() - b.mean():+.2f}", p, f"{r:.2f}".replace("0.", ".")]

rows = []
for g, f, label in PAIRS:
    rows.append(["item", f"{label} ({g} / {f})", *compare(item(L, g), item(T, f))])

L_G = pd.concat([item(L, f"G{i}") for i in range(1, 9)], axis=1).mean(axis=1)
T_F = pd.concat([item(T, f"F{i}") for i in range(1, 13)], axis=1).mean(axis=1)
rows.append(["total", "Overall preparedness scale", *compare(L_G, T_F)])
for loc in ["Rural", "Urban"]:
    rows.append(["sub", f"    {loc} schools", *compare(L_G[L[LOC] == loc], T_F[T[LOC] == loc])])

table = pd.DataFrame(rows, columns=["kind", "Aspect of preparedness", "Leaders M (SD)", "Teachers M (SD)",
                                    "Difference", "p", "r"])
n_lr, n_lu = int((L[LOC] == "Rural").sum()), int((L[LOC] == "Urban").sum())
n_tr, n_tu = int((T[LOC] == "Rural").sum()), int((T[LOC] == "Urban").sum())
note = (f"Leaders (n = {len(L)}; rural {n_lr}, urban {n_lu}) rated their teachers (G1–G8); teachers "
        f"(n = {len(T)}; rural {n_tr}, urban {n_tu}) rated themselves (matching F items shown in brackets). "
        "Scale: 1 = strongly disagree, 5 = strongly agree. Difference = leaders minus teachers. "
        "p = Mann-Whitney U test; r = rank-biserial effect size (negative = teachers rated higher; "
        "about .1 small, .3 medium, .5 large). Overall scale: leaders G1–G8, teachers F1–F12.")

print(TABLE_NO, TITLE)
print(table.drop(columns="kind").to_string(index=False))
print("Note.", note)

# ---------------------------------------------------------------- Word
def set_border(cell, **edges):
    tcPr = cell._tc.get_or_add_tcPr()
    b = tcPr.first_child_found_in("w:tcBorders")
    if b is None:
        b = OxmlElement("w:tcBorders"); tcPr.append(b)
    for edge, size in edges.items():
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(size)); el.set(qn("w:color"), "000000")
        b.append(el)

def write(cell, text, bold=False, italic=False, left=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(str(text)); run.bold, run.italic = bold, italic
    run.font.size = Pt(size); run.font.name = "Times New Roman"

doc = Document()
s = doc.sections[0]
s.left_margin = s.right_margin = Inches(0.8)
doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)

p = doc.add_paragraph(); p.add_run(TABLE_NO).bold = True; p.paragraph_format.space_after = Pt(0)
doc.add_paragraph().add_run(TITLE).italic = True
headers = ["Aspect of preparedness", "Leaders\nM (SD)", "Teachers\nM (SD)", "Difference", "p", "r"]
widths = [Inches(3.1), Inches(1.0), Inches(1.0), Inches(0.85), Inches(0.6), Inches(0.5)]
t = doc.add_table(rows=len(table) + 1, cols=len(headers))
t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
for row in t.rows:
    for j, w in enumerate(widths):
        row.cells[j].width = w
for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
    gc.set(qn("w:w"), str(int(widths[j].twips)))
for j, h in enumerate(headers):
    write(t.cell(0, j), h, bold=True, left=(j == 0))
    set_border(t.cell(0, j), top=12, bottom=8)
for i, r in enumerate(table.itertuples(index=False), start=1):
    bold = r.kind == "total"
    for j, v in enumerate(list(r)[1:]):
        label = ("   " + v.strip()) if (j == 0 and r.kind == "sub") else v
        write(t.cell(i, j), label, bold=bold, italic=(r.kind == "sub"), left=(j == 0))
        if bold:
            set_border(t.cell(i, j), top=4)
for j in range(len(headers)):
    set_border(t.cell(len(table), j), bottom=12)
n = doc.add_paragraph(); n.paragraph_format.space_before = Pt(4)
run = n.add_run("Note. "); run.italic = True; run.font.size = Pt(9)
n.add_run(note).font.size = Pt(9)
doc.save(OUT_DOCX)
print(f"Saved: {OUT_DOCX.name}")
