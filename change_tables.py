"""
Tables for Section 3.2 (nature of changes) and 3.3 (factors influencing the changes).

  Table 3.3  Nature of changes in leadership practice: themes from open-ended answers
             (leaders D12 "change you made", teachers D11 "change observed"), with an example answer
  Table 3.4  Factors influencing leaders' changes (multiple-response question D13), by location

Saves Change_Tables.docx (and prints both tables).
Requires: pip install pandas openpyxl python-docx
Run:      python change_tables.py
"""
import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

BASE = Path.cwd()
OUT_DOCX = BASE / "Change_Tables.docx"
THEME_TABLE_NO, FACTOR_TABLE_NO = "Table 3.3", "Table 3.4"
LOC = "A1. Location of the Senior School"

def find_file(*keywords):
    for p in BASE.glob("*.xlsx"):
        if all(k.lower() in p.name.lower() for k in keywords) and not p.name.startswith("~$"):
            return p
    raise FileNotFoundError(f"No .xlsx containing {keywords} in {BASE}")

L = pd.read_excel(find_file("SCHOOL_LEADERS", "Baseline"))
T = pd.read_excel(find_file("TEACHERS", "Baseline"))

def q(df, start):
    return next(c for c in df.columns if c.startswith(start))

# ---------------------------------------------------------------- Table 3.3: themes
EMPTY = re.compile(r"^\s*(n/?a|na|nan|none\.?|nothing|no|not applicable|no thank.*)\s*$", re.I)
CHANGES = {   # theme -> keywords (edit to refine the coding)
    "Collaborative and consultative decision-making":
        r"collab|consult|involv|decision|democratic|team|dialogue|inclusive|sharing|stakeholder|parent|communit",
    "Learner-centred teaching and learner participation":
        r"learner.?cent|learner involvement|student|learners|noisy|interactive|hands.?on|practical|teaching method|methodolog|approach",
    "Provision and mobilisation of resources": r"resource|material|facilit|provision|mobilis",
    "ICT and technology integration": r"ict|technolog|digital|computer|internet",
    "Monitoring, supervision and assessment": r"monitor|assess|supervis|performance|progress",
    "Restructuring of roles and structures":
        r"restructur|department|dean|pathway leader|implementation group|section head|employ|realign|plan",
}

def code(df, col):
    text = df[col].astype(str).str.replace("\n", " ").str.strip()
    valid = df[col].notna() & ~text.str.match(EMPTY)
    hits = pd.DataFrame({k: text.str.contains(p, case=False, regex=True) & valid for k, p in CHANGES.items()})
    return text, valid, hits

lt, lv, lh = code(L, q(L, "D12."))
tt, tv, th = code(T, q(T, "D11."))
n_l, n_t = int(lv.sum()), int(tv.sum())

# Hand-picked verbatim examples (leave a theme out to let the script pick one automatically)
EXAMPLES = {
    "Collaborative and consultative decision-making": "I have made leadership more collaborative",
    "Learner-centred teaching and learner participation":
        "The teaching methods have changed, from teacher centered to learner centered",
    "Provision and mobilisation of resources":
        "School support in requisitions of required teaching and learning resources",
    "Monitoring, supervision and assessment": "Monitoring of CBE related assessments and collaboration of teachers",
}

def example(theme):
    if theme in EXAMPLES:
        return f"\u201c{EXAMPLES[theme]}\u201d"
    """Shortest clear answer (25-160 characters) that falls under this theme only."""
    pool = []
    for text, hits in [(lt, lh), (tt, th)]:
        only = hits[theme] & (hits.sum(axis=1) == 1)
        pool += [s for s in text[only] if 25 <= len(s) <= 160]
    if not pool:
        for text, hits in [(lt, lh), (tt, th)]:
            pool += [s for s in text[hits[theme]] if 25 <= len(s) <= 160]
    return f"“{sorted(pool, key=len)[0].rstrip('.')}”" if pool else ""

theme_rows = []
for k in CHANGES:
    a, b = int(lh[k].sum()), int(th[k].sum())
    theme_rows.append([k, f"{a} ({a / n_l * 100:.1f})", f"{b} ({b / n_t * 100:.1f})", example(k), a + b])
themes = (pd.DataFrame(theme_rows, columns=["Theme", "Leaders", "Teachers", "Example response", "_n"])
          .sort_values("_n", ascending=False).drop(columns="_n"))

changed = L[q(L, "D11.")].value_counts()
theme_note = (f"Leaders: n = {n_l} who described a change they made (D12); teachers: n = {n_t} who described a "
              f"change observed in leadership (D11). Overall, {changed.get('Yes', 0)} of {len(L)} leaders "
              f"({changed.get('Yes', 0) / len(L) * 100:.1f}%) reported changing their leadership practices. "
              "Responses were coded by keyword; one response may fall under more than one theme, so "
              "percentages do not sum to 100.")

# ---------------------------------------------------------------- Table 3.4: factors (D13)
prefix = "D13. What has mainly influenced the changes you have made?/"
opts = [c for c in L.columns if c.startswith(prefix)]
answered = L[opts[0]].notna()
d = L[answered]
n_r, n_u = int((d[LOC] == "Rural").sum()), int((d[LOC] == "Urban").sum())
f_rows = []
for c in opts:
    x = pd.to_numeric(d[c], errors="coerce").fillna(0)
    r, u = int(x[d[LOC] == "Rural"].sum()), int(x[d[LOC] == "Urban"].sum())
    if r + u == 0:
        continue
    f_rows.append([c.split("/", 1)[1], f"{r} ({r / n_r * 100:.1f})", f"{u} ({u / n_u * 100:.1f})",
                   f"{r + u} ({(r + u) / len(d) * 100:.1f})", r + u])
factors = (pd.DataFrame(f_rows, columns=["Factor", "Rural", "Urban", "Total", "_n"])
           .sort_values("_n", ascending=False).drop(columns="_n"))
other = L.loc[answered, q(L, "Please specify other:.1")].dropna() if "Please specify other:.1" in L.columns else []
factor_note = (f"School leaders who reported changes, n = {len(d)} (rural n = {n_r}, urban n = {n_u}). "
               "Multiple responses allowed; percentages are of respondents and do not sum to 100."
               + (f" “Other”: {'; '.join(map(str, other))}." if len(other) else ""))

print(THEME_TABLE_NO, "\n", themes.to_string(index=False), "\n", theme_note, "\n")
print(FACTOR_TABLE_NO, "\n", factors.to_string(index=False), "\n", factor_note)

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

def add_table(doc, number, title, df, headers, widths, note, left_cols=(0,), italic_cols=()):
    p = doc.add_paragraph(); p.add_run(number).bold = True; p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph().add_run(title).italic = True
    t = doc.add_table(rows=len(df) + 1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = w
    for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j].twips)))
    for j, h in enumerate(headers):
        write(t.cell(0, j), h, bold=True, left=(j in left_cols))
        set_border(t.cell(0, j), top=12, bottom=8)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, v in enumerate(row):
            write(t.cell(i, j), v, left=(j in left_cols), italic=(j in italic_cols), size=9 if j in italic_cols else 10)
    for j in range(len(headers)):
        set_border(t.cell(len(df), j), bottom=12)
    n = doc.add_paragraph(); n.paragraph_format.space_before = Pt(4)
    r = n.add_run("Note. "); r.italic = True; r.font.size = Pt(9)
    n.add_run(note).font.size = Pt(9)
    doc.add_paragraph()

doc = Document()
s = doc.sections[0]
s.left_margin = s.right_margin = Inches(0.9)
doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)

add_table(doc, THEME_TABLE_NO, "Nature of Changes in Leadership Practices Since the Introduction of CBE",
          themes, ["Theme", f"Leaders\nn (%)", f"Teachers\nn (%)", "Example response"],
          [Inches(2.1), Inches(0.9), Inches(0.9), Inches(2.8)], theme_note, left_cols=(0, 3), italic_cols=(3,))
add_table(doc, FACTOR_TABLE_NO, "Factors Influencing Changes in Leadership Practices (School Leaders)",
          factors, ["Factor", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)"],
          [Inches(2.9), Inches(1.15), Inches(1.15), Inches(1.15)], factor_note)
doc.save(OUT_DOCX)
print(f"\nSaved: {OUT_DOCX.name}")
