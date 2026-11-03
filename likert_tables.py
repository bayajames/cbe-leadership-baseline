"""
Section tables for the Likert-scale items (leadership, adaptation, context, preparedness).

For every section it builds one table with, per item:
  Overall M (SD) | % rating 4-5 | Rural M (SD) | Urban M (SD) | p (Mann-Whitney U, rural vs urban)
and a final row for the composite (scale mean), with Cronbach's alpha in the note.

Saves:  Likert_Tables.docx (all tables, report style)  and  Likert_Tables.xlsx (one sheet per table)

Requires: pip install pandas numpy scipy openpyxl python-docx
Run:      python likert_tables.py
"""
import re
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

BASE = Path.cwd()
FIRST_TABLE = 3              # numbering continues after Table 1.2 -> 1.3, 1.4, ...
TABLE_PREFIX = "Table 1."
OUT_DOCX = BASE / "Likert_Tables.docx"
OUT_XLSX = BASE / "Likert_Tables.xlsx"
LOC = "A1. Location of the Senior School"

def find_file(*keywords):
    for p in BASE.glob("*.xlsx"):
        if all(k.lower() in p.name.lower() for k in keywords) and not p.name.startswith("~$"):
            return p
    raise FileNotFoundError(f"No .xlsx containing {keywords} in {BASE}")

L = pd.read_excel(find_file("SCHOOL_LEADERS", "Baseline"))
T = pd.read_excel(find_file("TEACHERS", "Baseline"))

def cols(df, pattern):
    return [c for c in df.columns if re.match(pattern, c)]

FREQ = ("1 = Never, 5 = Very often", "% Often/ Very often")
CHANGE = ("1 = Not changed, 5 = Changed to a very great extent", "% Great/ Very great")
EXTENT = ("1 = Not at all, 5 = To a very great extent", "% Great/ Very great")
AGREE = ("1 = Strongly disagree, 5 = Strongly agree", "% Agree/ Strongly agree")

# (title, respondent label, dataframe, item columns, scale) - one entry = one table
# For the leaders' context section, sub-scales are given as a list of (sub-heading, columns).
SECTIONS = [
    ("Management Leadership Practices Reported by School Leaders", "School leaders", L, cols(L, r"B1\.\d+\. "), FREQ),
    ("Management Leadership Practices as Observed by Teachers", "Teachers", T, cols(T, r"B1\.\d+\. "), FREQ),
    ("Instructional Leadership Practices Reported by School Leaders", "School leaders", L, cols(L, r"C\d+\. I "), FREQ),
    ("Instructional Leadership Practices as Observed by Teachers", "Teachers", T, cols(T, r"C\d\. School"), FREQ),
    ("Extent of Change in Leadership Practices Since CBE: School Leaders", "School leaders", L, cols(L, r"D([1-9]|10)\. "), CHANGE),
    ("Extent of Change in Leadership Practices Since CBE: Teachers' Observations", "Teachers", T, cols(T, r"D([1-9]|10)\. "), CHANGE),
    ("Contextual Conditions Affecting CBE Implementation: School Leaders", "School leaders", L,
     [("Resources and infrastructure", cols(L, r"E1\.\d\. ")),
      ("Human resources and workload", cols(L, r"E2\.\d\. ")),
      ("Professional and institutional support", cols(L, r"E3\.\d\. ")),
      ("School and community environment", cols(L, r"E4\.\d\. "))], EXTENT),
    ("Contextual Conditions Affecting CBE Implementation: Teachers", "Teachers", T, cols(T, r"E([1-9]|1[0-2])\. "), EXTENT),
    ("Leadership Support Provided to Teachers (School Leaders)", "School leaders", L, cols(L, r"F\d\. I "), FREQ),
    ("Teacher Preparedness for CBE: School Leaders' Assessment", "School leaders", L, cols(L, r"G[1-8]\. "), AGREE),
    ("Teacher Preparedness for CBE: Teachers' Self-Assessment", "Teachers", T, cols(T, r"F([1-9]|1[0-2])\. "), AGREE),
]

# ---------------------------------------------------------------- helpers
def to_num(s):
    return pd.to_numeric(s.astype(str).str.extract(r"^\s*(\d)\s*=")[0], errors="coerce")

def alpha(items):
    d = items.dropna()
    k = d.shape[1]
    return k / (k - 1) * (1 - d.var(ddof=1).sum() / d.sum(axis=1).var(ddof=1))

def mw_p(x, loc):
    r, u = x[loc == "Rural"].dropna(), x[loc == "Urban"].dropna()
    return stats.mannwhitneyu(r, u, alternative="two-sided").pvalue

def ms(x):
    return f"{x.mean():.2f} ({x.std():.2f})"

def fmt_p(p):
    return "< .001" if p < .001 else f"{p:.3f}".lstrip("0")

def short(item):
    text = item.split("\n")[0].strip()
    text = re.sub(r"^(B1\.\d+|[A-Z]\d+(\.\d+)?)\.\s*", lambda m: m.group(1) + " ", text)
    return text.rstrip(".")

def item_row(df, c, label=None):
    x = to_num(df[c])
    return {"Item": label or short(c), "Overall M (SD)": ms(x), "% 4-5": f"{(x >= 4).mean() * 100:.1f}",
            "Rural M (SD)": ms(x[df[LOC] == "Rural"]), "Urban M (SD)": ms(x[df[LOC] == "Urban"]),
            "p": fmt_p(mw_p(x, df[LOC])), "kind": "item"}

def composite_row(df, items, label):
    data = pd.concat([to_num(df[c]) for c in items], axis=1)
    x = data.mean(axis=1)
    row = {"Item": label, "Overall M (SD)": ms(x), "% 4-5": f"{(x >= 4).mean() * 100:.1f}",
           "Rural M (SD)": ms(x[df[LOC] == "Rural"]), "Urban M (SD)": ms(x[df[LOC] == "Urban"]),
           "p": fmt_p(mw_p(x, df[LOC])), "kind": "composite"}
    return row, alpha(data)

# ---------------------------------------------------------------- build tables
tables = []
for k, (title, who, df, spec, (scale_txt, pct_label)) in enumerate(SECTIONS):
    rows, alphas = [], []
    groups = spec if isinstance(spec[0], tuple) else [(None, spec)]
    all_items = []
    for sub, items in groups:
        if sub:
            rows.append({"Item": sub, "kind": "sub"})
        rows += [item_row(df, c) for c in items]
        all_items += items
        if sub:
            r, a = composite_row(df, items, f"{sub} (sub-scale mean)")
            rows.append(r)
            alphas.append(f"{sub.lower()} α = {a:.2f}".replace("0.", "."))
    r, a = composite_row(df, all_items, "Overall scale mean")
    r["kind"] = "total"
    rows.append(r)
    alphas.append(f"overall α = {a:.2f}".replace("0.", "."))
    n_r, n_u = int((df[LOC] == "Rural").sum()), int((df[LOC] == "Urban").sum())
    note = (f"{who}, N = {len(df)} (rural n = {n_r}, urban n = {n_u}). Scale: {scale_txt}. "
            f"% 4-5 = share rating 4 or 5. p = Mann-Whitney U test, rural vs urban. "
            f"Cronbach's {'; '.join(alphas)}.")
    tables.append((f"{TABLE_PREFIX}{FIRST_TABLE + k}", title, pd.DataFrame(rows), pct_label, note))

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
    run = p.add_run(str(text))
    run.bold, run.italic = bold, italic
    run.font.size = Pt(size); run.font.name = "Times New Roman"

doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = sec.page_height, sec.page_width
sec.left_margin = sec.right_margin = Inches(0.8)
doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)

COLS = ["Item", "Overall M (SD)", "% 4-5", "Rural M (SD)", "Urban M (SD)", "p"]
WIDTHS = [Inches(4.6), Inches(1.1), Inches(0.95), Inches(1.1), Inches(1.1), Inches(0.6)]

for i, (num, title, tb, pct_label, note) in enumerate(tables):
    if i:
        doc.add_page_break()
    p = doc.add_paragraph(); p.add_run(num).bold = True; p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph().add_run(title).italic = True
    headers = ["Item", "Overall\nM (SD)", pct_label.replace("/ ", "/\n"), "Rural\nM (SD)", "Urban\nM (SD)", "p"]
    t = doc.add_table(rows=len(tb) + 1, cols=len(COLS))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for row in t.rows:
        for j, w in enumerate(WIDTHS):
            row.cells[j].width = w
    for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(WIDTHS[j].twips)))
    trPr = t.rows[0]._tr.get_or_add_trPr()
    h = OxmlElement("w:tblHeader"); h.set(qn("w:val"), "true"); trPr.append(h)
    for j, txt in enumerate(headers):
        write(t.cell(0, j), txt, bold=True, left=(j == 0))
        set_border(t.cell(0, j), top=12, bottom=8)
    for r_i, r in enumerate(tb.to_dict("records"), start=1):
        if r["kind"] == "sub":
            m = t.cell(r_i, 0).merge(t.cell(r_i, len(COLS) - 1))
            write(m, r["Item"], bold=True, italic=True, left=True)
            continue
        bold = r["kind"] in ("composite", "total")
        for j, c in enumerate(COLS):
            write(t.cell(r_i, j), r[c], bold=bold, italic=(r["kind"] == "composite" and j == 0), left=(j == 0))
        if r["kind"] == "total":
            for j in range(len(COLS)):
                set_border(t.cell(r_i, j), top=4, bottom=12)
    n = doc.add_paragraph(); n.paragraph_format.space_before = Pt(4)
    run = n.add_run("Note. "); run.italic = True; run.font.size = Pt(9)
    n.add_run(note).font.size = Pt(9)
doc.save(OUT_DOCX)

# ---------------------------------------------------------------- Excel
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
    for num, title, tb, pct_label, note in tables:
        out = tb.drop(columns="kind").fillna("").rename(columns={"% 4-5": pct_label})
        sheet = num.replace("Table ", "T")
        pd.DataFrame([[f"{num}. {title}"]]).to_excel(xw, sheet_name=sheet, index=False, header=False)
        out.to_excel(xw, sheet_name=sheet, index=False, startrow=2)
        pd.DataFrame([["Note. " + note]]).to_excel(xw, sheet_name=sheet, index=False, header=False,
                                                    startrow=len(out) + 4)
        xw.sheets[sheet].column_dimensions["A"].width = 80

for num, title, tb, pct_label, note in tables:
    tot = tb[tb.kind == "total"].iloc[0]
    print(f"{num}  {title[:70]:70s} overall {tot['Overall M (SD)']}  rural {tot['Rural M (SD)']}  urban {tot['Urban M (SD)']}  p {tot['p']}")
print(f"\nSaved: {OUT_DOCX.name} and {OUT_XLSX.name}")
