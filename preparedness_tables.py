"""
Tables for Section 5.4 (areas requiring additional support) and 5.5 (leadership and preparedness).

  Table 5.4  Priority area for additional support: leaders' view of teachers (G9) and teachers' own (F13)
  Table 5.5  Support that would most strengthen teachers' preparedness (I4): themes, by location
  Table 5.6  Spearman correlations between leadership practices / context and teacher preparedness
  Table 5.7  Teachers' self-rated preparedness by CBE-related training received

Saves Preparedness_Tables.docx (and prints the tables).
Requires: pip install pandas numpy scipy openpyxl python-docx
Run:      python preparedness_tables.py
"""
import re
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

BASE = Path.cwd()
OUT_DOCX = BASE / "Preparedness_Tables.docx"
NUM = {"areas": "Table 5.4", "support": "Table 5.5", "corr": "Table 5.6", "training": "Table 5.7"}
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

def cols(df, pattern):
    return [c for c in df.columns if re.match(pattern, c)]

def to_num(s):
    return pd.to_numeric(s.astype(str).str.extract(r"^\s*(\d)\s*=")[0], errors="coerce")

def scale(df, items):
    return pd.concat([to_num(df[c]) for c in items], axis=1).mean(axis=1)

def fmt_p(p):
    return "< .001" if p < .001 else f"{p:.3f}".lstrip("0")

def npct(n, base):
    return f"{n} ({n / base * 100:.1f})"

# ---------------------------------------------------------------- Table 5.4: priority area (single choice)
SAME_AREA = {"Competency-based assessment": "Assessment",
             "Differentiating learning to address learner needs": "Differentiating learning"}
g9 = L[q(L, "G9.")].replace(SAME_AREA)
f13 = T[q(T, "F13.")].replace(SAME_AREA)
areas = sorted(set(g9.dropna()) | set(f13.dropna()), key=lambda a: -(f13 == a).sum())
tr, tu = T[LOC] == "Rural", T[LOC] == "Urban"
rows = [[a, npct(int((g9 == a).sum()), g9.notna().sum()),
         npct(int((f13[tr] == a).sum()), int(tr.sum())), npct(int((f13[tu] == a).sum()), int(tu.sum())),
         npct(int((f13 == a).sum()), f13.notna().sum())] for a in areas]
rows.append(["Total", npct(int(g9.notna().sum()), g9.notna().sum()), npct(int(tr.sum()), int(tr.sum())),
             npct(int(tu.sum()), int(tu.sum())), npct(int(f13.notna().sum()), f13.notna().sum())])
t54 = pd.DataFrame(rows, columns=["Area", "Leaders", "Teachers rural", "Teachers urban", "Teachers total"])
note54 = (f"Leaders (n = {g9.notna().sum()}) named the most important area in which teachers need additional support "
          f"(G9); teachers (n = {f13.notna().sum()}; rural {int(tr.sum())}, urban {int(tu.sum())}) named the area "
          "in which they themselves need most support (F13). One response per person. “Competency-based "
          "assessment” (leaders) and “Assessment” (teachers) are shown as one row. "
          "– = option not offered in that questionnaire.")
leader_opts = set(g9.dropna()) | {"Classroom management", "Assessment", "Differentiating learning",
                                  "ICT/digital integration", "Teaching and learning approaches",
                                  "Understanding CBE requirements"}
teacher_opts = {"Assessment", "Competency development", "Differentiating learning", "ICT/digital integration",
                "Teaching and learning approaches", "Understanding CBE requirements",
                "Use of learner performance information"} | set(f13.dropna())
for i, a in enumerate(t54["Area"]):
    if a != "Total" and a not in leader_opts:
        t54.loc[i, "Leaders"] = "–"
    if a != "Total" and a not in teacher_opts:
        t54.loc[i, ["Teachers rural", "Teachers urban", "Teachers total"]] = "–"

# ---------------------------------------------------------------- Table 5.5: support themes (I4)
EMPTY = re.compile(r"^\s*(n/?a|na|nan|none\.?|nothing|no|not applicable)\s*$", re.I)
SUPPORT = {   # theme -> keywords (edit to refine)
    "More training, retooling and capacity building":
        r"train|retool|re-tool|seminar|workshops? |workshops$|frequent workshop|regular workshop|capacity|in ?service|"
        r"refresher|course|development|induction|sensiti|awareness|knowledge|know more|guide",
    "ICT equipment and internet connectivity":
        r"ict|digital|computer|internet|wi-?fi|laptop|tablet|gadget|projector|smart ?board|technolog|electronic",
    "Teaching and learning resources / textbooks": r"resource|material|text ?books?|books|resourses",
    "Facilities and infrastructure": r"structure|construct|facilit|equip schools|stage",
    "Government, parent and stakeholder support": r"government|parent|stake ?holder",
    "Teacher collaboration and support": r"collaborat|teacher support|office support",
}
EXAMPLES = {
    "More training, retooling and capacity building": "More training on implementation of CBE",
    "ICT equipment and internet connectivity": "Provision of ICT equipment like laptops and smart boards",
    "Teaching and learning resources / textbooks": "Availability of teaching and learning materials",
    "Facilities and infrastructure": "Building structures for the specific learning areas",
    "Government, parent and stakeholder support": "Actively involving all the stake holders including the parents",
    "Teacher collaboration and support": "Collaboration of teachers to be more often",
}
text = T[q(T, "I4.")].astype(str).str.replace("\n", " ").str.strip()
valid = T[q(T, "I4.")].notna() & ~text.str.match(EMPTY)
hits = pd.DataFrame({k: text.str.contains(p, case=False, regex=True) & valid for k, p in SUPPORT.items()})
vr, vu = int((valid & tr).sum()), int((valid & tu).sum())
rows = [[k, npct(int(hits.loc[tr, k].sum()), vr), npct(int(hits.loc[tu, k].sum()), vu),
         npct(int(hits[k].sum()), int(valid.sum())), f"“{EXAMPLES[k]}”", int(hits[k].sum())]
        for k in SUPPORT]
t55 = (pd.DataFrame(rows, columns=["Support needed", "Rural", "Urban", "Total", "Example response", "_n"])
       .sort_values("_n", ascending=False).drop(columns="_n"))
uncoded = text[valid & ~hits.any(axis=1)].tolist()
note55 = (f"Teachers who answered I4, n = {int(valid.sum())} (rural {vr}, urban {vu}). Responses were coded by "
          "keyword; one response may fall under more than one theme, so percentages do not sum to 100."
          + (f" Not coded ({len(uncoded)}): {'; '.join(uncoded)}." if uncoded else ""))

# ---------------------------------------------------------------- Table 5.6: correlations
T_B, T_C, T_D = scale(T, cols(T, r"B1\.\d+\. ")), scale(T, cols(T, r"C\d\. School")), scale(T, cols(T, r"D([1-9]|10)\. "))
T_E, T_F = scale(T, cols(T, r"E([1-9]|1[0-2])\. ")), scale(T, cols(T, r"F([1-9]|1[0-2])\. "))
L_B, L_C, L_D = scale(L, cols(L, r"B1\.\d+\. ")), scale(L, cols(L, r"C\d+\. I ")), scale(L, cols(L, r"D([1-9]|10)\. "))
L_E, L_F, L_G = scale(L, cols(L, r"E[1-4]\.\d\. ")), scale(L, cols(L, r"F\d\. I ")), scale(L, cols(L, r"G[1-8]\. "))
order_map = {"1": 1, "2-3": 2, "4-5": 3, "More than 5": 4}
pd_n = T[q(T, "A7.")].astype(str).str.replace(r"\.0$", "", regex=True).map(order_map)

def rho(x, y):
    m = x.notna() & y.notna()
    r, p = stats.spearmanr(x[m], y[m])
    stars = "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else ""
    return f"{r:.2f}{stars}".replace("0.", "."), int(m.sum())

rows = []
for label, t_x, l_x in [("Management leadership (B1)", T_B, L_B), ("Instructional leadership (C)", T_C, L_C),
                        ("Adaptation of leadership practices (D)", T_D, L_D),
                        ("Leadership support to teachers (F, leaders)", None, L_F),
                        ("Contextual conditions (E)", T_E, L_E),
                        ("Number of CBE training activities (A7, teachers)", pd_n, None)]:
    tv = rho(t_x, T_F) if t_x is not None else ("–", "")
    lv = rho(l_x, L_G) if l_x is not None else ("–", "")
    rows.append([label, tv[0], lv[0]])
t56 = pd.DataFrame(rows, columns=["Variable", "Teachers", "Leaders"])
note56 = (f"Spearman's ρ. Teachers (n = {len(T)}): correlation with teachers' self-rated preparedness (F1–F12); "
          f"number of training activities n = {int(pd_n.notna().sum())} (teachers with CBE training). Leaders (n = {len(L)}): "
          "correlation with leaders' assessment of teacher preparedness (G1–G8). Both variables in each pair come "
          "from the same respondents. *p < .05, **p < .01, ***p < .001. – = not measured for that group.")

# ---------------------------------------------------------------- Table 5.7: preparedness by training
def msd(x):
    return f"{x.mean():.2f} ({x.std():.2f})"

yn = T[q(T, "A6.")]
rows = [["Received CBE-related training", "", "", ""]]
for v in ["Yes", "No"]:
    x = T_F[yn == v]
    rows.append([f"    {v}", str(len(x)), msd(x), ""])
U = stats.mannwhitneyu(T_F[yn == "Yes"], T_F[yn == "No"])
rows[0][3] = f"U = {U.statistic:.1f}, p = {fmt_p(U.pvalue)}"
rows.append(["CBE training activities in past 3 yearsᵃ", "", "", ""])
k = len(rows) - 1
labels = {1: "1", 2: "2–3", 3: "4–5", 4: "More than 5"}
groups = []
for code_, lab in labels.items():
    x = T_F[pd_n == code_]
    groups.append(x)
    rows.append([f"    {lab}", str(len(x)), msd(x), ""])
H = stats.kruskal(*[g for g in groups if len(g)])
r_dose, p_dose = stats.spearmanr(pd_n[pd_n.notna()], T_F[pd_n.notna()])
rows[k][3] = f"H({sum(len(g) > 0 for g in groups) - 1}) = {H.statistic:.2f}, p = {fmt_p(H.pvalue)}"
t57 = pd.DataFrame(rows, columns=["Training", "n", "Preparedness M (SD)", "Test"])
note57 = ("Preparedness = mean of F1–F12 (1 = strongly disagree, 5 = strongly agree). U = Mann-Whitney test; "
          f"H = Kruskal-Wallis test. Trend across number of activities: Spearman's ρ = {r_dose:.2f}, "
          f"p = {fmt_p(p_dose)}. ᵃ Teachers who received CBE-related training only; groups 4–5 and more than 5 "
          "are small.").replace("= 0.", "= .")

for key, df_, note in [("areas", t54, note54), ("support", t55, note55), ("corr", t56, note56), ("training", t57, note57)]:
    print(NUM[key]); print(df_.to_string(index=False)); print(note, "\n")

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

def add_table(doc, number, title, df, headers, widths, note, left_cols=(0,), italic_cols=(),
              bold_last=False, bold_rows=()):
    p = doc.add_paragraph(); p.add_run(number).bold = True; p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph().add_run(title).italic = True
    t = doc.add_table(rows=len(df) + 1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = w
    for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j].twips)))
    trPr = t.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader"); hdr.set(qn("w:val"), "true"); trPr.append(hdr)
    for row in t.rows:                       # keep rows together so short tables don't split
        for c in row.cells:
            c.paragraphs[0].paragraph_format.keep_with_next = True
    for j, h in enumerate(headers):
        write(t.cell(0, j), h, bold=True, left=(j in left_cols))
        set_border(t.cell(0, j), top=12, bottom=8)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        last = bold_last and i == len(df)
        for j, v in enumerate(row):
            write(t.cell(i, j), str(v).strip() if j == 0 and not str(v).startswith("    ") else
                  ("   " + str(v).strip() if j == 0 else v),
                  bold=last or (i - 1 in bold_rows and j == 0), left=(j in left_cols),
                  italic=(j in italic_cols), size=9 if j in italic_cols else 10)
            if last:
                set_border(t.cell(i, j), top=4)
    for j in range(len(headers)):
        set_border(t.cell(len(df), j), bottom=12)
    n = doc.add_paragraph(); n.paragraph_format.space_before = Pt(4)
    r = n.add_run("Note. "); r.italic = True; r.font.size = Pt(9)
    n.add_run(note).font.size = Pt(9)
    doc.add_paragraph()

doc = Document()
s = doc.sections[0]
s.left_margin = s.right_margin = Inches(0.8)
doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)

add_table(doc, NUM["areas"], "Priority Areas in Which Teachers Require Additional Support for CBE", t54,
          ["Area", "Leaders\nn (%)", "Teachers: rural\nn (%)", "Teachers: urban\nn (%)", "Teachers: total\nn (%)"],
          [Inches(2.5)] + [Inches(1.1)] * 4, note54, bold_last=True)
add_table(doc, NUM["support"], "Support That Would Most Strengthen Teachers' Preparedness to Implement CBE", t55,
          ["Support needed", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)", "Example response"],
          [Inches(2.0), Inches(0.8), Inches(0.8), Inches(0.8), Inches(2.5)], note55, left_cols=(0, 4), italic_cols=(4,))
add_table(doc, NUM["corr"], "Correlations between Leadership Practices, Context and Teacher Preparedness", t56,
          ["Variable", "Teachers: self-rated\npreparedness (ρ)", "Leaders: assessed\npreparedness (ρ)"],
          [Inches(3.3), Inches(1.6), Inches(1.6)], note56)
heading_rows = [i for i, v in enumerate(t57["Training"]) if not v.startswith("    ")]
add_table(doc, NUM["training"], "Teachers' Self-Rated Preparedness by CBE-Related Training Received", t57,
          ["Training", "n", "Preparedness\nM (SD)", "Test"],
          [Inches(2.9), Inches(0.6), Inches(1.3), Inches(1.9)], note57, bold_rows=heading_rows)
doc.save(OUT_DOCX)
print(f"Saved: {OUT_DOCX.name}")
