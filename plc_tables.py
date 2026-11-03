"""
Tables for Section 6.0 Professional Learning and PLC Readiness.

  Table 6.1  Existing professional learning opportunities: school leaders (H1-H3), by location
  Table 6.2  Existing professional learning opportunities: teachers (G1-G4), by location
  Table 6.3  Presence and perceived usefulness of a structured PLC (teachers H1; leaders H4; teachers H2)
  Table 6.4  Priority areas for PLC support (multiple response: leaders H5, teachers H3)
  Table 6.5  What would make a PLC useful (open-ended: leaders H7, teachers H4), themes with examples

Saves PLC_Tables.docx (and prints the tables).
Requires: pip install pandas openpyxl python-docx
Run:      python plc_tables.py
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
OUT_DOCX = BASE / "PLC_Tables.docx"
NUM = ["Table 6.1", "Table 6.2", "Table 6.3", "Table 6.4", "Table 6.5"]
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

def npct(n, base):
    return f"{n} ({n / base * 100:.1f})" if base else "0 (0.0)"

FREQ_MEET = ["More than once per month", "Monthly", "More than once per term", "Once per term",
             "Less than once a term", "Never"]
OFTEN = ["Very Often", "Often", "Sometimes", "Rarely", "Never"]
EXTENT = ["To a very great extent", "To a great extent", "To a moderate extent", "To a small extent", "Not at all"]
USEFUL = ["Extremely useful", "Very useful", "Moderately useful", "Slightly useful", "Not useful"]

# ---------------------------------------------------------------- panel tables (by location)
def panel_rows(spec):
    """spec: list of (heading, dataframe, column, category order). Returns rows with Rural/Urban/Total n (%)."""
    rows = []
    for heading, df, col, order in spec:
        d = df[[col, LOC]].dropna().copy()
        d[col] = d[col].astype(str).str.replace(r"^\d\s*=\s*", "", regex=True)      # "4 = To a great extent"
        rows.append(["var", heading, "", "", ""])
        ct = pd.crosstab(d[col], d[LOC]).reindex(columns=["Rural", "Urban"], fill_value=0)
        cats = [c for c in order if c in ct.index] + [c for c in ct.index if c not in order]
        ct = ct.reindex(cats, fill_value=0)
        base = ct.sum()
        for cat, r in ct.iterrows():
            rows.append(["cat", cat, npct(int(r["Rural"]), base["Rural"]), npct(int(r["Urban"]), base["Urban"]),
                         npct(int(r.sum()), int(base.sum()))])
    return pd.DataFrame(rows, columns=["kind", "Item", "Rural", "Urban", "Total"])

def loc_counts(df):
    return int((df[LOC] == "Rural").sum()), int((df[LOC] == "Urban").sum())

t61 = panel_rows([
    ("Opportunities to meet and discuss professional challenges (H1)", L, q(L, "H1."), ["Yes", "No"]),
    ("Frequency of professional discussions (H2)", L, q(L, "H2."), FREQ_MEET),
    ("Extent existing professional learning supports leaders with CBE challenges (H3)", L, q(L, "H3."), EXTENT),
])
lr, lu = loc_counts(L)
note61 = f"School leaders, N = {len(L)} (rural n = {lr}, urban n = {lu}). Column percentages."

t62 = panel_rows([
    ("Participation in professional discussions about CBE (G1)", T, q(T, "G1."), FREQ_MEET),
    ("Teachers share strategies for addressing CBE challenges (G2)", T, q(T, "G2."), OFTEN),
    ("School leaders create opportunities for teachers to learn from one another (G3)", T, q(T, "G3."), OFTEN),
    ("Extent professional learning has strengthened ability to implement CBE (G4)", T, q(T, "G4."), EXTENT),
])
tr, tu = loc_counts(T)
note62 = f"Teachers, N = {len(T)} (rural n = {tr}, urban n = {tu}). Column percentages."

t63 = panel_rows([
    ("School currently has a structured PLC (teachers, H1)", T, q(T, "H1."), ["Yes", "No", "Not sure"]),
    ("Usefulness of a structured PLC for strengthening leadership (leaders, H4)", L, q(L, "H4."), USEFUL),
    ("Usefulness of a structured PLC for addressing CBE challenges (teachers, H2)", T, q(T, "H2."), USEFUL),
])
note63 = (f"Teachers N = {len(T)} (rural {tr}, urban {tu}); school leaders N = {len(L)} (rural {lr}, urban {lu}). "
          "Column percentages within each question.")

# ---------------------------------------------------------------- Table 6.4: multiple response
SAME = {"Teacher collaboration": "Collaboration among teachers"}

def multi(df, prefix):
    opts = [c for c in df.columns if c.startswith(prefix + "/")]
    n = int(df[opts[0]].notna().sum())
    out = {}
    for c in opts:
        name = SAME.get(c.split("/", 1)[1], c.split("/", 1)[1])
        out[name] = int(pd.to_numeric(df[c], errors="coerce").fillna(0).sum())
    return out, n

lm, ln = multi(L, "H5. Which areas would you most want a PLC to support?")
tm, tn = multi(T, "H3. Which areas would you most want a PLC to support?")
areas = sorted((set(lm) | set(tm)) - {"Other"}, key=lambda a: -(lm.get(a, 0) + tm.get(a, 0)))
t64 = pd.DataFrame([["cat", a, npct(lm[a], ln) if a in lm else "–", npct(tm[a], tn) if a in tm else "–"]
                    for a in areas], columns=["kind", "Area", "Leaders", "Teachers"])
note64 = (f"School leaders n = {ln}; teachers n = {tn}. Multiple responses allowed; percentages are of respondents "
          "and do not sum to 100. – = option not offered in that questionnaire. No respondent chose “Other”.")

# ---------------------------------------------------------------- Table 6.5: open-ended themes
EMPTY = re.compile(r"^\s*(n/?a|na|nan|none\.?|nothing|no|yes|not sure.*|not applicable|everything about it\.?|"
                   r"if it'?s realized|availing them)\s*$", re.I)
USEFUL_THEMES = {   # theme -> keywords (edit to refine)
    "Training, capacity building and knowledge/skills":
        r"train|capacity|workshop|refresher|induction|development|equip|skill|competence|knowledge|know ?how|"
        r"understand|sensiti|seminar",
    "Collaboration and sharing of ideas and experiences":
        r"collab|shar|interaction|unity|consult|collectively|discussion|ideas",
    "Regular, scheduled meetings with adequate time":
        r"regular|often|frequent|contin|constant|time ?table|meeting",
    "Expert facilitation and professional guidance":
        r"expert|facilitator|proper knowledge|guidance|supervis",
    "Solving implementation challenges and improving practice":
        r"challenge|problem|improv|content delivery|easier to teach|practicab|teaching and learning practices",
    "CBE assessment": r"assess",
    "Resources and ICT": r"resource|ict|digital|wi-?fi",
    "Involvement of stakeholders and the community":
        r"stake ?holder|communit|involves all|involvement of many|parent",
}
EXAMPLES = {
    "Training, capacity building and knowledge/skills": "Capacity building of teachers for proper CBE implementation",
    "Collaboration and sharing of ideas and experiences": "Sharing with other teachers from across schools",
    "Regular, scheduled meetings with adequate time": "Meeting should be held often to discuss the challenges",
    "Expert facilitation and professional guidance": "Trained facilitators",
    "Solving implementation challenges and improving practice": "Help improve teaching and learning practices",
    "CBE assessment": "Better assessment approaches",
    "Resources and ICT": "provision of required resources",
    "Involvement of stakeholders and the community": "Stakeholders engagement",
}

def code(df, col):
    text = df[col].astype(str).str.replace("\n", " ").str.strip()
    valid = df[col].notna() & ~text.str.match(EMPTY)
    hits = pd.DataFrame({k: text.str.contains(p, case=False, regex=True) & valid for k, p in USEFUL_THEMES.items()})
    return text, valid, hits

lt, lv, lh = code(L, q(L, "H7."))
tt, tv, th = code(T, q(T, "H4."))
nl, nt = int(lv.sum()), int(tv.sum())
rows = [["cat", k, npct(int(lh[k].sum()), nl), npct(int(th[k].sum()), nt), f"“{EXAMPLES[k]}”",
         int(lh[k].sum() + th[k].sum())] for k in USEFUL_THEMES]
t65 = (pd.DataFrame(rows, columns=["kind", "Theme", "Leaders", "Teachers", "Example response", "_n"])
       .sort_values("_n", ascending=False).drop(columns="_n"))
unc_l = lt[lv & ~lh.any(axis=1)].tolist()
unc_t = tt[tv & ~th.any(axis=1)].tolist()
note65 = (f"Leaders: type of support that would make a PLC most useful (H7), n = {nl}; teachers: what would make "
          f"participation in a PLC meaningful (H4), n = {nt} substantive answers (answers such as “N/a” or "
          "“Not sure” excluded). Coded by keyword; one response may fall under more than one theme, so "
          f"percentages do not sum to 100. Responses not fitting any theme: leaders {len(unc_l)}, teachers {len(unc_t)}.")

TABLES = [
    (NUM[0], "Existing Professional Learning Opportunities Reported by School Leaders", t61,
     ["Item", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)"], [3.6, 1.0, 1.0, 1.0], note61),
    (NUM[1], "Existing Professional Learning Opportunities Reported by Teachers", t62,
     ["Item", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)"], [3.6, 1.0, 1.0, 1.0], note62),
    (NUM[2], "Presence and Perceived Usefulness of a Structured Professional Learning Community", t63,
     ["Item", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)"], [3.6, 1.0, 1.0, 1.0], note63),
    (NUM[3], "Priority Areas for PLC Support", t64,
     ["Area", "Leaders\nn (%)", "Teachers\nn (%)"], [3.6, 1.2, 1.2], note64),
    (NUM[4], "What Would Make a PLC Useful: Themes from Open-Ended Responses", t65,
     ["Theme", "Leaders\nn (%)", "Teachers\nn (%)", "Example response"], [2.3, 0.85, 0.85, 2.7], note65),
]
for num, title, df_, *_ , note in TABLES:
    print(num, title); print(df_.drop(columns="kind").to_string(index=False)); print("Note.", note, "\n")
if unc_l or unc_t:
    print("Not coded - leaders:", unc_l, "\nNot coded - teachers:", unc_t)

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

for k, (num, title, df_, headers, widths_in, note) in enumerate(TABLES):
    if k:
        doc.add_page_break()
    p = doc.add_paragraph(); p.add_run(num).bold = True; p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph().add_run(title).italic = True
    widths = [Inches(w) for w in widths_in]
    t = doc.add_table(rows=len(df_) + 1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = w
    for j, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j].twips)))
    trPr = t.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader"); hdr.set(qn("w:val"), "true"); trPr.append(hdr)
    for j, h in enumerate(headers):
        write(t.cell(0, j), h, bold=True, left=(j == 0 or h == "Example response"))
        set_border(t.cell(0, j), top=12, bottom=8)
    for i, r in enumerate(df_.to_dict("records"), start=1):
        vals = [v for key, v in r.items() if key != "kind"]
        if r["kind"] == "var":
            m = t.cell(i, 0).merge(t.cell(i, len(headers) - 1))
            write(m, vals[0], bold=True, left=True)
            continue
        indent = "   " if any(x["kind"] == "var" for x in df_.to_dict("records")) else ""
        for j, v in enumerate(vals):
            ex = headers[j] == "Example response"
            write(t.cell(i, j), (indent + v) if j == 0 else v, left=(j == 0 or ex), italic=ex, size=9 if ex else 10)
    for j in range(len(headers)):
        set_border(t.cell(len(df_), j), bottom=12)
    n = doc.add_paragraph(); n.paragraph_format.space_before = Pt(4)
    run = n.add_run("Note. "); run.italic = True; run.font.size = Pt(9)
    n.add_run(note).font.size = Pt(9)
doc.save(OUT_DOCX)
print(f"Saved: {OUT_DOCX.name}")
