"""
Tables for Section 4.3 (main challenges) and 4.4 (actions taken and their success).

  Table 4.3  Main challenges in CBE implementation: themes from open-ended answers
             (leaders E5, teachers E13), with an example answer per theme
  Table 4.4  Actions taken by schools to address the main challenge (leaders E7): themes
  Table 4.5  Perceived success of these actions (leaders E8), by location

Saves Challenge_Tables.docx (and prints the tables).
Requires: pip install pandas openpyxl python-docx
Run:      python challenge_tables.py
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
OUT_DOCX = BASE / "Challenge_Tables.docx"
NUM_CHALLENGES, NUM_ACTIONS, NUM_SUCCESS = "Table 4.3", "Table 4.4", "Table 4.5"
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

EMPTY = re.compile(r"^\s*(n/?a|na|nan|none\.?|no|not applicable|no thank.*)\s*$", re.I)

# ---------------------------------------------------------------- codebooks (edit keywords to refine)
CHALLENGES = {
    "Inadequate teaching and learning resources / textbooks": r"resource|material|text ?book|books",
    "Inadequate facilities, laboratories and infrastructure": r"practical|facilit|lab|infrastruct|workshop|room",
    "Inadequate ICT devices and connectivity":
        r"ict|digital|computer|internet|network|gadget|laptop|tablet|phone|electronic|technolog|electricity|online|smart",
    "Inadequate teacher training, retooling and CBE knowledge":
        r"train|retool|re-tool|capacity|induction|know|skill|understanding|technical (advice|guidance)|guidance|information",
    "Financial constraints": r"financ|fund|money|budget|expend",
    "Staffing shortages and workload":
        r"staff|personnel|teachers? (of|with)|human resource|workload|untrained learning|assigned areas|trained teachers",
    "Learner factors (absenteeism, readiness)": r"abse|learners are|prior knowledge|pathway|enrol|population",
    "Curriculum transition and design": r"curriculum|syllab|8.?4.?4|strand|assessment",
    "Limited parent/community support": r"parent|communit",
}
ACTIONS = {
    "Engaging parents, stakeholders and authorities":
        r"parent|stakeholder|moe|lobby|enlist|sensiti|supporting bodies",
    "Reallocating and prioritising school funds": r"fund|budget|levy|votehead|prioriti",
    "Sharing resources and collaborating with other schools/institutions":
        r"other schools|neighbour|junior|institution|consult",
    "Using available or personal resources (e.g. phones)": r"phone|personal|gadget|available resources|existing",
    "Acquiring ICT devices and learning resources": r"provision|ict devices|purchase|buying",
    "Staffing and curriculum adjustments": r"employ|subject combination|curriculum",
    "Little or no action taken": r"nothing",
}
# Hand-picked verbatim examples (leave a theme out to let the script pick one automatically)
EXAMPLES = {
    "Inadequate teaching and learning resources / textbooks": "Learning materials like books from government not adequate",
    "Inadequate ICT devices and connectivity": "Lack of ict and internet connection",
    "Inadequate teacher training, retooling and CBE knowledge":
        "Lack of most basic training for teachers to implement CBE and funds to support the program",
    "Learner factors (absenteeism, readiness)": "Chronic absenteeism",
    "Reallocating and prioritising school funds":
        "Mobilised resources from other votehead to purchase some of the learning resources",
    "Using available or personal resources (e.g. phones)": "Empowering teachers to use their phones while in class",
    "Little or no action taken": "Nothing much",
}

def code(df, col, book):
    text = df[col].astype(str).str.replace("\n", " ").str.strip()
    valid = df[col].notna() & ~text.str.match(EMPTY)
    hits = pd.DataFrame({k: text.str.contains(p, case=False, regex=True) & valid for k, p in book.items()})
    return text, valid, hits

def example(theme, sources):
    if theme in EXAMPLES:
        return f"“{EXAMPLES[theme]}”"
    pool = []
    for text, hits in sources:
        only = hits[theme] & (hits.sum(axis=1) == 1)
        pool += [s for s in text[only] if 15 <= len(s) <= 150]
    if not pool:
        for text, hits in sources:
            pool += [s for s in text[hits[theme]] if 15 <= len(s) <= 150]
    return f"“{sorted(pool, key=len)[0].rstrip('.')}”" if pool else ""

# ---------------------------------------------------------------- Table 4.3
lt, lv, lh = code(L, q(L, "E5."), CHALLENGES)
tt, tv, th = code(T, q(T, "E13."), CHALLENGES)
n_l, n_t = int(lv.sum()), int(tv.sum())
rows = []
for k in CHALLENGES:
    a, b = int(lh[k].sum()), int(th[k].sum())
    if a + b == 0:
        continue
    rows.append([k, f"{a} ({a / n_l * 100:.1f})", f"{b} ({b / n_t * 100:.1f})", example(k, [(lt, lh), (tt, th)]), b])
challenges = (pd.DataFrame(rows, columns=["Challenge", "Leaders", "Teachers", "Example response", "_n"])
              .sort_values("_n", ascending=False).drop(columns="_n"))
note_ch = (f"Leaders: n = {n_l} (single most significant contextual challenge, E5); teachers: n = {n_t} "
           "(most significant challenge affecting their ability to implement CBE, E13). Responses were coded by "
           "keyword; one response may fall under more than one theme, so percentages do not sum to 100.")

# ---------------------------------------------------------------- Table 4.4
at, av, ah = code(L, q(L, "E7."), ACTIONS)
n_a = int(av.sum())
n_r, n_u = int((L.loc[av, LOC] == "Rural").sum()), int((L.loc[av, LOC] == "Urban").sum())
rows = []
for k in ACTIONS:
    r, u = int(ah.loc[L[LOC] == "Rural", k].sum()), int(ah.loc[L[LOC] == "Urban", k].sum())
    if r + u == 0:
        continue
    rows.append([k, f"{r} ({r / n_r * 100:.1f})", f"{u} ({u / n_u * 100:.1f})",
                 f"{r + u} ({(r + u) / n_a * 100:.1f})", example(k, [(at, ah)]), r + u])
actions = (pd.DataFrame(rows, columns=["Action", "Rural", "Urban", "Total", "Example response", "_n"])
           .sort_values("_n", ascending=False).drop(columns="_n"))
uncoded = at[av & ~ah.any(axis=1)].tolist()
note_ac = (f"School leaders, n = {n_a} (rural n = {n_r}, urban n = {n_u}). Responses to E7 (action taken to address "
           "the main challenge) were coded by keyword; one response may fall under more than one theme."
           + (f" Not coded: {'; '.join(uncoded)}." if uncoded else ""))

# ---------------------------------------------------------------- Table 4.5
e8 = q(L, "E8.")
order = ["Very successful", "Moderately successful", "Slightly successful", "Not successful"]
ct = pd.crosstab(L[e8], L[LOC])
ct = ct.reindex([o for o in order if o in ct.index] + [o for o in ct.index if o not in order])
ct["Total"] = ct.sum(axis=1)
tot = ct.sum()
success = pd.DataFrame({"Level of success": ct.index,
                        **{c: [f"{int(v)} ({v / tot[c] * 100:.1f})" for v in ct[c]] for c in ["Rural", "Urban", "Total"]}})
success.loc[len(success)] = ["Total"] + [f"{int(tot[c])} (100.0)" for c in ["Rural", "Urban", "Total"]]
note_su = f"School leaders, N = {len(L)}. Percentages are column percentages."

for num, df_, note in [(NUM_CHALLENGES, challenges, note_ch), (NUM_ACTIONS, actions, note_ac), (NUM_SUCCESS, success, note_su)]:
    print(num); print(df_.to_string(index=False)); print(note, "\n")

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

def add_table(doc, number, title, df, headers, widths, note, left_cols=(0,), italic_cols=(), bold_last=False):
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
        last = bold_last and i == len(df)
        for j, v in enumerate(row):
            write(t.cell(i, j), v, bold=last, left=(j in left_cols), italic=(j in italic_cols),
                  size=9 if j in italic_cols else 10)
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

add_table(doc, NUM_CHALLENGES, "Main Challenges in CBE Implementation Reported by School Leaders and Teachers",
          challenges, ["Challenge", "Leaders\nn (%)", "Teachers\nn (%)", "Example response"],
          [Inches(2.4), Inches(0.85), Inches(0.85), Inches(2.8)], note_ch, left_cols=(0, 3), italic_cols=(3,))
add_table(doc, NUM_ACTIONS, "Actions Taken by Schools to Address the Main Challenge (School Leaders)",
          actions, ["Action", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)", "Example response"],
          [Inches(2.1), Inches(0.8), Inches(0.8), Inches(0.8), Inches(2.4)], note_ac, left_cols=(0, 4), italic_cols=(4,))
add_table(doc, NUM_SUCCESS, "Perceived Success of Actions Taken to Address the Main Challenge",
          success, ["Level of success", "Rural\nn (%)", "Urban\nn (%)", "Total\nn (%)"],
          [Inches(2.6), Inches(1.2), Inches(1.2), Inches(1.2)], note_su, bold_last=True)
doc.save(OUT_DOCX)
print(f"Saved: {OUT_DOCX.name}")
