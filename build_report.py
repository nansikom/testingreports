"""
EOY Nomination Report generator
--------------------------------
Reads EOY_Report_Data.xlsx (Summary, History, Departments, Locations, Nominations sheets)
and renders a modern, printable executive PDF.

Usage:  python build_report.py [input.xlsx] [output.pdf]
Requires: reportlab, openpyxl, matplotlib
"""
import sys, io, re
from collections import OrderedDict
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, Image, KeepTogether, NextPageTemplate, Flowable)

XLSX = sys.argv[1] if len(sys.argv) > 1 else "EOY_Report_Data.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "EOY_Nomination_Report_Executive.pdf"

# ---------- palette ----------
ORANGE = colors.HexColor("#F26522")
ORANGE_LT = colors.HexColor("#FDE9DD")
INK = colors.HexColor("#1F2933")
SLATE = colors.HexColor("#52606D")
MIST = colors.HexColor("#9AA5B1")
LINE = colors.HexColor("#E4E7EB")
PANEL = colors.HexColor("#F5F7FA")
BLUE = colors.HexColor("#1F5F8B")
BLUE_LT = colors.HexColor("#E3EEF6")
GREEN = colors.HexColor("#2E7D5B")
GREEN_LT = colors.HexColor("#E2F2EA")
WHITE = colors.white

W, H = letter
M = 0.7 * inch

# ---------- data ----------
wb = openpyxl.load_workbook(XLSX, data_only=True)

def _key(s): return re.sub(r"[^a-z0-9]", "", str(s).lower())
SHEETS = {_key(n): n for n in wb.sheetnames}            # forgiving lookup: 'History ', 'history', 'HISTORY' all work

def _header_of(ws):
    for r in ws.iter_rows(min_row=1, max_row=15, values_only=True):
        if sum(1 for v in r if v is not None and str(v).strip()) >= 2:
            return [_key(v) for v in r if v is not None]
    return []

def find_sheet(name, must_have=None):
    """Return the real sheet name for a logical name, or None. If must_have columns are given and no sheet
    matches by name, pick the first sheet whose header contains all of them (auto-detects a raw export)."""
    if _key(name) in SHEETS: return SHEETS[_key(name)]
    if must_have:
        for n in wb.sheetnames:
            hdr = _header_of(wb[n])
            if all(any(_key(c) in h for h in hdr) for c in must_have): return n
    return None

def has(name): return find_sheet(name) is not None

def rows(name, must_have=None):
    real = find_sheet(name, must_have)
    if real is None: return [], []
    ws = wb[real]
    data = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    # header = first row with at least 2 filled cells (skips instruction/note rows at the top)
    start = next((i for i, r in enumerate(data) if sum(1 for v in r if v is not None and str(v).strip()) >= 2), None)
    if start is None: return [], []
    return data[start], [r for r in data[start + 1:] if any(v is not None and str(v).strip() for v in r)]

def norm(s): return re.sub(r"[^a-z0-9]", "", str(s or "").lower())
def txt(v): return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""

hh, HIST = rows("History")
HIST = [r for r in HIST if r and r[0] is not None and str(r[0]).strip().isdigit()]

export_sheet = find_sheet("Export", must_have=["recno", "achievement"])
if export_sheet is not None:
    # ---- NEW FORMAT: raw SharePoint export (one row per employee) + optional Awards / Settings / History ----
    _, srows = rows("Settings")
    S = {r[0]: r[1] for r in srows if r[0] is not None and r[1] is not None and str(r[1]).strip() != ""}
    import datetime as _dt
    S.setdefault("Award Year", _dt.date.today().year - (1 if _dt.date.today().month <= 3 else 0))
    S.setdefault("Report Date", _dt.date.today().strftime("%B %d, %Y"))
    S.setdefault("Prepared For", "Executive Review Committee"); S.setdefault("Prepared By", "Achievement Awards Program")
    S.setdefault("Confidentiality Line", "Confidential - For executive review only")
    eh, EXP = rows(export_sheet)
    col = {norm(h): i for i, h in enumerate(eh)}
    def pick(*names):
        for n in names:
            if norm(n) in col: return col[norm(n)]
        return None
    C = dict(rec=pick("Rec. No.", "RecNo", "Record Number", "ID"), ach=pick("Achievement"), ins=pick("Inscription"),
             dept=pick("Department"), es=pick("E/S", "ES", "Status", "Employee Status"),
             loc=pick("WorkLocation", "Work Location", "Work city & state", "Work City State"), title=pick("Title"))
    missing = [k for k in ("rec", "ach", "es") if C[k] is None]
    if missing: raise RuntimeError(f"Export sheet is missing required column(s): {missing}. Found: {eh}")

    # Awards sheet: Rec. No. | EOY (Y/blank) | Nominator | Department override | Notes
    aw = {}
    if has("Awards"):
        ah, AW = rows("Awards"); ai = {norm(h): i for i, h in enumerate(ah)}
        for r in AW:
            key = txt(r[ai.get("recno", 0)])
            if key: aw[key] = {k: txt(r[ai[k]]) if k in ai and ai[k] < len(r) else "" for k in ("eoy", "nominator", "department", "notes")}

    from collections import Counter, defaultdict
    grp = OrderedDict()
    for r in EXP:
        rec = txt(r[C["rec"]])
        if not rec: continue
        g = grp.setdefault(rec, dict(rows=[], ach=Counter(), ins=Counter(), dept=Counter(), loc=Counter()))
        g["rows"].append(r)
        if C["ach"] is not None and txt(r[C["ach"]]): g["ach"][txt(r[C["ach"]])] += 1
        if C["ins"] is not None and txt(r[C["ins"]]): g["ins"][txt(r[C["ins"]])] += 1
        if C["dept"] is not None and txt(r[C["dept"]]): g["dept"][txt(r[C["dept"]])] += 1
        if C["loc"] is not None and txt(r[C["loc"]]): g["loc"][txt(r[C["loc"]])] += 1

    def is_exempt(v): return norm(v) in ("e", "exempt", "ex", "salaried")
    def is_sched(v): return norm(v) in ("s", "scheduled", "sch", "craft")
    def top(c): return c.most_common(1)[0][0] if c else ""

    NOMS = []
    for rec, g in grp.items():
        a = aw.get(rec, {})
        ex = sum(1 for r in g["rows"] if is_exempt(r[C["es"]])); sc = sum(1 for r in g["rows"] if is_sched(r[C["es"]]))
        NOMS.append(dict(Section="EOY" if norm(a.get("eoy", "")) in ("y", "yes", "eoy", "x", "true", "1") else "AA",
                         Department=a.get("department") or top(g["dept"]) or "Unassigned",
                         Inscription=top(g["ins"]) or (txt(g["rows"][0][C["title"]]) if C["title"] is not None else "") or f"Award {rec}",
                         Achievement=top(g["ach"]), Nominator=a.get("nominator", ""), Exempt=ex, Scheduled=sc,
                         Notes=a.get("notes", ""), Rec=rec))

    # Derived summary figures
    all_rows = [r for g in grp.values() for r in g["rows"]]
    S.setdefault("Total Achievement Awards", len(grp))
    S.setdefault("Total Employees Recognized", len(all_rows))
    S.setdefault("Exempt Employees", sum(1 for r in all_rows if is_exempt(r[C["es"]])))
    S.setdefault("Scheduled Employees", sum(1 for r in all_rows if is_sched(r[C["es"]])))
    eoy = [n for n in NOMS if n["Section"] == "EOY"]
    S.setdefault("EOY Awards", len(eoy))
    S.setdefault("EOY AA Honorees", sum(n["Exempt"] + n["Scheduled"] for n in eoy))
    S.setdefault("EOY AA Honorees Exempt", sum(n["Exempt"] for n in eoy))
    S.setdefault("EOY AA Honorees Scheduled", sum(n["Scheduled"] for n in eoy))
    S.setdefault("BoB (Exempt)", 0); S.setdefault("SEOY (Scheduled)", 0)
    S.setdefault("Total EOY Honorees", S["EOY AA Honorees"] + int(S["BoB (Exempt)"] or 0) + int(S["SEOY (Scheduled)"] or 0))
    for k in ("Report Title", "Prepared For", "Prepared By", "Report Date", "Confidentiality Line", "Rooms Reserved"):
        S.setdefault(k, "")

    # Departments table derived from awards
    dd = defaultdict(lambda: [0, 0, 0, 0, 0])
    for n in NOMS:
        d = dd[n["Department"]]; d[0] += 1; d[1] += n["Scheduled"]; d[2] += n["Exempt"]
        if n["Section"] == "EOY": d[3] += 1; d[4] += n["Exempt"] + n["Scheduled"]
    DEPTS = [[k] + v for k, v in sorted(dd.items(), key=lambda kv: kv[0].lower())]

    # Locations share (%) derived from employees
    lc = Counter(txt(r[C["loc"]]) for r in all_rows if C["loc"] is not None and txt(r[C["loc"]]))
    tot_l = sum(lc.values()) or 1
    LOCS = [[k, round(v * 100 / tot_l)] for k, v in lc.most_common(9)]
    LOCS.append(["Other", max(0, 100 - sum(x[1] for x in LOCS))])

    # Current year row for History (computed) unless user typed one
    yr = S.get("Award Year")
    if yr is not None and not any(str(r[0]) == str(yr) for r in HIST):
        HIST.insert(0, [yr, S["Total Achievement Awards"], S["EOY Awards"], S["EOY AA Honorees"], S["EOY AA Honorees Scheduled"],
                        S["EOY AA Honorees Exempt"], S["BoB (Exempt)"], S["SEOY (Scheduled)"], S["Total EOY Honorees"]])
elif not has("Summary"):
    raise RuntimeError("No sheet with the SharePoint export was found. The workbook needs a sheet (any name) whose header row "
                     f"includes 'Rec. No.' and 'Achievement'. Sheets found: {wb.sheetnames}")
else:
    # ---- ORIGINAL FORMAT: pre-aggregated sheets ----
    _, srows = rows("Summary")
    S = {r[0]: r[1] for r in srows}
    dh, DEPTS = rows("Departments")
    lh, LOCS = rows("Locations")
    nh, NOMS = rows("Nominations")
    NOMS = [dict(zip(nh, r)) for r in NOMS]
    for n in NOMS:
        n["Exempt"] = int(n["Exempt"] or 0); n["Scheduled"] = int(n["Scheduled"] or 0)
        for k in ("Inscription", "Achievement", "Nominator", "Notes", "Department", "Section"):
            n[k] = (n[k] or "").strip() if isinstance(n[k], str) else (n[k] or "")
YEAR = S["Award Year"]
# Keep the most recent 8 years so the trend chart and history table always fit on one page
HIST = sorted(HIST, key=lambda r: int(r[0]), reverse=True)[:8]

# ---------- styles ----------
def ps(name, **kw):
    base = dict(fontName="Helvetica", fontSize=10, leading=14, textColor=INK)
    base.update(kw); return ParagraphStyle(name, **base)
st = dict(
    h1=ps("h1", fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=INK),
    h2=ps("h2", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=INK, spaceAfter=4),
    eyebrow=ps("eyebrow", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=ORANGE),
    body=ps("body", fontSize=9.5, leading=13.5, textColor=SLATE),
    small=ps("small", fontSize=8, leading=10.5, textColor=MIST),
    card_title=ps("ct", fontName="Helvetica-Bold", fontSize=10.5, leading=13.5, textColor=INK),
    card_body=ps("cb", fontSize=8.8, leading=12, textColor=SLATE),
    card_meta=ps("cm", fontSize=8, leading=10.5, textColor=MIST),
    kpi_num=ps("kn", fontName="Helvetica-Bold", fontSize=26, leading=28, textColor=INK, alignment=TA_CENTER),
    kpi_lbl=ps("kl", fontSize=8.5, leading=11, textColor=SLATE, alignment=TA_CENTER),
    th=ps("th", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=WHITE),
    td=ps("td", fontSize=9, leading=12, textColor=INK),
    tdr=ps("tdr", fontSize=9, leading=12, textColor=INK, alignment=TA_CENTER),
    tdb=ps("tdb", fontName="Helvetica-Bold", fontSize=9, leading=12, textColor=INK, alignment=TA_CENTER),
)

def esc(s):
    s = str(s or "")
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s

def fmt_nominator(s):
    """'Last, First,Title,Dept' -> 'Last, First · Title · Dept'"""
    parts = [p.strip() for p in re.split(r",(?=\S)", s) if p.strip()]
    if len(parts) >= 2 and len(parts[0].split()) <= 2 and parts[1] and parts[1][0].isupper() and len(parts[1].split()) <= 2 and parts[1].replace("(","").replace(")","").replace(".","").isalpha():
        name = parts[0] + ", " + parts[1]; rest = parts[2:]
    else:
        name = parts[0] if parts else s; rest = parts[1:]
    rest = [re.sub(r"\s+", " ", p) for p in rest]
    rest = [(p.replace(" AL", "al").title() if p.isupper() else p) for p in rest]
    return name, " · ".join(rest)

# ---------- charts ----------
def fig_to_img(fig, width):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", transparent=True); plt.close(fig)
    buf.seek(0); img = Image(buf); ratio = img.imageHeight / img.imageWidth
    img.drawWidth = width; img.drawHeight = width * ratio; return img

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8, "axes.edgecolor": "#E4E7EB",
                     "axes.labelcolor": "#52606D", "xtick.color": "#52606D", "ytick.color": "#52606D"})

def trend_chart():
    yrs = [r[0] for r in HIST][::-1]; aa = [r[1] for r in HIST][::-1]; eoy = [r[2] for r in HIST][::-1]
    fig, ax = plt.subplots(figsize=(7.0, 1.3))
    x = range(len(yrs))
    ax.bar([i - 0.2 for i in x], aa, 0.4, color="#F26522", label="Achievement Awards")
    ax.bar([i + 0.2 for i in x], eoy, 0.4, color="#1F5F8B", label="EOY Awards")
    for i, v in enumerate(aa): ax.text(i - 0.2, v + 2, str(v), ha="center", fontsize=7, color="#1F2933")
    for i, v in enumerate(eoy): ax.text(i + 0.2, v + 2, str(v), ha="center", fontsize=7, color="#1F2933")
    ax.set_xticks(list(x)); ax.set_xticklabels(yrs); ax.set_yticks([])
    for s in ("top", "right", "left"): ax.spines[s].set_visible(False)
    ax.legend(frameon=False, loc="upper left", ncol=2, fontsize=7.5)
    ax.set_ylim(0, max(aa) * 1.25)
    return fig

def honoree_trend():
    yrs = [r[0] for r in HIST][::-1]; tot = [r[8] for r in HIST][::-1]
    fig, ax = plt.subplots(figsize=(3.4, 2.2))
    ax.plot(yrs, tot, color="#F26522", lw=2, marker="o", ms=4)
    for x_, v in zip(yrs, tot): ax.text(x_, v + 14, str(v), ha="center", fontsize=6.5, color="#1F2933")
    ax.set_yticks([]); ax.set_ylim(0, max(tot) * 1.25)
    for s in ("top", "right", "left"): ax.spines[s].set_visible(False)
    ax.tick_params(axis="x", labelsize=7)
    return fig

def dept_bar():
    d = sorted(DEPTS, key=lambda r: r[1], reverse=True)
    names = [r[0] for r in d]; aa = [r[1] for r in d]; eoy = [r[4] or 0 for r in d]
    fig, ax = plt.subplots(figsize=(3.6, 3.2))
    y = range(len(names))
    ax.barh(y, aa, color="#FDE9DD", height=0.62)
    ax.barh(y, eoy, color="#F26522", height=0.62)
    for i, (a, e) in enumerate(zip(aa, eoy)):
        ax.text(a + 0.4, i, f"{a}", va="center", fontsize=8, color="#1F2933")
    ax.set_yticks(list(y)); ax.set_yticklabels(names, fontsize=8); ax.invert_yaxis(); ax.set_xticks([])
    for s in ("top", "right", "bottom"): ax.spines[s].set_visible(False)
    return fig

def donut(vals, labels, cols, center):
    fig, ax = plt.subplots(figsize=(2.2, 2.2))
    if sum(v or 0 for v in vals) <= 0:      # nothing to show yet - draw an empty grey ring
        vals, cols = [1], ["#E4E7EB"]
    ax.pie([v or 0 for v in vals], colors=cols, startangle=90, counterclock=False, wedgeprops=dict(width=0.32, edgecolor="white"))
    ax.text(0, 0.05, center[0], ha="center", va="center", fontsize=13, fontweight="bold", color="#1F2933")
    ax.text(0, -0.22, center[1], ha="center", va="center", fontsize=6.5, color="#52606D")
    return fig

def location_bar():
    d = [r for r in LOCS if r[0] != "Other"]
    fig, ax = plt.subplots(figsize=(3.3, 2.2))
    y = range(len(d))
    ax.barh(y, [r[1] for r in d], color="#1F5F8B", height=0.6)
    for i, r in enumerate(d): ax.text(r[1] + 0.5, i, f"{r[1]}%", va="center", fontsize=7, color="#1F2933")
    ax.set_yticks(list(y)); ax.set_yticklabels([r[0] for r in d], fontsize=7); ax.invert_yaxis(); ax.set_xticks([])
    for s in ("top", "right", "bottom"): ax.spines[s].set_visible(False)
    return fig

# ---------- flowable helpers ----------
def kpi(num, label, accent=ORANGE):
    t = Table([[Paragraph(f"{num:,}" if isinstance(num, int) else str(num), st["kpi_num"])],
               [Paragraph(label, st["kpi_lbl"])]], colWidths=[1.55 * inch])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), PANEL), ("LINEABOVE", (0, 0), (-1, 0), 3, accent),
                           ("TOPPADDING", (0, 0), (-1, 0), 10), ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
                           ("ROUNDEDCORNERS", [6, 6, 6, 6])]))
    return t

def kpi_row(items):
    t = Table([[kpi(*i) for i in items]], colWidths=[(W - 2 * M) / len(items)] * len(items))
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]))
    return t

def section_header(eyebrow, title, sub=None):
    out = [Paragraph(eyebrow.upper(), st["eyebrow"]), Paragraph(title, st["h1"])]
    if sub: out.append(Paragraph(sub, st["body"]))
    out.append(Spacer(1, 4)); out.append(Rule()); out.append(Spacer(1, 10))
    return out

class Rule(Flowable):
    def __init__(self, color=ORANGE, thickness=2, width=1.2 * inch):
        super().__init__(); self.color = color; self.t = thickness; self.w = width; self.height = thickness
    def wrap(self, aw, ah): return (aw, self.t)
    def draw(self):
        self.canv.setStrokeColor(self.color); self.canv.setLineWidth(self.t); self.canv.line(0, 0, self.w, 0)

def styled_table(headers, data, widths, align_from=1, bold_last=False, zebra=True):
    hdr = [Paragraph(h, st["th"]) for h in headers]
    body = []
    for r in data:
        row = [Paragraph(esc(r[0]), st["td"])] + [Paragraph("—" if (v is None or v == 0 or v == "") else esc(v), st["tdr"]) for v in r[1:]]
        body.append(row)
    if bold_last:
        body[-1] = [Paragraph(f"<b>{esc(data[-1][0])}</b>", st["td"])] + [Paragraph(f"<b>{esc(v)}</b>", st["tdb"]) for v in data[-1][1:]]
    t = Table([hdr] + body, colWidths=widths, repeatRows=1)
    sty = [("BACKGROUND", (0, 0), (-1, 0), INK), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
           ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
           ("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE)]
    if zebra:
        for i in range(1, len(body) + 1):
            if i % 2 == 0: sty.append(("BACKGROUND", (0, i), (-1, i), PANEL))
    if bold_last:
        sty += [("BACKGROUND", (0, len(body)), (-1, len(body)), ORANGE_LT), ("LINEABOVE", (0, len(body)), (-1, len(body)), 1, ORANGE)]
    t.setStyle(TableStyle(sty)); return t

def chip(text, bg, fg):
    p = Paragraph(f"<font color='{fg.hexval()}'><b>{esc(text)}</b></font>", ps("chip", fontSize=7.5, leading=9, alignment=TA_CENTER))
    t = Table([[p]], colWidths=[None]); t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), bg), ("TOPPADDING", (0, 0), (-1, -1), 2),
               ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
               ("ROUNDEDCORNERS", [4, 4, 4, 4])])); return t

def nomination_card(n, idx, compact=False):
    name, title = fmt_nominator(n["Nominator"])
    chips = []
    if n["Exempt"]: chips.append(chip(f"{n['Exempt']} Exempt", BLUE_LT, BLUE))
    if n["Scheduled"]: chips.append(chip(f"{n['Scheduled']} Scheduled", GREEN_LT, GREEN))
    if n["Notes"]: chips.append(chip(n["Notes"], ORANGE_LT, ORANGE))
    total = n["Exempt"] + n["Scheduled"]
    chiprow = Table([[c] for c in chips], colWidths=[1.2 * inch], hAlign="CENTER") if chips else Spacer(1, 1)
    if chips: chiprow.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    body_style = st["card_body"]
    ach = esc(n["Achievement"])
    left = [Paragraph(f"<font color='{MIST.hexval()}'>#{idx:02d}</font>&nbsp;&nbsp;{esc(n['Inscription'])}", st["card_title"]), Spacer(1, 3),
            Paragraph(ach, body_style), Spacer(1, 5),
            Paragraph((f"<b>Nominated by</b> {esc(name)}" + (f" <font color='{MIST.hexval()}'>· {esc(title)}</font>" if title else "") if name.strip()
                       else f"<font color='{MIST.hexval()}'>Rec. No. {esc(n.get('Rec', ''))}</font>" if n.get("Rec") else ""), st["card_meta"])]
    right = [Paragraph(f"{total}", ps("hn", fontName="Helvetica-Bold", fontSize=20, leading=22, textColor=INK, alignment=TA_CENTER)),
             Paragraph("honoree" + ("s" if total != 1 else ""), ps("hl", fontSize=7.5, leading=9, textColor=MIST, alignment=TA_CENTER)), Spacer(1, 6), chiprow]
    t = Table([[left, right]], colWidths=[W - 2 * M - 1.45 * inch - 0.06 * inch, 1.45 * inch])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (-1, -1), WHITE),
                           ("BOX", (0, 0), (-1, -1), 0.6, LINE), ("LINEBEFORE", (0, 0), (0, -1), 3, ORANGE),
                           ("BACKGROUND", (1, 0), (1, -1), PANEL),
                           ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                           ("LEFTPADDING", (0, 0), (0, -1), 11), ("RIGHTPADDING", (0, 0), (0, -1), 10),
                           ("LEFTPADDING", (1, 0), (1, -1), 6), ("RIGHTPADDING", (1, 0), (1, -1), 6)]))
    return [t, Spacer(1, 8)]

def dept_banner(dept, count, honorees):
    t = Table([[Paragraph(esc(dept).upper(), ps("db", fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=WHITE)),
                Paragraph(f"{count} nomination{'s' if count != 1 else ''} · {honorees} honoree{'s' if honorees != 1 else ''}",
                          ps("dbr", fontSize=8.5, leading=12, textColor=WHITE, alignment=2))]],
              colWidths=[(W - 2 * M) * 0.6, (W - 2 * M) * 0.4])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), INK), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                           ("LEFTPADDING", (0, 0), (0, 0), 10), ("RIGHTPADDING", (-1, 0), (-1, 0), 10), ("ROUNDEDCORNERS", [4, 4, 4, 4])]))
    return t

# ---------- page decorations ----------
def cover_page(c, doc):
    c.saveState()
    c.setFillColor(INK); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(ORANGE); c.rect(0, H - 0.35 * inch, W, 0.35 * inch, fill=1, stroke=0)
    c.setFillColor(ORANGE); c.rect(M, H * 0.55, 1.5 * inch, 4, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 11); c.drawString(M, H - 0.9 * inch, "BNSF  |  ACHIEVEMENT AWARDS PROGRAM")
    c.setFont("Helvetica-Bold", 40); c.drawString(M, H * 0.55 + 0.9 * inch + 44, f"{YEAR} Employee")
    c.drawString(M, H * 0.55 + 0.9 * inch, "of the Year")
    c.setFont("Helvetica", 16); c.setFillColor(colors.HexColor("#CBD2D9"))
    c.drawString(M, H * 0.55 - 0.3 * inch, "Nomination Report  ·  Executive Review Package")
    c.setFont("Helvetica", 10.5)
    c.drawString(M, H * 0.55 - 0.65 * inch, f"Prepared for {S.get('Prepared For','')}   ·   {S.get('Report Date','')}")
    # KPI strip
    kp = [(S["Total Achievement Awards"], "Achievement Awards"), (S["Total Employees Recognized"], "Employees Recognized"),
          (S["EOY Awards"], "EOY Awards"), (S["Total EOY Honorees"], "EOY Honorees")]
    bw = (W - 2 * M - 3 * 0.2 * inch) / 4; y = 1.6 * inch
    for i, (v, l) in enumerate(kp):
        x = M + i * (bw + 0.2 * inch)
        c.setFillColor(colors.HexColor("#2B3743")); c.roundRect(x, y, bw, 1.15 * inch, 6, fill=1, stroke=0)
        c.setFillColor(ORANGE); c.rect(x, y + 1.15 * inch - 3, bw, 3, fill=1, stroke=0)
        c.setFillColor(WHITE); c.setFont("Helvetica-Bold", 26); c.drawCentredString(x + bw / 2, y + 0.55 * inch, f"{v:,}")
        c.setFillColor(colors.HexColor("#CBD2D9")); c.setFont("Helvetica", 8.5); c.drawCentredString(x + bw / 2, y + 0.28 * inch, l)
    c.setFillColor(MIST); c.setFont("Helvetica", 8)
    c.drawString(M, 0.8 * inch, S.get("Confidentiality Line", ""))
    c.restoreState()

def body_page(c, doc):
    c.saveState()
    c.setFillColor(ORANGE); c.rect(0, H - 0.18 * inch, W, 0.18 * inch, fill=1, stroke=0)
    c.setFillColor(MIST); c.setFont("Helvetica", 8)
    c.drawString(M, H - 0.5 * inch, f"{YEAR} Employee of the Year Nomination Report")
    c.drawRightString(W - M, H - 0.5 * inch, "Executive Review Package")
    c.setStrokeColor(LINE); c.setLineWidth(0.5); c.line(M, H - 0.58 * inch, W - M, H - 0.58 * inch)
    c.line(M, 0.6 * inch, W - M, 0.6 * inch)
    c.drawString(M, 0.42 * inch, S.get("Confidentiality Line", ""))
    c.drawRightString(W - M, 0.42 * inch, f"Page {doc.page}")
    c.restoreState()

# ---------- build ----------
doc = BaseDocTemplate(OUT, pagesize=letter, leftMargin=M, rightMargin=M, topMargin=0.85 * inch, bottomMargin=0.85 * inch,
                      title=f"{YEAR} EOY Nomination Report", author=str(S.get("Prepared By", "")))
frame = Frame(M, 0.85 * inch, W - 2 * M, H - 1.7 * inch, id="f", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
doc.addPageTemplates([PageTemplate(id="cover", frames=[frame], onPage=cover_page),
                      PageTemplate(id="body", frames=[frame], onPage=body_page)])
story = [NextPageTemplate("body"), PageBreak()]

# --- Executive summary ---
story += section_header("Section 01", "Executive Summary",
    f"Snapshot of the {YEAR} Achievement Award program and Employee of the Year (EOY) nominations submitted for executive review.")
story.append(kpi_row([(S["Total Achievement Awards"], "Achievement Awards"), (S["Total Employees Recognized"], "Employees recognized"),
                      (S["EOY Awards"], "EOY awards nominated"), (S["EOY AA Honorees"], "EOY AA honorees")]))
story.append(Spacer(1, 8))

ex, sc = S["Exempt Employees"], S["Scheduled Employees"]
d1 = fig_to_img(donut([ex, sc], ["Exempt", "Scheduled"], ["#1F5F8B", "#F26522"], (f"{ex/(ex+sc)*100:.0f}%", "Exempt")), 1.2 * inch)
bob, seoy, aah = S["BoB (Exempt)"], S["SEOY (Scheduled)"], S["EOY AA Honorees"]
d2 = fig_to_img(donut([aah, bob, seoy], ["AA", "BoB", "SEOY"], ["#F26522", "#1F5F8B", "#2E7D5B"], (f"{S['Total EOY Honorees']}", "EOY honorees")), 1.2 * inch)
legend1 = Paragraph(f"<font color='#1F5F8B'><b>■</b></font> Exempt {ex:,}&nbsp;&nbsp;&nbsp;<font color='#F26522'><b>■</b></font> Scheduled {sc:,}", st["small"])
legend2 = Paragraph(f"<font color='#F26522'><b>■</b></font> AA {aah}&nbsp;&nbsp;<font color='#1F5F8B'><b>■</b></font> BoB {bob}&nbsp;&nbsp;<font color='#2E7D5B'><b>■</b></font> SEOY {seoy}", st["small"])
has_loc = any(r[0] != "Other" and (r[1] or 0) > 0 for r in LOCS)
other_pct = next((r[1] for r in LOCS if r[0] == "Other"), 0)
cells = [[Paragraph("Employee status", st["h2"]), d1, legend1], [Paragraph("EOY honoree mix", st["h2"]), d2, legend2]]
widths = [2.15 * inch, 2.15 * inch]
if has_loc:
    loc = fig_to_img(location_bar(), 2.5 * inch)
    cells.append([Paragraph("Top work locations", st["h2"]), loc,
                  Paragraph(f"Share of recognized employees." + (f" Remaining {other_pct}% across other locations." if other_pct else ""), st["small"])])
    widths.append(W - 2 * M - 4.3 * inch)
else:
    widths = [(W - 2 * M) / 2] * 2
panel = Table([cells], colWidths=widths)
panel.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 8)]))
story.append(panel); story.append(Spacer(1, 6))

if len(HIST) >= 2:
    story.append(Paragraph(f"{len(HIST)}-year trend", st["h2"]))
    story.append(Paragraph("Achievement Awards submitted and EOY awards selected each year.", st["small"]))
    story.append(fig_to_img(trend_chart(), W - 2 * M))
    story.append(Spacer(1, 2))
    hist_rows = [[r[0], r[1], r[2], r[3], f"{r[4]}/{r[5]}", r[6], r[7], r[8]] for r in HIST]
    ht = styled_table(["Year", "Total AA", "EOY Awards", "AA Honorees", "Sched / Exempt", "BoB (Exempt)", "SEOY (Sched)", "Total EOY Honorees"],
                      hist_rows, [0.7 * inch] + [(W - 2 * M - 0.7 * inch) / 7] * 7)
    ht.setStyle(TableStyle([("BACKGROUND", (0, 1), (-1, 1), ORANGE_LT), ("TEXTCOLOR", (0, 1), (-1, 1), ORANGE)]))
    story.append(ht)
else:
    story.append(Paragraph("Historical trend", st["h2"]))
    story.append(Paragraph("Add prior years to the History sheet to show the multi-year trend here.", st["small"]))
story.append(PageBreak())

# --- Department breakdown ---
story += section_header("Section 02", "Department Breakdown",
    "Achievement Awards submitted by nominator department, with the EOY awards and honorees currently under consideration.")
tot = ["Total"] + [sum((r[i] or 0) for r in DEPTS) for i in range(1, 6)]
dep_rows = [[r[0], r[1], r[2], r[3], r[4], r[5]] for r in DEPTS] + [tot]
cw = W - 2 * M
dt = styled_table(["Nominator Department", "Total AA", "Scheduled", "Exempt", "EOY Awards", "EOY Honorees"], dep_rows,
                  [cw - 5 * 0.95 * inch] + [0.95 * inch] * 5, bold_last=True)
story.append(dt); story.append(Spacer(1, 16))
bar = fig_to_img(dept_bar(), 3.4 * inch)
left = [Paragraph("Awards by department", st["h2"]),
        Paragraph("<font color='#F26522'><b>■</b></font> EOY awards&nbsp;&nbsp;&nbsp;<font color='#F5C9B0'><b>■</b></font> All Achievement Awards", st["small"]), Spacer(1, 4), bar]
right = [Paragraph("EOY honorees over time", st["h2"]), Paragraph("Total EOY honorees (AA + BoB + SEOY) by year.", st["small"]), Spacer(1, 4)]
right.append(fig_to_img(honoree_trend(), 3.2 * inch) if len(HIST) >= 2 else Paragraph("Not enough history yet.", st["small"]))
grid = Table([[left, right]], colWidths=[cw / 2, cw / 2])
grid.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
story.append(grid)
story.append(PageBreak())

# --- Nomination sections ---
DEPT_ORDER = [r[0] for r in DEPTS]
def grouped(section):
    g = OrderedDict((d, []) for d in DEPT_ORDER)
    for n in NOMS:
        if n["Section"] == section:
            g.setdefault(n["Department"], []).append(n)
    return [(d, v) for d, v in g.items() if v]

def nomination_section(num, section, title, sub):
    items = grouped(section)
    n_total = sum(len(v) for _, v in items); h_total = sum(x["Exempt"] + x["Scheduled"] for _, v in items for x in v)
    out = section_header(f"Section {num:02d}", title, sub)
    out.append(kpi_row([(n_total, "Nominations"), (len(items), "Departments"),
                        (h_total, "Honorees listed"), (sum(x["Exempt"] for _, v in items for x in v), "Exempt honorees")]))
    out.append(Spacer(1, 10))
    # contents strip
    toc = [[Paragraph(esc(d), st["td"]), Paragraph(str(len(v)), st["tdr"]), Paragraph(str(sum(x["Exempt"] + x["Scheduled"] for x in v)), st["tdr"])] for d, v in items]
    tt = Table([[Paragraph("Department", st["th"]), Paragraph("Nominations", st["th"]), Paragraph("Honorees", st["th"])]] + toc,
               colWidths=[3.2 * inch, 1.2 * inch, 1.2 * inch], hAlign="LEFT")
    tt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), INK), ("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE), ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    out += [Paragraph("Contents", st["h2"]), tt, Spacer(1, 16)]
    idx = 0
    for d, v in items:
        banner = [dept_banner(d, len(v), sum(x["Exempt"] + x["Scheduled"] for x in v)), Spacer(1, 8)]
        for j, n in enumerate(v):
            idx += 1
            card = nomination_card(n, idx)
            out.append(KeepTogether(banner + card if j == 0 else card))
        out.append(Spacer(1, 6))
    out.append(PageBreak())
    return out

story += nomination_section(3, "EOY", "Employee of the Year Nominations",
    "Achievement Awards nominated for Employee of the Year recognition. Each card shows the inscription, achievement summary, nominator and honoree count.")
story += nomination_section(4, "AA", f"Remaining {YEAR} Achievement Awards",
    "All other Achievement Awards submitted during the year, grouped by nominator department, for reference.")

# --- Legend / notes ---
story += section_header("Appendix", "Definitions & Notes")
defs = [("Achievement Award (AA)", "Recognition submitted by a nominator for an individual or team accomplishment during the award year."),
        ("Employee of the Year (EOY)", "Achievement Awards selected for elevated, company-wide recognition."),
        ("BoB (Exempt)", "Best of the Best honorees from the exempt (salaried) workforce."),
        ("SEOY (Scheduled)", "Scheduled Employee of the Year honorees from the scheduled (craft) workforce."),
        ("Exempt / Scheduled", "Honoree counts by employee status on each nomination card."),
        ("Hero", "Nominations flagged for life-saving or emergency response actions.")]
for k, v in defs:
    story.append(Paragraph(f"<b>{k}</b> — {v}", st["body"])); story.append(Spacer(1, 4))
story.append(Spacer(1, 10))
story.append(Paragraph(f"{S.get('Rooms Reserved','')} total rooms reserved for {YEAR} honorees.", st["body"])); story.append(Spacer(1, 6))
story.append(Paragraph(f"Source: Achievement Awards SharePoint list, exported {S.get('Report Date','')}. Generated automatically from EOY_Report_Data.xlsx.", st["small"]))
if story and isinstance(story[-1], PageBreak): story.pop()
doc.build(story)
print("wrote", OUT)
