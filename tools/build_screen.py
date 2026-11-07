import openpyxl, re, sys, datetime, collections
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
AS_OF = datetime.date(2026, 9, 30)

# ------------------------------------------------------------------ styles
F = "Arial"
font = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
hdr_font = Font(name=F, size=10, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F3864")
inp_font = Font(name=F, size=10, color="0000FF")
inp_fill = PatternFill("solid", fgColor="FFFF00")
sub_fill = PatternFill("solid", fgColor="D9E1F2")
title_font = Font(name=F, size=14, bold=True, color="1F3864")
thin = Side(style="thin", color="BFBFBF")
border = Border(top=thin, bottom=thin, left=thin, right=thin)

def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.border = hdr_font, hdr_fill, border
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30

def put(ws, r, c, v, f=font, fill=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = f
    cell.border = border
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    return cell

# ------------------------------------------------------------------ source
src = openpyxl.load_workbook(SRC, data_only=True)["Investment List"]
rows = [r for r in src.iter_rows(min_row=8, values_only=True) if r[0] and r[2]]

def num(x):
    return x if isinstance(x, (int, float)) else None

CLASS_RE = re.compile(
    r"\s+(Class|Cl\.?)\s+\S+(\s+Shares?)?$|"
    r"\s+(Institutional Premium|Institutional Plus|Institutional|Instl|Inst|Admiral|Investor|Premium|Service|Select|Retirement|Advisor|Founders|Ultra|Supra|Plan|Platform|Super)\s+(Class|Shares?)(\s+Shares?)?$|"
    r"\s+(I|I2|I3|I5|IS|Y|Z|Z6|N|F|F2|F3|K|K6|R|R2|R3|R4|R5|R6|A|C|T|U|P|W|X|Q|M|TR)\s+(Class|Shares?)$|"
    r"\s+(Class|Cl\.?)\s*(I|I2|I3|I5|IS|Y|Z|Z6|N|F|F2|F3|K|K6|R|R2|R3|R4|R5|R6|A|C|T|U|P|W|X|Q|M|AAA|TR|R-\d)$|"
    r"\s+\((7-Day Yield.*?)\)$",
    re.I,
)
def base_name(n):
    n = re.sub(r"\s*\(7-Day Yield[^)]*\)", "", n)
    prev = None
    while prev != n:
        prev = n
        n = CLASS_RE.sub("", n).strip()
        n = re.sub(r"\s+(Shares|Share Class)$", "", n, flags=re.I).strip()
    return n

# ------------------------------------------------------------------ slots
# (slot name, tier, categories, 3(21) min, 3(38) target, expense cap override or None)
SLOTS = [
    ("Large Cap Value", "Core", ["Large Value"], 3, 1, None),
    ("Large Cap Blend", "Core", ["Large Blend"], 3, 1, None),
    ("Large Cap Growth", "Core", ["Large Growth"], 3, 1, None),
    ("Mid Cap Value", "Core", ["Mid-Cap Value"], 3, 1, None),
    ("Mid Cap Blend", "Core", ["Mid-Cap Blend"], 3, 1, None),
    ("Mid Cap Growth", "Core", ["Mid-Cap Growth"], 3, 1, None),
    ("Small Cap Value", "Core", ["Small Value"], 3, 1, None),
    ("Small Cap Blend", "Core", ["Small Blend"], 3, 1, None),
    ("Small Cap Growth", "Core", ["Small Growth"], 3, 1, None),
    ("International Equity", "Core", ["Foreign Large Blend", "Foreign Large Value", "Foreign Large Growth"], 3, 1, None),
    ("Emerging Markets Equity", "Core", ["Diversified Emerging Mkts"], 3, 1, None),
    ("Global Equity", "Optional", ["Global Large-Stock Blend", "Global Large-Stock Growth", "Global Large-Stock Value"], 2, 1, None),
    ("Real Estate", "Optional", ["Real Estate"], 2, 1, None),
    ("Core Bond", "Core", ["Intermediate Core Bond"], 3, 1, 0.50),
    ("Core-Plus Bond", "Core", ["Intermediate Core-Plus Bond"], 3, 1, 0.50),
    ("Short-Term Bond", "Core", ["Short-Term Bond"], 3, 1, 0.50),
    ("High Yield Bond", "Optional", ["High Yield Bond"], 2, 1, 0.75),
    ("Inflation-Protected Bond", "Optional", ["Inflation-Protected Bond"], 2, 1, 0.50),
    ("Multisector Bond", "Optional", ["Multisector Bond"], 2, 1, 0.75),
    ("Capital Preservation", "Core", ["Money Market-Taxable"], 2, 1, 0.35),
    ("Balanced", "Optional", ["Moderate Allocation"], 2, 1, None),
    ("Target-Date Series", "Core (QDIA)", ["Target-Date *"], 2, 1, 0.50),
]
cat2slot = {}
for s in SLOTS:
    for c in s[2]:
        cat2slot[c] = s[0]
td_cats = sorted({r[5] for r in rows if r[5] and r[5].startswith("Target-Date")})
for c in td_cats:
    cat2slot[c] = "Target-Date Series"

wb = openpyxl.Workbook()

# ================================================================== README
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 120
lines = [
    ("BCA Financial Group — 401(k) Fund Screening Workbook", title_font),
    (f"Universe: Fidelity 401(k) Direct Program (FPRS) platform list, data as of {AS_OF:%B %d, %Y}. Source: Fidelity 'FPRS_Direct_Platform_List_Q3_2026' (supplied by Jim Bianchi).", font),
    ("", font),
    ("PURPOSE", bold),
    ("Step 1 of building BCA's two covered fund lists: a 3(21) Acceptable Universe (several funds per lineup slot) and a 3(38) Select Lineup (one fund per slot, a strict subset).", font),
    ("This workbook applies the mechanical screens that need no judgment. Fiduciary scoring (Fi360 Fiduciary Score / LPL 12-Point) is layered on afterwards in the Candidates tab, and the final lists are created as Covered Fund Lists in the Broadridge Fiduciary Focus Toolkit.", font),
    ("", font),
    ("HOW TO USE", bold),
    ("1. Screens tab — the policy parameters. Blue text on yellow = editable. Everything else is a formula.", font),
    ("2. Slots tab — the lineup structure: which Morningstar categories feed each slot, and how many funds each list needs per slot. Editable columns are blue on yellow.", font),
    ("3. Fidelity Universe tab — every share class on the platform, with live Pass/Fail formulas for each screen. Change a parameter on Screens and these recalculate.", font),
    ("4. Candidates tab — the survivors (one share class per fund, lowest net expense), ready for scoring. Fill in Fi360 Score and LPL 12-Pt, then mark 3(21) Include and 3(38) Select. This tab is a snapshot: if you change a screen, re-run the build script to regenerate it (the Universe tab stays live either way).", font),
    ("5. Target-Date tab — target-date series that pass the screens, summarized at the series level.", font),
    ("6. Slot Summary tab — counts per slot, and whether each list has enough selections yet.", font),
    ("", font),
    ("SCREENS APPLIED (see Screens tab for the values)", bold),
    ("  • On the Fidelity platform (On Platform = Y)", font),
    ("  • Status = open to all investors (closed and restricted share classes excluded)", font),
    ("  • Zero recordkeeping offset — no revenue sharing, for BOTH lists (decision: Jim Bianchi, 2026-11-06)", font),
    ("  • Net expense ratio at or below the cap (global cap, with per-slot overrides for fixed income and capital preservation)", font),
    ("  • Minimum track record, measured from inception date to the as-of date", font),
    ("  • Category maps to a defined lineup slot", font),
    ("  • Lowest-cost share class per fund (share-class names normalized by script; spot-check duplicates)", font),
    ("", font),
    ("NOT YET APPLIED (next step)", bold),
    ("  • Fi360 Fiduciary Score® and LPL 12-Point score — look up in the Toolkit and enter on the Candidates tab", font),
    ("  • Manager tenure, style consistency, qualitative review — documented per the Monitoring Policy", font),
    ("", font),
    ("LEGEND", bold),
    ("  Blue text / yellow fill = input you may change    Black = formula    Green = link to another tab", font),
    ("", font),
    ("Performance data is Fidelity's as of the date above; past performance is no guarantee of future results. For plan sponsor and investment professional use only.", Font(name=F, size=9, italic=True)),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = f; c.alignment = Alignment(wrap_text=True, vertical="top")

# ================================================================== Screens
sc = wb.create_sheet("Screens")
sc.column_dimensions["A"].width = 48; sc.column_dimensions["B"].width = 34; sc.column_dimensions["C"].width = 70
sc["A1"] = "Screening Parameters"; sc["A1"].font = title_font
header(sc, 3, ["Parameter", "Value", "Notes"])
params = [
    ("As-of date", AS_OF, "Fidelity data date. Track record is measured to this date.", "mm/dd/yyyy"),
    ("Require On Platform = ", "Y", "Fidelity 'On Platform (Y/N)' column", None),
    ("Allowed status", "WI-DC: Open to all investors", "Exact Fidelity status text. Closed/restricted classes fail.", None),
    ("Maximum recordkeeping offset (%)", 0, "0 = zero revenue share. Applies to BOTH the 3(21) and 3(38) lists.", "0.00"),
    ("Global maximum net expense ratio (%)", 0.75, "Used when a slot has no override on the Slots tab.", "0.00"),
    ("Minimum track record (years)", 5, "From inception date. 3(38) selections should generally also have a 10-year record; see Candidates tab '10 Yr Record?' column.", "0"),
    ("Preferred 10-year record for 3(38) (years)", 10, "Informational flag on Candidates; not a hard screen.", "0"),
    ("3(21) target funds per slot (default)", 3, "Default; Slots tab can override per slot.", "0"),
    ("3(38) target funds per slot (default)", 1, "Default; Slots tab can override per slot.", "0"),
]
for i, (p, v, n, fmt) in enumerate(params, 4):
    put(sc, i, 1, p, bold)
    c = put(sc, i, 2, v, inp_font, inp_fill, fmt)
    put(sc, i, 3, n)
# named refs
P = {  # cell addresses on Screens
    "asof": "Screens!$B$4", "platform": "Screens!$B$5", "status": "Screens!$B$6",
    "offset": "Screens!$B$7", "er": "Screens!$B$8", "track": "Screens!$B$9", "track10": "Screens!$B$10",
    "n321": "Screens!$B$11", "n338": "Screens!$B$12",
}
sc["A15"] = "Decision log"; sc["A15"].font = bold
sc["A16"] = "2026-11-06 — Zero revenue-share share classes required for both the 3(21) Acceptable Universe and the 3(38) Select Lineup (Jim Bianchi)."
sc["A16"].font = font

# ================================================================== Slots
sl = wb.create_sheet("Slots")
sl["A1"] = "Lineup Slots and Category Map"; sl["A1"].font = title_font
header(sl, 3, ["Slot", "Tier", "Morningstar categories (feed this slot)", "3(21) target # funds", "3(38) target # funds", "Net ER cap override (%)", "Effective ER cap (%)"],
       [28, 14, 70, 14, 14, 16, 14])
for i, (name, tier, cats, n21, n38, cap) in enumerate(SLOTS, 4):
    put(sl, i, 1, name, bold)
    put(sl, i, 2, tier)
    put(sl, i, 3, ", ".join(cats))
    put(sl, i, 4, n21, inp_font, inp_fill, "0")
    put(sl, i, 5, n38, inp_font, inp_fill, "0")
    put(sl, i, 6, cap, inp_font, inp_fill, "0.00")
    put(sl, i, 7, f'=IF(F{i}="",{P["er"]},F{i})', fmt="0.00")
slot_first, slot_last = 4, 3 + len(SLOTS)
# category map table
mr = slot_last + 3
sl.cell(row=mr - 1, column=1, value="Category → Slot map (used by the Universe tab lookup; add a row to map a new category)").font = bold
header(sl, mr, ["Morningstar Category", "Slot"])
map_rows = sorted(cat2slot.items())
for i, (cat, slot) in enumerate(map_rows, mr + 1):
    put(sl, i, 1, cat, inp_font, inp_fill)
    put(sl, i, 2, slot, inp_font, inp_fill)
map_first, map_last = mr + 1, mr + len(map_rows)
MAPCAT = f"Slots!$A${map_first}:$A${map_last}"
MAPSLOT = f"Slots!$B${map_first}:$B${map_last}"
SLOTNAMES = f"Slots!$A${slot_first}:$A${slot_last}"
SLOTCAP = f"Slots!$G${slot_first}:$G${slot_last}"
sl.freeze_panes = "A4"

# ================================================================== Fidelity Universe
un = wb.create_sheet("Fidelity Universe")
cols = ["Fund Name", "CUSIP", "Ticker", "Status", "On Platform", "Morningstar Category",
        "YTD %", "1 Yr %", "3 Yr %", "5 Yr %", "10 Yr %", "Inception", "Gross ER %", "Net ER %", "RK Offset %", "Index Fund",
        "Fund (share-class stripped)", "Slot", "Track Record (yrs)", "Slot ER Cap %",
        "Pass: Platform", "Pass: Status", "Pass: Offset", "Pass: Expense", "Pass: Track", "Pass: Slot", "PASS ALL", "Lowest-Cost Class", "CANDIDATE"]
header(un, 1, cols, [52, 11, 8, 28, 9, 26, 8, 8, 8, 8, 8, 11, 8, 8, 9, 8, 46, 24, 10, 9, 9, 9, 9, 9, 9, 9, 9, 10, 11])
n = len(rows)
last = n + 1
for i, r in enumerate(rows, 2):
    vals = [r[0], r[1], r[2], r[3], r[4], r[5], num(r[6]), num(r[7]), num(r[8]), num(r[9]), num(r[10]),
            r[12].date() if hasattr(r[12], "date") else r[12], num(r[14]), num(r[15]), num(r[16]), r[17], base_name(r[0])]
    for j, v in enumerate(vals, 1):
        c = un.cell(row=i, column=j, value=v); c.font = font
        if 7 <= j <= 11 or 13 <= j <= 15: c.number_format = "0.00"
        if j == 12: c.number_format = "mm/dd/yyyy"
    f = {
        18: f'=IFERROR(INDEX({MAPSLOT},MATCH(F{i},{MAPCAT},0)),"")',
        19: f'=IF(L{i}="","",({P["asof"]}-L{i})/365.25)',
        20: f'=IFERROR(INDEX({SLOTCAP},MATCH(R{i},{SLOTNAMES},0)),{P["er"]})',
        21: f'=IF(E{i}={P["platform"]},"Y","N")',
        22: f'=IF(D{i}={P["status"]},"Y","N")',
        23: f'=IF(AND(ISNUMBER(O{i}),O{i}<={P["offset"]}),"Y","N")',
        24: f'=IF(AND(ISNUMBER(N{i}),N{i}<=T{i}),"Y","N")',
        25: f'=IF(AND(ISNUMBER(S{i}),S{i}>={P["track"]}),"Y","N")',
        26: f'=IF(R{i}<>"","Y","N")',
        27: f'=IF(COUNTIF(U{i}:Z{i},"Y")=6,"Y","N")',
        28: f'=IF(N{i}=_xlfn.MINIFS($N$2:$N${last},$Q$2:$Q${last},Q{i},$AA$2:$AA${last},"Y"),"Y","N")',
        29: f'=IF(AND(AA{i}="Y",AB{i}="Y"),"Y","N")',
    }
    for j, fx in f.items():
        c = un.cell(row=i, column=j, value=fx); c.font = font
        if j == 19: c.number_format = "0.0"
        if j == 20: c.number_format = "0.00"
un.freeze_panes = "D2"
un.auto_filter.ref = f"A1:AC{last}"
un["Q1"].comment = Comment("Share-class suffix removed by script so the lowest-cost class of each fund can be identified. Heuristic — spot-check funds that appear twice in Candidates.", "BCA")

# ================================================================== Python-side evaluation (for Candidates snapshot)
def passes(r):
    cat = r[5]; slot = cat2slot.get(cat)
    if not slot or slot == "Target-Date Series": return False
    cap = next((s[5] for s in SLOTS if s[0] == slot), None) or 0.75
    ner, off = num(r[15]), num(r[16])
    inc = r[12].date()
    yrs = (AS_OF - inc).days / 365.25
    return (r[4] == "Y" and r[3] == "WI-DC: Open to all investors" and off is not None and off <= 0
            and ner is not None and ner <= cap and yrs >= 5)
surv = [r for r in rows if passes(r)]
best = {}
for r in surv:
    k = (base_name(r[0]), r[5])
    if k not in best or num(r[15]) < num(best[k][15]): best[k] = r
cands = sorted(best.values(), key=lambda r: (cat2slot[r[5]], r[5], num(r[15])))

# ================================================================== Candidates
ca = wb.create_sheet("Candidates")
ccols = ["Slot", "Morningstar Category", "Fund Name", "Ticker", "CUSIP", "Net ER %", "1 Yr %", "3 Yr %", "5 Yr %", "10 Yr %",
         "Track Record (yrs)", "10 Yr Record?", "Index Fund", "ER rank in slot (1 = cheapest)", "5-Yr return rank in slot", "10-Yr return rank in slot",
         "Fi360 Fiduciary Score®", "Fi360 Score (Qtr-1)", "Fi360 Score (Qtr-2)", "Fi360 Score (Qtr-3)", "LPL 12-Point", "Score Status",
         "3(21) Include (Y/N)", "3(38) Select (Y/N)", "Reviewer", "Notes / rationale"]
header(ca, 1, ccols, [24, 26, 52, 8, 11, 8, 8, 8, 8, 8, 10, 9, 8, 12, 12, 12, 11, 10, 10, 10, 10, 14, 10, 10, 12, 50])
cl = len(cands) + 1
for i, r in enumerate(cands, 2):
    yrs = (AS_OF - r[12].date()).days / 365.25
    vals = [cat2slot[r[5]], r[5], r[0], r[2], r[1], num(r[15]), num(r[7]), num(r[8]), num(r[9]), num(r[10]), round(yrs, 1)]
    for j, v in enumerate(vals, 1):
        c = ca.cell(row=i, column=j, value=v); c.font = font
        if 6 <= j <= 10: c.number_format = "0.00"
        if j == 11: c.number_format = "0.0"
    ca.cell(row=i, column=13, value=r[17]).font = font
    fx = {
        12: f'=IF(K{i}>={P["track10"]},"Y","N")',
        14: f'=COUNTIFS($A$2:$A${cl},A{i},$F$2:$F${cl},"<"&F{i})+1',
        15: f'=IF(ISNUMBER(I{i}),COUNTIFS($A$2:$A${cl},A{i},$I$2:$I${cl},">"&I{i})+1,"")',
        16: f'=IF(ISNUMBER(J{i}),COUNTIFS($A$2:$A${cl},A{i},$J$2:$J${cl},">"&J{i})+1,"")',
        22: f'=IF(Q{i}="","Not scored",IF(Q{i}<=25,"Top quartile",IF(Q{i}<=50,"Top half","Below median")))',
    }
    for j, f_ in fx.items():
        ca.cell(row=i, column=j, value=f_).font = font
    for j in (17, 18, 19, 20, 21, 23, 24, 25, 26):
        c = ca.cell(row=i, column=j); c.font = inp_font; c.fill = inp_fill; c.border = border
ca.freeze_panes = "D2"
ca.auto_filter.ref = f"A1:Z{cl}"
ca["Q1"].comment = Comment("Enter from the Broadridge Fiduciary Focus Toolkit (0 = best, 100 = worst). Quartile: ≤25 top quartile, ≤50 top half.", "BCA")
ca["U1"].comment = Comment("LPL 12-Point score (0–12). LPL watch-list trigger in the Toolkit: ≤6 fails.", "BCA")
ca["W1"].comment = Comment("Y = include in the 3(21) Acceptable Universe covered fund list.", "BCA")
ca["X1"].comment = Comment("Y = the single 3(38) selection for this slot. Must also be Y in the 3(21) column.", "BCA")
# example row marker
ca.cell(row=cl + 2, column=1, value="Example of a completed row (illustrative values, not a recommendation):").font = Font(name=F, size=9, italic=True)
ex = ["Large Cap Blend", "Large Blend", "Example Index Fund — Institutional", "XXXXX", "000000000", 0.02, 17.5, 23.1, 15.2, 14.6, 25.0, "Y", "Y", 1, 12, 9, 18, 20, 22, 19, 11, "Top quartile", "Y", "Y", "NDA", "Passive core; lowest cost on platform; Fi360 top quartile 4 consecutive quarters"]
for j, v in enumerate(ex, 1):
    c = ca.cell(row=cl + 3, column=j, value=v); c.font = Font(name=F, size=9, italic=True, color="808080")

# ================================================================== Target-Date
td = wb.create_sheet("Target-Date")
td["A1"] = "Target-Date Series (passing: on platform, open, zero offset, net ER ≤ cap)"; td["A1"].font = title_font
header(td, 3, ["Series (share-class stripped)", "Vintages on platform", "Avg Net ER %", "Max Net ER %", "Index-based?", "Earliest inception", "Avg 5 Yr % (2040 vintage if present)",
               "Fi360 Score (2040 vintage)", "LPL 12-Point", "3(21) Include (Y/N)", "3(38) Select (Y/N)", "Notes"],
       [56, 10, 10, 10, 10, 14, 16, 12, 10, 10, 10, 50])
series = collections.defaultdict(list)
tdcap = 0.50
for r in rows:
    if r[5] in td_cats and r[4] == "Y" and r[3] == "WI-DC: Open to all investors" and (num(r[16]) or 0) <= 0 and num(r[15]) is not None and num(r[15]) <= tdcap:
        nm = re.sub(r"\s*(19|20)\d\d\s*", " ", base_name(r[0])).replace("  ", " ").strip()
        nm = re.sub(r"\s+(Fund|Portfolio|Trust)$", "", nm).strip()
        series[nm].append(r)
ser = sorted(series.items(), key=lambda kv: (sum(num(x[15]) for x in kv[1]) / len(kv[1])))
for i, (nm, rs) in enumerate(ser, 4):
    ers = [num(x[15]) for x in rs]
    v2040 = [x for x in rs if x[5] == "Target-Date 2040"]
    put(td, i, 1, nm, bold)
    put(td, i, 2, len(rs), fmt="0")
    put(td, i, 3, sum(ers) / len(ers), fmt="0.00")
    put(td, i, 4, max(ers), fmt="0.00")
    put(td, i, 5, "Y" if any(x[17] == "Y" for x in rs) else "N")
    put(td, i, 6, min(x[12] for x in rs).date(), fmt="mm/dd/yyyy")
    put(td, i, 7, num(v2040[0][9]) if v2040 else None, fmt="0.00")
    for j in (8, 9, 10, 11, 12):
        put(td, i, j, None, inp_font, inp_fill)
td.cell(row=len(ser) + 6, column=1, value=f"Series are grouped by normalized name; a series with fewer than ~8 vintages may have some vintages failing a screen or carrying a different share class. Target-date ER cap used: {tdcap}%. Treat the series as the unit of selection; the Toolkit scores each vintage.").font = Font(name=F, size=9, italic=True)
td.freeze_panes = "B4"

# ================================================================== Slot Summary
ss = wb.create_sheet("Slot Summary")
ss["A1"] = "Slot Summary"; ss["A1"].font = title_font
ss["A2"] = "Live counts from Fidelity Universe (recalculate when Screens change) and from the Candidates tab selections."; ss["A2"].font = font
header(ss, 4, ["Slot", "Tier", "Share classes in slot", "Pass all screens", "Candidates (lowest-cost class)", "3(21) target", "3(21) selected", "3(21) status",
               "3(38) target", "3(38) selected", "3(38) status", "Candidates scored"], [28, 14, 12, 12, 14, 10, 10, 14, 10, 10, 14, 12])
for i, s in enumerate(SLOTS, 5):
    r_ = slot_first + i - 5
    put(ss, i, 1, f"=Slots!A{r_}", Font(name=F, size=10, color="008000"))
    put(ss, i, 2, f"=Slots!B{r_}", Font(name=F, size=10, color="008000"))
    put(ss, i, 3, f"=COUNTIF('Fidelity Universe'!$R$2:$R${last},A{i})", fmt="#,##0")
    put(ss, i, 4, f"=COUNTIFS('Fidelity Universe'!$R$2:$R${last},A{i},'Fidelity Universe'!$AA$2:$AA${last},\"Y\")", fmt="#,##0")
    put(ss, i, 5, f"=COUNTIFS('Fidelity Universe'!$R$2:$R${last},A{i},'Fidelity Universe'!$AC$2:$AC${last},\"Y\")", fmt="#,##0")
    put(ss, i, 6, f"=Slots!D{r_}", Font(name=F, size=10, color="008000"), fmt="0")
    put(ss, i, 7, f"=COUNTIFS(Candidates!$A$2:$A${cl},A{i},Candidates!$W$2:$W${cl},\"Y\")", fmt="0")
    put(ss, i, 8, f'=IF(G{i}>=F{i},"Complete",IF(G{i}=0,"Not started","In progress"))')
    put(ss, i, 9, f"=Slots!E{r_}", Font(name=F, size=10, color="008000"), fmt="0")
    put(ss, i, 10, f"=COUNTIFS(Candidates!$A$2:$A${cl},A{i},Candidates!$X$2:$X${cl},\"Y\")", fmt="0")
    put(ss, i, 11, f'=IF(J{i}=I{i},"Complete",IF(J{i}>I{i},"Too many",IF(J{i}=0,"Not started","In progress")))')
    put(ss, i, 12, f"=COUNTIFS(Candidates!$A$2:$A${cl},A{i},Candidates!$Q$2:$Q${cl},\"<>\")", fmt="0")
tr = 5 + len(SLOTS)
put(ss, tr, 1, "Total", bold)
for col in (3, 4, 5, 7, 10, 12):
    L = get_column_letter(col)
    put(ss, tr, col, f"=SUM({L}5:{L}{tr-1})", bold, fmt="#,##0")
ss.cell(row=tr + 2, column=1, value="Target-Date Series row counts the Universe only; selections for target-date are made on the Target-Date tab.").font = Font(name=F, size=9, italic=True)
ss.freeze_panes = "B5"

wb.move_sheet("Slot Summary", offset=-(len(wb.sheetnames) - 2))
wb.save(OUT)
print("candidates", len(cands), "td series", len(ser), "universe rows", n)
