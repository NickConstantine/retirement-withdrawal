"""Build the Retirement Withdrawal Strategies planner (based on the WSJ 4% rule article)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.comments import Comment

# ---- Style helpers -------------------------------------------------
ARIAL = "Arial"
BLUE = "0000FF"
BLACK = "000000"
GREEN = "008000"
WHITE = "FFFFFF"
MS_BLUE = "0078D4"
LIGHT_GREY = "F3F2F1"
YELLOW = "FFFF00"
PALE_BLUE = "DEECF9"
RED_FILL = "FFC7CE"
RED_FONT = "9C0006"
GREEN_FILL = "C6EFCE"
GREEN_FONT = "006100"

UNLOCK = Protection(locked=False)

def font(bold=False, color=BLACK, size=11, italic=False):
    return Font(name=ARIAL, bold=bold, color=color, size=size, italic=italic)

thin = Side(style="thin", color="BFBFBF")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
bottom_only = Border(bottom=Side(style="thin", color="808080"))
box = Border(left=Side(style="thin", color=MS_BLUE), right=Side(style="thin", color=MS_BLUE),
             top=Side(style="thin", color=MS_BLUE), bottom=Side(style="thin", color=MS_BLUE))

CUR = '$#,##0;($#,##0);"-"'
PCT = '0.0%'
PCT2 = '0.00%'
NUM = '0'
NUM1 = '0.0'

def header_cell(c, text):
    c.value = text
    c.font = font(bold=True, color=WHITE, size=11)
    c.fill = PatternFill("solid", start_color=MS_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all

def title(ws, cell, text, size=16):
    ws[cell] = text
    ws[cell].font = font(bold=True, color="242424", size=size)

def input_cell(c, val, fmt):
    c.value = val
    c.font = font(color=BLACK)
    c.number_format = fmt
    c.fill = PatternFill("solid", start_color=YELLOW)
    c.border = bottom_only
    c.alignment = Alignment(horizontal="right")
    c.protection = UNLOCK

red_rule = CellIsRule(operator="lessThanOrEqual", formula=["0"],
                      fill=PatternFill("solid", start_color=RED_FILL), font=Font(name=ARIAL, color=RED_FONT))

wb = Workbook()

# ===================================================================
# SHEET: Inputs & Summary
# ===================================================================
ws = wb.active
ws.title = "Inputs & Summary"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 48
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 3
ws.column_dimensions["E"].width = 50
ws.column_dimensions["F"].width = 16

title(ws, "B2", "Retirement Withdrawal Strategies")
ws["B3"] = ("Compare the classic 4% rule with a dynamic strategy (life-expectancy + guardrails + shock absorber). "
            "Yellow = inputs you can change · Black = calculated · Green = pulled from another tab. See the Instructions tab.")
ws["B3"].font = font(italic=True, color="595959", size=9)
ws.merge_cells("B3:F3")

# ---- Verdict box (row 4) ----
ws.merge_cells("B4:F4")
ws.row_dimensions[4].height = 46
vcell = ws["B4"]
vcell.fill = PatternFill("solid", start_color=PALE_BLUE)
vcell.border = box
vcell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
vcell.font = font(bold=True, color="1F3864", size=11)
vcell.value = (
    '="At a starting portfolio of "&TEXT($C$8,"$#,##0")&":     '
    '4% rule "&IF($F$10="YES","lasts to age "&$C$7,"runs out before age "&$C$7)&" (about "&TEXT($F$9,"$#,##0")&" left).     '
    'Dynamic strategy "&IF($F$16="YES","lasts to age "&$C$7,"runs out before age "&$C$7)&", with annual withdrawals of "'
    '&TEXT($F$13,"$#,##0")&" to "&TEXT($F$14,"$#,##0")&"."'
)

# ---- INPUTS ----
ws["B6"] = "ASSUMPTIONS (edit the yellow cells)"
ws["B6"].font = font(bold=True, color=WHITE)
ws["B6"].fill = PatternFill("solid", start_color=MS_BLUE)
ws["C6"].fill = PatternFill("solid", start_color=MS_BLUE)

inputs = [
    ("Retirement start age (or current age, if retired)", 50, NUM, "age"),
    ("Plan-to age (planning horizon)", 90, NUM, "age"),
    ("Starting portfolio at retirement (or current portfolio amount, if retired)", 3500000, CUR, "money"),
    ("Expected annual return (nominal, default)", 0.06, PCT, "rate"),
    ("Inflation rate", 0.03, PCT, "rate"),
    ("4% rule starting rate", 0.04, PCT, "rate"),
    ("Guardrail floor (minimum withdrawal rate)", 0.03, PCT, "rate"),
    ("Guardrail ceiling (maximum withdrawal rate)", 0.06, PCT, "rate"),
    ("Shock absorber — max spending change per year", 0.05, PCT, "rate"),
    ("Essential annual spending (today's $)", 100000, CUR, "money"),
    ("Full Social Security benefit at age 67 / FRA (today's $/yr)", 30000, CUR, "money"),
    ("Social Security start age (62–70)", 67, NUM, "ssage"),
    ("Current 4% rule withdrawal if already retired (today's $/yr; 0 if not)", 0, CUR, "money"),
]
r = 7
irow = {}
for label, val, fmt, kind in inputs:
    ws.cell(r, 2, label).font = font()
    input_cell(ws.cell(r, 3), val, fmt)
    irow[label] = (r, kind)
    r += 1

def A(label):
    return f"$C${irow[label][0]}"

start_age = A("Retirement start age (or current age, if retired)")
plan_age = A("Plan-to age (planning horizon)")
port0 = A("Starting portfolio at retirement (or current portfolio amount, if retired)")
exp_ret = A("Expected annual return (nominal, default)")
infl = A("Inflation rate")
rate4 = A("4% rule starting rate")
floor = A("Guardrail floor (minimum withdrawal rate)")
ceiling = A("Guardrail ceiling (maximum withdrawal rate)")
shock = A("Shock absorber — max spending change per year")
essential = A("Essential annual spending (today's $)")
full_ss = A("Full Social Security benefit at age 67 / FRA (today's $/yr)")
ss_age = A("Social Security start age (62–70)")
current_4pct = A("Current 4% rule withdrawal if already retired (today's $/yr; 0 if not)")

S = "'Inputs & Summary'!"

# data validation on inputs
dv_age = DataValidation(type="whole", operator="between", formula1="40", formula2="110", allow_blank=False)
dv_age.error = "Enter a whole age between 40 and 110."
dv_rate = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.25", allow_blank=False)
dv_rate.error = "Enter a rate between 0% and 25%."
dv_money = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=False)
dv_money.error = "Enter a non-negative dollar amount."
dv_ssage = DataValidation(type="whole", operator="between", formula1="62", formula2="70", allow_blank=False)
dv_ssage.error = "Social Security start age must be a whole number between 62 and 70."
for d in (dv_age, dv_rate, dv_money, dv_ssage):
    ws.add_data_validation(d)
for label, (rr, kind) in irow.items():
    ref = f"C{rr}"
    if kind == "age":
        dv_age.add(ref)
    elif kind == "rate":
        dv_rate.add(ref)
    elif kind == "ssage":
        dv_ssage.add(ref)
    else:
        dv_money.add(ref)

# ---- MARKET SCENARIO selector (rows 21-22) ----
ws.cell(21, 2, "MARKET SCENARIO").font = font(bold=True, color=WHITE)
ws.cell(21, 2).fill = PatternFill("solid", start_color=MS_BLUE)
ws.cell(21, 3).fill = PatternFill("solid", start_color=MS_BLUE)
ws.cell(22, 2, "Market scenario (drives the return column)").font = font()
sc = ws.cell(22, 3, "Steady")
sc.font = font(color=BLACK); sc.fill = PatternFill("solid", start_color=YELLOW)
sc.border = bottom_only; sc.alignment = Alignment(horizontal="right"); sc.protection = UNLOCK
dv_scen = DataValidation(type="list", formula1='"Steady,Early crash,Stagflation"', allow_blank=False)
dv_scen.error = "Choose Steady, Early crash, or Stagflation."
ws.add_data_validation(dv_scen); dv_scen.add("C22")
scenario_ref = "$C$22"

# ---- LIFE EXPECTANCY ANCHORS (rows 24-29) ----
ws.cell(24, 2, "LIFE EXPECTANCY ANCHORS (SSA cohort)").font = font(bold=True, color=WHITE)
ws.cell(24, 2).fill = PatternFill("solid", start_color=MS_BLUE)
ws.cell(24, 3).fill = PatternFill("solid", start_color=MS_BLUE)
le_inputs = [
    ("Current age (for life expectancy)", 49, NUM, "age"),
    ("Remaining life expectancy at current age", 32.6, NUM1, "le"),
    ("Remaining life expectancy at age 62", 22.4, NUM1, "le"),
    ("Remaining life expectancy at age 67", 18.8, NUM1, "le"),
    ("Remaining life expectancy at age 70", 16.7, NUM1, "le"),
]
dv_le2 = DataValidation(type="decimal", operator="between", formula1="0", formula2="60", allow_blank=False)
dv_le2.error = "Enter remaining years between 0 and 60."
ws.add_data_validation(dv_le2)
rr2 = 25
for label, val, fmt, kind in le_inputs:
    ws.cell(rr2, 2, label).font = font()
    input_cell(ws.cell(rr2, 3), val, fmt)
    if kind == "age":
        dv_age.add(f"C{rr2}")
    else:
        dv_le2.add(f"C{rr2}")
    rr2 += 1
cur_age_le = "$C$25"; le_cur = "$C$26"; le62 = "$C$27"; le67 = "$C$28"; le70 = "$C$29"

# ---- SCENARIO DETAIL table (cols H-J): annual returns used by each scenario ----
ws.column_dimensions["H"].width = 8
ws.column_dimensions["I"].width = 13
ws.column_dimensions["J"].width = 13
ws.merge_cells("H5:J5")
ws.cell(5, 8, "SCENARIO DETAIL — annual returns (edit yellow cells)").font = font(bold=True, color=WHITE)
for cc_ in (8, 9, 10):
    ws.cell(5, cc_).fill = PatternFill("solid", start_color=MS_BLUE)
header_cell(ws.cell(6, 8), "Year")
header_cell(ws.cell(6, 9), "Early crash")
header_cell(ws.cell(6, 10), "Stagflation")
crash_vals = {1: -0.20, 2: -0.10, 3: 0.15, 4: 0.15, 5: 0.12}
stag_years = 8
dv_scenret = DataValidation(type="decimal", operator="between", formula1="-0.6", formula2="0.6", allow_blank=False)
dv_scenret.error = "Enter an annual return between -60% and 60%."
ws.add_data_validation(dv_scenret)
for j in range(10):
    yrow = 7 + j
    yv = j + 1
    yc = ws.cell(yrow, 8, yv); yc.font = font(); yc.alignment = Alignment(horizontal="center")
    if yv in crash_vals:
        cc = ws.cell(yrow, 9, crash_vals[yv]); cc.font = font(color=BLACK)
        cc.fill = PatternFill("solid", start_color=YELLOW); cc.protection = UNLOCK; dv_scenret.add(f"I{yrow}")
    else:
        cc = ws.cell(yrow, 9, f"={exp_ret}"); cc.font = font(color="595959")
    cc.number_format = PCT
    if yv <= stag_years:
        sg = ws.cell(yrow, 10, 0.01); sg.font = font(color=BLACK)
        sg.fill = PatternFill("solid", start_color=YELLOW); sg.protection = UNLOCK; dv_scenret.add(f"J{yrow}")
    else:
        sg = ws.cell(yrow, 10, f"={exp_ret}"); sg.font = font(color="595959")
    sg.number_format = PCT
ws.cell(17, 8, "Years beyond 10 use your expected return.").font = font(size=8, italic=True, color="808080")
ws.merge_cells("H17:J17")
SCEN_TBL = f"{S}$H$7:$J$16"

# ---- SOCIAL SECURITY claiming factors (cols H-I, rows 19-30) : born 1960 or later, FRA 67 ----
ws.merge_cells("H19:I19")
ws.cell(19, 8, "SOCIAL SECURITY — % of full benefit by start age").font = font(bold=True, color=WHITE)
for cc_ in (8, 9):
    ws.cell(19, cc_).fill = PatternFill("solid", start_color=MS_BLUE)
header_cell(ws.cell(20, 8), "Start age")
header_cell(ws.cell(20, 9), "% of full")
ss_factors = [(62, 0.700), (63, 0.750), (64, 0.800), (65, 0.867), (66, 0.933),
              (67, 1.000), (68, 1.080), (69, 1.160), (70, 1.240)]
for i, (age_, fac_) in enumerate(ss_factors):
    rr_ = 21 + i
    ac_ = ws.cell(rr_, 8, age_); ac_.font = font(); ac_.alignment = Alignment(horizontal="center")
    fc_ = ws.cell(rr_, 9, fac_); fc_.font = font(color="595959"); fc_.number_format = "0.0%"
    fc_.alignment = Alignment(horizontal="center")
ws.cell(30, 8, "Source: SSA, born 1960 or later (FRA 67). ssa.gov/benefits/retirement/planner/1960.html").font = font(size=8, italic=True, color="808080")
ws.merge_cells("H30:J30")
SS_FACTORS = "$H$21:$I$29"            # Inputs-local
SS_FACTORS_Q = f"{S}$H$21:$I$29"      # sheet-qualified for other tabs
# actual annual Social Security (today's $) = full benefit x claiming factor
ACTUAL_SS = f"{full_ss}*VLOOKUP({ss_age},{SS_FACTORS},2,FALSE)"

# ---- RESULTS (fixed rows 6-18) ----
ws["E6"] = "RESULTS"
ws["E6"].font = font(bold=True, color=WHITE)
ws["E6"].fill = PatternFill("solid", start_color=MS_BLUE)
ws["F6"].fill = PatternFill("solid", start_color=MS_BLUE)

def res(row, label, formula, fmt=CUR, color=GREEN, section=False):
    lc = ws.cell(row, 5, label)
    lc.font = font(bold=True, color=MS_BLUE) if section else font()
    if not section:
        c = ws.cell(row, 6, formula)
        c.font = font(color=color)
        if fmt != "General":
            c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
        c.border = bottom_only
        return c

res(7, "Years modeled (start -> plan-to age)", f"={plan_age}-{start_age}", NUM, BLACK)
res(8, "— 4% RULE —", None, section=True)
# pos = clamp horizon to available rows
POS = f"MIN(MAX({plan_age}-{start_age}+1,1),46)"
res(9, "First-year withdrawal (today's $)", "='4% Rule'!$E$5")
res(10, "Balance at plan-to age (future $)", f"=INDEX('4% Rule'!$H$5:$H$50,{POS})")
res(11, "Money lasts to plan-to age?", f"=IF(INDEX('4% Rule'!$H$5:$H$50,{POS})>0,\"YES\",\"NO\")", "General")
res(12, "— DYNAMIC STRATEGY —", None, section=True)
res(13, "First-year withdrawal (today's $)", "='Dynamic Strategy'!$H$5")
res(14, "Lowest annual withdrawal (future $)", "=MIN('Dynamic Strategy'!$H$5:$H$50)")
res(15, "Highest annual withdrawal (future $)", "=MAX('Dynamic Strategy'!$H$5:$H$50)")
res(16, "Balance at plan-to age (future $)", f"=INDEX('Dynamic Strategy'!$L$5:$L$50,{POS})")
res(17, "Balance at plan-to age (today's $)", f"=INDEX('Dynamic Strategy'!$M$5:$M$50,{POS})")
res(18, "Money lasts to plan-to age?", f"=IF(INDEX('Dynamic Strategy'!$L$5:$L$50,{POS})>0,\"YES\",\"NO\")", "General")
res(19, "Years total spending falls below essentials",
    f"=SUMPRODUCT(('Dynamic Strategy'!$B$5:$B$50<={plan_age})*('Dynamic Strategy'!$P$5:$P$50<{essential}))",
    NUM, BLACK)
# green when 0 such years, red when 1+
ws.conditional_formatting.add("F19", CellIsRule(operator="equal", formula=["0"],
    fill=PatternFill("solid", start_color=GREEN_FILL), font=Font(name=ARIAL, bold=True, color=GREEN_FONT)))
ws.conditional_formatting.add("F19", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
    fill=PatternFill("solid", start_color=RED_FILL), font=Font(name=ARIAL, bold=True, color=RED_FONT)))
res(20, "— ESSENTIALS CHECK —", None, section=True)
res(21, "Essential spending not covered by Social Security, once collecting (today's $)", f"=MAX({essential}-({ACTUAL_SS}),0)", CUR, BLACK)

# NOTE: result row numbers used by the verdict box:
#   4% balance=F10, 4% lasts=F11 ; dyn low=F14, high=F15, lasts=F17
# adjust verdict to match these rows
vcell.value = (
    '="At a starting portfolio of "&TEXT($C$9,"$#,##0")&":     '
    '4% rule "&IF($F$11="YES","lasts to age "&$C$8,"runs out before age "&$C$8)&" (about "&TEXT($F$10,"$#,##0")&" left).     '
    'Dynamic strategy "&IF($F$18="YES","lasts to age "&$C$8,"runs out before age "&$C$8)&", with annual withdrawals of "'
    '&TEXT($F$14,"$#,##0")&" to "&TEXT($F$15,"$#,##0")&"."'
)

# conditional formatting on the two "money lasts" cells
green_yes = CellIsRule(operator="equal", formula=['"YES"'],
                       fill=PatternFill("solid", start_color=GREEN_FILL), font=Font(name=ARIAL, bold=True, color=GREEN_FONT))
red_no = CellIsRule(operator="equal", formula=['"NO"'],
                    fill=PatternFill("solid", start_color=RED_FILL), font=Font(name=ARIAL, bold=True, color=RED_FONT))
for cellref in ("F11", "F18"):
    ws.conditional_formatting.add(cellref, green_yes)
    ws.conditional_formatting.add(cellref, red_no)

# Visual comparison heading (charts are added later, once the strategy tabs exist)
ws.cell(31, 2, "Visual comparison").font = font(bold=True, color="242424", size=12)
ws.cell(64, 2,
        "Educational illustration, not financial advice. Source: WSJ, \"The 4% Rule for Retirement Is Too Simple. "
        "Here's a Better Way.\" Life-expectancy figures are SSA cohort life-table values (born 1960 or later).").font = font(size=8, italic=True, color="808080")
ws.merge_cells(start_row=64, start_column=2, end_row=64, end_column=10)

# ===================================================================
# SHEET: Life Expectancy  (fully calculated from the anchors on Inputs & Summary)
# ===================================================================
le = wb.create_sheet("Life Expectancy")
le.sheet_view.showGridLines = False
title(le, "A1", "Remaining Life Expectancy — calculated", 14)
le["A2"] = ("Every value is interpolated from the Life Expectancy Anchors you set on the Inputs & Summary tab "
            "(current age, 62, 67, 70). Ages beyond 70 or below your current age are extrapolated. Table covers ages 35-110.")
le["A2"].font = font(italic=True, color="595959", size=9)
le["A3"] = "Anchor source: SSA Life Expectancy Calculator (ssa.gov/OACT/population/longevity.html)."
le["A3"].font = font(italic=True, color="808080", size=8)
le.column_dimensions["A"].width = 10
le.column_dimensions["B"].width = 18
le.column_dimensions["C"].width = 30
le.column_dimensions["E"].width = 12
le.column_dimensions["F"].width = 16

header_cell(le.cell(5, 1), "Age")
header_cell(le.cell(5, 2), "Remaining years")
header_cell(le.cell(5, 3), "Basis")

# ---- Anchor table (cols E/F, rows 6-17) — driven by the summary inputs ----
header_cell(le.cell(5, 5), "Anchor age")
header_cell(le.cell(5, 6), "Anchor years")
anchor_ages = [f"={S}{cur_age_le}", 62, 67, 70, 75, 80, 85, 90, 95, 100, 105, 110]
ratios = [None, None, None, None, 0.802, 0.623, 0.467, 0.335, 0.234, 0.162, 0.114, 0.084]
anchor_le = [f"={S}{le_cur}", f"={S}{le62}", f"={S}{le67}", f"={S}{le70}"] + \
            [f"={S}{le70}*{rt}" for rt in ratios[4:]]
for i in range(12):
    ar_ = 6 + i
    ca = le.cell(ar_, 5, anchor_ages[i]); ca.font = font(color=("808080" if i == 0 else BLACK)); ca.number_format = NUM
    cl = le.cell(ar_, 6, anchor_le[i]); cl.font = font(color="808080"); cl.number_format = NUM1
le.cell(18, 5, "Anchors 75+ scale from the age-70 value.").font = font(size=8, italic=True, color="808080")
le.merge_cells("E18:F18")

# ---- Main table (ages 35-110), interpolated by formula ----
le_start = 6
for i, age in enumerate(range(35, 111)):
    rr = le_start + i
    ac = le.cell(rr, 1, age); ac.font = font(); ac.alignment = Alignment(horizontal="center")
    m = f"IFERROR(MATCH(A{rr},$E$6:$E$17,1),1)"
    al = f"INDEX($E$6:$E$17,{m})"
    au = f"INDEX($E$6:$E$17,MIN({m}+1,12))"
    ll = f"INDEX($F$6:$F$17,{m})"
    lu = f"INDEX($F$6:$F$17,MIN({m}+1,12))"
    b = le.cell(rr, 2, f"=ROUND({ll}+({al}<>{au})*(A{rr}-{al})/MAX({au}-{al},1)*({lu}-{ll}),1)")
    b.font = font(); b.number_format = NUM1
    basis = (f'=IF(OR(A{rr}=62,A{rr}=67,A{rr}=70),"Anchor (from inputs)",'
             f'IF(AND(A{rr}>={S}{cur_age_le},A{rr}<=70),"Interpolated (calculated)","Extrapolated (calculated)"))')
    le.cell(rr, 3, basis).font = font(size=9, color="808080")
LE_RANGE = f"'Life Expectancy'!$A${le_start}:$B${le_start + 75}"

# ===================================================================
# SHEET: 4% Rule (year by year, holds shared returns)
# ===================================================================
MAXROWS = 46
fr = wb.create_sheet("4% Rule")
fr.sheet_view.showGridLines = False
title(fr, "A1", "4% Rule — Year by Year", 14)
fr["A2"] = ("Year-1 withdrawal = 4% of the starting portfolio, then grown by inflation. The Annual return "
            "column is calculated from the Market scenario you choose on the Inputs & Summary tab.")
fr["A2"].font = font(italic=True, color="595959", size=9)

cols = ["Year", "Age", "Annual return", "Beginning balance", "Withdrawal",
        "Withdrawal (today's $)", "Investment return", "Ending balance", "Ending balance (today's $)"]
widths = [7, 7, 13, 17, 15, 16, 15, 16, 18]
for i, (cn, w) in enumerate(zip(cols, widths), start=1):
    header_cell(fr.cell(4, i), cn)
    fr.column_dimensions[get_column_letter(i)].width = w

fstart = 5
for k in range(MAXROWS):
    rr = fstart + k
    yr = k + 1
    fr.cell(rr, 1, yr).font = font()
    fr.cell(rr, 2, f"={S}{start_age}+{yr}-1").font = font()
    # scenario-driven return (calculated from the Market scenario on Inputs & Summary)
    ret_formula = (f'=IF({S}{scenario_ref}="Steady",{S}{exp_ret},'
                   f'IF({S}{scenario_ref}="Early crash",IFERROR(VLOOKUP(A{rr},{SCEN_TBL},2,FALSE),{S}{exp_ret}),'
                   f'IFERROR(VLOOKUP(A{rr},{SCEN_TBL},3,FALSE),{S}{exp_ret})))')
    rc = fr.cell(rr, 3, ret_formula)
    rc.font = font(color=GREEN); rc.number_format = PCT
    beg = f"={S}{port0}" if k == 0 else f"=MAX(H{rr-1},0)"
    fr.cell(rr, 4, beg).font = font()
    if k == 0:
        # if already retired, continue your current 4% withdrawal; else 4% of the portfolio
        wd = f"=MIN(IF({S}{current_4pct}>0,{S}{current_4pct},{S}{port0}*{S}{rate4}),D{rr})"
    else:
        wd = f"=MIN(E{rr-1}*(1+{S}{infl}),D{rr})"
    fr.cell(rr, 5, wd).font = font()
    fr.cell(rr, 6, f"=E{rr}/(1+{S}{infl})^({yr}-1)").font = font(color="595959")
    fr.cell(rr, 7, f"=(D{rr}-E{rr})*C{rr}").font = font()
    fr.cell(rr, 8, f"=D{rr}-E{rr}+G{rr}").font = font()
    fr.cell(rr, 9, f"=H{rr}/(1+{S}{infl})^({yr}-1)").font = font(color="595959")
    for col in range(1, 10):
        cell = fr.cell(rr, col)
        if col in (4, 5, 6, 7, 8, 9):
            cell.number_format = CUR
        if k % 2 == 1:
            cell.fill = PatternFill("solid", start_color=LIGHT_GREY)
flast = fstart + MAXROWS - 1
fr.conditional_formatting.add(f"H{fstart}:H{flast}", red_rule)

# ===================================================================
# SHEET: Dynamic Strategy (year by year)
# ===================================================================
dy = wb.create_sheet("Dynamic Strategy")
dy.sheet_view.showGridLines = False
title(dy, "A1", "Dynamic Strategy — Year by Year", 14)
dy["A2"] = ("Withdrawal = balance / remaining years, kept inside the guardrails, then limited by the shock absorber. "
            "Returns are linked to the 4% Rule tab for an apples-to-apples comparison.")
dy["A2"].font = font(italic=True, color="595959", size=9)

cols2 = ["Year", "Age", "Remaining years", "Annual return", "Beginning balance",
         "Life-exp. withdrawal", "After guardrails", "Final withdrawal", "Final withdrawal (today's $)",
         "Implied rate", "Investment return", "Ending balance", "Ending balance (today's $)",
         "Social Security", "Total spending", "Total spending (today's $)"]
notes2 = [
    "Year of retirement, starting at 1. Year 1 is your retirement start age.",
    "Your age in this year (retirement start age + year - 1).",
    "Your remaining life expectancy at this age, calculated from the anchors on the Inputs tab. The dynamic withdrawal divides your balance by this number.",
    "The investment return applied this year. It is linked from the 4% Rule tab so both strategies see the same markets, and is set by your chosen Market scenario.",
    "Portfolio value at the start of the year, before any withdrawal. It equals last year's ending balance.",
    "The raw actuarial withdrawal: beginning balance divided by remaining years. This is before any guardrails or smoothing.",
    "The life-expectancy withdrawal after being capped inside your floor/ceiling band (for example 3% to 6% of the balance).",
    "The actual amount withdrawn this year, after the shock absorber limits how much your spending can change from last year.",
    "The final withdrawal restated in today's buying power, so you can compare it fairly across years despite inflation.",
    "The final withdrawal as a percent of the beginning balance. It shows where you landed inside the guardrail band.",
    "Dollar growth or loss on the portfolio this year, applied after the withdrawal is taken out.",
    "Portfolio value at year end: beginning balance minus withdrawal plus investment return. It carries into next year.",
    "The ending balance restated in today's buying power.",
    "Guaranteed income this year. It is zero until your Social Security start age, then your full benefit times the claiming factor, grown by inflation.",
    "What you actually get to spend: portfolio withdrawal plus Social Security, in nominal future dollars.",
    "Total spending in today's buying power. It turns red in any year it falls below your essential spending, a sign the plan underfunds your must-haves that year.",
]
widths2 = [7, 7, 12, 12, 16, 15, 14, 14, 17, 10, 14, 16, 18, 14, 14, 18]
for i, (cn, w) in enumerate(zip(cols2, widths2), start=1):
    hc = dy.cell(4, i)
    header_cell(hc, cn)
    cm = Comment(notes2[i - 1], "Planner")
    cm.width = 260; cm.height = 120
    hc.comment = cm
    dy.column_dimensions[get_column_letter(i)].width = w

dstart = 5
for k in range(MAXROWS):
    rr = dstart + k
    yr = k + 1
    dy.cell(rr, 1, yr).font = font()
    dy.cell(rr, 2, f"={S}{start_age}+{yr}-1").font = font()
    dy.cell(rr, 3, f"=IFERROR(VLOOKUP(B{rr},{LE_RANGE},2,FALSE),1.5)").font = font()
    dy.cell(rr, 4, f"='4% Rule'!C{rr}").font = font(color=GREEN); dy.cell(rr, 4).number_format = PCT
    beg = f"={S}{port0}" if k == 0 else f"=MAX(L{rr-1},0)"
    dy.cell(rr, 5, beg).font = font()
    dy.cell(rr, 6, f"=E{rr}/MAX(C{rr},1)").font = font()
    dy.cell(rr, 7, f"=MIN({S}{ceiling}*E{rr},MAX({S}{floor}*E{rr},F{rr}))").font = font()
    if k == 0:
        fw = f"=MIN(G{rr},E{rr})"
    else:
        fw = f"=MIN(MIN(H{rr-1}*(1+{S}{shock}),MAX(H{rr-1}*(1-{S}{shock}),G{rr})),E{rr})"
    dy.cell(rr, 8, fw).font = font(bold=True)
    dy.cell(rr, 9, f"=H{rr}/(1+{S}{infl})^({yr}-1)").font = font(color="595959")
    dy.cell(rr, 10, f"=IFERROR(H{rr}/E{rr},0)").font = font(); dy.cell(rr, 10).number_format = PCT2
    dy.cell(rr, 11, f"=(E{rr}-H{rr})*D{rr}").font = font()
    dy.cell(rr, 12, f"=E{rr}-H{rr}+K{rr}").font = font()
    dy.cell(rr, 13, f"=L{rr}/(1+{S}{infl})^({yr}-1)").font = font(color="595959")
    dy.cell(rr, 14, f"=IF(B{rr}>={S}{ss_age},{S}{full_ss}*VLOOKUP({S}{ss_age},{SS_FACTORS_Q},2,FALSE)*(1+{S}{infl})^({yr}-1),0)").font = font(color="808080")
    dy.cell(rr, 15, f"=H{rr}+N{rr}").font = font()
    dy.cell(rr, 16, f"=O{rr}/(1+{S}{infl})^({yr}-1)").font = font(color="595959")
    for col in range(1, 17):
        cell = dy.cell(rr, col)
        if col in (5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16):
            cell.number_format = CUR
        if k % 2 == 1:
            cell.fill = PatternFill("solid", start_color=LIGHT_GREY)
dlast = dstart + MAXROWS - 1
dy.conditional_formatting.add(f"L{dstart}:L{dlast}", red_rule)
# red when total spending (today's $) drops below essentials, within the plan horizon
below_ess_rule = FormulaRule(formula=[f"AND(P{dstart}<{S}{essential},B{dstart}<={S}{plan_age})"],
                             fill=PatternFill("solid", start_color=RED_FILL),
                             font=Font(name=ARIAL, color=RED_FONT))
dy.conditional_formatting.add(f"P{dstart}:P{dlast}", below_ess_rule)

# ===================================================================
# CHARTS (placed on the Inputs & Summary tab)
# ===================================================================
# Chart 1: portfolio balance over time (nominal)
c1 = LineChart()
c1.title = "Portfolio balance over time"
c1.height = 8; c1.width = 17
c1.y_axis.title = "Balance ($)"
c1.x_axis.title = "Age"
c1.x_axis.delete = False; c1.y_axis.delete = False
s1 = Series(Reference(fr, min_col=8, min_row=fstart, max_row=flast), title="4% rule")
s2 = Series(Reference(dy, min_col=12, min_row=dstart, max_row=dlast), title="Dynamic strategy")
c1.series.append(s1); c1.series.append(s2)
c1.set_categories(Reference(fr, min_col=2, min_row=fstart, max_row=flast))
ws.add_chart(c1, "B32")

# Chart 2: annual portfolio withdrawal in today's dollars
c2 = LineChart()
c2.title = "Annual portfolio withdrawal (today's dollars)"
c2.height = 8; c2.width = 17
c2.y_axis.title = "Withdrawal ($, today's buying power)"
c2.x_axis.title = "Age"
c2.x_axis.delete = False; c2.y_axis.delete = False
s3 = Series(Reference(fr, min_col=6, min_row=fstart, max_row=flast), title="4% rule")
s4 = Series(Reference(dy, min_col=9, min_row=dstart, max_row=dlast), title="Dynamic strategy")
c2.series.append(s3); c2.series.append(s4)
c2.set_categories(Reference(fr, min_col=2, min_row=fstart, max_row=flast))
ws.add_chart(c2, "B48")

# ===================================================================
# SHEET: Article (narrative)
# ===================================================================
art = wb.create_sheet("Article")
art.sheet_view.showGridLines = False
art.column_dimensions["A"].width = 3
art.column_dimensions["B"].width = 110

article = [
    ("h1", "The 4% Rule for Retirement Is Too Simple. Here's a Better Way."),
    ("src", "The Wall Street Journal · Personal Finance / Retirement"),
    ("url", "https://www.wsj.com/personal-finance/retirement/retirement-planning-4-percent-rule-d5ba3a0b"),
    ("p", "Most retirees are familiar with the 4% rule."),
    ("p", "It's the guide many financial pros have extolled for decades for the decumulation phase of "
          "retirement – the time when you finally start tapping the money you've spent decades (ideally) "
          "saving and investing."),
    ("p", "Withdraw 4% of your portfolio in the first year of retirement, then adjust that amount for "
          "inflation every year thereafter. If your portfolio starts out at $1 million, for instance, you "
          "spend $40,000 the first year, then $40,000 plus inflation the next year, and so on."),
    ("p", "The rule was devised in 1994 by financial adviser William Bengen after research showed that "
          "retirees with a balanced stock-and-bond portfolio who followed that math wouldn't have run out "
          "of money over any 30-year period since 1926 – even when economic conditions were bad."),
    ("p", "It's pretty simple. Maybe too simple."),
    ("p", "The 4% rule has some blind spots. It assumes a 30-year retirement, but some retirees will need "
          "their money to last longer than that. It also assumes the markets will perform as well as they "
          "have over the past century. As a result, researchers have been lowering or raising the figure "
          "for years. Investment research firm Morningstar revises its estimate annually; it was 3.3% in "
          "2021 and 3.9% this year. Bengen himself has said that a safe number could now be 4.7%."),
    ("p", "The rule's deeper problem is its rigidity. Your portfolio amount could double in a market boom "
          "during retirement and the rule would still limit you to the same inflation-adjusted amount, "
          "which would be unnecessarily frugal. More concerning, your assets could fall by a third and the "
          "rule wouldn't tell you to curtail your spending. You would have to liquidate a bigger portion of "
          "your portfolio to withdraw the prescribed amount. If that happens a few years in a row, you run "
          "a real risk of eventually outliving your money, despite what the historical data say."),
    ("p", "The good news is that there are ways to make retirement spending more responsive to how long "
          "you're likely to live and how markets perform. Use the 4% rule as a reference point, rather than "
          "have it be the full plan."),
    ("h2", "Your life expectancy matters"),
    ("p", "Instead of spending a fixed inflation-adjusted dollar amount every year, you need to factor in "
          "your life expectancy and the size of your portfolio and then adjust accordingly."),
    ("p", "You divide your current portfolio balance by the number of years you can expect to live, and "
          "spend that much in the coming year. The following year, you do the math again with your new "
          "portfolio balance and your new remaining life expectancy. (You can find your life expectancy at "
          "your current age in the Social Security Administration's life tables, which are updated annually. "
          "Keep in mind that tables don't apply to any specific person; they don't factor in health issues "
          "or family medical history. You may want to be conservative and go beyond your life expectancy, "
          "or take a chance and reduce it.)"),
    ("p", "If you're a 70-year-old woman, for instance, your remaining life expectancy is about 16 years. "
          "So you would divide your portfolio into sixteenths and spend one of them this year. A year later, "
          "you would recalculate using your new balance and your new life expectancy of 15 years, and so on. "
          "This allows your spending to rise when your portfolio is up and fall when it's down."),
    ("p", "This actuarial approach lessens the chance you'll run out of money, and it lets good market "
          "returns flow through to higher spending."),
    ("h2", "Spending too much – or not enough"),
    ("p", "But it introduces a new problem: too much fluctuation in your spending."),
    ("p", "If markets drop 30% in a bad year, your spending drops by roughly 30%, too. If markets surge in "
          "a year when the life tables say your remaining expectancy is five years, the formula tells you to "
          "spend a fifth of your portfolio – probably far more than you need or want."),
    ("p", "The fix is to put guardrails on the actuarial approach, using the 4% rule as your anchor. You set "
          "an upper and lower band around it – say, 3% of assets on the low end and 6% on the high end. Each "
          "year, you do the actuarial calculation and check the result against the band."),
    ("p", "If the implied withdrawal rate falls inside the guardrails, you're fine. If your spending would "
          "exceed 6% of your portfolio, you cap your spending at 6%. If your spending would fall below 3%, "
          "you spend at least 3%. The guardrails work like a financial thermostat. The 4% rule sits near the "
          "center of the band as the familiar reference point most retirees already know."),
    ("p", "But if markets fluctuate a lot, you could still be in for a shock. Say markets fall sharply and "
          "the guardrails tell you to cut spending by a quarter to get back inside the band. That could be "
          "wrenching for a lot of retirees."),
    ("p", "So here's a third fix: a shock absorber, like the ones some actual roadside guardrails have."),
    ("p", "You cap year-to-year spending changes at, say, plus or minus 5%. Even if the guardrails say to "
          "cut more aggressively, you never cut by more than 5% in a single year. Same on the upside. Your "
          "spending won't swing chaotically between feast and famine."),
    ("h2", "Essential vs. discretionary"),
    ("p", "The trade-off for the shock absorption is that in a deep, sustained downturn, your portfolio will "
          "be exposed to more drawdown than a stricter rule would allow. So the shock absorber works best "
          "for retirees with some cushion above their essential spending."),
    ("p", "That points to the next refinement."),
    ("p", "None of the changes so far distinguish between the spending you have to do and the spending you "
          "would like to do. If you feel any of these strategies leaves you too vulnerable to market swings "
          "or a surprisingly long life, you can calculate the expenses you absolutely need to cover – food, "
          "healthcare, rent and the like – and pay for those with guaranteed income such as Social Security "
          "benefits or an annuity."),
    ("p", "Then you can apply a dynamic withdrawal strategy, with the guardrails and shock absorber "
          "discussed above, to the spending that remains. You can afford to take more risk, and probably "
          "earn higher returns, with your discretionary portfolio because you know the essentials are "
          "covered."),
    ("p", "None of these refinements is perfect. Each retiree has personal factors that could change the "
          "math, and the account order above will have exceptions in particular tax situations. There are "
          "software tools that run the calculations back and forth until they converge on an answer for a "
          "given person's circumstances. Realistically, though, most people won't use such tools routinely, "
          "if at all."),
    ("p", "A dynamic strategy – with the 4% rule as a reference point – is practicable without specialized "
          "software. And it should help you enjoy your nest egg and not outlive it."),
]
ar = 2
for kind, text in article:
    if kind == "h2":
        ar += 1  # blank row before each section
    c = art.cell(ar, 2, text)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "h1":
        c.font = font(bold=True, color="242424", size=18); art.row_dimensions[ar].height = 26
    elif kind == "src":
        c.font = font(italic=True, color="808080", size=9)
    elif kind == "url":
        c.font = Font(name=ARIAL, color="0563C1", size=10, underline="single")
        c.hyperlink = text
        ar += 1  # blank row between the URL and the article body
    elif kind == "h2":
        c.font = font(bold=True, color=MS_BLUE, size=13); art.row_dimensions[ar].height = 22
    else:
        c.font = font(color="242424", size=11)
        lines = max(1, (len(text) // 105) + 1)
        art.row_dimensions[ar].height = 15 * lines + 4
    ar += 1

# ===================================================================
# SHEET: Instructions
# ===================================================================
ins = wb.create_sheet("Instructions")
ins.sheet_view.showGridLines = False
ins.column_dimensions["A"].width = 3
ins.column_dimensions["B"].width = 110

instructions = [
    ("h1", "How to Use This Planner"),
    ("p", "This workbook compares two ways to spend down a retirement portfolio: the classic 4% rule and a "
          "more flexible \"dynamic\" strategy from the Wall Street Journal article (see the Article tab). "
          "It is an educational illustration, not financial advice."),
    ("h2", "How the strategies work"),
    ("p", "4% rule — withdraw 4% of the starting portfolio in year one, then raise that dollar amount by "
          "inflation every year (created by William Bengen, 1994). Simple, but rigid."),
    ("p", "Dynamic strategy — three layered fixes the article recommends:"),
    ("p", "   1) Life expectancy: each year divide your balance by your remaining years and spend that share."),
    ("p", "   2) Guardrails: keep the result inside a band (floor / ceiling) so spending never gets extreme."),
    ("p", "   3) Shock absorber: never change spending by more than a set % in a single year."),
    ("p", "Essentials vs. discretionary — cover must-haves with guaranteed income; apply the dynamic "
          "withdrawals to the rest. The Essentials Check shows how much of your essentials the portfolio "
          "must still cover."),
    ("h2", "Quick start"),
    ("p", "1. Open the Inputs & Summary tab."),
    ("p", "2. Change any of the yellow cells to match your situation (age, portfolio size, returns, etc.)."),
    ("p", "3. Read the blue verdict box at the top and the Results panel on the right — they update automatically."),
    ("p", "4. Scroll down on that tab to see the two charts compare the strategies side by side."),
    ("h2", "What each tab does"),
    ("p", "• Inputs & Summary — your assumptions, the market scenario, life-expectancy anchors, results, and charts."),
    ("p", "• Dynamic Strategy — life-expectancy + guardrails + shock absorber, year by year."),
    ("p", "• 4% Rule — the classic strategy, year by year. Also holds the shared annual returns."),
    ("p", "• Life Expectancy — remaining-years table, fully calculated from your anchor inputs."),
    ("p", "• Article — the full WSJ narrative the model is based on."),
    ("h2", "Reading the numbers: nominal vs. today's dollars"),
    ("p", "Most columns are \"nominal\" (actual future dollars). Because of inflation, a $90,000 withdrawal "
          "30 years from now buys far less than $90,000 today. The columns labeled \"(today's $)\" restate "
          "those amounts in today's buying power, so you can compare spending fairly across the years."),
    ("h2", "Testing market scenarios"),
    ("p", "On the Inputs & Summary tab, use the Market scenario dropdown to switch between three return paths:"),
    ("p", "   • Steady — every year earns your expected return."),
    ("p", "   • Early crash — a sharp drop in the first couple of retirement years, then recovery. This tests "
          "\"sequence-of-returns\" risk, the most dangerous case: bad markets right as you start withdrawing."),
    ("p", "   • Stagflation — several years of very low returns. With inflation still eroding value, this squeezes "
          "real spending."),
    ("p", "You can fine-tune the exact return numbers for each scenario in the yellow Scenario detail "
          "table on the Inputs & Summary tab. The Annual return column on the 4% Rule tab is calculated "
          "from your scenario choice; both strategies share it, so the comparison stays fair."),
    ("h2", "Life-expectancy anchors"),
    ("p", "The dynamic strategy needs your remaining life expectancy at each age. On the Inputs & Summary tab "
          "you enter just four anchor values (at your current age, 62, 67 and 70 — from the SSA calculator). "
          "The Life Expectancy tab then calculates every other age automatically: it interpolates between your "
          "anchors and extrapolates beyond age 70 from the age-70 value. Update the anchors and the whole table "
          "and both strategies refresh."),
    ("h2", "Social Security"),
    ("p", "Enter your full Social Security benefit (the amount at your full retirement age of 67, in today's "
          "dollars per year) and the age you plan to start collecting (62 to 70). The model looks up the SSA "
          "claiming factor — 70% of full at 62, 100% at 67, up to 124% at 70 (born 1960 or later) — and pays "
          "that income only from your start age onward, growing with inflation. Before your start age, the "
          "model assumes no Social Security."),
    ("p", "The claiming-factor table sits on the Inputs & Summary tab. Source: SSA, "
          "ssa.gov/benefits/retirement/planner/1960.html."),
    ("h2", "Already retired? Using this mid-stream"),
    ("p", "The dynamic strategy is designed to be re-run from today every year, so it works mid-retirement. "
          "Set the Retirement start age to your current age and the Starting portfolio to your current "
          "balance (both inputs are labeled for this). Year 1 of the tables then becomes next year, and the "
          "dynamic first-year withdrawal is your suggested amount for the coming year. Come back annually, "
          "update your age and balance, and read the new figure."),
    ("p", "One nuance for the 4% Rule column: by default it restarts at 4% of your current balance. If you "
          "want it to continue your original, inflation-adjusted 4% amount instead, enter that dollar amount "
          "in \"Current 4% rule withdrawal if already retired\" — leave it 0 if you are not yet retired."),
    ("h2", "Today's dollars vs. future dollars in the Results"),
    ("p", "Result rows are labeled either (today's $) or (future $). First-year withdrawals are in today's "
          "dollars (year 1 is now). Balances at plan-to age are shown both ways. The lowest/highest annual "
          "withdrawals are future (nominal) dollars, so later years look larger purely because of inflation."),
    ("h2", "Why the spreadsheet is locked"),
    ("p", "To prevent accidental edits, every tab is protected so only the yellow input cells (and the return "
          "and scenario cells) can be changed. The formulas are locked."),
    ("p", "To unlock a tab so you can edit formulas:"),
    ("p", "   1. Click the tab you want to change."),
    ("p", "   2. Go to the Review ribbon in Excel."),
    ("p", "   3. Click \"Unprotect Sheet\" (there is no password — it unlocks immediately)."),
    ("p", "To turn protection back on: Review ribbon → \"Protect Sheet\" → OK."),
    ("h2", "Color key"),
    ("p", "• Yellow cells = inputs you can edit."),
    ("p", "• Black text = calculated values."),
    ("p", "• Green text = pulled from another tab."),
    ("p", "• Grey text = shown in today's dollars."),
    ("p", "• Red highlight = a balance has run out (reached $0)."),
    ("h2", "A note on assumptions"),
    ("p", "Defaults: 6% expected return (nominal), 3% inflation, 4% rule rate, 3%–6% guardrails, 5% shock "
          "absorber. A 3% inflation assumption is reasonable and slightly conservative (long-run US "
          "inflation has averaged about 2.5–3%). Adjust any of these to your own outlook."),
]
ir2 = 2
for kind, text in instructions:
    if kind == "h2":
        ir2 += 1  # blank row before each section header
    c = ins.cell(ir2, 2, text)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "h1":
        c.font = font(bold=True, color="242424", size=18); ins.row_dimensions[ir2].height = 26
    elif kind == "h2":
        c.font = font(bold=True, color=MS_BLUE, size=13); ins.row_dimensions[ir2].height = 22
    else:
        c.font = font(color="242424", size=11)
        lines = max(1, (len(text) // 105) + 1)
        ins.row_dimensions[ir2].height = 15 * lines + 4
    ir2 += 1

# ===================================================================
# Tab order: Instructions, Inputs & Summary, Article, Charts, Life Expectancy, 4% Rule, Dynamic
# ===================================================================
order = ["Instructions", "Inputs & Summary", "Dynamic Strategy", "4% Rule", "Life Expectancy", "Article"]
wb._sheets.sort(key=lambda s: order.index(s.title))

# ===================================================================
# Enable protection on every sheet (no password; input cells already unlocked)
# ===================================================================
for sh in wb.worksheets:
    sh.protection.sheet = True
    sh.protection.selectLockedCells = False
    sh.protection.selectUnlockedCells = False
    sh.protection.formatCells = False

import os
# generic document properties (no personal information)
wb.properties.creator = "Retirement Planner"
wb.properties.lastModifiedBy = "Retirement Planner"
wb.properties.title = "Retirement Withdrawal Strategies Planner"
wb.properties.description = "Educational 4% rule vs. dynamic withdrawal planner."
wb.properties.keywords = None
wb.properties.company = None
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Retirement Withdrawal Strategies Planner.xlsx")
wb.save(out)
print("saved:", out)
