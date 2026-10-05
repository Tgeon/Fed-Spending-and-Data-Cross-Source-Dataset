#!/usr/bin/env python3
"""Build the Excel filter workbook from the exported cubes, profile and dictionary.
Usage: python3 build_filter_workbook.py <build_inputs_dir> <out.xlsx>"""
import sys, json, csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.worksheet.properties import PageSetupProperties

IN, OUTF = sys.argv[1], sys.argv[2]
OVR = json.loads(sys.argv[3]) if len(sys.argv) > 3 else {}
FONT = "Times New Roman"; NAVY = "1E5631"; MIDG = "548235"; BG = "E2F0D9"; LINKC = "006100"; BARC = "2E7D32"
lists = json.load(open(IN + "/lists.json")); prof = json.load(open(IN + "/profile.json")); dic = json.load(open(IN + "/dictionary.json"))
def num(x):
    try: return int(x)
    except ValueError: return float(x)
def readcsv(p, numcols):
    with open(p, newline="") as f:
        r = csv.reader(f); hdr = next(r)
        return hdr, [[num(x) if i in numcols else x for i, x in enumerate(row)] for row in r]
mh, main = readcsv(IN + "/cube_main.csv", {1, 4, 5, 6}); nh, month = readcsv(IN + "/cube_month.csv", {1, 2, 5, 6})
for c in lists["categories"]:
    assert not any(ch in c for ch in "*?~"), c
N1, N2 = len(main) + 1, len(month) + 1   # last data row of each cube
SRC, YEARS, STATES, CATS, BANDS = lists["sources"], lists["years"], lists["states"], lists["categories"], lists["bands"]

wb = Workbook()
wb._named_styles["Normal"].font = Font(name=FONT, size=11)
NF_N = "#,##0"; NF_D = "$#,##0;($#,##0);-"; NF_P = "0.0%"
thin = Side(style="thin", color="BFBFBF")
def put(ws, a, v, kind="lbl", nf=None, wrap=False, h=None, bold=None):
    c = ws[a]; c.value = v
    f = dict(name=FONT, size=11)
    if kind == "title": f.update(size=16, bold=True, color=NAVY)
    elif kind == "sub": f.update(italic=True, color="3B3B3B", size=10)
    elif kind == "sec": f.update(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
    elif kind == "hdr": f.update(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=MIDG); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    elif kind == "in": f.update(color="0000FF"); c.fill = PatternFill("solid", fgColor="FFF2CC"); c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    elif kind == "inp": f.update(color="0000FF")
    elif kind == "link": f.update(color=LINKC)
    elif kind == "note": f.update(italic=True, color="3B3B3B", size=10)
    elif kind == "warn": f.update(bold=True, color="C00000")
    elif kind == "b": f.update(bold=True)
    if bold: f["bold"] = True
    c.font = Font(**f)
    if nf: c.number_format = nf
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    if h: c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return c
def sec(ws, row, text, c1=1, c2=4):
    for c in range(c1, c2 + 1): put(ws, L(c) + str(row), None, "sec")
    put(ws, L(c1) + str(row), text, "sec")
def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[L(i)].width = x

wb.properties.title = "Cross-Source Dataset: Federal Spending and Financial Data, FY1998 to FY2026"
ws_readme = wb.active; ws_readme.title = "README"
ws = wb.create_sheet("Explorer"); ws_ext = wb.create_sheet("Extract_Request"); ws_prof = wb.create_sheet("Profile")
ws_dic = wb.create_sheet("Data_Dictionary"); ws_l = wb.create_sheet("Lists"); ws_calc = wb.create_sheet("Calc")
ws_cm = wb.create_sheet("Cube_Main"); ws_cmo = wb.create_sheet("Cube_Month")

# ------------------------------------------------------------------ cube data sheets
def datasheet(w, hdr, rows):
    w.append(hdr)
    for r in rows: w.append(r)
    for i in range(1, len(hdr) + 1): put(w, L(i) + "1", hdr[i - 1], "hdr")
    w.freeze_panes = "A2"; w.auto_filter.ref = "A1:" + L(len(hdr)) + str(len(rows) + 1)
datasheet(ws_cm, ["Source", "Year", "State", "Category", "Band", "Records", "Amount"], main)
datasheet(ws_cmo, ["Source", "Year", "Month", "State", "Category", "Records", "Amount"], month)
widths(ws_cm, [14, 8, 9, 56, 7, 11, 20]); widths(ws_cmo, [14, 8, 8, 9, 56, 11, 20])

# ------------------------------------------------------------------ Lists
widths(ws_l, [14, 8, 10, 56, 18, 8, 12, 14, 14])
put(ws_l, "A1", "Sources", "hdr"); put(ws_l, "B1", "Years", "hdr"); put(ws_l, "C1", "States", "hdr"); put(ws_l, "D1", "Categories", "hdr")
put(ws_l, "E1", "Amount band", "hdr"); put(ws_l, "F1", "Band id", "hdr"); put(ws_l, "G1", "Rank by", "hdr"); put(ws_l, "H1", "Band lower ($)", "hdr"); put(ws_l, "I1", "Band upper ($)", "hdr")
for i, v in enumerate(["All"] + SRC): put(ws_l, "A" + str(i + 2), v, "inp")
for i, v in enumerate(YEARS): put(ws_l, "B" + str(i + 2), v, "inp")
for i, v in enumerate(["All"] + STATES): put(ws_l, "C" + str(i + 2), v, "inp")
for i, v in enumerate(["All"] + CATS): put(ws_l, "D" + str(i + 2), v, "inp")
LO = [None, None, 0, 0.01, 10000, 100000, 1000000, 10000000, 100000000]; HI = [None, -0.01, 0, 9999.99, 99999.99, 999999.99, 9999999.99, 99999999.99, None]
for i, b in enumerate(BANDS):
    r = i + 2; put(ws_l, "E" + str(r), b, "inp"); put(ws_l, "F" + str(r), i, "inp")
    if LO[i] is not None: put(ws_l, "H" + str(r), LO[i], "inp", "#,##0.00")
    if HI[i] is not None: put(ws_l, "I" + str(r), HI[i], "inp", "#,##0.00")
put(ws_l, "G2", "Records", "inp"); put(ws_l, "G3", "Amount", "inp")
put(ws_l, "K1", "Dropdown lists and band limits. Do not edit.", "note")
NSRC, NY, NST, NCAT, NB = len(SRC) + 1, len(YEARS), len(STATES) + 1, len(CATS) + 1, len(BANDS)

# ------------------------------------------------------------------ Explorer controls
widths(ws, [42, 22, 24, 12, 2] + [11] * 14)
put(ws, "A1", "Cross-Source Data Explorer", "title")
put(ws, "A2", "Change the yellow cells.", "sub")
sec(ws, 4, "Filters", 1, 4)
ctrl = [("Source", "All", "Dataset, or All"), ("Year from", YEARS[0], "Calendar year"), ("Year to", YEARS[-1], ""),
        ("State", "All", "Two-letter code"), ("Category", "All", ""),
        ("Amount band from", BANDS[0], "Coarse size filter"), ("Amount band to", BANDS[-1], ""),
        ("Rank top 10 lists by", "Records", "Records or Amount"), ("Exact minimum amount ($)", None, "Optional; Extract_Request only"),
        ("Exact maximum amount ($)", None, "Optional; Extract_Request only")]
for i, (lab, val, note) in enumerate(ctrl):
    r = 5 + i; val = OVR.get("B" + str(r), val); put(ws, "A" + str(r), lab, "b"); put(ws, "B" + str(r), val, "in", "#,##0.00" if i >= 8 else None, h="left"); put(ws, "C" + str(r), note, "note")
def dv(formula, cell, err=True):
    d = DataValidation(type="list", formula1=formula, allow_blank=False, showErrorMessage=err); ws.add_data_validation(d); d.add(cell)
dv("=Lists!$A$2:$A$" + str(NSRC + 1), "B5"); dv("=Lists!$B$2:$B$" + str(NY + 1), "B6"); dv("=Lists!$B$2:$B$" + str(NY + 1), "B7")
dv("=Lists!$C$2:$C$" + str(NST + 1), "B8"); dv("=Lists!$D$2:$D$" + str(NCAT + 1), "B9"); dv("=Lists!$E$2:$E$" + str(NB + 1), "B10"); dv("=Lists!$E$2:$E$" + str(NB + 1), "B11"); dv("=Lists!$G$2:$G$3", "B12")
dn = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="-1000000000000", allow_blank=True); ws.add_data_validation(dn); dn.add("B13"); dn.add("B14")
put(ws, "C7", '=IF(B6>B7,"Year from is after year to","")', "warn")
put(ws, "C9", '=IF(B9="All","",IF(AND(B5<>"All",LEFT(B9,FIND(" | ",B9)-1)<>B5),"Other source: results will be zero","Source: "&LEFT(B9,FIND(" | ",B9)-1)))', "note")
ws["C11"].value = '=IF(MATCH(B10,Lists!$E$2:$E$' + str(NB + 1) + ',0)>MATCH(B11,Lists!$E$2:$E$' + str(NB + 1) + ',0),"Band from is above band to","")'; ws["C11"].font = Font(name=FONT, size=11, bold=True, color="C00000")

# ------------------------------------------------------------------ Calc: criteria
widths(ws_calc, [30, 22, 3, 18, 14, 22, 16, 3, 56, 14, 22, 16])
put(ws_calc, "A1", "Helper cells (do not edit)", "sec"); put(ws_calc, "B1", None, "sec")
crit = [("f_src", '=IF(Explorer!B5="All","*",Explorer!B5)'), ("f_state", '=IF(Explorer!B8="All","*",Explorer!B8)'), ("f_cat", '=IF(Explorer!B9="All","*",Explorer!B9)'),
        ("f_blo", '=MATCH(Explorer!B10,Lists!$E$2:$E$' + str(NB + 1) + ',0)-1'), ("f_bhi", '=MATCH(Explorer!B11,Lists!$E$2:$E$' + str(NB + 1) + ',0)-1'),
        ("f_ylo", "=Explorer!B6"), ("f_yhi", "=Explorer!B7"),
        ("f_srcamt", '=IF(AND(Explorer!B5="All",Explorer!B9="All"),"<>SEC",f_src)'), ("f_metric", "=Explorer!B12"), ("xl_rows", 1048575), ("tiebreak", 0.000001)]
for i, (nme, fm) in enumerate(crit):
    r = i + 2; put(ws_calc, "A" + str(r), nme, "lbl"); put(ws_calc, "B" + str(r), fm, "link" if isinstance(fm, str) else "inp", None if isinstance(fm, str) else ("#,##0" if fm > 1 else "0.000000"))
    wb.defined_names[nme] = DefinedName(nme, attr_text="Calc!$B$" + str(r))
put(ws_calc, "A13", "f_srcamt: leaves SEC out of dollar totals when Source and Category are All. xl_rows: Excel row limit. tiebreak: separates equal ranks.", "note")

# ------------------------------------------------------------------ SUMIFS builders
def M(col): return "Cube_Main!$" + col + "$2:$" + col + "$" + str(N1)
def O(col): return "Cube_Month!$" + col + "$2:$" + col + "$" + str(N2)
def sumifs(cube, valcol, dims):
    """dims: dict dim -> 'std' | ('eq', cellref) | None(skip)"""
    if cube == "main": rng, cols = M, dict(src="A", year="B", state="C", cat="D", band="E")
    else: rng, cols = O, dict(src="A", year="B", month="C", state="D", cat="E")
    parts = [rng(valcol)]
    def add(dim, expr): parts.extend([rng(cols[dim]), expr])
    d = dims
    s = d.get("src", "std"); add("src", "f_src" if s == "std" else ("f_srcamt" if s == "amt" else s[1])) if s is not None else None
    y = d.get("year", "std")
    if y == "std": add("year", '">="&f_ylo'); add("year", '"<="&f_yhi')
    elif y is not None: add("year", y[1])
    st = d.get("state", "std")
    if st is not None: add("state", "f_state" if st == "std" else st[1])
    c = d.get("cat", "std")
    if c is not None: add("cat", "f_cat" if c == "std" else c[1])
    if cube == "main":
        b = d.get("band", "std")
        if b == "std": add("band", '">="&f_blo'); add("band", '"<="&f_bhi')
        elif b is not None: add("band", b[1])
    else:
        mo = d.get("month")
        if mo is not None: add("month", mo[1])
    return "SUMIFS(" + ",".join(parts) + ")"
def rec(dims, cube="main"): return sumifs(cube, "F", dims)
def amt(dims, cube="main"):
    dims = dict(dims)
    if dims.get("src", "std") == "std": dims["src"] = "amt"
    return sumifs(cube, "G", dims)

# ------------------------------------------------------------------ Explorer results
sec(ws, 16, "Results", 1, 4)
put(ws, "A17", "Records", "b"); put(ws, "B17", "=" + rec({}), "f", NF_N)
put(ws, "A18", "Records in dollar total", "lbl"); put(ws, "B18", "=" + sumifs("main", "F", dict(src="amt")), "f", NF_N)
put(ws, "A19", "Total amount ($)", "b"); put(ws, "B19", "=" + amt({}), "f", NF_D)
put(ws, "A20", "Average amount per record ($)", "lbl"); put(ws, "B20", "=IF(B18=0,0,B19/B18)", "f", NF_D)
put(ws, "A21", "Share of all records", "lbl"); put(ws, "B21", "=B17/SUM(" + M("F") + ")", "f", NF_P)
put(ws, "A22", '=IF(AND(B9="All",OR(B5="All",B5="SEC")),"SEC dollars are separate tags: pick one tag."&IF(B5="All"," SEC is left out of totals.",""),IF(B9="SEC | Other","SEC Other mixes tags: dollars are not meaningful.",""))&IF(B5="All"," Dollars differ by source: pick one source for a clean total.","")', "note")
ws.merge_cells("A22:D22"); ws["A22"].alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[22].height = 44
for r in range(17, 22): ws["B" + str(r)].font = Font(name=FONT, size=11, bold=(r in (17, 19)))

# by year
sec(ws, 24, "By year", 1, 4)
for j, t in enumerate(["Year", "Records", "Total amount ($)"]): put(ws, L(j + 1) + "25", t, "hdr")
Y0 = 26
for i, y in enumerate(YEARS):
    r = Y0 + i; put(ws, "A" + str(r), "=Lists!B" + str(i + 2), "link", "0", h="left")
    cond = "AND($A" + str(r) + ">=f_ylo,$A" + str(r) + "<=f_yhi)"
    put(ws, "B" + str(r), "=IF(" + cond + "," + rec(dict(year=("eq", "$A" + str(r)))) + ",0)", "f", NF_N)
    put(ws, "C" + str(r), "=IF(" + cond + "," + amt(dict(year=("eq", "$A" + str(r)))) + ",0)", "f", NF_D)
Y1 = Y0 + len(YEARS) - 1
# by month
MO0 = Y1 + 4
sec(ws, MO0 - 2, "By month (no size filter)", 1, 4)
for j, t in enumerate(["Month", "Records", "Total amount ($)", "Month no."]): put(ws, L(j + 1) + str(MO0 - 1), t, "hdr")
MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for i in range(12):
    r = MO0 + i; put(ws, "A" + str(r), MN[i], "lbl"); put(ws, "D" + str(r), i + 1, "note", "0")
    put(ws, "B" + str(r), "=" + rec(dict(month=("eq", "$D" + str(r))), "month"), "f", NF_N)
    put(ws, "C" + str(r), "=" + amt(dict(month=("eq", "$D" + str(r))), "month"), "f", NF_D)
MO1 = MO0 + 11
# by source
S0 = MO1 + 5
sec(ws, S0 - 2, "By source (no source filter)", 1, 4)
for j, t in enumerate(["Source", "Records", "Total amount ($)"]): put(ws, L(j + 1) + str(S0 - 1), t, "hdr")
for i, s in enumerate(SRC):
    r = S0 + i; put(ws, "A" + str(r), "=Lists!A" + str(i + 3), "link", h="left")
    put(ws, "B" + str(r), "=" + rec(dict(src=("eq", "$A" + str(r)))), "f", NF_N)
    put(ws, "C" + str(r), '=IF(AND($A' + str(r) + '="SEC",f_cat="*"),"n/a: pick a tag",' + sumifs("main", "G", dict(src=("eq", "$A" + str(r)))) + ")", "f", NF_D, h="right")
S1 = S0 + len(SRC) - 1
# by band
B0 = S1 + 5
sec(ws, B0 - 2, "By size (no size filter)", 1, 4)
for j, t in enumerate(["Band", "Records", "Total amount ($)", "Band id"]): put(ws, L(j + 1) + str(B0 - 1), t, "hdr")
for i in range(NB):
    r = B0 + i; put(ws, "A" + str(r), "=Lists!E" + str(i + 2), "link", h="left"); put(ws, "D" + str(r), "=Lists!F" + str(i + 2), "note", "0")
    put(ws, "B" + str(r), "=" + rec(dict(band=("eq", "$D" + str(r)))), "f", NF_N)
    put(ws, "C" + str(r), "=" + amt(dict(band=("eq", "$D" + str(r)))), "f", NF_D)
B1 = B0 + NB - 1

# ---- Calc tables for rankings
put(ws_calc, "D1", "State", "hdr"); put(ws_calc, "E1", "Records", "hdr"); put(ws_calc, "F1", "Amount ($)", "hdr"); put(ws_calc, "G1", "Rank value", "hdr")
for i in range(NST):
    r = i + 2; put(ws_calc, "D" + str(r), "=Lists!C" + str(i + 3 - 1 + 0), "link") if False else put(ws_calc, "D" + str(r), "=Lists!C" + str(i + 2), "link")
    put(ws_calc, "E" + str(r), "=" + rec(dict(state=("eq", "$D" + str(r)))), "f", NF_N); put(ws_calc, "F" + str(r), "=" + amt(dict(state=("eq", "$D" + str(r)))), "f", NF_D)
    put(ws_calc, "G" + str(r), '=IF(f_metric="Amount",F' + str(r) + ",E" + str(r) + ")+ROW()*tiebreak", "f", "0.000000")
put(ws_calc, "I1", "Category", "hdr"); put(ws_calc, "J1", "Records", "hdr"); put(ws_calc, "K1", "Amount ($)", "hdr"); put(ws_calc, "L1", "Rank value", "hdr")
for i in range(NCAT):
    r = i + 2; put(ws_calc, "I" + str(r), "=Lists!D" + str(i + 2), "link")
    put(ws_calc, "J" + str(r), "=" + rec(dict(cat=("eq", "$I" + str(r)))), "f", NF_N)
    put(ws_calc, "K" + str(r), '=IF(AND(LEFT($I' + str(r) + ',3)="SEC",OR(Explorer!$B$9="All",RIGHT($I' + str(r) + ',8)=" | Other",RIGHT($I' + str(r) + ',13)=" | Not stated")),0,' + amt(dict(cat=("eq", "$I" + str(r)))) + ")", "f", NF_D)
    put(ws_calc, "L" + str(r), '=IF(f_metric="Amount",K' + str(r) + ",J" + str(r) + ")+ROW()*tiebreak", "f", "0.000000")
# state/category rows start at Lists row 2 = "All"; skip "All" by ranking only rows 3..
SR0, SR1 = 3, NST + 1; CR0, CR1 = 3, NCAT + 1
T0 = B1 + 5
sec(ws, T0 - 2, "Top 10 states (no state filter)", 1, 4)
for j, t in enumerate(["Rank / State", "Records", "Total amount ($)", "Share of records"]): put(ws, L(j + 1) + str(T0 - 1), t, "hdr")
for k in range(10):
    r = T0 + k; m = "MATCH(LARGE(Calc!$G$" + str(SR0) + ":$G$" + str(SR1) + "," + str(k + 1) + "),Calc!$G$" + str(SR0) + ":$G$" + str(SR1) + ",0)"
    put(ws, "A" + str(r), "=" + str(k + 1) + '&". "&INDEX(Calc!$D$' + str(SR0) + ":$D$" + str(SR1) + "," + m + ")", "f")
    put(ws, "B" + str(r), "=INDEX(Calc!$E$" + str(SR0) + ":$E$" + str(SR1) + "," + m + ")", "f", NF_N)
    put(ws, "C" + str(r), "=INDEX(Calc!$F$" + str(SR0) + ":$F$" + str(SR1) + "," + m + ")", "f", NF_D)
    put(ws, "D" + str(r), "=IF(SUM(Calc!$E$" + str(SR0) + ":$E$" + str(SR1) + ")=0,0,B" + str(r) + "/SUM(Calc!$E$" + str(SR0) + ":$E$" + str(SR1) + "))", "f", NF_P)
C0 = T0 + 15
sec(ws, C0 - 2, "Top 10 categories (no category filter)", 1, 4)
for j, t in enumerate(["Rank / Category", "Records", "Total amount ($)", "Share of records"]): put(ws, L(j + 1) + str(C0 - 1), t, "hdr")
for k in range(10):
    r = C0 + k; m = "MATCH(LARGE(Calc!$L$" + str(CR0) + ":$L$" + str(CR1) + "," + str(k + 1) + "),Calc!$L$" + str(CR0) + ":$L$" + str(CR1) + ",0)"
    put(ws, "A" + str(r), "=" + str(k + 1) + '&". "&INDEX(Calc!$I$' + str(CR0) + ":$I$" + str(CR1) + "," + m + ")", "f")
    put(ws, "B" + str(r), "=INDEX(Calc!$J$" + str(CR0) + ":$J$" + str(CR1) + "," + m + ")", "f", NF_N)
    put(ws, "C" + str(r), "=INDEX(Calc!$K$" + str(CR0) + ":$K$" + str(CR1) + "," + m + ")", "f", NF_D)
    put(ws, "D" + str(r), "=IF(SUM(Calc!$J$" + str(CR0) + ":$J$" + str(CR1) + ")=0,0,B" + str(r) + "/SUM(Calc!$J$" + str(CR0) + ":$J$" + str(CR1) + "))", "f", NF_P)
    ws["A" + str(r)].alignment = Alignment(wrap_text=False)
ws.column_dimensions["A"].width = 50
put(ws, "A" + str(C0 + 12), "Top 10 ranked by B12. Shares are of filtered records.", "note")

# ---- charts on Explorer
def chart(title, r0, r1, col, anchor, yt, nf):
    ch = BarChart(); ch.type = "col"; ch.title = title; ch.style = 10; ch.width = 18; ch.height = 7.6; ch.legend = None
    ch.add_data(Reference(ws, min_col=col, min_row=r0 - 1, max_row=r1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=r1))
    ch.y_axis.numFmt = nf; ch.y_axis.majorGridlines = None; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.title = yt
    gp = GraphicalProperties(solidFill=BARC); gp.line.solidFill = BARC; ch.series[0].graphicalProperties = gp
    ws.add_chart(ch, anchor)
chart("Records by year", Y0, Y1, 2, "F4", "Records", "#,##0")
chart("Total amount by year", Y0, Y1, 3, "F20", "Dollars", "$#,##0,,\"M\"")
chart("Records by month", MO0, MO1, 2, "F36", "Records", "#,##0")
chart("Records by amount size", B0, B1, 2, "F52", "Records", "#,##0")
ws.freeze_panes = "A4"

# ------------------------------------------------------------------ Extract_Request
widths(ws_ext, [34, 100])
put(ws_ext, "A1", "Extract request", "title")
put(ws_ext, "A2", "Command and SQL for your Explorer filters.", "sub")
sec(ws_ext, 4, "Your filters (from Explorer)", 1, 2)
labs = [("Source", "B5"), ("Year from", "B6"), ("Year to", "B7"), ("State", "B8"), ("Category", "B9"), ("Amount band from", "B10"), ("Amount band to", "B11")]
for i, (a, c) in enumerate(labs): put(ws_ext, "A" + str(5 + i), a, "lbl"); put(ws_ext, "B" + str(5 + i), '=Explorer!' + c, "link", h="left")
put(ws_ext, "A12", "Records matching (Explorer)", "b"); put(ws_ext, "B12", "=Explorer!B17", "link", NF_N, h="left")
put(ws_ext, "A13", "Minimum amount used ($)", "lbl"); put(ws_ext, "B13", '=IF(Explorer!B13<>"",Explorer!B13,IF(f_blo<=1,"",INDEX(Lists!$H$2:$H$' + str(NB + 1) + ",f_blo+1)))", "f", "#,##0.00", h="left")
put(ws_ext, "A14", "Maximum amount used ($)", "lbl"); put(ws_ext, "B14", '=IF(Explorer!B14<>"",Explorer!B14,IF(f_bhi=8,"",INDEX(Lists!$I$2:$I$' + str(NB + 1) + ",f_bhi+1)))", "f", "#,##0.00", h="left")
put(ws_ext, "A15", "Source used", "lbl"); put(ws_ext, "B15", '=IF(Explorer!B5<>"All",Explorer!B5,IF(Explorer!B9<>"All",LEFT(Explorer!B9,FIND(" | ",Explorer!B9)-1),"All"))', "f", h="left")
put(ws_ext, "A16", "Category name used", "lbl"); put(ws_ext, "B16", '=IF(Explorer!B9="All","",MID(Explorer!B9,FIND(" | ",Explorer!B9)+3,200))', "f", h="left")
put(ws_ext, "A17", "Check", "lbl"); put(ws_ext, "B17", '=IF(OR(RIGHT(Explorer!B9,8)=" | Other",RIGHT(Explorer!B9,13)=" | Not stated"),"Other and Not stated cannot be extracted by name. Pick a named category.",IF(B12>xl_rows,"Over 1,048,575 records: use CSV or Parquet, not Excel.",IF(B12=0,"No records match.","OK")))', "warn", h="left")
Q = "CHAR(34)"
cmd = ('="python3 03_Pipeline/extract.py --table unified"&IF(B15="All",""," --source "&' + Q + '&B15&' + Q + ')&" --year-from "&Explorer!B6&" --year-to "&Explorer!B7'
       '&IF(Explorer!B8="All",""," --state "&Explorer!B8)&IF(B16="",""," --category "&' + Q + '&B16&' + Q + ')&IF(B13="",""," --min-amount "&B13)&IF(B14="",""," --max-amount "&B14)'
       '&" --format csv --out 04_Extracts/my_extract.csv"')
sec(ws_ext, 19, "Command (run from project folder)", 1, 2)
put(ws_ext, "A20", cmd, "f", wrap=True); ws_ext.merge_cells("A20:B20"); ws_ext.row_dimensions[20].height = 64
Q1 = "CHAR(39)"
sql = ('="SELECT * FROM unified_records WHERE calendar_year BETWEEN "&Explorer!B6&" AND "&Explorer!B7'
       '&IF(B15="All","",CHAR(10)&"  AND source = "&' + Q1 + '&B15&' + Q1 + ')&IF(Explorer!B8="All","",IF(Explorer!B8="Unknown",CHAR(10)&"  AND state IS NULL",CHAR(10)&"  AND state = "&' + Q1 + '&Explorer!B8&' + Q1 + '))'
       '&IF(B16="","",CHAR(10)&"  AND category = "&' + Q1 + '&SUBSTITUTE(B16,' + Q1 + ',' + Q1 + '&' + Q1 + ')&' + Q1 + ')&IF(B13="","",CHAR(10)&"  AND amount >= "&B13)&IF(B14="","",CHAR(10)&"  AND amount <= "&B14)&";"')
sec(ws_ext, 22, "SQL (DuckDB, 02_Clean_Data/analytics.duckdb)", 1, 2)
put(ws_ext, "A23", sql, "f", wrap=True); ws_ext.merge_cells("A23:B23"); ws_ext.row_dimensions[23].height = 110
put(ws_ext, "A25", "For all columns, use a source table (see Data_Dictionary). Change --format to xlsx or parquet.", "note")
ws_ext["A25"].alignment = Alignment(wrap_text=True, vertical="top"); ws_ext.merge_cells("A25:B25"); ws_ext.row_dimensions[25].height = 32

# ------------------------------------------------------------------ Profile
widths(ws_prof, [26, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 16, 16])
put(ws_prof, "A1", "Data profile", "title"); put(ws_prof, "A2", "Build snapshot (blue). Percentages are formulas.", "sub")
GR = {"unified_records": "All five sources", "usaspending_transactions": "Contract transaction",
      "usaspending_awards": "One row per contract award", "fema_pa_projects": "One row per funded project", "fac_reports": "One row per audit report",
      "fac_awards": "Program in a report", "fac_findings": "One row per audit finding", "sec_filings": "One row per filing",
      "sec_tags": "Tag labels and definitions", "sec_facts": "One row per reported USD value", "sba_ppp_loans": "One row per loan"}
sec(ws_prof, 4, "A. Tables", 1, 4)
for j, t in enumerate(["Table", "Rows", "Columns", "What one row is"]): put(ws_prof, L(j + 1) + "5", t, "hdr")
tabs = sorted(prof["tables"], key=lambda x: (x["table"] != "unified_records", x["table"]))
for i, t in enumerate(tabs):
    r = 6 + i; put(ws_prof, "A" + str(r), t["table"], "lbl"); put(ws_prof, "B" + str(r), t["rows"], "inp", NF_N); put(ws_prof, "C" + str(r), t["columns"], "inp", "0", h="center"); put(ws_prof, "D" + str(r), GR.get(t["table"], ""), "lbl")
UR = 6   # unified row in this table
P0 = 6 + len(tabs) + 3
sec(ws_prof, P0 - 2, "B. Unified table by source", 1, 15)
H = ["Source", "Records", "Date from", "Date to", "% with state", "% with category", "% with entity name", "% with amount", "Negative amounts", "Zero amounts", "Minimum ($)", "Median ($)", "90th percentile ($)", "Maximum ($)", "Total ($)"]
for j, t in enumerate(H): put(ws_prof, L(j + 1) + str(P0 - 1), t, "hdr")
ws_prof.row_dimensions[P0 - 1].height = 32
bs = {d["source"]: d for d in prof["by_source"]}
for i, s in enumerate(SRC):
    d = bs[s]; r = P0 + i
    put(ws_prof, "A" + str(r), s, "lbl"); put(ws_prof, "B" + str(r), d["rows"], "inp", NF_N); put(ws_prof, "C" + str(r), d["date_min"], "inp", h="center"); put(ws_prof, "D" + str(r), d["date_max"], "inp", h="center")
    for col, key in zip("EFGH", ["with_state", "with_category", "with_entity", "with_amount"]):
        # hidden counts live in columns R..U so the percentage stays a formula
        cc = {"E": "R", "F": "S", "G": "T", "H": "U"}[col]; put(ws_prof, cc + str(r), d[key], "inp", NF_N)
        put(ws_prof, col + str(r), "=IF($B" + str(r) + "=0,0," + cc + str(r) + "/$B" + str(r) + ")", "f", NF_P)
    for col, key in zip("IJKLMNO", ["negative", "zero", "amt_min", "amt_median", "amt_p90", "amt_max", "amt_sum"]):
        put(ws_prof, col + str(r), d[key], "inp", NF_N if col in "IJ" else NF_D)
for c, t in zip("RSTU", ["Count with state", "Count with category", "Count with entity", "Count with amount"]): put(ws_prof, c + str(P0 - 1), t, "hdr")
widths_extra = {"R": 14, "S": 14, "T": 14, "U": 14}
for k, v in widths_extra.items(): ws_prof.column_dimensions[k].width = v
PE = P0 + len(SRC) - 1
put(ws_prof, "A" + str(PE + 1), "SEC dollar stats mix tags; not meaningful.", "note")
put(ws_prof, "A" + str(PE + 2), "OpenFEMA has no names (0%).", "note")
# tie-outs
K0 = PE + 5
sec(ws_prof, K0 - 1, "C. Checks against the database", 1, 4)
for j, t in enumerate(["Check", "Workbook", "Database", "Result"]): put(ws_prof, L(j + 1) + str(K0), t, "hdr")
put(ws_prof, "A" + str(K0 + 1), "Records in Cube_Main", "lbl"); put(ws_prof, "B" + str(K0 + 1), "=SUM(" + M("F") + ")", "f", NF_N); put(ws_prof, "C" + str(K0 + 1), "=B" + str(UR), "f", NF_N); put(ws_prof, "D" + str(K0 + 1), '=IF(B' + str(K0 + 1) + "=C" + str(K0 + 1) + ',"PASS","FAIL")', "f", h="center")
put(ws_prof, "A" + str(K0 + 2), "Records in Cube_Month", "lbl"); put(ws_prof, "B" + str(K0 + 2), "=SUM(" + O("F") + ")", "f", NF_N); put(ws_prof, "C" + str(K0 + 2), "=B" + str(UR), "f", NF_N); put(ws_prof, "D" + str(K0 + 2), '=IF(B' + str(K0 + 2) + "=C" + str(K0 + 2) + ',"PASS","FAIL")', "f", h="center")
put(ws_prof, "A" + str(K0 + 3), "Records by source (sum of rows above)", "lbl"); put(ws_prof, "B" + str(K0 + 3), "=SUM(B" + str(P0) + ":B" + str(PE) + ")", "f", NF_N); put(ws_prof, "C" + str(K0 + 3), "=B" + str(UR), "f", NF_N); put(ws_prof, "D" + str(K0 + 3), '=IF(B' + str(K0 + 3) + "=C" + str(K0 + 3) + ',"PASS","FAIL")', "f", h="center")
for i, s in enumerate([x for x in SRC if x != "SEC"]):
    r = K0 + 4 + i; pr = P0 + SRC.index(s)
    put(ws_prof, "A" + str(r), "Amount, " + s + " (cube vs database)", "lbl"); put(ws_prof, "B" + str(r), "=SUMIFS(" + M("G") + "," + M("A") + ',"' + s + '")', "f", NF_D); put(ws_prof, "C" + str(r), "=O" + str(pr), "f", NF_D)
    put(ws_prof, "D" + str(r), '=IF(ABS(B' + str(r) + "-C" + str(r) + ")<=1,\"PASS\",\"FAIL\")", "f", h="center")
KE = K0 + 4 + len(SRC) - 2
put(ws_prof, "A" + str(KE + 1), "All checks pass?", "b"); put(ws_prof, "D" + str(KE + 1), '=IF(COUNTIF(D' + str(K0 + 1) + ":D" + str(KE) + ',"FAIL")=0,"PASS","REVIEW")', "f", h="center", bold=True)
CHECKCELL = "Profile!D" + str(KE + 1)

# ------------------------------------------------------------------ Data dictionary
widths(ws_dic, [24, 28, 12, 10, 12, 30, 60, 70])
put(ws_dic, "A1", "Data dictionary", "title"); put(ws_dic, "A2", "Every column and how it was made. nz trim, to_dt date, to_num number, st state check, std_name matching name.", "sub")
for j, t in enumerate(["Table", "Column", "Type", "% missing", "Distinct values (approx.)", "Example (lowest value)", "How it was made", "Note"]): put(ws_dic, L(j + 1) + "4", t, "hdr")
ws_dic.row_dimensions[4].height = 32
UN = {"source": "Which dataset the row is from", "record_id": "Key from the source",
      "record_date": "Main date: action date (USAspending), first obligation (OpenFEMA), accepted date (FAC), approval date (SBA PPP), period end (SEC)",
      "calendar_year": "Year of record_date", "calendar_month": "Month of record_date", "fiscal_year": "Federal fiscal year (Oct to Sep)",
      "state": "State code: place of work (USAspending), state (OpenFEMA), auditee (FAC), borrower (SBA PPP), business address (SEC)",
      "county": "County name, where the source has one",
      "category": "Main group: agency (USAspending, FAC), damage category (OpenFEMA), business type (SBA PPP), SEC tag",
      "subcategory": "Second group: NAICS (USAspending), incident type (OpenFEMA), program (FAC), loan status (SBA PPP), form (SEC)",
      "industry_code": "NAICS (USAspending, SBA PPP) or SIC (SEC)",
      "entity_name": "Recipient, auditee, borrower or filer (OpenFEMA has IDs only)", "entity_name_std": "Standard name for matching across sources",
      "entity_id": "Recipient UEI, applicant ID, auditee UEI or SEC CIK", "amount": "Main dollar value (see amount_measure)",
      "amount_measure": "What amount means for this source", "data_flag": "Special case: date_from_declaration, migration_record or qtrs_N"}
SHORT = {"record_date": "Main date for this table", "calendar_year": "Year of record_date", "calendar_month": "Month of record_date", "fiscal_year": "Federal fiscal year (Oct to Sep)",
         "record_date_basis": "Which date was used", "recipient_name_std": "Standard name for matching", "auditee_name_std": "Standard name for matching",
         "company_name_std": "Standard name for matching", "borrower_name_std": "Standard name for matching", "is_migration_record": "TRUE = from the older FAC system",
         "assistance_listing": "Program number (agency prefix.extension)", "federal_agency": "Agency name from the two-digit prefix"}
d2 = sorted(dic, key=lambda x: (x["table"] != "unified_records", x["table"]))
for i, d in enumerate(d2):
    r = 5 + i; note = UN.get(d["column"], "") if d["table"] == "unified_records" else SHORT.get(d["column"], "")
    put(ws_dic, "A" + str(r), d["table"], "lbl"); put(ws_dic, "B" + str(r), d["column"], "lbl"); put(ws_dic, "C" + str(r), d["type"], "lbl"); put(ws_dic, "D" + str(r), d["null_pct"], "inp", NF_P)
    put(ws_dic, "E" + str(r), d["distinct_approx"], "inp", NF_N); put(ws_dic, "F" + str(r), ("Text: " + d["example"].lstrip("#")) if d["example"].startswith("#") else d["example"], "lbl"); ws_dic["F" + str(r)].data_type = "s"; put(ws_dic, "G" + str(r), d["derivation"], "lbl"); put(ws_dic, "H" + str(r), note, "lbl")
ws_dic.freeze_panes = "A5"; ws_dic.auto_filter.ref = "A4:H" + str(4 + len(d2))

# ------------------------------------------------------------------ README
widths(ws_readme, [26, 110])
put(ws_readme, "A1", "Cross-Source Dataset: Federal Spending and Financial Data, FY1998 to FY2026", "title")
put(ws_readme, "A2", "Filter, summarize and extract five cleaned datasets.", "sub")
rows = [
 ("What this is", "Cleaned data from five public sources."),
 ("Datasets", "USAspending, OpenFEMA, FAC, SEC, SBA PPP."),
 ("Step 1: filter", "Change the yellow cells on Explorer."),
 ("Step 2: extract", "Extract_Request gives the command for exact records."),
 ("Step 3: learn", "Profile shows coverage. Data_Dictionary lists columns."),
 ("Amount filter", "Size bands only. Exact limits: Explorer B13 and B14 (extract only)."),
 ("Categories", "Top 14 per source are named; the rest are Other."),
 ("SEC dollars", "SEC amounts are separate tags. Pick one tag."),
 ("Dates", "One date per source. Year is the calendar year."),
 ("Cleaning", "Dates and numbers converted. Bad states blanked. No personal details. Totals match raw files."),
 ("Colors", "Blue = typed. Yellow = filters. Black = formulas. Green = links."),
 ("Refresh", "Re-run the scripts in 03_Pipeline."),
 ("Check status", None),
]
for i, (a, b) in enumerate(rows):
    r = 4 + i; put(ws_readme, "A" + str(r), a, "b")
    if b is None: put(ws_readme, "B" + str(r), "=" + CHECKCELL, "link", h="left")
    else: put(ws_readme, "B" + str(r), b, "lbl", wrap=True); ws_readme.row_dimensions[r].height = max(18, 15 * (len(b) // 105 + 1))
    ws_readme["A" + str(r)].alignment = Alignment(vertical="top")
sec(ws_readme, 18, "Sheets", 1, 2)
for i, (a, b) in enumerate([("Explorer", "Filters, results, charts"), ("Extract_Request", "Command and SQL"), ("Profile", "Coverage and checks"), ("Data_Dictionary", "All columns"),
                            ("Lists, Calc", "Helpers (do not edit)"), ("Cube_Main, Cube_Month", "Explorer data")]):
    put(ws_readme, "A" + str(19 + i), a, "lbl"); put(ws_readme, "B" + str(19 + i), b, "lbl")

for w in wb.worksheets:
    w.sheet_view.showGridLines = False if w.title in ("README", "Explorer", "Extract_Request", "Profile") else True
    w.page_setup.orientation = "landscape"
    if w.title in ("README", "Explorer", "Extract_Request", "Profile", "Data_Dictionary"):
        w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0; w.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
for w in wb.worksheets:
    w.sheet_properties.tabColor = MIDG if w.title in ("README", "Explorer", "Extract_Request", "Profile", "Data_Dictionary") else "A9D18E"
wb.save(OUTF); print("saved", OUTF)
