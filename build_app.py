#!/usr/bin/env python3
"""Generate a full 'application-style' Digital Bank Savings Tracker (.xlsx)
using only the Python standard library (no openpyxl available offline).

Sheets: Dashboard | Deposit | Transactions | Balance | Interest
- App look: navy header band, nav buttons (hyperlinks), KPI cards, hidden
  gridlines, frozen panes, consistent theme on every tab.
- Deposit & Transactions are Excel Tables => built-in filter dropdowns.
- Dashboard has 3 charts: Savings Growth (line), Savings by Bank (pie),
  Monthly Interest Earned (column).
- Interest auto-calculates daily/monthly earnings that compound; balances
  flow Deposit/Transactions -> Balance -> Interest -> Dashboard.
"""
import zipfile
from datetime import date

# ----------------------------------------------------------------------------
# DATA MODEL
# ----------------------------------------------------------------------------
# (Bank, Account, Type, Label, AnnualRate, Frequency)  -- Maya Savings = 10%
ACCOUNTS = [
    ("Maya",     "Savings",        "Savings",      "Maya - Savings",        0.10,   "Daily"),
    ("Maya",     "Time Deposit 1", "Time Deposit", "Maya - Time Deposit 1", 0.06,   "Monthly"),
    ("Maya",     "Time Deposit 2", "Time Deposit", "Maya - Time Deposit 2", 0.06,   "Monthly"),
    ("Maya",     "Personal Goal",  "Goal",         "Maya - Personal Goal",  0.056,  "Monthly"),
    ("UNO",      "Digi Savings",   "Savings",      "UNO - Digi Savings",    0.035,  "Daily"),
    ("UNO",      "Time Deposit 1", "Time Deposit", "UNO - Time Deposit 1",  0.06,   "Monthly"),
    ("UNO",      "Time Deposit 2", "Time Deposit", "UNO - Time Deposit 2",  0.06,   "Monthly"),
    ("Tonik",    "Stash 1",        "Savings",      "Tonik - Stash 1",       0.04,   "Monthly"),
    ("Tonik",    "Stash 2",        "Savings",      "Tonik - Stash 2",       0.04,   "Monthly"),
    ("Tonik",    "Time Deposit",   "Time Deposit", "Tonik - Time Deposit",  0.08,   "Monthly"),
    ("Banko",    "Savings",        "Savings",      "Banko - Savings",       0.05,   "Daily"),
    ("CIMB",     "Savings",        "Savings",      "CIMB - Savings",        0.023,  "Daily"),
    ("GoTyme",   "GoSave",         "Savings",      "GoTyme - GoSave",       0.03,   "Monthly"),
    ("Maribank", "Savings",        "Savings",      "Maribank - Savings",    0.0325, "Daily"),
]
BANKS = ["Maya", "UNO", "Tonik", "Banko", "CIMB", "GoTyme", "Maribank"]

EPOCH = date(1899, 12, 30)
def serial(y, m, d):
    return (date(y, m, d) - EPOCH).days

# Seed example deposits (label -> amount) so the app shows data immediately.
SEED_DEP = {
    "Maya - Savings": 50000, "Maya - Time Deposit 1": 100000,
    "Maya - Time Deposit 2": 100000, "Maya - Personal Goal": 25000,
    "UNO - Digi Savings": 20000, "UNO - Time Deposit 1": 50000,
    "UNO - Time Deposit 2": 50000, "Tonik - Stash 1": 10000,
    "Tonik - Stash 2": 10000, "Tonik - Time Deposit": 30000,
    "Banko - Savings": 15000, "CIMB - Savings": 10000,
    "GoTyme - GoSave": 12000, "Maribank - Savings": 8000,
}
# Deposit rows: (dateserial, bank, label, amount, note)
DEP_ROWS = [(serial(2026, 6, 1), b, lbl, SEED_DEP[lbl], "Opening balance (example)")
            for (b, a, t, lbl, r, f) in ACCOUNTS]

# Transaction rows: (dateserial, type, bank, label, signed amount, note)
TX_ROWS = [
    (serial(2026, 6, 15), "Withdrawal", "Maya",  "Maya - Savings",     -2000,  "ATM withdrawal (example)"),
    (serial(2026, 6, 25), "Transfer",   "UNO",   "UNO - Digi Savings",  5000,  "Transfer in (example)"),
]
# Interest is now accrued automatically (see Interest sheet), so it is NOT
# posted as a manual transaction here (that would double-count).

# ---- compute cached values (so the file shows numbers before recalculation) --
dep_by_label = {lbl: 0.0 for (_, _, _, lbl, _, _) in ACCOUNTS}
for (_, b, lbl, amt, _) in DEP_ROWS:
    dep_by_label[lbl] += amt
nettx_by_label = {lbl: 0.0 for (_, _, _, lbl, _, _) in ACCOUNTS}
for (_, ty, b, lbl, amt, _) in TX_ROWS:
    nettx_by_label[lbl] += amt

calc = {}   # label -> dict of computed values
for (bank, acct, atype, lbl, rate, freq) in ACCOUNTS:
    dep = dep_by_label[lbl]; ntx = nettx_by_label[lbl]
    principal = dep + ntx
    if freq == "Daily":
        pr = rate / 365; n = 365
        mfactor = (1 + rate / 365) ** (365 / 12)
    else:
        pr = rate / 12; n = 12
        mfactor = 1 + rate / 12
    iper = principal * pr
    projbal = principal * (1 + pr) ** n
    int1yr = projbal - principal
    # ---- live accrual (grows with TODAY): from earliest deposit date to "today"
    START = serial(2026, 6, 1)          # earliest deposit date for the seed data
    REF_TODAY = serial(2026, 7, 5)      # matches the current date for cached preview
    if freq == "Daily":
        elapsed = max(0, REF_TODAY - START)                 # days
    else:
        # whole months between START and REF_TODAY (like DATEDIF "m")
        elapsed = max(0, (2026 - 2026) * 12 + (7 - 6) - (1 if 5 < 1 else 0))
    i2d = principal * ((1 + pr) ** elapsed - 1)             # interest earned to date
    bal_today = principal + i2d
    calc[lbl] = dict(bank=bank, dep=dep, ntx=ntx, principal=principal, rate=rate,
                     freq=freq, pr=pr, iper=iper, n=n, projbal=projbal,
                     int1yr=int1yr, mfactor=mfactor, elapsed=elapsed,
                     i2d=i2d, bal_today=bal_today, start=START)

bank_totals = {b: 0.0 for b in BANKS}      # principal per bank
bank_today = {b: 0.0 for b in BANKS}       # balance-today (incl. accrued interest) per bank
for lbl, c in calc.items():
    bank_totals[c["bank"]] += c["principal"]
    bank_today[c["bank"]] += c["bal_today"]

MONTHS = list(range(0, 13))
def total_at(m):
    return sum(c["principal"] * (c["mfactor"] ** m) for c in calc.values())
growth_total = [total_at(m) for m in MONTHS]
monthly_interest = [0.0] + [growth_total[i] - growth_total[i - 1] for i in range(1, len(MONTHS))]

KPI_total_savings = sum(c["principal"] for c in calc.values())
KPI_total_deposits = sum(c["dep"] for c in calc.values())
KPI_balance_today = sum(c["bal_today"] for c in calc.values())      # grows with TODAY()
KPI_interest_to_date = sum(c["i2d"] for c in calc.values())         # grows with TODAY()
KPI_proj_interest = sum(c["int1yr"] for c in calc.values())
KPI_proj_balance = sum(c["projbal"] for c in calc.values())

# Key-metric extras + goal tracker
interest_credited = sum(a for (_, ty, _, _, a, _) in TX_ROWS if ty == "Interest")
total_withdrawn = -sum(a for (_, ty, _, _, a, _) in TX_ROWS if ty == "Withdrawal")
GOAL_TARGET = 100000
maya_goal_balance = calc["Maya - Personal Goal"]["bal_today"]
goal_progress = (maya_goal_balance / GOAL_TARGET) if GOAL_TARGET else 0

# 24-month projection: one cumulative series per bank + total monthly interest
MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MONTH_LABELS = [f"{MONTH_ABBR[mo]} {y}" for y in (2026, 2027) for mo in range(12)]
bank_series = {b: [sum(c['principal'] * (c['mfactor'] ** m) for c in calc.values() if c['bank'] == b)
                   for m in range(24)] for b in BANKS}
total_series24 = [sum(bank_series[b][m] for b in BANKS) for m in range(24)]
monthly_int24 = [0.0] + [total_series24[m] - total_series24[m - 1] for m in range(1, 24)]
# per-bank monthly interest earned (for the stacked column chart)
bank_monthly_int = {b: [0.0] + [bank_series[b][m] - bank_series[b][m - 1] for m in range(1, 24)]
                    for b in BANKS}
# distinct, well-separated palette (one color per bank, shared by both charts)
BANK_COLORS = {"Maya": "4472C4", "UNO": "ED7D31", "Tonik": "70AD47", "Banko": "FFC000",
               "CIMB": "7030A0", "GoTyme": "C00000", "Maribank": "595959"}

# ----------------------------------------------------------------------------
# XML helpers
# ----------------------------------------------------------------------------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

COLS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"  # column letters

def colL(i):            # 0-indexed column number -> letter(s)
    s = ""; i += 1
    while i > 0:
        i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s

def ct(ref, s, val):    # inline text cell
    return f'<c r="{ref}" t="inlineStr" s="{s}"><is><t xml:space="preserve">{esc(val)}</t></is></c>'
def cn(ref, s, val):    # number cell
    return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
def cf(ref, s, formula, cached=None):   # formula cell (+ optional cached value)
    v = f'<v>{cached}</v>' if cached is not None else ''
    return f'<c r="{ref}" s="{s}"><f>{esc(formula)}</f>{v}</c>'
def ce(ref, s):         # empty styled cell
    return f'<c r="{ref}" s="{s}"/>'
def cf_str(ref, s, formula, cached_text):   # formula cell that returns TEXT
    return f'<c r="{ref}" s="{s}" t="str"><f>{esc(formula)}</f><v>{esc(cached_text)}</v></c>'

def band_row(rownum, style, text=None, text_style=None, ncols=13):
    """A full-width colored band row across columns A..(ncols)."""
    cells = []
    for i in range(ncols):
        ref = f"{colL(i)}{rownum}"
        if i == 0 and text is not None:
            cells.append(ct(ref, text_style if text_style else style, text))
        else:
            cells.append(ce(ref, style))
    return f'<row r="{rownum}" ht="22" customHeight="1">' + "".join(cells) + "</row>"

# Navigation bar (row 2). active = current sheet name.
NAV = [("Dashboard", "\U0001F3E0 Dashboard"), ("Deposit", "\u2795 Deposit"),
       ("Transactions", "\U0001F504 Transactions"), ("Balance", "\U0001F4BC Balance"),
       ("Interest", "\U0001F4C8 Interest")]
NAV_PAIRS = ["A2:B2", "C2:D2", "E2:F2", "G2:H2", "I2:J2"]

def nav_row(active, ncols=13):
    cells, hyper, merges = [], [], []
    # button pairs A..J
    for idx, (sheet, label) in enumerate(NAV):
        left = colL(idx * 2); right = colL(idx * 2 + 1)
        s = S["nav_active"] if sheet == active else S["nav"]
        cells.append(ct(f"{left}2", s, label))
        cells.append(ce(f"{right}2", s))
        merges.append(f"{left}2:{right}2")
        hyper.append(f'<hyperlink ref="{left}2:{right}2" location="{sheet}!A1" display="{esc(label)}"/>')
    # filler navy cells beyond the nav buttons
    for i in range(10, ncols):
        cells.append(ce(f"{colL(i)}2", S["nav"]))
    row = f'<row r="2" ht="24" customHeight="1">' + "".join(cells) + "</row>"
    return row, merges, hyper

# ----------------------------------------------------------------------------
# STYLES
# ----------------------------------------------------------------------------
# numFmts: 164 currency, builtins 10 percent, 14 date, 1 integer, 3 thousands
CURR = 164
FONTS = [
    '<font><sz val="11"/><name val="Calibri"/><color rgb="FF000000"/></font>',                 # 0 default
    '<font><b/><sz val="18"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>',              # 1 title
    '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>',              # 2 white bold
    '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF1F3864"/></font>',              # 3 navy bold
    '<font><b/><sz val="20"/><name val="Calibri"/><color rgb="FF1F3864"/></font>',              # 4 kpi value navy
    '<font><i/><sz val="10"/><name val="Calibri"/><color rgb="FF808080"/></font>',              # 5 subtitle gray
    '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF000000"/></font>',              # 6 bold black
    '<font><b/><sz val="12"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>',              # 7 section white
    '<font><sz val="9"/><name val="Calibri"/><color rgb="FF808080"/></font>',                   # 8 small gray
    '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF375623"/></font>',              # 9 green bold
]
FILLS = [
    '<fill><patternFill patternType="none"/></fill>',                                            # 0
    '<fill><patternFill patternType="gray125"/></fill>',                                         # 1
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>',  # 2 navy
    '<fill><patternFill patternType="solid"><fgColor rgb="FF548235"/><bgColor indexed="64"/></patternFill></fill>',  # 3 green
    '<fill><patternFill patternType="solid"><fgColor rgb="FF2E5496"/><bgColor indexed="64"/></patternFill></fill>',  # 4 med blue
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>',  # 5 light gray
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>',  # 6 input yellow
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>',  # 7 kpi light blue
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>',  # 8 light green
]
BORDERS = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',  # 0 none
    ('<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>'
     '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'),  # 1 thin gray
]

# Each xf: (numFmt, font, fill, border, extra-attrs, alignment-or-None)
XF = []
S = {}
def add(name, numFmt, font, fill, border, align=None):
    apply = f'applyFont="1" applyFill="1" applyBorder="1"'
    if numFmt: apply += ' applyNumberFormat="1"'
    al = ''
    if align:
        apply += ' applyAlignment="1"'
        al = f'<alignment {align}/>'
    XF.append(f'<xf numFmtId="{numFmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" {apply}>{al}</xf>'
              if al else
              f'<xf numFmtId="{numFmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" {apply}/>')
    S[name] = len(XF) - 1

add("default", 0, 0, 0, 0)
add("title",   0, 1, 2, 0, 'horizontal="left" vertical="center"')
add("nav",     0, 2, 2, 0, 'horizontal="center" vertical="center"')
add("nav_active", 0, 2, 3, 0, 'horizontal="center" vertical="center"')
add("section", 0, 7, 4, 0, 'horizontal="left" vertical="center"')
add("subtitle", 0, 5, 5, 0, 'horizontal="left" vertical="center"')
add("thead",   0, 2, 4, 1, 'horizontal="center" vertical="center" wrapText="1"')
add("txt",     0, 0, 0, 1, 'vertical="center"')
add("curr",    CURR, 0, 0, 1, 'vertical="center"')
add("pct",     10, 0, 0, 1, 'horizontal="center" vertical="center"')
add("intc",    1, 0, 0, 1, 'horizontal="center" vertical="center"')
add("date",    14, 0, 0, 1, 'horizontal="center" vertical="center"')
add("kpi_lbl", 0, 3, 7, 1, 'horizontal="center" vertical="center"')
add("kpi_curr", CURR, 4, 7, 1, 'horizontal="center" vertical="center"')
add("kpi_int",  1, 4, 7, 1, 'horizontal="center" vertical="center"')
add("tot_txt", 0, 6, 8, 1, 'horizontal="left" vertical="center"')
add("tot_curr", CURR, 6, 8, 1, 'vertical="center"')
add("note",    0, 8, 5, 0, 'vertical="center" wrapText="1"')
add("txt_c",   0, 0, 0, 1, 'horizontal="center" vertical="center"')
add("curr_col", CURR, 0, 0, 0, 'vertical="center"')   # table column (no border)
add("date_col", 14, 0, 0, 0, 'horizontal="center" vertical="center"')
add("txt_col",  0, 0, 0, 0, 'vertical="center"')
add("dhead",   0, 3, 5, 1, 'horizontal="left" vertical="center"')  # data-region header
add("green_curr", CURR, 9, 8, 1, 'vertical="center"')
# Dashboard v2 styles
add("subsection", 0, 3, 5, 0, 'horizontal="left" vertical="center"')          # navy bold on light gray band
add("subsection_r", 0, 3, 5, 0, 'horizontal="right" vertical="center"')       # right-aligned variant (date stamp)
add("m_lbl", 0, 3, 7, 1, 'horizontal="left" vertical="center"')               # metric label: navy bold, light blue
add("m_val", CURR, 3, 0, 1, 'horizontal="right" vertical="center"')           # metric value: currency right, white
add("m_val_pct", 10, 9, 0, 1, 'horizontal="right" vertical="center"')         # progress %: green bold
add("m_val_input", CURR, 3, 6, 1, 'horizontal="right" vertical="center"')     # editable goal target: yellow
add("bank_val", CURR, 0, 0, 1, 'horizontal="right" vertical="center"')        # balance-by-bank value
add("num", 3, 0, 0, 1, 'horizontal="right" vertical="center"')                # compact #,##0 (chart-data matrix)
add("date_input", 14, 0, 6, 1, 'horizontal="center" vertical="center"')       # editable date (yellow) - Interest Start

STYLES_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    f'<numFmts count="1"><numFmt numFmtId="{CURR}" formatCode="[$\u20b1-3409]#,##0.00"/></numFmts>'
    f'<fonts count="{len(FONTS)}">' + "".join(FONTS) + '</fonts>'
    f'<fills count="{len(FILLS)}">' + "".join(FILLS) + '</fills>'
    f'<borders count="{len(BORDERS)}">' + "".join(BORDERS) + '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    f'<cellXfs count="{len(XF)}">' + "".join(XF) + '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>')

# ----------------------------------------------------------------------------
# Worksheet scaffolding
# ----------------------------------------------------------------------------
def sheet_open(active, subtitle, ysplit, tab_selected=False, ncols=13):
    """Return (rows_list, merges_list, hyper_list) pre-populated with the
    title band, nav bar and subtitle band spanning `ncols` columns."""
    last = colL(ncols - 1)
    rows, merges, hyper = [], [], []
    # Row 1 title
    rows.append(band_row(1, S["title"], "\U0001F4B0 Digital Bank Savings Tracker",
                         text_style=S["title"], ncols=ncols))
    rows[-1] = rows[-1].replace('ht="22"', 'ht="30"')
    merges.append(f"A1:{last}1")
    # Row 2 nav
    nrow, nmerges, nhyper = nav_row(active, ncols=ncols)
    rows.append(nrow); merges += nmerges; hyper += nhyper
    # Row 3 subtitle
    rows.append(band_row(3, S["subtitle"], subtitle, text_style=S["subtitle"], ncols=ncols))
    merges.append(f"A3:{last}3")
    return rows, merges, hyper

def cols_xml(widths):
    out = []
    for i, w in enumerate(widths):
        out.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    return "<cols>" + "".join(out) + "</cols>"

def cols_xml_styled(specs):
    # specs: list of (width, style_or_None)
    out = []
    for i, (w, st) in enumerate(specs):
        s = f' style="{st}"' if st is not None else ''
        out.append(f'<col min="{i+1}" max="{i+1}" width="{w}"{s} customWidth="1"/>')
    return "<cols>" + "".join(out) + "</cols>"

def dv_list(sqref, formula1, title, prompt):
    """A dropdown (list) data-validation with an input hint."""
    return (f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            f'promptTitle="{esc(title)}" prompt="{esc(prompt)}" sqref="{sqref}">'
            f'<formula1>{formula1}</formula1></dataValidation>')

def databar(sqref, color="FF63C384", lo=0, hi=1):
    return (f'<conditionalFormatting sqref="{sqref}"><cfRule type="dataBar" priority="1"><dataBar>'
            f'<cfvo type="num" val="{lo}"/><cfvo type="num" val="{hi}"/>'
            f'<color rgb="{color}"/></dataBar></cfRule></conditionalFormatting>')

def worksheet(rows, merges, hyper, ysplit, colsxml, tab_selected=False,
              drawing_rid=None, tablepart_rid=None, cond_fmt="", validations=None):
    sel = ' tabSelected="1"' if tab_selected else ''
    top = ysplit + 1
    view = (f'<sheetViews><sheetView showGridLines="0"{sel} workbookViewId="0">'
            f'<pane ySplit="{ysplit}" topLeftCell="A{top}" activePane="bottomLeft" state="frozen"/>'
            f'<selection pane="bottomLeft" activeCell="A{top}" sqref="A{top}"/></sheetView></sheetViews>')
    mc = ''
    if merges:
        mc = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>'
    dv = ''
    if validations:
        dv = f'<dataValidations count="{len(validations)}">' + "".join(validations) + '</dataValidations>'
    hl = ''
    if hyper:
        hl = '<hyperlinks>' + "".join(hyper) + '</hyperlinks>'
    dr = f'<drawing r:id="{drawing_rid}"/>' if drawing_rid else ''
    tp = f'<tableParts count="1"><tablePart r:id="{tablepart_rid}"/></tableParts>' if tablepart_rid else ''
    data = "".join(rows)
    # schema order: sheetData, mergeCells, conditionalFormatting, dataValidations, hyperlinks, drawing, tableParts
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + view + '<sheetFormatPr defaultRowHeight="15"/>' + colsxml
        + '<sheetData>' + data + '</sheetData>' + mc + cond_fmt + dv + hl + dr + tp + '</worksheet>')

# ============================================================================
# SHEET: DASHBOARD (sheet1)
# ============================================================================
def build_dashboard():
    rows, merges, hyper = sheet_open("Dashboard",
        "A live overview of everything in this workbook. Add deposits on the Deposit tab and this page updates automatically.", 3, True)

    def two_col(rn, label, valcells, ht=18):
        cells = [ct(f"A{rn}", S["m_lbl"], label)] + [ce(f"{COLS[i]}{rn}", S["m_lbl"]) for i in (1, 2, 3)] + valcells
        merges.append(f"A{rn}:D{rn}"); merges.append(f"E{rn}:H{rn}")
        return f'<row r="{rn}" ht="{ht}" customHeight="1">' + "".join(cells) + "</row>"

    def val_block(rn, style, formula=None, cached=None, number=None):
        top = (cf(f"E{rn}", style, formula, cached) if formula is not None
               else cn(f"E{rn}", style, number))
        return [top] + [ce(f"{COLS[i]}{rn}", style) for i in (5, 6, 7)]

    # ---- Key Metrics (with a live "figures as of" date on the right) --------
    krow = [ct("A4", S["subsection"], "\U0001F511 Key Metrics")]
    krow += [ce(f"{colL(i)}4", S["subsection"]) for i in range(1, 8)]
    krow.append(cf_str("I4", S["subsection_r"], '"Figures as of "&TEXT(TODAY(),"mmm d, yyyy")', "Figures as of today"))
    krow += [ce(f"{colL(i)}4", S["subsection_r"]) for i in range(9, 13)]
    rows.append('<row r="4" ht="20" customHeight="1">' + "".join(krow) + "</row>")
    merges.append("A4:H4"); merges.append("I4:M4")
    metrics = [
        ("Total Balance Today (incl. interest)", "SUM(Balance!$H$6:$H$19)", round(KPI_balance_today, 2)),
        ("Total Deposited (All-Time)",           "SUM(Balance!$D$6:$D$19)", round(KPI_total_deposits, 2)),
        ("Interest Earned to Date (auto)",       "SUM(Balance!$G$6:$G$19)", round(KPI_interest_to_date, 2)),
        ("Total Withdrawn",                      '-SUMIFS(Transactions!$E$6:$E$205,Transactions!$B$6:$B$205,"Withdrawal")', round(total_withdrawn, 2)),
    ]
    r = 5
    for (lbl, formula, cached) in metrics:
        rows.append(two_col(r, lbl, val_block(r, S["m_val"], formula=formula, cached=cached)))
        r += 1
    rows.append('<row r="9" ht="6" customHeight="1"/>')

    # ---- Personal Goal Tracker ---------------------------------------------
    rows.append(band_row(10, S["subsection"], "\U0001F3AF Personal Goal Tracker")); merges.append("A10:M10")
    rows.append(two_col(11, "Your Savings Goal Target",
                        val_block(11, S["m_val_input"], number=GOAL_TARGET)))
    rows.append(two_col(12, "Maya Personal Goal \u2014 Balance Today",
                        val_block(12, S["m_val"],
                                  formula='IFERROR(VLOOKUP("Maya - Personal Goal",Balance!$C$6:$H$19,6,FALSE),0)',
                                  cached=round(calc["Maya - Personal Goal"]["bal_today"], 2))))
    rows.append(two_col(13, "Progress Toward Goal",
                        val_block(13, S["m_val_pct"], formula="IFERROR(E12/E11,0)",
                                  cached=round(goal_progress, 4))))
    rows.append('<row r="14" ht="6" customHeight="1"/>')

    # ---- Balance by Bank (also the pie source) ------------------------------
    rows.append(band_row(15, S["subsection"], "\U0001F3E6 Balance by Bank")); merges.append("A15:M15")
    rows.append('<row r="16" ht="18" customHeight="1">'
                + ct("A16", S["thead"], "Bank") + ct("B16", S["thead"], "Balance") + "</row>")
    rr = 17
    for b in BANKS:
        rows.append(f'<row r="{rr}">' + ct(f"A{rr}", S["txt"], b)
                    + cf(f"B{rr}", S["bank_val"],
                         f"SUMIF(Balance!$A$6:$A$19,A{rr},Balance!$H$6:$H$19)",
                         round(bank_today[b], 2)) + "</row>")
        rr += 1
    rows.append('<row r="24">' + ct("A24", S["tot_txt"], "Total")
                + cf("B24", S["tot_curr"], "SUM(B17:B23)", round(sum(bank_today.values()), 2)) + "</row>")

    # ---- chart section titles (charts float over the empty rows) ------------
    rows.append(band_row(27, S["subsection"], "\U0001F4C8 Savings Growth by Bank (Cumulative Balance)")); merges.append("A27:M27")
    rows.append(band_row(51, S["subsection"], "\U0001F4CA Monthly Interest Earned by Bank")); merges.append("A51:M51")

    # ---- Chart data matrix: 24 months x 7 banks + total interest -----------
    rows.append(band_row(77, S["section"], "\U0001F5C2 Chart Data (auto-calculated \u2014 do not edit)")); merges.append("A77:M77")
    hdr = [ct("A79", S["dhead"], "Month")]
    for i, b in enumerate(BANKS):
        hdr.append(ct(f"{COLS[i+1]}79", S["dhead"], b))
    hdr.append(ct("I79", S["dhead"], "Total Interest"))
    rows.append('<row r="79">' + "".join(hdr) + "</row>")
    for m in range(24):
        r = 80 + m
        cells = [ct(f"A{r}", S["txt"], MONTH_LABELS[m])]
        for i, b in enumerate(BANKS):
            col = COLS[i + 1]
            cells.append(cf(f"{col}{r}", S["curr"],
                            f"SUMPRODUCT((Interest!$A$6:$A$19={col}$79)*Interest!$G$6:$G$19*Interest!$M$6:$M$19^{m})",
                            round(bank_series[b][m], 2)))
        if m == 0:
            cells.append(cn(f"I{r}", S["curr"], 0))
        else:
            cells.append(cf(f"I{r}", S["curr"], f"SUM(B{r}:H{r})-SUM(B{r-1}:H{r-1})",
                            round(monthly_int24[m], 2)))
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")

    # ---- second matrix: monthly interest earned PER BANK (stacked column) ---
    rows.append(band_row(105, S["subsection"], "\U0001F5C2 Monthly Interest Earned by Bank (auto)")); merges.append("A105:M105")
    hdr2 = [ct("A106", S["dhead"], "Month")]
    for i, b in enumerate(BANKS):
        hdr2.append(ct(f"{COLS[i+1]}106", S["dhead"], b))
    rows.append('<row r="106">' + "".join(hdr2) + "</row>")
    for m in range(24):
        r = 107 + m          # cumulative counterpart is at row (r - 27)
        cells = [ct(f"A{r}", S["txt"], MONTH_LABELS[m])]
        for i, b in enumerate(BANKS):
            col = COLS[i + 1]
            if m == 0:
                cells.append(cn(f"{col}{r}", S["curr"], 0))
            else:
                cells.append(cf(f"{col}{r}", S["curr"], f"{col}{r-27}-{col}{r-28}",
                                round(bank_monthly_int[b][m], 2)))
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")

    # wide enough for "PHP ###,###.##" currency in the chart-data matrix
    widths = [20, 16, 16, 16, 16, 16, 16, 16, 16, 12, 12, 12, 12]
    colsxml = cols_xml(widths)
    return worksheet(rows, merges, hyper, 3, colsxml, tab_selected=True,
                     drawing_rid="rId1", cond_fmt=databar("E13"))

# ============================================================================
# SHEET: DEPOSIT (sheet2)  -- Excel Table with filters
# ============================================================================
def build_deposit():
    rows, merges, hyper = sheet_open("Deposit",
        "Log every deposit here. Type a new deposit on the next empty row of the table \u2014 balances update automatically. Use the filter arrows to search.", 5)
    rows.append(band_row(4, S["section"], "\u2795 Add / View Deposits  (enter new deposits on the next empty row)"))
    merges.append("A4:M4")
    # Table header row 5
    heads = ["Date", "Bank", "Account", "Amount", "Notes"]
    hcells = [ct(f"{COLS[i]}5", S["thead"], heads[i]) for i in range(5)]
    for i in range(5, 13):
        hcells.append(ce(f"{COLS[i]}5", S["thead"]))
    rows.append('<row r="5" ht="20" customHeight="1">' + "".join(hcells) + "</row>")
    # Seed rows 6..19 then empty to 205
    for idx, (ds, bank, lbl, amt, note) in enumerate(DEP_ROWS):
        r = 6 + idx
        rows.append(f'<row r="{r}">' +
            cn(f"A{r}", S["date_col"], ds) + ct(f"B{r}", S["txt_col"], bank) +
            ct(f"C{r}", S["txt_col"], lbl) + cn(f"D{r}", S["curr_col"], amt) +
            ct(f"E{r}", S["txt_col"], note) + "</row>")
    # table range A5:E205
    table_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'id="1" name="tblDeposits" displayName="tblDeposits" ref="A5:E205" totalsRowShown="0">'
        '<autoFilter ref="A5:E205"/>'
        '<tableColumns count="5">'
        '<tableColumn id="1" name="Date"/><tableColumn id="2" name="Bank"/>'
        '<tableColumn id="3" name="Account"/><tableColumn id="4" name="Amount"/>'
        '<tableColumn id="5" name="Notes"/></tableColumns>'
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
        'showRowStripes="1" showColumnStripes="0"/></table>')
    specs = [(13, S["date_col"]), (12, S["txt_col"]), (24, S["txt_col"]),
             (15, S["curr_col"]), (28, S["txt_col"]), (10, None), (10, None),
             (10, None), (12, None), (12, None), (12, None), (12, None), (12, None)]
    colsxml = cols_xml_styled(specs)
    vals = [
        dv_list("B6:B205", "BankList", "Pick a bank", "Choose the bank from the drop-down list."),
        dv_list("C6:C205", "AccountList", "Pick an account", "Choose the account from the drop-down list so the balance matches correctly."),
    ]
    ws = worksheet(rows, merges, hyper, 5, colsxml, tablepart_rid="rId1", validations=vals)
    return ws, table_xml

# ============================================================================
# SHEET: TRANSACTIONS (sheet3) -- Excel Table with filters
# ============================================================================
def build_transactions():
    rows, merges, hyper = sheet_open("Transactions",
        "All money movements: deposits, withdrawals, transfers and interest postings. Use negative amounts for money out. Filter by any column.", 5)
    rows.append(band_row(4, S["section"], "\U0001F504 Transaction Log  (use the filter arrows to search / sort)"))
    merges.append("A4:M4")
    heads = ["Date", "Type", "Bank", "Account", "Amount", "Notes"]
    hcells = [ct(f"{COLS[i]}5", S["thead"], heads[i]) for i in range(6)]
    for i in range(6, 13):
        hcells.append(ce(f"{COLS[i]}5", S["thead"]))
    rows.append('<row r="5" ht="20" customHeight="1">' + "".join(hcells) + "</row>")
    for idx, (ds, ty, bank, lbl, amt, note) in enumerate(TX_ROWS):
        r = 6 + idx
        rows.append(f'<row r="{r}">' +
            cn(f"A{r}", S["date_col"], ds) + ct(f"B{r}", S["txt_col"], ty) +
            ct(f"C{r}", S["txt_col"], bank) + ct(f"D{r}", S["txt_col"], lbl) +
            cn(f"E{r}", S["curr_col"], amt) + ct(f"F{r}", S["txt_col"], note) + "</row>")
    table_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'id="2" name="tblTransactions" displayName="tblTransactions" ref="A5:F205" totalsRowShown="0">'
        '<autoFilter ref="A5:F205"/>'
        '<tableColumns count="6">'
        '<tableColumn id="1" name="Date"/><tableColumn id="2" name="Type"/>'
        '<tableColumn id="3" name="Bank"/><tableColumn id="4" name="Account"/>'
        '<tableColumn id="5" name="Amount"/><tableColumn id="6" name="Notes"/></tableColumns>'
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
        'showRowStripes="1" showColumnStripes="0"/></table>')
    specs = [(13, S["date_col"]), (14, S["txt_col"]), (12, S["txt_col"]),
             (24, S["txt_col"]), (15, S["curr_col"]), (28, S["txt_col"]),
             (10, None), (10, None), (12, None), (12, None), (12, None), (12, None), (12, None)]
    colsxml = cols_xml_styled(specs)
    vals = [
        dv_list("B6:B205", '"Deposit,Withdrawal,Transfer,Adjustment,Fee"', "Pick a type",
                "Choose the transaction type. (Interest is added automatically - no need to record it here.)"),
        dv_list("C6:C205", "BankList", "Pick a bank", "Choose the bank from the drop-down list."),
        dv_list("D6:D205", "AccountList", "Pick an account", "Choose the account from the drop-down list."),
    ]
    ws = worksheet(rows, merges, hyper, 5, colsxml, tablepart_rid="rId1", validations=vals)
    return ws, table_xml

# ============================================================================
# SHEET: BALANCE (sheet4)
# ============================================================================
def build_balance():
    rows, merges, hyper = sheet_open("Balance",
        "Balance Today grows automatically with interest each day/month. Current Balance = Deposits + Net Transactions (money you moved); Balance Today adds the interest accrued up to today.", 5)
    rows.append(band_row(4, S["section"], "\U0001F4BC Balances by Account"))
    merges.append("A4:M4")
    heads = ["Bank", "Account", "Account Label", "Deposits", "Net Transactions",
             "Current Balance", "Interest to Date", "Balance Today",
             "Interest (1 Yr)", "Proj. Balance (1 Yr)"]
    hcells = [ct(f"{COLS[i]}5", S["thead"], heads[i]) for i in range(10)]
    for i in range(10, 13):
        hcells.append(ce(f"{COLS[i]}5", S["thead"]))
    rows.append('<row r="5" ht="28" customHeight="1">' + "".join(hcells) + "</row>")
    for idx, (bank, acct, atype, lbl, rate, freq) in enumerate(ACCOUNTS):
        r = 6 + idx; c = calc[lbl]
        cells = [
            ct(f"A{r}", S["txt"], bank), ct(f"B{r}", S["txt"], acct),
            ct(f"C{r}", S["txt"], lbl),
            cf(f"D{r}", S["curr"], f"SUMIFS(Deposit!$D$6:$D$205,Deposit!$C$6:$C$205,C{r})", round(c["dep"], 2)),
            cf(f"E{r}", S["curr"], f"SUMIFS(Transactions!$E$6:$E$205,Transactions!$D$6:$D$205,C{r})", round(c["ntx"], 2)),
            cf(f"F{r}", S["curr"], f"D{r}+E{r}", round(c["principal"], 2)),
            cf(f"G{r}", S["green_curr"], f"IFERROR(VLOOKUP(C{r},Interest!$D$6:$Q$19,13,FALSE),0)", round(c["i2d"], 2)),
            cf(f"H{r}", S["green_curr"], f"IFERROR(VLOOKUP(C{r},Interest!$D$6:$Q$19,14,FALSE),0)", round(c["bal_today"], 2)),
            cf(f"I{r}", S["curr"], f"IFERROR(VLOOKUP(C{r},Interest!$D$6:$L$19,9,FALSE),0)", round(c["int1yr"], 2)),
            cf(f"J{r}", S["curr"], f"F{r}+I{r}", round(c["principal"] + c["int1yr"], 2)),
        ]
        for i in range(10, 13):
            cells.append(ce(f"{COLS[i]}{r}", S["txt_col"]))
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    # totals row 20
    tr = 20
    totmap = {'D': 'dep', 'E': 'ntx', 'F': 'principal', 'G': 'i2d', 'H': 'bal_today', 'I': 'int1yr'}
    tcells = [ct(f"A{tr}", S["tot_txt"], "TOTAL")] + [ce(f"{COLS[i]}{tr}", S["tot_txt"]) for i in (1, 2)]
    for col in "DEFGHIJ":
        if col == 'J':
            val = sum(calc[l]['principal'] + calc[l]['int1yr'] for l in dep_by_label)
        else:
            val = sum(calc[l][totmap[col]] for l in dep_by_label)
        tcells.append(cf(f"{col}{tr}", S["tot_curr"], f"SUM({col}6:{col}19)", round(val, 2)))
    for i in range(10, 13):
        tcells.append(ce(f"{COLS[i]}{tr}", S["tot_curr"]))
    rows.append(f'<row r="{tr}" ht="18" customHeight="1">' + "".join(tcells) + "</row>")

    widths = [12, 16, 24, 15, 16, 16, 16, 17, 15, 17, 10, 12, 12]
    colsxml = cols_xml(widths)
    return worksheet(rows, merges, hyper, 5, colsxml)

# ============================================================================
# SHEET: INTEREST (sheet5)
# ============================================================================
def build_interest():
    NC = 17  # A..Q
    rows, merges, hyper = sheet_open("Interest",
        "Interest auto-calculates and GROWS ON ITS OWN. 'Balance Today' accrues from the yellow 'Interest Start' "
        "date up to TODAY() every time you open the file - daily accounts step up each day, monthly accounts each "
        "month. Interest Start is pre-filled with your first deposit date; edit it if an account started elsewhere. Maya Savings = 10%.", 5, ncols=NC)
    rows.append(band_row(4, S["section"], "\U0001F4C8 Interest Rates, Earnings & Live Balance", ncols=NC))
    merges.append("A4:Q4")
    heads = ["Bank", "Account", "Account Type", "Account Label", "Annual Rate",
             "Frequency", "Current Balance", "Periodic Rate", "Interest / Period",
             "Periods / Yr", "Proj. Balance (1 Yr)", "Interest (1 Yr)", "Monthly Factor",
             "Interest Start", "Periods Elapsed", "Interest to Date", "Balance Today"]
    hcells = [ct(f"{colL(i)}5", S["thead"], heads[i]) for i in range(NC)]
    rows.append('<row r="5" ht="30" customHeight="1">' + "".join(hcells) + "</row>")
    for idx, (bank, acct, atype, lbl, rate, freq) in enumerate(ACCOUNTS):
        r = 6 + idx; c = calc[lbl]
        cells = [
            ct(f"A{r}", S["txt"], bank), ct(f"B{r}", S["txt"], acct),
            ct(f"C{r}", S["txt"], atype), ct(f"D{r}", S["txt"], lbl),
            cn(f"E{r}", S["pct"], rate), ct(f"F{r}", S["txt_c"], freq),
            cf(f"G{r}", S["curr"], f"IFERROR(VLOOKUP(D{r},Balance!$C$6:$H$19,4,FALSE),0)", round(c["principal"], 2)),
            cf(f"H{r}", S["pct"], f'IF(F{r}="Daily",E{r}/365,E{r}/12)', round(c["pr"], 8)),
            cf(f"I{r}", S["curr"], f"G{r}*H{r}", round(c["iper"], 2)),
            cf(f"J{r}", S["intc"], f'IF(F{r}="Daily",365,12)', c["n"]),
            cf(f"K{r}", S["curr"], f"G{r}*(1+H{r})^J{r}", round(c["projbal"], 2)),
            cf(f"L{r}", S["curr"], f"K{r}-G{r}", round(c["int1yr"], 2)),
            cf(f"M{r}", S["pct"], f'IF(F{r}="Daily",(1+E{r}/365)^(365/12),1+E{r}/12)', round(c["mfactor"], 8)),
            # ---- live accrual (grows with TODAY) ----
            # Interest Start = a plain EDITABLE date (yellow), pre-filled with the first
            # deposit date. No lookup functions -> works on every Excel version.
            cn(f"N{r}", S["date_input"], c["start"]),
            cf(f"O{r}", S["intc"], f'IF(ISNUMBER(N{r}),IF(F{r}="Daily",MAX(0,TODAY()-N{r}),MAX(0,DATEDIF(N{r},TODAY(),"m"))),0)', c["elapsed"]),
            cf(f"P{r}", S["curr"], f"G{r}*((1+H{r})^O{r}-1)", round(c["i2d"], 2)),
            cf(f"Q{r}", S["curr"], f"G{r}+P{r}", round(c["bal_today"], 2)),
        ]
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    tr = 20
    def totcur(col, val):
        return cf(f"{col}{tr}", S["tot_curr"], f"SUM({col}6:{col}19)", round(val, 2))
    tcells = [ct(f"A{tr}", S["tot_txt"], "TOTAL")] + [ce(f"{colL(i)}{tr}", S["tot_txt"]) for i in range(1, 6)]
    tcells.append(totcur("G", sum(c['principal'] for c in calc.values())))
    tcells.append(ce(f"H{tr}", S["tot_txt"]))
    tcells.append(totcur("I", sum(c['iper'] for c in calc.values())))
    tcells.append(ce(f"J{tr}", S["tot_txt"]))
    tcells.append(totcur("K", sum(c['projbal'] for c in calc.values())))
    tcells.append(totcur("L", sum(c['int1yr'] for c in calc.values())))
    tcells += [ce(f"M{tr}", S["tot_txt"]), ce(f"N{tr}", S["tot_txt"]), ce(f"O{tr}", S["tot_txt"])]
    tcells.append(totcur("P", sum(c['i2d'] for c in calc.values())))
    tcells.append(totcur("Q", sum(c['bal_today'] for c in calc.values())))
    rows.append(f'<row r="{tr}" ht="18" customHeight="1">' + "".join(tcells) + "</row>")

    widths = [11, 15, 13, 22, 11, 11, 15, 11, 15, 10, 16, 15, 12, 13, 12, 16, 16]
    # Hide the internal calc columns so the sheet fits on screen (no horizontal
    # scroll -> the top nav buttons stay visible & clickable):
    # H Periodic Rate(7), J Periods/Yr(9), M Monthly Factor(12), O Periods Elapsed(14)
    hidden = {7, 9, 12, 14}
    _cols = []
    for i, w in enumerate(widths):
        hid = ' hidden="1"' if i in hidden else ''
        _cols.append(f'<col min="{i+1}" max="{i+1}" width="{w}"{hid} customWidth="1"/>')
    colsxml = "<cols>" + "".join(_cols) + "</cols>"
    return worksheet(rows, merges, hyper, 5, colsxml)

# ============================================================================
# CHARTS + DRAWING
# ============================================================================
CNS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def num_cache(vals, fmt="General"):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{round(v,4)}</c:v></c:pt>' for i, v in enumerate(vals))
    return f'<c:numCache><c:formatCode>{fmt}</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numCache>'
def str_cache(vals):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{esc(v)}</c:v></c:pt>' for i, v in enumerate(vals))
    return f'<c:strCache><c:ptCount val="{len(vals)}"/>{pts}</c:strCache>'

def title_el(text):
    return (f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1">'
            f'<a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')

# ---- shared chart styling helpers ------------------------------------------
def axis_title(text, vertical=False):
    rot = ' rot="-5400000" vert="horz"' if vertical else ''
    return (f'<c:title><c:tx><c:rich><a:bodyPr{rot}/><a:lstStyle/>'
            f'<a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="900" b="1"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')

# subtle horizontal gridlines
GRIDLINES = ('<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill>'
             '<a:srgbClr val="E7E7E7"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>')
# rotated, smaller month labels (24 categories)
CAT_TXPR = ('<c:txPr><a:bodyPr rot="-2700000" spcFirstLastPara="1" vertOverflow="ellipsis" '
            'vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
            '<a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>')
PESO_FMT = '<c:numFmt formatCode="[$\u20b1-3409]#,##0" sourceLinked="0"/>'

def chart_pie(title, catf, catv, valf, valv, sername):
    ser = (f'<c:ser><c:idx val="0"/><c:order val="0"/>'
           f'<c:tx><c:strRef><c:f>{sername[0]}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(sername[1])}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
           f'<c:cat><c:strRef><c:f>{catf}</c:f>{str_cache(catv)}</c:strRef></c:cat>'
           f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        f'<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>{ser}'
        '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
        '<c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>'
        '<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>'
        '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def chart_line(title, catf, catv, valf, valv, sername):
    ser = (f'<c:ser><c:idx val="0"/><c:order val="0"/>'
           f'<c:tx><c:strRef><c:f>{sername[0]}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(sername[1])}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
           '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="2E5496"/></a:solidFill></a:ln></c:spPr>'
           '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>'
           f'<c:cat><c:numRef><c:f>{catf}</c:f>{num_cache(catv)}</c:numRef></c:cat>'
           f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val>'
           '<c:smooth val="0"/></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        + ser +
        '<c:marker val="1"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>'
        '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:crossAx val="222"/></c:catAx>'
        '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="111"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def chart_bar(title, catf, catv, valf, valv, sername):
    ser = (f'<c:ser><c:idx val="0"/><c:order val="0"/>'
           f'<c:tx><c:strRef><c:f>{sername[0]}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(sername[1])}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
           '<c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill></c:spPr>'
           f'<c:cat><c:numRef><c:f>{catf}</c:f>{num_cache(catv)}</c:numRef></c:cat>'
           f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        + ser +
        '<c:axId val="333"/><c:axId val="444"/></c:barChart>'
        '<c:catAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:crossAx val="444"/></c:catAx>'
        '<c:valAx><c:axId val="444"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="333"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def chart_line_multi(title, catf, catlabels, series):
    """series: list of (name_cell_ref, name, val_ref, val_cache_list, colorhex)."""
    sers = []
    for i, (namecell, name, valf, valv, color) in enumerate(series):
        sers.append(
            f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx><c:strRef><c:f>{namecell}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            f'<c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker>'
            f'<c:cat><c:strRef><c:f>{catf}</c:f>{str_cache(catlabels)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val>'
            f'<c:smooth val="0"/></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:roundedCorners val="0"/><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        + "".join(sers) +
        '<c:marker val="1"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>'
        '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/>' + axis_title("Month") + CAT_TXPR +
        '<c:crossAx val="222"/><c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/>' + GRIDLINES + axis_title("Cumulative Balance", vertical=True) + PESO_FMT +
        '<c:crossAx val="111"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def chart_bar_stacked(title, catf, catlabels, series):
    """Stacked column chart: one series (segment) per bank. series: list of
    (name_cell, name, val_ref, val_cache_list, colorhex)."""
    sers = []
    for i, (namecell, name, valf, valv, color) in enumerate(series):
        sers.append(
            f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx><c:strRef><c:f>{namecell}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>'
            f'<c:cat><c:strRef><c:f>{catf}</c:f>{str_cache(catlabels)}</c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:roundedCorners val="0"/><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
        + "".join(sers) +
        '<c:gapWidth val="40"/><c:overlap val="100"/><c:axId val="333"/><c:axId val="444"/></c:barChart>'
        '<c:catAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/>' + axis_title("Month") + CAT_TXPR +
        '<c:crossAx val="444"/><c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        '<c:valAx><c:axId val="444"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/>' + GRIDLINES + axis_title("Interest Earned", vertical=True) + PESO_FMT +
        '<c:crossAx val="333"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def chart_bar_cat(title, catf, catlabels, valf, valv, sername):
    ser = (f'<c:ser><c:idx val="0"/><c:order val="0"/>'
           f'<c:tx><c:strRef><c:f>{sername[0]}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(sername[1])}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
           '<c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill></c:spPr>'
           f'<c:cat><c:strRef><c:f>{catf}</c:f>{str_cache(catlabels)}</c:strRef></c:cat>'
           f'<c:val><c:numRef><c:f>{valf}</c:f>{num_cache(valv)}</c:numRef></c:val></c:ser>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<c:chartSpace {CNS}><c:chart>{title_el(title)}<c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        + ser +
        '<c:axId val="333"/><c:axId val="444"/></c:barChart>'
        '<c:catAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:crossAx val="444"/></c:catAx>'
        '<c:valAx><c:axId val="444"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:r><a:t>Interest (PHP)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="333"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

def anchor(frm, to, rid, name):
    fc, fr = frm; tc, tr = to
    return (f'<xdr:twoCellAnchor editAs="oneCell">'
        f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
        f'<xdr:cNvPr id="{rid}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
        f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{rid-1}"/>'
        f'</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')

def build_drawing():
    # chart ids: cNvPr id = 2,3,4 -> r:id = rId1,rId2,rId3
    a1 = anchor((0, 27), (12, 50), 2, "Savings Growth")   # line, under its title (row 28-50)
    a2 = anchor((3, 14), (11, 26), 3, "Savings by Bank")  # pie, beside Balance-by-Bank table
    a3 = anchor((0, 51), (12, 74), 4, "Monthly Interest") # column, under its title (row 52-74)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + a1 + a2 + a3 + '</xdr:wsDr>')

# ----------------------------------------------------------------------------
# Assemble package
# ----------------------------------------------------------------------------
dash = build_dashboard()
dep_ws, dep_tbl = build_deposit()
tx_ws, tx_tbl = build_transactions()
bal = build_balance()
intr = build_interest()

growth_series = [(f"Dashboard!${c}$79", b, f"Dashboard!${c}$80:${c}$103", bank_series[b], BANK_COLORS[b])
                 for c, b in zip("BCDEFGH", BANKS)]
chart1 = chart_line_multi("Savings Growth by Bank (Cumulative Balance)",
                          "Dashboard!$A$80:$A$103", MONTH_LABELS, growth_series)
chart2 = chart_pie("\U0001F967 Savings by Bank",
                   "Dashboard!$A$17:$A$23", BANKS, "Dashboard!$B$17:$B$23",
                   [round(bank_today[b], 2) for b in BANKS], ("Dashboard!$B$16", "Balance"))
interest_series = [(f"Dashboard!${c}$106", b, f"Dashboard!${c}$107:${c}$130", bank_monthly_int[b], BANK_COLORS[b])
                   for c, b in zip("BCDEFGH", BANKS)]
chart3 = chart_bar_stacked("Monthly Interest Earned by Bank",
                           "Dashboard!$A$107:$A$130", MONTH_LABELS, interest_series)
drawing = build_drawing()

workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    '<sheet name="Dashboard" sheetId="1" r:id="rId1"/>'
    '<sheet name="Deposit" sheetId="2" r:id="rId2"/>'
    '<sheet name="Transactions" sheetId="3" r:id="rId3"/>'
    '<sheet name="Balance" sheetId="4" r:id="rId4"/>'
    '<sheet name="Interest" sheetId="5" r:id="rId5"/>'
    '</sheets>'
    '<definedNames>'
    '<definedName name="BankList">Dashboard!$A$17:$A$23</definedName>'
    '<definedName name="AccountList">Interest!$D$6:$D$19</definedName>'
    '</definedNames>'
    '<calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1"/></workbook>')

workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>'
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>'
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>')

sheet1_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
    '</Relationships>')
sheet2_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
    '</Relationships>')
sheet3_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>'
    '</Relationships>')
drawing_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>'
    '</Relationships>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 6))
    + '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    + "".join(f'<Override PartName="/xl/charts/chart{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' for i in range(1, 4))
    + '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
    + '<Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
    + '</Types>')

root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>')

OUT = "/projects/sandbox/Kiro/Digital-Bank-Savings-Tracker.xlsx"
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("xl/workbook.xml", workbook_xml)
    z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
    z.writestr("xl/styles.xml", STYLES_XML)
    z.writestr("xl/worksheets/sheet1.xml", dash)
    z.writestr("xl/worksheets/sheet2.xml", dep_ws)
    z.writestr("xl/worksheets/sheet3.xml", tx_ws)
    z.writestr("xl/worksheets/sheet4.xml", bal)
    z.writestr("xl/worksheets/sheet5.xml", intr)
    z.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet1_rels)
    z.writestr("xl/worksheets/_rels/sheet2.xml.rels", sheet2_rels)
    z.writestr("xl/worksheets/_rels/sheet3.xml.rels", sheet3_rels)
    z.writestr("xl/drawings/drawing1.xml", drawing)
    z.writestr("xl/drawings/_rels/drawing1.xml.rels", drawing_rels)
    z.writestr("xl/charts/chart1.xml", chart1)
    z.writestr("xl/charts/chart2.xml", chart2)
    z.writestr("xl/charts/chart3.xml", chart3)
    z.writestr("xl/tables/table1.xml", dep_tbl)
    z.writestr("xl/tables/table2.xml", tx_tbl)

print("Wrote", OUT)
print(f"Totals: savings={KPI_total_savings:.2f} deposits={KPI_total_deposits:.2f} "
      f"proj_interest={KPI_proj_interest:.2f} proj_balance={KPI_proj_balance:.2f}")
