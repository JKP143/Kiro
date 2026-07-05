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
    (serial(2026, 6, 20), "Interest",   "Banko", "Banko - Savings",     62.50, "Interest posting (example)"),
    (serial(2026, 6, 25), "Transfer",   "UNO",   "UNO - Digi Savings",  5000,  "Transfer in (example)"),
]

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
    calc[lbl] = dict(bank=bank, dep=dep, ntx=ntx, principal=principal, rate=rate,
                     freq=freq, pr=pr, iper=iper, n=n, projbal=projbal,
                     int1yr=int1yr, mfactor=mfactor)

bank_totals = {b: 0.0 for b in BANKS}
for lbl, c in calc.items():
    bank_totals[c["bank"]] += c["principal"]

MONTHS = list(range(0, 13))
def total_at(m):
    return sum(c["principal"] * (c["mfactor"] ** m) for c in calc.values())
growth_total = [total_at(m) for m in MONTHS]
monthly_interest = [0.0] + [growth_total[i] - growth_total[i - 1] for i in range(1, len(MONTHS))]

KPI_total_savings = sum(c["principal"] for c in calc.values())
KPI_total_deposits = sum(c["dep"] for c in calc.values())
KPI_proj_interest = sum(c["int1yr"] for c in calc.values())
KPI_proj_balance = sum(c["projbal"] for c in calc.values())

# ----------------------------------------------------------------------------
# XML helpers
# ----------------------------------------------------------------------------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

COLS = "ABCDEFGHIJKLM"  # 13 columns

def ct(ref, s, val):    # inline text cell
    return f'<c r="{ref}" t="inlineStr" s="{s}"><is><t xml:space="preserve">{esc(val)}</t></is></c>'
def cn(ref, s, val):    # number cell
    return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
def cf(ref, s, formula, cached=None):   # formula cell (+ optional cached value)
    v = f'<v>{cached}</v>' if cached is not None else ''
    return f'<c r="{ref}" s="{s}"><f>{esc(formula)}</f>{v}</c>'
def ce(ref, s):         # empty styled cell
    return f'<c r="{ref}" s="{s}"/>'

def band_row(rownum, style, text=None, text_style=None, ncols=13):
    """A full-width colored band row across columns A..(ncols)."""
    cells = []
    for i in range(ncols):
        ref = f"{COLS[i]}{rownum}"
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

def nav_row(active):
    cells, hyper, merges = [], [], []
    # button pairs A..J
    for idx, (sheet, label) in enumerate(NAV):
        left = COLS[idx * 2]; right = COLS[idx * 2 + 1]
        s = S["nav_active"] if sheet == active else S["nav"]
        cells.append(ct(f"{left}2", s, label))
        cells.append(ce(f"{right}2", s))
        merges.append(f"{left}2:{right}2")
        hyper.append(f'<hyperlink ref="{left}2:{right}2" location="{sheet}!A1" display="{esc(label)}"/>')
    # filler K,L,M navy
    for i in range(10, 13):
        cells.append(ce(f"{COLS[i]}2", S["nav"]))
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

STYLES_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    f'<numFmts count="1"><numFmt numFmtId="{CURR}" formatCode="&quot;PHP &quot;#,##0.00"/></numFmts>'
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
def sheet_open(active, subtitle, ysplit, tab_selected=False,
               colwidths=None, extra_sheetpr="", firstcol_end=13):
    """Return (rows_list, merges_list, hyper_list) pre-populated with the
    title band, nav bar and subtitle band."""
    rows, merges, hyper = [], [], []
    # Row 1 title
    rows.append(band_row(1, S["title"], "\U0001F4B0 Digital Bank Savings Tracker",
                         text_style=S["title"]))
    rows[-1] = rows[-1].replace('ht="22"', 'ht="30"')
    merges.append("A1:M1")
    # Row 2 nav
    nrow, nmerges, nhyper = nav_row(active)
    rows.append(nrow); merges += nmerges; hyper += nhyper
    # Row 3 subtitle
    rows.append(band_row(3, S["subtitle"], subtitle, text_style=S["subtitle"]))
    merges.append("A3:M3")
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

def worksheet(rows, merges, hyper, ysplit, colsxml, tab_selected=False,
              drawing_rid=None, tablepart_rid=None):
    sel = ' tabSelected="1"' if tab_selected else ''
    top = ysplit + 1
    view = (f'<sheetViews><sheetView showGridLines="0"{sel} workbookViewId="0">'
            f'<pane ySplit="{ysplit}" topLeftCell="A{top}" activePane="bottomLeft" state="frozen"/>'
            f'<selection pane="bottomLeft" activeCell="A{top}" sqref="A{top}"/></sheetView></sheetViews>')
    mc = ''
    if merges:
        mc = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>'
    hl = ''
    if hyper:
        hl = '<hyperlinks>' + "".join(hyper) + '</hyperlinks>'
    dr = f'<drawing r:id="{drawing_rid}"/>' if drawing_rid else ''
    tp = f'<tableParts count="1"><tablePart r:id="{tablepart_rid}"/></tableParts>' if tablepart_rid else ''
    data = "".join(rows)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + view + '<sheetFormatPr defaultRowHeight="15"/>' + colsxml
        + '<sheetData>' + data + '</sheetData>' + mc + hl + dr + tp + '</worksheet>')

# ============================================================================
# SHEET: DASHBOARD (sheet1)
# ============================================================================
def build_dashboard():
    rows, merges, hyper = sheet_open("Dashboard",
        "Overview of your savings across all digital banks. Figures update automatically from the Deposit, Transactions and Interest tabs.", 2, True)
    # Row 4 section band
    rows.append(band_row(4, S["section"], "\U0001F4CA Overview"))
    merges.append("A4:M4")
    # KPI cards: labels row5, values row6  (pairs A:B C:D E:F G:H)
    kpi_defs = [
        ("Total Savings",         S["kpi_curr"], "SUM(Balance!$F$6:$F$19)", round(KPI_total_savings, 2)),
        ("Total Deposits",        S["kpi_curr"], "SUM(Balance!$D$6:$D$19)", round(KPI_total_deposits, 2)),
        ("Proj. Interest (1 Yr)", S["kpi_curr"], "SUM(Balance!$G$6:$G$19)", round(KPI_proj_interest, 2)),
        ("Proj. Balance (1 Yr)",  S["kpi_curr"], "SUM(Balance!$H$6:$H$19)", round(KPI_proj_balance, 2)),
    ]
    lbl_cells, val_cells = [], []
    for i, (lbl, vs, formula, cached) in enumerate(kpi_defs):
        left = COLS[i * 2]; right = COLS[i * 2 + 1]
        lbl_cells.append(ct(f"{left}5", S["kpi_lbl"], lbl)); lbl_cells.append(ce(f"{right}5", S["kpi_lbl"]))
        val_cells.append(cf(f"{left}6", vs, formula, cached)); val_cells.append(ce(f"{right}6", vs))
        merges.append(f"{left}5:{right}5"); merges.append(f"{left}6:{right}6")
    # fill I,J,K,L,M for rows 5-6 with light gray so the band is clean
    for i in range(8, 13):
        lbl_cells.append(ce(f"{COLS[i]}5", S["kpi_lbl"]))
        val_cells.append(ce(f"{COLS[i]}6", S["kpi_lbl"]))
    rows.append(f'<row r="5" ht="18" customHeight="1">' + "".join(lbl_cells) + "</row>")
    rows.append(f'<row r="6" ht="34" customHeight="1">' + "".join(val_cells) + "</row>")
    # spacer row 7
    rows.append('<row r="7" ht="8" customHeight="1"/>')
    # (charts float over rows 8..49)

    # ---- Chart data region (rows 52+) --------------------------------------
    rows.append(band_row(52, S["section"], "\U0001F5C2 Chart Data (auto-calculated \u2014 do not edit)"))
    merges.append("A52:M52")
    # headers row 54
    hdr = [ct("A54", S["dhead"], "Bank"), ct("B54", S["dhead"], "Total Savings"),
           ce("C54", S["dhead"]),
           ct("D54", S["dhead"], "Month"), ct("E54", S["dhead"], "Total Savings"),
           ct("F54", S["dhead"], "Interest Earned")]
    for i in range(6, 13):
        hdr.append(ce(f"{COLS[i]}54", S["dhead"]))
    rows.append('<row r="54">' + "".join(hdr) + "</row>")
    # Bank table rows 55..61 ; Projection rows 55..67
    maxrow = 67
    for r in range(55, maxrow + 1):
        cells = []
        bi = r - 55  # 0..
        # bank table (only 7 banks -> rows 55..61)
        if 0 <= bi < len(BANKS):
            bank = BANKS[bi]
            cells.append(ct(f"A{r}", S["txt"], bank))
            cells.append(cf(f"B{r}", S["curr"],
                            f"SUMIF(Balance!$B$6:$B$19,A{r},Balance!$F$6:$F$19)",
                            round(bank_totals[bank], 2)))
        else:
            cells.append(ce(f"A{r}", S["txt"])); cells.append(ce(f"B{r}", S["curr"]))
        cells.append(ce(f"C{r}", S["txt"]))
        # projection rows 55..67 -> month 0..12
        mi = r - 55
        if 0 <= mi <= 12:
            cells.append(cn(f"D{r}", S["intc"], mi))
            cells.append(cf(f"E{r}", S["curr"],
                            f"SUMPRODUCT(Interest!$G$6:$G$19,Interest!$M$6:$M$19^D{r})",
                            round(growth_total[mi], 2)))
            if mi == 0:
                cells.append(cn(f"F{r}", S["curr"], 0))
            else:
                cells.append(cf(f"F{r}", S["curr"], f"E{r}-E{r-1}",
                                round(monthly_interest[mi], 2)))
        else:
            cells += [ce(f"D{r}", S["intc"]), ce(f"E{r}", S["curr"]), ce(f"F{r}", S["curr"])]
        for i in range(6, 13):
            cells.append(ce(f"{COLS[i]}{r}", S["txt_col"]))
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")

    widths = [16, 14, 4, 10, 16, 16, 12, 12, 12, 12, 12, 12, 12]
    colsxml = cols_xml(widths)
    return worksheet(rows, merges, hyper, 2, colsxml, tab_selected=True,
                     drawing_rid="rId1")

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
    ws = worksheet(rows, merges, hyper, 5, colsxml, tablepart_rid="rId1")
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
    ws = worksheet(rows, merges, hyper, 5, colsxml, tablepart_rid="rId1")
    return ws, table_xml

# ============================================================================
# SHEET: BALANCE (sheet4)
# ============================================================================
def build_balance():
    rows, merges, hyper = sheet_open("Balance",
        "Current balance per account = Deposits + Net Transactions. Projected columns include compounding interest from the Interest tab.", 5)
    rows.append(band_row(4, S["section"], "\U0001F4BC Balances by Account"))
    merges.append("A4:M4")
    heads = ["Bank", "Account", "Account Label", "Deposits", "Net Transactions",
             "Current Balance", "Interest (1 Yr)", "Proj. Balance (1 Yr)"]
    hcells = [ct(f"{COLS[i]}5", S["thead"], heads[i]) for i in range(8)]
    for i in range(8, 13):
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
            cf(f"G{r}", S["curr"], f"IFERROR(VLOOKUP(C{r},Interest!$D$6:$L$19,9,FALSE),0)", round(c["int1yr"], 2)),
            cf(f"H{r}", S["curr"], f"F{r}+G{r}", round(c["principal"] + c["int1yr"], 2)),
        ]
        for i in range(8, 13):
            cells.append(ce(f"{COLS[i]}{r}", S["txt_col"]))
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    # totals row 20
    tr = 20
    tcells = [ct(f"A{tr}", S["tot_txt"], "TOTAL")] + [ce(f"{COLS[i]}{tr}", S["tot_txt"]) for i in (1, 2)]
    for col in "DEFGH":
        tcells.append(cf(f"{col}{tr}", S["tot_curr"], f"SUM({col}6:{col}19)",
                         round(sum(calc[l][{'D':'dep','E':'ntx','F':'principal','G':'int1yr','H':'projbal'}[col]] if col!='H' else (calc[l]['principal']+calc[l]['int1yr']) for l in dep_by_label), 2) if col!='H' else round(sum(calc[l]['principal']+calc[l]['int1yr'] for l in dep_by_label),2)))
    for i in range(8, 13):
        tcells.append(ce(f"{COLS[i]}{tr}", S["tot_curr"]))
    rows.append(f'<row r="{tr}" ht="18" customHeight="1">' + "".join(tcells) + "</row>")

    widths = [12, 16, 24, 15, 16, 16, 16, 18, 10, 12, 12, 12, 12]
    colsxml = cols_xml(widths)
    return worksheet(rows, merges, hyper, 5, colsxml)

# ============================================================================
# SHEET: INTEREST (sheet5)
# ============================================================================
def build_interest():
    rows, merges, hyper = sheet_open("Interest",
        "Interest rates & auto-calculated earnings. Daily accounts compound daily (rate/365); monthly accounts compound monthly (rate/12). Maya Savings overridden to 10%.", 5)
    rows.append(band_row(4, S["section"], "\U0001F4C8 Interest Rates & Earnings"))
    merges.append("A4:M4")
    heads = ["Bank", "Account", "Account Type", "Account Label", "Annual Rate",
             "Frequency", "Current Balance", "Periodic Rate", "Interest / Period",
             "Periods / Yr", "Proj. Balance (1 Yr)", "Interest (1 Yr)", "Monthly Factor"]
    hcells = [ct(f"{COLS[i]}5", S["thead"], heads[i]) for i in range(13)]
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
        ]
        rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    tr = 20
    tcells = [ct(f"A{tr}", S["tot_txt"], "TOTAL")] + [ce(f"{COLS[i]}{tr}", S["tot_txt"]) for i in range(1, 6)]
    tcells.append(cf(f"G{tr}", S["tot_curr"], "SUM(G6:G19)", round(sum(c['principal'] for c in calc.values()), 2)))
    tcells.append(ce(f"H{tr}", S["tot_txt"]))
    tcells.append(cf(f"I{tr}", S["tot_curr"], "SUM(I6:I19)", round(sum(c['iper'] for c in calc.values()), 2)))
    tcells.append(ce(f"J{tr}", S["tot_txt"]))
    tcells.append(cf(f"K{tr}", S["tot_curr"], "SUM(K6:K19)", round(sum(c['projbal'] for c in calc.values()), 2)))
    tcells.append(cf(f"L{tr}", S["tot_curr"], "SUM(L6:L19)", round(sum(c['int1yr'] for c in calc.values()), 2)))
    tcells.append(ce(f"M{tr}", S["tot_txt"]))
    rows.append(f'<row r="{tr}" ht="18" customHeight="1">' + "".join(tcells) + "</row>")

    widths = [11, 15, 13, 24, 11, 11, 16, 12, 16, 11, 18, 16, 13]
    colsxml = cols_xml(widths)
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
    a1 = anchor((0, 7), (6, 27), 2, "Savings Growth")
    a2 = anchor((6, 7), (13, 27), 3, "Savings by Bank")
    a3 = anchor((0, 28), (13, 49), 4, "Monthly Interest")
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

chart1 = chart_line("\U0001F4C8 Savings Growth (12-Month Projection)",
                    "Dashboard!$D$55:$D$67", MONTHS, "Dashboard!$E$55:$E$67", growth_total,
                    ("Dashboard!$E$54", "Total Savings"))
chart2 = chart_pie("\U0001F967 Savings by Bank",
                   "Dashboard!$A$55:$A$61", BANKS, "Dashboard!$B$55:$B$61",
                   [round(bank_totals[b], 2) for b in BANKS], ("Dashboard!$B$54", "Total Savings"))
chart3 = chart_bar("\U0001F4CA Monthly Interest Earned",
                   "Dashboard!$D$56:$D$67", MONTHS[1:], "Dashboard!$F$56:$F$67", monthly_interest[1:],
                   ("Dashboard!$F$54", "Interest Earned"))
drawing = build_drawing()

workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    '<sheet name="Dashboard" sheetId="1" r:id="rId1"/>'
    '<sheet name="Deposit" sheetId="2" r:id="rId2"/>'
    '<sheet name="Transactions" sheetId="3" r:id="rId3"/>'
    '<sheet name="Balance" sheetId="4" r:id="rId4"/>'
    '<sheet name="Interest" sheetId="5" r:id="rId5"/>'
    '</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')

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
