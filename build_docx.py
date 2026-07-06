#!/usr/bin/env python3
"""Generate a .docx documenting every formula in the Digital Bank Savings
Tracker, using only the Python standard library."""
import zipfile

def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

def para(text="", style=None, runs=None, spacing_before=None):
    ppr = ""
    inner = ""
    if style or spacing_before:
        inner += "<w:pPr>"
        if style:
            inner += f'<w:pStyle w:val="{style}"/>'
        if spacing_before:
            inner += f'<w:spacing w:before="{spacing_before}"/>'
        inner += "</w:pPr>"
    if runs:
        inner += runs
    else:
        inner += f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    return f"<w:p>{inner}</w:p>"

def run(text, bold=False, color=None, mono=False, size=None, italic=False):
    rpr = "<w:rPr>"
    if bold: rpr += "<w:b/>"
    if italic: rpr += "<w:i/>"
    if mono: rpr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>'
    if color: rpr += f'<w:color w:val="{color}"/>'
    if size: rpr += f'<w:sz w:val="{size}"/>'
    rpr += "</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def cell(width, content_xml, fill=None):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
            f'<w:vAlign w:val="center"/></w:tcPr>{content_xml}</w:tc>')

def formula_table(rows):
    grid = '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="4300"/><w:gridCol w:w="3800"/></w:tblGrid>'
    borders = ('<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:left w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:right w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/></w:tblBorders>')
    tblpr = f'<w:tblPr><w:tblW w:w="9600" w:type="dxa"/>{borders}</w:tblPr>'
    # header
    hdr = ('<w:tr>'
        + cell(1500, para(runs=run("Cell(s)", bold=True, color="FFFFFF")), fill="1F3864")
        + cell(4300, para(runs=run("Formula", bold=True, color="FFFFFF")), fill="1F3864")
        + cell(3800, para(runs=run("What it does", bold=True, color="FFFFFF")), fill="1F3864")
        + '</w:tr>')
    trs = [hdr]
    for cells, formula, expl in rows:
        c1 = cell(1500, para(runs=run(cells, bold=True, size="18")))
        c2 = cell(4300, para(runs=run("=" + formula, mono=True, size="18", color="1F3864")))
        c3 = cell(3800, para(runs=run(expl, size="18")))
        trs.append(f"<w:tr>{c1}{c2}{c3}</w:tr>")
    return f"<w:tbl>{tblpr}{grid}{''.join(trs)}</w:tbl>"

# ---------------------------------------------------------------------------
# CONTENT
# ---------------------------------------------------------------------------
body = []
body.append(para(runs=run("Digital Bank Savings Tracker", bold=True, color="1F3864", size="40"), style="Title"))
body.append(para(runs=run("Formula Reference Guide", bold=True, color="2E5496", size="30")))
body.append(para(runs=run("This document lists every calculation used in the workbook, grouped by sheet. "
                 "Formulas repeat down each column, so one representative example is shown per pattern along "
                 "with the range of cells it covers. All amounts use the Philippine Peso currency format.",
                 italic=True, color="808080")))

body.append(para(runs=run("How the sheets connect", bold=True, color="1F3864", size="26"), style="Heading1", spacing_before="240"))
body.append(para("Deposit + Transactions  ->  Balance  ->  Interest  ->  Dashboard. "
                 "You type raw entries on the Deposit and Transactions tabs; the Balance sheet totals them per account; "
                 "the Interest sheet computes compounding earnings on those balances; and the Dashboard summarises everything with charts."))

# ---- Dashboard ----
body.append(para(runs=run("1.  Dashboard", bold=True, color="1F3864", size="28"), style="Heading1", spacing_before="240"))
body.append(para(runs=run("Key Metrics", bold=True, color="2E5496", size="24")))
body.append(formula_table([
    ("E5", "SUM(Balance!$H$6:$H$19)", "Total Balance Today (incl. interest): adds every account's Balance Today. Grows on its own as time passes."),
    ("E6", "SUM(Balance!$D$6:$D$19)", "Total Deposited (All-Time): sum of every account's total deposits."),
    ("E7", "SUM(Balance!$G$6:$G$19)",
        "Interest Earned to Date (auto): total interest accrued so far across all accounts; increases each day/month."),
    ("E8", '-SUMIFS(Transactions!$E$6:$E$205,Transactions!$B$6:$B$205,"Withdrawal")',
        'Total Withdrawn: sums all "Withdrawal" amounts; the leading minus turns the stored negatives into a positive total.'),
]))
body.append(para(runs=run("Personal Goal Tracker", bold=True, color="2E5496", size="24"), spacing_before="160"))
body.append(formula_table([
    ("E11", "(you type your goal, e.g. 100000)", "Your Savings Goal Target: a value you enter (highlighted yellow), not a formula."),
    ("E12", 'IFERROR(VLOOKUP("Maya - Personal Goal",Balance!$C$6:$H$19,6,FALSE),0)',
        "Maya Personal Goal Balance Today: looks up that account and returns its Balance Today (principal + accrued interest)."),
    ("E13", "IFERROR(E12/E11,0)", "Progress Toward Goal: current goal balance divided by target (shown as a %)."),
]))
body.append(para(runs=run("Balance by Bank (also feeds the pie chart)", bold=True, color="2E5496", size="24"), spacing_before="160"))
body.append(formula_table([
    ("B17:B23", "SUMIF(Balance!$A$6:$A$19,A17,Balance!$H$6:$H$19)",
        "Totals the Balance Today of every account belonging to the bank named in column A (e.g. all Maya accounts)."),
    ("B24", "SUM(B17:B23)", "Grand total of all bank balances."),
]))
body.append(para(runs=run("Chart data - Savings Growth (cumulative balance, rows 80-103)", bold=True, color="2E5496", size="24"), spacing_before="160"))
body.append(formula_table([
    ("B80:H103", "SUMPRODUCT((Interest!$A$6:$A$19=B$79)*Interest!$G$6:$G$19*Interest!$M$6:$M$19^0)",
        "Projected cumulative balance for one bank in a given month. For each account it multiplies the current "
        "balance by the bank's Monthly Growth Factor raised to the month number, and only counts accounts of that "
        "bank (the =B$79 test). The exponent ^0, ^1 ... ^23 is the month index (Jan 2026 = 0)."),
    ("I81:I103", "SUM(B81:H81)-SUM(B80:H80)",
        "Total interest earned that month = this month's total across all banks minus last month's total (I80 = 0)."),
]))
body.append(para(runs=run("Chart data - Monthly Interest by Bank (rows 107-130)", bold=True, color="2E5496", size="24"), spacing_before="160"))
body.append(formula_table([
    ("B108:H130", "B81-B80",
        "Each bank's interest for the month = its cumulative balance this month minus last month "
        "(references the Savings Growth block above). Row 107 = 0. This drives the stacked column chart."),
]))

# ---- Deposit / Transactions ----
body.append(para(runs=run("2.  Deposit  &  3.  Transactions", bold=True, color="1F3864", size="28"), style="Heading1", spacing_before="240"))
body.append(para("These two tabs contain no formulas - they are data-entry tables (Excel Tables with filter "
                 "dropdowns). You type each deposit/transaction on a new row and the Balance and Dashboard sheets read "
                 "from them automatically. Columns: Deposit = Date, Bank, Account, Amount, Notes; "
                 "Transactions = Date, Type, Bank, Account, Amount (use negatives for money out), Notes."))

# ---- Balance ----
body.append(para(runs=run("4.  Balance", bold=True, color="1F3864", size="28"), style="Heading1", spacing_before="240"))
body.append(para("One row per account (rows 6-19). Each column below repeats down all 14 accounts; the example shows row 6."))
body.append(formula_table([
    ("D6:D19", "SUMIFS(Deposit!$D$6:$D$205,Deposit!$C$6:$C$205,C6)",
        "Deposits: total of all Deposit-tab amounts recorded for this account (matched by Account Label)."),
    ("E6:E19", "SUMIFS(Transactions!$E$6:$E$205,Transactions!$D$6:$D$205,C6)",
        "Net Transactions: sum of all Transaction amounts for this account (withdrawals are negative, so they subtract)."),
    ("F6:F19", "D6+E6", "Current Balance: Deposits plus Net Transactions (money you moved)."),
    ("G6:G19", "IFERROR(VLOOKUP(C6,Interest!$D$6:$Q$19,13,FALSE),0)",
        "Interest to Date: the auto-accrued interest for this account (from the Interest sheet); grows with the date."),
    ("H6:H19", "IFERROR(VLOOKUP(C6,Interest!$D$6:$Q$19,14,FALSE),0)",
        "Balance Today: current balance plus interest accrued to date. This is the number that increases on its own."),
    ("I6:I19", "IFERROR(VLOOKUP(C6,Interest!$D$6:$L$19,9,FALSE),0)",
        "Interest (1 Yr): projected 1-year interest for this account."),
    ("J6:J19", "F6+I6", "Projected Balance (1 Yr): current balance plus projected 1-year interest."),
    ("D20:J20", "SUM(D6:D19)", "Column totals row (same SUM pattern for E, F, G, H, I and J)."),
]))

# ---- Interest ----
body.append(para(runs=run("5.  Interest", bold=True, color="1F3864", size="28"), style="Heading1", spacing_before="240"))
body.append(para("One row per account (rows 6-19). Daily accounts (Maya/UNO/Banko/CIMB/Maribank savings) compound "
                 "daily; the rest compound monthly. Maya Savings rate is set to 10%. Example shows row 6. "
                 "Columns N-Q make the balance grow automatically over time using the TODAY() function."))
body.append(formula_table([
    ("G6:G19", "IFERROR(VLOOKUP(D6,Balance!$C$6:$H$19,4,FALSE),0)",
        "Current Balance: pulls this account's current balance from the Balance sheet."),
    ("H6:H19", 'IF(F6="Daily",E6/365,E6/12)',
        "Periodic Rate: annual rate / 365 for daily accounts, or / 12 for monthly accounts."),
    ("I6:I19", "G6*H6",
        "Interest / Period: balance x periodic rate = the daily earning (daily accounts) or monthly earning (monthly accounts)."),
    ("J6:J19", 'IF(F6="Daily",365,12)', "Periods / Yr: 365 for daily accounts, 12 for monthly."),
    ("K6:K19", "G6*(1+H6)^J6",
        "Projected Balance (1 Yr): compound-interest formula balance x (1 + periodic rate) ^ periods. This is what makes interest grow on itself."),
    ("L6:L19", "K6-G6", "Interest (1 Yr): projected balance minus starting balance = total interest earned in a year."),
    ("M6:M19", 'IF(F6="Daily",(1+E6/365)^(365/12),1+E6/12)',
        "Monthly Growth Factor: how much 1 peso grows in one month (used by the Dashboard growth/interest charts)."),
    ("N6:N19", "(editable date - e.g. 6/1/2026)",
        "Interest Start: a plain, editable date (highlighted yellow) from which interest accrues. Pre-filled with "
        "your first deposit date; type a different date to change it. Kept as a value (not a lookup) so it works on every Excel version."),
    ("O6:O19", 'IF(ISNUMBER(N6),IF(F6="Daily",MAX(0,TODAY()-N6),MAX(0,DATEDIF(N6,TODAY(),"m"))),0)',
        "Periods Elapsed: days since the start date for daily accounts, or whole months for monthly accounts, up to TODAY(). "
        "Because it uses TODAY(), it grows by itself each time you open the file."),
    ("P6:P19", "G6*((1+H6)^O6-1)",
        "Interest to Date: compound interest actually earned from the start date until today."),
    ("Q6:Q19", "G6+P6",
        "Balance Today: current balance plus interest to date - the amount that increases on its own each day/month."),
    ("G20,I20,K20,L20,P20,Q20", "SUM(G6:G19)", "Column totals row (same SUM pattern for I, K, L, P and Q)."),
]))

body.append(para(runs=run("How the automatic growth works", bold=True, color="1F3864", size="26"), style="Heading1", spacing_before="240"))
body.append(para("The Interest sheet uses TODAY(), which Excel refreshes every time the file is opened (or recalculated). "
                 "So each new day, daily-interest accounts grow a little, and each new month, monthly-interest accounts step "
                 "up - the Balance Today, the Dashboard totals and the Savings-by-Bank pie all rise on their own without you "
                 "editing anything. Note: this is an estimate of what your money earns; because interest is now accrued "
                 "automatically, do NOT also record interest as a manual transaction (that would double-count). Actual bank "
                 "postings and any taxes may differ slightly."))
body.append(para(runs=run("These figures currently use example seed data. Replace the entries on the Deposit and "
                          "Transactions tabs with your real amounts and every formula above recalculates automatically.",
                          italic=True, color="808080"), spacing_before="120"))

sectpr = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
          '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
          'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W}><w:body>{"".join(body)}{sectpr}</w:body></w:document>')

styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles {W}>'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
    '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
    '<w:pPr><w:spacing w:after="80"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr></w:style>'
    '</w:styles>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '</Types>')

root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')

doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>')

OUT = "/projects/sandbox/Kiro/Digital-Bank-Savings-Tracker-Formulas.docx"
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/_rels/document.xml.rels", doc_rels)
print("Wrote", OUT)
