#!/usr/bin/env python3
"""Generate an Interest workbook (.xlsx) implementing daily/monthly compounding
interest for the digital bank savings tracker, using only the stdlib."""
import zipfile

# ---- Account data (from the Interest tab screenshot) -------------------------
# Maya Savings rate overridden to 10% per user instruction.
# Daily-compounding: Maya/UNO/Banko/CIMB/Maribank *Savings* accounts.
# Everything else compounds monthly.
# (Bank, Account, Type, Label, AnnualRate, Frequency, ExampleBalance)
ACCOUNTS = [
    ("Maya", "Savings",        "Savings",      "Maya - Savings",        0.10,   "Daily",   50000),
    ("Maya", "Time Deposit 1", "Time Deposit", "Maya - Time Deposit 1", 0.06,   "Monthly", 100000),
    ("Maya", "Time Deposit 2", "Time Deposit", "Maya - Time Deposit 2", 0.06,   "Monthly", 100000),
    ("Maya", "Personal Goal",  "Goal",         "Maya - Personal Goal",  0.056,  "Monthly", 25000),
    ("UNO",  "Digi Savings",   "Savings",      "UNO - Digi Savings",    0.035,  "Daily",   20000),
    ("UNO",  "Time Deposit 1", "Time Deposit", "UNO - Time Deposit 1",  0.06,   "Monthly", 50000),
    ("UNO",  "Time Deposit 2", "Time Deposit", "UNO - Time Deposit 2",  0.06,   "Monthly", 50000),
    ("Tonik","Stash 1",        "Savings",      "Tonik - Stash 1",       0.04,   "Monthly", 10000),
    ("Tonik","Stash 2",        "Savings",      "Tonik - Stash 2",       0.04,   "Monthly", 10000),
    ("Tonik","Time Deposit",   "Time Deposit", "Tonik - Time Deposit",  0.08,   "Monthly", 30000),
    ("Banko","Savings",        "Savings",      "Banko - Savings",       0.05,   "Daily",   15000),
    ("CIMB", "Savings",        "Savings",      "CIMB - Savings",        0.023,  "Daily",   10000),
    ("GoTyme","GoSave",        "Savings",      "GoTyme - GoSave",       0.03,   "Monthly", 12000),
    ("Maribank","Savings",     "Savings",      "Maribank - Savings",    0.0325, "Daily",   8000),
]

HEADERS = ["Bank", "Account", "Account Type", "Account Label", "Annual Rate",
           "Frequency", "Current Balance", "Periodic Rate", "Interest / Period",
           "Periods / Yr", "Proj. Balance (1 Yr)", "Interest Earned (1 Yr)"]

def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

# Style indices used below:
# 1 title | 2 header | 3 percent | 4 currency | 5 integer | 6 text
# 7 input-currency(highlight) | 8 totals-currency | 9 totals-text | 10 note
def cell_text(ref, val, style):
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{esc(val)}</t></is></c>'

def cell_num(ref, val, style):
    return f'<c r="{ref}" s="{style}"><v>{val}</v></c>'

def cell_formula(ref, formula, style):
    return f'<c r="{ref}" s="{style}"><f>{esc(formula)}</f></c>'

rows_xml = []

# Row 1: title (merged A1:L1)
rows_xml.append('<row r="1" ht="24" customHeight="1">'
                + cell_text("A1", "Interest Rates & Earnings by Account", 1)
                + '</row>')
# Row 2: note (merged A2:L2)
note = ("Rates from your bank notes (Maya Savings overridden to 10%). "
        "Edit only the yellow Current Balance column - all other columns calculate "
        "automatically. Daily accounts compound daily (rate/365); Monthly accounts compound monthly (rate/12).")
rows_xml.append('<row r="2" ht="30" customHeight="1">' + cell_text("A2", note, 10) + '</row>')

# Row 3: headers
cols = "ABCDEFGHIJKL"
hdr_cells = "".join(cell_text(f"{cols[i]}3", HEADERS[i], 2) for i in range(len(HEADERS)))
rows_xml.append('<row r="3" ht="30" customHeight="1">' + hdr_cells + '</row>')

# Data rows 4..
start = 4
for i, (bank, acct, atype, label, rate, freq, bal) in enumerate(ACCOUNTS):
    r = start + i
    c = []
    c.append(cell_text(f"A{r}", bank, 6))
    c.append(cell_text(f"B{r}", acct, 6))
    c.append(cell_text(f"C{r}", atype, 6))
    c.append(cell_text(f"D{r}", label, 6))
    c.append(cell_num(f"E{r}", rate, 3))                       # annual rate (percent)
    c.append(cell_text(f"F{r}", freq, 6))                      # Daily / Monthly
    c.append(cell_num(f"G{r}", bal, 7))                        # editable balance
    c.append(cell_formula(f"H{r}", f'IF(F{r}="Daily",E{r}/365,E{r}/12)', 3))   # periodic rate
    c.append(cell_formula(f"I{r}", f'G{r}*H{r}', 4))          # interest per period (auto)
    c.append(cell_formula(f"J{r}", f'IF(F{r}="Daily",365,12)', 5))             # periods/yr
    c.append(cell_formula(f"K{r}", f'G{r}*(1+H{r})^J{r}', 4)) # compounded balance after 1 yr
    c.append(cell_formula(f"L{r}", f'K{r}-G{r}', 4))          # total interest earned in 1 yr
    rows_xml.append(f'<row r="{r}">' + "".join(c) + '</row>')

# Totals row
last = start + len(ACCOUNTS) - 1
tr = last + 1
tot = []
tot.append(cell_text(f"A{tr}", "TOTAL", 9))
for col in "BCDEF":
    tot.append(cell_text(f"{col}{tr}", "", 9))
tot.append(cell_formula(f"G{tr}", f'SUM(G{start}:G{last})', 8))
tot.append(cell_text(f"H{tr}", "", 9))
tot.append(cell_formula(f"I{tr}", f'SUM(I{start}:I{last})', 8))
tot.append(cell_text(f"J{tr}", "", 9))
tot.append(cell_formula(f"K{tr}", f'SUM(K{start}:K{last})', 8))
tot.append(cell_formula(f"L{tr}", f'SUM(L{start}:L{last})', 8))
rows_xml.append(f'<row r="{tr}">' + "".join(tot) + '</row>')

sheet_data = "".join(rows_xml)

col_widths = {"A":10,"B":16,"C":14,"D":22,"E":12,"F":11,"G":17,"H":12,"I":18,"J":11,"K":21,"L":20}
cols_xml = "".join(
    f'<col min="{i+1}" max="{i+1}" width="{col_widths[cols[i]]}" customWidth="1"/>'
    for i in range(12))

merges = f'<mergeCell ref="A1:L1"/><mergeCell ref="A2:L2"/>'

sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>{cols_xml}</cols>
<sheetData>{sheet_data}</sheetData>
<mergeCells count="2">{merges}</mergeCells>
</worksheet>'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;PHP &quot;#,##0.00"/></numFmts>
<fonts count="4">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Interest" sheetId="1" r:id="rId1"/></sheets>
<calcPr calcId="0" fullCalcOnLoad="1"/>
</workbook>'''

workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

OUT = "/projects/sandbox/Bank-Savings-Tracker-Interest.xlsx"
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("xl/workbook.xml", workbook_xml)
    z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
    z.writestr("xl/styles.xml", styles_xml)
    z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

print("Wrote", OUT)
