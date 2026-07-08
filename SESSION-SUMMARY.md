# Digital Bank Savings Tracker — Session Summary

A running log of everything built and changed in this session. The tracker is a
pure-Excel "application" (no macros) generated from Python scripts, plus a Word
document that documents every formula.

---

## Files in the repository

| File | What it is |
|------|------------|
| `Digital-Bank-Savings-Tracker.xlsx` | The main workbook (5 tabs: Dashboard, Deposit, Transactions, Balance, Interest). |
| `Digital-Bank-Savings-Tracker-Formulas.docx` | Reference document listing every formula in the workbook, by sheet. |
| `build_app.py` | Generator that builds the `.xlsx` from raw Office XML (stdlib only — no openpyxl). |
| `build_docx.py` | Generator that builds the formula-reference `.docx`. |
| `Bank-Savings-Tracker-Interest.xlsx` | The first, interest-only prototype workbook. |
| `build_tracker.py` | Generator for that first prototype. |
| `.gitattributes` | Marks Office files as binary so Git never alters their bytes. |

> The workbooks were hand-built as Office Open XML because `openpyxl` was not
> available offline. Every build is validated (valid ZIP, well-formed XML,
> matching style/relationship references, no duplicate cells).

---

## The workbook at a glance

**Sheets:** Dashboard · Deposit · Transactions · Balance · Interest

- **App-style look:** navy header band, clickable nav buttons (hyperlinks),
  hidden gridlines, frozen panes, consistent theme, Philippine Peso currency
  formatting throughout.
- **Data flow:** `Deposit + Transactions -> Balance -> Interest -> Dashboard`
  (no circular references).
- **Accounts (14):** Maya (Savings 10%, 2 Time Deposits, Personal Goal),
  UNO (Digi Savings, 2 Time Deposits), Tonik (2 Stash, Time Deposit),
  Banko, CIMB, GoTyme, Maribank.

---

## What we did — in order

### 1. First interest workbook
- Built an initial `.xlsx` with an Interest sheet: daily/monthly compounding,
  Maya Savings set to **10%**, auto-calculated earnings. Pushed to `JKP143/Kiro`.

### 2. Full 5-sheet application
- Created the full workbook: **Dashboard, Deposit, Transactions, Balance,
  Interest**, app-style header + nav, filterable tables, and 3 Dashboard charts
  (Savings Growth, Savings by Bank, Monthly Interest).

### 3. Dashboard redesign
- Added **Key Metrics** (Total Balance, Total Deposited, Interest, Withdrawn),
  a **Personal Goal Tracker** (editable target + progress %), a **Balance by
  Bank** table, a **multi-line Savings Growth** chart (one line per bank,
  24 months), and a **Monthly Interest** chart.

### 4. Number-format fixes
- Fixed `#######` in the chart-data area (columns too narrow) by widening and
  using compact formats.
- Made **all amount cells** use a proper currency format.
- Switched from a custom `"PHP "` text format to **`[$₱-3409]#,##0.00`** so
  Excel recognizes it as the **Currency** category (not "Custom").

### 5. Chart correctness + upgrades
- **Fixed the empty "Savings by Bank" pie:** the Balance-by-Bank `SUMIF` was
  matching the Account column instead of the **Bank** column, returning ₱0.00.
- Changed **Monthly Interest** to a **stacked column** chart with one segment
  per bank.
- Polished both time-series charts: distinct 7-color bank palette, peso-
  formatted axes, subtle gridlines, rotated month labels, bottom legend.

### 6. Formula reference document
- Generated `Digital-Bank-Savings-Tracker-Formulas.docx` documenting every
  formula grouped by sheet, in plain English.

### 7. Automatic, self-growing interest (the big one)
- Added **automatic interest accrual** driven by `TODAY()`: each account has an
  **Interest Start** date and the balance grows on its own every time the file
  is opened/recalculated.
- **Compatibility fixes:** `MINIFS` (Excel 2019+) caused `#NAME?` on older Excel
  and showed ₱0.00 — replaced with `AGGREGATE`, then finally made **Interest
  Start a plain editable date** so the accrual uses only universal functions
  (`TODAY`, `DATEDIF`, `IF`, `MAX`, `ISNUMBER`).
- Set the workbook to **automatic calculation + full recalc on open** so
  `TODAY()` refreshes reliably.

### 8. Interest posted as transactions + folded into the balance
- The **Balance** sheet's headline **Current Balance now includes accrued
  interest** (Principal + interest).
- Added an **auto-posted Interest ledger** on the Transactions tab (read-only,
  one line per account). Balances use structured references
  (`SUMIFS(tblTransactions[Amount], ...)`) so the ledger is never double-counted
  and there is no circular reference.
- Reworked the Transactions layout so filtering/sorting the log can't hide the
  interest, ending with the **interest ledger on top** and the **Transaction
  Log below** (the log grows freely with nothing beneath it).

### 9. User-friendly upgrades
- **Dropdown menus** (Bank / Account / Type) on Deposit & Transactions to
  prevent typos that would break the balance matching.
- **Goal progress data bar** on the Dashboard.
- Live **"Figures as of [date]"** stamp on the Dashboard.
- Hardened the file: proper string-formula cell, `.gitattributes` for binaries.

### 10. Layout / usability fixes
- **Interest tab nav buttons unclickable when scrolled:** hid the internal calc
  columns so the sheet fits on screen (no horizontal scroll).
- Diagnosed the **"file format is not valid"** error as a **download problem**
  (GitHub `raw` rate-limit / HTML error page saved as `.xlsx`) — the file itself
  was always valid. Recommended downloading via the repo ZIP.

### 11. Interest presented per DAY
- Removed the misleading **Interest / Period** total (it summed per-day and
  per-month figures).
- Renamed the column to **Interest / Day** (`Balance × Rate ÷ 365`) for every
  account and added an **Interest / Month** column beside it (both labeled est.).
- Changed the **accrual itself to DAILY for every account** (was monthly for the
  monthly accounts) so interest grows every day — matching how banks compute
  interest on the daily balance.
- The **Transactions ledger** and the **Balance sheet** now both show the
  per-day interest amount; the cumulative "Interest to Date" that makes up the
  balance stays on the Interest tab and inside **Current Balance**.

---

## Key formulas (current)

| Where | Formula | Meaning |
|-------|---------|---------|
| Interest — Interest / Day | `G*E/365` | Balance × annual rate ÷ 365 |
| Interest — Interest / Month | `G*E/12` | Balance × annual rate ÷ 12 |
| Interest — Days Elapsed | `IF(ISNUMBER(N),MAX(0,TODAY()-N),0)` | days since Interest Start |
| Interest — Interest to Date | `G*((1+E/365)^O-1)` | daily-compounded interest so far |
| Interest — Balance Today | `G+P` | balance incl. accrued interest |
| Balance — Current Balance | `VLOOKUP(...Interest Balance Today)` | Principal + interest to date |
| Balance — Net Transactions | `SUMIFS(tblTransactions[Amount],tblTransactions[Account],C)` | your logged movements only |

---

## How to use it

1. **Download** via the repo ZIP (avoids GitHub raw rate limits):
   `https://github.com/JKP143/Kiro/archive/refs/heads/main.zip`
2. Open the workbook and click **Enable Editing** if you see the yellow bar.
3. Press **F9** to refresh `TODAY()` — the Dashboard "Figures as of" stamp should
   show today's date, and daily interest updates.
4. Enter deposits on the **Deposit** tab and other movements on the
   **Transactions** log (use the dropdowns). Do **not** log interest manually —
   it is auto-posted.
5. Set each account's **Interest Start** date (yellow) on the Interest tab if it
   differs from the pre-filled first-deposit date.

> Figures currently use example seed data. Replace them with your real amounts
> and everything recalculates automatically.

---

*This summary reflects the state of the workbook at the end of the session.*
