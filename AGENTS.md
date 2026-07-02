# Retirement Withdrawal Strategies Planner

An educational Excel planner that compares two ways to spend down a retirement
portfolio:

1. **The classic 4% rule** — withdraw 4% of the starting portfolio in year one, then
   raise that dollar amount by inflation every year (William Bengen, 1994).
2. **A dynamic strategy** — recompute spending each year from your current balance and
   remaining life expectancy, kept inside **guardrails** (a floor/ceiling withdrawal
   band) and smoothed by a **shock absorber** (a cap on year-to-year spending changes).

It is inspired by the Wall Street Journal article *"The 4% Rule for Retirement Is Too
Simple. Here's a Better Way."* This is an educational illustration, **not financial
advice**.

The workbook is generated entirely from code. `build_retirement_planner.py` is the
single source of truth — edit it and re-run the workflow below to regenerate the
spreadsheet. Do not hand-edit the `.xlsx`; your changes will be overwritten on the
next build.

## Repository contents

| File | Purpose |
| --- | --- |
| `build_retirement_planner.py` | Builds the workbook (all tabs, formulas, charts, conditional formatting, data validation, sheet protection) with **openpyxl**. Outputs `Retirement Withdrawal Strategies Planner.xlsx` next to the script. |
| `recalc_excel.ps1` | Opens the workbook in Microsoft Excel via COM, forces a full recalculation, scans every cell for formula errors (`#REF!`, `#DIV/0!`, etc.), and re-saves. Reports `STATUS: success - zero formula errors` or lists the offending cells. |
| `patch_metadata.py` | Strips personal author metadata that Excel stamps on save, resetting the document author fields to a generic value. Leaves all worksheet content, charts, and conditional formatting intact. |
| `AGENTS.md` | This guide. |

## Prerequisites

- **Python 3.9+** with `openpyxl`:
  ```bash
  pip install openpyxl
  ```
- **Windows with Microsoft Excel** — required only for `recalc_excel.ps1`, which uses
  Excel COM automation to compute formula values and validate the file. (openpyxl
  writes formulas but does not evaluate them.)
- **PowerShell** (bundled with Windows) to run the `.ps1` script.

> No Excel? You can still build the workbook with `build_retirement_planner.py`; the
> formulas will simply be uncalculated until first opened in a spreadsheet app. If you
> have LibreOffice, you can substitute a headless `soffice --calc` recalculation for
> step 2.

## How to (re)generate the spreadsheet

Run these three steps from the repository folder, in order:

```powershell
# 1. Build the workbook (formulas written, not yet calculated)
python build_retirement_planner.py

# 2. Recalculate + validate (requires Excel on Windows)
powershell -ExecutionPolicy Bypass -File recalc_excel.ps1

# 3. Scrub personal metadata Excel stamped on save
python patch_metadata.py
```

Expected output:

```
saved: ...\Retirement Withdrawal Strategies Planner.xlsx
STATUS: success - zero formula errors
metadata patched. creator: Retirement Planner | lastModifiedBy: Retirement Planner
```

The result is `Retirement Withdrawal Strategies Planner.xlsx` in the same folder.

## Workbook structure

The generated workbook has these tabs (in order):

- **Instructions** — how to use the planner, color key, and how to unlock the sheets.
- **Inputs & Summary** — all editable inputs (yellow cells), a plain-English verdict
  box, the results panel, a market-scenario selector, life-expectancy anchors, a
  Social Security claiming-factor table, and two comparison charts.
- **Dynamic Strategy** — year-by-year dynamic withdrawals (life expectancy → guardrails
  → shock absorber), with hover notes on every column header.
- **4% Rule** — year-by-year classic 4% withdrawals.
- **Life Expectancy** — remaining-years table, fully interpolated from four anchor
  inputs (current age, 62, 67, 70) sourced from the SSA life-expectancy calculator.
- **Article** — the narrative the model is based on, with the source link.

### Key model conventions

- **Nominal vs. today's dollars** — most columns are nominal (future) dollars; columns
  and results labeled *(today's $)* are restated in today's buying power so figures are
  comparable across years.
- **Social Security** — enter the full benefit (at full retirement age 67) and a start
  age (62–70). The model applies the SSA claiming factor (70% at 62 … 124% at 70, for
  those born 1960 or later) and pays it only from the start age onward.
- **Mid-retirement use** — the dynamic strategy is meant to be re-run each year: set the
  start age to your current age and the portfolio to your current balance.
- **Color coding** — yellow = inputs, black = calculated, green = pulled from another
  tab, grey = today's-dollars restatement, red highlight = a balance/spending shortfall.
- Every build is validated to contain **zero formula errors**.

## For AI agents

- Treat `build_retirement_planner.py` as the single source of truth. Make all changes
  there, then run the three-step workflow above and confirm
  `STATUS: success - zero formula errors`.
- Result-row numbers on the **Inputs & Summary** tab are referenced by the verdict box
  and by conditional formatting. If you insert or remove result rows, update those
  references together.
- Keep the workbook free of personal information; always run `patch_metadata.py` last.
