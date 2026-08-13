# Retirement Withdrawal Strategies Planner

An Excel retirement planner that compares two ways to spend down a portfolio:

1. **Fixed spending** — pick a target and take it regardless of what the portfolio does.
   The target is either *4% of the starting portfolio* grown by inflation (William
   Bengen, 1994) or *your actual spending needs*, re-solved after tax every year. This
   is the strategy that can genuinely run out, which is what makes it the useful stress
   test.
2. **A dynamic strategy** — recompute spending each year from your current balance and
   remaining life expectancy, kept inside **guardrails** (a floor/ceiling withdrawal
   band) and smoothed by a **shock absorber** (a cap on year-to-year spending changes).

The two answer different questions. The dynamic strategy is portfolio-led: it optimises
for the money lasting and treats meeting your spending as a hoped-for consequence, which
is why its ending balance does not move at all when you change essential spending. Fixed
spending on the needs basis is spending-led: it asks whether the portfolio can pay for
the life you actually plan to live.

It is inspired by the Wall Street Journal article *"The 4% Rule for Retirement Is Too
Simple. Here's a Better Way."*, but it is built to be **used**, not just to illustrate
the idea: every figure is after tax, after fees, and after the costs that usually decide
whether a plan actually works — pre-Medicare healthcare, long-term care, and RMDs.
It is still an educational tool, **not financial advice**.

The workbook is generated entirely from code. Do not hand-edit the `.xlsx`; your changes
will be overwritten on the next build.

> **The generated `.xlsx` is not committed.** At ~40 MB it exceeds GitHub's warning
> threshold, and because every rebuild rewrites the whole binary it would add that much to
> history each time it changed. It is in `.gitignore`; regenerate it with the workflow
> below, or publish it as a GitHub Release asset when you want to share it.
> `Retirement Withdrawal Strategies Planner (v1 - original).xlsx` *is* committed — it is the
> small hand-built original, kept for reference.

## Repository contents

| File | Purpose |
| --- | --- |
| `model_spec.py` | **Shared source of truth for data and defaults**: historical returns/CPI 1928–2025, SSA claiming factors, IRS Uniform Lifetime (RMD) divisors, scenario definitions, and every default input value. Imported by the builder *and* the validator so the two cannot drift apart. |
| `build_retirement_planner.py` | Builds the workbook (tabs, formulas, charts, conditional formatting, data validation, comments, sheet protection) with **openpyxl**. |
| `recalc_excel.ps1` | Opens the workbook via Excel COM, forces a full recalculation, scans for error values using `SpecialCells` (fast — not a per-cell walk), and re-saves. Reports `STATUS: success - zero formula errors`. |
| `validate_model.py` | **The real correctness gate.** Recomputes the entire model independently in Python and compares ~18,500 cells against the values Excel calculated — every row of both strategy tabs, the life-expectancy engine, all 1,000 Monte Carlo outcomes, and the results panel. |
| `test_scenarios.py` | Drives the workbook through 50 alternate input sets **and 8 Monte Carlo configurations** via COM (edge-case ages, couples, stagflation, account mixes, both ceiling bases, both shortfall modes, parametric vs historical bootstrap…) and re-checks every one against the shadow model. This is what catches bugs that only appear when inputs move. |
| `test_inputs_live.py` | Perturbs all 47 inputs — each in a context where it *should* matter — and asserts something downstream changes. Catches an input that is added to the sheet and to `DEFAULTS` but never wired into a formula: the workbook would still build, still recalculate without error, and silently ignore the user. Pure Python, no Excel needed. |
| `patch_metadata.py` | Strips personal author metadata that Excel stamps on save. |
| `.gitignore` | Excludes the generated workbook, Excel `~$` lock files and `__pycache__`. |

## Prerequisites

- **Python 3.9+** with `openpyxl` (and `pywin32` for `test_scenarios.py`):
  ```bash
  pip install openpyxl pywin32
  ```
- **Windows with Microsoft Excel** — required for `recalc_excel.ps1` and
  `test_scenarios.py`. openpyxl writes formulas but does not evaluate them.

## How to (re)generate the spreadsheet

```powershell
python build_retirement_planner.py                          # 1. write formulas   (~20s)
powershell -ExecutionPolicy Bypass -File recalc_excel.ps1    # 2. calculate + scan (~90s)
python validate_model.py                                    # 3. verify the numbers
python test_inputs_live.py                                  # 4. verify every input is live
python test_scenarios.py                                    # 5. verify other inputs (~5min)
python patch_metadata.py                                    # 6. scrub metadata
```

Expected output:

```
saved: ...\Retirement Withdrawal Strategies Planner.xlsx
recalculated in 95.3s
STATUS: success - zero formula errors
cells compared: 18,532
STATUS: success - model matches Excel on every compared cell
STATUS: success - every input changes the model
STATUS: success - all 50 input scenarios and 8 Monte Carlo scenarios match the shadow model
metadata patched. creator: Retirement Planner | lastModifiedBy: Retirement Planner
```

> **Steps 3-5 are not optional in spirit.** Step 2 only proves no cell contains an
> error value. It cannot catch a formula that computes the wrong number — which is
> exactly how the previous version shipped an operator-precedence bug that inflated
> every Monte Carlo ending balance. Run step 5 before any commit, not on every edit.

> **`test_scenarios.py` drives Excel in manual-calculation mode** and recalculates only
> the sheets a change can reach. That is what took it from ~100 minutes to ~5, but the
> **sheet order is load-bearing**: `Worksheet.Calculate()` reads whatever the *other*
> sheets currently hold, so a sheet calculated too early sees stale precedents. The order
> in `recalc_deterministic()` is Monte Carlo → Historical Returns → Life Expectancy →
> Fixed Spending → Dynamic Strategy → MC Engine → MC Outcomes. Getting it wrong fails
> loudly (mismatches, never false passes) — but it does fail, so do not reorder casually.

> **The suite uses `DispatchEx`, not `Dispatch`.** Plain `Dispatch` attaches to an
> already-running Excel instance, so the suite would drive *your* open workbook and die
> the moment you closed a window. `DispatchEx` forces its own isolated process.

## Workbook structure

- **Instructions** — how to use it, what the metrics mean, and what the model still omits.
- **Inputs & Summary** — all editable inputs (yellow, grouped by topic with hover notes),
  a verdict box, the results panel, **plan diagnostics**, the scenario detail table, and
  two charts.
- **Monte Carlo** — parametric and historical-bootstrap pressure tests over 1,000 shared paths.
- **Dynamic Strategy** / **Fixed Spending** — year-by-year detail, hover notes on every column.
- **Life Expectancy** — interpolated remaining-years table plus the survival curve and
  last-survivor (joint) life expectancy.
- **Historical Returns** — S&P 500, 10-year Treasury and CPI, 1928–2025.
- **Reference Tables** — SSA claiming factors and IRS RMD divisors.
- **Article** — the narrative the model is based on.
- Hidden: **Chart Data** (trims chart series to the horizon), **MC Engine**, **MC Outcomes**.

## Key model conventions

- **The Fixed Spending target basis is a single global switch.** `fixed_basis` chooses
  between the inflation-grown percentage chain and a fresh needs solve every year, and
  because the whole column is one or the other there is no cross-contamination: on the
  percentage basis `target = prev_target × (1 + inflation)` reads a previous *percentage*
  target, never a previous solved requirement. `current_4pct` is ignored on the needs
  basis by construction.
- **On the needs basis, "essentials met" and "money lasts" collapse into one metric.**
  The withdrawal is `min(target, balance)`, so the only way to miss the need is to have
  taken the entire balance — which leaves exactly zero and can never recover. The Monte
  Carlo run confirms this: both figures come back identical. The one theoretical gap is a
  final year that drains the portfolio to *exactly* zero while still funding the need
  (`lasts` tests `balance > 0`, `met` tests income within `SHORTFALL_TOL`), which is a
  measure-zero boundary rather than a bug. A large divergence means something is wrong.
- **Horizon.** Years modelled = `plan_age - start_age`, covering ages `start_age` through
  `plan_age - 1`. The "balance at plan-to age" is the **ending** balance of the last of
  those years. Getting this off by one shifts every headline number.
- **Everything is after tax.** Essential spending is entered as an after-tax figure.
  Three account balances are tracked — taxable, traditional and Roth — and a gross
  withdrawal is sourced **taxable first, then traditional, then Roth**, the conventional
  order that lets sheltered money compound longest. Only the *gain* portion of a taxable
  sale is taxed (at the capital-gains rate), traditional dollars are taxed at the
  **age-banded** ordinary rate (early / Social-Security / RMD phase), Roth is untaxed, and
  the taxable account's dividends are taxed annually whether or not they are spent.
- **RMDs are a tax drag, not forced spending.** If the RMD exceeds what the strategy
  wanted to spend, the excess is withdrawn from the traditional account, taxed, and the
  after-tax remainder is **reinvested into the taxable account** (raising its cost basis).
  Only the tax leaves the plan.
- **Inflation is per-year, not a constant.** Both strategy tabs carry an inflation column
  and a cumulative price index. Stagflation raises inflation as well as lowering returns;
  the historical bootstrap samples the return *and* that same year's CPI together.
- **Order of operations in the dynamic rule:** life-expectancy withdrawal → guardrail band
  → shock absorber → RMD tax. The shock absorber runs last and can therefore hold spending
  *outside* the guardrail band; that is intended, and a diagnostic reports when it happens.
- **The guardrail band is an age switch, not a market response.** The actuarial withdrawal
  is `balance / remaining years`, so its implied *rate* is always `1 / remaining years` —
  the balance cancels. The ceiling therefore binds at a fixed age (whenever remaining years
  drop below `1 / ceiling`) regardless of market performance. The *shock absorber* is the
  part that reacts to markets. A flat 6% ceiling switches the strategy off from about age
  71, so the default basis is **age-graduated**: `multiple × IRS Uniform Lifetime rate`,
  clamped to the table's 72–120 range, which at the default 2.0× first binds around age 83.
  Set the basis to *Fixed %* for the original article behaviour. Ages outside the Life
  Expectancy tab's 35–110 range fall back to the flat ceiling via `IFERROR`; the shadow
  model mirrors this.
- **The headline metric is "covers essentials every year", not "money lasts".** A rule that
  withdraws a percentage of the remaining balance can almost never reach zero, so portfolio
  survival flatters it by construction — in *Follow the rule* mode it is close to
  meaningless. Only in *Withdraw enough to cover needs* mode can the dynamic strategy
  genuinely deplete, because the spending floor overrides the guardrails in bad years.
  Both are shown; essentials leads.
- **The spending floor is solved, not searched.** `required_gross()` walks the sourcing
  layers (taxable → traditional → Roth) and grosses each slice up by its own tax factor.
  It is written that way in both the builder and the shadow model deliberately: the closed-
  form algebra is easy to get sign-wrong, and a sign error there silently changes results
  rather than raising an error.
- **Formula text is a real cost on the 71,000-row engine.** `required_gross()` repeats
  its `short` and `a` sub-expressions three times each. Inlining both for the two tracks
  pushed the workbook to 68 MB; pulling them into three helper columns (`net_need` shared
  by both tracks, plus one `taxable_net` per track) brought it back to 40 MB with
  bit-identical results. When you add anything to the engine, factor repeated
  sub-expressions into columns first.
- **Coverage tests use `SHORTFALL_TOL`, not a raw `< 0`.** In *cover needs* mode income
  equals the need *exactly*, so floating-point noise leaves surpluses around 1e-11 — and
  Excel and Python round those to opposite sides of zero. Every "did this year fall short"
  test compares against `-SHORTFALL_TOL` (one dollar) in both the workbook and the shadow
  model. Changing it in `model_spec.py` changes both.
- **Monte Carlo** shares every path between strategies and uses the fixed `MC_SEED` for
  reproducibility. Change `MC_SEED` in `model_spec.py` for an independent batch.

## Known simplifications (all bias the plan to look *better* than reality)

**Spending needs drain the portfolio only in one mode.** The input *"If the Dynamic
Strategy rule falls short of spending needs"* controls this. In **Follow the rule**
(the default) each year's need — essentials, pre-65 healthcare, long-term care — is
*compared* against after-tax income but never forces a withdrawal, so a care event shows
up as extra shortfall years while leaving the ending balance unchanged. In **Withdraw
enough to cover needs** the model solves for the gross withdrawal that nets the need after
tax and takes it, so hard bills actually deplete the portfolio. Neither is wholly right:
discretionary spending really is flexible, unavoidable costs really are not. Run both and
read the spread. The Fixed Spending tab has its own version of the same choice — the
*target basis* — and setting it to your actual spending needs is the cleanest way to ask
whether the portfolio can pay for your life, because that track has no guardrails to hide
behind.

Tax is a single effective rate per phase rather than real brackets, so there is no
Social-Security tax torpedo and no bracket-filling logic; Roth conversions and
withdrawal-order optimisation cannot be expressed; no IRMAA surcharges, no
state-specific rules, no survivor-benefit changes, no haircut for a Social Security
trust-fund shortfall; historical returns are US-only and resampled IID, so no mean
reversion and no sustained bear markets. Treat a marginal result as a failing one.

## For AI agents

- `model_spec.py` and `build_retirement_planner.py` are the source of truth. Never edit the
  `.xlsx`.
- **Put shared data and defaults in `model_spec.py`.** If the builder and the validator
  disagree about a constant, the validator's guarantee is worthless.
- Input cells are assigned rows dynamically by the `InputBlock` registry, and result rows
  are recorded in the `RES` dict. Do not hard-code row numbers — reference `RES[key]`.
  `test_scenarios.py` finds inputs by their **label text**, so if you rename a label,
  update `LABEL_TO_KEY` (Inputs & Summary) or `MC_LABEL_TO_KEY` (Monte Carlo) there.
- When adding a formula, add its shadow implementation to `validate_model.py` in the same
  change. A formula with no shadow is unverified.
- **When adding an input, add it to a scenario in `test_scenarios.py` and to `CASES` in
  `test_inputs_live.py`.** Monte Carlo inputs live on a different sheet and were missed by
  the address map for exactly this reason, leaving the historical-bootstrap engine path
  unvalidated. An input that no test perturbs is an input nobody has proven is connected.
- Beware operator precedence when interpolating formula fragments: `f"={a}/{b}"` where `b`
  is `"x*(1+y)"` silently becomes `(a/x)*(1+y)`. Parenthesise interpolated denominators.
- `MIN(IF(...))`-style array formulas require Ctrl+Shift+Enter in older Excel; prefer
  `INDEX`/`MATCH`.
- The **Chart Data** sheet deliberately returns `#N/A` past the horizon so chart lines
  stop. That is not an error.
- Always finish with `patch_metadata.py` and keep the workbook free of personal data.
