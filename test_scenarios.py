"""Drive the workbook through alternate inputs and check Excel still agrees.

`validate_model.py` proves the workbook is right at its default inputs. This proves
it stays right when the inputs move — which is where the original build broke. Each
scenario below targets a specific defect that was fixed:

  le-anchor-65   the life-expectancy anchors used to be fed to MATCH() unsorted
                 whenever the "current age" anchor was above 62, silently producing
                 remaining-life values like 63 years at age 62.
  horizon-*      the balance "at plan-to age" used to be read one row too far,
                 reporting the balance a year after the plan ended.
  couple         last-survivor life expectancy, which did not exist before.
  stagflation    inflation is now per-year and scenario-driven, not a constant.
  nominal-shock  the shock absorber band basis.
  short-horizon  the smallest legal horizon, where off-by-one errors surface.

Requires Excel (COM). Run after build_retirement_planner.py.

    python test_scenarios.py
"""
import os
import sys

import win32com.client as w32

import model_spec as spec
import validate_model as V

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "Retirement Withdrawal Strategies Planner.xlsx")

# label on Inputs & Summary column B -> key in model_spec.DEFAULTS
LABEL_TO_KEY = {
    "Retirement start age (or your current age, if already retired)": "start_age",
    "Plan-to age (planning horizon)": "plan_age",
    "Household": "household",
    "Spouse is younger by (years; negative = older)": "spouse_gap",
    "Taxable / brokerage balance": "taxable0",
    "Cost basis in the taxable account (% of its value)": "basis_share",
    "Traditional 401(k) / IRA balance (tax-deferred)": "traditional0",
    "Roth balance (tax-free)": "roth0",
    "Expected annual return (nominal, BEFORE fees)": "exp_ret",
    "Annual fees (fund expenses + any advisory fee)": "fee",
    "Inflation rate (long-run average)": "infl",
    "Market scenario (drives the return AND inflation columns)": "scenario",
    "Essential annual spending (today's $, after tax)": "essential",
    "Extra healthcare premiums before age 65 (today's $/yr)": "healthcare",
    "Real spending drift per year (spending smile)": "drift",
    "4% rule starting rate": "rate4",
    "Guardrail floor (minimum withdrawal rate)": "floor",
    "Guardrail ceiling (maximum withdrawal rate)": "ceiling",
    "Guardrail ceiling basis": "ceiling_basis",
    "Age-graduated ceiling — multiple of the IRS rate": "ceiling_mult",
    "Shock absorber — max spending change per year": "shock",
    "Shock absorber applies to": "shock_basis",
    "Current 4% rule withdrawal if already retired (today's $/yr; 0 if not)": "current_4pct",
    "Effective tax rate — before Social Security starts": "tax_early",
    "Effective tax rate — Social Security until RMDs": "tax_ss",
    "Effective tax rate — once RMDs begin": "tax_rmd",
    "Long-term capital gains effective rate": "cg_rate",
    "Taxable account dividend yield (taxed every year)": "div_yield",
    "RMD start age": "rmd_age",
    "Full Social Security benefit at FRA 67 (today's $/yr, before tax)": "full_ss",
    "Social Security start age (62-70)": "ss_age",
    "Include a long-term care event?": "ltc_on",
    "Long-term care cost (today's $/yr)": "ltc_cost",
    "Long-term care duration (years)": "ltc_years",
    "Long-term care starting age": "ltc_age",
    "Current age (for life expectancy)": "cur_age_le",
    "Remaining life expectancy at that current age": "le_cur",
    "Remaining life expectancy at age 62": "le62",
    "Remaining life expectancy at age 67": "le67",
    "Remaining life expectancy at age 70": "le70",
}

SCENARIOS = [
    ("defaults", {}),
    ("le-anchor-65", {"cur_age_le": 65, "le_cur": 20.2, "start_age": 65, "plan_age": 92}),
    ("le-anchor-70", {"cur_age_le": 70, "le_cur": 16.7, "start_age": 70, "plan_age": 95}),
    ("couple", {"household": "Couple", "spouse_gap": 3}),
    ("couple-older-spouse", {"household": "Couple", "spouse_gap": -6}),
    ("horizon-65-95", {"start_age": 65, "plan_age": 95, "cur_age_le": 65, "le_cur": 20.2}),
    ("short-horizon", {"start_age": 60, "plan_age": 61}),
    ("long-horizon", {"start_age": 40, "plan_age": 105, "cur_age_le": 40, "le_cur": 41.0}),
    ("stagflation", {"scenario": "Stagflation"}),
    ("early-crash", {"scenario": "Early crash"}),
    ("nominal-shock", {"shock_basis": "Nominal"}),
    ("nominal-shock-crash", {"shock_basis": "Nominal", "scenario": "Early crash"}),
    ("real-shock-crash", {"shock_basis": "Real", "scenario": "Early crash"}),
    ("nominal-shock-stagflation", {"shock_basis": "Nominal", "scenario": "Stagflation"}),
    ("spending-smile", {"drift": -0.01}),
    ("couple-at-65", {"household": "Couple", "start_age": 65, "plan_age": 95,
                      "cur_age_le": 65, "le_cur": 20.2}),
    ("no-tax-no-ltc", {"tax_early": 0.0, "tax_ss": 0.0, "tax_rmd": 0.0, "cg_rate": 0.0,
                       "div_yield": 0.0, "ltc_on": "No"}),
    ("all-roth", {"taxable0": 0, "traditional0": 0, "roth0": 3_500_000}),
    ("all-traditional", {"taxable0": 0, "traditional0": 3_500_000, "roth0": 0}),
    ("all-taxable", {"taxable0": 3_500_000, "traditional0": 0, "roth0": 0}),
    ("taxable-zero-basis", {"basis_share": 0.0}),
    ("taxable-full-basis", {"basis_share": 1.0}),
    ("flat-tax-bands", {"tax_early": 0.18, "tax_ss": 0.18, "tax_rmd": 0.18}),
    ("steep-tax-bands", {"tax_early": 0.05, "tax_ss": 0.20, "tax_rmd": 0.32}),
    ("high-tax-early-ss", {"tax_early": 0.35, "tax_ss": 0.35, "tax_rmd": 0.35,
                           "ss_age": 62, "rmd_age": 73}),
    ("rmd-heavy", {"taxable0": 100_000, "traditional0": 3_300_000, "roth0": 100_000,
                   "rmd_age": 73}),
    ("already-retired-4pct", {"start_age": 70, "plan_age": 95, "current_4pct": 90_000,
                              "cur_age_le": 70, "le_cur": 16.7}),
    ("tight-ceiling", {"ceiling": 0.045, "floor": 0.025, "ceiling_basis": "Fixed %"}),
    ("ceiling-fixed-6", {"ceiling_basis": "Fixed %", "ceiling": 0.06}),
    ("ceiling-grad-1.5x", {"ceiling_basis": "Age-graduated", "ceiling_mult": 1.5}),
    ("ceiling-grad-2.5x", {"ceiling_basis": "Age-graduated", "ceiling_mult": 2.5}),
    ("ceiling-grad-old-start", {"ceiling_basis": "Age-graduated", "start_age": 80,
                                "plan_age": 100, "cur_age_le": 80, "le_cur": 10.4}),
    ("ceiling-grad-couple", {"ceiling_basis": "Age-graduated", "household": "Couple"}),
]

# Monte Carlo inputs live on the Monte Carlo sheet, not Inputs & Summary, which is
# why they were previously untested: the address map above never reached them. The
# historical-bootstrap branch of the MC engine had no validation at all as a result.
MC_LABEL_TO_KEY = {
    "Return method": "mc_method",
    f"Simulations included (100-{spec.MAX_SIMS})": "mc_sims",
    "Annual return volatility (parametric)": "mc_vol",
    "Annual inflation volatility (parametric)": "mc_infl_vol",
    "US stock allocation (historical method)": "mc_stock",
    "Chance of the long-term care event occurring": "mc_ltc_prob",
}

# Each of these forces a full recalculation of the 71,000-row engine, so keep the
# list tight and targeted rather than exhaustive.
MC_SCENARIOS = [
    ("mc-parametric (default)", {}),
    ("mc-historical-bootstrap", {"mc_method": "Historical bootstrap"}),
    ("mc-historical-95pct-stock", {"mc_method": "Historical bootstrap", "mc_stock": 0.95}),
    ("mc-high-volatility", {"mc_vol": 0.25, "mc_infl_vol": 0.06}),
    ("mc-certain-ltc", {"mc_ltc_prob": 1.0}),
]

# MC Outcomes column -> shadow-model key
MC_OUT_COLS = [
    (2, "four_lasts"), (3, "four_ess"), (4, "four_bal"), (5, "four_low"),
    (6, "four_sum"), (7, "dyn_lasts"), (8, "dyn_ess"), (9, "dyn_bal"),
    (10, "dyn_low"), (11, "dyn_sum"), (12, "dyn_worst"),
    (13, "four_share"), (14, "dyn_share"),
]


def main():
    if not os.path.exists(BOOK):
        print("MISSING WORKBOOK - run build_retirement_planner.py first")
        return 1
    # DispatchEx forces a dedicated Excel process. Plain Dispatch() attaches to an
    # already-running instance, which means the suite would drive whatever workbook
    # the user happens to have open — and closing that window kills the run mid-way
    # with "The object invoked has disconnected from its clients".
    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.UserName = "Retirement Planner"
    wb = xl.Workbooks.Open(BOOK)
    try:
        sh = wb.Worksheets("Inputs & Summary")
        # locate every input cell by its label text
        addr = {}
        for r in range(1, 90):
            lab = sh.Cells(r, 2).Value
            if isinstance(lab, str):
                key = LABEL_TO_KEY.get(lab.strip())
                if key:
                    addr[key] = r
        missing = set(LABEL_TO_KEY.values()) - set(addr)
        if missing:
            print(f"FAILED to locate input rows for: {sorted(missing)}")
            return 1

        fr = wb.Worksheets("4% Rule")
        dysh = wb.Worksheets("Dynamic Strategy")
        lesh = wb.Worksheets("Life Expectancy")
        last_row = V.R0 + N_ROWS - 1
        le_last = 6 + len(V.AGES) - 1

        def near(a, b):
            if a is None or b is None:
                return False
            return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))

        total_fail = 0
        for name, overrides in SCENARIOS:
            for key, row in addr.items():
                sh.Cells(row, 3).Value = spec.DEFAULTS[key]
            for key, val in overrides.items():
                sh.Cells(addr[key], 3).Value = val
            xl.CalculateFullRebuild()

            V.configure(overrides)
            shared = V.four_rule()
            dyn = V.dynamic(shared)
            fails = []

            # One COM round trip per block instead of one per cell.
            le_vals = lesh.Range(f"A6:M{le_last}").Value
            four_vals = fr.Range(f"C{V.R0}:AF{last_row}").Value
            dyn_vals = dysh.Range(f"C{V.R0}:AK{last_row}").Value
            dyn_age = dysh.Range(f"B{V.R0}:B{last_row}").Value

            COLS4 = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P",
                     "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC",
                     "AD", "AE", "AF"]
            COLSD = COLS4 + ["AG", "AH", "AI", "AJ", "AK"]

            prev = None
            for i, age in enumerate(V.AGES):
                got = le_vals[i][1]
                if not near(got, V.SINGLE[age]):
                    fails.append(f"LE!B{6+i} age {age}: shadow {V.SINGLE[age]} vs excel {got}")
                if got is None or got <= 0 or got > 60:
                    fails.append(f"LE!B{6+i} age {age}: implausible remaining years {got}")
                elif prev is not None and got > prev + 1e-9:
                    fails.append(f"LE!B{6+i} age {age}: remaining years ROSE {prev} -> {got}")
                prev = got
                if not near(le_vals[i][11], V.PLANNING[age]):
                    fails.append(f"LE!L{6+i} age {age}: shadow {V.PLANNING[age]} "
                                 f"vs excel {le_vals[i][11]}")
                if not near(le_vals[i][12], V.ceiling_for(age)):
                    fails.append(f"LE!M{6+i} age {age}: shadow {V.ceiling_for(age)} "
                                 f"vs excel {le_vals[i][12]}")

            for k in range(N_ROWS):
                for ci, letter in enumerate(COLS4):
                    if not near(four_vals[k][ci], shared[k][letter]):
                        fails.append(f"4%!{letter}{V.R0+k}: shadow {shared[k][letter]!r} "
                                     f"vs excel {four_vals[k][ci]!r}")
                for ci, letter in enumerate(COLSD):
                    if not near(dyn_vals[k][ci], dyn[k][letter]):
                        fails.append(f"DYN!{letter}{V.R0+k}: shadow {dyn[k][letter]!r} "
                                     f"vs excel {dyn_vals[k][ci]!r}")

            # the reported balance must be the ENDING balance of the last modelled year
            horizon = min(max(V.D["plan_age"] - V.D["start_age"], 1), N_ROWS)
            want_age = V.D["start_age"] + horizon - 1
            got_age = dyn_age[horizon - 1][0]
            if got_age != want_age:
                fails.append(f"horizon row age {got_age} != {want_age} "
                             f"(plan-to {V.D['plan_age']})")

            status = "OK  " if not fails else "FAIL"
            print(f"  [{status}] {name:<26} horizon={horizon:>3}  "
                  f"dyn yr1 after-tax={dyn[0]['AB']:>10,.0f}  "
                  f"balance@plan={dyn[horizon - 1]['AJ']:>13,.0f}")
            for f in fails[:6]:
                print(f"           {f}")
            if len(fails) > 6:
                print(f"           ... and {len(fails) - 6} more")
            total_fail += len(fails)

        # ---- Monte Carlo pass -------------------------------------------------
        # Restore the deterministic inputs first so MC scenarios vary only MC inputs.
        for key, row in addr.items():
            sh.Cells(row, 3).Value = spec.DEFAULTS[key]

        mcsh = wb.Worksheets("Monte Carlo")
        mc_addr = {}
        for r in range(1, 30):
            lab = mcsh.Cells(r, 2).Value
            if isinstance(lab, str):
                key = MC_LABEL_TO_KEY.get(lab.strip())
                if key:
                    mc_addr[key] = r
        mc_missing = set(MC_LABEL_TO_KEY.values()) - set(mc_addr)
        if mc_missing:
            print(f"FAILED to locate Monte Carlo input rows for: {sorted(mc_missing)}")
            return 1

        outsh = wb.Worksheets("MC Outcomes")
        last_out = 5 + spec.MAX_SIMS - 1
        print()
        for name, overrides in MC_SCENARIOS:
            for key, row in mc_addr.items():
                mcsh.Cells(row, 3).Value = spec.DEFAULTS[key]
            for key, val in overrides.items():
                mcsh.Cells(mc_addr[key], 3).Value = val
            xl.CalculateFullRebuild()

            V.configure(overrides)
            expected = V.monte_carlo()
            got = outsh.Range(f"B5:N{last_out}").Value
            fails = []
            for sim, exp in enumerate(expected):
                for col, key in MC_OUT_COLS:
                    val = got[sim][col - 2]
                    if not near(val, exp[key]):
                        fails.append(f"MCOUT!{chr(64 + col)}{5 + sim}: "
                                     f"shadow {exp[key]!r} vs excel {val!r}")
                if len(fails) > 6:
                    break
            status = "OK  " if not fails else "FAIL"
            surv = sum(e["dyn_lasts"] for e in expected) / len(expected)
            ess = sum(e["dyn_ess"] for e in expected) / len(expected)
            print(f"  [{status}] {name:<26} dyn lasts={surv:>6.1%}  "
                  f"essentials met={ess:>6.1%}")
            for f in fails[:6]:
                print(f"           {f}")
            total_fail += len(fails)

        # restore defaults so the saved workbook ships in its default state
        for key, row in mc_addr.items():
            mcsh.Cells(row, 3).Value = spec.DEFAULTS[key]
        for key, row in addr.items():
            sh.Cells(row, 3).Value = spec.DEFAULTS[key]
        xl.CalculateFullRebuild()
        wb.Save()
    finally:
        # cleanup must not raise, or it masks whatever actually went wrong
        try:
            wb.Close(True)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    print()
    if total_fail:
        print(f"STATUS: FAILURES - {total_fail} mismatches across scenarios")
        return 1
    print(f"STATUS: success - all {len(SCENARIOS)} input scenarios and "
          f"{len(MC_SCENARIOS)} Monte Carlo scenarios match the shadow model")
    return 0


N_ROWS = spec.TABLE_ROWS

if __name__ == "__main__":
    sys.exit(main())
