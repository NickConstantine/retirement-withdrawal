"""Check every declared input actually changes the model's output.

Class of bug this catches: an input is added to the Inputs tab and to DEFAULTS,
but is never referenced by a formula (or is referenced only in a branch that the
current configuration never reaches). The workbook still builds, still shows the
yellow cell, still recalculates without error -- and silently ignores the user.

Each input is perturbed inside a configuration chosen so that it SHOULD matter.
If nothing downstream moves, the input is not wired up.
"""
import model_spec as spec
import validate_model as V


def fingerprint():
    """A number that depends on essentially every part of the model."""
    shared = V.four_rule()
    dyn = V.dynamic(shared)
    n = min(max(V.D["plan_age"] - V.D["start_age"], 1), V.N)
    parts = []
    for rows, keys in ((shared, ("X", "Z", "AE", "AF")),
                       (dyn, ("AB", "AD", "AI", "AJ", "AK"))):
        for k in keys:
            parts.append(sum(abs(r[k]) for r in rows[:n]))
    parts.append(sum(V.PLANNING[a] for a in V.AGES))
    parts.append(sum(V.ceiling_for(a) for a in V.AGES))
    return sum(parts)


def mc_fingerprint():
    out = V.monte_carlo()
    # include the essentials metrics: a long-term care event changes what you NEED,
    # which shows up in the coverage tests rather than in balances or withdrawals
    return sum(o["dyn_bal"] + o["four_bal"] + o["dyn_low"] + o["four_low"]
               + o["dyn_ess"] * 1e6 + o["four_ess"] * 1e6
               + o["dyn_share"] * 1e6 + o["four_share"] * 1e6 for o in out[:60])


# base config, then per-input (context, perturbed value).
# The context makes the input relevant; without it some inputs are legitimately inert.
CASES = [
    ("start_age", {}, 55),
    ("plan_age", {}, 95),
    ("household", {}, "Couple"),
    ("spouse_gap", {"household": "Couple"}, -8),
    ("taxable0", {}, 1_500_000),
    ("basis_share", {}, 0.20),
    ("traditional0", {}, 2_600_000),
    ("roth0", {}, 900_000),
    ("exp_ret", {}, 0.07),
    ("fee", {}, 0.015),
    ("infl", {}, 0.045),
    ("scenario", {}, "Stagflation"),
    ("essential", {}, 130_000),
    ("healthcare", {}, 35_000),
    ("drift", {}, -0.01),
    ("rate4", {}, 0.05),
    ("floor", {"ceiling_basis": "Fixed %"}, 0.045),
    ("ceiling", {"ceiling_basis": "Fixed %"}, 0.05),
    ("ceiling_basis", {}, "Fixed %"),
    ("ceiling_mult", {"ceiling_basis": "Age-graduated"}, 1.4),
    ("shock", {}, 0.02),
    ("shock_basis", {"scenario": "Early crash"}, "Nominal"),
    ("current_4pct", {}, 150_000),
    ("tax_early", {}, 0.25),
    ("tax_ss", {}, 0.30),
    ("tax_rmd", {}, 0.35),
    ("cg_rate", {}, 0.30),
    ("div_yield", {}, 0.05),
    ("rmd_age", {}, 73),
    ("full_ss", {}, 45_000),
    ("ss_age", {}, 62),
    ("ltc_on", {}, "No"),
    ("ltc_cost", {}, 200_000),
    ("ltc_years", {}, 6),
    ("ltc_age", {}, 70),
    ("cur_age_le", {}, 55),
    ("le_cur", {}, 25.0),
    ("le62", {}, 18.0),
    ("le67", {}, 15.0),
    ("le70", {}, 12.0),
]

MC_CASES = [
    ("mc_method", {}, "Historical bootstrap"),
    ("mc_vol", {}, 0.25),
    ("mc_infl_vol", {}, 0.06),
    ("mc_stock", {"mc_method": "Historical bootstrap"}, 0.95),
    ("mc_ltc_prob", {}, 1.0),
]

print("=" * 78)
print("DETERMINISTIC MODEL — does each input change the result?")
print("=" * 78)
dead = []
for key, ctx, new in CASES:
    V.configure(ctx)
    before = fingerprint()
    V.configure(dict(ctx, **{key: new}))
    after = fingerprint()
    moved = abs(after - before) > 1e-9 * max(1.0, abs(before))
    if not moved:
        dead.append(key)
    print(f"  {'OK  ' if moved else 'DEAD'}  {key:<16} {str(new):>22}  "
          f"delta {abs(after - before):>18,.2f}")

print()
print("=" * 78)
print("MONTE CARLO INPUTS")
print("=" * 78)
for key, ctx, new in MC_CASES:
    V.configure(ctx)
    before = mc_fingerprint()
    V.configure(dict(ctx, **{key: new}))
    after = mc_fingerprint()
    moved = abs(after - before) > 1e-9 * max(1.0, abs(before))
    if not moved:
        dead.append(key)
    print(f"  {'OK  ' if moved else 'DEAD'}  {key:<16} {str(new):>22}  "
          f"delta {abs(after - before):>18,.2f}")

print()
missing = sorted(set(spec.DEFAULTS) - {c[0] for c in CASES} - {c[0] for c in MC_CASES})
if missing:
    print(f"NOT COVERED BY THIS CHECK: {missing}")
if dead:
    print(f"STATUS: {len(dead)} input(s) had no effect: {dead}")
else:
    print("STATUS: success - every input changes the model")
