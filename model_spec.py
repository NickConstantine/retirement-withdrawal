"""Shared constants and reference data for the Retirement Withdrawal Strategies planner.

Imported by both `build_retirement_planner.py` (which writes the Excel formulas)
and `validate_model.py` (which recomputes the same model in Python and compares
its answers against the values Excel calculated). Keeping the data in one module
guarantees the workbook and its validator cannot drift apart.
"""

MAX_SIMS = 1000
TABLE_ROWS = 71          # ages 40..110 is 71 rows; the horizon can never exceed this
MC_SEED = 20260811

# Age -> remaining-years anchors are entered by the user; the table itself spans:
LE_FIRST_AGE = 35
LE_LAST_AGE = 110

# ---------------------------------------------------------------------------
# Annual returns and inflation, 1928-2025.
#   stock  = S&P 500 total return
#   bond   = 10-year US Treasury total return
#   Source: Aswath Damodaran, NYU Stern, "Annual Returns on Investments".
#   cpi    = US CPI-U annual average inflation
#   Source: US Bureau of Labor Statistics.
# Sampling all three from the SAME year keeps the return/inflation relationship
# intact, which is what makes the 1970s stagflation years bite in the bootstrap.
# ---------------------------------------------------------------------------
HISTORICAL = [
    (1928, 0.4381, 0.0084, -0.0117), (1929, -0.0830, 0.0420, 0.0000),
    (1930, -0.2512, 0.0454, -0.0234), (1931, -0.4384, -0.0256, -0.0902),
    (1932, -0.0864, 0.0879, -0.0987), (1933, 0.4998, 0.0186, -0.0514),
    (1934, -0.0119, 0.0796, 0.0311), (1935, 0.4674, 0.0447, 0.0224),
    (1936, 0.3194, 0.0502, 0.0151), (1937, -0.3534, 0.0138, 0.0361),
    (1938, 0.2928, 0.0421, -0.0208), (1939, -0.0110, 0.0441, -0.0144),
    (1940, -0.1067, 0.0540, 0.0072), (1941, -0.1277, -0.0202, 0.0500),
    (1942, 0.1917, 0.0229, 0.1088), (1943, 0.2506, 0.0249, 0.0613),
    (1944, 0.1903, 0.0258, 0.0169), (1945, 0.3582, 0.0380, 0.0227),
    (1946, -0.0843, 0.0313, 0.0830), (1947, 0.0520, 0.0092, 0.1436),
    (1948, 0.0570, 0.0195, 0.0807), (1949, 0.1830, 0.0466, -0.0120),
    (1950, 0.3081, 0.0043, 0.0130), (1951, 0.2368, -0.0030, 0.0788),
    (1952, 0.1815, 0.0227, 0.0192), (1953, -0.0121, 0.0414, 0.0075),
    (1954, 0.5256, 0.0329, 0.0075), (1955, 0.3260, -0.0134, -0.0037),
    (1956, 0.0744, -0.0226, 0.0149), (1957, -0.1046, 0.0680, 0.0331),
    (1958, 0.4372, -0.0210, 0.0276), (1959, 0.1206, -0.0265, 0.0072),
    (1960, 0.0034, 0.1164, 0.0170), (1961, 0.2664, 0.0206, 0.0101),
    (1962, -0.0881, 0.0569, 0.0100), (1963, 0.2261, 0.0168, 0.0132),
    (1964, 0.1642, 0.0373, 0.0132), (1965, 0.1240, 0.0072, 0.0161),
    (1966, -0.0997, 0.0291, 0.0286), (1967, 0.2380, -0.0158, 0.0309),
    (1968, 0.1081, 0.0327, 0.0427), (1969, -0.0824, -0.0501, 0.0545),
    (1970, 0.0356, 0.1675, 0.0572), (1971, 0.1422, 0.0979, 0.0443),
    (1972, 0.1876, 0.0282, 0.0327), (1973, -0.1431, 0.0366, 0.0622),
    (1974, -0.2590, 0.0199, 0.1103), (1975, 0.3700, 0.0361, 0.0912),
    (1976, 0.2383, 0.1598, 0.0576), (1977, -0.0698, 0.0129, 0.0650),
    (1978, 0.0651, -0.0078, 0.0762), (1979, 0.1852, 0.0067, 0.1125),
    (1980, 0.3174, -0.0299, 0.1350), (1981, -0.0470, 0.0820, 0.1032),
    (1982, 0.2042, 0.3281, 0.0616), (1983, 0.2234, 0.0320, 0.0321),
    (1984, 0.0615, 0.1373, 0.0430), (1985, 0.3124, 0.2571, 0.0355),
    (1986, 0.1849, 0.2428, 0.0186), (1987, 0.0581, -0.0496, 0.0365),
    (1988, 0.1654, 0.0822, 0.0414), (1989, 0.3148, 0.1769, 0.0482),
    (1990, -0.0306, 0.0624, 0.0539), (1991, 0.3023, 0.1500, 0.0423),
    (1992, 0.0749, 0.0936, 0.0303), (1993, 0.0997, 0.1421, 0.0296),
    (1994, 0.0133, -0.0804, 0.0261), (1995, 0.3720, 0.2348, 0.0281),
    (1996, 0.2268, 0.0143, 0.0293), (1997, 0.3310, 0.0994, 0.0234),
    (1998, 0.2834, 0.1492, 0.0155), (1999, 0.2089, -0.0825, 0.0219),
    (2000, -0.0903, 0.1666, 0.0338), (2001, -0.1185, 0.0557, 0.0283),
    (2002, -0.2197, 0.1512, 0.0159), (2003, 0.2836, 0.0038, 0.0227),
    (2004, 0.1074, 0.0449, 0.0268), (2005, 0.0483, 0.0287, 0.0339),
    (2006, 0.1561, 0.0196, 0.0323), (2007, 0.0548, 0.1021, 0.0285),
    (2008, -0.3655, 0.2010, 0.0384), (2009, 0.2594, -0.1112, -0.0036),
    (2010, 0.1482, 0.0846, 0.0164), (2011, 0.0210, 0.1604, 0.0316),
    (2012, 0.1589, 0.0297, 0.0207), (2013, 0.3215, -0.0910, 0.0146),
    (2014, 0.1352, 0.1075, 0.0162), (2015, 0.0138, 0.0128, 0.0012),
    (2016, 0.1177, 0.0069, 0.0126), (2017, 0.2161, 0.0280, 0.0213),
    (2018, -0.0423, -0.0002, 0.0244), (2019, 0.3121, 0.0964, 0.0181),
    (2020, 0.1802, 0.1133, 0.0123), (2021, 0.2847, -0.0442, 0.0470),
    (2022, -0.1804, -0.1783, 0.0800), (2023, 0.2606, 0.0388, 0.0413),
    (2024, 0.2488, -0.0164, 0.0295), (2025, 0.1778, 0.0780, 0.0273),
]

# ---------------------------------------------------------------------------
# Social Security claiming factors, born 1960 or later (full retirement age 67).
# Source: SSA, ssa.gov/benefits/retirement/planner/1960.html
# ---------------------------------------------------------------------------
SS_FACTORS = [(62, 0.700), (63, 0.750), (64, 0.800), (65, 0.867), (66, 0.933),
              (67, 1.000), (68, 1.080), (69, 1.160), (70, 1.240)]

# ---------------------------------------------------------------------------
# IRS Uniform Lifetime Table (used for RMDs), effective 2022 onward.
# Source: IRS Publication 590-B, Appendix B, Table III.
# ---------------------------------------------------------------------------
RMD_DIVISORS = [
    (72, 27.4), (73, 26.5), (74, 25.5), (75, 24.6), (76, 23.7), (77, 22.9),
    (78, 22.0), (79, 21.1), (80, 20.2), (81, 19.4), (82, 18.5), (83, 17.7),
    (84, 16.8), (85, 16.0), (86, 15.2), (87, 14.4), (88, 13.7), (89, 12.9),
    (90, 12.2), (91, 11.5), (92, 10.8), (93, 10.1), (94, 9.5), (95, 8.9),
    (96, 8.4), (97, 7.8), (98, 7.3), (99, 6.8), (100, 6.4), (101, 6.0),
    (102, 5.6), (103, 5.2), (104, 4.9), (105, 4.5), (106, 4.2), (107, 3.9),
    (108, 3.7), (109, 3.4), (110, 3.1), (111, 2.9), (112, 2.6), (113, 2.4),
    (114, 2.1), (115, 1.9), (116, 1.7), (117, 1.5), (118, 1.4), (119, 1.2),
    (120, 2.0),
]

# Scenario detail defaults (year -> value). Years past the table use the base input.
SCENARIO_YEARS = 10
CRASH_RETURNS = {1: -0.20, 2: -0.10, 3: 0.15, 4: 0.15, 5: 0.12}
STAGFLATION_YEARS = 8
STAGFLATION_RETURN = 0.01
STAGFLATION_INFLATION = 0.07     # real stagflation is high inflation, not just weak returns

# Anchor scaling used for ages 75+ on the Life Expectancy tab.
LE_TAIL_ANCHORS = [(75, 0.802), (80, 0.623), (85, 0.467), (90, 0.335),
                   (95, 0.234), (100, 0.162), (105, 0.114), (110, 0.084)]

# ---------------------------------------------------------------------------
# Default input values. The builder writes these into the yellow cells and the
# validator uses them as its baseline, so "defaults" are defined exactly once.
# ---------------------------------------------------------------------------
DEFAULTS = {
    "start_age": 50,
    "plan_age": 90,
    "household": "Single",
    "spouse_gap": 3,
    # ---- accounts (the starting portfolio is the sum of these three) ----
    "taxable0": 1_000_000,
    "basis_share": 0.65,
    "traditional0": 2_100_000,
    "roth0": 400_000,
    "exp_ret": 0.06,
    "fee": 0.005,
    "infl": 0.03,
    "scenario": "Steady",
    "essential": 100_000,
    "healthcare": 20_000,
    "drift": 0.0,
    "rate4": 0.04,
    "fixed_basis": "4% of starting portfolio",
    "floor": 0.03,
    "ceiling": 0.06,
    "ceiling_basis": "Age-graduated",
    "ceiling_mult": 2.0,
    "shock": 0.05,
    "shock_basis": "Real",
    "shortfall_mode": "Follow the rule",
    "current_4pct": 0,
    # ---- taxes: effective ordinary rates by retirement phase ----
    "tax_early": 0.10,
    "tax_ss": 0.15,
    "tax_rmd": 0.20,
    "cg_rate": 0.15,
    "div_yield": 0.02,
    "rmd_age": 75,
    "full_ss": 30_000,
    "ss_age": 67,
    "ltc_on": "Yes",
    "ltc_cost": 75_000,
    "ltc_years": 3,
    "ltc_age": 84,
    "cur_age_le": 49,
    "le_cur": 32.6,
    "le62": 22.4,
    "le67": 18.8,
    "le70": 16.7,
    "mc_method": "Parametric",
    "mc_sims": MAX_SIMS,
    "mc_vol": 0.12,
    "mc_infl_vol": 0.02,
    "mc_stock": 0.60,
    "mc_ltc_prob": 0.50,
}

SS_TAXABLE_SHARE = 0.85          # simplification: 85% of the benefit is taxable
INFLATION_FLOOR = -0.05          # parametric inflation draws are floored here

# A year counts as a shortfall only if income misses the need by more than this, in
# today's dollars. In "Withdraw enough to cover needs" mode the withdrawal is solved so
# that income equals the need EXACTLY, which leaves a long tail of results sitting on
# zero to within floating-point dust. Without a tolerance, Excel and Python round those
# to opposite sides and disagree about whether the year was funded.
SHORTFALL_TOL = 1.0

# The two tabs are named here so the builder and both test harnesses cannot disagree
# about them. The fixed-spending tab is no longer only the 4% rule: with the needs
# basis it targets your actual spending instead of a percentage of the opening balance.
FIXED_SHEET = "Fixed Spending"
BASIS_PCT = "4% of starting portfolio"
BASIS_NEEDS = "Your actual spending needs"
