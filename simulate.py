"""
simulate.py: Monte Carlo expected claims and reinsurance calculator.

Steps, in order:
  1. Ask for run settings (press Enter to accept the default in brackets).
  2. Read the input workbook (tabs "Policies" and "Mortality").
  3. Check the data. If anything is wrong, stop and list the problem rows.
  4. Simulate one year of deaths, many times over.
  5. Split every claim into retained and ceded parts at the retention limit.
  6. Calculate the risk metrics (means, percentiles, VaR 99%, CTE 99%)
     for claim amounts and for claim frequency (number of claims).
  7. Save everything to output/results_R<retention>_N<sims>_S<seed>.xlsx
  8. Pick a representative year and the VaR 99% year, replay only those two
     simulations from the same random stream, and save the policy-level
     ledgers to output/granularity_R<retention>_N<sims>_S<seed>.xlsx

A "claim" is a death of a life with sum assured > 0. A death with nothing
insured is not a claim and is not counted in claim frequency.
"""

import hashlib
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------- defaults
DEFAULT_INPUT = "sample_input.xlsx"
DEFAULT_RETENTION = 3_400_000
DEFAULT_SIMULATIONS = 10_000
DEFAULT_SEED = 42
DEFAULT_PERCENTILES = [50, 75, 90, 95, 99, 99.5]
TAIL_LEVEL = 99            # VaR and CTE are measured at this percentile
TAIL_RANK_BY = "Retained"  # column used to pick the VaR / CTE years ("Gross" or "Retained")
FREQUENCY_RANK_BY = "Claims"   # column used to pick the frequency VaR / CTE years
OUTPUT_FOLDER = Path("output")
# Mortality adjustment by sum assured band: qx is DIVIDED by the divisor of the band the life falls in.
# Each entry is (lower bound, divisor); the lower bound is included, the upper bound is the next band's
# lower bound. Keep the list from the highest bound to the lowest. A divisor below 1 raises qx.
# Lives below the lowest bound are not adjusted. Use an empty list to switch the adjustment off.
MORTALITY_BANDS = [(20_000_000, 2.3),
                   (10_000_000, 1.25),
                   (5_000_000, 0.9)]
CELLS_PER_BATCH = 10_000_000   # memory control only; does not change results
# Sum assured bands for the claim frequency breakdown: (lower bound included, upper bound excluded).
# The last band has no upper bound. Lives with sum assured 0 never claim, so they are in no band.
CLAIM_BANDS = [(0, 5_000_000), (5_000_000, 10_000_000), (10_000_000, 20_000_000), (20_000_000, None)]
EXPORT_GRANULARITY = True      # default answer to the granularity prompt; the prompt can override it
EXCEL_MAX_DATA_ROWS = 1_048_575   # Excel allows 1,048,576 rows including the header


# ---------------------------------------------------------------- 1. settings
def ask(question, default, convert):
    """Ask a question; an empty answer uses the default shown in brackets."""
    while True:
        answer = input(f"{question} [{default}]: ").strip() or str(default)
        try:
            return convert(answer)
        except ValueError as problem:
            print(f"  Not valid: {problem}. Try again.")


def positive_number(text):
    value = float(text.replace(",", ""))
    if not np.isfinite(value) or value <= 0:
        raise ValueError("must be a finite number greater than zero")
    return value


def positive_whole_number(text):
    value = int(text.replace(",", ""))
    if value <= 0:
        raise ValueError("must be a whole number greater than zero")
    return value


def non_negative_whole_number(text):
    value = int(text.replace(",", ""))
    if value < 0:
        raise ValueError("must be a whole number, zero or more")
    return value


def yes_no(text):
    answer = text.strip().lower()
    if answer in ("y", "yes", "true", "1", "on"):
        return True
    if answer in ("n", "no", "false", "0", "off"):
        return False
    raise ValueError("answer yes or no")


def percentile_list(text):
    values = sorted({float(part) for part in text.replace(",", " ").split()})
    if not values or any(not np.isfinite(v) or v <= 0 or v >= 100 for v in values):
        raise ValueError("each percentile must be between 0 and 100")
    return values


def ask_settings():
    print("\nMonte Carlo Claims & Reinsurance Calculator\n")
    return {
        "input_file": ask("Input workbook", DEFAULT_INPUT, str),
        "retention": ask("Retention limit", f"{DEFAULT_RETENTION:,}", positive_number),
        "simulations": ask("Number of simulations", f"{DEFAULT_SIMULATIONS:,}", positive_whole_number),
        "seed": ask("Random seed", DEFAULT_SEED, non_negative_whole_number),
        "percentiles": ask("Percentiles", ", ".join(f"{p:g}" for p in DEFAULT_PERCENTILES),
                           percentile_list),
        "granularity": ask("Export policy-level granularity (yes/no)",
                           "yes" if EXPORT_GRANULARITY else "no", yes_no),
    }


# ---------------------------------------------------------------- 2-3. read and check
def read_inputs(path):
    if not Path(path).exists():
        raise SystemExit(f"\nSTOPPED: input file not found: {path}")
    tabs = pd.ExcelFile(path).sheet_names
    for tab in ["Policies", "Mortality"]:
        if tab not in tabs:
            raise SystemExit(f"\nSTOPPED: the workbook has no tab named '{tab}' (found: {', '.join(tabs)})")
    policies = pd.read_excel(path, sheet_name="Policies")
    mortality = pd.read_excel(path, sheet_name="Mortality")
    for name, table, needed in [("Policies", policies, ["PolicyID", "Age", "SumAssured"]),
                                ("Mortality", mortality, ["Age", "qx"])]:
        missing = [c for c in needed if c not in table.columns]
        if missing:
            raise SystemExit(f"\nSTOPPED: the {name} tab is missing column(s): {', '.join(missing)}")
        if table.empty:
            raise SystemExit(f"\nSTOPPED: the {name} tab has no rows")
    return policies, mortality


def is_whole_number(series):
    numbers = pd.to_numeric(series, errors="coerce")
    return numbers.notna() & (numbers == np.floor(numbers))


def check_inputs(policies, mortality):
    """Return policies with a qx column attached, or stop and list every problem."""
    problems = []   # each entry: (tab, excel row, column, value, issue)

    def flag(tab, mask, column, issue):
        for i in mask[mask].index:
            value = (policies if tab == "Policies" else mortality).at[i, column]
            problems.append((tab, i + 2, column, value, issue))   # +2: header row, 1-based

    # Mortality table
    flag("Mortality", ~is_whole_number(mortality["Age"]), "Age", "age is missing or not a whole number")
    flag("Mortality", mortality["Age"].duplicated(keep=False), "Age", "age listed more than once")
    qx = pd.to_numeric(mortality["qx"], errors="coerce")
    flag("Mortality", qx.isna() | (qx < 0) | (qx > 1), "qx", "qx must be a number between 0 and 1")

    # Policies
    flag("Policies", policies["PolicyID"].isna(), "PolicyID", "PolicyID is blank")
    flag("Policies", policies["PolicyID"].duplicated(keep=False) & policies["PolicyID"].notna(),
         "PolicyID", "PolicyID appears more than once")
    whole_age = is_whole_number(policies["Age"])
    flag("Policies", ~whole_age, "Age", "age is missing or not a whole number")
    known_ages = set(pd.to_numeric(mortality["Age"], errors="coerce").dropna())
    flag("Policies", whole_age & ~pd.to_numeric(policies["Age"], errors="coerce").isin(known_ages),
         "Age", "age not found in the Mortality tab")
    sa = pd.to_numeric(policies["SumAssured"], errors="coerce")
    # Zero is allowed (nil cover, paid-up). Such lives can die but never produce a claim.
    flag("Policies", sa.isna() | (sa < 0) | ~np.isfinite(sa),
         "SumAssured", "sum assured is blank, negative or not a number")

    if problems:
        report = pd.DataFrame(problems, columns=["Tab", "Excel row", "Column", "Value", "Problem"])
        OUTPUT_FOLDER.mkdir(exist_ok=True)
        report.to_csv(OUTPUT_FOLDER / "input_problems.csv", index=False)
        print(f"\nSTOPPED: found {len(report)} problem(s) in the input. First 20:\n")
        print(report.head(20).to_string(index=False))
        raise SystemExit(f"\nFull list saved to {OUTPUT_FOLDER / 'input_problems.csv'}. Fix them and run again.")

    # Attach each life's qx by looking up its age
    rates = pd.Series(qx.values, index=mortality["Age"].astype(int))
    checked = policies.copy()
    checked["Age"] = checked["Age"].astype(int)
    checked["SumAssured"] = sa
    checked["qx_table"] = checked["Age"].map(rates)
    checked["qx_divisor"] = band_divisors(checked["SumAssured"])
    checked["qx"] = np.minimum(checked["qx_table"] / checked["qx_divisor"], 1.0)   # qx can never exceed 1
    return checked


def band_divisors(sum_assured):
    """The mortality divisor for each life, from MORTALITY_BANDS (1.0 if below every band)."""
    sa = sum_assured.to_numpy(dtype=float)
    return np.select([sa >= lower for lower, _ in MORTALITY_BANDS],
                     [divisor for _, divisor in MORTALITY_BANDS], default=1.0)


def band_label(index):
    lower, divisor = MORTALITY_BANDS[index]
    upper = MORTALITY_BANDS[index - 1][0] if index > 0 else None
    return f"{lower:,} to below {upper:,}" if upper else f"{lower:,} and above"


def portfolio_fingerprint(policies):
    """A short code that changes if any age, sum assured, qx or row order changes."""
    data = policies[["Age", "SumAssured", "qx"]].to_numpy(dtype=np.float64).tobytes()
    return hashlib.md5(data).hexdigest()[:12]


def claim_band_labels():
    labels = []
    for lower, upper in CLAIM_BANDS:
        labels.append(f"Claims {lower / 1e6:g}M+" if upper is None
                      else f"Claims {lower / 1e6:g}-{upper / 1e6:g}M")
    return labels


def claim_band_masks(sum_assured):
    """For each band label, a True/False array over the lives (sum assured > 0 and inside the band)."""
    sa = np.asarray(sum_assured, dtype=float)
    return {label: (sa > 0) & (sa >= lower) & ((sa < upper) if upper is not None else True)
            for label, (lower, upper) in zip(claim_band_labels(), CLAIM_BANDS)}


# ---------------------------------------------------------------- 4-5. simulate
def simulate(qx, sum_assured, retention, simulations, seed):
    """
    Each row of a batch is one simulated year; each column is one life.
    A life dies in a simulation if its random number U < qx.

    Because a claim is either the full sum assured or nothing, each life's
    amounts can be worked out once, up front:
        claim count       = 1 if sum assured > 0, else 0
        ceded claim count = 1 if sum assured > retention, else 0
        retained          = min(sum assured, retention)
        ceded             = max(0, sum assured - retention)
    A simulation's total is then the sum of these amounts over the lives that died.
    """
    rng = np.random.default_rng(seed)
    retained_if_dies = np.minimum(sum_assured, retention)
    ceded_if_dies = sum_assured - retained_if_dies
    band_masks = claim_band_masks(sum_assured)
    columns = ["Claims", "Ceded claims", "Gross", "Retained", "Ceded"] + list(band_masks)
    amounts = np.column_stack([(sum_assured > 0).astype(float), (sum_assured > retention).astype(float),
                               sum_assured, retained_if_dies, ceded_if_dies]
                              + [mask.astype(float) for mask in band_masks.values()])

    totals = np.zeros((simulations, len(columns)))
    batch = max(1, CELLS_PER_BATCH // len(qx))      # simulations handled per pass
    for start in range(0, simulations, batch):
        end = min(start + batch, simulations)
        dies = rng.random((end - start, len(qx))) < qx           # True = death
        totals[start:end] = dies.astype(np.float64) @ amounts     # add up amounts of those who died
        print(f"  simulated {end:,} of {simulations:,}", end="\r")
    print()

    table = pd.DataFrame(totals, columns=columns)
    table.insert(0, "Simulation", np.arange(1, simulations + 1))
    return table


# ---------------------------------------------------------------- 6. metrics
def exact_values(policies, retention):
    """Exact expected values (sum of probability x amount) for every simulated column."""
    qx, sa = policies["qx"], policies["SumAssured"]
    exact = {"Claims": qx[sa > 0].sum(),
             "Ceded claims": qx[sa > retention].sum(),
             "Gross": (qx * sa).sum(),
             "Retained": (qx * np.minimum(sa, retention)).sum(),
             "Ceded": (qx * np.maximum(0, sa - retention)).sum()}
    for label, mask in claim_band_masks(sa).items():
        exact[label] = qx.to_numpy()[mask].sum()
    return exact


def var_scenarios(sims, rank_by, level=TAIL_LEVEL):
    """
    The simulation(s) whose rank_by total is closest to its level% percentile.
    If several years tie at that value (common for claim counts), all of them
    are used and VaR for every column is their average, so the result does not
    depend on which tied year happens to come first. For amounts, retained + ceded = gross.
    """
    values = sims[rank_by].to_numpy()
    distance = np.abs(values - np.percentile(values, level))
    return np.flatnonzero(distance == distance.min())


def cte_weights(sims, rank_by, level=TAIL_LEVEL):
    """
    Weights over all simulations that define the worst (100 - level)% by rank_by.
    Years strictly above the cut-off get weight 1. Years tied at the cut-off share
    the remaining places equally, so the result does not depend on sort order.
    The weights add up to the tail size k. CTE of a column = sum(weight x value) / k.
    """
    values = sims[rank_by].to_numpy()
    k = max(1, int(round(len(values) * (100 - level) / 100)))
    cutoff = np.sort(values)[-k]
    weights = (values > cutoff).astype(float)
    tied = np.flatnonzero(values == cutoff)
    weights[tied] = (k - weights.sum()) / len(tied)
    return weights, k


def metrics_table(sims, columns, exact, rank_by):
    """Mean, spread and tail measures for the given columns. Returns (table, VaR rows)."""
    var_rows = var_scenarios(sims, rank_by)
    weights, k = cte_weights(sims, rank_by)
    rows = {}
    for col in columns:
        x = sims[col].to_numpy()
        rows[col] = {
            "Mean (simulated)": x.mean(),
            "Standard error of mean": x.std(ddof=1) / np.sqrt(len(x)),
            "Expected value (exact formula)": exact[col],
            "Standard deviation": x.std(ddof=1),
            f"VaR {TAIL_LEVEL}%": x[var_rows].mean(),
            f"CTE {TAIL_LEVEL}%": (x * weights).sum() / k,
            "Largest simulated total": x.max(),
        }
    return pd.DataFrame(rows).rename_axis("Metric").reset_index(), var_rows


def list_simulations(sims, rows, show=10):
    """Simulation numbers for Run Info, e.g. '3 years: 12, 480, 9044'."""
    numbers = sims["Simulation"].to_numpy()[rows]
    shown = ", ".join(str(n) for n in numbers[:show]) + (", ..." if len(numbers) > show else "")
    return f"{len(numbers)} year{'s' if len(numbers) > 1 else ''}: {shown}"


def calculate_metrics(sims, exact, percentiles):
    # Amounts (severity)
    amount_cols = ["Gross", "Retained", "Ceded"]
    summary, var_rows = metrics_table(sims, amount_cols, exact, TAIL_RANK_BY)
    # Same basis as VaR in the Summary sheet: for each percentile, find the simulated year(s) whose
    # TAIL_RANK_BY total is closest to that percentile, then report every column for those same years.
    pct_rows = [var_scenarios(sims, TAIL_RANK_BY, level=p) for p in percentiles]
    pct = pd.DataFrame({"Percentile": [f"{p:g}%" for p in percentiles]})
    for col in amount_cols:
        values = sims[col].to_numpy()
        pct[col] = [values[rows].mean() for rows in pct_rows]
    pct["Scenario (simulation numbers)"] = [list_simulations(sims, rows) for rows in pct_rows]

    # Frequency (number of claims)
    count_cols = ["Claims", "Ceded claims"] + claim_band_labels()
    frequency, freq_var_rows = metrics_table(sims, count_cols, exact, FREQUENCY_RANK_BY)
    # Same basis as the frequency VaR: years ranked by FREQUENCY_RANK_BY, every column read from those years.
    freq_pct_rows = [var_scenarios(sims, FREQUENCY_RANK_BY, level=p) for p in percentiles]
    freq_pct = pd.DataFrame({"Metric": [f"Percentile {p:g}%" for p in percentiles]})
    for col in count_cols:
        values = sims[col].to_numpy()
        freq_pct[col] = [values[rows].mean() for rows in freq_pct_rows]
    freq_pct["Scenario (simulation numbers)"] = [list_simulations(sims, rows) for rows in freq_pct_rows]
    frequency = pd.concat([frequency, freq_pct], ignore_index=True)

    var_sims = {"amounts": list_simulations(sims, var_rows),
                "frequency": list_simulations(sims, freq_var_rows)}
    return summary, pct, frequency, var_sims, var_rows


# ---------------------------------------------------------------- 7. save results
def file_tag(retention, simulations, seed):
    return f"R{retention:.0f}_N{simulations}_S{seed}"


def write_run_info(writer, run_info, formats):
    """Run Info mixes text, whole numbers and decimals, so each cell gets its own format."""
    sheet = writer.book.add_worksheet("Run Info")
    writer.sheets["Run Info"] = sheet
    sheet.write(0, 0, "Setting", formats["plain"])
    sheet.write(0, 1, "Value", formats["plain"])
    for r, (setting, value) in enumerate(run_info.itertuples(index=False), start=1):
        sheet.write(r, 0, setting, formats["plain"])
        if isinstance(value, (int, np.integer)):
            sheet.write_number(r, 1, int(value), formats["money"])
        elif isinstance(value, (float, np.floating)):
            whole = float(value).is_integer()
            sheet.write_number(r, 1, float(value), formats["money"] if whole else formats["count"])
        else:
            sheet.write_string(r, 1, str(value), formats["plain"])
    sheet.set_column(0, 0, 44)
    sheet.set_column(1, 1, 60)
    sheet.freeze_panes(1, 0)


def save_results(path, run_info, summary, pct, frequency, sims):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        formats = {"money": book.add_format({"num_format": "#,##0", "font_name": "Arial"}),
                   "count": book.add_format({"num_format": "#,##0.00", "font_name": "Arial"}),
                   "plain": book.add_format({"font_name": "Arial"})}
        write_run_info(writer, run_info, formats)
        tables = [("Summary", summary), ("Percentiles", pct), ("Frequency", frequency)]
        if len(sims) <= EXCEL_MAX_DATA_ROWS:
            tables.append(("Simulations", sims))
        else:
            csv_path = path.with_name(path.stem + "_simulations.csv")
            sims.to_csv(csv_path, index=False)
            print(f"  Too many simulations for an Excel tab; saved the Simulations table to {csv_path}")
        for name, table in tables:
            table.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.set_column(0, 0, 32, formats["plain"])
            sheet.set_column(1, len(table.columns) - 1, 20,
                             formats["count"] if name == "Frequency" else formats["money"])
            sheet.freeze_panes(1, 0)


# ---------------------------------------------------------------- 8. policy-level granularity
LEDGER_COLUMNS = ["PolicyID", "Age", "SumAssured", "qx", "Claim",
                  "Gross Claim", "Retained Claim", "Ceded Claim"]


def representative_simulation(sims, exact):
    """
    Row of the simulation closest to the exact expected values, measured as
    sqrt(Z_claims^2 + Z_gross^2 + Z_retained^2 + Z_ceded^2), where each Z is the
    deviation from the exact expected value divided by that column's standard deviation.
    """
    distance_sq = np.zeros(len(sims))
    for col in ["Claims", "Gross", "Retained", "Ceded"]:
        x = sims[col].to_numpy()
        spread = x.std(ddof=1) if len(x) > 1 else 0.0
        if spread > 0:
            distance_sq += ((x - exact[col]) / spread) ** 2
    return int(np.argmin(distance_sq))


def replay_deaths(qx, seed, rows):
    """
    Re-create the death flags of the given simulation rows (0-based) from the same
    random stream as simulate(). Nothing else is kept, so memory stays small.
    The stream is the same whatever the batch size, because numbers are drawn row by row.
    """
    rng = np.random.default_rng(seed)
    wanted = sorted(set(rows))
    last = wanted[-1]
    batch = max(1, CELLS_PER_BATCH // len(qx))
    found = {}
    for start in range(0, last + 1, batch):
        end = min(start + batch, last + 1)
        draws = rng.random((end - start, len(qx)))
        for row in wanted:
            if start <= row < end:
                found[row] = draws[row - start] < qx
        print(f"  replayed {end:,} of {last + 1:,}", end="\r")
    print()
    return found


def build_ledger(policies, died, retention):
    sa = policies["SumAssured"].to_numpy(dtype=float)
    claim = died & (sa > 0)
    return pd.DataFrame({
        "PolicyID": policies["PolicyID"].to_numpy(),
        "Age": policies["Age"].to_numpy(),
        "SumAssured": sa,
        "qx": policies["qx"].to_numpy(),
        "Claim": np.where(claim, "YES", "NO"),
        "Gross Claim": np.where(claim, sa, 0.0),
        "Retained Claim": np.where(claim, np.minimum(sa, retention), 0.0),
        "Ceded Claim": np.where(claim, np.maximum(0.0, sa - retention), 0.0),
    })[LEDGER_COLUMNS]


def check_ledger(ledger, sims, row, label):
    """Stop if the replayed year does not reproduce the first-pass totals."""
    for ledger_total, column in [(ledger["Gross Claim"].sum(), "Gross"),
                                 (ledger["Retained Claim"].sum(), "Retained"),
                                 (ledger["Ceded Claim"].sum(), "Ceded"),
                                 ((ledger["Claim"] == "YES").sum(), "Claims")]:
        if not np.isclose(ledger_total, sims[column].iloc[row], rtol=1e-9, atol=1e-6):
            raise SystemExit(f"\nSTOPPED: the replayed {label} year does not match the simulation "
                             f"({column}: {ledger_total:,.2f} vs {sims[column].iloc[row]:,.2f}).")


def write_ledger(writer, name, table, formats):
    table.to_excel(writer, sheet_name=name, index=False, header=False, startrow=1)
    sheet = writer.sheets[name]
    for c, heading in enumerate(table.columns):
        sheet.write(0, c, heading, formats["header"])
    column_format = [formats["plain"], formats["money"], formats["money"], formats["qx"],
                     formats["centre"], formats["money"], formats["money"], formats["money"]]
    widths = [18, 8, 16, 12, 10, 18, 18, 18]
    for c, (width, fmt) in enumerate(zip(widths, column_format)):
        sheet.set_column(c, c, width, fmt)
    if len(table) > 0:
        sheet.conditional_format(1, 4, len(table), 4, {"type": "cell", "criteria": "==",
                                 "value": '"YES"', "format": formats["yes"]})
        sheet.conditional_format(1, 4, len(table), 4, {"type": "cell", "criteria": "==",
                                 "value": '"NO"', "format": formats["no"]})
    sheet.freeze_panes(1, 0)


def export_granularity(path, policies, sims, retention, seed, rep_row, extreme_row):
    """Save the four ledger tabs. Returns the two ledgers keyed by simulation row."""
    died = replay_deaths(policies["qx"].to_numpy(), seed, [rep_row, extreme_row])
    ledgers = {}
    for row, label in [(rep_row, "representative"), (extreme_row, "VaR 99%")]:
        ledgers[row] = build_ledger(policies, died[row], retention)
        check_ledger(ledgers[row], sims, row, label)

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        arial = {"font_name": "Arial"}
        formats = {
            "header": book.add_format({**arial, "bold": True, "font_size": 10, "font_color": "#FFFFFF",
                                       "bg_color": "#1F2A44", "border": 1}),
            "plain": book.add_format(arial),
            "money": book.add_format({**arial, "num_format": "#,##0"}),
            "qx": book.add_format({**arial, "num_format": "0.000000"}),
            "centre": book.add_format({**arial, "align": "center"}),
            "yes": book.add_format({"bold": True, "font_color": "#B23A48"}),
            "no": book.add_format({"font_color": "#6B7280"}),
        }
        for prefix, row in [("Rep", rep_row), ("Extrm", extreme_row)]:
            number = int(sims["Simulation"].iloc[row])
            ledger = ledgers[row]
            claimants = ledger[ledger["Claim"] == "YES"].sort_values("SumAssured", ascending=False,
                                                                     kind="mergesort")
            write_ledger(writer, f"{prefix}_Sim{number}_Claimants", claimants, formats)
            write_ledger(writer, f"{prefix}_Sim{number}_All_Lives", ledger, formats)
    return ledgers


# ---------------------------------------------------------------- main
def main():
    settings = ask_settings()
    started = time.time()
    retention = settings["retention"]

    print("\nReading and checking input...")
    policies, mortality = read_inputs(settings["input_file"])
    policies = check_inputs(policies, mortality)
    print(f"  {len(policies):,} lives, total sum assured {policies['SumAssured'].sum():,.0f}")
    for i, (_, divisor) in enumerate(MORTALITY_BANDS):
        in_band = int((policies["qx_divisor"] == divisor).sum())
        print(f"  qx divided by {divisor:g} for {in_band:,} lives with sum assured {band_label(i)}")

    print("Simulating...")
    sims = simulate(policies["qx"].to_numpy(), policies["SumAssured"].to_numpy(dtype=float),
                    retention, settings["simulations"], settings["seed"])
    exact = exact_values(policies, retention)
    summary, pct, frequency, var_sims, var_rows = calculate_metrics(sims, exact, settings["percentiles"])

    if settings["granularity"]:
        rep_row = representative_simulation(sims, exact)
        extreme_row = int(var_rows[0])     # lowest simulation number if several years tie at the VaR value
        rep_id = int(sims["Simulation"].iloc[rep_row])
        extreme_id = int(sims["Simulation"].iloc[extreme_row])
    else:
        rep_id = extreme_id = "not exported (granularity off)"

    run_info = pd.DataFrame([
        ("Input file", Path(settings["input_file"]).name),
        ("Portfolio fingerprint", portfolio_fingerprint(policies)),
        ("Number of lives", len(policies)),
        ("Total sum assured", policies["SumAssured"].sum()),
        *[(f"qx divisor, sum assured {band_label(i)}",
           f"{divisor:g}  ({int((policies['qx_divisor'] == divisor).sum()):,} lives)")
          for i, (_, divisor) in enumerate(MORTALITY_BANDS)],
        ("Expected deaths before adjustment (table qx)", policies["qx_table"].sum()),
        ("Expected deaths (sum of adjusted qx)", policies["qx"].sum()),
        ("Expected claims (sum of qx, sum assured > 0)", exact["Claims"]),
        ("Retention limit", retention),
        ("Lives above retention", int((policies["SumAssured"] > retention).sum())),
        *[(f"Lives in band, {label.lower()}", int(mask.sum()))
          for label, mask in claim_band_masks(policies["SumAssured"]).items()],
        ("Number of simulations", settings["simulations"]),
        ("Random seed", settings["seed"]),
        ("Percentiles", " ".join(f"{p:g}" for p in settings["percentiles"])),
        ("Tail level for VaR / CTE", f"{TAIL_LEVEL}%"),
        ("VaR / CTE years ranked by", TAIL_RANK_BY),
        ("VaR scenario (simulation numbers)", var_sims["amounts"]),
        ("Frequency VaR / CTE years ranked by", FREQUENCY_RANK_BY),
        ("Frequency VaR scenario (simulation numbers)", var_sims["frequency"]),
        ("Representative simulation (granularity)", rep_id),
        ("Extreme simulation (granularity, VaR year)", extreme_id),
        ("Run date", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Run time (seconds)", round(time.time() - started, 1)),
    ], columns=["Setting", "Value"])

    OUTPUT_FOLDER.mkdir(exist_ok=True)
    tag = file_tag(retention, settings["simulations"], settings["seed"])
    path = OUTPUT_FOLDER / f"results_{tag}.xlsx"
    save_results(path, run_info, summary, pct, frequency, sims)

    print("\n" + summary.to_string(index=False, float_format=lambda v: f"{v:,.0f}"))
    print("\n" + frequency.head(len(summary)).drop(columns="Scenario (simulation numbers)").to_string(index=False, float_format=lambda v: f"{v:,.2f}"))
    print(f"\nSaved: {path}")

    if not settings["granularity"]:
        return

    print("\nExporting policy-level granularity...")
    granularity_path = OUTPUT_FOLDER / f"granularity_{tag}.xlsx"
    export_granularity(granularity_path, policies, sims, retention, settings["seed"], rep_row, extreme_row)
    for label, number, row in [("Representative Year", rep_id, rep_row),
                               (f"Extreme Year (VaR {TAIL_LEVEL}%)", extreme_id, extreme_row)]:
        print(f"  {label}: Simulation #{number} "
              f"(claims: {sims['Claims'].iloc[row]:,.0f}, gross: {sims['Gross'].iloc[row]:,.0f} SAR)")
    print(f"  Saved granularity workbook: {granularity_path}")

    check = pd.DataFrame(
        [["Exact actuarial expected value", exact["Claims"], exact["Gross"], exact["Retained"], exact["Ceded"]],
         [f"Representative simulation #{rep_id}"] + [sims[c].iloc[rep_row] for c in ["Claims", "Gross", "Retained", "Ceded"]],
         [f"Extreme VaR {TAIL_LEVEL}% simulation #{extreme_id}"] + [sims[c].iloc[extreme_row] for c in ["Claims", "Gross", "Retained", "Ceded"]]],
        columns=["", "Claims", "Gross", "Retained", "Ceded"])
    print("\n" + check.to_string(index=False, float_format=lambda v: f"{v:,.0f}"))


if __name__ == "__main__":
    main()
