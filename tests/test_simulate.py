"""Tests for simulate.py. Run from the project folder with: pytest -v."""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import simulate as sim


@pytest.fixture
def small_portfolio():
    """Small portfolio covering nil-cover, below-retention, and above-retention lives."""
    return pd.DataFrame({
        "PolicyID": ["A", "B", "C", "D"],
        "Age": [30, 40, 50, 60],
        "SumAssured": [0, 1_000_000, 5_000_000, 20_000_000],
        "qx": [0.01, 0.02, 0.05, 0.10],
    })


RETENTION = 3_400_000


def _run_sim(portfolio, simulations=2_000, seed=1):
    """Run simulate() on a portfolio DataFrame using the test retention."""
    return sim.simulate(
        portfolio["qx"].to_numpy(),
        portfolio["SumAssured"].to_numpy(dtype=float),
        RETENTION,
        simulations=simulations,
        seed=seed,
    )


def test_retained_plus_ceded_equals_gross(small_portfolio):
    """Every simulated year splits exactly into retained and ceded claims."""
    sims = _run_sim(small_portfolio)
    assert np.allclose(sims["Retained"] + sims["Ceded"], sims["Gross"])


def test_zero_sum_assured_is_never_a_claim(small_portfolio):
    """Lives with zero sum assured never add to claim counts or amounts."""
    sims = _run_sim(small_portfolio[small_portfolio["SumAssured"] == 0])
    assert sims["Claims"].sum() == 0
    assert sims["Gross"].sum() == 0


def test_simulated_mean_matches_exact_expected_value(small_portfolio):
    """Simulated means stay within four standard errors of exact expectations."""
    sims = _run_sim(small_portfolio, simulations=20_000, seed=42)
    exact = sim.exact_values(small_portfolio, RETENTION)
    for col in ("Gross", "Retained", "Ceded", "Claims"):
        values = sims[col].to_numpy()
        standard_error = values.std(ddof=1) / np.sqrt(len(values))
        assert abs(values.mean() - exact[col]) <= 4 * standard_error


def test_same_seed_gives_identical_results(small_portfolio):
    """Matching seeds reproduce identical results while different seeds diverge."""
    run_a = _run_sim(small_portfolio, simulations=1_000, seed=42)
    run_b = _run_sim(small_portfolio, simulations=1_000, seed=42)
    run_c = _run_sim(small_portfolio, simulations=1_000, seed=99)

    pd.testing.assert_frame_equal(run_a, run_b)
    assert not run_a.equals(run_c)


def test_cte_is_at_least_var(small_portfolio):
    """CTE is at least VaR for gross, retained, and ceded claims."""
    sims = _run_sim(small_portfolio)
    for col in ("Gross", "Retained", "Ceded"):
        var_rows = sim.var_scenarios(sims, col)
        weights, k = sim.cte_weights(sims, col)
        var = sims[col].to_numpy()[var_rows].mean()
        cte = (sims[col].to_numpy() * weights).sum() / k
        assert cte >= var


def test_tied_years_do_not_depend_on_row_order():
    """VaR and CTE tie handling is invariant to simulation row ordering."""
    claims = np.array([0] * 100 + [1] * 70 + [2] * 25 + [3] * 4 + [4])
    gross = np.arange(1, 201, dtype=float) * 100_000
    sims = pd.DataFrame({"Claims": claims, "Gross": gross})
    shuffled = sims.sample(frac=1.0, random_state=7).reset_index(drop=True)

    exact = {"Claims": sims["Claims"].mean(), "Gross": sims["Gross"].mean()}
    table_orig, _ = sim.metrics_table(sims, ["Claims", "Gross"], exact, rank_by="Claims")
    table_shuf, _ = sim.metrics_table(shuffled, ["Claims", "Gross"], exact, rank_by="Claims")

    pd.testing.assert_frame_equal(table_orig, table_shuf)


def test_replayed_year_matches_first_pass(small_portfolio):
    """Replaying a simulation row reproduces the first-pass totals."""
    seed = 123
    sims = _run_sim(small_portfolio, simulations=500, seed=seed)

    claim_rows = sims.index[sims["Claims"] > 0]
    assert len(claim_rows) > 0
    row = int(claim_rows[0])

    died = sim.replay_deaths(small_portfolio["qx"].to_numpy(), seed, [row])[row]
    ledger = sim.build_ledger(small_portfolio, died, RETENTION)

    assert (ledger["Claim"] == "YES").sum() == sims["Claims"].iloc[row]
    assert np.isclose(ledger["Gross Claim"].sum(), sims["Gross"].iloc[row])
    assert np.isclose(ledger["Retained Claim"].sum(), sims["Retained"].iloc[row])
    assert np.isclose(ledger["Ceded Claim"].sum(), sims["Ceded"].iloc[row])
    sim.check_ledger(ledger, sims, row, "replayed")


def test_check_inputs_stops_on_bad_data(tmp_path, monkeypatch):
    """Input validation stops on duplicate IDs, negative sums assured, or unknown ages."""
    monkeypatch.setattr(sim, "OUTPUT_FOLDER", tmp_path)
    mortality = pd.DataFrame({"Age": [30, 40], "qx": [0.01, 0.02]})
    valid = pd.DataFrame({
        "PolicyID": ["P1", "P2"],
        "Age": [30, 40],
        "SumAssured": [1_000_000, 2_000_000],
    })

    dup_id = valid.copy()
    dup_id["PolicyID"] = ["P1", "P1"]

    neg_sa = valid.copy()
    neg_sa.loc[0, "SumAssured"] = -500

    missing_age = valid.copy()
    missing_age.loc[1, "Age"] = 99

    for bad_policies, expected_col in ((dup_id, "PolicyID"), (neg_sa, "SumAssured"), (missing_age, "Age")):
        with pytest.raises(SystemExit):
            sim.check_inputs(bad_policies, mortality)
        problems = pd.read_csv(tmp_path / "input_problems.csv")
        assert expected_col in set(problems["Column"])
