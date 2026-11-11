"""
Tests for simulate.py. Run from the project folder with:  pytest -v

Two tests are written out as models. The rest are stubs for you to complete.
Writing them yourself is the point: each one forces you to state what the
code is supposed to guarantee, which is what an interviewer will ask.
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import simulate as sim


@pytest.fixture
def small_portfolio():
    """A tiny portfolio with known answers: one nil-cover life, one life above retention."""
    return pd.DataFrame({
        "PolicyID": ["A", "B", "C", "D"],
        "Age": [30, 40, 50, 60],
        "SumAssured": [0, 1_000_000, 5_000_000, 20_000_000],
        "qx": [0.01, 0.02, 0.05, 0.10],
    })


RETENTION = 3_400_000


# ------------------------------------------------------------- written out
def test_retained_plus_ceded_equals_gross(small_portfolio):
    """Every simulated year must split exactly into retained and ceded."""
    sims = sim.simulate(small_portfolio["qx"].to_numpy(),
                        small_portfolio["SumAssured"].to_numpy(dtype=float),
                        RETENTION, simulations=2_000, seed=1)
    assert np.allclose(sims["Retained"] + sims["Ceded"], sims["Gross"])


def test_zero_sum_assured_is_never_a_claim(small_portfolio):
    """Life A can die but must never add to the claim count (the bug you flagged as critical)."""
    only_nil = small_portfolio[small_portfolio["SumAssured"] == 0]
    sims = sim.simulate(only_nil["qx"].to_numpy(),
                        only_nil["SumAssured"].to_numpy(dtype=float),
                        RETENTION, simulations=2_000, seed=1)
    assert sims["Claims"].sum() == 0
    assert sims["Gross"].sum() == 0


# ------------------------------------------------------------- yours to write
@pytest.mark.skip(reason="TODO")
def test_simulated_mean_matches_exact_expected_value(small_portfolio):
    """
    Run many simulations, compare sims["Gross"].mean() with sim.exact_values(...)["Gross"].
    Decide the tolerance from the standard error, e.g. within 4 standard errors.
    Why 4 and not 1? Be ready to explain.
    """


@pytest.mark.skip(reason="TODO")
def test_same_seed_gives_identical_results(small_portfolio):
    """Two runs with the same seed must match exactly. A different seed must not."""


@pytest.mark.skip(reason="TODO")
def test_cte_is_at_least_var(small_portfolio):
    """Use sim.var_scenarios and sim.cte_weights. CTE >= VaR for Gross, Retained and Ceded."""


@pytest.mark.skip(reason="TODO")
def test_tied_years_do_not_depend_on_row_order():
    """
    Build a sims table with many tied claim counts, shuffle its rows, and check that
    VaR and CTE (via sim.metrics_table) come out the same. This protects the tie handling.
    """


@pytest.mark.skip(reason="TODO")
def test_replayed_year_matches_first_pass(small_portfolio):
    """
    Use sim.replay_deaths and sim.build_ledger for one simulation row, then check the ledger
    totals equal that row of the simulate() output. This is what sim.check_ledger does at runtime.
    """


@pytest.mark.skip(reason="TODO")
def test_check_inputs_stops_on_bad_data():
    """Duplicate PolicyID, negative sum assured, or an age missing from Mortality should raise SystemExit."""
