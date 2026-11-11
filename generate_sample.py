"""
generate_sample.py: builds a synthetic input workbook for simulate.py.

Policies are invented, so no real policyholder data is used. The Mortality
tab is copied from mortality.csv, the SOA recommended mortality table for
Pakistan based on the SLIC dataset.

Output: sample_input.xlsx with tabs "Policies" and "Mortality".
Usage:  python generate_sample.py [lives] [seed]
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd

LIVES = 50_000
SEED = 2026

MORTALITY_FILE = Path(__file__).with_name("mortality.csv")


def load_mortality():
    """SOA recommended mortality for Pakistan (SLIC based), columns Age and qx."""
    if not MORTALITY_FILE.exists():
        raise SystemExit(f"Missing {MORTALITY_FILE.name}: put the Age, qx table next to this script.")
    return pd.read_csv(MORTALITY_FILE)


def build_policies(lives, rng):
    # Working ages 20 to 64, centred on the late 30s (group-life style portfolio)
    ages = np.clip(rng.normal(38, 9, lives).round(), 20, 64).astype(int)

    # Sum assured: lognormal body, so most lives are small and a few are very large.
    # Rounded to the nearest 50,000. A small share has nil cover to exercise that case.
    sum_assured = np.round(rng.lognormal(mean=14.2, sigma=1.0, size=lives) / 50_000) * 50_000
    sum_assured = np.clip(sum_assured, 50_000, 100_000_000)
    sum_assured[rng.random(lives) < 0.01] = 0

    return pd.DataFrame({"PolicyID": [f"SYN{i:07d}" for i in range(1, lives + 1)],
                         "Age": ages,
                         "SumAssured": sum_assured.astype(np.int64)})


def main():
    lives = int(sys.argv[1]) if len(sys.argv) > 1 else LIVES
    seed = int(sys.argv[2]) if len(sys.argv) > 2 else SEED
    rng = np.random.default_rng(seed)
    with pd.ExcelWriter("sample_input.xlsx", engine="xlsxwriter") as writer:
        build_policies(lives, rng).to_excel(writer, sheet_name="Policies", index=False)
        load_mortality().to_excel(writer, sheet_name="Mortality", index=False)
    print(f"Saved sample_input.xlsx ({lives:,} lives, seed {seed})")


if __name__ == "__main__":
    main()
