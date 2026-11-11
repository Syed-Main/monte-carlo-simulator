"""
Compare simulate.py results across retention limits.

Requires runs to share the same portfolio fingerprint, simulation count, and seed.
Output: output/comparison_R<lowest>-R<highest>_N<sims>_S<seed>.xlsx
"""

from pathlib import Path

import numpy as np
import pandas as pd

OUTPUT_FOLDER = Path("output")
MUST_MATCH = ["Portfolio fingerprint", "Number of simulations", "Random seed"]


def choose_files():
    print("\nRetention comparison\n")
    answer = input(
        "Results files to compare (Enter = all results files in output/, "
        "or type names separated by commas): "
    ).strip()
    if answer:
        files = [Path(name.strip()) for name in answer.split(",")]
    else:
        files = sorted(OUTPUT_FOLDER.glob("results_*.xlsx"))
    missing = [str(f) for f in files if not f.exists()]
    if missing:
        raise SystemExit(f"\nSTOPPED: file(s) not found: {', '.join(missing)}")
    if len(files) < 2:
        raise SystemExit("\nSTOPPED: need at least two results files to compare.")
    return files


def read_run(path):
    sheets = pd.read_excel(path, sheet_name=["Run Info", "Summary", "Frequency", "Simulations"])
    info = dict(zip(sheets["Run Info"]["Setting"], sheets["Run Info"]["Value"]))
    return {
        "file": path.name,
        "info": info,
        "summary": sheets["Summary"].set_index("Metric"),
        "frequency": sheets["Frequency"].set_index("Metric"),
        "retained": sheets["Simulations"]["Retained"].to_numpy(),
    }


def check_runs_match(runs):
    table = pd.DataFrame([{"File": r["file"], **{k: r["info"][k] for k in MUST_MATCH}} for r in runs])
    if any(table[k].astype(str).nunique() > 1 for k in MUST_MATCH):
        print("\nSTOPPED: these runs cannot be compared fairly; they differ in:\n")
        print(table.to_string(index=False))
        raise SystemExit("\nRe-run simulate.py with the same input file, number of simulations and seed.")


def comparison_table(runs):
    rows = []
    for r in runs:
        s = r["summary"]
        var = [m for m in s.index if m.startswith("VaR")][0]
        cte = [m for m in s.index if m.startswith("CTE")][0]
        mean = s.loc["Mean (simulated)"]
        fq = r["frequency"]
        rows.append({
            "Retention limit": float(r["info"]["Retention limit"]),
            "Mean gross": mean["Gross"],
            "Mean retained": mean["Retained"],
            "Mean ceded": mean["Ceded"],
            "Share ceded": mean["Ceded"] / mean["Gross"],
            "Std dev retained": s.at["Standard deviation", "Retained"],
            f"{var} retained": s.at[var, "Retained"],
            f"{cte} retained": s.at[cte, "Retained"],
            "Mean claims": fq.at["Mean (simulated)", "Claims"],
            "Mean ceded claims": fq.at["Mean (simulated)", "Ceded claims"],
            "Share of claims ceded": (
                fq.at["Mean (simulated)", "Ceded claims"] / fq.at["Mean (simulated)", "Claims"]
            ),
            f"{var} ceded claims": fq.at[var, "Ceded claims"],
            f"{cte} ceded claims": fq.at[cte, "Ceded claims"],
        })
    table = pd.DataFrame(rows)
    if table["Retention limit"].duplicated().any():
        print("  Note: some retention limits appear in more than one file; keeping the first of each.")
    keep = ~table["Retention limit"].duplicated()
    sorted_table = table[keep].sort_values("Retention limit").reset_index(drop=True)
    kept_runs = [r for r, k in zip(runs, keep) if k]
    return sorted_table, kept_runs


def save_comparison(path, table, runs):
    order = np.argsort([float(r["info"]["Retention limit"]) for r in runs])
    runs = [runs[i] for i in order]
    probability = np.linspace(0, 100, 201)
    cdf = pd.DataFrame({"Probability": probability / 100})
    for r in runs:
        cdf[f"Retention {float(r['info']['Retention limit']):,.0f}"] = np.percentile(
            r["retained"], probability
        )

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book

        def fmt(**kwargs):
            return book.add_format({"font_name": "Arial", **kwargs})

        head = fmt(bold=True, font_color="white", bg_color="#1F2A44", border=1)
        money = fmt(num_format="#,##0")
        share = fmt(num_format="0.0%")
        decimals = fmt(num_format="#,##0.00")

        table.to_excel(writer, sheet_name="Comparison", index=False)
        ws = writer.sheets["Comparison"]
        for c, name in enumerate(table.columns):
            ws.write(0, c, name, head)
            col_fmt = (
                share if name.startswith("Share")
                else decimals if "claims" in name and "retained" not in name
                else money
            )
            ws.set_column(c, c, 22, col_fmt)
        ws.freeze_panes(1, 1)

        charts = book.add_worksheet("Charts")
        cdf.to_excel(writer, sheet_name="Chart Data", index=False)
        writer.sheets["Chart Data"].set_column(0, 0, 12, share)
        writer.sheets["Chart Data"].set_column(1, len(cdf.columns), 20, money)
        n = len(table)
        last = len(cdf)

        risk = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
        for col in [2, 6, 7]:
            risk.add_series({
                "name": ["Comparison", 0, col],
                "categories": ["Comparison", 1, 0, n, 0],
                "values": ["Comparison", 1, col, n, col],
            })
        risk.set_title({"name": "Retained risk by retention limit"})
        risk.set_x_axis({"name": "Retention limit", "num_format": "#,##0"})
        risk.set_y_axis({"name": "Total retained claims", "num_format": "#,##0"})
        risk.set_legend({"position": "bottom"})
        risk.set_size({"width": 900, "height": 420})
        charts.insert_chart("B2", risk)

        split = book.add_chart({"type": "column", "subtype": "stacked"})
        for col, colour in [(2, "#1E6F72"), (3, "#C08A2E")]:
            split.add_series({
                "name": ["Comparison", 0, col],
                "categories": ["Comparison", 1, 0, n, 0],
                "values": ["Comparison", 1, col, n, col],
                "fill": {"color": colour},
            })
        split.set_title({"name": "Expected claims split: retained vs ceded"})
        split.set_x_axis({"name": "Retention limit", "num_format": "#,##0"})
        split.set_y_axis({"num_format": "#,##0"})
        split.set_legend({"position": "bottom"})
        split.set_size({"width": 900, "height": 420})
        charts.insert_chart("B24", split)

        dist = book.add_chart({"type": "scatter", "subtype": "straight"})
        for col in range(1, len(cdf.columns)):
            dist.add_series({
                "name": ["Chart Data", 0, col],
                "categories": ["Chart Data", 1, col, last, col],
                "values": ["Chart Data", 1, 0, last, 0],
            })
        dist.set_title({"name": "Cumulative distribution of retained claims by retention limit"})
        dist.set_x_axis({"name": "Total retained claims", "num_format": "#,##0"})
        dist.set_y_axis({"name": "Probability", "num_format": "0%", "min": 0, "max": 1})
        dist.set_legend({"position": "bottom"})
        dist.set_size({"width": 900, "height": 460})
        charts.insert_chart("B46", dist)

        col = list(table.columns).index("Mean ceded claims")
        freq = book.add_chart({"type": "column"})
        freq.add_series({
            "name": "Mean ceded claims",
            "categories": ["Comparison", 1, 0, n, 0],
            "values": ["Comparison", 1, col, n, col],
            "fill": {"color": "#C08A2E"},
            "data_labels": {"value": True, "num_format": "#,##0.0"},
        })
        freq.set_title({"name": "Expected number of claims reaching reinsurance"})
        freq.set_x_axis({"name": "Retention limit", "num_format": "#,##0"})
        freq.set_y_axis({"name": "Claims per year"})
        freq.set_legend({"none": True})
        freq.set_size({"width": 900, "height": 380})
        charts.insert_chart("B70", freq)

        info = runs[0]["info"]
        run_meta = [("Files compared", ", ".join(r["file"] for r in runs))] + [
            (k, info[k]) for k in ["Input file"] + MUST_MATCH
        ]
        pd.DataFrame(run_meta, columns=["Setting", "Value"]).to_excel(
            writer, sheet_name="Run Info", index=False
        )
        writer.sheets["Run Info"].set_column(0, 1, 40)


def main():
    runs = [read_run(f) for f in choose_files()]
    check_runs_match(runs)
    table, runs = comparison_table(runs)
    info = runs[0]["info"]
    low, high = table["Retention limit"].min(), table["Retention limit"].max()
    path = OUTPUT_FOLDER / (
        f"comparison_R{low:.0f}-R{high:.0f}"
        f"_N{int(info['Number of simulations'])}_S{int(info['Random seed'])}.xlsx"
    )
    save_comparison(path, table, runs)
    print("\n" + table.to_string(index=False, float_format=lambda v: f"{v:,.1f}" if v > 1 else f"{v:.1%}"))
    print(f"\nSaved: {path}")


if __name__ == "__main__":
    main()
