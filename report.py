"""
report.py: turns a results file from simulate.py into a stakeholder report.

  * Always creates:  output/report_R<retention>_N<sims>_S<seed>.xlsx  (native Excel charts)
  * Optionally:      output/report_R<retention>_N<sims>_S<seed>.html  (interactive charts)

To get a PDF, open the HTML report in a browser and choose Print > Save as PDF.
"""

import html
from pathlib import Path

import numpy as np
import pandas as pd

OUTPUT_FOLDER = Path("output")
HISTOGRAM_BINS = 40
COLOURS = {"Gross": "#6B7280", "Retained": "#1E6F72", "Ceded": "#C08A2E",
           "Claims": "#1E6F72", "Ceded claims": "#C08A2E",
           "tail": "#B23A48", "ink": "#1F2A44"}


# ---------------------------------------------------------------- read
def ask(question, default):
    return input(f"{question} [{default}]: ").strip() or default


def newest_results_file():
    files = sorted(OUTPUT_FOLDER.glob("results_*.xlsx"), key=lambda f: f.stat().st_mtime)
    return str(files[-1]) if files else ""


def read_results(path):
    if not Path(path).exists():
        raise SystemExit(f"\nSTOPPED: results file not found: {path}")
    sheets = pd.read_excel(path, sheet_name=None)
    info = dict(zip(sheets["Run Info"]["Setting"], sheets["Run Info"]["Value"]))
    summary = sheets["Summary"].set_index("Metric")
    frequency = sheets["Frequency"].set_index("Metric")
    return info, summary, sheets["Percentiles"], frequency, sheets["Simulations"]


def file_tag(info):
    return f"R{float(info['Retention limit']):.0f}_N{int(info['Number of simulations'])}_S{int(info['Random seed'])}"


def headline(info, summary):
    """The one-sentence takeaway that opens both reports."""
    var = [m for m in summary.index if m.startswith("VaR")][0]
    return (f"With a retention of {float(info['Retention limit']):,.0f}, expected retained claims are "
            f"{summary.at['Mean (simulated)', 'Retained']:,.0f}, and in a 1-in-100 year they reach "
            f"{summary.at[var, 'Retained']:,.0f} (gross: {summary.at[var, 'Gross']:,.0f}).")


# ---------------------------------------------------------------- chart data
def histogram_table(retained, var):
    counts, edges = np.histogram(retained, bins=HISTOGRAM_BINS)
    middle = (edges[:-1] + edges[1:]) / 2
    beyond = edges[:-1] >= var                       # bins that sit in the tail
    return pd.DataFrame({"Retained claims (bin middle)": middle,
                         "Simulations below VaR": np.where(beyond, 0, counts),
                         "Simulations beyond VaR": np.where(beyond, counts, 0)})


def count_table(sims, column):
    """How many simulations had each possible number of claims."""
    counts = sims[column].astype(int).value_counts().sort_index()
    full = range(counts.index.min(), counts.index.max() + 1)
    return counts.reindex(full, fill_value=0).rename_axis(column).rename("Simulations").reset_index()


def cdf_table(sims):
    probability = np.linspace(0, 100, 201)
    return pd.DataFrame({"Probability": probability / 100,
                         "Gross": np.percentile(sims["Gross"], probability),
                         "Retained": np.percentile(sims["Retained"], probability)})


# ---------------------------------------------------------------- Excel report
def excel_report(path, info, summary, pct, frequency, sims):
    var_name = [m for m in summary.index if m.startswith("VaR")][0]
    cte_name = [m for m in summary.index if m.startswith("CTE")][0]
    hist = histogram_table(sims["Retained"], summary.at[var_name, "Retained"])
    cdf = cdf_table(sims)
    means = summary.loc["Mean (simulated)", ["Gross", "Retained", "Ceded"]]

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        f = lambda **k: book.add_format({"font_name": "Arial", **k})
        title, sub, head = f(bold=True, font_size=16, font_color=COLOURS["ink"]), f(italic=True, font_color="#555555"), \
            f(bold=True, font_color="white", bg_color=COLOURS["ink"], border=1)
        label, money, pctf, text = f(bold=True), f(num_format="#,##0"), f(num_format="0%"), f()
        big = f(bold=True, font_size=13, num_format="#,##0", font_color=COLOURS["ink"])

        # --- Report sheet
        ws = book.add_worksheet("Report")
        ws.set_column(0, 0, 34); ws.set_column(1, 3, 18)
        ws.write(0, 0, f"Claims & Reinsurance Report: retention {float(info['Retention limit']):,.0f}", title)
        ws.write(1, 0, f"{int(info['Number of lives']):,} lives, {int(info['Number of simulations']):,} simulations, "
                       f"seed {int(info['Random seed'])}, run {info['Run date']}", sub)
        ws.write(2, 0, headline(info, summary), text)

        ws.write(4, 0, "Key figures (retained)", head); ws.write(4, 1, "", head)
        for r, (name, value) in enumerate([("Expected retained claims", means["Retained"]),
                                           (f"{var_name} (1-in-100 year)", summary.at[var_name, "Retained"]),
                                           (f"{cte_name} (average of worst 1%)", summary.at[cte_name, "Retained"]),
                                           ("Expected reinsurance recoveries", means["Ceded"])], start=5):
            ws.write(r, 0, name, label); ws.write(r, 1, value, big)

        start = 11
        ws.write_row(start, 0, ["Metric", "Gross", "Retained", "Ceded"], head)
        for r, (metric, row) in enumerate(summary.iterrows(), start=start + 1):
            ws.write(r, 0, metric, label); ws.write_row(r, 1, row.tolist(), money)

        start += len(summary) + 3
        ws.write_row(start, 0, ["Percentile", "Gross", "Retained", "Ceded"], head)
        for r, row in enumerate(pct.itertuples(index=False), start=start + 1):
            ws.write(r, 0, row[0], label); ws.write_row(r, 1, list(row[1:]), money)

        ch = book.add_worksheet("Charts")     # created first so it sits before the data tab

        # --- Frequency sheet: number of claims
        fq = book.add_worksheet("Frequency")
        decimals = f(num_format="#,##0.00")
        fq.set_column(0, 0, 34); fq.set_column(1, 2, 18)
        fq.write(0, 0, "Claim frequency: number of claims per year", title)
        fq.write(1, 0, "Ceded claims = deaths where the sum assured exceeds the retention limit.", sub)
        fq.write_row(3, 0, ["Metric"] + list(frequency.columns), head)
        for r, (metric, row) in enumerate(frequency.iterrows(), start=4):
            fq.write(r, 0, metric, label); fq.write_row(r, 1, row.tolist(), decimals)

        # --- Chart data sheet (charts read from here)
        hist.to_excel(writer, sheet_name="Chart Data", index=False, startcol=0)
        cdf.to_excel(writer, sheet_name="Chart Data", index=False, startcol=4)
        means.rename("Mean").rename_axis("Claim type").reset_index() \
            .to_excel(writer, sheet_name="Chart Data", index=False, startcol=8)
        counts = {col: count_table(sims, col) for col in frequency.columns}
        for i, col in enumerate(frequency.columns):
            counts[col].to_excel(writer, sheet_name="Chart Data", index=False, startcol=11 + 3 * i)
        cd = writer.sheets["Chart Data"]
        cd.set_column(0, 16, 18, money); cd.set_column(4, 4, 12, pctf)
        n_hist, n_cdf = len(hist), len(cdf)

        # --- Frequency charts (on the Frequency sheet)
        for i, col in enumerate(frequency.columns):
            c, n = 11 + 3 * i, len(counts[col])
            chart = book.add_chart({"type": "column"})
            chart.add_series({"name": col, "categories": ["Chart Data", 1, c, n, c],
                              "values": ["Chart Data", 1, c + 1, n, c + 1],
                              "fill": {"color": COLOURS[col]}, "gap": 20})
            chart.set_title({"name": f"Distribution of number of {col.lower()} per year"})
            chart.set_x_axis({"name": f"Number of {col.lower()}", "label_position": "low"})
            chart.set_y_axis({"name": "Number of simulations"}); chart.set_legend({"none": True})
            chart.set_size({"width": 800, "height": 360})
            fq.insert_chart(4 + len(frequency) + 2 + 19 * i, 0, chart)

        # --- Charts sheet
        hist_chart = book.add_chart({"type": "column", "subtype": "stacked"})
        for col, name, colour in [(1, "Below VaR", COLOURS["Retained"]), (2, "Beyond VaR (tail)", COLOURS["tail"])]:
            hist_chart.add_series({"name": name, "categories": ["Chart Data", 1, 0, n_hist, 0],
                                   "values": ["Chart Data", 1, col, n_hist, col],
                                   "fill": {"color": colour}, "gap": 10})
        hist_chart.set_title({"name": "Distribution of total retained claims"})
        hist_chart.set_x_axis({"name": "Total retained claims", "num_format": "#,##0", "label_position": "low"})
        hist_chart.set_y_axis({"name": "Number of simulations"})
        hist_chart.set_legend({"position": "bottom"}); hist_chart.set_size({"width": 900, "height": 420})
        ch.insert_chart("B2", hist_chart)

        cdf_chart = book.add_chart({"type": "scatter", "subtype": "straight"})
        for col, name in [(5, "Gross"), (6, "Retained")]:
            cdf_chart.add_series({"name": name, "categories": ["Chart Data", 1, col, n_cdf, col],
                                  "values": ["Chart Data", 1, 4, n_cdf, 4],
                                  "line": {"color": COLOURS[name], "width": 2.25}})
        cdf_chart.set_title({"name": "Cumulative distribution: chance total claims are at or below an amount"})
        cdf_chart.set_x_axis({"name": "Total claims", "num_format": "#,##0"})
        cdf_chart.set_y_axis({"name": "Probability", "num_format": "0%", "min": 0, "max": 1})
        cdf_chart.set_legend({"position": "bottom"}); cdf_chart.set_size({"width": 900, "height": 420})
        ch.insert_chart("B24", cdf_chart)

        bar = book.add_chart({"type": "column"})
        bar.add_series({"name": "Expected claims", "categories": ["Chart Data", 1, 8, 3, 8],
                        "values": ["Chart Data", 1, 9, 3, 9], "data_labels": {"value": True, "num_format": "#,##0"},
                        "points": [{"fill": {"color": COLOURS[k]}} for k in ["Gross", "Retained", "Ceded"]]})
        bar.set_title({"name": "Expected claims: gross, retained and ceded"})
        bar.set_y_axis({"num_format": "#,##0"}); bar.set_legend({"none": True})
        bar.set_size({"width": 900, "height": 380})
        ch.insert_chart("B46", bar)

        # --- Run info sheet
        pd.DataFrame(info.items(), columns=["Setting", "Value"]).to_excel(writer, sheet_name="Run Info", index=False)
        writer.sheets["Run Info"].set_column(0, 1, 32, text)


# ---------------------------------------------------------------- HTML report
def html_table(df, money_cols, decimals=0):
    out = df.copy()
    for c in money_cols:
        out[c] = out[c].map(lambda v: f"{v:,.{decimals}f}")
    return out.to_html(index=False, border=0, classes="data", escape=True)


def html_report(path, info, summary, pct, frequency, sims):
    import plotly.graph_objects as go

    var_name = [m for m in summary.index if m.startswith("VaR")][0]
    cte_name = [m for m in summary.index if m.startswith("CTE")][0]
    var, cte = summary.at[var_name, "Retained"], summary.at[cte_name, "Retained"]
    layout = dict(template="simple_white", font=dict(family="Segoe UI, Arial, sans-serif", size=13),
                  margin=dict(l=60, r=20, t=50, b=50), height=420, separators=".,",
                  legend=dict(orientation="h", y=-0.2))

    hist = go.Figure(go.Histogram(x=sims["Retained"], nbinsx=60, marker_color=COLOURS["Retained"],
                                  hovertemplate="%{x}<br>%{y} simulations<extra></extra>"))
    hist.update_layout(title="Distribution of total retained claims", xaxis_title="Total retained claims",
                       yaxis_title="Number of simulations", xaxis_tickformat=",.0f", **layout)
    hist.add_vline(x=var, line_color=COLOURS["tail"], line_dash="dash", line_width=2, opacity=1,
                   annotation_text=f"{var_name}: {var:,.0f}", annotation_position="top left")
    hist.add_vline(x=cte, line_color=COLOURS["tail"], line_width=2, opacity=1,
                   annotation_text=f"{cte_name}: {cte:,.0f}", annotation_position="top right")

    cdf = cdf_table(sims)
    cdf_fig = go.Figure([go.Scatter(x=cdf[k], y=cdf["Probability"], name=k, line=dict(color=COLOURS[k], width=2.5),
                                    hovertemplate="%{x:,.0f}<br>%{y:.1%} of simulations at or below<extra>" + k + "</extra>")
                         for k in ["Gross", "Retained"]])
    cdf_fig.update_layout(title="Cumulative distribution", xaxis_title="Total claims", yaxis_title="Probability",
                          xaxis_tickformat=",.0f", yaxis_tickformat=".0%", **layout)

    means = summary.loc["Mean (simulated)", ["Gross", "Retained", "Ceded"]]
    bar = go.Figure(go.Bar(x=means.index, y=means.values, marker_color=[COLOURS[k] for k in means.index],
                           text=[f"{v:,.0f}" for v in means.values], textposition="outside",
                           hovertemplate="%{x}: %{y:,.0f}<extra></extra>"))
    bar.update_layout(title="Expected claims", yaxis_tickformat=",.0f", **{**layout, "height": 380})

    freq_figs = []
    for col in frequency.columns:
        counts = count_table(sims, col)
        fig = go.Figure(go.Bar(x=counts[col], y=counts["Simulations"], marker_color=COLOURS[col],
                               hovertemplate="%{x} " + col.lower() + "<br>%{y} simulations<extra></extra>"))
        fig.update_layout(title=f"Distribution of number of {col.lower()} per year",
                          xaxis_title=f"Number of {col.lower()}", yaxis_title="Number of simulations",
                          bargap=0.1, **{**layout, "height": 380})
        f_var, f_cte = frequency.at[var_name, col], frequency.at[cte_name, col]
        var_side, cte_side = ("left", "right") if f_var <= f_cte else ("right", "left")   # keep labels apart
        fig.add_vline(x=f_var, line_color=COLOURS["tail"], line_dash="dash", line_width=2, opacity=1,
                      annotation_text=f"{var_name}: {f_var:,.0f}", annotation_position=f"top {var_side}")
        fig.add_vline(x=f_cte, line_color=COLOURS["tail"], line_width=2, opacity=1,
                      annotation_text=f"{cte_name}: {f_cte:,.1f}", annotation_position=f"top {cte_side}")
        freq_figs.append(fig)

    config = {"displaylogo": False}
    charts = [hist.to_html(full_html=False, include_plotlyjs=True, config=config)] + \
             [f.to_html(full_html=False, include_plotlyjs=False, config=config) for f in (cdf_fig, bar, *freq_figs)]

    key = [("Expected retained claims", means["Retained"]), (f"{var_name}, 1-in-100 year", var),
           (f"{cte_name}, average of worst 1%", cte), ("Expected reinsurance recoveries", means["Ceded"])]
    key_html = "".join(f"<div><dt>{html.escape(k)}</dt><dd>{v:,.0f}</dd></div>" for k, v in key)
    run_info = pd.DataFrame(info.items(), columns=["Setting", "Value"]).astype(str)

    page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Claims &amp; Reinsurance Report: retention {float(info['Retention limit']):,.0f}</title>
<style>
  :root {{ --ink:{COLOURS['ink']}; --muted:#5B6475; --rule:#D9DEE7; --retained:{COLOURS['Retained']}; }}
  body {{ margin:0; background:#fff; color:var(--ink); font:15px/1.55 "Segoe UI", Arial, sans-serif; }}
  main {{ max-width:960px; margin:0 auto; padding:40px 24px 64px; }}
  h1 {{ font:600 30px/1.2 Georgia, "Times New Roman", serif; margin:0 0 6px; }}
  h2 {{ font:600 20px/1.3 Georgia, "Times New Roman", serif; margin:44px 0 10px; }}
  .meta {{ color:var(--muted); margin:0 0 22px; }}
  .lead {{ font:400 21px/1.45 Georgia, "Times New Roman", serif; max-width:46em;
           border-left:4px solid var(--retained); padding-left:16px; margin:0 0 28px; }}
  dl {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:0; margin:0;
        border-top:2px solid var(--ink); }}
  dl div {{ padding:12px 14px 12px 0; border-bottom:1px solid var(--rule); }}
  dt {{ color:var(--muted); font-size:13px; }}
  dd {{ margin:2px 0 0; font-size:22px; font-weight:600; font-variant-numeric:tabular-nums; }}
  .chart {{ margin:18px 0; break-inside:avoid; }}
  .scroll {{ overflow-x:auto; }}
  table.data {{ border-collapse:collapse; width:100%; font-variant-numeric:tabular-nums; }}
  table.data th {{ text-align:right; border-bottom:2px solid var(--ink); padding:6px 10px; font-weight:600; }}
  table.data td {{ text-align:right; border-bottom:1px solid var(--rule); padding:6px 10px; }}
  table.data th:first-child, table.data td:first-child {{ text-align:left; }}
  .note {{ background:#EEF4F4; padding:10px 14px; color:var(--muted); font-size:14px; }}
  @media print {{ .no-print {{ display:none; }} main {{ padding:0; }} h2 {{ break-after:avoid; }} }}
</style></head>
<body><main>
  <p class="note no-print">To save as PDF: press Ctrl+P (Cmd+P on Mac) and choose "Save as PDF". Hover over charts to see exact values.</p>
  <h1>Claims &amp; reinsurance report</h1>
  <p class="meta">Retention {float(info['Retention limit']):,.0f}, {int(info['Number of lives']):,} lives,
     {int(info['Number of simulations']):,} simulations, seed {int(info['Random seed'])}, run {html.escape(str(info['Run date']))}</p>
  <p class="lead">{html.escape(headline(info, summary))}</p>
  <dl>{key_html}</dl>
  <h2>Loss distribution</h2>
  <div class="chart">{charts[0]}</div>
  <div class="chart">{charts[1]}</div>
  <div class="chart">{charts[2]}</div>
  <h2>Summary</h2>
  <div class="scroll">{html_table(summary.reset_index(), ["Gross", "Retained", "Ceded"])}</div>
  <h2>Percentiles</h2>
  <div class="scroll">{html_table(pct, ["Gross", "Retained", "Ceded"])}</div>
  <h2>Claim frequency</h2>
  <p>Number of claims per year. Ceded claims are deaths where the sum assured exceeds the retention limit.</p>
  <div class="chart">{charts[3]}</div>
  <div class="chart">{charts[4]}</div>
  <div class="scroll">{html_table(frequency.reset_index(), list(frequency.columns), decimals=2)}</div>
  <h2>Run details</h2>
  <div class="scroll">{run_info.to_html(index=False, border=0, classes="data")}</div>
</main></body></html>"""
    Path(path).write_text(page, encoding="utf-8")


# ---------------------------------------------------------------- main
def main():
    print("\nReport builder\n")
    source = ask("Results file", newest_results_file())
    make_html = ask("Also create interactive HTML report? (y/n)", "n").lower().startswith("y")

    info, summary, pct, frequency, sims = read_results(source)
    name = OUTPUT_FOLDER / f"report_{file_tag(info)}"

    excel_report(f"{name}.xlsx", info, summary, pct, frequency, sims)
    print(f"Saved: {name}.xlsx")
    if make_html:
        html_report(f"{name}.html", info, summary, pct, frequency, sims)
        print(f"Saved: {name}.html  (open in a browser; Print > Save as PDF for a PDF)")


if __name__ == "__main__":
    main()
