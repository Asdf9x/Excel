import os
import numpy as np
import pandas as pd

OUTPUT_PATH = "reports/NVDA_Returns_Analysis.xlsx"
DATA_PATH = "data/nvda_us_m.csv"

def main():
    # 1) Load data
    df = pd.read_csv(DATA_PATH, parse_dates=["Date"])
    df = df.sort_values("Date").reset_index(drop=True)

    # 2) Compute simple monthly returns from Close
    # Return_t = Close_t / Close_{t-1} - 1
    df["Return"] = df["Close"].pct_change()

    # Prepare Returns series (drop the first NaN)
    returns = df.loc[1:, ["Date", "Return"]].reset_index(drop=True)

    # 3) Descriptive statistics (based on returns)
    r = returns["Return"].dropna().to_numpy()
    n = r.size
    mean = float(np.mean(r)) if n else np.nan
    median = float(np.median(r)) if n else np.nan
    # mode: may be empty; take first if exists
    mode_vals = pd.Series(r).mode()
    mode = float(mode_vals.iloc[0]) if len(mode_vals) > 0 else np.nan
    r_min = float(np.min(r)) if n else np.nan
    r_max = float(np.max(r)) if n else np.nan
    r_range = r_max - r_min if n else np.nan
    var_s = float(np.var(r, ddof=1)) if n > 1 else np.nan
    std_s = float(np.std(r, ddof=1)) if n > 1 else np.nan

    p20 = float(np.percentile(r, 20)) if n else np.nan
    p60 = float(np.percentile(r, 60)) if n else np.nan
    p90 = float(np.percentile(r, 90)) if n else np.nan

    q1 = float(np.percentile(r, 25)) if n else np.nan
    q3 = float(np.percentile(r, 75)) if n else np.nan
    iqr = q3 - q1 if n else np.nan
    lower_bound = q1 - 1.5 * iqr if n else np.nan
    upper_bound = q3 + 1.5 * iqr if n else np.nan

    # Outliers by IQR rule
    outliers_df = pd.DataFrame(columns=["Date", "Return"])
    if n:
        mask_out = (returns["Return"] < lower_bound) | (returns["Return"] > upper_bound)
        outliers_df = returns.loc[mask_out, ["Date", "Return"]].copy().reset_index(drop=True)

    # 4) Frequency table with 10% class interval
    bin_width = 0.10  # 10%
    if n:
        lower = np.floor(r.min() / bin_width) * bin_width
        upper = np.ceil(r.max() / bin_width) * bin_width
        # ensure upper >= lower
        if upper == lower:
            upper = lower + bin_width
        edges = np.arange(lower, upper + bin_width * 1.0001, bin_width)
        counts, edges = np.histogram(r, bins=edges)
        labels = [f"{edges[i]:.0%} to {edges[i+1]:.0%}" for i in range(len(edges) - 1)]
        rel = counts / counts.sum() if counts.sum() > 0 else np.zeros_like(counts, dtype=float)
        freq_df = pd.DataFrame({
            "Class Interval": labels,
            "Frequency": counts,
            "Relative Frequency": rel
        })
    else:
        freq_df = pd.DataFrame(columns=["Class Interval", "Frequency", "Relative Frequency"])

    # 5) Write Excel with embedded histogram chart
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book
        # Formats
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        pct1_fmt = wb.add_format({"num_format": "0.0%"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        num_fmt = wb.add_format({"num_format": "0.0000"})
        bold = wb.add_format({"bold": True})

        # Sheet: Data (original)
        df.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        # Format date column
        ws_data.set_column("A:A", 12, date_fmt)
        ws_data.set_column("B:F", 12)

        # Sheet: Returns (use formula for Return so it stays dynamic if Data changes)
        ws_ret_name = "Returns"
        ws_ret = wb.add_worksheet(ws_ret_name)
        ws_ret.write(0, 0, "Date", bold)
        ws_ret.write(0, 1, "Return (Simple %)", bold)

        # Fill dates (starting from second row of Data) and formulas for returns
        # Data headers at row 1; first data at row 2 in Excel
        n_rows = len(df)
        ret_row = 1
        for i in range(1, n_rows):
            data_row_excel = i + 1  # 0-based i -> Excel row number (headers at 1)
            # Write date from Data sheet via formula to keep consistent
            # Or write the value directly; here we write the value for better readability
            ws_ret.write_datetime(ret_row, 0, df.loc[i, "Date"].to_pydatetime(), date_fmt)
            # Return formula referencing Data!E (Close)
            formula = f"=(Data!E{data_row_excel+1})/(Data!E{data_row_excel})-1"
            ws_ret.write_formula(ret_row, 1, formula, pct_fmt)
            ret_row += 1
        ws_ret.set_column("A:A", 12, date_fmt)
        ws_ret.set_column("B:B", 16, pct_fmt)

        # Sheet: Descriptive Stats
        stats = [
            ("Count (months)", n, None),
            ("Mean (monthly)", mean, "pct"),
            ("Median", median, "pct"),
            ("Mode", mode, "pct"),
            ("Range", r_range, "pct"),
            ("Variance (sample)", var_s, None),
            ("Std Dev (sample)", std_s, None),
            ("20th percentile", p20, "pct"),
            ("60th percentile", p60, "pct"),
            ("90th percentile", p90, "pct"),
            ("Min", r_min, "pct"),
            ("Q1", q1, "pct"),
            ("Median (Q2)", median, "pct"),
            ("Q3", q3, "pct"),
            ("Max", r_max, "pct"),
            ("IQR (Q3 - Q1)", iqr, "pct"),
            ("Lower bound (IQR)", lower_bound, "pct"),
            ("Upper bound (IQR)", upper_bound, "pct"),
        ]
        ws_stats = wb.add_worksheet("Descriptive Stats")
        ws_stats.write(0, 0, "Metric", bold)
        ws_stats.write(0, 1, "Value", bold)
        for i, (k, v, kind) in enumerate(stats, start=1):
            ws_stats.write(i, 0, k)
            if kind == "pct":
                ws_stats.write_number(i, 1, v, pct_fmt)
            else:
                ws_stats.write_number(i, 1, v)

        # Outliers table
        ws_stats.write(0, 3, "Outliers (IQR rule)", bold)
        ws_stats.write(1, 3, "Date", bold)
        ws_stats.write(1, 4, "Return", bold)
        for j in range(len(outliers_df)):
            ws_stats.write_datetime(j + 2, 3, pd.to_datetime(outliers_df.loc[j, "Date"]).to_pydatetime(), date_fmt)
            ws_stats.write_number(j + 2, 4, float(outliers_df.loc[j, "Return"]), pct_fmt)

        ws_stats.set_column("A:A", 26)
        ws_stats.set_column("B:B", 18)
        ws_stats.set_column("D:D", 12, date_fmt)
        ws_stats.set_column("E:E", 16, pct_fmt)

        # Sheet: Frequency
        freq_sheet = wb.add_worksheet("Frequency")
        freq_sheet.write(0, 0, "Class Interval", bold)
        freq_sheet.write(0, 1, "Frequency", bold)
        freq_sheet.write(0, 2, "Relative Frequency", bold)
        for i in range(len(freq_df)):
            freq_sheet.write(i + 1, 0, str(freq_df.loc[i, "Class Interval"]))
            freq_sheet.write_number(i + 1, 1, int(freq_df.loc[i, "Frequency"]))
            freq_sheet.write_number(i + 1, 2, float(freq_df.loc[i, "Relative Frequency"]), pct1_fmt)
        freq_sheet.set_column("A:A", 22)
        freq_sheet.set_column("B:B", 12)
        freq_sheet.set_column("C:C", 18, pct1_fmt)

        # Sheet: Histogram (embedded chart)
        hist_sheet = wb.add_worksheet("Histogram")
        # Create column chart with categories from Frequency!A and values from Frequency!B
        chart = wb.add_chart({"type": "column"})
        last_row = len(freq_df) + 1  # header + data
        if len(freq_df) > 0:
            chart.add_series({
                "name": "Frequency",
                "categories": f"=Frequency!$A$2:$A${last_row}",
                "values":     f"=Frequency!$B$2:$B${last_row}",
                "data_labels": {"value": True},
            })
        chart.set_title({"name": "Monthly Return Distribution (10% bins)"})
        chart.set_x_axis({"name": "Return Interval", "num_font": {"size": 9}})
        chart.set_y_axis({"name": "Frequency"})
        chart.set_legend({"position": "none"})
        hist_sheet.insert_chart("A1", chart, {"x_scale": 1.5, "y_scale": 1.5})

        # Sheet: Summary (textual observations and conclusion)
        ws_sum = wb.add_worksheet("Summary")
        ws_sum.set_column("A:A", 110)
        lines = []
        lines.append(f"Sample size: {n} months.")
        if n:
            lines.append(f"Mean monthly return: {mean:.2%}; Std Dev (monthly): {std_s:.2%}.")
            lines.append(f"Min: {r_min:.2%}, 20th pct: {p20:.2%}, Median: {median:.2%}, 60th pct: {p60:.2%}, 90th pct: {p90:.2%}, Max: {r_max:.2%}.")
            lines.append(f"IQR: {iqr:.2%}; IQR outlier bounds: [{lower_bound:.2%}, {upper_bound:.2%}]. Outliers detected: {len(outliers_df)}.")
            lines.append("Distribution insight: Wide spread and fat tails are typical for equities; large positive outliers drive the right tail, while drawdowns produce left-tail risk.")
            # Investment conclusion with Ecclesiastes 11:2
            lines.append("")
            lines.append("Conclusion:")
            lines.append(
                "Based on historical monthly returns, the stock exhibits high volatility but strong upside over time. "
                "A data-driven approach suggests it can be part of a growth portfolio, provided risk controls (position sizing, rebalancing) are used."
            )
            lines.append(
                "In light of Ecclesiastes 11:2 (“Invest in seven ventures, yes, in eight; for you do not know what disaster may come upon the land”), "
                "this supports a diversified approach: consider investing in this stock as one of multiple holdings, not a concentrated bet."
            )
        text = "\n".join(lines) if lines else "No data."
        ws_sum.write(0, 0, text)

    print(f"Wrote {OUTPUT_PATH}")

if __name__ == "__main__":
    main()