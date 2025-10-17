import os
import pandas as pd
from datetime import datetime

OUTPUT_PATH = "reports/NVDA_Returns_Analysis.xlsx"
DATA_PATH = "data/nvda_us_m.csv"

def main():
    # 1) Load data
    df = pd.read_csv(DATA_PATH, parse_dates=["Date"])
    df = df.sort_values("Date").reset_index(drop=True)

    # 2) Ensure folder
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    # 3) Write Excel (all metrics via formulas, high compatibility)
    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        # Formats
        pct_fmt  = wb.add_format({"num_format": "0.00%"})
        pct1_fmt = wb.add_format({"num_format": "0.0%"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        num_fmt  = wb.add_format({"num_format": "0.000000"})
        int_fmt  = wb.add_format({"num_format": "0"})
        bold     = wb.add_format({"bold": True})

        # Sheet: Data
        df.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        ws_data.set_column("A:A", 12, date_fmt)
        ws_data.set_column("B:F", 12)

        # Sheet: Returns (simple returns via formula from Close)
        ws_ret = wb.add_worksheet("Returns")
        ws_ret.write(0, 0, "Date", bold)
        ws_ret.write(0, 1, "Return", bold)
        ws_ret.write(0, 2, "IsOutlier", bold)   # helper
        ws_ret.write(0, 3, "OutlierRow", bold)  # helper (row numbers for outliers)

        n_rows = len(df)
        ret_row = 1
        for i in range(1, n_rows):
            data_row_excel = i + 1  # header row=1, first data row=2
            ws_ret.write_datetime(ret_row, 0, df.loc[i, "Date"].to_pydatetime(), date_fmt)
            # Return_t = Close_t / Close_{t-1} - 1
            formula_ret = f"=(Data!E{data_row_excel+1})/(Data!E{data_row_excel})-1"
            ws_ret.write_formula(ret_row, 1, formula_ret, pct_fmt)
            # IsOutlier will be filled after Stats sheet exists (we still can write formula now)
            # =OR(Bk < Stats!$B$17, Bk > Stats!$B$18)
            row_idx = ret_row + 1  # Excel row number
            formula_flag = f"=OR(B{row_idx}<'Descriptive Stats'!$B$17, B{row_idx}>'Descriptive Stats'!$B$18)"
            ws_ret.write_formula(ret_row, 2, formula_flag)
            # OutlierRow = IF(IsOutlier, ROW(), "")
            formula_row = f"=IF(C{row_idx}, ROW(), \"\")"
            ws_ret.write_formula(ret_row, 3, formula_row)
            ret_row += 1

        ws_ret.set_column("A:A", 12, date_fmt)
        ws_ret.set_column("B:B", 16, pct_fmt)
        ws_ret.set_column("C:C", 12)
        ws_ret.set_column("D:D", 12)

        # Dynamic range for returns (as formula string)
        # Returns from B2 to last numeric cell in column B
        RET_RANGE = "Returns!$B$2:INDEX(Returns!$B:$B, MATCH(1E+99, Returns!$B:$B))"

        # Sheet: Descriptive Stats (all via formulas, no structured refs)
        ws_stats = wb.add_worksheet("Descriptive Stats")
        ws_stats.write(0, 0, "Metric", bold)
        ws_stats.write(0, 1, "Value", bold)

        stats_rows = [
            ("Count (months)",    f"=COUNT({RET_RANGE})",                            "int"),
            ("Mean (monthly)",    f"=AVERAGE({RET_RANGE})",                          "pct"),
            ("Median",            f"=MEDIAN({RET_RANGE})",                           "pct"),
            ("Mode",              f"=MODE.SNGL({RET_RANGE})",                        "pct"),
            ("Range",             f"=(MAX({RET_RANGE})-MIN({RET_RANGE}))",           "pct"),
            ("Variance (sample)", f"=VAR.S({RET_RANGE})",                             "num"),
            ("Std Dev (sample)",  f"=STDEV.S({RET_RANGE})",                           "pct"),
            ("20th percentile",   f"=PERCENTILE.INC({RET_RANGE},0.2)",               "pct"),
            ("60th percentile",   f"=PERCENTILE.INC({RET_RANGE},0.6)",               "pct"),
            ("90th percentile",   f"=PERCENTILE.INC({RET_RANGE},0.9)",               "pct"),
            ("Min",               f"=MIN({RET_RANGE})",                              "pct"),
            ("Q1",                f"=QUARTILE.INC({RET_RANGE},1)",                   "pct"),
            ("Median (Q2)",       f"=MEDIAN({RET_RANGE})",                           "pct"),
            ("Q3",                f"=QUARTILE.INC({RET_RANGE},3)",                   "pct"),
            ("Max",               f"=MAX({RET_RANGE})",                              "pct"),
            ("IQR (Q3 - Q1)",     f"=(QUARTILE.INC({RET_RANGE},3)-QUARTILE.INC({RET_RANGE},1))", "pct"),
            ("Lower bound (IQR)", f"=(QUARTILE.INC({RET_RANGE},1) - 1.5*(QUARTILE.INC({RET_RANGE},3)-QUARTILE.INC({RET_RANGE},1)))", "pct"),
            ("Upper bound (IQR)", f"=(QUARTILE.INC({RET_RANGE},3) + 1.5*(QUARTILE.INC({RET_RANGE},3)-QUARTILE.INC({RET_RANGE},1)))", "pct"),
            ("Outliers (count)",  f"=COUNTIFS({RET_RANGE},\"<\"&B17)+COUNTIFS({RET_RANGE},\">\"&B18)", "int"),
        ]

        for i, (label, formula, kind) in enumerate(stats_rows, start=1):
            ws_stats.write(i, 0, label)
            if kind == "pct":
                ws_stats.write_formula(i, 1, formula, pct_fmt)
            elif kind == "int":
                ws_stats.write_formula(i, 1, formula, int_fmt)
            elif kind == "num":
                ws_stats.write_formula(i, 1, formula, num_fmt)
            else:
                ws_stats.write_formula(i, 1, formula)

        ws_stats.set_column("A:A", 26)
        ws_stats.set_column("B:B", 18)

        # Sheet: Outliers (no FILTER; use SMALL over helper OutlierRow)
        ws_out = wb.add_worksheet("Outliers")
        ws_out.write(0, 0, "Date", bold)
        ws_out.write(0, 1, "Return", bold)
        # Pre-fill 200 rows (more than enough)
        for k in range(1, 201):
            row = k + 1
            # A2: =IFERROR(INDEX(Returns!A:A, SMALL(Returns!$D:$D, ROWS($A$1:A1))), "")
            fA = f'=IFERROR(INDEX(Returns!A:A, SMALL(Returns!$D:$D, ROWS($A$1:A{row-1}))), "")'
            fB = f'=IFERROR(INDEX(Returns!B:B, SMALL(Returns!$D:$D, ROWS($A$1:A{row-1}))), "")'
            ws_out.write_formula(row-1, 0, fA, date_fmt)
            ws_out.write_formula(row-1, 1, fB, pct_fmt)
        ws_out.set_column("A:A", 12, date_fmt)
        ws_out.set_column("B:B", 16, pct_fmt)

        # Sheet: Frequency (no dynamic arrays; formulas copied for many rows)
        ws_freq = wb.add_worksheet("Frequency")
        ws_freq.write(0, 0, "Class Lower", bold)
        ws_freq.write(0, 1, "Class Interval", bold)
        ws_freq.write(0, 2, "Frequency", bold)
        ws_freq.write(0, 3, "Relative Frequency", bold)
        ws_freq.write(0, 5, "Bin width", bold)   # F1
        ws_freq.write_number(0, 6, 0.10, pct_fmt) # G1 = 10%
        ws_freq.write(1, 5, "Lower Edge", bold)  # F2
        ws_freq.write(2, 5, "Upper Edge", bold)  # F3

        # Lower edge: INT(MIN/width)*width works with negatives
        ws_freq.write_formula(1, 6, f"=INT(MIN({RET_RANGE})/$G$1)*$G$1", pct_fmt)  # G2
        # Upper edge: CEILING for positive max
        ws_freq.write_formula(2, 6, f"=CEILING(MAX({RET_RANGE}),$G$1)", pct_fmt)   # G3

        # Pre-fill up to 60 bins; trailing bins auto-blank
        max_bins = 60
        for i in range(max_bins):
            r = i + 1  # 1-based offset for data rows (A2 is i=0)
            row = r + 1
            if i == 0:
                f_lower = "=G$2"
            else:
                # If previous row blank or exceeds range, stay blank; else add width
                f_lower = f'=IF(OR(A{row-1}="", A{row-1}>$G$3-$G$1), "", A{row-1}+$G$1)'
            ws_freq.write_formula(row-1, 0, f_lower, pct_fmt)

            # Label
            f_label = f'=IF(A{row}="","",TEXT(A{row},"0%")&" to "&TEXT(A{row}+$G$1,"0%"))'
            ws_freq.write_formula(row-1, 1, f_label)

            # Frequency
            f_freq = f'=IF(A{row}="","",COUNTIFS(Returns!$B:$B,">="&A{row}, Returns!$B:$B,"<"&A{row}+$G$1))'
            ws_freq.write_formula(row-1, 2, f_freq)

            # Relative frequency (sum over the 60 rows)
            f_relf = f'=IF(C{row}="","", C{row}/SUM($C$2:$C${max_bins+1}))'
            ws_freq.write_formula(row-1, 3, f_relf, pct1_fmt)

        ws_freq.set_column("A:A", 14, pct_fmt)
        ws_freq.set_column("B:B", 20)
        ws_freq.set_column("C:C", 12)
        ws_freq.set_column("D:D", 18, pct1_fmt)
        ws_freq.set_column("F:F", 12)
        ws_freq.set_column("G:G", 14)

        # Sheet: Histogram (embedded chart)
        ws_hist = wb.add_worksheet("Histogram")
        chart = wb.add_chart({"type": "column"})
        # Categories/values: use a conservative fixed range covering the 60 bins
        chart.add_series({
            "name": "Frequency",
            "categories": f"=Frequency!$B$2:$B${max_bins+1}",
            "values":     f"=Frequency!$C$2:$C${max_bins+1}",
            "data_labels": {"value": True},
        })
        chart.set_title({"name": "Monthly Return Distribution (10% bins)"})
        chart.set_x_axis({"name": "Return Interval", "num_font": {"size": 9}})
        chart.set_y_axis({"name": "Frequency"})
        chart.set_legend({"position": "none"})
        ws_hist.insert_chart("A1", chart, {"x_scale": 1.5, "y_scale": 1.5})

        # Sheet: Summary (formulas referencing Stats)
        ws_sum = wb.add_worksheet("Summary")
        ws_sum.set_column("A:A", 110)
        ws_sum.write(0, 0, "Conclusion:", bold)
        ws_sum.write_formula(1, 0, '="Sample size: "&TEXT(\'Descriptive Stats\'!B1,"0")&" months."')
        ws_sum.write_formula(2, 0, '="Mean monthly return: "&TEXT(\'Descriptive Stats\'!B2,"0.00%")&"; Std Dev (monthly): "&TEXT(\'Descriptive Stats\'!B7,"0.00%")&"."')
        ws_sum.write_formula(3, 0, '="Min: "&TEXT(\'Descriptive Stats\'!B11,"0.00%")&", 20th pct: "&TEXT(\'Descriptive Stats\'!B8,"0.00%")&", Median: "&TEXT(\'Descriptive Stats\'!B3,"0.00%")&", 60th pct: "&TEXT(\'Descriptive Stats\'!B9,"0.00%")&", 90th pct: "&TEXT(\'Descriptive Stats\'!B10,"0.00%")&", Max: "&TEXT(\'Descriptive Stats\'!B15,"0.00%")&"."')
        ws_sum.write_formula(4, 0, '="IQR: "&TEXT(\'Descriptive Stats\'!B16,"0.00%")&"; Outlier bounds: ["&TEXT(\'Descriptive Stats\'!B17,"0.00%")&", "&TEXT(\'Descriptive Stats\'!B18,"0.00%")&"]; Count: "&TEXT(\'Descriptive Stats\'!B19,"0")')
        ws_sum.write(6, 0,
            "Based on historical monthly returns, the stock is volatile with meaningful upside. "
            "Consider it as one of multiple holdings in a diversified, risk-managed portfolio. "
            "Ecclesiastes 11:2 encourages diversification; investing in this stock as part of 7-8 ventures supports that worldview."
        )

    print(f"Wrote {OUTPUT_PATH} at {datetime.utcnow().isoformat()}Z")

if __name__ == "__main__":
    main()
