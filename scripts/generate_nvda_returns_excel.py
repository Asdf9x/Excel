import os
import numpy as np
import pandas as pd

from datetime import datetime

OUTPUT_PATH = "reports/NVDA_Returns_Analysis.xlsx"
DATA_PATH = "data/nvda_us_m.csv"

def main():
    # 1) Load data
    df = pd.read_csv(DATA_PATH, parse_dates=["Date"])
    df = df.sort_values("Date").reset_index(drop=True)

    # 2) Prepare folders
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    # 3) Write Excel with formulas (not pre-computed numbers)
    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        # Formats
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        pct1_fmt = wb.add_format({"num_format": "0.0%"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        num_fmt = wb.add_format({"num_format": "0.000000"})
        int_fmt = wb.add_format({"num_format": "0"})
        bold = wb.add_format({"bold": True})

        # Sheet: Data (原始数据)
        df.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        ws_data.set_column("A:A", 12, date_fmt)
        ws_data.set_column("B:F", 12)

        # Sheet: Returns（用公式从 Data!Close 计算简单月度收益）
        ws_ret = wb.add_worksheet("Returns")
        ws_ret.write(0, 0, "Date", bold)
        ws_ret.write(0, 1, "Return", bold)

        n_rows = len(df)
        # 写入日期和收益率公式
        ret_row = 1
        for i in range(1, n_rows):
            data_row_excel = i + 1  # Data表中的Excel行号（从1开始，含表头）
            # 日期直接写值
            ws_ret.write_datetime(ret_row, 0, df.loc[i, "Date"].to_pydatetime(), date_fmt)
            # 简单收益率公式： (Close_t / Close_{t-1}) - 1
            formula = f"=(Data!E{data_row_excel+1})/(Data!E{data_row_excel})-1"
            ws_ret.write_formula(ret_row, 1, formula, pct_fmt)
            ret_row += 1

        ws_ret.set_column("A:A", 12, date_fmt)
        ws_ret.set_column("B:B", 16, pct_fmt)

        # 将 Returns 区域添加为表，便于结构化引用（tReturns）
        # 表范围：从(0,0)到(last_row,1)，包含表头
        last_table_row = ret_row - 1  # 数据最后一行的0基索引
        ws_ret.add_table(0, 0, last_table_row, 1, {
            "name": "tReturns",
            "columns": [{"name": "Date"}, {"name": "Return"}]
        })

        # Sheet: Descriptive Stats（全部用公式）
        ws_stats = wb.add_worksheet("Descriptive Stats")
        ws_stats.write(0, 0, "Metric", bold)
        ws_stats.write(0, 1, "Value", bold)

        stats_rows = [
            ("Count (months)",        "=ROWS(tReturns[Return])",                                  "int"),
            ("Mean (monthly)",        "=AVERAGE(tReturns[Return])",                               "pct"),
            ("Median",                "=MEDIAN(tReturns[Return])",                                "pct"),
            ("Mode",                  "=MODE.SNGL(tReturns[Return])",                             "pct"),
            ("Range",                 "=(MAX(tReturns[Return]) - MIN(tReturns[Return]))",         "pct"),
            ("Variance (sample)",     "=VAR.S(tReturns[Return])",                                 "num"),
            ("Std Dev (sample)",      "=STDEV.S(tReturns[Return])",                               "pct"),
            ("20th percentile",       "=PERCENTILE.INC(tReturns[Return],0.2)",                    "pct"),
            ("60th percentile",       "=PERCENTILE.INC(tReturns[Return],0.6)",                    "pct"),
            ("90th percentile",       "=PERCENTILE.INC(tReturns[Return],0.9)",                    "pct"),
            ("Min",                   "=MIN(tReturns[Return])",                                   "pct"),
            ("Q1",                    "=QUARTILE.INC(tReturns[Return],1)",                        "pct"),
            ("Median (Q2)",           "=MEDIAN(tReturns[Return])",                                "pct"),
            ("Q3",                    "=QUARTILE.INC(tReturns[Return],3)",                        "pct"),
            ("Max",                   "=MAX(tReturns[Return])",                                   "pct"),
            ("IQR (Q3 - Q1)",         "=(QUARTILE.INC(tReturns[Return],3)-QUARTILE.INC(tReturns[Return],1))", "pct"),
            ("Lower bound (IQR)",     "=(QUARTILE.INC(tReturns[Return],1) - 1.5*(QUARTILE.INC(tReturns[Return],3)-QUARTILE.INC(tReturns[Return],1)))", "pct"),
            ("Upper bound (IQR)",     "=(QUARTILE.INC(tReturns[Return],3) + 1.5*(QUARTILE.INC(tReturns[Return],3)-QUARTILE.INC(tReturns[Return],1)))", "pct"),
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

        # Outliers 区（动态数组，自动溢出）
        ws_stats.write(0, 3, "Outliers (IQR rule)", bold)
        ws_stats.write(1, 3, "Date", bold)
        ws_stats.write(1, 4, "Return", bold)

        # 用 FILTER 返回两列（Date, Return）
        # 条件：Return < LowerBound 或 Return > UpperBound（直接在公式内重算IQR上下界，避免跨表绝对引用）
        outlier_formula = (
            "=FILTER("
            "tReturns[[Date]:[Return]],"
            "(tReturns[Return]<(QUARTILE.INC(tReturns[Return],1) - 1.5*(QUARTILE.INC(tReturns[Return],3)-QUARTILE.INC(tReturns[Return],1))))+"
            "(tReturns[Return]>(QUARTILE.INC(tReturns[Return],3) + 1.5*(QUARTILE.INC(tReturns[Return],3)-QUARTILE.INC(tReturns[Return],1))))"
            ")"
        )
        ws_stats.write_formula(2, 3, outlier_formula)
        ws_stats.set_column("A:A", 26)
        ws_stats.set_column("B:B", 18)
        ws_stats.set_column("D:D", 12, date_fmt)
        ws_stats.set_column("E:E", 16, pct_fmt)

        # Sheet: Frequency（纯公式生成10%组距、频数、相对频数；动态数组）
        ws_freq = wb.add_worksheet("Frequency")
        ws_freq.write(0, 0, "Class Lower", bold)
        ws_freq.write(0, 1, "Class Interval", bold)
        ws_freq.write(0, 2, "Frequency", bold)
        ws_freq.write(0, 3, "Relative Frequency", bold)

        ws_freq.write(0, 5, "Bin width", bold)   # F1
        ws_freq.write_number(0, 6, 0.10, pct_fmt)  # G1 = 10%
        ws_freq.write(1, 5, "Lower Edge", bold)  # F2
        ws_freq.write(2, 5, "Upper Edge", bold)  # F3
        ws_freq.write(3, 5, "Bin Count", bold)   # F4

        # F2: lower edge = floor(min to 10%)
        ws_freq.write_formula(1, 6, "=FLOOR.MATH(MIN(tReturns[Return])/$G$1,1)*$G$1", pct_fmt)
        # F3: upper edge = ceiling(max to 10%)
        ws_freq.write_formula(2, 6, "=CEILING.MATH(MAX(tReturns[Return])/$G$1,1)*$G$1", pct_fmt)
        # F4: bin count = INT((upper-lower)/bin)
        ws_freq.write_formula(3, 6, "=INT((G3-G2)/$G$1)", int_fmt)

        # A2: 下界序列（动态溢出）
        ws_freq.write_formula(1, 0, "=SEQUENCE($G$4,1,$G$2,$G$1)", pct_fmt)
        # B2: 文字标签（动态溢出，依赖A2#）
        ws_freq.write_formula(1, 1, '=TEXT(A2#,"0%")&" to "&TEXT(A2# + $G$1,"0%")')
        # C2: 频数（动态溢出）
        ws_freq.write_formula(1, 2, '=COUNTIFS(tReturns[Return],">="&A2#, tReturns[Return],"<"&A2# + $G$1)')
        # D2: 相对频数（动态溢出）
        ws_freq.write_formula(1, 3, "=C2#/SUM(C2#)")

        ws_freq.set_column("A:A", 14, pct_fmt)
        ws_freq.set_column("B:B", 20)
        ws_freq.set_column("C:C", 12)
        ws_freq.set_column("D:D", 18, pct1_fmt)
        ws_freq.set_column("F:F", 12)
        ws_freq.set_column("G:G", 14)

        # Sheet: Histogram（嵌入式图表，引用溢出区域）
        ws_hist = wb.add_worksheet("Histogram")
        chart = wb.add_chart({"type": "column"})
        # 使用溢出区域引用作为分类与数值
        chart.add_series({
            "name": "Frequency",
            "categories": "=Frequency!$B$2#",
            "values":     "=Frequency!$C$2#",
            "data_labels": {"value": True},
        })
        chart.set_title({"name": "Monthly Return Distribution (10% bins)"})
        chart.set_x_axis({"name": "Return Interval", "num_font": {"size": 9}})
        chart.set_y_axis({"name": "Frequency"})
        chart.set_legend({"position": "none"})
        ws_hist.insert_chart("A1", chart, {"x_scale": 1.5, "y_scale": 1.5})

        # Sheet: Summary（用公式拼接，自动引用统计结果）
        ws_sum = wb.add_worksheet("Summary")
        ws_sum.set_column("A:A", 110)

        ws_sum.write_formula(0, 0, '="Sample size: "&TEXT(\'Descriptive Stats\'!B1,"0")&" months." )')
        ws_sum.write_formula(1, 0, '="Mean monthly return: "&TEXT(\'Descriptive Stats\'!B2,"0.00%")&"; Std Dev (monthly): "&TEXT(\'Descriptive Stats\'!B7,"0.00%")&"."')
        ws_sum.write_formula(2, 0, '="Min: "&TEXT(\'Descriptive Stats\'!B11,"0.00%")&", 20th pct: "&TEXT(\'Descriptive Stats\'!B8,"0.00%")&", Median: "&TEXT(\'Descriptive Stats\'!B3,"0.00%")&", 60th pct: "&TEXT(\'Descriptive Stats\'!B9,"0.00%")&", 90th pct: "&TEXT(\'Descriptive Stats\'!B10,"0.00%")&", Max: "&TEXT(\'Descriptive Stats\'!B15,"0.00%")&"."')
        ws_sum.write_formula(3, 0, '="IQR: "&TEXT(\'Descriptive Stats\'!B16,"0.00%")&"; IQR outlier bounds: ["&TEXT(\'Descriptive Stats\'!B17,"0.00%")&", "&TEXT(\'Descriptive Stats\'!B18,"0.00%")&"]."')
        ws_sum.write(5, 0, "Conclusion:", bold)
        ws_sum.write(6, 0,
            "Based on historical monthly returns, the stock exhibits high volatility but strong upside over time. "
            "Consider it as part of a diversified, risk-managed portfolio (position sizing, rebalancing). "
            "Ecclesiastes 11:2 encourages diversification; investing in this stock as one of multiple holdings supports that worldview."
        )

    print(f"Wrote {OUTPUT_PATH} at {datetime.utcnow().isoformat()}Z")

if __name__ == "__main__":
    main()
