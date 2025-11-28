import os
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ============================================================
# CONFIG
# ============================================================
FILE_NAME = "student_grades_2027-2028.xlsx"
ALERT_THRESHOLD = 60
SUBJECT_COLS = ["math", "english", "science", "history"]
NEW_COHORT_CANDIDATES = ["23-24", "24-25", "25-26"]
CLEAN_CSV_DIR = "cleaned_data"         # folder for cleaned CSVs
DASHBOARD_FILE = "Final_Dashboard.xlsx"

# Color palette
CHART_COLORS = ["#ff9999", "#66b3ff", "#99ff99"]  # red, blue, green


def series_color(i: int) -> str:
    """Cycle through the red / blue / green palette."""
    return CHART_COLORS[i % len(CHART_COLORS)]


# ============================================================
# DATA LOADING & CLEANING
# ============================================================
def load_excel_file(file_name: str) -> Optional[pd.ExcelFile]:
    file_path = os.path.join(os.getcwd(), file_name)
    try:
        xls = pd.ExcelFile(file_path)
        return xls
    except FileNotFoundError:
        print(f"Error: File '{file_name}' not found.")
        return None
    except Exception as e:
        print("An unexpected error occurred while loading Excel:", e)
        return None


def load_track_sheets(xls: pd.ExcelFile) -> Tuple[Optional[pd.DataFrame],
                                                  Optional[pd.DataFrame],
                                                  Optional[pd.DataFrame]]:
    data_df = pd.read_excel(xls, sheet_name="Data") if "Data" in xls.sheet_names else None
    fin_df = pd.read_excel(xls, sheet_name="Finance") if "Finance" in xls.sheet_names else None
    bm_df = pd.read_excel(xls, sheet_name="BM") if "BM" in xls.sheet_names else None

    print("Sheets loaded:")
    print("Data:", data_df is not None)
    print("Finance:", fin_df is not None)
    print("BM:", bm_df is not None)

    return data_df, fin_df, bm_df


def concat_with_track_labels(
    data_df: Optional[pd.DataFrame],
    fin_df: Optional[pd.DataFrame],
    bm_df: Optional[pd.DataFrame],
) -> Optional[pd.DataFrame]:
    dfs_to_concat: List[pd.DataFrame] = []

    if data_df is not None:
        df = data_df.copy()
        df["track"] = "Data"
        dfs_to_concat.append(df)

    if fin_df is not None:
        df = fin_df.copy()
        df["track"] = "Finance"
        dfs_to_concat.append(df)

    if bm_df is not None:
        df = bm_df.copy()
        df["track"] = "BM"
        dfs_to_concat.append(df)

    if not dfs_to_concat:
        print("No sheets available to concatenate.")
        return None

    return pd.concat(dfs_to_concat, ignore_index=True)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df_clean = df.copy()

    rename_map = {
        "StudentID": "student_id",
        "FirstName": "first_name",
        "LastName": "last_name",
        "Class": "class_name",
        "Term": "term",
        "Math": "math",
        "English": "english",
        "Science": "science",
        "History": "history",
        "Attendance (%)": "attendance",
        "ProjectScore": "project_score",
        "Passed (Y/N)": "passed",
        "IncomeStudent": "income_student",
        "Cohort": "cohort",
        "Track": "track",
    }

    safe_map = {old: new for old, new in rename_map.items() if old in df_clean.columns}
    df_clean = df_clean.rename(columns=safe_map)

    df_clean.columns = (
        df_clean.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_", regex=False)
    )

    bad_strings = ["", "n/a", "waived"]

    mask = (
        df_clean.isna()
        | df_clean.apply(
            lambda col: col.astype(str)
            .str.strip()
            .str.lower()
            .isin(bad_strings)
        )
    )

    df_clean = df_clean[~mask.any(axis=1)].copy()

    required_cols = ["student_id", "passed", "income_student"]
    missing = [c for c in required_cols if c not in df_clean.columns]
    if missing:
        raise KeyError(
            f"Missing required columns after cleaning: {missing}. "
            f"Current columns: {list(df_clean.columns)}"
        )

    df_clean["student_id"] = (
        df_clean["student_id"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .astype(int)
    )

    df_clean["passed"] = df_clean["passed"].replace({"Y": 1, "N": 0})
    df_clean["income_student"] = df_clean["income_student"].replace(
        {"True": 1, "False": 0, True: 1, False: 0}
    )

    num_cols_float = ["math", "english", "science", "history", "project_score", "attendance"]
    num_cols_int = ["term", "passed", "income_student"]

    num_cols_float_existing = [c for c in num_cols_float if c in df_clean.columns]
    num_cols_int_existing = [c for c in num_cols_int if c in df_clean.columns]

    df_clean[num_cols_float_existing + num_cols_int_existing] = df_clean[
        num_cols_float_existing + num_cols_int_existing
    ].apply(lambda col: pd.to_numeric(col, errors="coerce"))

    df_clean = df_clean.dropna(subset=num_cols_float_existing + num_cols_int_existing)

    if num_cols_int_existing:
        df_clean[num_cols_int_existing] = df_clean[num_cols_int_existing].astype(int)

    score_cols = [col for col in num_cols_float_existing if col != "attendance"]
    if score_cols:
        df_clean[score_cols] = df_clean[score_cols].astype(int)

    df_clean = df_clean.set_index("student_id")

    return df_clean


# ============================================================
# SUMMARY TABLE BUILDERS
# ============================================================
def get_track_counts(data_df: pd.DataFrame,
                     fin_df: pd.DataFrame,
                     bm_df: pd.DataFrame) -> pd.DataFrame:
    num_d = len(data_df)
    num_f = len(fin_df)
    num_bm = len(bm_df)

    print(
        f"There are {num_d} students in the Data track, "
        f"{num_f} students in the Finance track, "
        f"and {num_bm} students in the Business track."
    )

    return pd.DataFrame({
        "track": ["Data", "Finance", "Business"],
        "count": [num_d, num_f, num_bm],
    })


def make_track_summary(uni_df: pd.DataFrame) -> pd.DataFrame:
    track_summary = uni_df.groupby("track", as_index=False).agg(
        {
            "math": "mean",
            "english": "mean",
            "science": "mean",
            "history": "mean",
            "attendance": "mean",
            "project_score": "mean",
        }
    )
    print("\nTrack summary:")
    print(track_summary)
    return track_summary


def make_pass_fail_summary(uni_df: pd.DataFrame) -> pd.DataFrame:
    vc = uni_df["passed"].value_counts()
    passed = vc.get(1, 0)
    failed = vc.get(0, 0)
    total = passed + failed if (passed + failed) > 0 else 1

    print(
        f"There are {round(passed / total * 100, 2)}% students who passed "
        f"and {round(failed / total * 100, 2)}% students who failed."
    )

    return pd.DataFrame({
        "Status": ["Pass", "Fail"],
        "Count": [passed, failed],
        "Percent": [passed / total * 100, failed / total * 100],
    })


def summarize_by_cohort(uni_df: pd.DataFrame) -> pd.DataFrame:
    cohort_summary = uni_df.groupby("cohort").agg(
        avg_math=("math", "mean"),
        avg_english=("english", "mean"),
        avg_science=("science", "mean"),
        avg_history=("history", "mean"),
        avg_attendance=("attendance", "mean"),
        avg_project=("project_score", "mean"),
        pass_rate=("passed", "mean"),
        count=("cohort", "size"),
    ).reset_index()

    print("\nCohort-Level Summary:")
    print(cohort_summary)

    return cohort_summary


def summarize_by_income(uni_df: pd.DataFrame) -> pd.DataFrame:
    income_summary = uni_df.groupby("income_student").agg(
        avg_math=("math", "mean"),
        avg_english=("english", "mean"),
        avg_science=("science", "mean"),
        avg_history=("history", "mean"),
        avg_attendance=("attendance", "mean"),
        avg_score=("project_score", "mean"),
        pass_rate=("passed", "mean"),
        count=("income_student", "size"),
    ).reset_index()

    income_summary["Group"] = income_summary["income_student"].map(
        {1: "Income Supported", 0: "Not Supported"}
    )

    print("\nIncome Student Performance Comparison:")
    print(income_summary)

    return income_summary


def generate_performance_alert_table(uni_df: pd.DataFrame,
                                     alert_threshold: int = ALERT_THRESHOLD) -> pd.DataFrame:
    df_alert = uni_df.copy()

    agg = (
        df_alert.groupby(["track", "cohort"])[SUBJECT_COLS]
        .mean()
        .reset_index()
    )

    agg["AvgScore"] = agg[SUBJECT_COLS].mean(axis=1)
    agg["LowPerformance"] = agg["AvgScore"] < alert_threshold
    agg["Alert"] = agg["LowPerformance"].apply(
        lambda x: "⚠ LOW PERFORMANCE" if x else "OK"
    )

    print("\n===== Performance Alert Table =====")
    print(agg.sort_values(["track", "cohort"]))

    return agg.sort_values(["track", "cohort"]).reset_index(drop=True)


# ============================================================
# SYNTHETIC COHORT GENERATION & TRENDS
# ============================================================
def generate_synthetic_cohorts(uni_df: pd.DataFrame,
                               candidates: List[str]) -> pd.DataFrame:
    base_df = uni_df.reset_index().copy()

    existing_cohorts = set(base_df["cohort"].astype(str).unique())
    new_cohorts = [c for c in candidates if c not in existing_cohorts]

    print("Existing cohorts:", existing_cohorts)
    print("New cohorts to generate:", new_cohorts)

    track_sizes: Dict[str, int] = (
        base_df.groupby("track")["student_id"].count().to_dict()
    )
    print("track sizes:", track_sizes)

    synthetic_rows = []
    max_id = base_df["student_id"].max()

    for cohort in new_cohorts:
        for track, size in track_sizes.items():
            track_df = base_df[base_df["track"] == track]
            sampled = track_df.sample(n=size, replace=True).copy()
            sampled["student_id"] = range(max_id + 1, max_id + 1 + size)
            max_id += size
            sampled["cohort"] = cohort

            for col in SUBJECT_COLS:
                sampled[col] = (
                    sampled[col].astype(float)
                    + np.random.normal(loc=0, scale=5, size=len(sampled))
                ).clip(0, 100).round().astype(int)

            synthetic_rows.append(sampled)

    synthetic_df = pd.concat(synthetic_rows, ignore_index=True)
    uni_df_all = pd.concat([base_df, synthetic_df], ignore_index=True)
    uni_df_all = uni_df_all.set_index("student_id")

    print("Cohort counts in extended dataset:")
    print(uni_df_all["cohort"].value_counts().sort_index())

    return uni_df_all


def build_trend_pivots(uni_df_all: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    uni_df_all = uni_df_all.copy()
    uni_df_all["cohort"] = uni_df_all["cohort"].astype(str)

    cohort_order = sorted(
        uni_df_all["cohort"].unique(), key=lambda x: int(x.split("-")[0])
    )
    uni_df_all["cohort"] = pd.Categorical(
        uni_df_all["cohort"], categories=cohort_order, ordered=True
    )

    trend_pivots: Dict[str, pd.DataFrame] = {}

    for subj in SUBJECT_COLS:
        mean_scores = (
            uni_df_all.groupby(["cohort", "track"])[subj]
            .mean()
            .reset_index()
        )
        pivot = mean_scores.pivot(index="cohort", columns="track", values=subj)
        trend_pivots[subj] = pivot.reset_index()

    return trend_pivots


# ============================================================
# EXPORTS & DASHBOARD
# ============================================================
def export_cleaned_data(
    data_df: pd.DataFrame,
    fin_df: pd.DataFrame,
    bm_df: pd.DataFrame,
    uni_df: pd.DataFrame,
) -> None:
    os.makedirs(CLEAN_CSV_DIR, exist_ok=True)

    data_df.reset_index().to_csv(os.path.join(CLEAN_CSV_DIR, "data_df_cleaned.csv"), index=False)
    fin_df.reset_index().to_csv(os.path.join(CLEAN_CSV_DIR, "fin_df_cleaned.csv"), index=False)
    bm_df.reset_index().to_csv(os.path.join(CLEAN_CSV_DIR, "bm_df_cleaned.csv"), index=False)
    uni_df.reset_index().to_csv(os.path.join(CLEAN_CSV_DIR, "uni_df_cleaned.csv"), index=False)

    print(f"CSV files exported successfully to folder: '{CLEAN_CSV_DIR}'")


def create_dashboard_excel(
    track_counts: pd.DataFrame,
    track_summary: pd.DataFrame,
    pass_fail_df: pd.DataFrame,
    cohort_summary: pd.DataFrame,
    income_summary: pd.DataFrame,
    alert_table: pd.DataFrame,
    trend_pivots: Dict[str, pd.DataFrame],
    output_file: str = DASHBOARD_FILE,
) -> None:
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#D9E1F2",
                "font_color": "#000000",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        num_fmt = workbook.add_format({"num_format": "0.00"})
        percent_fmt = workbook.add_format({"num_format": "0.0%", "align": "center"})
        center_fmt = workbook.add_format({"align": "center", "valign": "vcenter"})
        title_fmt = workbook.add_format(
            {"bold": True, "font_size": 16, "font_color": "#305496"}
        )

        # =========================================================
        # Overview
        # =========================================================
        ws_overview = workbook.add_worksheet("Overview")
        ws_overview.set_zoom(90)
        ws_overview.freeze_panes(5, 0)

        ws_overview.merge_range("A1:H1", "Student Performance Dashboard", title_fmt)

        # Track counts
        start_row_tc = 3
        ws_overview.write(start_row_tc, 0, "Track", header_fmt)
        ws_overview.write(start_row_tc, 1, "Count", header_fmt)
        for i, row in track_counts.iterrows():
            ws_overview.write(start_row_tc + 1 + i, 0, row["track"], center_fmt)
            ws_overview.write_number(start_row_tc + 1 + i, 1, row["count"])
        ws_overview.set_column("A:A", 15)
        ws_overview.set_column("B:B", 10)

        # Pass/fail
        start_row_pf = 3
        ws_overview.write(start_row_pf, 3, "Status", header_fmt)
        ws_overview.write(start_row_pf, 4, "Count", header_fmt)
        ws_overview.write(start_row_pf, 5, "Percent", header_fmt)
        for i, row in pass_fail_df.iterrows():
            ws_overview.write(start_row_pf + 1 + i, 3, row["Status"], center_fmt)
            ws_overview.write_number(start_row_pf + 1 + i, 4, row["Count"])
            ws_overview.write_number(start_row_pf + 1 + i, 5, row["Percent"] / 100, percent_fmt)
        ws_overview.set_column("D:D", 10)
        ws_overview.set_column("E:E", 10)
        ws_overview.set_column("F:F", 10)

        # Track pie (red/blue/green)
        chart_track_pie = workbook.add_chart({"type": "pie"})
        chart_track_pie.add_series({
            "name": "Track Distribution",
            "categories": ["Overview", start_row_tc + 1, 0, start_row_tc + len(track_counts), 0],
            "values": ["Overview", start_row_tc + 1, 1, start_row_tc + len(track_counts), 1],
            "points": [
                {"fill": {"color": series_color(0)}},
                {"fill": {"color": series_color(1)}},
                {"fill": {"color": series_color(2)}},
            ],
        })
        chart_track_pie.set_title({"name": "Track Distribution"})
        ws_overview.insert_chart("A10", chart_track_pie, {"x_scale": 1.1, "y_scale": 1.1})

        # Pass/fail pie (use red/green)
        chart_pf_pie = workbook.add_chart({"type": "pie"})
        chart_pf_pie.add_series({
            "name": "Pass / Fail",
            "categories": ["Overview", start_row_pf + 1, 3, start_row_pf + len(pass_fail_df), 3],
            "values": ["Overview", start_row_pf + 1, 4, start_row_pf + len(pass_fail_df), 4],
            "points": [
                {"fill": {"color": series_color(2)}},  # Pass -> green
                {"fill": {"color": series_color(0)}},  # Fail -> red
            ],
        })
        chart_pf_pie.set_title({"name": "Pass / Fail Ratio"})
        ws_overview.insert_chart("H10", chart_pf_pie, {"x_scale": 1.1, "y_scale": 1.1})

        # =========================================================
        # Track Summary
        # =========================================================
        track_summary.to_excel(writer, sheet_name="track_Summary", index=False)
        ws_track = writer.sheets["track_Summary"]
        ws_track.set_zoom(90)
        ws_track.freeze_panes(1, 1)

        rows_t = len(track_summary)
        cols_t = len(track_summary.columns)
        col_idx_map_t = {col: i for i, col in enumerate(track_summary.columns)}

        ws_track.set_row(0, 18, header_fmt)
        ws_track.set_column("A:A", 12, center_fmt)
        ws_track.set_column("B:G", 12, num_fmt)

        ws_track.add_table(
            0, 0, rows_t, cols_t - 1,
            {
                "style": "Table Style Medium 9",
                "columns": [{"header": col} for col in track_summary.columns],
            },
        )

        chart_track = workbook.add_chart({"type": "column"})
        max_row_track = rows_t + 1  # header + data

        for i, subj in enumerate(SUBJECT_COLS):  # math, english, science, history
            if subj in col_idx_map_t:
                cidx = col_idx_map_t[subj]
                chart_track.add_series(
                    {
                        "name": ["track_Summary", 0, cidx],
                        "categories": ["track_Summary", 1, 0, max_row_track - 1, 0],
                        "values": ["track_Summary", 1, cidx, max_row_track - 1, cidx],
                        "fill": {"color": series_color(i)},
                    }
                )

        chart_track.set_title({"name": "Average Scores by Track"})
        chart_track.set_x_axis({"name": "Track"})
        chart_track.set_y_axis({"name": "Score"})
        ws_track.insert_chart("I2", chart_track, {"x_scale": 1.2, "y_scale": 1.2})

        # =========================================================
        # Cohort Summary
        # =========================================================
        cohort_summary.to_excel(writer, sheet_name="Cohort_Summary", index=False)
        ws_cohort = writer.sheets["Cohort_Summary"]
        ws_cohort.set_zoom(90)
        ws_cohort.freeze_panes(1, 1)

        rows_c = len(cohort_summary)
        cols_c = len(cohort_summary.columns)
        col_idx_map_c = {col: i for i, col in enumerate(cohort_summary.columns)}

        ws_cohort.set_row(0, 18, header_fmt)
        ws_cohort.set_column("A:A", 10, center_fmt)
        ws_cohort.set_column("B:H", 12, num_fmt)
        ws_cohort.set_column("I:I", 10, num_fmt)

        ws_cohort.add_table(
            0, 0, rows_c, cols_c - 1,
            {
                "style": "Table Style Medium 9",
                "columns": [{"header": col} for col in cohort_summary.columns],
            },
        )

        chart_cohort = workbook.add_chart({"type": "column"})
        max_row_cohort = rows_c + 1

        cohort_subject_cols = ["avg_math", "avg_english", "avg_science", "avg_history"]
        for i, subj in enumerate(cohort_subject_cols):
            if subj in col_idx_map_c:
                cidx = col_idx_map_c[subj]
                chart_cohort.add_series(
                    {
                        "name": ["Cohort_Summary", 0, cidx],
                        "categories": ["Cohort_Summary", 1, 0, max_row_cohort - 1, 0],
                        "values": [
                            "Cohort_Summary",
                            1,
                            cidx,
                            max_row_cohort - 1,
                            cidx,
                        ],
                        "fill": {"color": series_color(i)},
                    }
                )

        chart_cohort.set_title({"name": "Average Scores by Cohort"})
        chart_cohort.set_x_axis({"name": "Cohort"})
        chart_cohort.set_y_axis({"name": "Score"})
        ws_cohort.insert_chart("J2", chart_cohort, {"x_scale": 1.1, "y_scale": 1.1})

        # Pass rate chart (blue)
        chart_passrate = workbook.add_chart({"type": "column"})
        if "pass_rate" in col_idx_map_c:
            pr_idx = col_idx_map_c["pass_rate"]
            chart_passrate.add_series({
                "name": "Pass Rate",
                "categories": ["Cohort_Summary", 1, 0, max_row_cohort - 1, 0],
                "values": ["Cohort_Summary", 1, pr_idx, max_row_cohort - 1, pr_idx],
                "fill": {"color": series_color(1)},  # blue
            })
        chart_passrate.set_title({"name": "Pass Rate by Cohort"})
        chart_passrate.set_y_axis({"name": "Pass Rate"})
        ws_cohort.insert_chart("J20", chart_passrate, {"x_scale": 1.1, "y_scale": 1.1})

        # =========================================================
        # Income Summary
        # =========================================================
        income_summary.to_excel(writer, sheet_name="Income_Summary", index=False)
        ws_income = writer.sheets["Income_Summary"]
        ws_income.set_zoom(90)
        ws_income.freeze_panes(1, 1)

        rows_i = len(income_summary)
        cols_i = len(income_summary.columns)
        col_idx_map_i = {col: i for i, col in enumerate(income_summary.columns)}

        ws_income.set_row(0, 18, header_fmt)
        ws_income.set_column("A:A", 14, center_fmt)
        ws_income.set_column("B:H", 12, num_fmt)
        ws_income.set_column("I:I", 10, num_fmt)
        ws_income.set_column("J:J", 16, center_fmt)

        ws_income.add_table(
            0, 0, rows_i, cols_i - 1,
            {
                "style": "Table Style Medium 9",
                "columns": [{"header": col} for col in income_summary.columns],
            },
        )

        chart_income = workbook.add_chart({"type": "column"})
        max_row_income = rows_i + 1

        income_subject_cols = ["avg_math", "avg_english", "avg_science", "avg_history"]
        for i, subj in enumerate(income_subject_cols):
            if subj in col_idx_map_i and "Group" in col_idx_map_i:
                sidx = col_idx_map_i[subj]
                gidx = col_idx_map_i["Group"]
                chart_income.add_series(
                    {
                        "name": ["Income_Summary", 0, sidx],
                        "categories": ["Income_Summary", 1, gidx, max_row_income - 1, gidx],
                        "values": [
                            "Income_Summary",
                            1,
                            sidx,
                            max_row_income - 1,
                            sidx,
                        ],
                        "fill": {"color": series_color(i)},
                    }
                )

        chart_income.set_title({"name": "Subject Averages by Income Group"})
        chart_income.set_x_axis({"name": "Group"})
        chart_income.set_y_axis({"name": "Average Score"})
        ws_income.insert_chart("L2", chart_income, {"x_scale": 1.1, "y_scale": 1.1})

        # =========================================================
        # Alerts
        # =========================================================
        alert_table.to_excel(writer, sheet_name="Alerts", index=False)
        ws_alert = writer.sheets["Alerts"]
        ws_alert.set_zoom(90)
        ws_alert.freeze_panes(1, 1)

        rows_a = len(alert_table)
        cols_a = len(alert_table.columns)
        col_idx_map_a = {col: i for i, col in enumerate(alert_table.columns)}

        ws_alert.set_row(0, 18, header_fmt)
        ws_alert.set_column("A:A", 10, center_fmt)
        ws_alert.set_column("B:B", 10, center_fmt)
        ws_alert.set_column("C:F", 12, num_fmt)
        ws_alert.set_column("G:G", 12, num_fmt)
        ws_alert.set_column("H:H", 14, center_fmt)
        ws_alert.set_column("I:I", 18, center_fmt)

        ws_alert.add_table(
            0, 0, rows_a, cols_a - 1,
            {
                "style": "Table Style Medium 10",
                "columns": [{"header": col} for col in alert_table.columns],
            },
        )

        if "AvgScore" in col_idx_map_a:
            avg_idx = col_idx_map_a["AvgScore"]
            ws_alert.conditional_format(
                1,
                avg_idx,
                rows_a,
                avg_idx,
                {
                    "type": "3_color_scale",
                    "min_color": "#FF0000",
                    "mid_color": "#FFFF00",
                    "max_color": "#00B050",
                },
            )

        # =========================================================
        # Trends
        # =========================================================
        ws_trends = workbook.add_worksheet("Trends")
        ws_trends.set_zoom(90)
        row_offset = 0

        for subj, pivot_df in trend_pivots.items():
            ws_trends.write(row_offset, 0, f"{subj.capitalize()} Trend", title_fmt)
            start_row = row_offset + 2
            start_col = 0

            ws_trends.set_row(start_row, 18, header_fmt)
            for col_idx, col_name in enumerate(pivot_df.columns):
                ws_trends.write(start_row, start_col + col_idx, col_name)

            for r in range(len(pivot_df)):
                for c in range(len(pivot_df.columns)):
                    value = pivot_df.iloc[r, c]
                    if c == 0:
                        ws_trends.write(start_row + 1 + r, start_col + c, value, center_fmt)
                    else:
                        ws_trends.write_number(start_row + 1 + r, start_col + c, value)

            ws_trends.set_column(start_col, start_col, 10, center_fmt)
            ws_trends.set_column(start_col + 1,
                                 start_col + len(pivot_df.columns) - 1,
                                 12,
                                 num_fmt)

            chart_trend = workbook.add_chart({"type": "line"})
            num_rows = len(pivot_df)
            num_cols = len(pivot_df.columns)

            cat_range = ["Trends", start_row + 1, start_col, start_row + num_rows, start_col]

            for i in range(1, num_cols):
                chart_trend.add_series({
                    "name": ["Trends", start_row, start_col + i],
                    "categories": cat_range,
                    "values": [
                        "Trends",
                        start_row + 1,
                        start_col + i,
                        start_row + num_rows,
                        start_col + i,
                    ],
                    "line": {"color": series_color(i - 1)},
                })

            chart_trend.set_title({"name": f"{subj.capitalize()} Average Score Trend"})
            chart_trend.set_x_axis({"name": "Cohort"})
            chart_trend.set_y_axis({"name": "Average Score"})
            ws_trends.insert_chart(start_row, start_col + num_cols + 2,
                                   chart_trend,
                                   {"x_scale": 1.1, "y_scale": 1.1})

            row_offset = start_row + num_rows + 6

    print(f"Dashboard generated: {output_file}")


# ============================================================
# MAIN
# ============================================================
def main() -> None:
    xls = load_excel_file(FILE_NAME)
    if xls is None:
        return

    raw_data_df, raw_fin_df, raw_bm_df = load_track_sheets(xls)
    uni_raw = concat_with_track_labels(raw_data_df, raw_fin_df, raw_bm_df)
    if uni_raw is None:
        return

    uni_df = clean_data(uni_raw)
    data_df = clean_data(raw_data_df)
    fin_df = clean_data(raw_fin_df)
    bm_df = clean_data(raw_bm_df)

    track_counts = get_track_counts(data_df, fin_df, bm_df)
    track_summary = make_track_summary(uni_df)
    pass_fail_df = make_pass_fail_summary(uni_df)
    cohort_summary = summarize_by_cohort(uni_df)
    income_summary = summarize_by_income(uni_df)
    alert_table = generate_performance_alert_table(uni_df, ALERT_THRESHOLD)

    export_cleaned_data(data_df, fin_df, bm_df, uni_df)

    uni_df_all = generate_synthetic_cohorts(uni_df, NEW_COHORT_CANDIDATES)
    trend_pivots = build_trend_pivots(uni_df_all)

    create_dashboard_excel(
        track_counts=track_counts,
        track_summary=track_summary,
        pass_fail_df=pass_fail_df,
        cohort_summary=cohort_summary,
        income_summary=income_summary,
        alert_table=alert_table,
        trend_pivots=trend_pivots,
        output_file=DASHBOARD_FILE,
    )


if __name__ == "__main__":
    main()
