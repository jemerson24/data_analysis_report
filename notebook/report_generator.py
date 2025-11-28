# In[ ]:
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
import os

# In[ ]:
file_name = "student_grades_2027-2028.xlsx"
file_path = os.path.join(os.getcwd(), file_name)

try:
    xls = pd.ExcelFile(file_path)

    # Load selected sheets if they exist
    data_df = pd.read_excel(xls, sheet_name="Data") if "Data" in xls.sheet_names else None
    fin_df = pd.read_excel(xls, sheet_name="Finance") if "Finance" in xls.sheet_names else None
    bm_df = pd.read_excel(xls, sheet_name="BM") if "BM" in xls.sheet_names else None

    # Collect non-empty DFs into a list
    dfs_to_concat = []
    if data_df is not None:
        data_df["Track"] = "Data"
        dfs_to_concat.append(data_df)

    if fin_df is not None:
        fin_df["Track"] = "Finance"
        dfs_to_concat.append(fin_df)

    if bm_df is not None:
        bm_df["Track"] = "BM"
        dfs_to_concat.append(bm_df)

    # Concatenate into one dataframe
    uni_df = pd.concat(dfs_to_concat, ignore_index=True) if dfs_to_concat else None

    print("Sheets loaded:")
    print("Data:", data_df is not None)
    print("Finance:", fin_df is not None)
    print("BM:", bm_df is not None)

except FileNotFoundError:
    print(f"Error: File '{file_name}' not found.")
except Exception as e:
    print("An unexpected error occurred:", e)



# In[ ]:
def clean_data(df):
    df_clean = df.copy()

    # 0. Rename columns from raw Excel headers to standard names (if present)
    rename_map = {
        'StudentID': 'student_id',
        'FirstName': 'first_name',
        'LastName': 'last_name',
        'Class': 'class_name',
        'Term': 'term',
        'Math': 'math',
        'English': 'english',
        'Science': 'science',
        'History': 'history',
        'Attendance (%)': 'attendance',
        'ProjectScore': 'project_score',
        'Passed (Y/N)': 'passed',
        'IncomeStudent': 'income_student',
        'Cohort': 'cohort'
    }

    # Only rename columns that actually exist in the dataframe
    safe_map = {old: new for old, new in rename_map.items() if old in df_clean.columns}
    df_clean = df_clean.rename(columns=safe_map)

    # 1. Normalise column names: strip, lowercase, spaces -> underscores
    df_clean.columns = (
        df_clean.columns
            .str.strip()
            .str.lower()
            .str.replace(" ", "_", regex=False)
    )

    # 2. Define bad string values (lowercase for comparison)
    bad_strings = ["", "n/a", "waived"]

    # 3. Build mask of bad values
    mask = (
        df_clean.isna() |
        df_clean.apply(
            lambda col: col.astype(str)
                          .str.strip()
                          .str.lower()
                          .isin(bad_strings)
        )
    )

    # Keep only rows with NO bad values
    df_clean = df_clean[~mask.any(axis=1)].copy()

    # 4. Check required columns exist after renaming/normalising
    required_cols = ["student_id", "passed", "income_student"]
    missing = [c for c in required_cols if c not in df_clean.columns]
    if missing:
        raise KeyError(
            f"Missing required columns after cleaning: {missing}. "
            f"Current columns: {list(df_clean.columns)}"
        )

    # 5. Convert student_id to integer safely
    df_clean["student_id"] = (
        df_clean["student_id"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)  # handle "123.0" style values
        .astype(int)
    )

    # 6. Convert passed ("Y"/"N") → 1/0
    df_clean["passed"] = df_clean["passed"].replace({"Y": 1, "N": 0})

    # 7. Convert income_student ("True"/"False") → 1/0
    df_clean["income_student"] = df_clean["income_student"].replace(
        {"True": 1, "False": 0, True: 1, False: 0}
    )

    # 8. Columns that must be numeric (after normalisation)
    num_cols_float = ["math", "english", "science", "history", "project_score", "attendance"]
    num_cols_int = ["term", "passed", "income_student"]

    # Only use columns that actually exist
    num_cols_float_existing = [c for c in num_cols_float if c in df_clean.columns]
    num_cols_int_existing = [c for c in num_cols_int if c in df_clean.columns]

    # First: convert to numeric, coercing bad values to NaN
    df_clean[num_cols_float_existing + num_cols_int_existing] = df_clean[
        num_cols_float_existing + num_cols_int_existing
    ].apply(lambda col: pd.to_numeric(col, errors="coerce"))

    # Drop any rows that still have NaNs in numeric columns
    df_clean = df_clean.dropna(subset=num_cols_float_existing + num_cols_int_existing)

    # Cast ints (no NaNs left)
    if num_cols_int_existing:
        df_clean[num_cols_int_existing] = df_clean[num_cols_int_existing].astype(int)

    # Cast score columns (math, science, etc.) to int, keeping attendance as float
    score_cols = [col for col in num_cols_float_existing if col != "attendance"]
    if score_cols:
        df_clean[score_cols] = df_clean[score_cols].astype(int)
        
    # Set index at the end
    df_clean = df_clean.set_index("student_id")

    return df_clean


# In[ ]:
uni_df = clean_data(uni_df)
data_df = clean_data(data_df)
fin_df = clean_data(fin_df)
bm_df = clean_data(bm_df)

# In[ ]:
# Count cohort sizes & set variables 
num_d = len(data_df)
num_f = len(fin_df)
num_bm = len(bm_df)
num_students = [num_d, num_f, num_bm]
labels = ["Data", "Finance", "Business"]
print(f"There are {num_d} students in the Data track, {num_f} students in the Finance track, and {num_bm} students in the Business track.")

# In[ ]:
# Plot pie chart
plt.pie(num_students, 
        labels=labels, 
        autopct='%1.1f%%', 
        startangle=90, 
        shadow=True, 
        colors=['#ff9999','#66b3ff','#99ff99']
        )
plt.title('Major Distribution')
plt.show()

# In[ ]:
track_summary = uni_df.groupby("track", as_index=False).agg({
    "math": "mean",
    "english": "mean",
    "science": "mean",
    "history": "mean",
    "attendance": "mean",
    "project_score": "mean",
})

print(track_summary)


# In[ ]:
passed = uni_df["passed"].value_counts()[1]
failed = uni_df["passed"].value_counts()[0]
pf_ratio = [passed, failed]
labels = ["Pass", "Fail"]

# In[ ]:
# Plot pie chart
plt.pie(pf_ratio,
        labels=labels, 
        autopct='%1.1f%%', 
        startangle=90, 
        shadow=True, 
        colors=['#99ff99','#ff9999']
        )
plt.title('Pass/Fail Ratio')
plt.show()
print(f"There are {round(passed/(failed+passed)*100, 2)}% students who passed and {round(failed/(failed+passed)*100, 2)}% students who failed.")

# In[ ]:
plt.figure()

tracks = [data_df, fin_df, bm_df]
labels = ["Data", "Finance", "BM"]

for track, label in zip(tracks, labels):
    plt.hist(track["history"], bins=10, alpha=0.5, label=label)

plt.legend()
plt.title("History Grade Distribution (Histogram)")
plt.xlabel("History Score")
plt.ylabel("Frequency")
plt.show()

# In[ ]:
plt.figure()

plt.boxplot(
    [data_df["history"], fin_df["history"], bm_df["history"]],
    labels=["Data", "Finance", "BM"]
)

plt.title("History Scores by Track (Boxplot)")
plt.ylabel("Scores")
plt.show()


# In[ ]:
temp = pd.DataFrame({
    "Track": ["Data"] * len(data_df) +
             ["Finance"] * len(fin_df) +
             ["BM"] * len(bm_df),
    "Math": pd.concat([data_df["math"], fin_df["math"], bm_df["math"]], axis=0)
})

# Group, then reorder the index
avg_math = (
    temp.groupby("Track")["Math"]
        .mean()
        .reindex(["Data", "Finance", "BM"])   # <-- force the order
)

plt.figure()
avg_math.plot(kind="bar", color=["#ff9999", "#66b3ff", "#99ff99"])
plt.title("Average Mathematics Score by Track")
plt.ylabel("Average Math Score")
plt.show()


# In[ ]:
tracks = {
    "Data Science": data_df,
    "Finance": fin_df,
    "Business": bm_df
}

colors = {
    "Data Science": "#ff9999",
    "Finance": "#66b3ff",
    "Business": "#99ff99"
}

plt.figure(figsize=(15, 4))

for i, (name, df) in enumerate(tracks.items(), 1):

    plt.subplot(1, 3, i)
    
    # Scatter
    plt.scatter(df["attendance"], df["project_score"],
                alpha=0.6, color=colors[name])
    
    # Regression line
    m, b = np.polyfit(df["attendance"], df["project_score"], 1)
    x = np.linspace(df["attendance"].min(), df["attendance"].max(), 100)
    plt.plot(x, m*x + b, color=colors[name], linewidth=2)
    
    # Correlation
    r = df[["attendance", "project_score"]].corr().iloc[0, 1]
    
    # Display correlation on the subplot
    plt.text(
        0.05, 0.95,
        f"r = {r:.2f}",
        transform=plt.gca().transAxes,
        fontsize=12,
        fontweight="bold",
        color="black",
        verticalalignment="top"
    )
    
    plt.title(name)
    plt.xlabel("Attendance")
    plt.ylabel("Project Score")

plt.tight_layout()
plt.show()


# In[ ]:
cohort_summary = uni_df.groupby("cohort").agg(
    avg_math=("math", "mean"),
    avg_english=("english", "mean"),
    avg_science=("science", "mean"),
    avg_history=("history", "mean"),
    avg_attendance=("attendance", "mean"),
    avg_project=("project_score", "mean"),
    pass_rate=("passed", "mean"),   # passed already converted to 0/1
    count=("cohort", "size")
).reset_index()

print("\nCohort-Level Summary:")
print(cohort_summary)


# In[ ]:
plt.figure()
plt.bar(cohort_summary["cohort"], cohort_summary["pass_rate"] * 100)
plt.title("Pass Rate by Cohort (%)")
plt.xlabel("Cohort")
plt.ylabel("Pass Rate (%)")
plt.ylim(80, 100)
plt.xticks(rotation=0)
plt.show()


# In[ ]:
income_summary = uni_df.groupby("income_student").agg(
    avg_math=("math", "mean"),
    avg_english=("english", "mean"),
    avg_science=("science", "mean"),
    avg_history=("history", "mean"),
    avg_attendance=("attendance", "mean"),
    avg_score=("project_score", "mean"),
    pass_rate=("passed", "mean"),
    count=("income_student", "size")
).reset_index()

# Map 1/0 → labels
income_summary["Group"] = income_summary["income_student"].map({
    1: "Income Supported",
    0: "Not Supported"
})

print("\nIncome Student Performance Comparison:")
print(income_summary)



# In[ ]:
avg_academic = income_summary[[
    "avg_math", "avg_english", "avg_science", "avg_history"
]].mean(axis=1)

plt.figure()
plt.bar(income_summary["Group"], avg_academic, color=["#66b3ff", "#ff9999"])
plt.title("Average Academic Performance by Income Group")
plt.ylabel("Average Grade")
plt.ylim(50, 75)
plt.show()


# In[ ]:
# Export each cleaned dataframe to CSV
data_df.to_csv("data_df_cleaned.csv", index=False)
fin_df.to_csv("fin_df_cleaned.csv", index=False)
bm_df.to_csv("bm_df_cleaned.csv", index=False)
uni_df.to_csv("uni_df_cleaned.csv", index=False)

print("CSV files exported successfully!")


# In[ ]:
# Track-level summary
track_summary = (
    uni_df
    .groupby("track")
    .agg({
        "math": "mean",
        "english": "mean",
        "science": "mean",
        "history": "mean",
        "attendance": "mean",
        "project_score": "mean"
    })
    .reset_index()
)

# Cohort-level summary
cohort_summary = (
    uni_df
    .groupby("cohort")
    .agg({
        "math": "mean",
        "english": "mean",
        "science": "mean",
        "history": "mean",
        "attendance": "mean",
        "project_score": "mean"
    })
    .reset_index()
)

# In[ ]:

def create_dashboard_excel(track_df, cohort_df, output_file='Final_Dashboard.xlsx'):

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # 1. input summary sheet
        track_df.to_excel(writer, sheet_name='track_Summary', index=False)
        cohort_df.to_excel(writer, sheet_name='Cohort_Summary', index=False)

        # 2. workbook / worksheet
        workbook = writer.book
        worksheet_track = writer.sheets['track_Summary']
        worksheet_cohort = writer.sheets['Cohort_Summary']

        # 3. setting the format
        header_fmt = workbook.add_format({
            'bold': True,
            'bg_color': '#D7E4BC',
            'border': 1
        })
        num_fmt = workbook.add_format({'num_format': '0.00'})

        worksheet_track.set_column('B:G', 12, num_fmt)
        worksheet_track.set_row(0, None, header_fmt)

        worksheet_cohort.set_column('B:G', 12, num_fmt)
        worksheet_cohort.set_row(0, None, header_fmt)

        # =========================================================
        # 4. track_Summary plot
        # =========================================================
        chart_track = workbook.add_chart({'type': 'column'})

        max_row_track = len(track_df) + 1 


        # Math
        chart_track.add_series({
            'name':       ['track_Summary', 0, 1],
            'categories': ['track_Summary', 1, 0, max_row_track - 1, 0],
            'values':     ['track_Summary', 1, 1, max_row_track - 1, 1],
        })

        # English
        chart_track.add_series({
            'name':       ['track_Summary', 0, 2],
            'categories': ['track_Summary', 1, 0, max_row_track - 1, 0],
            'values':     ['track_Summary', 1, 2, max_row_track - 1, 2],
        })

        # Science
        chart_track.add_series({
            'name':       ['track_Summary', 0, 3],
            'categories': ['track_Summary', 1, 0, max_row_track - 1, 0],
            'values':     ['track_Summary', 1, 3, max_row_track - 1, 3],
        })

        # History
        chart_track.add_series({
            'name':       ['track_Summary', 0, 4],
            'categories': ['track_Summary', 1, 0, max_row_track - 1, 0],
            'values':     ['track_Summary', 1, 4, max_row_track - 1, 4],
        })

        chart_track.set_title({'name': 'Average Scores by track'})
        chart_track.set_x_axis({'name': 'track'})
        chart_track.set_y_axis({'name': 'Score'})

        worksheet_track.insert_chart('I2', chart_track)

        # =========================================================
        # 5. Cohort_Summary plot
        # =========================================================
        chart_cohort = workbook.add_chart({'type': 'column'})

        max_row_cohort = len(cohort_df) + 1  

        # Math
        chart_cohort.add_series({
            'name':       ['Cohort_Summary', 0, 1],
            'categories': ['Cohort_Summary', 1, 0, max_row_cohort - 1, 0],
            'values':     ['Cohort_Summary', 1, 1, max_row_cohort - 1, 1],
        })

        # English
        chart_cohort.add_series({
            'name':       ['Cohort_Summary', 0, 2],
            'categories': ['Cohort_Summary', 1, 0, max_row_cohort - 1, 0],
            'values':     ['Cohort_Summary', 1, 2, max_row_cohort - 1, 2],
        })

        # Science
        chart_cohort.add_series({
            'name':       ['Cohort_Summary', 0, 3],
            'categories': ['Cohort_Summary', 1, 0, max_row_cohort - 1, 0],
            'values':     ['Cohort_Summary', 1, 3, max_row_cohort - 1, 3],
        })

        # History
        chart_cohort.add_series({
            'name':       ['Cohort_Summary', 0, 4],
            'categories': ['Cohort_Summary', 1, 0, max_row_cohort - 1, 0],
            'values':     ['Cohort_Summary', 1, 4, max_row_cohort - 1, 4],
        })

        chart_cohort.set_title({'name': 'Average Scores by Cohort'})
        chart_cohort.set_x_axis({'name': 'Cohort'})
        chart_cohort.set_y_axis({'name': 'Score'})
        
        worksheet_cohort.insert_chart('I2', chart_cohort)

    print(f"Dashboard generated: {output_file}")


# In[ ]:
create_dashboard_excel(track_summary, cohort_summary, output_file='Final_Dashboard.xlsx')

# In[ ]:
# ============================================================
# Includes:
# A. Threshold-based Alert Table
# B. Heatmap Visualization
# C. Text Alert Report
# ============================================================
df_alert = uni_df.copy()

subject_cols = ["math", "english", "science", "history"]

# ---------------------------------------------
# STEP A: alert threshold settings
# ---------------------------------------------
alert_threshold = 60  # we can change the threshold here

agg = (
    df_alert
    .groupby(["track", "cohort"])[subject_cols]
    .mean()
    .reset_index()
)

agg["AvgScore"] = agg[subject_cols].mean(axis=1)
agg["LowPerformance"] = agg["AvgScore"] < alert_threshold
agg["Alert"] = agg["LowPerformance"].apply(lambda x: "⚠ LOW PERFORMANCE" if x else "OK")

print("===== Performance Alert Table =====")
display(agg.sort_values(["track", "cohort"]))


# ---------------------------------------------
# STEP B: performance heating map
# ---------------------------------------------
plt.figure(figsize=(10, 6))
pivot = agg.pivot(index="cohort", columns="track", values="AvgScore")
sns.heatmap(pivot, annot=True, fmt=".1f", cmap="RdYlGn", center=70)
plt.title("Performance Heatmap (Average Score)")
plt.ylabel("Cohort")
plt.xlabel("track")
plt.show()


# ---------------------------------------------
# STEP C: text alarm
# ---------------------------------------------
alerts = agg[agg["LowPerformance"] == True]

print("\n===== Automated Text Alert Report =====\n")

if len(alerts) == 0:
    print("No alerts. All tracks and cohorts meet performance standards.")
else:
    for _, row in alerts.iterrows():
        print(
            f"• {row['track']} track in cohort {row['cohort']} "
            f"is low-performing (Average Score = {row['AvgScore']:.1f})"
        )


# In[ ]:
base_df = uni_df.reset_index().copy()


existing_cohorts = set(base_df["cohort"].astype(str).unique())
candidate_new = ["23-24", "24-25", "25-26"]  
new_cohorts = [c for c in candidate_new if c not in existing_cohorts]

print("Existing cohorts:", existing_cohorts)
print("New cohorts to generate:", new_cohorts)

track_sizes = base_df.groupby("track")["student_id"].count().to_dict()
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

       
        for col in ["math", "english", "science", "history"]:
            sampled[col] = (
                sampled[col].astype(float)
                + np.random.normal(loc=0, scale=5, size=len(sampled))  
            ).clip(0, 100).round().astype(int)  

        synthetic_rows.append(sampled)

# combine
synthetic_df = pd.concat(synthetic_rows, ignore_index=True)
uni_df_all = pd.concat([base_df, synthetic_df], ignore_index=True)
uni_df_all = uni_df_all.set_index("student_id")
print("Cohort counts in extended dataset:")
print(uni_df_all["cohort"].value_counts().sort_index())


# In[ ]:

uni_df_all = uni_df_all.copy()
uni_df_all["cohort"] = uni_df_all["cohort"].astype(str)

cohort_order = sorted(
    uni_df_all["cohort"].unique(),
    key=lambda x: int(x.split("-")[0])
)
uni_df_all["cohort"] = pd.Categorical(
    uni_df_all["cohort"],
    categories=cohort_order,
    ordered=True
)

subjects = ["math", "english", "science", "history"]

for subj in subjects:
    plt.figure(figsize=(8, 5))

    mean_scores = (
        uni_df_all
        .groupby(["cohort", "track"])[subj]
        .mean()
        .reset_index()
    )

    for track in mean_scores["track"].unique():
        track_data = mean_scores[mean_scores["track"] == track]
        plt.plot(
            track_data["cohort"],
            track_data[subj],
            marker="o",
            label=track
        )

    plt.title(f"{subj.capitalize()} Average Score Trend by Cohort and track")
    plt.xlabel("Cohort")
    plt.ylabel("Average Score")
    plt.ylim(0, 100)
    plt.legend(title="track")
    plt.grid(True)
    plt.tight_layout()
    plt.show()


# In[ ]:


