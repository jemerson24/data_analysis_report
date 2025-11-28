# Student Performance Analysis

This project analyzes student performance data from three academic tracks (Data, Finance, BM).  
The notebook cleans and merges the raw Excel sheets, runs multiple analysis modules, generates visualizations, and exports an Excel dashboard.  

---

## 📌 Features

- **Data Cleaning**  
  Handles missing values, standardizes column formats, converts types, and unifies boolean fields.

- **Track-Level Analysis**  
  Total number of students per track, average subject scores, attendance, project score comparison, and pass/fail ratios.

- **Cross-Track Comparison**  
  Visualization of specific scores from various tracks and attendance–project correlations for each track.

- **Cohort Analysis**  
  Cohort-level performance, pass rates, and comparison of income-supported vs non-supported students.

- **Trend Analysis**  
  Synthetic cohorts are generated to show multi-year performance trends across all tracks.

- **Performance Alerts**  
  Flags low-performing Track–Cohort pairs, prints text alerts, and displays a heatmap.

- **Excel Dashboard**  
  Generates `Final_Dashboard.xlsx` using XlsxWriter, with formatted tables and embedded charts.

- **CLI Menu**  
  Run all analyses from the terminal without using Jupyter Notebook.

---

## 🚀 Quickstart

This analysis is presented both as a Jupyter Notebook and a Python script. If you decide to run the analysis in Jupter Notebook:
- Ensure student_grades_2027-2028.xlsx is in the root folder.
- Open data_analysis_report.ipynb.
- Click "Run All".
- Check the generated Final_Dashboard.xlsx and images folder.

Install dependencies:

```bash
pip install pandas numpy matplotlib seaborn xlsxwriter openpyxl


📁 Project Structure
├── data_analysis_report.ipynb        # Analysis source code
├── report_generator.py        # Analysis source code
├── student_grades_2027-2028.xlsx     # Input dataset
│
├── Data Files
│   └── uni_df_cleaned.csv        # Cleaned merged dataset
    └── bm_df_cleaned.csv 
    └── data_df_cleaned.csv
    └── fin_df_cleaned.csv       
│
├── Excel Reports
│   ├── Summary_Statistics_Report.xlsx # Report with charts
│   └── Final_Dashboard.xlsx           # Dashboard-style report
│
└── Generated Visualizations
│   ├── fig_track_major_distribution.png
│   ├── fig_track_pass_fail.png
│   ├── fig_history_grade_distribution.png
│   ├── fig_history_scores_by_track.png
│   ├── fig_average_math_scores_by_track.png
│   ├── fig_att_scores_correlations.png
│   ├── fig_pass_rate_by_cohort.png
│   └── fig_avg_performance_by_income_group.png
│
└── README.md
```

## 📊 How the Analysis Works

### 1. Data Loading and Cleaning
- Loads all Excel sheets (Data, Finance, BM)
- Renames columns for consistency
- Converts numeric and boolean fields
- Removes invalid rows
- Produces a unified dataset `uni_df`

### 2. Summary Statistics
- Track-level averages  
- Pass/fail distribution  
- Basic visualizations (pie charts, bar plots)

### 3. Cross-Track Analysis
- History distribution (histogram + boxplot)
- Math comparison
- Attendance vs project score correlation per track

### 4. Cohort Analysis
- Cohort-level averages
- Pass rates by cohort
- Income-supported student performance

### 5. Trend Analysis
- Synthetic additional cohorts (e.g., 23–24, 24–25, 25–26)
- Produces four trend plots (Math/English/Science/History)

### 6. Performance Alerts
- Threshold-based alert table
- Heatmap visualization
- Text alert report
- “LOW” marks added to trend charts

### 7. Excel Dashboard
- Track summary sheet
- Cohort summary sheet
- Embedded bar charts
- Styled headers and numeric formatting

## 📤 Outputs

Running the notebook or CLI will generate:

### Cleaned dataset
- `uni_df_cleaned.csv`
- `bm_df_cleaned.csv`
- `fin_df_cleaned.csv`
- `data_df_cleaned.csv`

### Excel Dashboard
- `Final_Dashboard.xlsx`  
  Contains formatted summary tables and embedded charts for track-level and cohort-level analysis.

### Summary Statistics Report
- `Summary_Statistics_Report.xlsx`  
  Consolidates Track-Level Summary Statistics, Cross-Track Comparative Analysis, and Cohort-Level Analysis into one Excel file, including all related data tables and charts.

### Plot Images (used for embedding charts)
These PNG files are automatically generated when running the notebook:
- `fig_track_major_distribution.png  `
- `fig_track_avg_math.png`  
- `fig_track_pass_fail.png`  
- `fig_history_grade_distribution.png`  
- `fig_history_scores_by_track.png`  
- `fig_average_math_scores_by_track.png`  
- `fig_att_scores_correlations.png`  
- `fig_cohort_pass_rate.png`  
- `fig_cohort_income_group.png`

## 📝 Notes

The synthetic cohort data is used only for trend visualization and includes mock data.

The project assumes consistent sheet names: Data, Finance, BM.