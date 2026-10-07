"""Clean the synthetic employee dataset with pandas.

Local use:
    pip install pandas openpyxl
    python clean_employees.py

Put employees_data_cleaning.xlsx beside this script.
--json-only exports the same cleaned data for a separate workbook formatter.
"""
from pathlib import Path
import argparse
import json
import pandas as pd


def clean_employees(source):
    df = pd.read_excel(source, sheet_name="Employees")
    input_rows = len(df)

    # Remove extra spaces and standardize spelling/capitalization.
    text_columns = ["Employee_ID", "Name", "City", "Department", "Job_Title"]
    for column in text_columns:
        df[column] = (df[column].astype("string").str.strip()
                      .str.replace(r"\s+", " ", regex=True).replace("", pd.NA))
    df["Employee_ID"] = df["Employee_ID"].str.upper()
    for column in ["Name", "City", "Department", "Job_Title"]:
        df[column] = df[column].str.title()
    df["Department"] = df["Department"].replace({"It": "IT"})
    df["Job_Title"] = (df["Job_Title"].str.replace(r"\bIt\b", "IT", regex=True)
                        .str.replace(r"\bHr\b", "HR", regex=True))

    # Convert mixed salary formats such as "1,250" and "950 JOD".
    salary_text = (df["Salary_JOD"].astype("string").str.upper()
                   .str.replace("JOD", "", regex=False)
                   .str.replace(",", "", regex=False).str.strip())
    df["Salary_JOD"] = pd.to_numeric(salary_text, errors="coerce")
    df["Age"] = pd.to_numeric(df["Age"], errors="coerce")
    df["Hire_Date"] = pd.to_datetime(df["Hire_Date"], format="mixed", errors="coerce")

    # Remove matching records. Never silently choose between conflicting IDs.
    df = df.drop_duplicates().reset_index(drop=True)
    removed = input_rows - len(df)
    if df["Employee_ID"].isna().any() or df["Employee_ID"].duplicated().any():
        raise ValueError("Missing or conflicting employee IDs require review.")

    # Exercise rule: age must be a whole number between 18 and 65.
    valid_age = df["Age"].between(18, 65) & df["Age"].mod(1).eq(0)
    median_age = df.loc[valid_age, "Age"].median()
    if pd.isna(median_age):
        raise ValueError("No valid ages are available to calculate a median.")
    df["Age_Estimated"] = ~valid_age
    df.loc[~valid_age, "Age"] = round(median_age)
    df["Age"] = df["Age"].astype("Int64")

    # Preserve unknown salaries rather than inventing payroll amounts.
    negative_salary = df["Salary_JOD"].lt(0)
    df.loc[negative_salary, "Salary_JOD"] = float("nan")
    df["Salary_JOD"] = df["Salary_JOD"].astype(float)

    # High salary is a review flag, not proof of an error.
    # 5,000 JOD is a review threshold for this synthetic exercise only.
    flags = [[] for _ in range(len(df))]
    def flag(mask, message):
        for index in df.index[mask.fillna(False)]:
            flags[index].append(message)
    flag(negative_salary, "Negative salary removed; verify source")
    flag(df["Salary_JOD"].isna() & ~negative_salary, "Salary missing; verify source")
    flag(df["Salary_JOD"].gt(5000), "High salary retained; verify source")
    flag(df["City"].isna(), "City missing")
    flag(df["Department"].isna(), "Department missing")
    flag(df["Hire_Date"].isna(), "Hire date missing or invalid")
    df["Review_Notes"] = ["; ".join(items) for items in flags]

    report = pd.DataFrame([
        ["Dataset", "Synthetic training data, not real employee records"],
        ["Input records", input_rows],
        ["Exact duplicates removed", removed],
        ["Output employees", len(df)],
        ["Unique employee IDs", df["Employee_ID"].nunique()],
        ["Age estimates", int(df["Age_Estimated"].sum())],
        ["Age estimate used", round(median_age)],
        ["Age policy", "Valid whole ages 18–65; missing/invalid ages use the rounded valid median"],
        ["Age estimates marker", "Age_Estimated = TRUE means estimated, not verified"],
        ["Negative salaries removed", int(negative_salary.sum())],
        ["Salary blanks retained", int(df["Salary_JOD"].isna().sum())],
        ["High salaries awaiting review", int(df["Salary_JOD"].gt(5000).sum())],
        ["Salary policy", "Monthly base salary in JOD; missing amounts are not estimated"],
        ["High salary policy", "Values above 5,000 JOD are retained and flagged for this exercise"],
        ["Missing cities retained", int(df["City"].isna().sum())],
        ["Missing departments retained", int(df["Department"].isna().sum())],
        ["Rows requiring source review", int(df["Review_Notes"].ne("").sum())],
        ["Text cleaning", "Trimmed spaces and standardized case; preserved IT and HR"],
        ["Dates", "Converted to real dates; displayed as yyyy-mm-dd"],
        ["Original file", "Preserved unchanged"],
    ], columns=["Metric", "Result"])
    return df, report


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=Path(__file__).with_name("employees_data_cleaning.xlsx"))
    parser.add_argument("--output", type=Path, default=Path(__file__).with_name("new employees data.xlsx"))
    parser.add_argument("--json-only", type=Path)
    args = parser.parse_args()
    df, report = clean_employees(args.input)
    if args.json_only:
        payload = {"employees": json.loads(df.to_json(orient="split", date_format="iso")),
                   "report": json.loads(report.to_json(orient="split"))}
        args.json_only.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    else:
        # Portable Excel saving for your own computer.
        with pd.ExcelWriter(args.output, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
            df.to_excel(writer, sheet_name="Employees", index=False)
            report.to_excel(writer, sheet_name="Cleaning_Report", index=False)
        print(f"Saved: {args.output}")
    print(report.to_string(index=False))


if __name__ == "__main__":
    main()
