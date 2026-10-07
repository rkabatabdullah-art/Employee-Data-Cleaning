# Employee Data Cleaning with Python and Pandas

A portfolio project that cleans an Excel employee dataset, documents the cleaning decisions, and exports a new workbook with a cleaning report.

The source contains **130 records representing 124 unique employees**, including six duplicate rows. All names, salaries, and other employee details are **synthetic training data**, not real employee records or client work. Salaries represent monthly base pay in Jordanian dinars (JOD).

## Project files

| File | Purpose |
| --- | --- |
| `employees_data_cleaning.xlsx` | Original dataset with intentional data quality issues |
| `clean_employees.py` | Python cleaning script |
| `new_employees_data.xlsx` | Cleaned dataset and cleaning report |
| `README.md` | Project overview, results, and usage instructions |

## Tools

- Python
- Pandas for reading, cleaning, and checking the data
- OpenPyXL for Excel input/output in the standalone script
- Excel workbooks for the source and deliverable

## Data quality issues

The input includes duplicate records, missing ages and salaries, inconsistent capitalization, extra spaces, invalid ages, a negative salary, and an unusually high salary. Some salaries contain currency labels or commas, and some hire dates are stored as text.

## Cleaning process

1. Trim leading and trailing spaces and collapse repeated spaces in text fields.
2. Standardize employee IDs, names, cities, departments, and job titles while preserving abbreviations such as `IT` and `HR`.
3. Remove `JOD` and thousands separators from salary text, convert ages and salaries to numbers, and parse hire dates.
4. Remove duplicate records after standardization. Stop if employee IDs are missing or still duplicated, rather than silently choosing between conflicting records.
5. Replace missing or invalid ages with the rounded median of valid ages. Mark every estimate using `Age_Estimated`.
6. Replace the negative salary with a blank. Keep missing salaries, cities, and departments blank and document them in `Review_Notes`.
7. Retain unusually high salaries and flag them for source verification.
8. Export the cleaned employee data and a report to a separate workbook, preserving the original file.

## Results

| Measure | Result |
| --- | ---: |
| Original records | 130 |
| Duplicate rows removed | 6 |
| Final employee records | 124 |
| Unique employee IDs | 124 |
| Missing or invalid ages estimated | 11 |
| Median age used for estimates | 37 years |
| Negative salary values removed | 1 |
| Salary values left blank | 7 |
| High salary values flagged | 1 |
| City values left blank | 4 |
| Department values left blank | 3 |
| Employees with source-review notes | 9 |

The 11 age estimates cover eight missing ages and three invalid ages. The seven blank salaries include six originally missing salaries and the removed negative salary. Review categories overlap, so their counts should not be added to calculate the number of affected employees.

## Cleaning decisions and limitations

- **Age:** Whole-number ages from 18 to 65 are accepted for this exercise. Estimated ages are not verified personal information. `Age_Estimated = TRUE` identifies these records separately from source-review notes.
- **Salary:** Missing amounts are not estimated because that would invent payroll figures. The salary of **85,000 JOD** is retained and flagged, not silently corrected.
- **Review threshold:** Salaries above **5,000 JOD** trigger a review flag in this exercise. This is a training rule, not a market benchmark or proof of an error.
- **Missing categories:** Cities and departments remain blank when their values are unknown.

The output is standardized and deduplicated, but unresolved values still require source verification. It should not be presented as a fully verified payroll dataset.

## How to run

Place `clean_employees.py` and `employees_data_cleaning.xlsx` in the same folder. Open a terminal in that folder.

Install the dependencies:

```bash
python -m pip install pandas openpyxl
```

Run the script:

```bash
python clean_employees.py --output new_employees_data.xlsx
```

The output contains two worksheets:

- **Employees:** 124 cleaned employee records, including `Age_Estimated` and `Review_Notes`.
- **Cleaning_Report:** Counts, cleaning rules, and remaining issues.

The explicit output argument matches the sample filename above. Running without it saves `new employees data.xlsx` beside the script. Re-running overwrites the chosen output file. The script reproduces the cleaned data and report; visual formatting may differ from the formatted sample workbook.

## Skills demonstrated

Data inspection, text normalization, duplicate removal, numeric and date conversion, missing-value handling, documented imputation, outlier review, and Excel export.
