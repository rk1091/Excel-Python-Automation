# Excel-Python-Automation

**A batch autograder for finance coursework — validates dozens of numeric answers across student Excel submissions against expected ranges, and compiles results into a single report.**

## What it does

This script automates grading for a corporate finance assignment (capital budgeting: NPV, IRR, tax, working capital, and cash flow problems). It scans a folder for every student's `.xlsx` submission, checks each required answer cell against an expected tolerance range from the answer key, and appends a pass/fail (1/0) result vector per student into a single consolidated `OUTPUT SHEET.xlsx`. Built for and actively used by finance students/TAs at the University of Texas to replace manual, cell-by-cell grading of dozens of submissions.

## Tech Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel_I%2FO-217346?style=flat&logo=microsoftexcel&logoColor=white)

## Architecture / How it works

1. Uses `glob` to discover every `.xlsx` file in the working directory (excluding the output file)
2. Loads each workbook with `openpyxl` and reads a fixed set of answer cells (revenue, costs, taxes, NWC, operating/non-operating cash flows, IRR, accept/reject decision — across multiple case scenarios in the assignment)
3. Validates each value against a tolerance-banded expected range from the answer key, appending a 1 (correct) or 0 (incorrect) per check
4. Aggregates every student's result row into `OUTPUT SHEET.xlsx` for the grader to review in one place
5. Wraps the run in error handling for common submission issues — malformed/non-numeric cells (`TypeError`), missing files (`FileNotFoundError`), and a locked output file still open in Excel

## Setup & Run

```bash
git clone https://github.com/rk1091/Excel-Python-Automation.git
cd Excel-Python-Automation
pip install -r requirements.txt

# Place all student submission .xlsx files + OUTPUT SHEET.xlsx in this folder
python FINAL_SCRIPT_2.1.py
```
A compiled `.exe` build (`FINAL_SCRIPT_2.1.exe`) is also included for graders without a Python environment set up — just drop it in the submissions folder and run it.

## What I learned / Key challenges

- Working with `openpyxl` to read/write specific cells across many workbooks in a batch pipeline, rather than one file at a time
- Designing tolerance-band validation (ranges, not exact-match) to account for legitimate floating-point/rounding differences in student calculations
- Defensive error handling for real-world messy input — non-numeric cells, missing files, and files locked by the OS — since this ran unattended against submissions I didn't control
- Packaging the script as a standalone `.exe` so non-technical graders could run it without a Python setup

## Notes

The answer-key ranges and cell references are hardcoded to this specific assignment's format — this was built as a practical, one-off automation tool for a real grading workload, not a general-purpose spreadsheet validator. That's a deliberate tradeoff: it optimizes for "solves a real batch-grading problem correctly" over "reusable for any spreadsheet."
