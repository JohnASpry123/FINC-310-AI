# FINC 310 — Procter & Gamble Financial Story Dashboard Package

## Purpose

This is a complete pilot package for an AI-enabled FINC 310 group assignment using The Procter & Gamble Company’s audited fiscal 2025 Form 10-K. Students use a standardized Excel workbook, AI, and Python to create and audit four complementary financial visualizations:

1. Income statement Sankey
2. Free cash flow waterfall
3. Five-year net sales and GAAP diluted EPS trend
4. DuPont ROE decomposition

The instructional emphasis is verification: AI may draft code, but students must reconcile every number to the provided data, understand the program, document assumptions and limitations in REM comments, and correct weaknesses or errors.

## Data scope and conventions

- Fiscal year: year ended June 30, 2025
- Units: USD millions, except per-share amounts
- FY2023–FY2025 financial statement values: P&G fiscal 2025 Form 10-K
- FY2021–FY2022 net sales and GAAP diluted EPS: P&G fiscal 2022 Form 10-K
- P&G adjusted free cash flow is explicitly labeled as a company-defined non-GAAP measure.
- Textbook FCF and cash FCF are shown separately and reconciled rather than treated as interchangeable.
- The DuPont analysis uses average fiscal 2024/fiscal 2025 assets and total shareholders’ equity.

## Folder and file map

### Files to post for students

- `PG_Student_Group_Assignment.docx` — complete student instructions
- `PG_FinancialStory_Student.xlsx` — standardized workbook with audited input data and blank student calculation/output areas
- `PG_Model_AI_Prompt_and_REM_Guide.docx` — model prompt, REM template, AI audit questions, and code-understanding worksheet
- `PG_Grading_Rubric.docx` — 100-point analytic rubric
- `Source_Files/PG_FY2025_Annual_Report_10-K.pdf` — audited source report
- `requirements.txt` — Python packages required

### Canvas setup file

- `PG_Canvas_Assignment_Copy.docx` — assignment text ready to paste into Canvas and customize with dates

### Instructor-only files

- `PG_Instructor_Solution_Manual.docx` — calculations, interpretations, answer key, oral questions, and grading notes
- `PG_FinancialStory_Instructor.xlsx` — completed workbook with formulas, 30 named ranges, source citations, color conventions, verification checks, and embedded dashboard images
- `pg_financial_story_solution.py` — executable, commented Python solution with REM sections and automatic sanity checks
- `PG_Sanity_Checks.csv` — expected validation results; all six checks pass
- `PG_01_Income_Statement_Sankey.png`
- `PG_02_Free_Cash_Flow_Waterfall.png`
- `PG_03_Five_Year_Revenue_EPS_Trend.png`
- `PG_04_DuPont_ROE_Decomposition.png`
- `PG_Financial_Story_Dashboard.png`
- `Source_Files/PG_FY2025_IR_Financial_Statements.xlsx` — original P&G investor-relations spreadsheet

## Python installation and run commands

From a terminal in the package folder:

```bash
python -m pip install -r requirements.txt
python pg_financial_story_solution.py \
  --workbook PG_FinancialStory_Student.xlsx \
  --output PG_Output
```

Expected files in `PG_Output`:

- `PG_01_Income_Statement_Sankey.png`
- `PG_02_Free_Cash_Flow_Waterfall.png`
- `PG_03_Five_Year_Revenue_EPS_Trend.png`
- `PG_04_DuPont_ROE_Decomposition.png`
- `PG_Financial_Story_Dashboard.png`
- `PG_Sanity_Checks.csv`

The program stops with an error if a required accounting or finance identity fails.

## Instructor validation benchmarks

All figures are in USD millions unless indicated otherwise.

| Benchmark | Expected result |
|---|---:|
| FY2025 net sales | 84,284 |
| FY2025 operating income (EBIT) | 20,451 |
| FY2025 net earnings | 16,065 |
| FY2025 cash flow from operations | 17,817 |
| FY2025 capital expenditures | 3,773 |
| P&G adjusted free cash flow | 14,606 |
| Simplified DuPont ROE | 31.24% |
| FY2021–FY2025 net sales CAGR | 2.58% |
| FY2021–FY2025 diluted EPS CAGR | 4.30% |

## Recommended Canvas release sequence

1. Post the annual report, student workbook, student assignment, prompt/REM guide, rubric, and software instructions.
2. Keep the instructor workbook, instructor manual, Python solution, completed figures, and sanity-check key unpublished.
3. Require one group submission but reserve the right to adjust individual grades based on oral code/formula questions.
4. Require every team to submit its exact AI prompt history, REM block, annotated code, verification log, and at least one meaningful human revision.

## Quality-control status

- Python solution executed successfully against the student workbook.
- Six accounting and finance sanity checks returned `PASS`.
- Instructor workbook formulas were checked for common Excel errors.
- Workbook-defined names, embedded images, and source links were verified.
- Every page of each Word document was rendered and visually reviewed.
