# HR Analytics & Employee Records Dashboard

An HR operations project: employee records, hiring pipeline, onboarding tracker,
attendance, and payroll input validation, with an Excel workbook and a Python (Streamlit) dashboard.

## Files
| File | Purpose |
|---|---|
| `generate_workbook.py` | Builds `HR_Dashboard.xlsx` (sheets, drop-downs, formulas, charts, sample data) |
| `HR_Dashboard.xlsx` | Ready-to-use workbook (Google Sheets: File > Import) |
| `validators.py` | Python versions of the checks (payroll rules, attendance %, onboarding %, KPIs) |
| `app.py` | Streamlit dashboard reading the workbook |
| `test_validators.py` | Simple tests for the rules |

## Run it
```bash
pip install -r requirements.txt
python generate_workbook.py      # optional: rebuilds the workbook
python test_validators.py        # checks the rules
streamlit run app.py             # opens the dashboard in your browser
```
Note: if you rebuild with `generate_workbook.py`, open and save it once in Excel/LibreOffice
so formula values are calculated before the Streamlit app reads it.

## Payroll validation rules
Missing data, unknown employee, duplicate (Emp ID + Month), deductions above gross,
basic below the policy % of gross (default 40%, an assumption - set your own).

## Notes
- All data is sample data. Last 4 payroll rows are deliberate test cases.
- Do not store PAN, Aadhaar or bank numbers in this file.
