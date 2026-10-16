"""
validators.py
Pure-pandas HR data checks, so the same rules used in the Excel sheet
can run in Python (and be unit-tested) without opening Excel.
"""
import pandas as pd

MIN_BASIC_PCT = 0.40  # policy assumption - change to match your company


def load_data(path: str) -> dict:
    """Read the raw input columns of every sheet (formula columns are recomputed below)."""
    sheets = {}
    for name in ["Employees", "Candidates", "Onboarding", "Attendance", "Payroll_Input"]:
        df = pd.read_excel(path, sheet_name=name)
        df = df.dropna(how="all")
        sheets[name] = df
    # drop rows with no key
    sheets["Employees"] = sheets["Employees"].dropna(subset=["Emp ID"])
    sheets["Candidates"] = sheets["Candidates"].dropna(subset=["Candidate ID"])
    sheets["Onboarding"] = sheets["Onboarding"].dropna(subset=["Emp ID"])
    sheets["Attendance"] = sheets["Attendance"].dropna(subset=["Emp ID"])
    sheets["Payroll_Input"] = sheets["Payroll_Input"].dropna(subset=["Emp ID"])
    return sheets


def add_attendance_pct(att: pd.DataFrame) -> pd.DataFrame:
    att = att.copy()
    att["Attendance %"] = (att["Days Present"] / att["Working Days"]).where(att["Working Days"] > 0, 0)
    return att


def add_onboarding_pct(onb: pd.DataFrame) -> pd.DataFrame:
    onb = onb.copy()
    task_cols = [c for c in onb.columns if c not in ("Emp ID", "Employee Name", "Completion %")]
    onb["Completion %"] = (onb[task_cols] == "Yes").sum(axis=1) / len(task_cols)
    return onb


def validate_payroll(pay: pd.DataFrame, employees: pd.DataFrame,
                     min_basic_pct: float = MIN_BASIC_PCT) -> pd.DataFrame:
    """Return payroll rows with Gross, Net Pay and a 'Validation Check' column."""
    pay = pay.copy()
    for c in ["Basic (INR)", "HRA (INR)", "Allowances (INR)", "Deductions (INR)"]:
        pay[c] = pd.to_numeric(pay[c], errors="coerce")
    pay["Gross (INR)"] = pay[["Basic (INR)", "HRA (INR)", "Allowances (INR)"]].sum(axis=1)
    pay["Net Pay (INR)"] = pay["Gross (INR)"] - pay["Deductions (INR)"].fillna(0)

    known_ids = set(employees["Emp ID"].astype(str))
    dup = pay.duplicated(subset=["Emp ID", "Month"], keep=False)

    def check(row, is_dup):
        if pd.isna(row["Month"]) or pd.isna(row["Basic (INR)"]) or pd.isna(row["Deductions (INR)"]):
            return "Missing data"
        if str(row["Emp ID"]) not in known_ids:
            return "Unknown employee"
        if is_dup:
            return "Duplicate entry"
        if row["Deductions (INR)"] > row["Gross (INR)"]:
            return "Deductions exceed gross"
        if row["Basic (INR)"] < min_basic_pct * row["Gross (INR)"]:
            return "Basic below policy %"
        return "OK"

    pay["Validation Check"] = [check(r, d) for (_, r), d in zip(pay.iterrows(), dup)]
    return pay


def kpis(sheets: dict) -> dict:
    emp, can = sheets["Employees"], sheets["Candidates"]
    active = int((emp["Status"] == "Active").sum())
    resigned = int((emp["Status"] == "Resigned").sum())
    total = active + resigned
    joined = int((can["Stage"] == "Joined").sum())
    return {
        "active": active,
        "resigned": resigned,
        "attrition": resigned / total if total else 0,
        "candidates": len(can),
        "joined": joined,
        "hire_conversion": joined / len(can) if len(can) else 0,
    }
