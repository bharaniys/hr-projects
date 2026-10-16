"""Run:  python test_validators.py   (or pytest)"""
from validators import load_data, validate_payroll, add_attendance_pct, add_onboarding_pct, kpis

data = load_data("HR_Dashboard.xlsx")
pay = validate_payroll(data["Payroll_Input"], data["Employees"])
counts = pay["Validation Check"].value_counts().to_dict()


def test_payroll_rules():
    assert counts["OK"] == 9
    assert counts["Duplicate entry"] == 2
    assert counts["Unknown employee"] == 1
    assert counts["Deductions exceed gross"] == 1
    assert counts["Missing data"] == 1


def test_kpis():
    k = kpis(data)
    assert k["active"] == 10 and k["resigned"] == 2
    assert round(k["attrition"], 3) == 0.167


def test_pct_columns():
    assert 0 <= add_attendance_pct(data["Attendance"])["Attendance %"].max() <= 1
    assert round(add_onboarding_pct(data["Onboarding"])["Completion %"].mean(), 2) == 0.76


if __name__ == "__main__":
    test_payroll_rules(); test_kpis(); test_pct_columns()
    print("All checks passed.", counts)
