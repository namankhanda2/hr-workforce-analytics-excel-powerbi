"""
Generate a realistic synthetic HR workforce dataset.

Two outputs:
  1. data/raw/hr_employees_raw.csv   -> messy, as collected (blanks, dupes, typos)
  2. data/hr_cleaned.csv             -> cleaned flat dataset
  3. data/powerbi/*.csv              -> star-schema tables for Power BI

Run:  .venv/bin/python generate_hr_data.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(2026)

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "data" / "raw"
PBI_DIR = BASE / "data" / "powerbi"

SNAPSHOT = date(2025, 12, 31)

DEPARTMENTS = {
    "Executive": ["CEO", "CFO", "COO"],
    "Finance": ["Financial Analyst", "Accountant", "Credit Analyst"],
    "Human Resources": ["HR Generalist", "Recruiter", "HR Analyst"],
    "IT": ["Software Engineer", "Data Analyst", "DevOps Engineer", "IT Support"],
    "Operations": ["Operations Analyst", "Logistics Coordinator", "Quality Analyst"],
    "Marketing": ["Marketing Analyst", "Content Writer", "SEO Specialist"],
    "Sales": ["Sales Executive", "Account Executive", "Business Development"],
    "Customer Support": ["Support Agent", "Support Lead", "Escalation Specialist"],
    "Research & Development": ["Research Scientist", "Data Scientist", "Product Engineer"],
}

EDUCATION = ["High School", "Diploma", "Bachelor", "Master", "PhD"]
OFFICES = ["Delhi", "Mumbai", "Bengaluru", "Hyderabad", "Chennai", "Remote"]
EMPLOYMENT = ["Full-Time", "Part-Time", "Contract"]
EXIT_REASONS = {
    0.55: "Resigned",
    0.20: "Retirement",
    0.10: "Performance / Termination",
    0.08: "Layoff",
    0.07: "Relocation",
}

BASE_SALARY_MULT = {
    "Executive": 3.2, "Finance": 1.4, "Human Resources": 1.2, "IT": 1.8,
    "Operations": 1.2, "Marketing": 1.1, "Sales": 1.3,
    "Customer Support": 0.9, "Research & Development": 2.0,
}

DEP_ATTRITION_WEIGHT = {
    "Executive": 0.02, "Finance": 0.15, "Human Resources": 0.13, "IT": 0.17,
    "Operations": 0.24, "Marketing": 0.20, "Sales": 0.30,
    "Customer Support": 0.34, "Research & Development": 0.14,
}


def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))


def make_employee(emp_id: int) -> dict:
    hire_date = random_date(date(2020, 1, 1), date(2025, 12, 31))
    department = random.choices(list(DEPARTMENTS.keys()),
                                weights=[1, 9, 6, 12, 14, 8, 13, 12, 8], k=1)[0]
    job_role = random.choice(DEPARTMENTS[department])
    job_level = random.choices(range(1, 8), weights=[22, 24, 20, 14, 9, 6, 5], k=1)[0]
    age = max(21, 24 + job_level * 2 + random.randint(-3, 8))
    tenure_months = max(1, (SNAPSHOT - hire_date).days // 30)
    performance = max(1, min(5, random.choices([1, 2, 3, 4, 5],
                                               weights=[2, 8, 35, 35, 20], k=1)[0]))
    satisfaction = max(1, min(10, int(random.gauss(6.5, 1.8))))

    # Attrition decision (keeps enough churn for a juicy analysis)
    attrition_prob = min(
        0.80,
        0.55 * (
            DEP_ATTRITION_WEIGHT[department]
            + (0.25 if job_level <= 2 else 0)
            + (0.15 if tenure_months < 18 else 0)
            + (0.08 if satisfaction < 4 else 0)
            - (0.03 if performance >= 4 else 0)
        ),
    )
    is_active = random.random() >= attrition_prob

    base = 52000 + job_level * 9000
    salary = int(base * BASE_SALARY_MULT[department] * random.uniform(0.75, 1.35))
    salary = salary - salary % 500

    bonus_pct = random.choices(
        [0, 5, 10, 15, 20], weights=[40, 25, 20, 10, 5], k=1
    )[0]

    return {
        "EmployeeID": f"E{emp_id:05d}",
        "Gender": random.choice(["Male", "Female"]),
        "Age": age,
        "MaritalStatus": random.choices(
            ["Single", "Married", "Married", "Divorced"],
            weights=[35, 45, 45, 10], k=1,
        )[0],
        "Education": random.choice(EDUCATION),
        "Department": department,
        "JobRole": job_role,
        "JobLevel": job_level,
        "OfficeLocation": random.choice(OFFICES),
        "EmploymentType": None,  # filled below (needed by attrition calc)
        "HireDate": hire_date.isoformat(),
        "ExitDate": None,
        "IsActive": is_active,
        "ExitReason": None,
        "SalaryAnnual": salary,
        "BonusPct": bonus_pct,
        "PerformanceRating": performance,
        "SatisfactionScore": satisfaction,
        "AbsenteeDays": max(0, min(45, int(random.gauss(6, 6)))),
    }


def build_rows(n=3000) -> list:
    rows = []
    for i in range(1, n + 1):
        rows.append(make_employee(i))

    # employment type needs to be set before finalising exit reason/story
    for r in rows:
        r["EmploymentType"] = random.choices(
            ["Full-Time", "Full-Time", "Full-Time", "Part-Time", "Contract"],
            weights=[60, 60, 60, 15, 10], k=1,
        )[0]
        if not r["IsActive"]:
            exit_after = timedelta(days=random.randint(30, 150)) if r["HireDate"] < (SNAPSHOT - timedelta(days=60)).isoformat() else timedelta(days=15)
            max_ok = SNAPSHOT - date.fromisoformat(r["HireDate"]) - timedelta(days=1)
            exit_days = max(1, min(exit_after.days, max_ok.days))
            r["ExitDate"] = (date.fromisoformat(r["HireDate"]) + timedelta(days=exit_days)).isoformat()
            r["ExitReason"] = random.choices(
                list(EXIT_REASONS.values()), weights=list(EXIT_REASONS.keys()), k=1
            )[0]

    for r in rows:
        if not r["IsActive"] and r["ExitDate"] is not None:
            tenure = (date.fromisoformat(r["ExitDate"]) - date.fromisoformat(r["HireDate"])).days // 30
        else:
            tenure = (SNAPSHOT - date.fromisoformat(r["HireDate"])).days // 30
        r["TenureMonths"] = max(1, tenure)
    return rows


def inject_raw_data_issues(rows: list) -> list:
    """Add realistic data-quality problems so the cleaning step is meaningful."""
    raw = [dict(r) for r in rows]

    for i in raw:
        # lowercase a few genders + some 'Unknown'
        g = random.random()
        if g < 0.04:
            i["Gender"] = i["Gender"].lower()
        elif g < 0.06:
            i["Gender"] = "Unknown"
        # a few blank salaries and departments
        if random.random() < 0.012:
            i["SalaryAnnual"] = None
        if random.random() < 0.010:
            i["Department"] = None
        # a couple of insane age typos
        if random.random() < 0.006:
            i["Age"] = random.choice([200, 210, 250])
        # a couple of salary typos (outliers)
        if random.random() < 0.004:
            i["SalaryAnnual"] = random.choice([999999, 1000000, 1500000])

    # exact duplicate rows
    for d in random.sample(range(len(raw)), 10):
        raw.append(dict(raw[d]))

    random.shuffle(raw)
    return raw


def main():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PBI_DIR.mkdir(parents=True, exist_ok=True)

    rows = build_rows(3000)
    clean = sorted(rows, key=lambda r: r["EmployeeID"])

    # ---- cleaned flat CSV -------------------------------------------------
    clean_cols = [
        "EmployeeID", "Gender", "Age", "MaritalStatus", "Education",
        "Department", "JobRole", "JobLevel", "OfficeLocation",
        "EmploymentType", "HireDate", "ExitDate", "IsActive", "ExitReason",
        "SalaryAnnual", "BonusPct", "PerformanceRating", "SatisfactionScore",
        "TenureMonths", "AbsenteeDays",
    ]
    with open(BASE / "data" / "hr_cleaned.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=clean_cols)
        w.writeheader()
        for r in clean:
            w.writerow({c: r.get(c) for c in clean_cols})

    # ---- messy raw CSV ----------------------------------------------------
    messy = inject_raw_data_issues(clean)
    with open(RAW_DIR / "hr_employees_raw.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=clean_cols)
        w.writeheader()
        w.writerows(messy)

    # ---- Power BI star schema ---------------------------------------------
    dim_employees = []
    fact_events = []
    fact_comp = []
    for r in clean:
        dim_employees.append({
            "EmployeeID": r["EmployeeID"],
            "Gender": r["Gender"],
            "Age": r["Age"],
            "AgeBand": "Under 30" if r["Age"] < 30 else ("30-39" if r["Age"] < 40 else ("40-49" if r["Age"] < 50 else "50+")),
            "MaritalStatus": r["MaritalStatus"],
            "Education": r["Education"],
            "Department": r["Department"],
            "JobRole": r["JobRole"],
            "JobLevel": r["JobLevel"],
            "OfficeLocation": r["OfficeLocation"],
            "EmploymentType": r["EmploymentType"],
            "HireDate": r["HireDate"],
            "IsActive": r["IsActive"],
        })
        fact_events.append({
            "EventID": f"EVT-HIRE-{r['EmployeeID']}",
            "EmployeeID": r["EmployeeID"],
            "EventType": "Hire",
            "EventDate": r["HireDate"],
            "Detail": "",
        })
        if not r["IsActive"]:
            fact_events.append({
                "EventID": f"EVT-EXIT-{r['EmployeeID']}",
                "EmployeeID": r["EmployeeID"],
                "EventType": "Exit",
                "EventDate": r["ExitDate"],
                "Detail": r["ExitReason"],
            })
        fact_comp.append({
            "EmployeeID": r["EmployeeID"],
            "SalaryAnnual": r["SalaryAnnual"],
            "BonusPct": r["BonusPct"],
            "PerformanceRating": r["PerformanceRating"],
            "SatisfactionScore": r["SatisfactionScore"],
            "SalaryBand": "Low (<60k)" if (r["SalaryAnnual"] or 0) < 60000 else ("Mid (60-100k)" if (r["SalaryAnnual"] or 0) < 100000 else "High (100k+)"),
        })

    with open(PBI_DIR / "DimEmployee.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(dim_employees[0]))
        w.writeheader()
        w.writerows(dim_employees)
    with open(PBI_DIR / "FactEmploymentEvents.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(fact_events[0]))
        w.writeheader()
        w.writerows(fact_events)
    with open(PBI_DIR / "FactCompensation.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(fact_comp[0]))
        w.writeheader()
        w.writerows(fact_comp)

    # DimDate
    d = date(2019, 1, 1)
    dim_date = []
    while d <= date(2026, 12, 31):
        dim_date.append({
            "DateKey": d.strftime("%Y-%m-%d"),
            "Date": d.isoformat(),
            "Year": d.year,
            "Quarter": f"Q{(d.month - 1) // 3 + 1}",
            "YearQuarter": f"{d.year}-Q{(d.month - 1) // 3 + 1}",
            "MonthNumber": d.month,
            "MonthName": d.strftime("%B"),
            "YearMonth": f"{d.year}-{d.month:02d}",
            "WeekNumber": d.isocalendar()[1],
            "IsWeekend": "Yes" if d.weekday() >= 5 else "No",
        })
        d += timedelta(days=1)
    with open(PBI_DIR / "DimDate.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(dim_date[0]))
        w.writeheader()
        w.writerows(dim_date)

    active = sum(1 for r in clean if r["IsActive"])
    print(f"Total employees        : {len(clean)}")
    print(f"Active                 : {active}")
    print(f"Attrited (inactive)    : {len(clean) - active}  ({100 * (len(clean)-active)/len(clean):.1f}%)")
    print(f"Hire events            : {len(clean)}")
    print(f"Exit events            : {len(clean) - active}")
    print("Star-schema CSVs + flattened data written under data/.")


if __name__ == "__main__":
    main()