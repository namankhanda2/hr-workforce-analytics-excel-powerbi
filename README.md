# HR Workforce Analytics — Excel + Power BI Project

End-to-end workforce analytics project: a messy HR export is cleaned, analysed
and visualised in **Excel**, and the same data is exported ready to load into
**Power BI** as a star schema.

```
Messy export  ->  pandas clean  ->  Excel workbook            (this repo)
                                  ->  Power BI star schema    (data/powerbi/)
```

## What it answers

- Headcount, attrition rate, average salary and tenure by **department**.
- **Gender pay-gap** by job level (Excel table + Power BI matrix).
- **Hiring trend** by year and **departures by exit reason**.
- KPI dashboard page with the numbers an HR manager cares about.

## The Excel workbook

`excel/HR_Workforce_Analytics.xlsx` contains 5 sheets:

| Sheet        | Content                                                          |
|--------------|------------------------------------------------------------------|
| Guide        | methodology + cleaning log + navigation                          |
| Raw Data     | as-collected file — duplicates, blank salary, typos, outliers    |
| Clean Data   | cleaned + derived columns via live formulas; exported as Excel Table `tCleanData` |
| Insights     | SUMIFS / COUNTIFS / AVERAGEIFS summaries + native Excel charts   |
| Dashboard    | KPI cards, charts, key findings                                  |

Highlighted Excel skills: `IFS`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`,
`COUNTIF`, `DATEDIF`-style date logic, conditional formatting, color scales,
Excel Tables, native charts, AutoFilter, number formatting.

## Key numbers (synthetic data — snapshot 31-Dec-2025)

| Metric              | Value  |
|---------------------|--------|
| Total employees     | 3,000  |
| Active              | 2,381  |
| Attrition rate      | 20.6%  |
| Highest attrition   | Customer Support (29.4%), Operations (25.0%), Sales (22.2%) |
| Pay gap (M vs F)    | uneven by level — see Insights / Power BI matrix |

## Live interactive dashboard

A Power BI-style dashboard (same KPIs and visuals as the report, rendered in the
browser) is deployed to GitHub Pages:

**https://namankhanda2.github.io/hr-workforce-analytics-excel-powerbi/**

Source: `dashboard/html_build.py` builds `docs/index.html` (single self-contained
file with an embedded dataset — works offline). It includes:

- KPI cards: total employees, active, attrition, avg salary, avg age, avg tenure
- Gender & age-band distribution, headcount by department, role distribution
- Hiring trend (hires vs exits by year), attrition by department and role
- Avg salary by department; department/gender/hire-year filters recompute everything

`docs/preview.png` is a static screenshot of the dashboard.

## Power BI

`data/powerbi/` ships a ready-to-import star schema:
- `DimEmployee.csv`, `DimDate.csv`, `FactEmploymentEvents.csv`, `FactCompensation.csv`

Full step-by-step build (load → relationships → DAX → report pages → publish):
**[powerbi/PowerBI_Build_Guide.md](powerbi/PowerBI_Build_Guide.md)**

## Project structure

```
├── generate_hr_data.py        # synthetic raw dataset + star-schema CSVs
├── build_excel.py             # builds excel/HR_Workforce_Analytics.xlsx
├── dashboard/
│   └── html_build.py          # builds docs/index.html (Power BI-style web dashboard)
├── docs/
│   ├── index.html             # live dashboard (deployed to GitHub Pages)
│   └── preview.png            # dashboard screenshot
├── data/
│   ├── raw/hr_employees_raw.csv    # messy, as-collected
│   ├── hr_cleaned.csv              # cleaned flat file
│   └── powerbi/                    # star schema for Power BI
├── excel/HR_Workforce_Analytics.xlsx
└── powerbi/PowerBI_Build_Guide.md
```

## Run it yourself

```bash
pip install pandas openpyxl
python generate_hr_data.py   # regenerate data (reproducible, seed fixed)
python build_excel.py        # rebuild the workbook
```

## Why this is a strong portfolio piece

- Shows a **real cleanup decision log**, not just pretty charts
- Excel formulas are **live** (change a KPI and the workbook recalculates)
- **Same data, two tools** — proves Excel + Power BI fluency
- The Power BI guide shows DAX, relationships, and report design, not just screenshots

## Notes

- All data is synthetic (generated with a fixed random seed) — no real employee
  information is used.

## License

MIT