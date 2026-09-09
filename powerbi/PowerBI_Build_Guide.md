# Power BI Build Guide — HR Workforce Analytics

The cleaned dataset is exported as a **star schema** in `data/powerbi/` so you can
load it into Power BI Desktop (Windows) in a couple of minutes.

## Star schema

```
            DimDate (1) ── (many) FactEmploymentEvents (many) ── (1) DimEmployee
                                     │                                     │
                                     │                                     └── (1) FactCompensation
                                     └──────────────── preamble: no dept link needed
```

| Table                   | Grain                    | Key(s)                    |
|-------------------------|--------------------------|---------------------------|
| `DimEmployee`           | one row per employee     | EmployeeID                |
| `DimDate`               | one row per calendar day | DateKey                   |
| `FactEmploymentEvents`  | one row per hire or exit | EventID (EmployeeID, EventDate) |
| `FactCompensation`      | snapshot per employee    | EmployeeID                |

## Step 1 — Load data

1. Power BI Desktop → **Get Data → Folder** → pick `data/powerbi/`
2. **Transform Data**:
   - Promote headers; keep only the four CSV files
   - `EventDate` / `HireDate` / `DateKey` → make **Date** data type
   - `DimEmployee` add column: Active Flag = `IF [IsActive] THEN 1 ELSE 0`
3. Close & Apply.

## Step 2 — Model relationships (Model view)

- `DimDate[DateKey] (1)` → `FactEmploymentEvents[EventDate] (*)`
- `DimEmployee[EmployeeID] (1)` → `FactEmploymentEvents[EmployeeID] (*)`
- `DimEmployee[EmployeeID] (1)` → `FactCompensation[EmployeeID] (*)`
- Cross-filter direction: *Single* is fine here.

## Step 3 — DAX measures (New measure)

```dax
Total Headcount = COUNTROWS('DimEmployee')

Active Headcount =
    CALCULATE( COUNTROWS('DimEmployee'), 'DimEmployee'[IsActive] = TRUE )

Attrition Rate =
    DIVIDE( [Total Headcount] - [Active Headcount], [Total Headcount] )

Avg Salary = AVERAGE('FactCompensation'[SalaryAnnual])

Avg Satisfaction = AVERAGE('FactCompensation'[SatisfactionScore])

Hires = CALCULATE( COUNTROWS('FactEmploymentEvents'), 'FactEmploymentEvents'[EventType] = "Hire" )

Exits = CALCULATE( COUNTROWS('FactEmploymentEvents'), 'FactEmploymentEvents'[EventType] = "Exit" )
```

## Step 4 — Report pages

**Page 1: Overview**
- Slicers: Year (from DimDate), Department (DimEmployee)
- KPI cards: Total Headcount, Attrition Rate, Avg Salary, Avg Satisfaction
- Clustered column: Headcount by Department
- Line: Hires vs Exits over `DimDate[YearMonth]`

**Page 2: Compensation**
- Clustered column: Avg Salary by Job Level (Legend = Gender) — the pay-gap visual
- Matrix: Rows = Department, Columns = Gender, Value = Avg Salary
- Table: Department, Female, Male, Gap = `DIVIDE(Male-Female, Female)`

**Page 3: People Movement**
- Line: Hires & Exits by YearMonth
- Donut: Exits by Detail (Reason)
- Table: Top attrition departments

**Page 4: Equity & Retention** (optional stretch)
- Avg Salary vs PerformanceRating (scatter)
- Attrition by AgeBand (clustered bar)

## Step 5 — Publish

File → **Publish** → pick a workspace → share the link. From there you can
add a scheduled refresh if you point sources at the CSV folder again — for a
portfolio demo, a static one-time refresh is enough.

## Expected numbers (so you can sanity-check)

- Total employees: 3,000 · Active: 2,381 · Attrition: ~20.6%
- Highest attrition: Customer Support (~29%), Operations (~25%), Sales (~22%)
- Highest paying: Executive, Research & Development (by average)
- Exits skew to "Resigned" (338) followed by Retirement (110)