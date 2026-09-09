"""
Builds docs/index.html — a self-contained, Power BI-style interactive dashboard
for the HR Workforce Analytics project.

- Reads data/hr_cleaned.csv
- Embeds a compact employee-level dataset into the page
- Renders KPIs + ECharts visuals, all filterable in the browser
- Zero build step (single static file, works on GitHub Pages)

Run:  .venv/bin/python dashboard/html_build.py
"""

import csv
import json
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CLEAN_CSV = BASE / "data" / "hr_cleaned.csv"
OUT = BASE / "docs" / "index.html"

COLS = [
    "EmployeeID", "Gender", "Age", "Education", "Department", "JobRole",
    "JobLevel", "OfficeLocation", "EmploymentType", "HireYear", "ExitReason",
    "Attrited", "Salary", "BonusPct", "Performance", "Satisfaction",
    "TenureMonths", "AbsenteeDays",
]


def load_rows():
    rows = []
    with open(CLEAN_CSV, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            hire_year = datetime.strptime(r["HireDate"], "%Y-%m-%d").year
            attrited = 0 if r["IsActive"] in ("True", "Yes") else 1
            rows.append([
                r["EmployeeID"], r["Gender"], int(r["Age"]), r["Education"],
                r["Department"], r["JobRole"], int(r["JobLevel"]),
                r["OfficeLocation"], r["EmploymentType"], hire_year,
                (r["ExitReason"] or ""), attrited, float(r["SalaryAnnual"]),
                float(r["BonusPct"]), int(r["PerformanceRating"]),
                int(r["SatisfactionScore"]), int(r["TenureMonths"]) if r["TenureMonths"] else 0,
                int(r["AbsenteeDays"]),
            ])
    return rows


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>HR Workforce Analytics Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<style>
  :root { --dark:#1F3864; --mid:#2E75B6; --accent:#548235; --red:#C00000;
           --bg:#F3F2F1; --card:#FFFFFF; --ink:#1A1A1A; --sub:#6E6E6E; }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI', system-ui, -apple-system, Arial, sans-serif;
         background:var(--bg); color:var(--ink); }
  header { background:linear-gradient(100deg, var(--dark), var(--mid)); color:#fff;
           padding:22px 32px; }
  header .title { font-size:24px; font-weight:700; letter-spacing:.5px; }
  header .sub  { font-size:13px; opacity:.85; margin-top:3px; }
  .wrap { max-width:1360px; margin:0 auto; padding:18px 24px 40px; }

  .filterbar { display:flex; gap:18px; align-items:flex-end; flex-wrap:wrap;
               background:var(--card); border:1px solid #E0E0E0; border-radius:8px;
               padding:12px 18px; margin:14px 0; box-shadow:0 1px 2px rgba(0,0,0,.05); }
  .filterbar .fld { display:flex; flex-direction:column; gap:4px; font-size:12px;
                    color:var(--sub); }
  .filterbar select { min-width:170px; padding:7px 10px; border:1px solid #C8C8C8;
                      border-radius:6px; font-size:13px; background:#fff; cursor:pointer; }
  .filterbar button { padding:8px 14px; border:none; border-radius:6px;
                      background:var(--mid); color:#fff; font-size:13px; font-weight:600;
                      cursor:pointer; }
  .filterbar button:hover { filter:brightness(1.08); }

  .kpis { display:grid; grid-template-columns:repeat(6,1fr); gap:14px; margin-bottom:16px; }
  .kpi { background:var(--card); border:1px solid #E0E0E0; border-left:5px solid var(--mid);
         border-radius:8px; padding:14px 16px; box-shadow:0 1px 2px rgba(0,0,0,.05); }
  .kpi.hl { border-left-color:var(--red); }
  .kpi.ok { border-left-color:var(--accent); }
  .kpi .lbl { font-size:11px; letter-spacing:.6px; color:var(--sub); text-transform:uppercase; }
  .kpi .val { font-size:26px; font-weight:700; color:var(--dark); margin-top:4px; }
  .kpi .foot { font-size:12px; color:var(--sub); margin-top:2px; }

  .grid { display:grid; grid-template-columns:repeat(3,1fr); gap:14px; }
  .card { background:var(--card); border:1px solid #E0E0E0; border-radius:8px;
          padding:12px 14px 4px; box-shadow:0 1px 2px rgba(0,0,0,.05); }
  .card.w2 { grid-column:span 2; }
  .card h3 { font-size:13px; font-weight:700; color:var(--dark);
             margin:2px 2px 8px; letter-spacing:.3px; }
  .chart { width:100%; height:300px; }
  .chart.tall { height:330px; }
  footer { text-align:center; color:var(--sub); font-size:12px; padding:14px 0 30px; }
  @media (max-width:1100px) {
    .kpis { grid-template-columns:repeat(3,1fr); }
    .grid { grid-template-columns:1fr 1fr; }
    .card.w2 { grid-column:span 2; }
  }
  @media (max-width:700px) {
    .kpis, .grid { grid-template-columns:1fr; }
    .card.w2 { grid-column:span 1; }
  }
</style>
</head>
<body>
<header>
  <div class="title">HR WORKFORCE ANALYTICS</div>
  <div class="sub">Power BI-style dashboard · Synthetic data, snapshot 31-Dec-2025 ·
     Excel + Power BI project by Naman Khanda</div>
</header>
<div class="wrap">
  <div class="filterbar" id="filterbar"></div>
  <div class="kpis" id="kpis"></div>
  <div class="grid">
    <div class="card"><h3 id="t_gender">Gender distribution</h3><div id="c_gender" class="chart"></div></div>
    <div class="card"><h3 id="t_age">Age band distribution</h3><div id="c_age" class="chart"></div></div>
    <div class="card"><h3 id="t_dept">Headcount by department</h3><div id="c_dept" class="chart tall"></div></div>
    <div class="card w2"><h3 id="t_role">Employee distribution by job role</h3><div id="c_role" class="chart tall"></div></div>
    <div class="card w2"><h3 id="t_hires">Hiring trend by year (hires vs exits)</h3><div id="c_hires" class="chart"></div></div>
    <div class="card"><h3 id="t_agedept">Attrition rate by department</h3><div id="c_agedept" class="chart"></div></div>
    <div class="card w2"><h3 id="t_atrrole">Attrition (employees left) by job role</h3><div id="c_atrrole" class="chart"></div></div>
    <div class="card"><h3 id="t_salary">Avg salary by department</h3><div id="c_salary" class="chart"></div></div>
  </div>
</div>
<footer>Data: synthetic HR export (3,000 employees) · cleaned in Excel ·
       Same dataset powers the Excel workbook &amp; Power BI report. Tags filter all visuals.</footer>

<script>
const AGG = __DATA__;
const COLS = __COLS__;
const I = Object.fromEntries(COLS.map((c, i) => [c, i]));
const rows = AGG.map(r => Object.fromEntries(COLS.map((c, i) => [c, r[i]])));

const PALETTE = ['#1F3864', '#2E75B6', '#548235', '#C00000', '#7030A0',
                 '#ED7D31', '#5B9BD5', '#A5A5A5', '#FFC000', '#264478'];
const fmtUS = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
const fmtUSD = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });

let state = { dept: 'All', gender: 'All', year: 'All' };

const filter = () => rows.filter(r =>
  (state.dept === 'All' || r.Department === state.dept) &&
  (state.gender === 'All' || r.Gender === state.gender) &&
  (state.year === 'All' || r.HireYear === state.year));

const groupBy = (arr, keyFn, valFn = r => r) => {
  const m = new Map();
  for (const r of arr) {
    const k = keyFn(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(valFn(r));
  }
  return m;
};

function kpiCard(label, value, foot, cls) {
  return `<div class="kpi ${cls}"><div class="lbl">${label}</div>
          <div class="val">${value}</div><div class="foot">${foot}</div></div>`;
}

function renderKPIs(d) {
  const n = d.length;
  const active = d.filter(r => !r.Attrited).length;
  const left = n - active;
  const atr = n ? left / n : 0;
  const sal = n ? d.reduce((s, r) => s + r.Salary, 0) / n : 0;
  const age = n ? d.reduce((s, r) => s + r.Age, 0) / n : 0;
  const ten = n ? d.reduce((s, r) => s + r.TenureMonths, 0) / n : 0;
  document.getElementById('kpis').innerHTML =
    kpiCard('Total employees', fmtUS.format(n), 'headcount (filtered)') +
    kpiCard('Active', fmtUS.format(active), `${(active / (n || 1) * 100).toFixed(1)}% of cohort`, 'ok') +
    kpiCard('Attrition / left', fmtUS.format(left), `${(atr * 100).toFixed(1)}% attrition rate`, 'hl') +
    kpiCard('Avg salary', fmtUSD.format(sal), 'annual base pay', 'ok') +
    kpiCard('Avg age', age.toFixed(1), 'years') +
    kpiCard('Avg tenure', ten.toFixed(1), 'months');
}

const baseText = { color: '#5A5A5A', fontSize: 11 };
const axLabel = { color: '#444', fontSize: 11 };
const gridOpt = { left: 8, right: 12, top: 20, bottom: 4, containLabel: true };

function donut(title = '') {
  return {
    tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
    legend: { bottom: 0, textStyle: baseText },
    series: [{ type: 'pie', radius: ['52%', '76%'], center: ['50%', '46%'],
      label: { show: false }, itemStyle: { borderRadius: 4, borderColor: '#fff', borderWidth: 2 },
      data: title.split ? [] : [] }],
  };
}

function barX() {
  return { grid: { ...gridOpt, right: 26 },
    xAxis: { type: 'value', axisLabel: axLabel, splitLine: { lineStyle: { color: '#EDEDED' } } },
    yAxis: { type: 'category', axisLabel: axLabel, axisLine: { show: false },
             axisTick: { show: false }, inverse: true },
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } } };
}

function barY() {
  return { grid: { ...gridOpt, left: 34, right: 8 },
    xAxis: { type: 'category', axisLabel: { ...axLabel, interval: 0 }, axisTick: { show: false },
             axisLine: { lineStyle: { color: '#CCC' } } },
    yAxis: { type: 'value', axisLabel: axLabel, splitLine: { lineStyle: { color: '#EDEDED' } } },
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } } };
}

const charts = {};
function mk(id) { if (!charts[id]) charts[id] = echarts.init(document.getElementById(id)); return charts[id]; }
function set(id, opt) { mk(id).setOption(opt, true); }

function renderCharts(d) {
  // 1 gender donut
  const g = [...groupBy(d, r => r.Gender).entries()]
    .map(([k, v]) => ({ name: k, value: v.length }))
    .sort((a, b) => b.value - a.value);
  set('c_gender', { ...donut(),
    series: [{ type: 'pie', radius: ['50%', '76%'], center: ['50%', '44%'],
      label: { show: false }, itemStyle: { borderRadius: 4, borderColor: '#fff', borderWidth: 2 },
      color: PALETTE, data: g }] });

  // 2 age band
  const bands = ['Under 30', '30-39', '40-49', '50+'];
  const ageCounts = bands.map(b => d.filter(r =>
    b === 'Under 30' ? r.Age < 30 : b === '30-39' ? r.Age < 40 :
    b === '40-49' ? r.Age < 50 : r.Age >= 50).length);
  set('c_age', { ...barY(), color: [PALETTE[1]],
    series: [{ type: 'bar', barWidth: 22, data: ageCounts,
      itemStyle: { borderRadius: [4, 4, 0, 0] },
      label: { show: true, position: 'top', fontSize: 11, color: '#444' } }],
    xAxis: { ...barY().xAxis, data: bands } });

  // 3 dept headcount
  const deptH = [...groupBy(d, r => r.Department).entries()]
    .map(([k, v]) => ({ name: k, n: v.length, atr: v.filter(x => x.Attrited).length / v.length }))
    .sort((a, b) => b.n - a.n);
  set('c_dept', { ...barX(), color: [PALETTE[0]],
    series: [{ type: 'bar', barWidth: 18, data: deptH.map(x => x.n),
      itemStyle: { borderRadius: [0, 4, 4, 0] },
      label: { show: true, position: 'right', fontSize: 11, color: '#444' } }],
    yAxis: { ...barX().yAxis, data: deptH.map(x => x.name) } });

  // 4 job role distribution
  const roleH = [...groupBy(d, r => r.JobRole).entries()]
    .map(([k, v]) => ({ name: k, n: v.length, atr: v.filter(x => x.Attrited).length / v.length }))
    .sort((a, b) => b.n - a.n).slice(0, 12);
  set('c_role', { ...barX(), color: [PALETTE[4]],
    series: [{ type: 'bar', barWidth: 13, data: roleH.map(x => x.n),
      itemStyle: { borderRadius: [0, 3, 3, 0] },
      label: { show: true, position: 'right', fontSize: 11, color: '#444' } }],
    yAxis: { ...barX().yAxis, data: roleH.map(x => x.name) } });

  // 5 hires vs exits by year
  const years = [...new Set(d.map(r => r.HireYear))].sort();
  const hires = years.map(y => d.filter(r => r.HireYear === y).length);
  const exits = years.map(y => d.filter(r => r.HireYear === y && r.Attrited).length);
  set('c_hires', { ...barY(), tooltip: { trigger: 'axis' },
    legend: { bottom: 0, textStyle: baseText },
    color: [PALETTE[1], PALETTE[3]],
    series: [
      { name: 'Hires', type: 'bar', barWidth: 18, data: hires,
        itemStyle: { borderRadius: [4, 4, 0, 0] } },
      { name: 'Exits', type: 'line', data: exits, smooth: true, symbolSize: 7,
        lineStyle: { width: 2 } },
    ],
    xAxis: { ...barY().xAxis, data: years } });

  // 6 attrition rate by dept
  const deptA = deptH.map(x => ({ name: x.name, atr: x.atr }))
    .sort((a, b) => b.atr - a.atr);
  set('c_agedept', { ...barX(), color: [PALETTE[3]],
    series: [{ type: 'bar', barWidth: 18, data: deptA.map(x => +(x.atr * 100).toFixed(1)),
      itemStyle: { borderRadius: [0, 4, 4, 0] },
      label: { show: true, position: 'right', formatter: '{c}%', fontSize: 11, color: '#444' } }],
    yAxis: { ...barX().yAxis, data: deptA.map(x => x.name) } });

  // 7 attrition by role (count left)
  const roleA = [...groupBy(d, r => r.JobRole).entries()]
    .map(([k, v]) => ({ name: k, left: v.filter(x => x.Attrited).length }))
    .filter(x => x.left > 0).sort((a, b) => b.left - a.left).slice(0, 12);
  set('c_atrrole', { ...barX(), color: [PALETTE[3]],
    series: [{ type: 'bar', barWidth: 16, data: roleA.map(x => x.left),
      itemStyle: { borderRadius: [0, 4, 4, 0] },
      label: { show: true, position: 'right', fontSize: 11, color: '#444' } }],
    yAxis: { ...barX().yAxis, data: roleA.map(x => x.name) } });

  // 8 avg salary by dept
  const deptS = deptH.map(x => ({ name: x.name,
    sal: d.filter(r => r.Department === x.name).reduce((s, r) => s + r.Salary, 0) /
         d.filter(r => r.Department === x.name).length }))
    .sort((a, b) => b.sal - a.sal);
  set('c_salary', { ...barX(), color: [PALETTE[2]],
    series: [{ type: 'bar', barWidth: 18, data: deptS.map(x => +(x.sal / 1000).toFixed(1)),
      itemStyle: { borderRadius: [0, 4, 4, 0] },
      label: { show: true, position: 'right', formatter: p => '$' + fmtUS.format(p.value * 1000),
               fontSize: 11, color: '#444' } }],
    yAxis: { ...barX().yAxis, data: deptS.map(x => x.name) } });
}

function buildFilters() {
  const depts = [...new Set(rows.map(r => r.Department))].sort();
  const genders = [...new Set(rows.map(r => r.Gender))].sort();
  const years = [...new Set(rows.map(r => r.HireYear))].sort();
  const sel = (id) => document.getElementById(id).value;
  document.getElementById('filterbar').innerHTML = `
    <div class="fld"><span>Department</span>
      <select id="f_dept"><option>All</option>${depts.map(x => `<option>${x}</option>`).join('')}</select></div>
    <div class="fld"><span>Gender</span>
      <select id="f_gender"><option>All</option>${genders.map(x => `<option>${x}</option>`).join('')}</select></div>
    <div class="fld"><span>Hire year</span>
      <select id="f_year"><option>All</option>${years.map(x => `<option>${x}</option>`).join('')}</select></div>
    <button id="f_reset">Reset filters</button>`;
  const refresh = () => {
    state = { dept: sel('f_dept'), gender: sel('f_gender'), year: Number(sel('f_year')) || 'All' };
    const d = filter();
    renderKPIs(d);
    renderCharts(d);
  };
  ['f_dept', 'f_gender', 'f_year'].forEach(id => document.getElementById(id).addEventListener('change', refresh));
  document.getElementById('f_reset').addEventListener('click', () => {
    ['f_dept', 'f_gender', 'f_year'].forEach(id => document.getElementById(id).value = 'All');
    refresh();
  });
}

window.addEventListener('resize', () => Object.values(charts).forEach(c => c.resize()));
buildFilters();
renderKPIs(filter());
renderCharts(filter());
</script>
</body>
</html>
"""


def main():
    rows = load_rows()
    out = (
        HTML_TEMPLATE
        .replace("__DATA__", json.dumps(rows, separators=(",", ":")))
        .replace("__COLS__", json.dumps(COLS, separators=(",", ":")))
    )
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(out, encoding="utf-8")
    print(f"Dashboard written -> {OUT}  ({len(out)//1024} KB, {len(rows)} employees)")


if __name__ == "__main__":
    main()