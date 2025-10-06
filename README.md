# Weekly Report Generator

Turn weekly repair-claim CSVs into an **Excel report** with KPIs, pivots, a chart, and validation checks — perfect for Data Specialist / QA workflows.

## Features
- ✅ Filters to a target week (Mon–Sun)
- ✅ KPI summary: totals, status mix, assignment vs non-assignment
- ✅ SLA checks: assignment lag & resolution time
- ✅ Sheets: Summary, By Branch, By Service, By PM, Raw Data, Errors
- ✅ Excel output with an embedded column chart

## Input schema (CSV)
Required columns:  
`claim_id, branch, line_of_service, is_assignment, received_date, assigned_pm, assigned_date, status, dash_job_id, completed_date`

## Quick start
```bash
# 1) Create venv (Windows PowerShell)
python -m venv .venv
. .venv/Scripts/Activate.ps1

# 2) Install deps
pip install -r requirements.txt

# 3) Run (last full week, auto)
python src/generate_report.py --input data/claims_sample.csv

# OR specify the week
python src/generate_report.py --week-start 2025-09-22

# Optional: tweak SLAs
python src/generate_report.py --sla-assign-days 1 --sla-complete-days 7
