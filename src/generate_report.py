import argparse
from datetime import datetime, timedelta, date
from dateutil import tz
import os
import pandas as pd
import numpy as np

ALLOWED_STATUSES = {"New", "In Progress", "Completed", "On Hold"}

REQUIRED_COLUMNS = [
    "claim_id",
    "branch",
    "line_of_service",
    "is_assignment",
    "received_date",
    "assigned_pm",
    "assigned_date",
    "status",
    "dash_job_id",
    "completed_date",
]

def parse_args():
    p = argparse.ArgumentParser(description="Weekly Repair Claims Report Generator")
    p.add_argument("--input", default="data/claims_sample.csv", help="Path to input CSV")
    p.add_argument("--outdir", default="outputs", help="Directory to write reports")
    p.add_argument("--week-start", help="ISO date (YYYY-MM-DD). If omitted, uses last full Mon–Sun", default=None)
    p.add_argument("--sla-assign-days", type=int, default=1, help="Days to assign (SLA)")
    p.add_argument("--sla-complete-days", type=int, default=7, help="Days to complete (SLA)")
    return p.parse_args()

def last_full_week():
    # Get last Monday-Sunday range that fully ended before today (local US/Eastern)
    today = datetime.now(tz=tz.gettz("America/New_York")).date()
    # Monday = 0
    weekday = today.weekday()
    # Go back to last Monday of the previous full week
    end = today - timedelta(days=weekday + 1)  # last Sunday
    start = end - timedelta(days=6)            # previous Monday
    return start, end

def load_data(path):
    df = pd.read_csv(path, dtype=str).fillna("")
    # Parse dates safely
    for c in ["received_date", "assigned_date", "completed_date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    # Normalize strings
    for c in ["branch", "line_of_service", "status", "is_assignment", "assigned_pm"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df

def filter_week(df, start, end):
    mask = (df["received_date"] >= start) & (df["received_date"] <= end)
    return df.loc[mask].copy()

def validate(df):
    errors = []

    # Required columns present?
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        errors.append({"row": "-", "field": "schema", "issue": f"Missing required columns: {missing_cols}"})
        return pd.DataFrame(errors)

    # Row-level checks
    for idx, row in df.iterrows():
        if row["status"] and row["status"] not in ALLOWED_STATUSES:
            errors.append({"row": idx, "field": "status", "issue": f"Invalid status '{row['status']}'"})

        # If assigned/active, should have DASH job id
        if row["status"] in {"In Progress", "Completed", "On Hold"} and not row["dash_job_id"]:
            errors.append({"row": idx, "field": "dash_job_id", "issue": "Missing DASH job id for active/closed claim"})

        # If assigned, should have assigned PM and date
        if row["is_assignment"].lower() in {"yes", "y", "true", "1"}:
            if not row["assigned_pm"]:
                errors.append({"row": idx, "field": "assigned_pm", "issue": "Missing assigned PM on assignment"})
            if pd.isna(row["assigned_date"]):
                errors.append({"row": idx, "field": "assigned_date", "issue": "Missing assigned_date on assignment"})

    return pd.DataFrame(errors)

def compute_kpis(df, sla_assign_days=1, sla_complete_days=7):
    out = {}

    total = len(df)
    out["Total Claims"] = total
    out["Assignments"] = int((df["is_assignment"].str.lower().isin(["yes","y","true","1"])).sum())
    out["Non-Assignments"] = total - out["Assignments"]

    # Status distribution
    for s in ["New", "In Progress", "Completed", "On Hold"]:
        out[f"Status: {s}"] = int((df["status"] == s).sum())

    # Lags
    df["assign_lag_days"] = (pd.to_datetime(df["assigned_date"]) - pd.to_datetime(df["received_date"])).dt.days
    df["resolution_days"] = (pd.to_datetime(df["completed_date"]) - pd.to_datetime(df["received_date"])).dt.days

    # Averages (drop NaN)
    out["Avg Assign Lag (days)"] = round(float(df["assign_lag_days"].dropna().mean()), 2) if df["assign_lag_days"].notna().any() else np.nan
    out["Avg Resolution (days)"] = round(float(df["resolution_days"].dropna().mean()), 2) if df["resolution_days"].notna().any() else np.nan

    # SLA breaches
    assign_breaches = ((df["assign_lag_days"] > sla_assign_days)).sum()
    # Only completed/in progress for completion SLA baseline
    comp_mask = df["resolution_days"].notna()
    complete_breaches = ((df.loc[comp_mask, "resolution_days"] > sla_complete_days)).sum()
    out[f"SLA Breaches: Assign>{sla_assign_days}d"] = int(assign_breaches)
    out[f"SLA Breaches: Complete>{sla_complete_days}d"] = int(complete_breaches)

    # Groupings
    by_branch = df.groupby("branch", dropna=False)["claim_id"].count().reset_index(name="count").sort_values("count", ascending=False)
    by_service = df.groupby("line_of_service", dropna=False)["claim_id"].count().reset_index(name="count").sort_values("count", ascending=False)
    by_pm = df.groupby("assigned_pm", dropna=False)["claim_id"].count().reset_index(name="count").sort_values("count", ascending=False)

    return out, by_branch, by_service, by_pm

def write_excel(outdir, report_date, kpis, by_branch, by_service, by_pm, df_week, errors_df):
    os.makedirs(outdir, exist_ok=True)
    path = os.path.join(outdir, f"weekly_report_{report_date}.xlsx")

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        # Summary sheet (KPIs)
        summary_df = pd.DataFrame(list(kpis.items()), columns=["Metric", "Value"])
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=1)

        # Other sheets
        by_branch.to_excel(writer, sheet_name="By Branch", index=False)
        by_service.to_excel(writer, sheet_name="By Service", index=False)
        by_pm.to_excel(writer, sheet_name="By PM", index=False)
        df_week.to_excel(writer, sheet_name="Raw Data", index=False)
        if not errors_df.empty:
            errors_df.to_excel(writer, sheet_name="Errors", index=False)

        # Simple chart on Summary
        workbook  = writer.book
        worksheet = writer.sheets["Summary"]
        worksheet.write(0, 0, "Weekly Repair Claims — Summary")

        # Add a column chart for By Branch counts
        chart_sheet = "By Branch"
        chart = workbook.add_chart({"type": "column"})
        # Categories (branches) and values (counts); rows start at 1 due to header
        chart.add_series({
            "name":       "Claims by Branch",
            "categories": f"='{chart_sheet}'!$A$2:$A${len(by_branch)+1}",
            "values":     f"='{chart_sheet}'!$B$2:$B${len(by_branch)+1}",
        })
        chart.set_title({"name": "Claims by Branch"})
        chart.set_x_axis({"name": "Branch"})
        chart.set_y_axis({"name": "Count"})

        # Insert chart on Summary
        worksheet.insert_chart("D2", chart)

    return path

def main():
    args = parse_args()

    if args.week_start:
        start = datetime.strptime(args.week_start, "%Y-%m-%d").date()
        end = start + timedelta(days=6)
    else:
        start, end = last_full_week()

    df = load_data(args.input)
    df_week = filter_week(df, start, end)

    errors_df = validate(df_week)
    kpis, by_branch, by_service, by_pm = compute_kpis(
        df_week,
        sla_assign_days=args.sla_assign_days,
        sla_complete_days=args.sla_complete_days,
    )

    report_date = end.isoformat()
    path = write_excel(args.outdir, report_date, kpis, by_branch, by_service, by_pm, df_week, errors_df)

    print(f"Report written to: {path}")
    print(f"Week range: {start} → {end}")
    if not errors_df.empty:
        print(f"Validation issues: {len(errors_df)} (see 'Errors' sheet)")

if __name__ == "__main__":
    main()
