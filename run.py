"""
run.py  —  One-click MIS Report Runner
========================================
Double-click this file OR run:  python run.py

Options:
  python run.py --sample        Use generated sample data (no real DB/Excel needed)
  python run.py --no-email      Generate report but skip emailing
  python run.py --no-powerbi    Skip saving the Power BI trend data export
"""

import sys
import os
import time
from datetime import datetime

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config

# Parse flags
USE_SAMPLE   = "--sample"    in sys.argv
SKIP_EMAIL   = "--no-email"  in sys.argv
SKIP_POWERBI = "--no-powerbi" in sys.argv


def banner(text: str):
    print(f"\n{'─'*50}")
    print(f"  {text}")
    print(f"{'─'*50}")


def step(n: int, label: str):
    print(f"\n[{n}/6] {label}...")


def main():
    start_time = time.time()
    print(f"\n{'='*50}")
    print(f"  MIS REPORT GENERATOR")
    print(f"  {config.COMPANY_NAME}")
    print(f"  {datetime.today().strftime('%A, %d %B %Y  %H:%M')}")
    print(f"{'='*50}")

    if USE_SAMPLE:
        print("  Mode: SAMPLE DATA (demo mode)")
    else:
        print("  Mode: LIVE DATA")

    # ── Step 1: Load data ────────────────────────────────────
    step(1, "Loading data sources")
    from src.data_loader import load_all_data
    df = load_all_data(
        excel_path = "data/raw/actuals.xlsx",
        sql_file   = "sql/queries.sql",
        use_sample = USE_SAMPLE,
    )

    if df.empty:
        print("  [ERROR] No data loaded. Exiting.")
        sys.exit(1)

    print(f"  Data loaded: {len(df)} rows")

    # ── Step 2: Compute KPIs ─────────────────────────────────
    step(2, "Computing KPIs and metrics")
    from src.kpi_engine import (
        compute_kpis,
        compute_department_kpis,
        compute_rolling_trends,
        detect_anomalies,
        get_risk_flags,
    )

    kpis        = compute_kpis(df)
    dept_df     = compute_department_kpis(df)
    trends_df   = compute_rolling_trends(df, weeks=8)
    anomalies   = detect_anomalies(df, columns=[
        config.COL_REVENUE, config.COL_GROSS_PROFIT, config.COL_SLA_BREACHED
    ])
    risk_flags  = get_risk_flags(kpis)

    print(f"  KPIs computed: {len(kpis)} metrics")
    print(f"  Anomalies detected: {len(anomalies)}")
    print(f"  Risk flags raised:  {len(risk_flags)}")

    # ── Step 3: Generate AI commentary ───────────────────────
    step(3, "Generating AI commentary (OpenAI)")
    from src.ai_commentary import (
        generate_executive_commentary,
        generate_anomaly_summary,
    )

    commentary       = generate_executive_commentary(kpis, anomalies, risk_flags, dept_df)
    anomaly_summary  = generate_anomaly_summary(anomalies)

    # ── Step 4: Build Excel report ────────────────────────────
    step(4, "Building Excel MIS report")
    from src.report_builder import build_excel_report

    date_str    = datetime.today().strftime("%Y-%m-%d")
    output_path = os.path.join(config.REPORT_OUTPUT_DIR, f"MIS_Report_{date_str}.xlsx")

    report_path = build_excel_report(
        kpis            = kpis,
        commentary      = commentary,
        anomaly_summary = anomaly_summary,
        risk_flags      = risk_flags,
        anomalies       = anomalies,
        dept_df         = dept_df,
        trends_df       = trends_df,
        output_path     = output_path,
    )

    # ── Step 5: Export Power BI data ──────────────────────────
    if not SKIP_POWERBI:
        step(5, "Exporting Power BI trend data")
        powerbi_path = os.path.join(config.REPORT_OUTPUT_DIR, f"PowerBI_Trends_{date_str}.xlsx")
        if trends_df is not None and not trends_df.empty:
            trends_df.to_excel(powerbi_path, index=False)
            print(f"  Power BI data saved: {powerbi_path}")
        if dept_df is not None and not dept_df.empty:
            dept_path = os.path.join(config.REPORT_OUTPUT_DIR, f"PowerBI_Departments_{date_str}.xlsx")
            dept_df.to_excel(dept_path, index=False)
            print(f"  Department data saved: {dept_path}")
    else:
        print("\n[5/6] Power BI export skipped (--no-powerbi)")

    # ── Step 6: Send email ────────────────────────────────────
    if not SKIP_EMAIL:
        step(6, "Sending email report")
        from src.email_sender import send_report
        send_report(
            kpis            = kpis,
            commentary      = commentary,
            risk_flags      = risk_flags,
            attachment_path = report_path,
        )
    else:
        print("\n[6/6] Email skipped (--no-email)")

    # ── Done ──────────────────────────────────────────────────
    elapsed = round(time.time() - start_time, 1)
    banner(f"COMPLETE  ({elapsed}s)")
    print(f"  Report:   {report_path}")
    print(f"  Metrics:  Revenue £{kpis.get('total_revenue',0):,.0f} | "
          f"Margin {kpis.get('gross_margin_pct',0):.1f}% | "
          f"Variance {kpis.get('budget_variance_pct',0):+.1f}%")
    print(f"  Flags:    {len(risk_flags)} risk flags | {len(anomalies)} anomalies")
    print()


if __name__ == "__main__":
    main()
