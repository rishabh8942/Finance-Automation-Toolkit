"""
src/kpi_engine.py  —  Compute KPIs, variances, MoM/YoY growth, anomaly detection
Returns structured dicts consumed by ai_commentary.py and report_builder.py
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import config


# ── Helpers ──────────────────────────────────────────────────────────────────

def safe_div(numerator: float, denominator: float, default: float = 0.0) -> float:
    """Division that never raises ZeroDivisionError."""
    return (numerator / denominator) if denominator else default


def pct_change(current: float, previous: float) -> float:
    """Percentage change from previous to current."""
    return safe_div((current - previous), abs(previous)) * 100


def filter_period(df: pd.DataFrame, weeks_back: int = 1) -> pd.DataFrame:
    """Return rows from the last N weeks."""
    cutoff = datetime.today() - timedelta(weeks=weeks_back)
    return df[df[config.COL_DATE] >= cutoff]


def filter_last_n_weeks(df: pd.DataFrame, n: int) -> pd.DataFrame:
    cutoff = datetime.today() - timedelta(weeks=n)
    return df[df[config.COL_DATE] >= cutoff]


# ── Core KPI computation ──────────────────────────────────────────────────────

def compute_kpis(df: pd.DataFrame) -> dict:
    """
    Compute all top-level KPIs for the current period.
    Returns a flat dict of named metrics.
    """
    if df.empty:
        return {}

    current  = filter_period(df, weeks_back=1)
    previous = filter_period(df, weeks_back=2)
    previous = previous[previous[config.COL_DATE] < (datetime.today() - timedelta(weeks=1))]

    def col_sum(frame, col):
        return frame[col].sum() if col in frame.columns else 0.0

    # Revenue & P&L
    rev_curr = col_sum(current,  config.COL_REVENUE)
    rev_prev = col_sum(previous, config.COL_REVENUE)
    gp_curr  = col_sum(current,  config.COL_GROSS_PROFIT)
    np_curr  = col_sum(current,  config.COL_NET_PROFIT)
    cogs     = col_sum(current,  config.COL_COGS)
    opex     = col_sum(current,  config.COL_OPEX)

    # Budget
    budget   = col_sum(current,  config.COL_BUDGET)
    actuals  = col_sum(current,  config.COL_ACTUALS)

    # SLA / Ops
    sla_breaches = col_sum(current, config.COL_SLA_BREACHED)
    tickets      = col_sum(current, config.COL_TICKETS)

    kpis = {
        # Revenue
        "total_revenue":          round(rev_curr, 2),
        "prev_revenue":           round(rev_prev, 2),
        "mom_revenue_change_pct": round(pct_change(rev_curr, rev_prev), 2),

        # P&L
        "gross_profit":           round(gp_curr, 2),
        "gross_margin_pct":       round(safe_div(gp_curr, rev_curr) * 100, 2),
        "net_profit":             round(np_curr, 2),
        "net_margin_pct":         round(safe_div(np_curr, rev_curr) * 100, 2),
        "cogs":                   round(cogs, 2),
        "opex":                   round(opex, 2),

        # Budget
        "total_budget":           round(budget, 2),
        "total_actuals":          round(actuals, 2),
        "budget_variance":        round(actuals - budget, 2),
        "budget_variance_pct":    round(pct_change(actuals, budget), 2),

        # Ops / SLA
        "sla_breaches":           int(sla_breaches),
        "total_tickets":          int(tickets),
        "sla_breach_rate_pct":    round(safe_div(sla_breaches, tickets) * 100, 2),

        # Metadata
        "period_start": current[config.COL_DATE].min().strftime("%d %b %Y") if not current.empty else "N/A",
        "period_end":   current[config.COL_DATE].max().strftime("%d %b %Y") if not current.empty else "N/A",
        "rows_processed": len(current),
    }

    return kpis


# ── Department breakdown ──────────────────────────────────────────────────────

def compute_department_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a DataFrame of KPIs broken down by department.
    Columns: department, revenue, budget, variance_pct, gross_margin_pct, sla_breach_rate_pct
    """
    if df.empty or config.COL_DEPARTMENT not in df.columns:
        return pd.DataFrame()

    current = filter_period(df, weeks_back=1)

    grp = current.groupby(config.COL_DEPARTMENT).agg(
        revenue      =(config.COL_REVENUE,      "sum"),
        budget       =(config.COL_BUDGET,       "sum"),
        actuals      =(config.COL_ACTUALS,      "sum"),
        gross_profit =(config.COL_GROSS_PROFIT, "sum"),
        sla_breaches =(config.COL_SLA_BREACHED, "sum"),
        tickets      =(config.COL_TICKETS,      "sum"),
    ).reset_index()

    grp["variance_pct"]       = grp.apply(lambda r: round(pct_change(r.actuals, r.budget), 2), axis=1)
    grp["gross_margin_pct"]   = grp.apply(lambda r: round(safe_div(r.gross_profit, r.revenue) * 100, 2), axis=1)
    grp["sla_breach_rate_pct"]= grp.apply(lambda r: round(safe_div(r.sla_breaches, r.tickets) * 100, 2), axis=1)

    return grp.sort_values("revenue", ascending=False)


# ── Rolling trends ────────────────────────────────────────────────────────────

def compute_rolling_trends(df: pd.DataFrame, weeks: int = 8) -> pd.DataFrame:
    """
    Weekly revenue, gross profit, and budget variance over the last N weeks.
    Used for trend charts in Power BI and the report.
    """
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df["week"] = df[config.COL_DATE].dt.to_period("W").apply(lambda p: p.start_time)

    trends = df.groupby("week").agg(
        revenue      =(config.COL_REVENUE,      "sum"),
        budget       =(config.COL_BUDGET,       "sum"),
        actuals      =(config.COL_ACTUALS,      "sum"),
        gross_profit =(config.COL_GROSS_PROFIT, "sum"),
        net_profit   =(config.COL_NET_PROFIT,   "sum"),
        sla_breaches =(config.COL_SLA_BREACHED, "sum"),
    ).reset_index()

    trends["budget_variance_pct"] = trends.apply(
        lambda r: round(pct_change(r.actuals, r.budget), 2), axis=1
    )
    trends["gross_margin_pct"] = trends.apply(
        lambda r: round(safe_div(r.gross_profit, r.revenue) * 100, 2), axis=1
    )

    cutoff = datetime.today() - timedelta(weeks=weeks)
    return trends[trends["week"] >= cutoff].tail(weeks)


# ── Anomaly detection ─────────────────────────────────────────────────────────

def detect_anomalies(df: pd.DataFrame, columns: list = None) -> list:
    """
    Detect statistical anomalies using z-score across specified columns.
    Returns a list of dicts with date, department, column, value, and severity.
    """
    if df.empty:
        return []

    if columns is None:
        columns = [config.COL_REVENUE, config.COL_GROSS_PROFIT, config.COL_SLA_BREACHED]

    anomalies = []
    threshold  = config.ANOMALY_ZSCORE_THRESHOLD

    for col in columns:
        if col not in df.columns:
            continue

        mean = df[col].mean()
        std  = df[col].std()
        if std == 0:
            continue

        df["_z"] = (df[col] - mean) / std
        outliers  = df[abs(df["_z"]) > threshold]

        for _, row in outliers.iterrows():
            severity = "high" if abs(row["_z"]) > threshold * 1.5 else "medium"
            dept     = row.get(config.COL_DEPARTMENT, "All")
            date_str = row[config.COL_DATE].strftime("%d %b %Y") if pd.notna(row[config.COL_DATE]) else "Unknown"

            anomalies.append({
                "date":       date_str,
                "department": dept,
                "metric":     col.replace("_", " ").title(),
                "value":      round(row[col], 2),
                "z_score":    round(abs(row["_z"]), 2),
                "direction":  "spike" if row["_z"] > 0 else "drop",
                "severity":   severity,
            })

        df.drop(columns=["_z"], inplace=True)

    anomalies.sort(key=lambda x: x["z_score"], reverse=True)
    return anomalies[:10]  # return top 10 anomalies


# ── Business rule flags ───────────────────────────────────────────────────────

def get_risk_flags(kpis: dict) -> list:
    """
    Apply business threshold rules and return a list of risk flags.
    Edit thresholds in config.py.
    """
    flags = []

    if kpis.get("budget_variance_pct", 0) < config.BUDGET_VARIANCE_WARN_PCT:
        flags.append({
            "flag":    "Budget overrun",
            "detail":  f"Actuals are {abs(kpis['budget_variance_pct']):.1f}% below budget",
            "severity":"high",
        })

    if kpis.get("sla_breach_rate_pct", 0) > config.SLA_BREACH_WARN_PCT:
        flags.append({
            "flag":    "SLA breach rate elevated",
            "detail":  f"SLA breach rate is {kpis['sla_breach_rate_pct']:.1f}% (threshold: {config.SLA_BREACH_WARN_PCT}%)",
            "severity":"high",
        })

    if kpis.get("gross_margin_pct", 100) < config.GROSS_MARGIN_WARN_PCT:
        flags.append({
            "flag":    "Gross margin below threshold",
            "detail":  f"Gross margin at {kpis['gross_margin_pct']:.1f}% — below the {config.GROSS_MARGIN_WARN_PCT}% minimum",
            "severity":"medium",
        })

    if kpis.get("mom_revenue_change_pct", 0) < -10:
        flags.append({
            "flag":    "Revenue decline",
            "detail":  f"Revenue dropped {abs(kpis['mom_revenue_change_pct']):.1f}% vs previous week",
            "severity":"medium",
        })

    if kpis.get("net_profit", 1) < 0:
        flags.append({
            "flag":    "Net loss recorded",
            "detail":  f"Net profit is negative: £{kpis['net_profit']:,.0f}",
            "severity":"high",
        })

    return flags
