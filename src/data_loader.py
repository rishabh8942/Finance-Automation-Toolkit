"""
src/data_loader.py  —  Load and merge all data sources
Handles: Excel files, SQL database, CSV files
Returns a single clean master DataFrame ready for KPI calculation.
"""

import os
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import config


# ── Excel ────────────────────────────────────────────────────────────────────

def load_excel_data(path: str, sheet_name: str = 0) -> pd.DataFrame:
    """
    Load an Excel file and return a normalised DataFrame.
    Column names are lowercased and stripped of whitespace.
    """
    if not os.path.exists(path):
        print(f"  [WARNING] Excel file not found: {path}")
        return pd.DataFrame()

    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    if config.COL_DATE in df.columns:
        df[config.COL_DATE] = pd.to_datetime(df[config.COL_DATE], errors="coerce")
        df = df.dropna(subset=[config.COL_DATE])

    print(f"  Loaded {len(df)} rows from {os.path.basename(path)}")
    return df


def load_multiple_excel(folder: str) -> pd.DataFrame:
    """Load and concatenate all .xlsx files from a folder."""
    frames = []
    for fname in os.listdir(folder):
        if fname.endswith(".xlsx") and not fname.startswith("~"):
            frames.append(load_excel_data(os.path.join(folder, fname)))
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# ── SQL ──────────────────────────────────────────────────────────────────────

def load_sql_data(query: str) -> pd.DataFrame:
    """
    Execute a SQL query and return results as a DataFrame.
    Connection string is read from config (pulled from .env).
    """
    try:
        engine = create_engine(config.DB_CONNECTION_STRING)
        with engine.connect() as conn:
            df = pd.read_sql(text(query), conn)
        print(f"  Loaded {len(df)} rows from SQL")
        return df
    except Exception as e:
        print(f"  [WARNING] SQL connection failed: {e}")
        return pd.DataFrame()


def load_sql_from_file(sql_file: str) -> pd.DataFrame:
    """Load SQL from a .sql file and execute it."""
    if not os.path.exists(sql_file):
        print(f"  [WARNING] SQL file not found: {sql_file}")
        return pd.DataFrame()
    with open(sql_file, "r") as f:
        query = f.read()
    return load_sql_data(query)


# ── CSV ──────────────────────────────────────────────────────────────────────

def load_csv_data(path: str) -> pd.DataFrame:
    """Load a CSV file and normalise column names."""
    if not os.path.exists(path):
        print(f"  [WARNING] CSV file not found: {path}")
        return pd.DataFrame()
    df = pd.read_csv(path)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    if config.COL_DATE in df.columns:
        df[config.COL_DATE] = pd.to_datetime(df[config.COL_DATE], errors="coerce")
    print(f"  Loaded {len(df)} rows from {os.path.basename(path)}")
    return df


# ── Sample data generator (for demo / testing) ───────────────────────────────

def generate_sample_data(weeks: int = 12) -> pd.DataFrame:
    """
    Generate realistic sample finance and ops data.
    Used when no real data source is available.
    Run:  python -c "from src.data_loader import generate_sample_data; generate_sample_data().to_excel('data/raw/sample_actuals.xlsx', index=False)"
    """
    import numpy as np
    np.random.seed(42)

    today = datetime.today()
    dates = [today - timedelta(weeks=i) for i in range(weeks, 0, -1)]
    departments = ["Sales", "Operations", "Finance", "Marketing", "IT"]
    regions = ["North", "South", "East", "West"]

    rows = []
    for d in dates:
        for dept in departments:
            base_rev    = {"Sales": 120000, "Operations": 80000, "Finance": 40000,
                           "Marketing": 55000, "IT": 30000}[dept]
            base_budget = base_rev * 1.05

            revenue     = base_rev * np.random.uniform(0.85, 1.20)
            budget      = base_budget * np.random.uniform(0.98, 1.02)
            cogs        = revenue * np.random.uniform(0.45, 0.60)
            opex        = revenue * np.random.uniform(0.15, 0.25)
            gross_profit= revenue - cogs
            net_profit  = gross_profit - opex
            sla_breached= int(np.random.poisson(2))
            tickets     = int(np.random.poisson(30))

            rows.append({
                config.COL_DATE:         d,
                config.COL_DEPARTMENT:   dept,
                config.COL_REGION:       np.random.choice(regions),
                config.COL_REVENUE:      round(revenue, 2),
                config.COL_BUDGET:       round(budget, 2),
                config.COL_ACTUALS:      round(revenue, 2),
                config.COL_COGS:         round(cogs, 2),
                config.COL_OPEX:         round(opex, 2),
                config.COL_GROSS_PROFIT: round(gross_profit, 2),
                config.COL_NET_PROFIT:   round(net_profit, 2),
                config.COL_SLA_BREACHED: sla_breached,
                config.COL_SLA_TARGET:   5,
                config.COL_TICKETS:      tickets,
            })

    df = pd.DataFrame(rows)
    # Inject one obvious anomaly for demo
    df.loc[df.index[5], config.COL_REVENUE] *= 0.40
    print(f"  Generated {len(df)} rows of sample data")
    return df


# ── Master loader ─────────────────────────────────────────────────────────────

def load_all_data(
    excel_path: str    = "data/raw/actuals.xlsx",
    sql_file:   str    = "sql/queries.sql",
    use_sample: bool   = False,
) -> pd.DataFrame:
    """
    Master entry point — loads all sources and returns one clean DataFrame.
    Set use_sample=True to use generated demo data.
    """
    if use_sample:
        print("  Using generated sample data")
        return generate_sample_data()

    frames = []

    excel_df = load_excel_data(excel_path)
    if not excel_df.empty:
        frames.append(excel_df)

    sql_df = load_sql_from_file(sql_file)
    if not sql_df.empty:
        frames.append(sql_df)

    if not frames:
        print("  No data sources found — falling back to sample data")
        return generate_sample_data()

    df = pd.concat(frames, ignore_index=True)
    df = df.sort_values(config.COL_DATE).reset_index(drop=True)
    print(f"  Master DataFrame: {len(df)} rows, {df.columns.tolist()}")
    return df
