"""
config.py  —  Central configuration for MIS Report Generator
All settings pulled from .env; change values there, not here.
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ── OpenAI ───────────────────────────────────────────────
OPENAI_API_KEY       = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL         = "gpt-4o"
OPENAI_MAX_TOKENS    = 600

# ── Database ─────────────────────────────────────────────
DB_CONNECTION_STRING = os.getenv("DB_CONNECTION_STRING", "sqlite:///data/raw/finance.db")

# ── Email ────────────────────────────────────────────────
EMAIL_SENDER         = os.getenv("EMAIL_SENDER", "")
EMAIL_PASSWORD       = os.getenv("EMAIL_PASSWORD", "")
EMAIL_RECIPIENTS     = [r.strip() for r in os.getenv("EMAIL_RECIPIENTS", "").split(",") if r.strip()]
SMTP_HOST            = "smtp.gmail.com"
SMTP_PORT            = 587

# ── Report ───────────────────────────────────────────────
COMPANY_NAME         = os.getenv("COMPANY_NAME", "Your Company")
REPORT_OUTPUT_DIR    = os.getenv("REPORT_OUTPUT_DIR", "reports/output")
REPORT_PERIOD        = os.getenv("REPORT_PERIOD", "weekly")

# ── KPI Thresholds (edit to match your business rules) ───
BUDGET_VARIANCE_WARN_PCT   = -5.0    # flag if actuals deviate > 5% below budget
SLA_BREACH_WARN_PCT        = 10.0   # flag if SLA breach rate > 10%
ANOMALY_ZSCORE_THRESHOLD   = 2.0    # z-score to classify a data point as anomaly
GROSS_MARGIN_WARN_PCT      = 30.0   # flag if gross margin drops below 30%

# ── Column Name Mapping (match your actual Excel/SQL column names) ──
COL_DATE          = "date"
COL_REVENUE       = "revenue"
COL_BUDGET        = "budget"
COL_ACTUALS       = "actuals"
COL_GROSS_PROFIT  = "gross_profit"
COL_COGS          = "cogs"
COL_OPEX          = "opex"
COL_NET_PROFIT    = "net_profit"
COL_DEPARTMENT    = "department"
COL_REGION        = "region"
COL_SLA_BREACHED  = "sla_breached"
COL_SLA_TARGET    = "sla_target"
COL_TICKETS       = "tickets"
