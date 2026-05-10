"""
src/ai_commentary.py  —  OpenAI-powered commentary generator
Generates plain-English MIS narrative, anomaly summaries, and risk briefings
for CFO and leadership consumption.
"""

import json
from datetime import datetime
from openai import OpenAI
import config


client = OpenAI(api_key=config.OPENAI_API_KEY)


# ── Prompt builders ───────────────────────────────────────────────────────────

def _build_executive_prompt(kpis: dict, anomalies: list, risk_flags: list, dept_summary: str) -> str:
    anomaly_text = ""
    if anomalies:
        anomaly_text = "\n".join([
            f"  - {a['metric']} {a['direction']} in {a['department']} on {a['date']}: value={a['value']:,} (z-score={a['z_score']})"
            for a in anomalies[:5]
        ])
    else:
        anomaly_text = "  None detected this period."

    flag_text = ""
    if risk_flags:
        flag_text = "\n".join([f"  - [{f['severity'].upper()}] {f['flag']}: {f['detail']}" for f in risk_flags])
    else:
        flag_text = "  No critical flags this period."

    return f"""You are a senior finance analyst preparing the weekly MIS (Management Information System) 
report for the CFO and leadership team of {config.COMPANY_NAME}.

Write a concise, professional executive summary with exactly THREE paragraphs:

Paragraph 1 — Performance overview:
Summarise overall financial performance. Comment on revenue ({kpis.get('total_revenue', 0):,.0f}), 
gross margin ({kpis.get('gross_margin_pct', 0):.1f}%), net profit ({kpis.get('net_profit', 0):,.0f}), 
and MoM revenue change ({kpis.get('mom_revenue_change_pct', 0):+.1f}%). 
Be direct — state whether the business is ahead or behind.

Paragraph 2 — Budget and operational performance:
Comment on budget variance ({kpis.get('budget_variance_pct', 0):+.1f}% vs plan), 
SLA breach rate ({kpis.get('sla_breach_rate_pct', 0):.1f}%), 
and highlight the worst-performing and best-performing departments:
{dept_summary}

Paragraph 3 — Risks, anomalies, and recommended actions:
Based on the risk flags and anomalies below, call out the top 2-3 concerns leadership 
must act on this week. Be specific — name the metric, the department, and the recommended action.

Risk flags:
{flag_text}

Statistical anomalies detected:
{anomaly_text}

Period: {kpis.get('period_start', 'N/A')} to {kpis.get('period_end', 'N/A')}

Rules:
- Write in plain English — no jargon, no bullet points inside paragraphs
- Be direct and confident, not hedging
- Each paragraph: 3-5 sentences
- Do NOT use markdown headers or formatting
- Sign off with: "Prepared by AI Finance Intelligence | {datetime.today().strftime('%d %b %Y')}"
"""


def _build_anomaly_prompt(anomalies: list) -> str:
    if not anomalies:
        return "No anomalies detected this period. All metrics are within normal ranges."

    anomaly_list = "\n".join([
        f"- {a['metric']} recorded a {a['direction']} in {a['department']} on {a['date']}: "
        f"value = {a['value']:,} (z-score = {a['z_score']}, severity = {a['severity']})"
        for a in anomalies
    ])

    return f"""You are a data analyst. Write a brief, plain-English anomaly alert 
(max 150 words) for a finance team. Explain what the anomaly means in business terms, 
why it might have happened, and what to investigate. No bullet points.

Anomalies detected:
{anomaly_list}
"""


def _build_department_prompt(dept_df) -> str:
    if dept_df is None or dept_df.empty:
        return ""
    rows = dept_df.head(5).to_dict("records")
    lines = []
    for r in rows:
        lines.append(
            f"  {r.get('department', 'N/A')}: revenue={r.get('revenue', 0):,.0f}, "
            f"variance={r.get('variance_pct', 0):+.1f}%, "
            f"gross_margin={r.get('gross_margin_pct', 0):.1f}%"
        )
    return "\n".join(lines)


# ── Commentary generators ─────────────────────────────────────────────────────

def generate_executive_commentary(
    kpis:       dict,
    anomalies:  list,
    risk_flags: list,
    dept_df=None,
) -> str:
    """
    Generate the main 3-paragraph executive MIS commentary using GPT-4o.
    Falls back to a plain summary if OpenAI is unavailable.
    """
    dept_summary = _build_department_prompt(dept_df)
    prompt = _build_executive_prompt(kpis, anomalies, risk_flags, dept_summary)

    try:
        response = client.chat.completions.create(
            model=config.OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=config.OPENAI_MAX_TOKENS,
            temperature=0.4,
        )
        commentary = response.choices[0].message.content.strip()
        print("  AI commentary generated successfully")
        return commentary

    except Exception as e:
        print(f"  [WARNING] OpenAI call failed ({e}) — using fallback commentary")
        return _fallback_commentary(kpis)


def generate_anomaly_summary(anomalies: list) -> str:
    """Generate a short anomaly alert paragraph."""
    prompt = _build_anomaly_prompt(anomalies)
    try:
        response = client.chat.completions.create(
            model=config.OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=200,
            temperature=0.3,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"  [WARNING] Anomaly summary failed ({e})")
        if not anomalies:
            return "No anomalies detected this period."
        return f"{len(anomalies)} anomalies detected. Please review the anomaly table for details."


def generate_department_narrative(dept_df) -> str:
    """Generate a short department performance narrative."""
    if dept_df is None or dept_df.empty:
        return "No department-level data available."

    dept_summary = _build_department_prompt(dept_df)
    prompt = f"""Write a concise 2-3 sentence department performance summary for leadership.
Highlight the best and worst performing departments based on budget variance and gross margin.
Be specific with numbers. No bullet points.

Department data:
{dept_summary}
"""
    try:
        response = client.chat.completions.create(
            model=config.OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=150,
            temperature=0.3,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"  [WARNING] Department narrative failed ({e})")
        return dept_summary


# ── Fallback ──────────────────────────────────────────────────────────────────

def _fallback_commentary(kpis: dict) -> str:
    """Plain-text fallback when OpenAI is unavailable."""
    direction = "ahead of" if kpis.get("budget_variance_pct", 0) >= 0 else "behind"
    return (
        f"This week, {config.COMPANY_NAME} recorded total revenue of "
        f"£{kpis.get('total_revenue', 0):,.0f} with a gross margin of "
        f"{kpis.get('gross_margin_pct', 0):.1f}% and net profit of "
        f"£{kpis.get('net_profit', 0):,.0f}. "
        f"Revenue moved {kpis.get('mom_revenue_change_pct', 0):+.1f}% vs last week. "
        f"Actuals are {direction} budget by {abs(kpis.get('budget_variance_pct', 0)):.1f}%. "
        f"SLA breach rate stands at {kpis.get('sla_breach_rate_pct', 0):.1f}%. "
        f"Please review the full report for department-level detail.\n\n"
        f"Prepared by MIS Report Generator | {datetime.today().strftime('%d %b %Y')}"
    )
