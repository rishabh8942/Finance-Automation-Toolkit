# 🤖 Finance Automation Toolkit — AI-Powered MIS Report Generator

> Automatically generates weekly MIS packs (P&L, SLA, KPI summaries) from Excel and SQL data, writes plain-English AI commentary using GPT-4o, and emails a formatted report to leadership — saving 4+ hours of manual work per week.

---

## 🧭 Project Overview

Every finance team has the same Monday morning ritual: pull data from 3 different sources, paste it into Excel, write a summary paragraph, format a report, email it out. This toolkit eliminates that entire workflow.

**One command. Full report. In under 60 seconds.**

```bash
python run.py --sample      # Demo with sample data (no DB needed)
python run.py               # Live run with your Excel/SQL sources
python run.py --no-email    # Generate report only, skip email
```

---

## ✨ What It Does

| Feature | Detail |
|---------|--------|
| **Data ingestion** | Pulls from Excel files, SQL databases, and CSV sources |
| **KPI computation** | Revenue, gross margin, net profit, MoM/YoY change, budget variance, SLA breach rate |
| **Anomaly detection** | Z-score based detection across all metrics — flags unusual spikes and drops |
| **AI commentary** | GPT-4o writes a 3-paragraph CFO-ready executive summary with risk analysis |
| **Excel MIS pack** | 5-sheet formatted report: Executive Summary, P&L, Departments, Anomalies, Trends |
| **Auto email** | HTML + plain-text email with Excel attachment sent to distribution list |
| **Power BI export** | Trend and department data exported to Excel for Power BI refresh |
| **One-click runner** | `run.py` ties everything together — non-technical users just double-click |

---

## 🏗️ Architecture

```
Excel / CSV files  ──┐
                      ├──► data_loader.py ──► kpi_engine.py ──► ai_commentary.py
SQL database       ──┘                                │                │
                                                       │                │
                                               report_builder.py ◄──────┘
                                                       │
                                         ┌─────────────┼─────────────┐
                                         ▼             ▼             ▼
                                   Excel MIS      Power BI      email_sender.py
                                   Report         Export         → CFO, Finance team
```

---

## 📁 Project Structure

```
mis-report-generator/
├── run.py                    ← One-click runner (start here)
├── config.py                 ← All settings (reads from .env)
├── requirements.txt
├── .env.example              ← Copy to .env and fill in your keys
│
├── src/
│   ├── data_loader.py        ← Excel, SQL, CSV ingestion + sample data generator
│   ├── kpi_engine.py         ← KPI computation, anomaly detection, risk flags
│   ├── ai_commentary.py      ← OpenAI GPT-4o commentary generation
│   ├── report_builder.py     ← Formatted Excel workbook builder (5 sheets)
│   └── email_sender.py       ← HTML email + Excel attachment sender
│
├── sql/
│   └── queries.sql           ← Edit to match your database schema
│
├── data/
│   ├── raw/                  ← Drop your Excel/CSV files here
│   └── processed/            ← Cleaned intermediate data
│
└── reports/
    └── output/               ← Generated reports saved here
```

---

## ⚙️ Setup

### 1. Clone the repository
```bash
git clone https://github.com/yourusername/mis-report-generator.git
cd mis-report-generator
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Configure your environment
```bash
cp .env.example .env
# Edit .env with your OpenAI key, database string, and email credentials
```

### 4. Run with sample data (no setup needed)
```bash
python run.py --sample --no-email
```

### 5. Connect your real data
- Drop your Excel actuals file into `data/raw/actuals.xlsx`
- Edit `sql/queries.sql` to match your database schema
- Update column names in `config.py` if needed
- Run: `python run.py`

---

## 📊 Output — Excel MIS Report (5 Sheets)

| Sheet | Contents |
|-------|---------|
| **Executive Summary** | KPI scorecards, AI commentary (3 paragraphs), risk flags with RAG coloring |
| **P&L Summary** | Revenue, COGS, Gross Profit, OPEX, Net Profit vs Budget with variance |
| **Departments** | Per-department breakdown: revenue, budget variance %, gross margin %, SLA breach % |
| **Anomalies** | Detected statistical anomalies with AI-generated explanation |
| **Trends (Power BI)** | 8-week rolling data — point Power BI directly at this sheet |

---

## 🤖 AI Commentary Sample

> *"This week, Acme Corp recorded total revenue of £680,034, representing a 3.2% increase over the prior week and performance broadly in line with budget (-0.2% variance). Gross margin held at 47.4%, reflecting stable cost structures across the Sales and Operations business units, though the Marketing department showed a 6.1% budget overrun driven by Q2 campaign activity.*
>
> *Operationally, the SLA breach rate of 6.8% is approaching the 10% threshold and warrants attention in the IT and Operations departments, which together accounted for 74% of this week's breaches. The Finance department posted the strongest gross margin at 54.2%, while Marketing came in lowest at 38.7%.*
>
> *The primary risk this week is the revenue anomaly detected in Operations on 05 May — a 58% drop vs the weekly average (z-score: 2.74). This requires immediate investigation to determine whether it reflects a data entry issue or a genuine operational disruption. Recommendation: Finance team to review Operations actuals before the Monday leadership call."*

---

## 📈 KPIs Computed

| Category | Metrics |
|---------|--------|
| **Revenue** | Total, MoM change %, YoY change % |
| **P&L** | Gross profit, gross margin %, net profit, net margin %, COGS, OPEX |
| **Budget** | Total budget, actuals, variance £, variance % |
| **Operations** | SLA breaches, breach rate %, total tickets |
| **Anomalies** | Z-score detection across all numeric columns |
| **Risk flags** | Budget overrun, SLA breach, gross margin drop, revenue decline, net loss |

---

## 🔄 Connecting Power BI

1. Open Power BI Desktop
2. `Get Data → Excel Workbook`
3. Point to `reports/output/MIS_Report_YYYY-MM-DD.xlsx`
4. Load the **Trends (Power BI)** and **Departments** sheets
5. Build your visuals — your existing Financial Reporting .pbix is a great starting point
6. Set up a scheduled refresh to auto-update after each `run.py` execution

---

## ⏰ Automating the Weekly Run

**Windows Task Scheduler:**
```
Action: Start a program
Program: python
Arguments: C:\path\to\mis-report-generator\run.py
Schedule: Every Monday at 08:00
```

**Linux / Mac (cron):**
```bash
# Run every Monday at 8am
0 8 * * 1 cd /path/to/mis-report-generator && python run.py
```

---

## 🛠️ Tech Stack

| Tool | Role |
|------|------|
| **Python 3.10+** | Core language |
| **OpenAI GPT-4o** | Executive commentary, anomaly narration |
| **pandas** | Data loading, cleaning, KPI computation |
| **SQLAlchemy** | Database connectivity (SQL Server, MySQL, PostgreSQL, SQLite) |
| **openpyxl** | Excel report generation with formatting |
| **smtplib** | Email delivery (Gmail, Outlook, any SMTP) |
| **Power BI** | Dashboard visualization (reads from report output) |
| **python-dotenv** | Secure credential management |

---

## 🔧 Configuration

All settings in `config.py` / `.env`:

| Setting | Description |
|---------|-------------|
| `OPENAI_API_KEY` | Your OpenAI API key |
| `DB_CONNECTION_STRING` | SQLAlchemy connection string |
| `EMAIL_SENDER` / `EMAIL_PASSWORD` | SMTP credentials |
| `EMAIL_RECIPIENTS` | Comma-separated list of report recipients |
| `BUDGET_VARIANCE_WARN_PCT` | Budget overrun threshold (default: -5%) |
| `SLA_BREACH_WARN_PCT` | SLA alert threshold (default: 10%) |
| `GROSS_MARGIN_WARN_PCT` | Margin floor threshold (default: 30%) |
| `ANOMALY_ZSCORE_THRESHOLD` | Statistical anomaly sensitivity (default: 2.0) |

---

## 👤 Author

**Rishabh Pandey**
- GitHub: [@rishabh8942](https://github.com/rishabh8942)
- LinkedIn: [rishabh-pandey2410](https://www.linkedin.com/in/rishabh-pandey2410/)

---

*Built to eliminate 4+ hours of manual finance reporting every week — fully automated, AI-narrated, and CFO-ready.*
