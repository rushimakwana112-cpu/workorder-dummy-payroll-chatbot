# Workorder Dummy — Payroll AI Chatbot

Ask payroll questions in plain English and get answers straight from a SQL Server table: salary totals, CTC,
margins per work order, TDS projections, month-on-month comparisons and more. Everything runs **on your own PC**:
the database, the AI model (Ollama) and the website. Nothing is sent to the cloud.

> This project runs on **dummy data only**. No real company, client or employee data is included, and the
> "Refresh database" button just explains that there is nothing to refresh from.

## What it can do

- **Plain-English questions**: *salary for sep-26*, *which workorder has highest margin in aug-26*,
  *compare aug-26 and sep-26 ctc and margin*, *projection of oct-26*, *top 10 employees by net salary*.
- **Read-only by design**: the AI only picks which report to run; the SQL is fixed, parameterised and
  checked so it can only `SELECT` from one table.
- **Work-order margins**: base sales value = invoice × 100 / 118 (GST removed), margin = base sales − CTC.
- **TDS tools**: proposed TDS per employee (new / old regime, FY 2026-27), a TDS calculator and a salary run.
- **Users and roles**: master, admin and user, with per-person limits (hidden columns, totals only,
  category / department limits) and an activity log.
- **Two-tone web UI** (Blue Panther theme) with charts, CSV downloads, light / dark mode and query history.

## Tech

Python · Starlette + Uvicorn (web, port 8080) · SQL Server Express · pyodbc · pandas ·
Ollama with `qwen2.5:3b` (answers) and `nomic-embed-text` (matching) · plain HTML/JS front end.

## Setup (Windows)

1. Install: Python 3.12+, SQL Server Express + SSMS, **ODBC Driver 17 for SQL Server**, and Ollama. Then:
   ```bash
   ollama pull qwen2.5:3b
   ```
   ```bash
   ollama pull nomic-embed-text
   ```
   ```bash
   pip install -r requirements.txt
   ```
2. Database: in SSMS run the scripts in `sql\` in order: `00` creates the database `payroll`, then `01`–`05`
   create the table `dbo.Payrollall1` (47 columns), the TDS functions and the helper procedures.
3. Load your own **dummy** data into `dbo.Payrollall1` (same 47 columns, same order), then run once:
   ```sql
   EXEC dbo.usp_RefreshPayrollall @ShowMarginReport = 0;
   ```
4. Point the bot at your server: top of `payroll_chatbot.py`
   ```python
   SERVER = r"localhost\SQLEXPRESS"
   DATABASE = "payroll"
   ```
5. Create the first (master) account: double-click **`Reset Password.bat`**.
6. Start: double-click **`Start Chatbot (no window).vbs`**, wait ~20 s, open <http://localhost:8080>.
   Stop with **`Stop Chatbot.bat`**.

## Settings worth knowing (top of `payroll_chatbot.py`)

| Setting | Default | Meaning |
|---|---|---|
| `DUMMY_DATA` | `True` | "Refresh database" only shows a dummy-data message |
| `REFRESH_SOURCE` | `""` | view/table to reload `Payrollall1` from (same 47 columns, same order) |
| `MODEL` | `qwen2.5:3b` | Ollama model for understanding questions |
| `GST_RATE` | `18` | used for base sales value |

The website listens on the local network (`HOST` in `web_server.py`); set it to `127.0.0.1` to keep it on
this PC only. A fresh install is valid for one year; the master renews it in **Settings → Chatbot validity**.

## Files that stay on your PC

Logins (`users.json`), the activity log, per-install settings, the validity key and logs are listed in
`.gitignore` and are never committed.
