"""
Payroll Chatbot - Finance AI (runs 100% on your PC)  - v13 (policy engine: per-user memory + admin data permissions; planner: Month = salary month, X-wise = X + salary month; query planner: whole question -> one plan; several categories at once, other than wox%; proposed TDS per employee from TDSOpt: Monthly TDS; open vem = OPENVEM; Qwen checks words the rules missed; projection: several months as columns; projection of future months like the proposed-data Excel; admin Refresh data button; every setting saved per user; without X / except X + category switches; work order-wise per month, this/last month; every work order report, download ALL rows as CSV; fast: instant answers, AI explains only when asked, query cache; chart/sort/headline follow the asked figure; hides PF/take-home/PT/ESIC/gross/TDS unless asked; margin % on CTC, calculated take-home, column guide; ChatGPT-style answers + follow-ups; months as MMM-YYYY, AI uses prepared view p; faster: fewer AI calls, employee ranking; new Workorder Dummy look + per-user light/dark, AI analyst mode, partial dept/category names, base sales in summaries, login system, per-user theme, asks back when unsure, base sales = invoice x 100/118)
Every question is understood by the AI (Ollama / Qwen) when ALWAYS_AI = True.
The AI only decides WHICH report to run; SQL stays fixed and read-only.
If Ollama is not reachable, the keyword rules are used as a backup.

Run with:   streamlit run payroll_chatbot.py
"""

import re
import json
import time
import functools
import html
import warnings
import datetime
from decimal import Decimal

import pandas as pd
import pyodbc
import ollama as _ollama_lib
import streamlit as st

warnings.filterwarnings("ignore", message=".*pandas only supports SQLAlchemy.*")

# ======================= CONFIG - CHANGE THESE =======================
SERVER = r"localhost\SQLEXPRESS"     # same server name you use in SSMS
DATABASE = "payroll"     # <-- your database name
TABLE = "Payrollall1"           # <-- your salary table name (must be the table dbo.usp_RefreshPayrollall fills)
DRIVER = "ODBC Driver 17 for SQL Server"
MODEL = "qwen2.5:3b"                 # ~2-3x faster on this CPU-only PC; more accurate but slower: "qwen2.5-coder:7b"
EMBED_MODEL = "nomic-embed-text"     # v14: matches response rules by meaning ("ollama pull nomic-embed-text"); "" = off
ROW_LIMIT = 1000                     # max rows shown for lists
CSV_LIMIT = 200000                   # max rows in a "Download all rows" CSV
SHOW_LAST_MESSAGES = 12              # older chat messages are hidden (keeps page fast)
# Work order value = INVOICE BASE SALES VALUE (not the invoice value).
# InvoiceAmount includes GST:  invoice 118 = base sales 100  ->  base = InvoiceAmount * 100 / 118
GST_RATE = 18                        # change if the GST rate is different
BASE_EXPR = f"[InvoiceAmount] * 100.0 / {100 + GST_RATE}"
# Take-home is CALCULATED (NetSalary / ChequeAmount in the table are not reliable):
# Take-home = Gross - TDS - PF (employee) - PT - ESIC (employee) - other Deduction
TAKEHOME_EXPR = ("(ISNULL([GrossAmount],0) - ISNULL([IncomeTax],0) - ISNULL([PF_Employee],0) - ISNULL([PT],0)"
                 " - ISNULL([ESIC_Employee],0) - ISNULL([Deduction],0))")
# Margin % (company formula) = Margin / CTC x 100.  Ask "margin % on base sales" for Margin / Base sales x 100.
MARGIN_PCT_DEFAULT = "ctc"
# InvoiceAmount is repeated on EVERY employee row of a work order (full value on each row),
# so the base sales value is taken ONCE per work order per month, while CTC is added up per employee.
# Margin = base sales value - CTC  (uses the column names produced by the query below)
MARGIN_FORMULA = "SUM([BaseSalesValue]) - SUM([CTC])"
# Optional read-only SQL login (see readonly_user_setup.sql). Leave blank to use Windows login.
SQL_USER = ""
SQL_PASSWORD = ""
SESSION_MINUTES = 0                  # 0 = never log out automatically (e.g. 30 = log out after 30 idle minutes)
MAX_LOGIN_TRIES = 5                  # wrong passwords before the ID is locked for 15 minutes
DEFAULT_THEME = "light"              # "light" or "dark" (can switch in the sidebar)
import os as _os
AI_THREADS = max(2, (_os.cpu_count() or 4) - 1)   # Ollama leaves 1 CPU core free so the PC doesn't lag
AI_TIMEOUT_SEC = 120                 # any single AI call is stopped after this many seconds (the bot then
                                     #   answers with the keyword rules or asks back, instead of never answering)
EXPLAIN_MODEL = ""                   # model for the written answer; "" = use qwen2.5:3b if installed (faster), else MODEL
AI_EXPLAIN = "ask"                   # "ask" = AI writes a long explanation only when the question says
                                     #   explain / why / insights ... (fast); "always" = every answer (slow)
CHAT_STYLE = True                    # ChatGPT-style: AI explains each answer in plain words (typing effect)
ANALYST_MODE = True                  # AI writes its own (checked, read-only) SQL for questions no fixed report covers
QUERY_TIMEOUT_SEC = 30               # any single query is stopped after this many seconds
# ---- Data refresh (ADMIN ONLY button). This is the ONLY place the app is allowed to change data. ----
# REFRESH_SOURCE = a view or table with the same 47 columns IN THE SAME ORDER as Payrollall1 (the copy is by
# position), e.g. "[MyServer].[MyDatabase].dbo.vw_PayrollSource". Empty = no refresh source (dummy data).
REFRESH_SOURCE = ""
REFRESH_PROC = "dbo.usp_RefreshPayrollall"        # fills the last two columns after the copy
REFRESH_TIMEOUT_SEC = 900                          # copying from the other server can take a few minutes
REFRESH_ALL_OR_NOTHING = True                      # True: if any step fails, payrollall stays exactly as it was
# Dummy data mode: the table holds made-up test data, so "Refresh database" only shows this message.
# Set DUMMY_DATA = False when the bot runs on the organisation's live data again.
DUMMY_DATA = True
DUMMY_DATA_MSG = ("This chatbot is running on dummy data, so it has no data update system. "
                  "Refresh database works only with the organisation's live data.")
ALWAYS_AI = True                     # True = every question goes to the AI (can switch in sidebar)
# =====================================================================

RAW_T = f"[{TABLE}]"                  # the real table - only used by admin refresh and internal lookups
ollama = _ollama_lib.Client(timeout=AI_TIMEOUT_SEC)   # same calls as the ollama module, but they can't hang forever
import threading as _threading
import importlib as _importlib
import policy_engine as pe
import teaching as tg
import reminders as rmd
import workorder_groups as wg


def _fresh_module(mod):
    """Streamlit keeps imported files in memory while it runs. If policy_engine.py / teaching.py were updated
    on disk, load the new version (otherwise the app crashes until it is restarted)."""
    try:
        m = _os.path.getmtime(mod.__file__)
        if getattr(mod, "_pb_loaded_mtime", None) != m:
            mod = _importlib.reload(mod)
            mod._pb_loaded_mtime = m
    except Exception:
        pass
    return mod


pe = _fresh_module(pe)
tg = _fresh_module(tg)
rmd = _fresh_module(rmd)
wg = _fresh_module(wg)
WO_GROUPS = wg.WorkOrderGroups(_os.path.dirname(_os.path.abspath(__file__)))   # NORTH / SOUTH ... saved in chat by admins
_POL = _threading.local()            # the signed-in user's effective policy for THIS request (per thread)


def _cast(t):
    t = str(t).lower()
    if t in ("int", "bigint", "smallint", "tinyint", "bit"):
        return "int"
    if t in ("decimal", "numeric", "float", "real", "money", "smallmoney"):
        return "float"
    if "date" in t or "time" in t:
        return "datetime"
    return "nvarchar"


def _table_columns():
    try:
        cols = table_schema()
        if cols:
            return cols
    except Exception:
        pass
    return [(c, "nvarchar") for c in pe.KNOWN_COLUMNS]


DATA_FROM_FILE = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "data_from.json")
_DATA_FROM = {"mtime": None, "val": {}}


def data_from():
    """{'from': 'YYYY-MM-01', 'by', 'at'} set by the master / an admin: EVERYONE (master too) sees only rows with
    liable month [month_year] from that month on. {} = all data. Re-read only when the file changes."""
    try:
        m = _os.path.getmtime(DATA_FROM_FILE)
    except OSError:
        return {}
    if _DATA_FROM["mtime"] != m:
        try:
            with open(DATA_FROM_FILE, encoding="utf-8") as fh:
                d = json.load(fh)
            datetime.date.fromisoformat(d["from"])
            _DATA_FROM.update(mtime=m, val=d)
        except Exception:
            _DATA_FROM.update(mtime=m, val={})
    return _DATA_FROM["val"]


def set_data_from(day, user):
    """day = date (any day of the month) -> show data from that month on; None -> show all data again."""
    if day is None:
        try:
            _os.remove(DATA_FROM_FILE)
        except FileNotFoundError:
            pass
        return {}
    d = {"from": day.replace(day=1).isoformat(), "by": user, "at": datetime.datetime.now().strftime("%d-%b-%Y %H:%M")}
    tmp = DATA_FROM_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, indent=2)
    _os.replace(tmp, DATA_FROM_FILE)
    return d


def _row_where():
    """' WHERE [Category] IN (...) AND [Department] IN (...)' for a user limited to some rows by the admin, plus
    '[month_year] >= ...' when a 'show data from' month is set, else ''.
    Values come from the admin's policy (picked from the real lists) and are quoted here, never from a question."""
    rows = getattr(_POL, "row_filter", None) or {}
    parts = []
    since = data_from().get("from")
    if since:                                   # validated ISO date (data_from), safe as a literal
        parts.append(f"[month_year] >= '{datetime.date.fromisoformat(since).isoformat()}'")
    for col in ("Category", "Department"):
        vals = [str(v) for v in rows.get(col) or []]
        if vals:
            parts.append(f"[{col}] IN (" + ", ".join("N'" + v.replace("'", "''") + "'" for v in vals) + ")")
    return (" WHERE " + " AND ".join(parts)) if parts else ""


class _PolicyTable:
    """The payroll table AS THE CURRENT USER MAY SEE IT. Every query built by the app reads FROM {T}, so
    restricted columns come back as NULL from SQL Server itself - they never reach Python."""

    def __str__(self):
        null = getattr(_POL, "null_cols", None)
        where = _row_where()
        if not null and not where:
            return RAW_T
        parts = [f"CAST(NULL AS {_cast(t)}) AS [{c}]" if c in (null or ()) else f"[{c}]" for c, t in _table_columns()]
        return f"(SELECT {', '.join(parts)} FROM {RAW_T}{where}) AS {RAW_T}"

    def __format__(self, spec):
        return str(self)


T = _PolicyTable()
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_MAP = {m.lower(): i + 1 for i, m in enumerate(MON)}
MONTH_WORDS = {"jan", "january", "feb", "february", "mar", "march", "apr", "april",
               "may", "jun", "june", "jul", "july", "aug", "august", "sep", "sept",
               "september", "oct", "october", "nov", "november", "dec", "december"}

REPORTS = {
    "summary": "total payroll / total salary for month(s). DEFAULT if unsure",
    "employee_list": "list of every employee with their salary",
    "department": "salary totals grouped by department",
    "category": "salary totals grouped by category",
    "deductions": "TDS / income tax, PF, PT, ESIC, deduction totals",
    "components": "salary breakup: basic, HRA, allowances, bonus, arrears",
    "employee_detail": "full salary details of ONE specific employee (name or code)",
    "workorder_detail": "all rows/details of ONE specific work order number (e.g. M244425094)",
    "workorder_summary": "totals PER work order with base sales value, CTC and MARGIN; use for highest/lowest/top work orders, margin, profit, work order wise",
    "trend": "month-wise trend / comparison over time",
    "joiners": "new joinees / employees who joined",
    "leavers": "employees who resigned / left",
    "chitchat": "greeting, thanks, small talk, or anything NOT asking for payroll data",
    "analysis": "ANY other data question: conditions on any column (dates like WOEndDate/joindate, greater/less than, before/after), counts, how many, averages, custom lists, comparisons no other report covers",
}

# keyword rules (checked in this order) - no AI needed when these match
REPORT_TITLES = {
    "summary": "Payroll summary", "employee_list": "Employee-wise list", "department": "Department-wise",
    "category": "Category-wise", "deductions": "Deductions", "components": "Salary components",
    "employee_detail": "Employee details", "workorder_detail": "Work order details",
    "workorder_summary": "Work order ranking / margin", "workorder_monthly": "Work order-wise (per month)",
    "projection": "Projection (future month)",
    "trend": "Monthly trend", "analysis": "AI analysis",
    "employee_ranking": "Employee ranking", "definition": "Definition / formula", "pref": "Setting saved",
    "policy": "Data permissions", "policy_confirm": "Confirm permission change", "admin_info": "Administrator info",
    "plan": "Query plan",
    "joiners": "New joiners", "leavers": "Leavers",
}

MARGIN_WORDS = {"margin", "margins", "profit", "profitability", "markup"}
RANK_WORDS = {"highest", "lowest", "top", "bottom", "max", "maximum", "min", "minimum", "most",
              "least", "best", "worst", "rank", "ranking", "biggest", "smallest", "wise",
              "each", "per", "workorderwise", "compare"}
LOW_WORDS = {"lowest", "bottom", "min", "minimum", "least", "worst", "smallest"}


WO_LIST_WORDS = {"every", "all", "each", "list", "details", "detail", "wise", "data", "csv", "excel", "workorderwise"}


def wants_wo_list(question):
    """'give me data workorder wise', 'details of every work order of aug-26' -> one row per WO per month."""
    t = set(re.findall(r"[a-z]+", question.lower()))
    if not (WO_WORDS.search(question) or "workorderwise" in t) or find_workorder(question)[0]:
        return False
    if t & {"department", "departments", "dept", "category", "categories", "employee", "employees", "emp"}:
        return False
    rank = RANK_WORDS - {"wise", "each", "per", "compare"}
    return bool(t & WO_LIST_WORDS) and not (t & rank) and not re.search(r"\btop\s+\d+|\bbottom\s+\d+", question.lower())


def wants_wo_summary(question):
    """True for 'which work order has highest margin', 'work order wise salary', 'top 5 work orders'."""
    t = set(re.findall(r"[a-z]+", question.lower()))
    if t & {"department", "departments", "dept", "category", "categories", "employee", "employees", "emp"}:
        return False                              # "category wise margin" -> category report
    if re.search(r"\bwhich\s+months?\b|\bmonths?\s+(?:had|has|with|was)\b|\bbest\s+month\b|\bworst\s+month\b", question, re.I):
        return False                              # "which month had the highest margin" -> monthly trend
    if WO_WORDS.search(question):
        if t & {"every", "all", "each", "list", "details", "detail", "wise", "summary", "data", "csv", "excel"} \
                and not find_workorder(question)[0]:
            return True                            # "details of every work order of aug-26"
        return bool(t & (RANK_WORDS | MARGIN_WORDS))
    # "margin for aug-26" = the month's total margin (summary); "highest margin" = work order ranking
    return bool(t & MARGIN_WORDS) and bool(t & (RANK_WORDS - {"wise", "each", "per", "compare"}))


def ranking(question):
    t = question.lower()
    top = re.search(r"\b(?:top|bottom|first|last)\s+(\d{1,4})\b", t)
    return ("asc" if set(re.findall(r"[a-z]+", t)) & LOW_WORDS else "desc",
            int(top.group(1)) if top else None)


RULES = [
    ("joiners", {"joined", "joiner", "joiners", "joinee", "joinees", "joining"}),
    ("leavers", {"resigned", "resign", "resignation", "resignations", "left", "leaver", "leavers", "exit"}),
    ("deductions", {"tds", "tax", "pf", "esic", "pt", "deduction", "deductions", "nps", "deducted", "deduct", "statutory"}),
    ("components", {"basic", "hra", "allowance", "allowances", "breakup", "bonus", "arrear", "arrears", "component", "components"}),
    ("department", {"department", "departments", "dept", "departmentwise"}),
    ("category", {"category", "categories", "categorywise"}),
    ("trend", {"trend", "compare", "comparison", "monthwise"}),
    ("employee_list", {"list", "each", "all", "employeewise"}),
]

# words that are safe to ignore - if a question has only these, AI is skipped
SIMPLE_WORDS = {
    "give", "me", "show", "get", "what", "is", "was", "the", "a", "of", "for", "in", "on",
    "please", "pls", "salary", "salaries", "payroll", "total", "net", "gross", "ctc",
    "month", "months", "year", "wise", "and", "by", "from", "to", "till", "until", "between",
    "employee", "employees", "emp", "who", "new", "details", "detail", "report", "data",
    "amount", "paid", "how", "much", "sum", "can", "you", "tell", "i", "want", "need",
    "income", "summary", "last", "this", "with", "find", "search", "work", "order",
    "workorder", "wo", "number", "no", "info", "information", "about", "all", "whose", "whom", "person",
    "people", "staff", "top", "highest", "lowest", "maximum", "minimum", "max", "min", "most", "least",
    "bottom", "biggest", "smallest", "take", "home", "takehome",
    "proper", "exact", "actual", "correct", "right", "value", "values", "amount", "amounts",
    "figure", "figures", "overall", "complete", "whole", "only", "also", "plus", "both",
    "workorders", "orders", "starting", "starts", "start", "beginning", "begins", "begin",
    "containing", "contains", "contain", "like", "ending", "ends", "end",
    "which", "has", "have", "had", "highest", "lowest", "top", "bottom", "max", "maximum",
    "min", "minimum", "most", "least", "best", "worst", "margin", "margins", "profit",
    "invoice", "billing", "each", "per", "rank", "ranking", "biggest", "smallest",
    "base", "sales", "sale", "value", "revenue", "many", "count",
    "liable", "liability", "payment", "disbursed",   # month words - month_mode_of() reads them
    "vs", "versus", "against", "compared", "comparison", "quarter", "quarterly", "half",
    "their", "them", "they", "those", "these", "gst", "including", "incl", "excluding", "excl", "spend", "spent",
    "performing", "performance", "performer", "performers", "profitable", "profitability", "earning", "earners",
    "earn", "earns", "loss", "losses", "lossmaking", "running", "active",
    "increase", "increased", "decrease", "decreased", "grew", "growth", "rise", "rose", "fell", "dropped",
    "here", "currently", "current", "company", "working", "work", "now", "today", "whats", "what's", "hows",
    # Hinglish fillers (meaning comes from the payroll words around them)
    "kitna", "kitni", "kitne", "hai", "hain", "ka", "ki", "ke", "ko", "mein", "mujhe", "batao", "bata", "dikhao",
    "dikha", "do", "de", "dena", "kya", "kaun", "kaunsa", "kaunsi", "sabse", "zyada", "jyada", "kam", "wala", "wali",
    "aur", "ye", "yeh", "woh", "tha", "thi", "mahina", "mahine", "mahine", "pichle", "is", "iss", "us", "total",
    "gives", "giving", "gets", "got", "makes", "made", "earns", "earned", "generates", "brings", "contributes",
    "contributed", "does", "did", "are", "were", "there", "our", "my", "we", "us", "its", "it", "be", "been",
    "employer", "contribution", "contributions", "combined", "together", "sum", "will", "would", "be", "much",
    "hua", "hui", "hue", "huwa", "me", "mai", "main", "raha", "rahi", "rahe", "hota", "hoti", "hote", "kar", "karo",
    "making", "make", "doing", "hike", "hiked", "raise", "impact", "effect", "happen", "happens",
}
COUNT_WORDS = {"many", "count"}      # "how many joined in sep-26" = the fixed joiners report already counts people
for _, words in RULES:
    SIMPLE_WORDS |= words


# ----------------------------- Database -----------------------------
# ---- GUARDRAIL: the chatbot can only READ. Nothing it runs can change SQL Server data. ----
FORBIDDEN_SQL = re.compile(
    r"\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|TRUNCATE|CREATE|EXEC|EXECUTE|GRANT|REVOKE|DENY|"
    r"BACKUP|RESTORE|SHUTDOWN|DBCC|BULK|OPENROWSET|OPENQUERY|OPENDATASOURCE|INTO|KILL|"
    r"SP_\w+|XP_\w+|WAITFOR|RECONFIGURE|USE|DECLARE|SET|GO|OPENXML|REPLICATE)\b", re.IGNORECASE)


def assert_read_only(sql):
    """Blocks anything that is not a single plain SELECT."""
    s = re.sub(r"'[^']*'", "''", sql)          # ignore text in quotes
    s = re.sub(r"\[[^\]]*\]", "[]", s)          # ignore [column names]
    bad = FORBIDDEN_SQL.search(s)
    if not re.match(r"\s*(SELECT|WITH)\b", s, re.IGNORECASE) or ";" in s or "--" in s or "/*" in s or bad:
        raise PermissionError("Blocked by read-only guard: only SELECT queries are allowed"
                              + (f" (found '{bad.group(1)}')" if bad else "") + ".")


def assert_ai_sql(sql):
    """Extra checks for SQL written by the AI: read-only AND only the payroll table."""
    assert_read_only(sql)
    s = re.sub(r"'[^']*'", "''", sql)
    low = s.lower()
    for bad in ("sys.", "information_schema", "master.", "msdb", "tempdb", "#", "@@", "@", ".."):
        if bad in low:
            raise PermissionError(f"Blocked: the AI query may only read the {TABLE} table.")
    allowed = {TABLE.lower(), f"[{TABLE.lower()}]", f"dbo.{TABLE.lower()}", f"[dbo].[{TABLE.lower()}]",
               f"dbo.[{TABLE.lower()}]"}
    ctes = {m.group(1).lower() for m in re.finditer(r"(?:\bWITH|,)\s*(\w+)\s+AS\s*\(", s, re.IGNORECASE)}
    for m in re.finditer(r"\b(?:FROM|JOIN)\s+([^\s(),]+)", s, re.IGNORECASE):
        target = m.group(1).lower()
        bare = re.sub(r"^\[?dbo\]?\.", "", target).strip("[]")
        if target not in allowed and bare not in ctes and bare != TABLE.lower():
            raise PermissionError(f"Blocked: the AI query may only read the {TABLE} table (found {m.group(1)}).")
    if re.search(r"\bFROM\s+[^\s(),]+(?:\s+(?:AS\s+)?\w+)?\s*,\s*[^\s(]", s, re.IGNORECASE):
        raise PermissionError("Blocked: comma joins are not allowed in AI queries.")


def run_ai_query(sql, limit=None):
    """Runs an AI-written query: checked, read-only, time-limited, max ROW_LIMIT rows."""
    assert_ai_sql(sql)
    view = str(T)
    if view != RAW_T:                       # the AI may name the real table: force it through this user's view too
        sql = _policy_table_sql(sql.replace(view, "\x00VIEW\x00")).replace("\x00VIEW\x00", view)
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(sql)
        cols = [d[0] or f"col{i + 1}" for i, d in enumerate(cur.description or [])]
        rows = cur.fetchmany(limit or ROW_LIMIT) if cols else []
        df = pd.DataFrame.from_records([tuple(r) for r in rows], columns=cols)
    finally:
        try:
            conn.rollback()
        finally:
            conn.close()
    for c in df.columns:
        if df[c].dtype == object and df[c].map(lambda v: isinstance(v, Decimal)).any():
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# source column -> table column names that are allowed to differ (same data, other name)
REFRESH_RENAMES = {("monthsalary", "ctc"), ("pfeeer", "totalpf"), ("incometax", "incometax2"), ("empoption", "tdsopt")}


def refresh_column_problems(src, dst):
    """The refresh copies by POSITION (INSERT ... SELECT *, '', ''), so column N of the view must be column N of the
    table. Returns a list of problems (empty = safe)."""
    norm = lambda c: re.sub(r"[^a-z0-9]", "", str(c).lower())
    if len(src) + 2 != len(dst):
        return [f"view has {len(src)} columns, {TABLE} needs {len(dst) - 2} (+ FinaluniqCode, month_year)"]
    return [f"column {i + 1}: view '{s}' would go into '{d}'" for i, (s, d) in enumerate(zip(src, dst))
            if norm(s) != norm(d) and (norm(s), norm(d)) not in REFRESH_RENAMES]


def refresh_payroll_data():
    """Admin-only refresh: copy fresh rows from the source view, then run the stored procedure.
    1) copies the source into a temporary table (reading only - payrollall untouched)
    2) stops if the copy is empty, so a dead link can never wipe the table
    3) in ONE transaction: TRUNCATE payrollall -> INSERT the copy -> EXEC the procedure -> COMMIT
       If anything fails, everything is rolled back and payrollall keeps its old rows.
    Uses your Windows login (never the read-only chatbot login). Returns (old_rows, new_rows, seconds)."""
    if DUMMY_DATA:
        raise RuntimeError(DUMMY_DATA_MSG)
    if not REFRESH_SOURCE:
        raise RuntimeError("No REFRESH_SOURCE is set, so there is nothing to refresh from.")
    t0 = time.time()
    cs = (f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;")
    conn = pyodbc.connect(cs, autocommit=True)
    try:
        conn.timeout = REFRESH_TIMEOUT_SEC
        cur = conn.cursor()
        old = cur.execute(f"SELECT COUNT(*) FROM {RAW_T}").fetchone()[0]
        cur.execute("SET NOCOUNT ON; IF OBJECT_ID('tempdb..#refresh_stage') IS NOT NULL DROP TABLE #refresh_stage;")
        cur.execute(f"SET NOCOUNT ON; SELECT * INTO #refresh_stage FROM {REFRESH_SOURCE};")
        staged = cur.execute("SELECT COUNT(*) FROM #refresh_stage").fetchone()[0]
        if not staged:
            raise RuntimeError("The source view returned 0 rows - nothing was changed. "
                               "Check REFRESH_SOURCE.")
        cur.execute("SELECT TOP 0 * FROM #refresh_stage")
        src = [d[0] for d in cur.description]
        cur.fetchall()
        dst = [r[0] for r in cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ? "
                                         "ORDER BY ORDINAL_POSITION", [TABLE]).fetchall()]
        bad = refresh_column_problems(src, dst)
        if bad:
            raise RuntimeError(f"The source view's columns do not line up with {TABLE} - nothing was changed. "
                               + "; ".join(bad[:5]))
        if REFRESH_ALL_OR_NOTHING:
            conn.autocommit = False
        try:
            cur.execute("SET XACT_ABORT ON;")
            cur.execute(f"TRUNCATE TABLE {RAW_T};")
            cur.execute(f"INSERT INTO {RAW_T} SELECT *, '', '' FROM #refresh_stage;")
            cur.execute(f"SET NOCOUNT ON; EXEC {REFRESH_PROC};")
            while cur.nextset():                     # let the procedure finish all its steps
                pass
            new = cur.execute(f"SELECT COUNT(*) FROM {RAW_T}").fetchone()[0]
            if REFRESH_ALL_OR_NOTHING:
                conn.commit()
        except Exception:
            if REFRESH_ALL_OR_NOTHING:
                conn.rollback()                      # payrollall goes back to how it was
            raise
        return old, new, time.time() - t0
    finally:
        conn.close()


def get_conn():
    login = (f"UID={SQL_USER};PWD={SQL_PASSWORD};" if SQL_USER else "Trusted_Connection=yes;")
    cs = (f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};{login}"
          "ApplicationIntent=ReadOnly;")
    try:
        conn = pyodbc.connect(cs, autocommit=False, readonly=True)   # read-only connection
    except TypeError:                                                 # very old pyodbc
        conn = pyodbc.connect(cs, autocommit=False)
    try:
        conn.timeout = QUERY_TIMEOUT_SEC                              # stop long-running queries
    except Exception:
        pass
    return conn


@st.cache_data(ttl=300, show_spinner=False, max_entries=200)
def _cached_query(sql, params):
    return _run_query_raw(sql, list(params))


def run_query(sql, params):
    """Same question again within 5 minutes is answered from memory (instant)."""
    return _cached_query(sql, tuple(params or [])).copy()


def _run_query_raw(sql, params):
    assert_read_only(sql)
    conn = get_conn()
    try:
        df = pd.read_sql(sql, conn, params=params)
    finally:
        try:
            conn.rollback()        # even if something slipped through, never commit
        finally:
            conn.close()
    for c in df.columns:  # SQL decimal -> float
        if df[c].dtype == object and df[c].map(lambda v: isinstance(v, Decimal)).any():
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def _distinct_values_raw(column):
    df = run_query(f"SELECT DISTINCT TOP 200 [{column}] AS v FROM {RAW_T} "
                   f"WHERE [{column}] IS NOT NULL", [])
    return sorted({str(v).strip() for v in df["v"] if str(v).strip()}, key=len, reverse=True)


def distinct_values(column):
    """Names used to understand questions - empty when the column is restricted for this user."""
    eff = current_eff()
    if eff is not None and (not eff.payroll_access or not eff.key_visible(pe.Vocab().by_norm.get(pe.norm(column), column))):
        return []
    limit = (getattr(eff, "row_filter", None) or {}).get(column) if eff is not None else None
    if limit:                                          # a limited user only learns the names they may see
        return [v for v in _distinct_values_raw(column) if v in set(limit)]
    return _distinct_values_raw(column)


# ------------------ Period mode: Month/Year  or  month_year ------------------
# Three month columns - never mix them up:
#   "salary"     = [Month] + [Year]          the month the salary is FOR        (Sep-26 salary -> 9, 2026)
#   "month_year" = [month_year]              LIABLE month = payment date - 1 month  (paid 07-10-26 -> 07-09-26)
#   "payment"    = [month_year] + 1 month    PAYMENT date - calculated here only, there is no such column in SQL
PAY_EXPR = "DATEADD(MONTH, 1, [month_year])"


def period_exprs(mode):
    if mode == "month_year":
        return "YEAR([month_year])", "MONTH([month_year])"
    if mode == "payment":
        return f"YEAR({PAY_EXPR})", f"MONTH({PAY_EXPR})"
    return "[Year]", "[Month]"


def period_filter(mode, periods):
    """Index-friendly filter for a list of (month, year)."""
    parts, params = [], []
    for m, y in periods:
        if mode in ("month_year", "payment"):
            if mode == "payment":                      # paid in Oct-26 = liable month Sep-26
                m, y = (12, y - 1) if m == 1 else (m - 1, y)
            start = datetime.date(y, m, 1)
            end = datetime.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
            parts.append("([month_year] >= ? AND [month_year] < ?)")
            params += [start, end]
        else:
            parts.append("([Year] = ? AND [Month] = ?)")
            params += [y, m]
    return "(" + " OR ".join(parts) + ")", params


def latest_period(mode):
    return _latest_period(mode, str(T))                 # per view: a limited user gets the latest month THEY can see


@st.cache_data(ttl=600, show_spinner=False)
def _latest_period(mode, view):
    Y, M = period_exprs(mode)
    df = run_query(f"SELECT MAX({Y}*100+{M}) AS p FROM {view}", [])
    if pd.isna(df["p"].iloc[0]):
        raise ValueError("No data available for your login.")
    p = int(df["p"].iloc[0])
    return (p % 100, p // 100)


def data_range(whole=False):
    """(first, last) liable month (month_year) as 'Apr-2026' labels: what THIS user can see, or with whole=True
    the whole table (for admins choosing a 'show data from' month). (None, None) when there is no data."""
    return _data_range(RAW_T if whole else str(T))


@st.cache_data(ttl=600, show_spinner=False)
def _data_range(view):
    df = run_query(f"SELECT MIN([month_year]) AS a, MAX([month_year]) AS b FROM {view}", [])
    a, b = df["a"].iloc[0], df["b"].iloc[0]
    if pd.isna(a) or pd.isna(b):
        return None, None
    a, b = pd.Timestamp(a), pd.Timestamp(b)
    return period_label(a.month, a.year), period_label(b.month, b.year)


# needs "from now on" / "only" / "hide": a plain "show data from jul-26 to sep-26" stays a normal question
DATA_FROM_SET = re.compile(r"\bfrom\s+now\s+on\b[^.?!]*\bdata\s+(?:only\s+)?(?:from|since|starting)\b"
                           r"|\bonly\s+show\s+(?:the\s+)?data\s+(?:from|since|starting)\b"
                           r"|\bshow\s+(?:the\s+)?data\s+only\s+(?:from|since|starting)\b"
                           r"|\b(?:hide|remove)\s+(?:all\s+)?(?:the\s+)?data\s+(?:before|older\s+than)\b", re.I)
DATA_FROM_CLEAR = re.compile(r"\bshow\s+all\s+(?:the\s+)?(?:data|months)\b"
                             r"|\b(?:remove|clear|reset|undo)\s+(?:the\s+)?(?:show\s+)?data\s+(?:from|start)\b", re.I)
DATA_RANGE_ASK = re.compile(r"\bfrom\s+which\s+month\b|\bdata\s+(?:range|availab\w*|period)\b"
                            r"|\b(?:months?\s+available|available\s+months?)\b"
                            r"|\bwhich\s+months?\s+(?:of\s+)?(?:data\s+)?(?:do|did|is|are)\s+(?:we|i|there)\b", re.I)


def data_range_text(admin):
    a, b = data_range()
    df_ = data_from()
    txt = (f"You can see data from **{a}** to **{b}** (liable month, month_year)." if a else
           "There is no data you can see.")
    if df_.get("from"):
        d = datetime.date.fromisoformat(df_["from"])
        txt += (f"\n\nOlder months are hidden for everyone: data is shown from **{period_label(d.month, d.year)}** "
                f"on (set by {df_.get('by', '?')} on {df_.get('at', '?')}).")
    if admin:
        wa, wb = data_range(whole=True)
        if wa:
            txt += f"\n\nThe table itself has **{wa}** to **{wb}**."
        txt += ("\n\nTo change what everyone sees: *from now on only show data from apr-26*, or *show all data*. "
                "Settings → Data has the same switch.")
    return txt


def data_from_command(question):
    """Master / admin: 'from now on only show data from apr-26' -> everyone (master too) sees month_year >= Apr-26;
    'show all data' -> no limit. Anyone: 'from which month to which month is data available'."""
    admin = getattr(_POL, "raw_role", None) in ("master", "admin")
    periods = parse_periods(question)
    setm = DATA_FROM_SET.search(question) and len(periods) <= 1
    clear = not setm and not periods and DATA_FROM_CLEAR.search(question)   # "show all data for sep-26" = a question
    if not (clear or setm):
        eff = current_eff()
        if DATA_RANGE_ASK.search(question) and not periods and (eff is None or eff.payroll_access):
            return _policy_reply(data_range_text(admin), "data_range")
        return None
    if not admin:
        return _policy_reply("🔒 Only the master or an admin can change which months everyone sees.", "blocked")
    uid = getattr(_POL, "uid", "?")
    if clear:
        set_data_from(None, uid)
        audit_policy_event("data_from_cleared", "show all months again")
        return _policy_reply("✅ Done. Everyone sees **all months** again.\n\n" + data_range_text(True), "data_range")
    if not periods:
        return _policy_reply("From which month? For example: *from now on only show data from apr-26*.", "clarify")
    m, y = periods[0]
    wa, wb = data_range(whole=True)
    if wb and datetime.date(y, m, 1) > pd.Timestamp(datetime.datetime.strptime(wb, "%b-%Y")).date():
        return _policy_reply(f"The table only has data up to **{wb}**, so nobody would see anything. Nothing changed.",
                             "clarify")
    set_data_from(datetime.date(y, m, 1), uid)
    audit_policy_event("data_from_set", f"show data from {period_label(m, y)}")
    return _policy_reply(f"✅ Done. From now on everyone (users, admins and the master) sees data only from "
                         f"**{period_label(m, y)}** on.\n\n" + data_range_text(True), "data_range")


def mode_label(mode):
    return {"month_year": "liable month (month_year)",
            "payment": "payment month (month_year + 1 month)"}.get(mode, "salary month (Month/Year)")


PERIOD_COL = {"month_year": "Liable month", "payment": "Payment month"}   # name of the Period column per mode


def month_mode_of(question, default_mode):
    """Which month column the question means. 'paid in / payment' -> payment, 'liable / month_year' -> month_year,
    'salary month / month/year' -> salary; otherwise the user's default (Settings)."""
    ql = question.lower()
    if re.search(r"\bpaid\s+(?:in|during)\s+(?:the\s+)?(?:liable|month[_\- ]year)", ql):   # "paid in liable month sep-26"
        return "month_year"
    if re.search(r"\bpayment\s*(date|month)?\b|\bpaid\s+(in|on|during|from|between)\b|\bdisburs\w*", ql):
        return "payment"
    if re.search(r"month[_\- ]year|\bliable\b|\bliability\s+month\b", ql):
        return "month_year"
    if re.search(r"\bsalary\s+month\b|\bmonth\s*/\s*year\b", ql):
        return "salary"
    return default_mode


def add_payment_date(df):
    """Shows [month_year] as 'Liable month' and adds 'Payment date' (= month_year + 1 month) right after it.
    Display only - nothing is added in SQL Server."""
    if df is None or "month_year" not in df.columns:
        return df
    df = df.copy()
    if "PaymentDate" not in df.columns:
        my = df["month_year"]

        def nxt(v):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                return v
            if isinstance(v, str):
                for fmt in ("%b-%Y", "%d-%b-%Y", "%Y-%m-%d"):
                    try:
                        d = datetime.datetime.strptime(v.strip(), fmt)
                        return (pd.Timestamp(d) + pd.DateOffset(months=1)).strftime(fmt)
                    except ValueError:
                        pass
                return None
            return (pd.Timestamp(v) + pd.DateOffset(months=1)).date()
        df.insert(list(df.columns).index("month_year") + 1, "PaymentDate", [nxt(v) for v in my])
    elif list(df.columns).index("PaymentDate") != list(df.columns).index("month_year") + 1:
        c_ = df.pop("PaymentDate")
        df.insert(list(df.columns).index("month_year") + 1, "PaymentDate", c_)
    return df.rename(columns={"month_year": "Liable month (month_year)", "PaymentDate": "Payment date"})


# --------------------------- Month parsing ---------------------------
def fix_year(y):
    y = int(y)
    return y + 2000 if y < 100 else y


def parse_periods(text):
    """Finds months like aug-26, Aug 2026, august'26, 08/26, 2026-08."""
    t = text.lower()
    found = []
    # real month names only - "margin" must NOT be read as "mar", "decision" not as "dec"
    names = (r"(jan|feb|mar|apr|may|jun|jul|aug|sept|sep|oct|nov|dec)"
             r"(?:uary|ruary|ch|il|e|y|ust|ember|tember|ober)?(?![a-z])")

    for m in re.finditer(rf"\b{names}[\s\-/',.]*(\d{{4}}|\d{{2}})\b", t):
        found.append((MONTH_MAP[m.group(1)[:3]], fix_year(m.group(2))))
    for m in re.finditer(r"\b(\d{4})[/\-](\d{1,2})\b", t):
        if 1 <= int(m.group(2)) <= 12:
            found.append((int(m.group(2)), int(m.group(1))))
    for m in re.finditer(r"\b(\d{1,2})[/\-](\d{4}|\d{2})\b", t):
        if 1 <= int(m.group(1)) <= 12:
            found.append((int(m.group(1)), fix_year(m.group(2))))

    if not found:  # month name without year -> most recent such month
        today = datetime.date.today()
        for m in re.finditer(rf"\b{names}\b", t):
            mon = MONTH_MAP[m.group(1)[:3]]
            found.append((mon, today.year if mon <= today.month else today.year - 1))

    found = sorted(set(found), key=lambda p: (p[1], p[0]))

    if len(found) == 2 and re.search(r"\b(to|till|until|between|from)\b", t):
        (m1, y1), (m2, y2) = found
        found, m, y = [], m1, y1
        while (y, m) <= (y2, m2):
            found.append((m, y))
            m, y = (1, y + 1) if m == 12 else (m + 1, y)
    return found


def period_label(m, y):
    return f"{MON[int(m) - 1]}-{int(y)}"          # e.g. Aug-2026


# ------------------------ Understand the question ------------------------
SHORT_WORDS = {"the", "and", "for", "you", "all", "new", "old", "per", "top", "who", "how", "why", "not", "any", "few",
               "pay", "net", "tds", "ctc", "gst", "hra", "esi", "nps", "has", "had", "was", "are", "can", "get", "set"}
NAME_FILLER = {"and", "the", "of", "pvt", "ltd", "limited", "private", "dept", "department",
               "team", "division", "services", "service", "group", "unit", "&"}


def find_known(question, column):
    """Finds a department/category written in the question (no AI).
    Full name ("Openvem Services") or its key words ("open vem") both work."""
    values = distinct_values(column)
    for val in values:                                   # 1) full name written
        if len(val) < 2:
            continue
        flags = 0 if len(val) <= 3 else re.IGNORECASE   # short names like IT: exact case
        if column == "Category" and len(val) == 3 and val.lower() not in SHORT_WORDS:
            flags = re.IGNORECASE                        # "gil margin" = GIL (not a normal word)
        if re.search(rf"(?<!\w){re.escape(val)}(?!\w)", question, flags):
            return val
    q_words = re.findall(r"[a-z0-9]+", question.lower())
    q_text = " " + " ".join(q_words) + " "
    best, best_n = None, 0
    for val in values:                                   # 2) 2+ key words of the name, in order
        words = [w for w in re.findall(r"[a-z0-9]+", val.lower()) if w not in NAME_FILLER and len(w) > 1]
        if len(words) < 2:
            continue
        for n in range(len(words), 1, -1):
            hit = any(f" {' '.join(words[i:i + n])} " in q_text for i in range(len(words) - n + 1))
            if hit:
                if n > best_n:
                    best, best_n = val, n
                break
    if best:
        return best
    joined = {"".join(q_words[i:i + n]) for n in range(2, 5) for i in range(len(q_words) - n + 1)}
    for val in values:                                   # 3a) typed apart: "open vem" -> OPENVEM
        v = re.sub(r"[^a-z0-9]", "", val.lower())
        if len(v) >= 5 and v in joined:
            return val
    compact = set(q_words)                               # 3) typed together: "openvem"
    for val in values:
        words = [w for w in re.findall(r"[a-z0-9]+", val.lower()) if w not in NAME_FILLER and len(w) > 1]
        for i in range(len(words) - 1):
            joined = words[i] + words[i + 1]
            if len(joined) >= 6 and joined in compact:
                return val
    return None


def name_words(val):
    return set(re.findall(r"[a-z0-9_]+", str(val or "").lower()))


# work order code: has a digit (M244425094) OR a wildcard % / * (wox%, *2508*)
_WO_CODE = r"([A-Za-z0-9\-/%*]*[%*][A-Za-z0-9\-/%*]*|[A-Za-z0-9][A-Za-z0-9\-/]*\d[A-Za-z0-9\-/%*]*)"
WO_PATTERN = re.compile(
    r"\b(?:work\s*order|workorder|w\.?o\.?)\s*(?:no\.?|number|#)?\s*[:#\-]?\s*" + _WO_CODE,
    re.IGNORECASE)
WO_CODE_ALONE = re.compile(r"\b([A-Za-z]{1,4}\d{5,}[A-Za-z0-9\-/%*]*)")   # e.g. M244425094
MONTH_CODE = re.compile(r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\-/']?\d{2,4}|"
                        r"\d{1,2}[\-/]\d{2,4}|\d{4}[\-/]\d{1,2}", re.IGNORECASE)   # aug-26, 08/26, 2026-08
WO_WORDS = re.compile(r"\b(?:work\s*order|workorder|w\.?o\.?)s?\b", re.IGNORECASE)
WO_PHRASE = re.compile(      # "work order starting with wox", "wo containing 2508"
    r"\b(start(?:s|ing)?\s+with|begin(?:s|ning)?\s+with|like|contain(?:s|ing)?|end(?:s|ing)?\s+with)"
    r"\s+['\"]?([A-Za-z0-9\-/]+)", re.IGNORECASE)


def find_workorder(question):
    """Returns (work order no or pattern, question with it removed) or (None, question).
    Patterns: wox% (starts with), %wox% (contains), * also works like %."""
    if WO_WORDS.search(question):
        m = WO_PHRASE.search(question)
        if m:
            word, code = m.group(1).lower(), m.group(2)
            pat = (f"{code}%" if word.startswith(("start", "begin"))
                   else f"%{code}" if word.startswith("end") else f"%{code}%")
            return pat.upper(), question[:m.start()] + " " + question[m.end():]
    m = WO_PATTERN.search(question)
    if m and MONTH_CODE.fullmatch(m.group(1)):          # "lowest margin work order aug-26": a month, not a WO no.
        m = None
    m = m or WO_CODE_ALONE.search(question)
    if not m:
        return None, question
    return m.group(1).upper().replace("*", "%"), question[:m.start(1)] + " " + question[m.end(1):]


def is_pattern(wo):
    return "%" in str(wo or "")


def valid_wo(wo):
    """A real work order no. has a digit (M244425094) or is a pattern with 2+ letters (WOX%)."""
    wo = str(wo or "")
    return bool(re.search(r"\d", wo)) or (is_pattern(wo) and len(re.sub(r"[^A-Za-z0-9]", "", wo)) >= 2)


def fast_intent(question):
    """Returns intent without AI, or None if the question needs the AI."""
    wo, question = find_workorder(question)
    t = question.lower()
    intent = {"report": "summary", "department": None, "category": None,
              "employee": None, "workorder": wo}
    if wo:
        intent["report"] = "workorder_detail"

    code = re.search(r"\b(?:employee|emp|code|id)\s*(?:no\.?|code|id)?\s*[:#\-]?\s*(\d{2,})\b", t)
    if code:
        intent.update(report="employee_detail", employee=code.group(1))

    intent["department"] = find_known(question, "Department")
    intent["category"] = find_known(question, "Category")

    tokens = set(re.findall(r"[a-z_]+", t.replace("month_year", " ")))
    if not code and not wo:
        for report, words in RULES:
            if tokens & words:
                intent["report"] = report
                break

    known = SIMPLE_WORDS | MONTH_WORDS
    for v in (intent["department"], intent["category"]):
        if v:
            known |= name_words(v)
            w = [x for x in re.findall(r"[a-z0-9]+", v.lower()) if x not in NAME_FILLER]
            known |= {w[i] + w[i + 1] for i in range(len(w) - 1)}
            vc = re.sub(r"[^a-z0-9]", "", v.lower())          # "open vem" typed apart = OPENVEM
            known |= {t for t in tokens if len(t) >= 3 and t in vc}
    unknown = tokens - known
    if unknown and not code and not wo:                 # "salary of Rohan Singh" -> no AI needed
        emp = find_employee_name(question)
        if emp:
            intent["employee"] = emp
            unknown -= name_words(emp)
    return intent if not unknown else None


def find_employee_name(question):
    """A real employee name (2+ words, in order) written in the question, else None. No AI."""
    eff = current_eff()
    if eff is not None and (not eff.payroll_access or not eff.key_visible(pe.Vocab().by_norm.get(pe.norm("Name"), "Name"))):
        return None
    q_text = " " + " ".join(re.findall(r"[a-z0-9]+", question.lower())) + " "
    best = None
    for v in run_query(f"SELECT DISTINCT [Name] AS v FROM {T} WHERE [Name] IS NOT NULL", [])["v"]:
        words = re.findall(r"[a-z0-9]+", str(v).lower())
        if len(words) >= 2 and f" {' '.join(words)} " in q_text and (best is None or len(str(v)) > len(best)):
            best = str(v).strip()
    return best


def _known_list(column, n=40):
    try:
        return ", ".join(distinct_values(column)[:n]) or "(none)"
    except Exception:
        return "(unknown)"


def _clean(v):
    """AI sometimes returns 'null', '', 'none' as text -> treat as empty."""
    if v is None:
        return None
    v = str(v).strip()
    return None if v.lower() in ("", "null", "none", "n/a", "na", "all") else v


def ai_intent(question):
    report_list = "\n".join(f'- "{k}": {v}' for k, v in REPORTS.items())
    system = f"""You classify payroll questions. Reply ONLY with JSON:
{{"report": "<key>", "department": null or text, "category": null or text, "employee": null or name or code, "workorder": null or work order number, "confident": true or false, "clarify": null or text}}
Report keys:
{report_list}
Known departments: {_known_list("Department")}
Known categories: {_known_list("Category")}
Rules: ignore months/years (handled separately). Keep % wildcards in work order numbers as typed (e.g. "wox%").
Set "workorder" ONLY if a specific work order number or pattern is written; otherwise null. Use a department/category only if the
question mentions it, spelled like the known list. "employee" = a person's name or numeric code.
If the question is unclear, could mean several reports, or matches none of them, set "confident": false and put
ONE short question to ask the user in "clarify". Never guess; asking back is better than a wrong report.
Examples:
"give me salary for aug-26" -> {{"report": "summary", "department": null, "category": null, "employee": null, "workorder": null}}
"salary of Amit Yadav" -> {{"report": "employee_detail", "department": null, "category": null, "employee": "Amit Yadav", "workorder": null}}
"how much tds was cut in july" -> {{"report": "deductions", "department": null, "category": null, "employee": null, "workorder": null}}
"which workorder has highest margin in aug-26" -> {{"report": "workorder_summary", "department": null, "category": null, "employee": null, "workorder": null}}
"show me the numbers" -> {{"report": "summary", "department": null, "category": null, "employee": null, "workorder": null, "confident": false, "clarify": "Which numbers would you like: total salary, department-wise, deductions or work order margin, and for which month?"}}
"which workorder end date is greater than jan-28" -> {{"report": "analysis", "department": null, "category": null, "employee": null, "workorder": null}}
"how many employees joined after 2024" -> {{"report": "analysis", "department": null, "category": null, "employee": null, "workorder": null}}
"great, lets go" -> {{"report": "chitchat", "department": null, "category": null, "employee": null, "workorder": null}}
"who left the company in aug-26" -> {{"report": "leavers", "department": null, "category": null, "employee": null, "workorder": null}}"""
    if taught_prompt():
        system += "\n" + taught_prompt()
    resp = ollama.chat(
        model=MODEL,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": question}],
        format="json",
        keep_alive="60m",                       # keep model in RAM -> no reload delay
        options={"temperature": 0, "num_predict": 120, "num_ctx": 2048, "num_thread": AI_THREADS},
    )
    data = json.loads(resp["message"]["content"])
    conf = data.get("confident", True)
    clar = _clean(data.get("clarify"))
    data = {k: _clean(data.get(k)) for k in ("report", "department", "category", "employee", "workorder")}
    data["confident"] = not (conf is False or str(conf).strip().lower() in ("false", "no", "0"))
    data["clarify"] = clar
    if data["report"] not in REPORTS:
        data["report"] = "summary"
    if not valid_wo(data["workorder"]):
        data["workorder"] = None
    # regex is more reliable for codes -> trust it over the AI when it finds one
    wo, _ = find_workorder(question)
    if wo:
        data["workorder"] = wo
    return data


# ------------------------------ SQL builder ------------------------------
NEG_WORDS = (r"without|except|excluding|exclude|excluded|excl\.?|other\s+th[ae]n|apart\s+from|besides|but\s+not|"
             r"not\s+including|minus|leave\s+out|leaving\s+out|leaving|ignore|ignoring|remove|removing|skip|"
             r"non[\s\-]?|no|bina|chhod\s*ke|chhodke|chod\s*ke|ko\s+chhod\s*kar|hata\s*ke|hatake")
NEG_AFTER = re.compile(rf"\b(?:{NEG_WORDS})\s+(?:the\s+|any\s+)?"
                       r"(.+?)(?=\s+(?:for|in|of|from|month|months|monthwise|wise|with|during|between|and\s+give|"
                       r"show|give|by|this|last|current|till|to)\b|"
                       r"\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-/'.]*\d|\s+\d{1,4}[/\-]\d|"
                       r"[,?!;]|\.(?=\s|$)|$)", re.IGNORECASE)
NEG_BEFORE = re.compile(r"((?:\S+\s+){0,4}?\S+)\s+(?:ke\s+bina|ko\s+chhod\s*kar|chhod\s*ke|chod\s*ke|hata\s*ke|ko\s+hata\s*ke)\b",
                        re.IGNORECASE)


def _find_all(text, column):
    out, t = [], text
    for _ in range(6):
        v = find_known(t, column)
        if not v:
            break
        out.append(v)
        t = re.sub(re.escape(v), " ", t, flags=re.IGNORECASE)
        for w in name_words(v):
            t = re.sub(rf"\b{re.escape(w)}\b", " ", t, flags=re.IGNORECASE)
    return out


def split_exclusions(question):
    """'margin month wise without efirm' -> ('margin month wise', {'Category': ['EFIRM']})."""
    excl, q = {"Category": [], "Department": [], "WorkOrderNo": []}, re.sub(r"(?i)\bnon[\s\-]+", "without ", question)
    for rx in (NEG_AFTER, NEG_BEFORE):
        for m in list(rx.finditer(q)):
            seg = m.group(1)
            if re.match(r"(?i)no\b", m.group(0)) and re.search(
                    r"(?i)(work\s*order|workorder|w\.?o\.?|wo|emp\w*|employee|code|sr|serial|invoice|bill)\s*$",
                    q[:m.start()]):
                continue                              # "workorder no wox%" = work order NUMBER, not "without"
            found = False
            for col in ("Category", "Department"):
                vals = _find_all(seg, col)
                if vals:
                    excl[col] += [v for v in vals if v not in excl[col]]
                    found = True
            for tok in re.findall(r"[A-Za-z0-9][A-Za-z0-9/\-_.()]*%|%[A-Za-z0-9][A-Za-z0-9/\-_.()]*%?", seg):
                excl["WorkOrderNo"].append(tok.upper())       # "wox%" -> work orders starting WOX
                found = True
            wo, _ = find_workorder(seg)
            if wo and not is_pattern(wo) and wo.upper() not in excl["WorkOrderNo"]:
                excl["WorkOrderNo"].append(wo.upper())        # one exact work order number
                found = True
            if not found and any(re.search(ask, seg, re.IGNORECASE) for ask, _ in HIDE_GROUPS.values()):
                found = True                  # "salary without pf" -> just don't ask for pf
            if found:
                q = q.replace(m.group(0), " ")
    return re.sub(r"\s+", " ", q).strip(), excl


def split_includes(question):
    """'... of openvem, gml ...' -> both categories. Only used when 2 or more are named
    (one name is handled by the normal filter). Returns (question without the names, {column: [values]})."""
    inc, q = {}, question
    for col in ("Category", "Department"):
        vals = _find_all(q, col)
        if len(vals) >= 2:
            inc[col] = vals
            for v in vals:
                q = re.sub(re.escape(v), " ", q, flags=re.IGNORECASE)
                vc = re.sub(r"[^a-z0-9]", "", v.lower())
                words = re.findall(r"[A-Za-z0-9]+", q)
                for n in range(4, 0, -1):                 # "open vem" typed apart
                    for i in range(len(words) - n + 1):
                        if "".join(words[i:i + n]).lower() == vc:
                            q = re.sub(r"\b" + r"\W+".join(map(re.escape, words[i:i + n])) + r"\b", " ", q,
                                       flags=re.IGNORECASE)
                for w in name_words(v):
                    q = re.sub(rf"\b{re.escape(w)}\b", " ", q, flags=re.IGNORECASE)
    q = re.sub(r"\s*,\s*(,\s*)*", " ", q)
    q = re.sub(r"\b(of|for|and|&)\s+(and|&|of)\b", r"\1", q, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", q).strip(), inc


def excluded_text():
    ex = current_exclusions()
    return ", ".join(ex.get("Category", []) + ex.get("Department", []) + ex.get("WorkOrderNo", []))


def current_exclusions():
    try:
        return st.session_state.get("_excl") or {}
    except Exception:
        return {}


def _excl_clauses():
    """SQL for the active exclusions - only real values from the data, so nothing can be injected."""
    out = []
    for key, vals in current_exclusions().items():
        if not vals:
            continue
        if key == "wo_group":                     # "NORTH margin" -> saved group names; SQL built from the saved rules
            c = wg.filter_sql(WO_GROUPS.load(), vals)
            if c:
                out.append(c)
            continue
        if key == "WorkOrderNo":                  # patterns like WOX% -> NOT LIKE (safe characters only)
            for v in vals:
                if re.fullmatch(r"[A-Za-z0-9/\-_.()% ]{2,80}", v):
                    out.append(f"ISNULL([WorkOrderNo], N'') NOT LIKE N'{v}'")
            continue
        col = key.replace("include_", "")
        if col not in ("Category", "Department"):
            continue
        try:
            known = set(distinct_values(col))
        except Exception:
            known = set()
        vals = [v for v in vals if v in known]
        if vals:
            lits = ", ".join("N'" + v.replace("'", "''") + "'" for v in vals)
            out.append(f"[{col}] IN ({lits})" if key.startswith("include_") else
                       f"ISNULL([{col}], N'') NOT IN ({lits})")
    return out


def build_where(mode, periods, f, date_col=None):
    clauses, params = list(_excl_clauses()), []
    if periods:
        if date_col:   # joiners / leavers
            exprs = f"(YEAR([{date_col}])*100+MONTH([{date_col}]))"
            clauses.append(f"{exprs} IN ({','.join('?' * len(periods))})")
            params += [y * 100 + m for m, y in periods]
        else:
            c, p = period_filter(mode, periods)
            clauses.append(c)
            params += p
    for col in ("Department", "Category"):
        val = (f.get(col.lower()) or "").strip()
        if val:
            if val in distinct_values(col):
                clauses.append(f"[{col}] = ?")
                params.append(val)
            else:
                clauses.append(f"[{col}] LIKE ?")
                params.append(f"%{val}%")
    wo = str(f.get("workorder") or "").strip()
    if wo:
        if is_pattern(wo):                      # wox%  -> all work orders starting WOX
            clauses.append("[WorkOrderNo] LIKE ?")
        else:
            clauses.append("[WorkOrderNo] = ?")
        params.append(wo)
    emp = str(f.get("employee") or "").strip()
    if emp:
        if emp.isdigit():
            clauses.append("[Employee] = ?")
            params.append(int(emp))
        else:
            clauses.append("[Name] LIKE ?")
            params.append(f"%{emp}%")
    return ("WHERE " + " AND ".join(clauses)) if clauses else "", params


TOTALS = """COUNT(DISTINCT [Employee]) AS Employees,
       SUM([CTC]) AS CTC, SUM([_Base]) AS BaseSalesValue, SUM([_Base]) - SUM([CTC]) AS Margin,
       SUM([_Inv]) AS InvoiceWithGST,
       CASE WHEN SUM([CTC]) <> 0 THEN (SUM([_Base]) - SUM([CTC])) * 100.0 / SUM([CTC]) END AS MarginPct,
       CASE WHEN SUM([_Base]) <> 0 THEN (SUM([_Base]) - SUM([CTC])) * 100.0 / SUM([_Base]) END AS MarginPctSales,
       SUM([GrossAmount]) AS GrossAmount,
       SUM([IncomeTax]) AS IncomeTax, SUM([Total PF]) AS TotalPF,
       SUM([PT]) AS PT, SUM([ESIC_Employee]) AS ESIC_Employee,
       SUM({TH}) AS TakeHome"""


TOTALS = TOTALS.replace("{TH}", TAKEHOME_EXPR)


def base_rows(Y, M, where):
    """Every row with its share of the work order's base sales value (invoice x 100/118).
    The invoice is repeated on each employee row, so it is split equally over that work order's
    rows in the month -> adding rows up counts each work order's base sales exactly once.
    ALWAYS per WorkOrderNo + salary [Year] + [Month]: one month_year can hold 2 salary months of the
    same work order (e.g. Jun and Jul both paid in Aug) - each has its own invoice and must count."""
    return (f"(SELECT *, MAX({BASE_EXPR}) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) * 1.0 "
            f"/ COUNT(*) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) AS [_Base], "
            f"MAX([InvoiceAmount]) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) * 1.0 "
            f"/ COUNT(*) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) AS [_Inv] "   # invoice WITH GST, once per WO
            f"FROM {T} {where}) AS r")


def _col_or_zero(col):
    """[col] if the table has it, else 0 (Payrollall1 has no ChildEduAllownce column)."""
    return f"[{col}]" if any(c.lower() == col.lower() for c, _ in _table_columns()) else "0"


def build_sql(report, mode, periods, f):
    Y, M = period_exprs(mode)
    P = f"{Y} AS [Year], {M} AS [Month]"
    G = f"GROUP BY {Y}, {M}"

    if report in ("joiners", "leavers"):
        col = "joindate" if report == "joiners" else "resigndate"
        where, params = build_where(mode, periods, f, date_col=col)
        # one row per person, with the CTC of their latest salary month ("... and what was their last ctc")
        return (f"""SELECT [Employee], [Name], [Department], [Category], [{col}],
       [Latest salary month], [Latest CTC] FROM (
  SELECT [Employee], [Name], [Department], [Category], [{col}],
         FORMAT(DATEFROMPARTS([Year], [Month], 1), 'MMM-yyyy') AS [Latest salary month],
         SUM([CTC]) OVER (PARTITION BY [Employee], [Year], [Month]) AS [Latest CTC],
         ROW_NUMBER() OVER (PARTITION BY [Employee] ORDER BY [Year] DESC, [Month] DESC) AS rn
  FROM {T} {where}) AS x
WHERE rn = 1 ORDER BY [{col}], [Name]"""), params

    where, params = build_where(mode, periods, f)

    SRC = base_rows(Y, M, where)
    if report == "summary":
        sql = f"SELECT {P}, {TOTALS} FROM {SRC} {G} ORDER BY [Year], [Month]"
    elif report == "trend":
        top = "" if periods else "TOP 24"
        sql = f"SELECT {top} {P}, {TOTALS} FROM {SRC} {G} ORDER BY [Year] DESC, [Month] DESC"
    elif report in ("department", "category"):
        col = report.capitalize()
        sql = (f"SELECT {P}, [{col}], {TOTALS} FROM {SRC} {G}, [{col}] "
               f"ORDER BY [Year], [Month], CTC DESC")
    elif report == "deductions":
        sql = f"""SELECT {P}, COUNT(DISTINCT [Employee]) AS Employees,
       SUM([IncomeTax]) AS IncomeTax, SUM([IncomeTax_2]) AS IncomeTax_2,
       SUM([PF_Employee]) AS PF_Employee, SUM([PF_Employer]) AS PF_Employer,
       SUM([Total PF]) AS TotalPF, SUM([PT]) AS PT,
       SUM([ESIC_Employee]) AS ESIC_Employee, SUM([ESIC]) AS ESIC,
       SUM([CORP_NPS]) AS CORP_NPS, SUM([Deduction]) AS Deduction
FROM {T} {where} {G} ORDER BY [Year], [Month]"""
    elif report == "components":
        sql = f"""SELECT {P}, COUNT(DISTINCT [Employee]) AS Employees,
       SUM([Basic]) AS Basic, SUM([HRA]) AS HRA,
       SUM([ConveyanceAllowance]) AS ConveyanceAllowance,
       SUM([SpecialAllowance]) AS SpecialAllowance,
       SUM([TransportAllowance]) AS TransportAllowance,
       SUM({_col_or_zero("ChildEduAllownce")}) AS ChildEduAllowance,
       SUM([MedicalAllownce]) AS MedicalAllowance,
       SUM([Bonus]) AS Bonus, SUM([Arrears]) AS Arrears,
       SUM([GrossAmount]) AS GrossAmount
FROM {T} {where} {G} ORDER BY [Year], [Month]"""
    elif report in ("employee_detail", "workorder_detail"):
        sql = (f"SELECT TOP ({ROW_LIMIT}) *, {PAY_EXPR} AS PaymentDate, {BASE_EXPR} AS BaseSalesValue, "
               f"{TAKEHOME_EXPR} AS TakeHome FROM {T} {where} ORDER BY {Y} DESC, {M} DESC")
    elif report == "workorder_summary":
        top = int(f.get("top") or ROW_LIMIT)
        order = "ASC" if f.get("sort") == "asc" else "DESC"
        # inner: one row per work order per month -> base sales counted once, CTC summed per employee
        sql = f"""SELECT TOP ({top}) [WorkOrderNo], MAX([Employees]) AS Employees, COUNT(*) AS Months,
       SUM([BaseSalesValue]) AS BaseSalesValue, SUM([InvoiceAmount]) AS InvoiceWithGST, SUM([CTC]) AS CTC,
       SUM([GrossAmount]) AS GrossAmount, SUM([TakeHome]) AS TakeHome,
       {MARGIN_FORMULA} AS Margin,
       CASE WHEN SUM([CTC]) <> 0 THEN ({MARGIN_FORMULA}) * 100.0 / SUM([CTC]) END AS MarginPct,
       CASE WHEN SUM([BaseSalesValue]) <> 0
            THEN ({MARGIN_FORMULA}) * 100.0 / SUM([BaseSalesValue]) END AS MarginPctSales
FROM (SELECT [WorkOrderNo], [Year] AS PY, [Month] AS PM, COUNT(DISTINCT [Employee]) AS Employees,
       MAX({BASE_EXPR}) AS BaseSalesValue, MAX([InvoiceAmount]) AS InvoiceAmount, SUM([CTC]) AS CTC,
       SUM([GrossAmount]) AS GrossAmount, SUM({TAKEHOME_EXPR}) AS TakeHome
      FROM {T} {where} GROUP BY [WorkOrderNo], [Year], [Month]) AS wo
GROUP BY [WorkOrderNo] ORDER BY Margin {order}"""
    elif report == "workorder_monthly":
        # one row per WorkOrderNo + salary Year + Month: invoice / base sales once, CTC = sum of its employees
        sql = f"""SELECT TOP ({ROW_LIMIT}) MAX([Category]) AS Category, MAX([Department]) AS Department,
       [WorkOrderNo], [Year], [Month], FORMAT(MAX([month_year]), 'MMM-yyyy') AS [month_year],
       FORMAT(MAX({PAY_EXPR}), 'MMM-yyyy') AS [PaymentDate],
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate, MAX([FinaluniqCode]) AS FinaluniqCode,
       COUNT(DISTINCT [Employee]) AS Employees, MAX([InvoiceAmount]) AS InvoiceAmount,
       ROUND(MAX({BASE_EXPR}), 2) AS BaseSalesValue, SUM([CTC]) AS CTC,
       ROUND(MAX({BASE_EXPR}), 2) - SUM([CTC]) AS Margin,
       CASE WHEN SUM([CTC]) <> 0 THEN (MAX({BASE_EXPR}) - SUM([CTC])) * 100.0 / SUM([CTC]) END AS MarginPct,
       CASE WHEN MAX({BASE_EXPR}) <> 0
            THEN (MAX({BASE_EXPR}) - SUM([CTC])) * 100.0 / MAX({BASE_EXPR}) END AS MarginPctSales
FROM {T} {where} GROUP BY [WorkOrderNo], [Year], [Month] ORDER BY [WorkOrderNo], [Year], [Month]"""
    else:  # employee_list
        sql = f"""SELECT TOP ({ROW_LIMIT}) {P}, [month_year], {PAY_EXPR} AS [PaymentDate], [Employee], [Name], [Department],
       [Category], [PaidDays], [GrossAmount], [IncomeTax], [Total PF], [PT],
       [ESIC_Employee], {TAKEHOME_EXPR} AS TakeHome
FROM {T} {where} ORDER BY [Year], [Month], [Department], [Name]"""
    return sql, params


# ------------------------------ Answer text ------------------------------
def inr(x):
    try:
        x = float(x or 0)
    except (TypeError, ValueError):
        return "–"
    if x != x or x in (float("inf"), float("-inf")):     # blank (NaN) value in the data
        return "–"
    sign = "-" if x < 0 else ""
    whole, dec = f"{abs(x):.2f}".split(".")
    if len(whole) > 3:
        rest = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", whole[:-3])
        whole = f"{rest},{whole[-3:]}"
    return f"{sign}₹{whole}.{dec}"


def _wo_rank_text(df, periods_text, intent):
    r = df.iloc[0]
    word = "Lowest" if (intent or {}).get("sort") == "asc" else "Highest"
    fl = (intent or {}).get("focus_label", "Margin")
    if (intent or {}).get("focus") not in (None, "Margin", "MarginPct", "CTC") or fl == "Margin %":
        word = f"{word} {fl.lower()}"
        fl_val = r.get((intent or {}).get("focus"))
        if fl_val is not None and pd.notna(fl_val):
            word += f" ({f'{fl_val:.2f}%' if fl == 'Margin %' else inr(fl_val)})"
        return (f"**{word} for {periods_text}: work order {r['WorkOrderNo']}** · Margin {inr(r['Margin'])} · "
                f"Base sales value {inr(r['BaseSalesValue'])} · CTC {inr(r['CTC'])}. "
                f"Ranked {len(df)} work order(s) below.")
    on = "base sales" if (intent or {}).get("pct_basis") == "sales" else "CTC"
    pct = f" ({r['MarginPct']:.2f}% on {on})" if pd.notna(r.get("MarginPct")) else ""
    return (f"**{word} margin amount for {periods_text}: work order {r['WorkOrderNo']}** · "
            f"Margin {inr(r['Margin'])}{pct} · Base sales value {inr(r['BaseSalesValue'])} · "
            f"CTC {inr(r['CTC'])} · {int(r['Employees'])} employee(s). "
            f"Ranked {len(df)} work order(s) below.")


def _wo_other_leader(df, intent):
    """'Highest margin' can mean the amount or the %. The answer leads with one; this names the other winner."""
    if len(df) < 2 or "MarginPct" not in df or (intent or {}).get("focus") not in (None, "Margin", "MarginPct"):
        return ""
    low = (intent or {}).get("sort") == "asc"
    word = "Lowest" if low else "Highest"
    by_pct = (intent or {}).get("focus") == "MarginPct" or (intent or {}).get("focus_label") == "Margin %"
    col = "Margin" if by_pct else "MarginPct"
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if s.empty:
        return ""
    r = df.loc[s.idxmin() if low else s.idxmax()]
    among = f" among these {len(df)}" if (intent or {}).get("top") else ""   # "top 5": not over all work orders
    if r["WorkOrderNo"] == df.iloc[0]["WorkOrderNo"]:
        return f"It is also the **{word.lower()} by {'margin amount' if by_pct else 'margin %'}{among}**."
    if by_pct:
        return (f"**{word} margin amount{among}:** work order {r['WorkOrderNo']} · Margin {inr(r['Margin'])}"
                + (f" ({r['MarginPct']:.2f}%)" if pd.notna(r.get("MarginPct")) else ""))
    return f"**{word} margin %{among}:** work order {r['WorkOrderNo']} · {r['MarginPct']:.2f}% · Margin {inr(r['Margin'])}"


DEDUCTION_ASK = [   # (words, column, label) - the deduction a question names is answered first
    (r"employer\s+(?:pf|contribution|share)|pf\s+employer", "PF_Employer", "PF (employer)"),
    (r"employee\s+(?:pf|contribution|share)|pf\s+employee", "PF_Employee", "PF (employee)"),
    (r"\bpf\b|provident", "TotalPF", "Total PF"),
    (r"\btds\b", "IncomeTax", "TDS"),
    (r"\bpt\b", "PT", "PT"),
    (r"\besic?\b", "ESIC_Employee", "ESIC (employee)"),
    (r"\besic?\b", "ESIC", "ESIC (employer)"),
    (r"\bnps\b", "CORP_NPS", "Corporate NPS"),
]
GST_ASK = re.compile(r"\bgst\s+(?:amount|value|collected|charged|part|portion|component|on|kitna|kitni)\b|"
                     r"\bamount\s+of\s+gst\b|\bhow\s+much\s+gst\b|\btotal\s+gst\b|(?<!with\s)\bgst\s+(?:for|in)\b", re.IGNORECASE)
COMPONENT_ASK = [   # salary components a question names -> answered first
    (r"\bbasic\b", "Basic", "Basic"), (r"\bhra\b", "HRA", "HRA"), (r"conveyance", "ConveyanceAllowance", "Conveyance"),
    (r"special\s+allowance", "SpecialAllowance", "Special allowance"), (r"transport", "TransportAllowance", "Transport"),
    (r"\bbonus\b", "Bonus", "Bonus"), (r"\barrears?\b", "Arrears", "Arrears"), (r"medical", "MedicalAllowance", "Medical"),
]


def make_answer(report, df, periods_text, intent=None):
    if df.empty:
        return f"No data found for {periods_text}."
    if report == "workorder_summary" and "Margin" in df and "WorkOrderNo" in df:
        text = _wo_rank_text(df, periods_text, intent)
        other = _wo_other_leader(df, intent)
        return text + (f"\n\n{other}" if other else "")

    if report == "workorder_monthly" and "CTC" in df:
        bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
        sales = (intent or {}).get("pct_basis") == "sales"
        den = bs if sales else ct
        pct = f" ({(bs - ct) * 100 / den:.2f}% on {'base sales' if sales else 'CTC'})" if den else ""
        scope = ""
        if (intent or {}).get("included") or (intent or {}).get("category") or (intent or {}).get("department"):
            scope += " — only " + ((intent or {}).get("included") or " · ".join(
                v for v in ((intent or {}).get("category"), (intent or {}).get("department")) if v))
        if (intent or {}).get("excluded"):
            scope += f" — excluding {intent['excluded']}"
        return (f"**Work order-wise for {periods_text}{scope}:** {df['WorkOrderNo'].nunique()} work order(s), "
                f"{len(df)} row(s) (one per work order per salary month) · Base sales value {inr(bs)} · "
                f"CTC {inr(ct)} · Margin {inr(bs - ct)}{pct}")
    if report in ("summary", "trend") and "CTC" in df:
        flt = " · ".join(str(v) for v in ((intent or {}).get("department"), (intent or {}).get("category")) if v)
        if (intent or {}).get("included"):
            flt = (flt + " · " if flt else "") + f"only {intent['included']}"
        if (intent or {}).get("excluded"):
            flt = (flt + " · " if flt else "") + f"excluding {intent['excluded']}"
        flt = f" ({flt})" if flt else ""
        base = ""
        if "BaseSalesValue" in df:
            bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
            sales = (intent or {}).get("pct_basis") == "sales"
            den = bs if sales else ct
            pct = f" ({(bs - ct) * 100 / den:.2f}% on {'base sales' if sales else 'CTC'})" if den else ""
            base = f" · Base sales value {inr(bs)} · Margin {inr(df['Margin'].sum())}{pct}"
        if len(df) == 1:
            r = df.iloc[0]
            inv = (f" · **Invoice with GST {inr(r['InvoiceWithGST'])}**" if "InvoiceWithGST" in df and
                   re.search(r"invoice|gst|billing|billed", str((intent or {}).get("_q", "")), re.I) else "")
            gst = ""
            if "InvoiceWithGST" in df and "BaseSalesValue" in df and GST_ASK.search(str((intent or {}).get("_q", ""))):
                gst = (f"**GST on invoices for {periods_text}{flt}: {inr(r['InvoiceWithGST'] - r['BaseSalesValue'])}** "
                       f"(invoice with GST {inr(r['InvoiceWithGST'])} − base sales value {inr(r['BaseSalesValue'])}, "
                       f"{GST_RATE}% GST)\n\n")
            return gst + (f"**Salary for {periods_text}{flt}:** {int(r['Employees'])} employees · "
                    f"CTC {inr(r['CTC'])}{base}{inv}"
                    + (f" · Gross {inr(r['GrossAmount'])}" if "GrossAmount" in df else "")
                    + (f" · Take-home {inr(r['TakeHome'])}" if "TakeHome" in df else ""))
        chg = ""
        if len(df) >= 2 and "Period" in df:                # changes between months (directors ask "compare")
            def _chg(a, b):
                parts = []
                for c, lab in (("CTC", "CTC"), ("BaseSalesValue", "Base sales"), ("Margin", "Margin")):
                    if c in df and pd.notna(a[c]) and pd.notna(b[c]):
                        d = b[c] - a[c]
                        parts.append(f"{lab} {'▲' if d >= 0 else '▼'} {inr(abs(d))}"
                                     + (f" ({abs(d) * 100 / abs(a[c]):.1f}%)" if a[c] else ""))
                return " · ".join(parts)
            if len(df) > 2:                                # whole range: first month -> last month
                chg += f"\n\n**{df.iloc[-1]['Period']} vs {df.iloc[0]['Period']}:** " + _chg(df.iloc[0], df.iloc[-1])
            chg += f"\n\n**{df.iloc[-1]['Period']} vs {df.iloc[-2]['Period']}:** " + _chg(df.iloc[-2], df.iloc[-1])
            if (intent or {}).get("ask_rank"):              # "which month had the highest margin"
                fc = (intent or {}).get("focus") or "CTC"
                fc = fc if fc in df else "CTC"
                s = pd.to_numeric(df[fc], errors="coerce")
                if fc == "MarginPct" and {"Margin", "CTC"} <= set(df.columns):
                    s = df["Margin"] * 100 / df["CTC"].replace(0, float("nan"))
                if s.notna().any():
                    hi, lo = s.idxmax(), s.idxmin()
                    val = (lambda v: f"{v:.2f}%") if fc == "MarginPct" else (lambda v: f"{int(v):,}") if fc == "Employees" else inr
                    name = {"MarginPct": "margin %", "BaseSalesValue": "base sales", "Employees": "headcount"}.get(fc, fc if fc.isupper() else fc.lower())
                    chg = (f"\n\n**Highest {name}: {df.loc[hi, 'Period']} — {val(s[hi])}** · lowest: "
                           f"{df.loc[lo, 'Period']} — {val(s[lo])}") + chg
        if re.search(r"\b(combined|together|total|sum|both|overall|milake|mila\s+ke)\b", str((intent or {}).get("_q", "")), re.I) \
                and "BaseSalesValue" in df:                # "total margin for aug and sep combined": the sum first
            bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
            chg = (f"**Combined {len(df)} months ({periods_text}){flt}: margin {inr(bs - ct)}"
                   + (f" ({(bs - ct) * 100 / ct:.2f}% on CTC)" if ct else "") + f"** · CTC {inr(ct)} · base sales value "
                   f"{inr(bs)}\n\n" + chg.lstrip("\n"))
            return chg.rstrip()
        return (f"**{len(df)} months** ({periods_text}){flt} · Total CTC {inr(df['CTC'].sum())}{base}"
                + (f" · Total take-home {inr(df['TakeHome'].sum())}" if "TakeHome" in df else "") + chg)
    if report == "workorder_detail" and "Name" in df:
        n_wo = df["WorkOrderNo"].nunique() if "WorkOrderNo" in df else 1
        wo = df["WorkOrderNo"].iloc[0] if n_wo == 1 else f"search: {n_wo} work orders matched"
        inv = ""
        if "BaseSalesValue" in df:
            keys = [c for c in ("WorkOrderNo", "Year", "Month") if c in df]
            base = pd.to_numeric(df["BaseSalesValue"], errors="coerce")
            value = base.groupby([df[k] for k in keys]).max().sum()    # once per work order per month
            ctc = df["CTC"].sum()
            pct = f" ({(value - ctc) * 100 / ctc:.2f}% on CTC)" if ctc else ""
            inv = (f" · Base sales value {inr(value)} · Margin {inr(value - ctc)}{pct}")
        return (f"**Work order {wo}:** {len(df)} record(s) · "
                f"{df['Employee'].nunique()} employee(s) · "
                + (f"Take-home {inr(df['TakeHome'].sum())} · " if "TakeHome" in df else "") +
                f"CTC {inr(df['CTC'].sum())}{inv}")
    if report in ("department", "category") and "CTC" in df:
        col = report.capitalize()
        fc, fl = (intent or {}).get("focus", "CTC"), (intent or {}).get("focus_label", "CTC")
        low = (intent or {}).get("sort") == "asc"
        head = ""
        if fc in df and len(df):
            g = df.groupby(col).agg({c: "sum" for c in ("CTC", "BaseSalesValue", "Margin", "Employees",
                                                         "TakeHome", "GrossAmount") if c in df})
            if fc == "MarginPct":
                sales = (intent or {}).get("pct_basis") == "sales"
                g["MarginPct"] = g["Margin"] * 100 / (g["BaseSalesValue"] if sales else g["CTC"]).replace(0, float("nan"))
            if fc in g:
                name = g[fc].idxmin() if low else g[fc].idxmax()
                v = g.loc[name, fc]
                val = f"{v:.2f}%" if fc == "MarginPct" else (f"{int(v):,}" if fc == "Employees" else inr(v))
                head = f"**{'Lowest' if low else 'Highest'} {fl.lower()}: {name} — {val}**\n\n"
        a_ = (intent or {}).get("all_rows")
        if a_:
            return (head + f"Top {len(df)} of {a_['n']} {col.lower()}(s) for {periods_text}, by {fl.lower()}. "
                    f"All {a_['n']}: CTC {inr(a_.get('CTC'))} · Base sales value {inr(a_.get('BaseSalesValue'))} · "
                    f"Margin {inr(a_.get('Margin'))}")
        return (head + f"{col}-wise for {periods_text}: {df[col].nunique()} {col.lower()}(s) · "
                f"CTC {inr(df['CTC'].sum())} · Base sales value {inr(df['BaseSalesValue'].sum())} · "
                f"Margin {inr(df['Margin'].sum())}")
    if report == "employee_detail" and "Name" in df:
        ctc = pd.to_numeric(df["CTC"], errors="coerce") if "CTC" in df else pd.Series(dtype=float)
        return (f"Found **{len(df)} salary record(s)** for {', '.join(df['Name'].astype(str).unique()[:5])}"
                + (f" · CTC {inr(ctc.sum())}" if ctc.notna().any() else "") + ".")
    if report in ("joiners", "leavers") and "Employee" in df:
        who = "joined" if report == "joiners" else "resigned / left"
        return f"**{df['Employee'].nunique():,} employee(s) {who}** in {periods_text}."
    if report in ("deductions", "components"):
        num = df.drop(columns=[c for c in ("Employees", "Period", "Year", "Month", "IncomeTax_2") if c in df]
                      ).select_dtypes("number")
        sums = num.sum()
        sums = sums[sums != 0]
        asked = [(c, lab) for rx, c, lab in (DEDUCTION_ASK if report == "deductions" else COMPONENT_ASK)
                 if re.search(rx, str((intent or {}).get("_q", "")), re.I) and c in num]
        if asked:                                       # "pf employer contribution", "esic", "bonus": that figure first
            if any(c in ("PF_Employer", "PF_Employee") for c, _ in asked):
                asked = [a for a in asked if a[0] != "TotalPF"]
            head = " · ".join(f"{lab} {inr(num[c].sum())}" for c, lab in asked)
            rest = [k for k in sums.index if k not in [c for c, _ in asked]]
            nice_ = {"IncomeTax": "TDS", "PF_Employee": "PF (employee)", "PF_Employer": "PF (employer)", "TotalPF": "Total PF",
                     "ESIC_Employee": "ESIC (employee)", "ESIC": "ESIC (employer)", "CORP_NPS": "Corporate NPS",
                     "ConveyanceAllowance": "Conveyance", "SpecialAllowance": "Special allowance",
                     "TransportAllowance": "Transport", "MedicalAllowance": "Medical", "GrossAmount": "Gross"}
            other = "Other deductions" if report == "deductions" else "Other components"
            return (f"**{head}** for {periods_text}."
                    + (f" {other}: " + " · ".join(f"{nice_.get(k, k)} {inr(sums[k])}" for k in rest) if rest else ""))
        if sums.empty and num.notna().any().any():      # "pf of <name>" with nothing deducted: say so, not "1 row"
            return (f"**{'Deductions' if report == 'deductions' else 'Salary components'} for {periods_text}:** "
                    "all ₹0.00.")
        nice = {"IncomeTax": "TDS", "PF_Employee": "PF (employee)", "PF_Employer": "PF (employer)", "TotalPF": "Total PF",
                "ESIC_Employee": "ESIC (employee)", "ESIC": "ESIC (employer)", "CORP_NPS": "Corporate NPS",
                "ConveyanceAllowance": "Conveyance", "SpecialAllowance": "Special allowance",
                "TransportAllowance": "Transport", "MedicalAllowance": "Medical", "GrossAmount": "Gross"}
        if not sums.empty:
            return (f"**{'Deductions' if report == 'deductions' else 'Salary components'} for {periods_text}:** "
                    + " · ".join(f"{nice.get(k, k)} {inr(v)}" for k, v in sums.items()))
    return f"Found **{len(df)} row(s)** for {periods_text}."


# ------------------------------ Main logic ------------------------------
WRITE_REQUEST = re.compile(
    r"^\s*(?:please\s+|pls\s+|can you\s+|kindly\s+)?(delete|drop|truncate|insert|update|alter|"
    r"modify|remove|change|edit|set|add|correct|overwrite|replace)\b"
    r"|\b(delete\s+from|drop\s+table|update\s+\S+\s+set|insert\s+into|truncate\s+table)\b",
    re.IGNORECASE)


PAYROLL_TERMS = {
    "salary", "salaries", "payroll", "ctc", "gross", "net", "pay", "paid", "wage", "wages",
    "numbers", "figures", "figure", "stats", "statistics", "kpi", "kpis", "dashboard", "overview",
    "employee", "employees", "emp", "staff", "headcount", "people", "department", "departments",
    "dept", "category", "categories", "workorder", "workorders", "wo", "order", "orders",
    "invoice", "billing", "base", "sales", "sale", "revenue", "margin", "margins", "profit", "tds", "tax", "pf", "esic", "pt", "nps",
    "basic", "hra", "bonus", "arrear", "arrears", "allowance", "allowances", "deduction",
    "deductions", "joiner", "joiners", "joined", "joinee", "joinees", "resigned", "leavers",
    "left", "exit", "trend", "compare", "cost", "amount", "total", "summary", "report", "list",
    "details", "detail", "breakup", "month", "months", "monthwise", "month_year", "code",
    "highest", "lowest", "top",
}
SMALLTALK = re.compile(
    r"^\s*(hi+|hello|hey|hlo|namaste|good\s+(?:morning|afternoon|evening|night)|thanks?|thank\s*you|"
    r"thx|ty|ok+|okay|k|great|nice|cool|awesome|super|perfect|excellent|good|fine|wow|"
    r"let'?s\s+go|lets\s+go|go\s+ahead|yes|yeah|yep|no|nope|bye|see\s+you|well\s+done|good\s+job|"
    r"hmm+|test(?:ing)?|who\s+are\s+you|what\s+can\s+you\s+do|help)\b", re.IGNORECASE)

SUGGEST = ("Try: *salary for aug-26* · *department wise salary for aug-26* · "
           "*which workorder has highest margin in aug-26* · *pf, tds and esic for aug-26*")


def has_payroll_signal(question):
    wo, q = find_workorder(question)
    if wo or parse_periods(q):
        return True
    tokens = set(re.findall(r"[a-z_]+", q.lower()))
    if tokens & PAYROLL_TERMS or re.search(r"\d{3,}", q):
        return True
    try:
        return bool(find_known(q, "Department") or find_known(q, "Category"))
    except Exception:
        return False


def chat_reply(question):
    t = question.lower()
    if re.search(r"\b(hi+|hello|hey|hlo|namaste|good\s+(morning|afternoon|evening))\b", t):
        text = "Hello! 👋 I'm the Workorder Dummy Payroll AI. Ask me anything about salary data."
    elif re.search(r"\b(thanks?|thank\s*you|thx|ty)\b", t):
        text = "You're welcome! 😊 What would you like to check next?"
    elif re.search(r"\b(bye|see\s+you)\b", t):
        text = "Bye! 👋 Come back any time you need payroll figures."
    elif re.search(r"\b(who\s+are\s+you|what\s+can\s+you\s+do|help)\b", t):
        text = ("I answer payroll questions from SQL Server (read-only): monthly totals, departments, "
                "categories, deductions, salary breakup, employees, work orders & margins, trends, "
                "joiners and leavers.")
    elif re.search(r"\b(let'?s\s+go|lets\s+go|go\s+ahead|yes|yeah|yep|ok+|okay)\b", t):
        text = "Sure! 🚀 Go ahead and ask your payroll question."
    elif re.search(r"\b(great|nice|cool|awesome|super|perfect|excellent|good|wow|well\s+done|good\s+job)\b", t):
        text = "Glad that helped! 👍 What would you like to see next?"
    else:
        text = "I didn't find a payroll question in that. 🤔 Could you tell me what you'd like to see?"
    return {"text": f"{text}\n\n{SUGGEST}", "note": "", "df": None, "report": "chat",
            "sql": "", "params": [], "intent": {}}


OPTION_TEMPLATES = {
    "summary": "salary summary for {p}", "department": "department wise salary for {p}",
    "category": "category wise salary for {p}", "deductions": "pf, tds and esic for {p}",
    "components": "salary breakup for {p}", "employee_list": "list all employees salary for {p}",
    "workorder_summary": "which workorder has highest margin in {p}", "joiners": "new joiners in {p}",
    "leavers": "leavers in {p}", "trend": "salary trend from {p}",
}


def period_phrase(question):
    ps = parse_periods(question)
    if not ps:
        return ""
    lab = [period_label(m, y).lower() for m, y in ps]
    return lab[0] if len(lab) == 1 else f"{lab[0]} to {lab[-1]}"


def suggest_options(question, first=None, n=4):
    p = period_phrase(question)
    order = [first] if first else []
    order += ["summary", "department", "workorder_summary", "deductions", "employee_list", "components"]
    out = []
    for rep in order:
        tpl = OPTION_TEMPLATES.get(rep)
        if not tpl or (rep == "trend" and " to " not in p):
            continue
        q = tpl.format(p=p).replace(" for ", " for ").strip()
        q = re.sub(r"\s+(for|in|from)\s*$", "", q)          # no month -> drop trailing "for"
        if q not in out:
            out.append(q)
    return out[:n]


def unknown_words(question, intent=None):
    """Words we don't recognise (ignores names/codes the question is about)."""
    _, q = find_workorder(question)
    skip = set()
    for v in ((intent or {}).get("employee"), (intent or {}).get("department"), (intent or {}).get("category")):
        if v:
            q = re.sub(re.escape(str(v)), " ", q, flags=re.IGNORECASE)
            skip |= name_words(v)
            w = [x for x in re.findall(r"[a-z0-9]+", str(v).lower()) if x not in NAME_FILLER]
            skip |= {w[i] + w[i + 1] for i in range(len(w) - 1)}
    tokens = {t for t in re.findall(r"[a-z_]+", q.lower().replace("month_year", " ")) if len(t) > 2}
    for v in ((intent or {}).get("department"), (intent or {}).get("category")):
        if v:
            vc = re.sub(r"[^a-z0-9]", "", str(v).lower())
            skip |= {t for t in tokens if len(t) >= 3 and t in vc}
    return tokens - SIMPLE_WORDS - MONTH_WORDS - PAYROLL_TERMS - skip


def clarify(question, text, options=None):
    return {"text": f"🤔 {text}", "note": "Pick an option below, or reply with a few words.",
            "df": None, "report": "clarify", "sql": "", "params": [], "intent": {},
            "options": options or suggest_options(question), "clarify_for": question}


def similar_values(column, text, limit=6):
    """Suggestions for 'did you mean' (read-only SELECT)."""
    word = re.sub(r"[^A-Za-z0-9\- ]", "", str(text)).strip()
    if len(word) < 2:
        return []
    key = word.split()[0][:4] if column == "WorkOrderNo" else word.split()[0]
    try:
        if column == "WorkOrderNo":
            sql, prm = (f"SELECT DISTINCT TOP ({limit}) [{column}] AS v FROM {T} "
                        f"WHERE [{column}] LIKE ? ORDER BY [{column}]"), [f"%{key}%"]
        else:   # match whole words only ("than" must not match "Rajasthan")
            sql, prm = (f"SELECT DISTINCT TOP ({limit}) [{column}] AS v FROM {T} "
                        f"WHERE [{column}] LIKE ? OR [{column}] LIKE ? ORDER BY [{column}]"), [f"{key}%", f"% {key}%"]
        df = run_query(sql, prm)
        return [str(v) for v in df["v"] if str(v).strip()]
    except Exception:
        return []


ANALYSIS_WORDS = {
    "greater", "less", "lesser", "more", "fewer", "above", "below", "exceed", "exceeds", "exceeding",
    "before", "after", "since", "older", "newer", "earlier", "later", "average", "avg", "mean", "median",
    "count", "many", "percentage", "percent", "ratio", "share", "expire", "expires", "expiring", "expired",
    "enddate", "startdate", "date", "dates", "duration", "tenure", "increase", "decrease", "growth",
    "change", "difference", "variance", "distinct", "unique", "missing", "blank", "null", "zero",
    "negative", "duplicate", "duplicates", "not", "without", "except", "excluding", "between",
    "minimum", "maximum", "min", "max", "sum",
}


@st.cache_data(ttl=3600, show_spinner=False)
def table_schema():
    df = run_query("SELECT COLUMN_NAME AS c, DATA_TYPE AS t FROM INFORMATION_SCHEMA.COLUMNS "
                   "WHERE TABLE_NAME = ? ORDER BY ORDINAL_POSITION", [TABLE])
    return [(str(c), str(t)) for c, t in zip(df["c"], df["t"])]


def wants_analysis(question, intent, unknown, keyword_sure):
    if not ANALYST_MODE:
        return False
    if intent.get("report") == "analysis":
        return True
    t = set(re.findall(r"[a-z]+", question.lower()))
    if t & ANALYSIS_WORDS:
        ps = parse_periods(question)          # "between apr-26 and aug-26" alone is a normal trend
        only_range = (t & ANALYSIS_WORDS) <= {"between"} and len(ps) >= 2
        only_count = (keyword_sure and (t & ANALYSIS_WORDS) <= COUNT_WORDS
                      and intent.get("report") in ("joiners", "leavers", "summary", "department", "category"))
        if not only_range and not only_count:
            return True
    return bool(unknown) and not keyword_sure


def p_view():
    """The AI queries this prepared view instead of the raw table, so the rules are already applied:
    BaseSales is each row's share of its work order's base sales (invoice x 100/118, once per WO per month),
    and months come ready as dates + 'Aug-2026' names."""
    body = (f"SELECT *, "
            f"MAX({BASE_EXPR}) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) * 1.0 "
            f"/ COUNT(*) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) AS [BaseSales], "
            f"MAX([InvoiceAmount]) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) * 1.0 "
            f"/ COUNT(*) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) AS [Inv], "
            f"DATEFROMPARTS([Year], [Month], 1) AS [SalaryMonth], "
            f"FORMAT(DATEFROMPARTS([Year], [Month], 1), 'MMM-yyyy') AS [SalaryMonthName], "
            f"DATEFROMPARTS(YEAR([month_year]), MONTH([month_year]), 1) AS [MYMonth], "
            f"FORMAT([month_year], 'MMM-yyyy') AS [MYMonthName], "
            f"{PAY_EXPR} AS [PaymentDate], "
            f"DATEFROMPARTS(YEAR({PAY_EXPR}), MONTH({PAY_EXPR}), 1) AS [PayMonth], "
            f"FORMAT({PAY_EXPR}, 'MMM-yyyy') AS [PayMonthName], "
            f"{TAKEHOME_EXPR} AS [TakeHome] "
            f"FROM {T}" + (" WHERE " + " AND ".join(_excl_clauses()) if _excl_clauses() else ""))
    tf = [(k, v["sql_expr"]) for k, v in _taught().items() if v.get("sql_expr")]   # v14: taught row formulas
    if not tf:
        return f"WITH p AS ({body})"
    return f"WITH p__base AS ({body}), p AS (SELECT *, " + ", ".join(f"({e}) AS [{k}]" for k, e in tf) + " FROM p__base)"


_NOT_ALIAS = r"(?:WHERE|GROUP|ORDER|JOIN|ON|INNER|LEFT|RIGHT|FULL|CROSS|OUTER|UNION|HAVING|WITH)\b"


def _rename_ai_p(q):
    """'WITH p AS (... FROM p ...) SELECT ... FROM p' -> the AI's own CTE becomes ai_p (no duplicate name)."""
    m = re.match(r"(?is)^with\s+(\[?p\]?)\s+as\s*\(", q)
    if not m:
        return q
    depth, i = 1, m.end()
    while i < len(q) and depth:
        depth += {"(": 1, ")": -1}.get(q[i], 0)
        i += 1
    body, rest = q[m.end():i], q[i:]

    def ref(mm):
        alias = mm.group(3)
        return f"{mm.group(1)} ai_p{alias}" if alias else f"{mm.group(1)} ai_p AS p"
    rest = re.sub(r"(?i)\b(FROM|JOIN)(\s+)\[?p\]?\b((?:\s+(?:AS\s+)?(?!" + _NOT_ALIAS + r")\w+)?)",
                  lambda mm: ref(mm), rest)
    return "WITH ai_p AS (" + body + rest


def with_view(ai_query):
    q = _rename_ai_p(ai_query.strip())
    if re.match(r"(?i)^with\b", q):                  # AI wrote its own WITH -> chain it
        return p_view() + ", " + re.sub(r"(?i)^with\s+", "", q, count=1)
    return p_view() + " " + q


def ai_sql(question, mode, error=None, bad_sql=None):
    eff = current_eff()
    cols = [(c, t) for c, t in table_schema() if eff is None or c in pe.TECHNICAL or eff.key_visible(c)]
    col_text = ", ".join(f"[{c}] {t}" for c, t in cols)
    hints = []
    for m, y in parse_periods(question):
        hints.append(f'"{period_label(m, y)}" = {MON[m - 1]} {y} -> \'{y}-{m:02d}-01\'')
    try:
        lm, ly = latest_period(mode)
        latest = f"{MON[lm - 1]}-{ly}"
    except Exception:
        latest = "unknown"
    PM, PN = {"month_year": ("MYMonth", "MYMonthName"),
              "payment": ("PayMonth", "PayMonthName")}.get(month_mode_of(question, mode), ("SalaryMonth", "SalaryMonthName"))
    system = f"""You write ONE Microsoft SQL Server (T-SQL) SELECT for a payroll analyst.
Query ONLY the prepared view [p] (write FROM p). [p] already exists - never create a CTE named p. One row = one employee's salary for one month.
Columns of p: {col_text}
What the columns mean (verified on the company's data):
- [Category] business line (e.g. Openvem, GML, Efirm). [Department] client department. [WorkOrderNo] client work order.
- [Employee] employee code, [Name]. [PaidDays] days paid in that salary month.
- [Salary] full-month salary rate. [CTC] = cost for that month = Salary x PaidDays / days in month.
- [GrossAmount] = CTC + Arrears - PF_Employer. [Arrears] = catch-up pay for EARLIER months that also have their
  own rows - never add Arrears to CTC (it would count those months twice).
- [IncomeTax] (= IncomeTax_2) is TDS. [Total PF] = PF_Employee + PF_Employer. [PT] (= PFTax) professional tax.
- [NetSalary] and [ChequeAmount] are NOT reliable - never use them; use [TakeHome] for take-home / net pay.
- Salary parts: Basic, HRA, SpecialAllowance, TransportAllowance, ConveyanceAllowance, Bonus, MedicalReimburse.
- [InvoiceAmount] = invoice incl. 18% GST, REPEATED on every employee row of the same WorkOrderNo + Year + Month.
  NEVER SUM([InvoiceAmount]) - it multiplies the invoice by the number of employees. Use SUM([Inv]) for invoice totals.
- [InvoiceDate]; [month_year] = LIABLE month = payment date - 1 month (paid 07-Oct-2026 -> 07-Sep-2026); it can
  include rows for older salary months. [Remarks] = salary month, or list of months covered by arrears.
- [WOStartDate]/[WOEndDate] work order period, [joindate]/[resigndate], [tdsopt] tax regime option.
Extra ready-made columns in p (use them!):
- [TakeHome] = GrossAmount - IncomeTax - PF_Employee - PT - ESIC_Employee - Deduction (real take-home).
- [Inv]: this row's share of its work order's invoice WITH GST. SUM([Inv]) over any rows = the correct invoice total.
- [BaseSales]: this row's share of its work order's base sales value (invoice without {GST_RATE}% GST).
  SUM([BaseSales]) over any rows = the correct base sales value. NEVER use [InvoiceAmount] for sales.
- [SalaryMonth] (date, 1st of month) and [SalaryMonthName] (text 'Aug-2026'): the month the salary is FOR
  (from [Year],[Month]; e.g. Aug-2026 salary is paid in September).
- [MYMonth] (date, 1st of month) and [MYMonthName] (text): LIABLE month = month of [month_year].
- [PaymentDate] (= month_year + 1 month), [PayMonth] (date, 1st of month), [PayMonthName] (text): the month the
  salary was PAID ("paid in oct-26", "payment month"). Liable month Sep-2026 = paid in Oct-2026.
Rules:
- Use [{PM}] / [{PN}] as the month for this question. Never group or filter by a raw date like [month_year].
- One month: WHERE [{PM}] = 'YYYY-MM-01'. Range: WHERE [{PM}] BETWEEN 'YYYY-MM-01' AND 'YYYY-MM-01'.
- Month-wise / each month: SELECT [{PN}] AS [Month] ... GROUP BY [{PM}], [{PN}] ORDER BY [{PM}].
- Show months as MMM-yyyy (use [{PN}], or FORMAT(date, 'MMM-yyyy') for other date columns).
- Margin = SUM([BaseSales]) - SUM([CTC]).
- Margin % (company formula, DEFAULT) = (SUM([BaseSales]) - SUM([CTC])) * 100.0 / NULLIF(SUM([CTC]), 0).
  ONLY if the user asks margin % "on base sales / on invoice / on sales / on revenue", divide by SUM([BaseSales]) instead.
- Round money and % with ROUND(x, 2). Count people with COUNT(DISTINCT [Employee]).
- Dates of other columns (WOEndDate, WOStartDate, joindate, resigndate, InvoiceDate): compare that column.
  "greater than jan-28" on a date means >= '2028-02-01'.
- Work-order lists: GROUP BY [WorkOrderNo]. [Total PF] has a space. Put column names in [brackets].
- Latest salary month in data: {latest}. Today: {datetime.date.today():%Y-%m-%d}.
- Only one SELECT. No INSERT/UPDATE/DELETE/INTO/EXEC/DECLARE/SET, no semicolons, no comments.
- Return ONLY the columns the question asks for. {'Do NOT add ' + ', '.join(sorted(set(hidden_now()) - asked_groups(question))) + ' columns unless asked.' if set(hidden_now()) - asked_groups(question) else ''}
Example - "total ctc, base sales, margin and margin % each month from apr-26 to aug-26":
SELECT [{PN}] AS [Month], ROUND(SUM([CTC]),2) AS [Total CTC], ROUND(SUM([BaseSales]),2) AS [Base Sales],
 ROUND(SUM([BaseSales])-SUM([CTC]),2) AS [Margin],
 ROUND((SUM([BaseSales])-SUM([CTC]))*100.0/NULLIF(SUM([CTC]),0),2) AS [Margin % on CTC]
FROM p WHERE [{PM}] BETWEEN '2026-04-01' AND '2026-08-01' GROUP BY [{PM}], [{PN}] ORDER BY [{PM}]
{taught_prompt()}
Reply ONLY as JSON: {{"sql": "<query>", "title": "<short title of the result>"}}"""
    user = question + ("\nMonth hints: " + "; ".join(hints) if hints else "")
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    if error:
        msgs += [{"role": "assistant", "content": json.dumps({"sql": bad_sql or ""})},
                 {"role": "user", "content": f"That query failed: {error}\nFix it. Same JSON format."}]
    resp = ollama.chat(model=MODEL, messages=msgs, format="json", keep_alive="60m",
                       options={"temperature": 0, "num_predict": 450, "num_ctx": 4096, "num_thread": AI_THREADS})
    data = json.loads(resp["message"]["content"])
    sql = str(data.get("sql") or "").strip().rstrip(";").strip()
    sql = re.sub(r"^```(?:sql)?|```$", "", sql).strip()
    return sql, str(data.get("title") or "Result").strip()


def prettify(df):
    """Month-like dates -> 'Aug-2026', other dates -> '15-Aug-2026', numbers rounded to 2 decimals."""
    df = df.copy()
    for c in df.columns:
        col = df[c]
        vals = col.dropna()
        if vals.empty:
            continue
        if pd.api.types.is_float_dtype(col):
            df[c] = col.round(2)
            continue
        if pd.api.types.is_datetime64_any_dtype(col) or vals.map(
                lambda v: isinstance(v, (datetime.date, datetime.datetime, pd.Timestamp))).all():
            d = pd.to_datetime(col, errors="coerce")
            monthly = (d.dropna().dt.day == 1).all()
            df[c] = d.dt.strftime("%b-%Y" if monthly else "%d-%b-%Y")
    return df


MONEY_COL = re.compile(r"amount|ctc|salary|base|margin|invoice|gross|net|pf|tax|value|cost|bonus|"
                       r"basic|hra|allow|arrear|esic|pt|nps|deduct|sales|revenue|pay|tds(?!\s*opt)|take.?home",
                       re.IGNORECASE)


def fmt_value(col, v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and pd.notna(v):
        if re.search(r"%|\bpct\b|percent|ratio", str(col), re.IGNORECASE):    # "Margin % on CTC" is a %, not ₹
            return f"{v:,.2f}%"
        if MONEY_COL.search(str(col)) and not re.search(r"count|employees|days|month|year|pct|percent",
                                                        str(col), re.IGNORECASE):
            return inr(v)
        return f"{v:,.2f}".rstrip("0").rstrip(".") if isinstance(v, float) else f"{v:,}"
    if isinstance(v, (datetime.date, datetime.datetime, pd.Timestamp)):
        t = pd.Timestamp(v)
        return t.strftime("%b-%Y" if t.day == 1 else "%d-%b-%Y")
    return str(v)


def analysis_text(df, title):
    if df.empty:
        return (f"**{title}:** no rows match. The filters may be too strict (for example month_year and "
                f"Month/Year pointing to different months) - please check the SQL below.")
    if df.shape == (1, 1):
        return f"**{title}:** {fmt_value(df.columns[0], df.iat[0, 0])}"
    if len(df) == 1:
        parts = [f"{c}: **{fmt_value(c, df.iloc[0][c])}**" for c in list(df.columns)[:8]]
        return f"**{title}** · " + " · ".join(parts)
    more = " (first 1,000 shown)" if len(df) >= ROW_LIMIT else ""
    return f"**{title}** · {len(df):,} row(s){more}."


def analyst_answer(question, mode, notes, t0):
    """AI writes a SELECT (checked), we run it read-only. Returns a result dict, or None if it failed."""
    err, sql, title = None, "", "Result"
    for attempt in range(2):                   # one retry with the error message
        try:
            ai_q, title = ai_sql(question, mode, err, sql)
            eff = current_eff()
            if eff is not None and eff.sql_violations(ai_q):
                audit_policy_event("sql_blocked", ", ".join(eff.sql_violations(ai_q)))
                raise PermissionError("Blocked: the query used a column that is not available for this login. "
                                      "Use only the columns listed.")
            sql = with_view(ai_q)
            t1 = time.time()
            df = hide_columns(prettify(run_ai_query(sql)), question)
            t_sql = time.time() - t1
            n = [f"Understood as: **AI analysis** (query written by AI, checked read-only)."] + notes
            if len(df) > 1 and df.duplicated().any():  # e.g. one row per employee when a work order list was asked
                if len(df) >= ROW_LIMIT:               # repeats filled the row limit -> read all rows, then de-duplicate
                    df = hide_columns(prettify(run_ai_query(sql, limit=CSV_LIMIT)), question)
                before = len(df)
                df = df.drop_duplicates().head(ROW_LIMIT).reset_index(drop=True)
                n.append(f"{before - len(df):,} repeated row(s) removed.")
            n.append("⚠️ Please check the SQL tab for how this was calculated.")
            n.append(f"⏱️ AI+SQL: {time.time() - t0:.1f}s · SQL: {t_sql:.1f}s")
            return {"text": analysis_text(df, title), "note": " ".join(n), "df": df,
                    "report": "analysis", "sql": sql, "params": [], "intent": {"title": title}}
        except PermissionError as e:
            err = str(e)
        except Exception as e:
            err = f"{e.__class__.__name__}: {str(e)[:300]}"
            if "timeout" in err.lower() or "timed out" in err.lower():
                err = f"the AI took longer than {AI_TIMEOUT_SEC}s"
                break                                # a retry would only make the user wait twice as long
    import logging
    logging.getLogger("payroll").warning("AI analysis failed for %r: %s", question, err)    # details for the admin log
    plain = ("the AI took too long" if "longer than" in str(err) else
             "that needs data the chatbot isn't allowed to show you" if "Blocked" in str(err) else
             "the AI couldn't write a working query for it")
    notes.append(f"I tried an AI-written query, but {plain}")
    return None


EMP_WORDS = {"employee", "employees", "emp", "who", "whose", "whom", "person", "people", "staff", "name", "names"}
METRICS = [   # (words in question, column, label)
    (("take home", "takehome", "net salary", "net pay", "net", "in hand", "inhand"), "TakeHome", "Take-home"),
    (("ctc", "cost to company"), "CTC", "CTC"),
    (("gross",), "GrossAmount", "Gross salary"),
    (("basic",), "Basic", "Basic"),
    (("hra",), "HRA", "HRA"),
    (("tds", "income tax", "tax"), "IncomeTax", "TDS / Income tax"),
    (("pf",), "Total PF", "Total PF"),
    (("bonus",), "Bonus", "Bonus"),
    (("arrear", "arrears"), "Arrears", "Arrears"),
    (("paid days", "paiddays"), "PaidDays", "Paid days"),
    (("salary", "pay", "paid"), "CTC", "CTC"),         # plain "salary" = CTC (take-home is hidden unless asked)
]
EMP_RANK_WORDS = {"highest", "lowest", "top", "bottom", "max", "maximum", "min", "minimum", "most", "least",
                  "biggest", "smallest", "best", "worst"}


def wants_emp_ranking(question):
    """'employee whose ctc is highest in aug-26', 'top 10 employees by net salary'."""
    t = question.lower()
    words = set(re.findall(r"[a-z]+", t))
    if (not (words & EMP_WORDS) and not WO_WORDS.search(question)
            and not (words & {"department", "departments", "dept", "category", "categories", "client"})
            and re.search(r"\b(?:highest|lowest|max\w*|min\w*|top|sabse\s+(?:zyada|jyada|kam))\s+(?:paid\s+)?(?:salary|salaries|ctc|pay|package)\b|"
                          r"\b(?:salary|ctc)\s+(?:is\s+)?(?:the\s+)?(?:highest|lowest|max\w*|min\w*)\b", t)):
        words = words | {"employee"}                     # "whats the highest salary" = who earns the most
    if not (words & EMP_RANK_WORDS) and not re.search(r"\bsabse\b", t):
        return None
    if not (words & EMP_WORDS) or WO_WORDS.search(question):
        return None
    if re.search(r"\bwhich\s+(?:departments?|depts?|categor(?:y|ies))\b", t):
        return None                                   # "which department has the most employees" ranks departments
    for keys, col, label in METRICS:
        if any(re.search(rf"\b{re.escape(k)}\b", t) for k in keys):
            return col, label
    return None


def employee_ranking(question, mode, col, label, notes, t0):
    sort, top = ranking(question)
    top = top or 10
    periods = parse_periods(question)
    if not periods:
        periods = [latest_period(mode)]
        notes.append("No month mentioned, showing the latest month in the data.")
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    f = understand_filters(question, f, notes)
    where, params = build_where(mode, periods, f)
    order = "ASC" if sort == "asc" else "DESC"
    sql = (f"SELECT TOP ({int(top)}) [Employee], MAX([Name]) AS Name, MAX([Department]) AS Department, "
           f"MAX([Category]) AS Category, MAX([WorkOrderNo]) AS WorkOrderNo, "
           f"SUM({TAKEHOME_EXPR if col == 'TakeHome' else '[' + col + ']'}) AS [{label}] "
           f"FROM {T} {where} GROUP BY [Employee] ORDER BY [{label}] {order}")
    t1 = time.time()
    df = hide_columns(run_query(sql, params), question)
    ptxt = ", ".join(period_label(m, y) for m, y in periods)
    if df.empty:
        text = f"No data found for {ptxt}."
    else:
        r = df.iloc[0]
        val = "" if label not in df else inr(r[label]) if col != "PaidDays" else f"{r[label]:,.1f}"
        text = (f"**{'Lowest' if sort == 'asc' else 'Highest'} {label} for {ptxt}: {r['Name']} "
                f"(emp {r['Employee']})** · {val} · {r['Department'] or ''}. Top {len(df)} below.")
    notes = [f"Understood as: **Employee ranking by {label}**."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI: {t1 - t0:.1f}s · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "employee_ranking", "sql": sql,
            "params": params, "intent": {"metric": label, "sort": sort, "top": top}}


MONTHLY_PHRASE = re.compile(r"\b(each|every|per|by)\s+month\b|\bmonth\s*-?\s*wise\b|\bmonthly\b|\bmonthwise\b|"
                            r"\bmonth\s+on\s+month\b|\bmom\b", re.IGNORECASE)
TOTAL_WORDS = {"ctc", "base", "sales", "margin", "margins", "gross", "net", "salary", "salaries", "payroll", "cost",
               "invoice", "profit", "percentage", "percent", "total", "totals", "employees", "headcount", "pay", "paid",
               "spend", "spent", "people"}
MONTH_RANK = re.compile(r"\bwhich\s+months?\b|\bmonths?\s+(?:had|has|with|was)\b|\bbest\s+month\b|\bworst\s+month\b", re.I)


def is_monthly_totals(question):
    """Month-by-month totals: 'every month', 'which month had the highest margin', 'how much did ctc increase from
    apr-26 to sep-26' (no department / category / work order / employee in it)."""
    t = set(re.findall(r"[a-z]+", question.lower()))
    if WO_WORDS.search(question) or t & {"department", "departments", "dept", "category", "categories",
                                          "whose", "who", "name", "names", "employee", "emp"}:
        return False
    if MONTH_RANK.search(question) and t & TOTAL_WORDS:
        return True
    if MOVE_RX.search(question) and len(parse_periods(question)) >= 2 and t & TOTAL_WORDS:
        return True
    if not MONTHLY_PHRASE.search(question) or not (t & TOTAL_WORDS):
        return False
    return not (t & {"top", "highest", "lowest"})


# ------------------------- definitions / formulas (no SQL) -------------------------
EXPLAIN_ASK = re.compile(
    r"\b(formula|formulas|formulae|how\s+(is|are|do\s+you|does\s+it|to)\s+\w*\s*(calculat|comput|work|arriv|deriv)"
    r"|calculat\w*\s+(of|for)|meaning|define|definition|explain|what\s+(is|are|does)|what's|mean[s]?)\b",
    re.IGNORECASE)
GLOSSARY = [   # (words that name the term, title, explanation)
    (("margin %", "margin percent", "margin percentage", "margin pct", "margin ratio"), "Margin %",
     lambda: "**Margin % = Margin ÷ CTC × 100**  (company formula, used by default)\n\n"
             "where **Margin = Base sales value − CTC**.\n\n"
             "Example: base sales ₹12,000, CTC ₹10,000 → margin ₹2,000 → **20.00%** (2,000 ÷ 10,000).\n\n"
             "Want it on sales instead? Ask *margin % on base sales* → Margin ÷ Base sales × 100 "
             "(same example: 16.67%)."),
    (("margin", "profit"), "Margin",
     lambda: "**Margin = Base sales value − CTC**\n\n"
             f"- **Base sales value** = Invoice amount × 100 / {100 + GST_RATE} (invoice without {GST_RATE}% GST), "
             "counted **once per work order + salary Month + Year** (the invoice is repeated on every employee row).\n"
             "- **CTC** = total of all employees' CTC in that work order / month.\n"
             "- **Margin %** = Margin ÷ CTC × 100 (company formula). On base sales: ask *margin % on base sales*.\n\n"
             "Example: work order ABC, Aug-2026 → invoice ₹14,160 → base sales ₹12,000; "
             "CTC ₹6,000 + ₹4,000 = ₹10,000 → **margin ₹2,000 (20.00% on CTC)**."),
    (("base sales", "base sale", "basic sales", "base value", "base amount", "sales value"), "Base sales value",
     lambda: f"**Base sales value = Invoice amount × 100 / {100 + GST_RATE}**  (invoice {100 + GST_RATE} = base 100)\n\n"
             "The invoice is repeated on every employee row of a work order, so it is taken **once per work "
             "order + salary Month + Year** (not per month_year), not added up per employee.\n\nExample: invoice ₹14,160 → base sales ₹12,000."),
    (("invoice",), "Invoice amount",
     lambda: f"**Invoice amount** is the billed value **including {GST_RATE}% GST**. It is stored on every employee "
             f"row of the work order. The chatbot uses Base sales value (= invoice × 100 / {100 + GST_RATE}) for sales and margin."),
    (("ctc", "cost to company"), "CTC",
     lambda: "**CTC (Cost to Company)** = the full monthly cost of the employee to the company, from the "
             "**CTC** column. Totals add CTC of every employee row (per month / work order / department)."),
    (("net salary", "take home", "net pay", "takehome", "in hand"), "Take-home",
     lambda: "**Take-home = Gross − TDS − PF (employee) − PT − ESIC (employee) − other deductions**\n\n"
             "The chatbot calculates this itself; the NetSalary and ChequeAmount columns are not used "
             "(cheque is blank for months paid through arrears)."),
    (("arrear", "arrears"), "Arrears",
     lambda: "**Arrears** = catch-up salary for earlier months, paid in a later run. Those earlier months also "
             "have their own rows (with their own CTC and invoice), so arrears are **not** added to CTC for margin."),
    (("gross",), "Gross salary",
     lambda: "**Gross = CTC + Arrears − PF (employer)**  (the GrossAmount column)."),
    (("month_year", "month year", "liable month", "payment date", "payment month"), "Liable month & payment date",
     lambda: "**Liable month** = the **month_year** column = payment date − 1 month (paid 07-Oct-2026 → "
             "07-Sep-2026). One liable month can contain rows for several older salary months (arrears).\n\n"
             "**Payment date** = month_year + 1 month. It is calculated by the chatbot (not stored in SQL Server). "
             "Ask e.g. *salary paid in oct-26*.\n\n**Salary month** = Month + Year columns: the month the salary "
             "is **for** (Sep-26 salary → Month 9, Year 2026)."),
    (("salary month", "month and year", "month/year"), "Salary month",
     lambda: "**Salary month** = the **Month** + **Year** columns: the month the salary is **for**. "
             "Example: Aug-2026 salary is paid in September."),
    (("pf", "provident fund"), "PF",
     lambda: "**Total PF = PF (employee) + PF (employer)**. Only the employee part is deducted from take-home."),
    (("tds", "income tax"), "TDS / Income tax",
     lambda: "**TDS / Income tax** = tax deducted from salary, from the **IncomeTax** column."),
]


STRONG_ASK = re.compile(r"\b(formula|formulas|formulae|define|definition|meaning|explain|calculated|calculation|"
                        r"kaise|kese|nikal\w*|matlab|samjha\w*|kya\s+(hai|he|hota|hoti)|"
                        r"computed|derived|how\s+(is|are|do|does|to)\b.*\b(calculat|comput|work|deriv))", re.IGNORECASE)
DATA_ASK = re.compile(r"\b(total|sum|show|give|list|how\s+much|amount\s+of|for\s+(each|every|the\s+month)|"
                      r"top|highest|lowest|wise)\b", re.IGNORECASE)


def wants_explanation(question):
    strong = STRONG_ASK.search(question)
    if not strong and not (EXPLAIN_ASK.search(question) and not DATA_ASK.search(question)):
        return None
    t = " " + re.sub(r"\s+", " ", re.sub(r"[^a-z0-9_%/ ]", " ", question.lower()).replace("%", " % ")) + " "
    for words, title, text in GLOSSARY:           # most specific first (margin % before margin)
        if any(f" {' '.join(w.replace('%', ' % ').split())} " in t for w in words):
            return title, text()
    return None


NO_EXPLAIN = {"definition", "chat", "chat_ai", "clarify", "blocked", "pref", "policy", "policy_confirm", "admin_info",
              "teach", "teach_confirm", "reminder"}


def _fmt_cell(col, v):
    try:
        return fmt_value(col, v)
    except Exception:
        return str(v)


def data_brief(df, max_rows=12):
    """Compact, pre-formatted facts for the AI - it must only use these numbers."""
    d = df.head(max_rows)
    lines = [" | ".join(str(c) for c in d.columns)]
    for _, r in d.iterrows():
        lines.append(" | ".join(_fmt_cell(c, r[c]) for c in d.columns))
    extra = []
    if len(df) > max_rows:
        extra.append(f"(showing {max_rows} of {len(df)} rows)")
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]) and MONEY_COL.search(str(c)) and len(df) > 1 \
                and not re.search(r"pct|percent|%|employee|month|year|days", str(c), re.IGNORECASE):
            extra.append(f"TOTAL {c} = {_fmt_cell(c, df[c].sum())}")
    return "\n".join(lines + extra)


@st.cache_data(ttl=600, show_spinner=False)
def explain_model():
    """A small model writes the text faster; SQL still uses MODEL."""
    if EXPLAIN_MODEL:
        return EXPLAIN_MODEL
    try:
        lst = ollama.list()
        models = lst.get("models", []) if isinstance(lst, dict) else getattr(lst, "models", [])
        names = [(m.get("model") or m.get("name")) if isinstance(m, dict) else (getattr(m, "model", "") or "")
                 for m in models]
        for want in ("qwen2.5:3b", "qwen2.5:3b-instruct", "llama3.2:3b", "qwen2.5:1.5b"):
            for n in names:
                if str(n).startswith(want):
                    return n
    except Exception:
        pass
    return MODEL


AI_EXPLAIN_ASK = re.compile(r"\b(explain|why|reason|insights?|analy[sz]e|analysis|comment|observations?|summari[sz]e|"
                         r"suggest|recommend|kyu|kyun|samjha\w*|batao\s+kyu)\b", re.IGNORECASE)


def _val(col, v):
    if col == "MarginPct":
        return f"{v:.2f}%"
    if col == "Employees":
        return f"{int(round(v)):,}"
    return inr(v)


def quick_explain(msg):
    """Short plain-language answer built from the result itself - instant and always correct."""
    df, report, it = msg.get("df"), msg.get("report"), (msg.get("intent") or {})
    if df is None or df.empty:
        return ""
    fc, fl = it.get("focus", "CTC"), it.get("focus_label", "CTC")
    sales = it.get("pct_basis") == "sales"
    basis = "base sales" if sales else "CTC"
    lab = {"trend": "Period", "summary": "Period", "department": "Department", "category": "Category",
           "workorder_summary": "WorkOrderNo"}.get(report)
    if not lab or lab not in df:
        return ""
    g = df.groupby(lab, sort=False).sum(numeric_only=True)
    if fc == "MarginPct":
        if not {"Margin", "CTC"} <= set(g.columns):
            return ""
        g["MarginPct"] = g["Margin"] * 100 / (g["BaseSalesValue"] if sales else g["CTC"]).replace(0, float("nan"))
    if fc not in g:
        fc, fl = ("CTC", "CTC") if "CTC" in g else (None, None)
    if fc is None or len(g) < 2:
        return ""
    ser = g[fc].dropna()
    if ser.empty:
        return ""
    name = fl + (f" (on {basis})" if fc == "MarginPct" else "")
    hi, lo = ser.idxmax(), ser.idxmin()
    out = []
    if lab == "Period":                       # month by month
        first, last = ser.index[0], ser.index[-1]
        if fc == "MarginPct":
            den = g["BaseSalesValue"] if sales else g["CTC"]
            out.append(f"**Overall {name} from {first} to {last} is {g['Margin'].sum() * 100 / den.sum():.2f}%.**")
        elif fc == "Employees":
            out.append(f"**Average headcount from {first} to {last} is {_val(fc, ser.mean())}.**")
        else:
            out.append(f"**Total {name} from {first} to {last} is {_val(fc, ser.sum())}** "
                       f"(average {_val(fc, ser.mean())} per month).")
        out.append(f"- **Highest:** {hi} — {_val(fc, ser[hi])}")
        out.append(f"- **Lowest:** {lo} — {_val(fc, ser[lo])}")
        prev = ser.index[-2]
        d = ser[last] - ser[prev]
        if fc == "MarginPct":
            out.append(f"- **{last} vs {prev}:** {'up' if d >= 0 else 'down'} {abs(d):.2f} points")
        elif ser[prev]:
            out.append(f"- **{last} vs {prev}:** {'up' if d >= 0 else 'down'} {_val(fc, abs(d))} "
                       f"({abs(d) * 100 / abs(ser[prev]):.1f}%)")
        if len(ser) > 2 and ser[first] and fc != "MarginPct":
            t = ser[last] - ser[first]
            out.append(f"- **Overall change since {first}:** {'up' if t >= 0 else 'down'} "
                       f"{abs(t) * 100 / abs(ser[first]):.1f}%")
    else:                                     # department / category / work order
        what = {"Department": "department", "Category": "category", "WorkOrderNo": "work order"}[lab]
        low_first = it.get("sort") == "asc"
        top, other = (lo, hi) if low_first else (hi, lo)
        out.append(f"**{'Lowest' if low_first else 'Highest'} {name}: {top} — {_val(fc, ser[top])}.**")
        out.append(f"- **{'Highest' if low_first else 'Lowest'}:** {other} — {_val(fc, ser[other])}")
        if fc != "MarginPct" and ser.sum() > 0 and ser[top] > 0:
            out.append(f"- {top} is **{ser[top] * 100 / ser.sum():.1f}%** of the total {fl if fl.isupper() else fl.lower()} "
                       f"({_val(fc, ser.sum())}) across {len(ser)} {what}(s)")
        else:
            out.append(f"- {len(ser)} {what}(s) compared")
        if fc in ("Margin", "MarginPct"):
            neg = int((ser < 0).sum())
            if neg:
                out.append(f"- ⚠️ {neg} {what}(s) have a **negative** margin")
    return "\n".join(out)


def explain_stream(msg):
    """Yields the ChatGPT-style answer word by word."""
    q = msg.get("question", "")
    hinglish = bool(re.search(r"\b(kya|hai|kaise|batao|dikhao|kitna|kitne|mujhe|ka|ki|ke)\b", q.lower()))
    lang = "Reply in simple Hinglish (Hindi in English letters), like the user." if hinglish else "Reply in English."
    if msg.get("report") == "chat_ai":
        system = ("You are the Payroll AI Assistant of Workorder Dummy, a friendly helper for the "
                  "finance team. Answer the user's general question clearly and briefly (max 150 words), like "
                  "ChatGPT. You CANNOT see payroll numbers in this reply - never invent company figures; if they "
                  "want data, suggest asking e.g. 'salary for aug-26'. "
                  "Be honest about what this chatbot can do. It CAN ONLY: answer payroll data questions (read-only), "
                  "explain formulas, remember display preferences, and set simple reminders inside this chatbot "
                  "('remind me tomorrow to ...'). It can NOT: change data, send e-mails or messages, make payments, "
                  "book leave or time off, submit timesheets, give financial planning advice, or see other systems. "
                  "Never offer or promise anything outside this list - say plainly it is not possible. " + lang)
        user = q
    else:
        system = ("You are a senior payroll/finance analyst at Workorder Dummy. Explain the result to the "
                  "user in plain words, like ChatGPT: start with the direct answer in one sentence, then 2-4 short "
                  "points on what stands out (highest/lowest, changes, share, anything unusual). Use ONLY the numbers "
                  "given below exactly as written (₹ with Indian commas, Cr/L are fine) - never invent or recalculate "
                  "new figures. Talk only about what the user asked; do not bring in PF, take-home, PT, ESIC, "
                  "gross or TDS unless the question mentions them. Do not mention SQL. Do NOT repeat every row of the "
                  "table (the user already sees it) - give only the direct answer and the key points. "
                  "Max 90 words. " + lang)
        fl = (msg.get("intent") or {}).get("focus_label")
        user = (f"User question: {q}\n" + (f"The user is asking about: {fl}\n" if fl else "") +
                f"What was calculated: {msg.get('text', '')}\n"
                f"Data:\n{data_brief(msg['df'], max_rows=8)}")
    buf, cut = "", False
    for chunk in ollama.chat(model=explain_model(), messages=[{"role": "system", "content": system},
                                                   {"role": "user", "content": user}],
                             stream=True, keep_alive="60m",
                             options={"temperature": 0.2, "num_predict": 400, "num_ctx": 2048,
                                      "num_thread": AI_THREADS}):
        buf += chunk["message"]["content"]
        if chunk.get("done"):
            cut = chunk.get("done_reason") == "length"
            break
        if "\n" in buf:                      # show complete lines only
            line, buf = buf.rsplit("\n", 1)
            yield line + "\n"
    buf = buf.strip()
    if cut:                                  # ran out of words: drop the half sentence
        m = re.search(r"^(.*[.!?%)])[^.!?%)]*$", buf, re.S)
        buf = m.group(1) if m and not buf.lstrip().startswith(("-", "*", "•")) else ""
    if buf.count("**") % 2:
        buf = buf.replace("**", "")
    if buf:
        yield buf


def chat_ai(question):
    return {"text": "", "note": "General answer by AI (not from payroll data).", "df": None,
            "report": "chat_ai", "sql": "", "params": [], "intent": {}}


# ---------------- follow-up questions: "and jul-26?", "what about IT?" ----------------
FOLLOW_UP = re.compile(r"^\s*(and|what\s+about|how\s+about|same\s+for|same|now|also|then|for|only|just|aur|ab)\b",
                       re.IGNORECASE)
PERIOD_TEXT = re.compile(
    r"\b(?:from\s+)?(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sept|sep|oct|nov|dec)(?:uary|ruary|ch|il|e|y|ust|ember|"
    r"tember|ober)?(?![a-z])[\s\-/',.]*(?:\d{4}|\d{2})|\d{4}[/\-]\d{1,2}|\d{1,2}[/\-](?:\d{4}|\d{2}))"
    r"(?:\s*(?:to|till|until|-)\s*(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sept|sep|oct|nov|dec)[a-z]*[\s\-/',.]*"
    r"(?:\d{4}|\d{2})))?", re.IGNORECASE)


def follow_up(question, prev):
    """Turns a short follow-up into a full question using the previous one. Returns None if not a follow-up."""
    if not prev:
        return None
    q = question.strip().rstrip("?").strip()
    short = len(q.split()) <= 5
    if not (FOLLOW_UP.match(q) or short) or wants_explanation(q):
        return None
    new_periods = PERIOD_TEXT.findall(q) and PERIOD_TEXT.search(q).group(0)
    rest = FOLLOW_UP.sub("", PERIOD_TEXT.sub(" ", q)).strip(" ,.?")
    rest = re.sub(r"\b(the|for|in|of|month|only|just)\b", " ", rest, flags=re.IGNORECASE).strip()
    dept = None
    try:
        dept = find_known(q, "Department") or find_known(q, "Category")
    except Exception:
        pass
    if not new_periods and not dept:
        return None
    if not FOLLOW_UP.match(q):
        # no "and / what about / same for": only a bare month or department counts as a follow-up
        left = rest
        if dept:
            left = re.sub(re.escape(dept), " ", left, flags=re.IGNORECASE)
            left = re.sub(r"\b(department|dept|category)\b", " ", left, flags=re.IGNORECASE)
        if left.strip(" ,.?"):
            return None                               # it's a new full question
    base = prev
    if new_periods:
        base = PERIOD_TEXT.sub(" ", base).strip() + " " + new_periods.strip()
    if dept and dept.lower() not in base.lower():
        base = f"{base} for {dept}"
    return re.sub(r"\s+", " ", base).strip()


SALES_BASIS = re.compile(
    r"\b(on|of|over|against|to|by|w\.?r\.?t\.?)\s+(the\s+)?(base\s*sales?|basic\s*sales?|base\s*amount|base\s*value|"
    r"sales?|invoice|billing|revenue|base|turnover)\b|\bsales\s*margin\b|\bmargin\s+on\s+sales\b", re.IGNORECASE)


FOCUS = [   # (pattern in question, column, label) - most specific first
    (r"margin\s*(%|percent|percentage|pct|ratio)|%\s*margin|margin\s+rate", "MarginPct", "Margin %"),
    (r"\bmargin|\bprofit", "Margin", "Margin"),
    (r"base\s*sales?|basic\s*sales?|\bsales\b|invoice|revenue|billing|turnover", "BaseSalesValue", "Base sales value"),
    (r"take\s*-?\s*home|takehome|in\s*-?hand|\bnet\b", "TakeHome", "Take-home"),
    (r"\bgross\b", "GrossAmount", "Gross"),
    (r"head\s*count|how\s+many\s+employees|no\.?\s+of\s+employees|number\s+of\s+employees|employee\s+count|"
     r"\b(?:most|least|highest|lowest|max\w*|min\w*|more|fewer|less)\s+(?:no\.?\s+of\s+|number\s+of\s+)?employees",
     "Employees", "Employees"),
    (r"\bctc\b|cost\s+to\s+company|\bcost\b|salary|salaries|payroll", "CTC", "CTC"),
]


def focus_metric(question):
    """Which figure the question is about, e.g. 'highest margin amount' -> Margin."""
    for pat, col, label in FOCUS:
        if re.search(pat, question, re.IGNORECASE):
            return col, label
    return "CTC", "CTC"


def pct_basis(question):
    """Margin % on base sales only when a MARGIN % is asked on sales ('gst amount on invoices' is not one)."""
    if SALES_BASIS.search(question) and re.search(r"margin|profit|%|percent|pct", question, re.IGNORECASE):
        return "sales"
    return MARGIN_PCT_DEFAULT


# amounts that are NOT shown unless the question asks for them (each user can change this)
HIDE_GROUPS = {   # name: (words in a question that ask for it, column-name pattern)
    "PF": (r"\bp\.?f\.?\b|provident", r"(^|[_ ])pf(_|$| )|totalpf|total pf|pf_employee|pf_employer|^pf"),
    "Take-home": (r"take\s*-?\s*home|takehome|in\s*-?hand|\bnet\b|net\s*(salary|pay)|cheque",
                  r"take.?home|net.?salary|netpay|cheque"),
    "PT": (r"\bpt\b|professional\s*tax|pftax", r"^pt$|pftax|professional"),
    "ESIC": (r"\besic\b|\besi\b", r"esic"),
    "Gross": (r"\bgross\b", r"gross"),
    "TDS": (r"\btds\b|income\s*-?\s*tax|\btax\b", r"tds|incometax|income tax|income_tax"),
}
DEFAULT_HIDE = list(HIDE_GROUPS)
SHOW_ALL = re.compile(r"\b(all\s+(details|columns|fields|data|info)|everything|full\s+details|complete\s+details|"
                      r"every\s+column|all\s+deductions?)\b", re.IGNORECASE)
PREF_SAVER = None          # set by the app: saves a user's hidden list


def hidden_now():
    try:
        return list(st.session_state.get("hide_cols", DEFAULT_HIDE))
    except Exception:
        return list(DEFAULT_HIDE)


def asked_groups(question):
    return {g for g, (ask, _) in HIDE_GROUPS.items() if re.search(ask, question, re.IGNORECASE)}


def hide_columns(df, question, report=""):
    """Drops PF / take-home / PT / ESIC / gross / TDS columns the user did not ask for."""
    if df is None or df.empty or SHOW_ALL.search(question):
        return df
    if report == "deductions" and re.search(r"\bdeductions?\b|statutory", question, re.IGNORECASE):
        return df
    hide = set(hidden_now()) - asked_groups(question)
    drop = []
    for c in df.columns:
        name = str(c).lower()
        for g in hide:
            if re.search(HIDE_GROUPS[g][1], name) and not (g == "PF" and re.search(r"pftax", name)):
                drop.append(c)
                break
    return df.drop(columns=drop)


PREF_ASK = re.compile(r"\b(from\s+now(\s+on)?|always|never|by\s+default|default|every\s*time|hereafter|aage\s+se)\b|"
                      r"\b(don'?t|dont|do\s+not|stop|no\s+need\s+to|not|hide|remove|start|begin)\s+"
                      r"(to\s+)?(show|showing|put|putting|include|including|display|displaying|give|giving|add|adding)\b",
                      re.IGNORECASE)


def preference_command(question):
    """'from now on don't show pf, tds...' -> saves the setting. Returns a reply dict, or None."""
    groups = asked_groups(question)
    if not groups or not PREF_ASK.search(question) or parse_periods(question):
        return None
    hide_words = re.search(r"\b(don'?t|dont|do\s+not|never|stop|no\s+need|not|hide|remove|without|mat)\b",
                           question, re.IGNORECASE)
    cur = set(hidden_now())
    cur = (cur | groups) if hide_words else (cur - groups)
    try:
        st.session_state.hide_cols = sorted(cur, key=DEFAULT_HIDE.index)
        if PREF_SAVER:
            PREF_SAVER(st.session_state.hide_cols)
    except Exception:
        pass
    names = ", ".join(g for g in DEFAULT_HIDE if g in groups)
    if hide_words:
        text = (f"✅ Got it. From now on I **won't show {names}** unless your question asks for it "
                f"(e.g. *pf and tds for aug-26*).")
    else:
        text = f"✅ Got it. From now on I'll **show {names}** in answers."
    now_hidden = [g for g in DEFAULT_HIDE if g in cur]
    note = ("Hidden unless asked: " + ", ".join(now_hidden)) if now_hidden else "Nothing is hidden."
    return {"text": text, "note": note + " · You can also change this in Settings.", "df": None,
            "report": "pref", "sql": "", "params": [], "intent": {}}


CAT_OFF_SAVER = None       # set by the app: saves a user's hidden categories (Settings → categories)


def category_pref_command(question):
    """'from now on do not show gil data' -> category GIL is left out of every answer for this login;
    'from now on show gil again' -> back. Returns a reply dict, or None."""
    if not PREF_ASK.search(question) or parse_periods(question) or re.search(r"\bonly\b", question, re.I):
        return None
    cats, q = [], question
    while True:                                          # every category named: "gml and efirm"
        c = find_known(q, "Category")
        if not c or c in cats:
            break
        cats.append(c)
        q = re.sub(rf"(?<!\w){re.escape(c)}(?!\w)", " ", q, flags=re.I)
    if not cats:
        return None
    hide = bool(re.search(r"\b(don'?t|dont|do\s+not|never|stop\s+showing|no\s+need|not|hide|remove|exclude|without|mat)\b",
                          question, re.I))
    try:
        cur = list(st.session_state.get("cat_off", []))
    except Exception:
        cur = []
    new = sorted(set(cur) | set(cats)) if hide else [c for c in cur if c not in cats]
    try:
        st.session_state.cat_off = new
    except Exception:
        pass
    if CAT_OFF_SAVER:
        CAT_OFF_SAVER(new)
    names = ", ".join(cats)
    if hide:
        text = (f"✅ Saved. From now on I **won't show {names} data** - category {names} is left out of your totals and "
                f"lists. Name it in a question to see it anyway (e.g. *{cats[0].lower()} ctc for sep-26*), or say "
                f"*from now on show {cats[0].lower()} again*.")
    else:
        text = f"✅ Saved. From now on I'll **show {names} data** again."
    audit_policy_event("preference_changed", json.dumps({"type": "cat_off", "hide": hide, "cats": cats})[:200])
    return {"text": text, "note": "Hidden categories: " + (", ".join(new) or "none") + " · saved for your login only.",
            "df": None, "report": "pref", "sql": "", "params": [], "intent": {}}


REL_MONTH = re.compile(
    r"\b(this|current|latest|recent|present|last|previous|prev|pichhle|pichla|is)\s+(month|mahine|mahina)(?:'?s)?\b|"
    r"\b(this|current|last|previous)\s+(year|fy|financial\s+year)\b|\bytd\b|\byear\s+to\s+date\b|\blast\s+(\d{1,2})\s+months\b",
    re.IGNORECASE)


def _shift(m, y, k):
    n = y * 12 + (m - 1) - k
    return n % 12 + 1, n // 12


QUARTER_RX = re.compile(r"\b(?:q([1-4])|(first|second|third|fourth|1st|2nd|3rd|4th)\s+quarter|quarter\s+([1-4]))\b"
                        r"(?:\s+(?:of\s+)?(?:(this|current|last|previous)\s+)?(?:financial\s+year|fy|year))?"
                        r"(?:\s*(?:fy\s*)?'?(\d{2,4})(?:\s*-\s*'?\d{2,4})?)?", re.I)
_QWORD = {"first": 1, "second": 2, "third": 3, "fourth": 4, "1st": 1, "2nd": 2, "3rd": 3, "4th": 4}


REL_QUARTER_RX = re.compile(r"\b(this|current|latest|last|previous|prev)\s+(quarter|qtr)\b", re.I)
HALF_RX = re.compile(r"\b(?:h([12])|(first|second|1st|2nd)\s+half)\b(?:\s+(?:of\s+)?(?:(this|current|last|previous)\s+)?"
                     r"(?:financial\s+year|fy|year))?", re.I)


def _fq(m, y):
    """(FY start year, quarter 1-4) of a month."""
    fy = y if m >= 4 else y - 1
    return fy, (m - 4) % 12 // 3 + 1


def expand_quarters(question, today=None, latest=None):
    """'q2 of this financial year' -> 'jul-26 to sep-26' (Indian FY: Q1 = Apr-Jun ... Q4 = Jan-Mar).
    'this quarter' / 'last quarter' = the quarter of the LATEST month in the data (latest = (m, y)) / the one before.
    'h1' / 'first half' = Apr-Sep, 'h2' = Oct-Mar."""
    today = today or datetime.date.today()
    cur_fy = today.year if today.month >= 4 else today.year - 1        # FY 2026-27 -> 2026

    def qtext(fy, q):
        sm = 4 + 3 * (q - 1)
        y1 = fy + (sm > 12)
        sm = (sm - 1) % 12 + 1
        em, y2 = (sm + 2, y1) if sm + 2 <= 12 else (sm + 2 - 12, y1 + 1)
        return f"{MON[sm - 1].lower()}-{y1 % 100:02d} to {MON[em - 1].lower()}-{y2 % 100:02d}"

    if latest and REL_QUARTER_RX.search(question):
        def rel(m):
            fy, q = _fq(*latest)
            if m.group(1).lower() in ("last", "previous", "prev"):
                fy, q = (fy - 1, 4) if q == 1 else (fy, q - 1)
            return qtext(fy, q)
        question = REL_QUARTER_RX.sub(rel, question)

    def half(m):
        h = 1 if (m.group(1) == "1" or (m.group(2) or "").lower() in ("first", "1st")) else 2
        fy = cur_fy - 1 if (m.group(3) or "").lower() in ("last", "previous") else cur_fy
        return f"apr-{fy % 100:02d} to sep-{fy % 100:02d}" if h == 1 else f"oct-{fy % 100:02d} to mar-{(fy + 1) % 100:02d}"
    question = HALF_RX.sub(half, question)

    def rep(m):
        q = int(m.group(1) or m.group(3) or _QWORD[m.group(2).lower()])
        fy = cur_fy - 1 if (m.group(4) or "").lower() in ("last", "previous") else cur_fy
        if m.group(5):                                                  # "q2 fy26" / "q2 2026-27" -> FY starting 2026
            y = int(m.group(5))
            fy = y + 2000 if y < 100 else y
        sm = 4 + 3 * (q - 1)
        y1 = fy + (sm > 12)
        sm = (sm - 1) % 12 + 1
        em, y2 = (sm + 2, y1) if sm + 2 <= 12 else (sm + 2 - 12, y1 + 1)
        return f"{MON[sm - 1].lower()}-{y1 % 100:02d} to {MON[em - 1].lower()}-{y2 % 100:02d}"
    return QUARTER_RX.sub(rep, question)


TYPO_VOCAB = ["employee", "employees", "salary", "salaries", "department", "departments", "category", "categories",
              "workorder", "workorders", "margin", "projection", "deduction", "deductions", "average", "highest",
              "lowest", "details", "summary", "compare", "percentage", "joiners", "leavers", "resigned", "between",
              "monthly", "invoice", "headcount", "maximum", "minimum", "employee", "revenue", "billing", "growth",
              "employer", "contribution", "combined", "received", "expired", "arrears", "liable", "payment",
              "forecast", "median", "number", "numbers", "change", "profit", "decrease", "increase"]
# words a "did you mean ...?" may suggest (a typo that is not close enough to fix without asking)
DYM_EXTRA = ["ratio", "share", "paid", "days", "gross", "bonus", "basic", "tds", "pf", "pt", "esic", "gst", "ctc",
             "work", "order", "orders", "month", "months", "salary", "margin", "highest", "lowest", "total", "average",
             "employees", "department", "category", "invoice", "revenue", "billing", "sales", "base", "value", "cost",
             "count", "list", "top", "bottom", "trend", "joined", "left", "expired", "started", "ending", "projection"]


HINGLISH = [(r"\bsabse\s+(?:zyada|jyada|adhik|jada|zada)\b", "highest"), (r"\bsabse\s+kam\b", "lowest"),
            (r"\b(?:kaun\s*sa|kaunsa|kaun\s*si|kaunsi|konsa|konsi|kon\s*sa)\b", "which"),
            (r"\bkitne\s+(?:log|logon|employees?)\b", "how many employees"), (r"\bkitna\b|\bkitni\b", "how much"),
            (r"\bpichhle\s+mahine\b|\bpichle\s+mahine\b|\bpichla\s+mahina\b", "last month"),
            (r"\bis\s+mahine\b|\biss\s+mahine\b", "this month"),
            (r"\b(?:leta|leti|lete|deta|deti|dete|milta|milti|milte|karta|karti|karte)\s+(?:hai|hain|he)\b", "has"),
            (r"\bkam\s+(?:wale|wala|wali)\b", "lowest"), (r"\bzyada\s+(?:wale|wala|wali)\b", "highest")]
SAME_MEANING = [(r"\bprofessional\s+tax\b", "pt"), (r"\bprovident\s+fund\b", "pf"), (r"\bincome\s*tax\b", "tds"),
                (r"\bclients\b|\bcustomers\b", "departments"), (r"\bclient\b|\bcustomer\b", "department"),
                (r"\btake\s*-?\s*home\b", "takehome")]          # words people use for the same column


def fix_typos(question):
    for rx, rep in HINGLISH + SAME_MEANING:             # Hinglish meaning -> English ("sabse zyada" = highest)
        question = re.sub(rx, rep, question, flags=re.IGNORECASE)
    return _fix_spelling(question)


def _fix_spelling(question):
    """'give me employe list', 'salry for sep', 'catagory wise' -> the intended payroll word (only words of 5+ letters
    that are not already known; names / codes / months are left alone)."""
    import difflib
    known = SIMPLE_WORDS | MONTH_WORDS | PAYROLL_TERMS | set(TYPO_VOCAB)

    def rep(m):
        w = m.group(0)
        lw = w.lower()
        if len(lw) < 5 or lw in known or any(ch.isdigit() for ch in lw):
            return w
        hit = difflib.get_close_matches(lw, TYPO_VOCAB, n=1, cutoff=0.84)
        if hit:
            _POL.typo_fixes = getattr(_POL, "typo_fixes", []) + [(w, hit[0])]   # told to the user in the note
        return hit[0] if hit else w
    _POL.typo_fixes = []
    return re.sub(r"[A-Za-z]+", rep, question)


def did_you_mean(question):
    """A typo too far from a known word to fix silently ('revnue', 'hedcount') -> the question with the closest known
    words, to ASK 'did you mean ...?' instead of guessing or waiting for the AI. None if nothing close."""
    import difflib
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category"),
         "employee": find_employee_name(question)}
    vocab = sorted(set(TYPO_VOCAB) | set(DYM_EXTRA) | {w for w in PAYROLL_TERMS if len(w) >= 4})   # payroll words only
    fixed, q2 = [], question
    for w in sorted(unknown_words(question, f), key=len, reverse=True):
        if len(w) < 4:
            continue
        hit = difflib.get_close_matches(w, vocab, n=1, cutoff=0.72)
        if hit and hit[0] != w:
            q2 = re.sub(rf"\b{re.escape(w)}\b", hit[0], q2, flags=re.IGNORECASE)
            fixed.append((w, hit[0]))
    return (q2, fixed) if fixed else None


# ---------------- direct answers for common analyst / HOD / CFO questions (exact SQL, no AI) ----------------
SHARE_RX = re.compile(r"\b(percent(?:age)?|share|portion|contribut\w*|what\s+part)\b|%\s+of\s+(?:the\s+)?total", re.I)


def share_answer(question, mode, notes, t0):
    """'what percentage of total ctc is openvem in aug-26' -> each category's / department's share of the total."""
    if not SHARE_RX.search(question):
        return None
    top = re.search(r"\b(?:top|biggest|largest)\s+(\d{1,3})\s+(work\s*-?\s*orders?|workorders?|wos|departments?|depts?)\b",
                    question, re.IGNORECASE)
    if top:                                              # "share of top 5 work orders in revenue" (concentration)
        return _top_share(question, mode, notes, int(top.group(1)),
                          "WorkOrderNo" if re.match(r"w", top.group(2), re.I) else "Department")
    cat, dep = find_known(question, "Category"), find_known(question, "Department")
    col = "Category" if cat or re.search(r"\bcategor", question, re.I) else \
        "Department" if dep or re.search(r"\bdepartment|\bdept\b|\bclient", question, re.I) else None
    if not col:
        return None
    val = cat if col == "Category" else dep
    metric, mlabel = {"BaseSalesValue": ("BaseSalesValue", "base sales value"), "Margin": ("Margin", "margin"),
                      "Employees": ("Employees", "headcount")}.get(focus_metric(question)[0], ("CTC", "CTC"))
    periods = parse_periods(question) or [latest_period(mode)]
    t1 = time.time()
    sql, params = build_sql(col.lower(), mode, periods, {"department": None, "category": None})
    df = run_query(sql, params)
    if df.empty or metric not in df:
        return None
    g = df.groupby(col)[metric].sum().sort_values(ascending=False)
    total = float(g.sum())
    if not total:
        return None
    out = pd.DataFrame({col: g.index, NICE_METRIC.get(metric, metric): g.values,
                        "Share %": (g.values * 100.0 / total).round(2)})
    ptxt = ", ".join(period_label(m, y) for m, y in periods) + f" ({PERIOD_COL.get(mode, 'Salary month').lower()})"
    fmt = (lambda v: f"{int(v):,}") if metric == "Employees" else inr
    if val and val in g.index:
        head = (f"**{val} is {g[val] * 100 / total:.2f}% of total {mlabel}** for {ptxt} "
                f"({fmt(g[val])} of {fmt(total)}).")
    else:
        head = (f"**Share of total {mlabel} by {col.lower()} for {ptxt}:** largest is **{g.index[0]}** "
                f"at {g.iloc[0] * 100 / total:.2f}% ({fmt(g.iloc[0])} of {fmt(total)}).")
    notes = [f"Understood as: **Share of total {mlabel}** by {col.lower()}."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": head, "note": " ".join(notes), "df": out, "report": "share", "sql": sql, "params": params,
            "intent": {"report": "share", "dim": col, "metric": metric}}


def _top_share(question, mode, notes, n, col):
    """Share of the top N work orders / departments in the month's total (ranked by that same figure)."""
    metric, mlabel = {"BaseSalesValue": ("BaseSalesValue", "base sales value"), "Margin": ("Margin", "margin")}.get(
        focus_metric(question)[0], ("CTC", "CTC"))
    periods = parse_periods(question) or [latest_period(mode)]
    f = {**_scope(question), "sort": "desc", "top": None}
    t1 = time.time()
    sql, params = build_sql("workorder_summary" if col == "WorkOrderNo" else "department", mode, periods, f)
    df = run_query(sql, params)
    if df.empty or metric not in df:
        return None
    g = df.groupby(col)[metric].sum().sort_values(ascending=False)
    total = float(g.sum())
    if not total:
        return None
    out = pd.DataFrame({"Rank": range(1, len(g) + 1), col: g.index, NICE_METRIC.get(metric, metric): g.values,
                        "Share %": (g.values * 100 / total).round(2)})
    out["Cumulative %"] = out["Share %"].cumsum().round(2)
    top_sum = float(g.head(n).sum())
    name = "work orders" if col == "WorkOrderNo" else "departments (clients)"
    text = (f"**Top {n} {name} = {top_sum * 100 / total:.2f}% of total {mlabel}** for {_period_text(periods, mode)} "
            f"({inr(top_sum)} of {inr(total)}, {len(g):,} {name} in all). Largest: **{g.index[0]}** "
            f"({g.iloc[0] * 100 / total:.2f}%).")
    notes = [f"Understood as: **Share of the top {n} {name}** in total {mlabel}."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out.head(max(n, 20)), "report": "share", "sql": sql,
            "params": params, "intent": {"report": "share", "dim": col, "metric": metric}}


NICE_METRIC = {"BaseSalesValue": "Base sales value", "Employees": "Employees", "Margin": "Margin", "CTC": "CTC"}
STATS_RX = re.compile(r"\b(median|percentiles?|quartiles?|distribution|bands?|brackets?|slabs?|ranges?|spread)\b", re.I)
BANDS = [(0, 20000), (20000, 30000), (30000, 40000), (40000, 50000), (50000, 75000), (75000, 100000),
         (100000, 200000), (200000, None)]


def stats_answer(question, mode, notes, t0):
    """'median ctc for sep-26', 'ctc distribution / salary bands for aug-26' (per employee, for the whole month)."""
    if not STATS_RX.search(question) or not re.search(r"\b(ctc|salary|salaries|pay|package|cost)\b", question, re.I):
        return None
    pq = re.search(r"\b(\d{1,2})(?:st|nd|rd|th)?\s+percentile\b", question, re.I)    # "90th percentile ctc"
    if re.search(r"\b(department|category|work\s*order|workorder)[\s\-]*wise\b|\bby\s+(department|category)", question, re.I):
        return None                                     # grouped statistics -> AI analyst
    periods = parse_periods(question) or [latest_period(mode)]
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where(mode, periods, f)
    sql = f"SELECT [Employee], SUM([CTC]) AS [CTC] FROM {T} {where} GROUP BY [Employee]"
    t1 = time.time()
    per = run_query(sql, params)
    c = pd.to_numeric(per["CTC"], errors="coerce").dropna()
    c = c[c > 0]
    if c.empty:
        return None
    ptxt = ", ".join(period_label(m, y) for m, y in periods) + f" ({PERIOD_COL.get(mode, 'Salary month').lower()})"
    scope = " · ".join(v for v in (f["department"], f["category"]) if v)
    scope = f" — {scope}" if scope else ""
    if re.search(r"\b(distribution|bands?|brackets?|slabs?|ranges?|spread)\b", question, re.I):
        rows = []
        for lo, hi in BANDS:
            sel = c[(c >= lo) & ((c < hi) if hi else True)]
            label = (f"₹{lo // 1000}k – ₹{hi // 1000}k" if hi and hi < 100000 else
                     f"₹{lo / 100000:g}L – ₹{hi / 100000:g}L" if hi else f"₹{lo / 100000:g}L and above")
            if lo == 75000 and hi == 100000:
                label = "₹75k – ₹1L"
            rows.append({"CTC band": label, "Employees": int(len(sel)), "Share of employees %": round(len(sel) * 100 / len(c), 2),
                         "CTC": float(sel.sum())})
        out = pd.DataFrame(rows)
        top = out.loc[out["Employees"].idxmax()]
        text = (f"**CTC bands for {ptxt}{scope}:** {len(c):,} employees · most are in **{top['CTC band']}** "
                f"({int(top['Employees']):,} people, {top['Share of employees %']:.1f}%) · median {inr(c.median())}.")
        report = "bands"
    else:
        out = pd.DataFrame([{"Employees": int(len(c)), "Average CTC": round(c.mean(), 2), "Median CTC": round(c.median(), 2),
                             "Lowest CTC": float(c.min()), "25th percentile": round(c.quantile(.25), 2),
                             "75th percentile": round(c.quantile(.75), 2), "Highest CTC": float(c.max())}])
        text = (f"**Median CTC for {ptxt}{scope}: {inr(c.median())}** · average {inr(c.mean())} · lowest {inr(c.min())} "
                f"· highest {inr(c.max())} · across {len(c):,} employees (CTC per employee for the month).")
        if pq and 1 <= int(pq.group(1)) <= 99:
            p_ = int(pq.group(1))
            v_ = c.quantile(p_ / 100)
            out.insert(len(out.columns) - 1, f"{p_}th percentile", round(v_, 2))
            text = (f"**{p_}th percentile CTC for {ptxt}{scope}: {inr(v_)}** — {p_}% of the {len(c):,} employees earn "
                    f"this or less. Median {inr(c.median())} · average {inr(c.mean())}.")
        report = "stats"
    notes = [f"Understood as: **CTC statistics** per employee."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": report, "sql": sql, "params": params,
            "intent": {"report": report}}


AVG_WO_MARGIN = re.compile(r"\baverage\s+margin\s*(?:%|percent\w*|pct)?\s+(?:per|by|of|for|across)\s+(?:each\s+|every\s+|all\s+)?"
                           r"(?:work\s*-?\s*orders?|workorders?|wos?)\b", re.I)


def avg_wo_margin_answer(question, mode, notes, t0):
    """'average margin % per work order in aug-26' -> the simple average over work orders AND the overall margin %."""
    if not AVG_WO_MARGIN.search(question):
        return None
    periods = parse_periods(question) or [latest_period(mode)]
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category"),
         "sort": "desc", "top": None}
    sql, params = build_sql("workorder_summary", mode, periods, f)
    t1 = time.time()
    df = run_query(sql, params)
    if df.empty or "MarginPct" not in df:
        return None
    s = pd.to_numeric(df["MarginPct"], errors="coerce").dropna()
    overall = df["Margin"].sum() * 100 / df["CTC"].sum() if df["CTC"].sum() else float("nan")
    ptxt = ", ".join(period_label(m, y) for m, y in periods) + f" ({PERIOD_COL.get(mode, 'Salary month').lower()})"
    df = df.sort_values("MarginPct", ascending=False).reset_index(drop=True)
    text = (f"**Average margin % per work order for {ptxt}: {s.mean():.2f}%** (simple average of {len(s):,} work orders) · "
            f"overall margin % on total CTC: **{overall:.2f}%** · median {s.median():.2f}% · best {df.iloc[0]['WorkOrderNo']} "
            f"({df.iloc[0]['MarginPct']:.2f}%) · worst {df.iloc[-1]['WorkOrderNo']} ({df.iloc[-1]['MarginPct']:.2f}%).")
    notes = ["Understood as: **Average margin % per work order**."] + notes + [
        "Simple average = every work order counts the same; overall = total margin ÷ total CTC (bigger work orders weigh more).",
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "workorder_summary", "sql": sql, "params": params,
            "intent": {"report": "workorder_summary", "focus": "MarginPct", "focus_label": "Margin %", "sort": "desc"}}


WO_GROUP_WISE = re.compile(r"\b(?:work\s*-?\s*orders?|workorders?|wo)[\s\-]*groups?\b|\bgroup[\s\-]*wise\b|"
                           r"\b(?:by|per|each|every)\s+groups?\b|\bgroups?\s+(?:vs\.?|versus|compar\w*)\b", re.I)


def wo_group_scope(question):
    """Saved work-order groups in a question: 'NORTH margin for sep-26' -> only NORTH; 'SOUTH vs NORTH CTC' /
    'work order group wise base sales' -> one row per group. None if the question has none of it."""
    try:
        groups = WO_GROUPS.load()
    except Exception:
        return None
    if not groups:
        return None
    names, rest = wg.find_names(question, groups)
    wise = bool(WO_GROUP_WISE.search(question)) or len(names) >= 2
    if not names and not wise:
        return None
    if not rest or not re.search(r"[a-z]{2}", rest, re.I):
        rest = "summary"
    return {"groups": groups, "names": names, "wise": wise, "question": rest}


def wo_group_answer(question, scope, mode, notes, t0):
    """One row per work-order group (and month): 'SOUTH vs NORTH CTC', 'work order group wise base sales'."""
    case = wg.case_sql(scope["groups"])
    Y, M = period_exprs(mode)
    periods = parse_periods(question) or [latest_period(mode)]
    where, params = build_where(mode, periods, _scope(question))       # also keeps only the groups named, if any
    sql = (f"SELECT {Y} AS [Year], {M} AS [Month], {case} AS [WorkOrderGroup], {TOTALS} FROM {base_rows(Y, M, where)} "
           f"GROUP BY {Y}, {M}, {case} ORDER BY [Year], [Month], CTC DESC")
    t1 = time.time()
    df = run_query(sql, params)
    metric, mlabel = focus_metric(question)
    if metric not in df.columns:
        metric, mlabel = "CTC", "CTC"
    ptxt = _period_text(periods, mode)
    if df.empty:
        text = f"No rows for the work-order groups in {ptxt}."
    else:
        g = df.groupby("WorkOrderGroup")[metric].sum().sort_values(ascending=False)
        if metric == "MarginPct":                          # a % can't be added up -> from the totals
            tot = df.groupby("WorkOrderGroup")[["Margin", "CTC"]].sum()
            g = (tot["Margin"] * 100 / tot["CTC"].where(tot["CTC"] != 0)).sort_values(ascending=False)
        text = (f"**{mlabel} by work-order group for {ptxt}:** " +
                " · ".join(f"**{k}** {_fmt_metric(metric, v)}" for k, v in g.items()) + ".")
    rules = "; ".join(f"{g_['name']} = starts with {' / '.join(g_['prefixes'])}" for g_ in scope["groups"])
    notes = [f"Understood as: **Work-order group wise {mlabel}** ({rules}; anything else = Other)."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "wo_group", "sql": sql, "params": params,
            "intent": {"report": "wo_group", "metric": metric}}


def _period_text(periods, mode):
    return ", ".join(period_label(m, y) for m, y in periods) + f" ({PERIOD_COL.get(mode, 'Salary month').lower()})"


def _prev_period(m, y):
    return (12, y - 1) if m == 1 else (m - 1, y)


def _scope(question):
    return {"department": find_known(question, "Department"), "category": find_known(question, "Category")}


def _month_row(mode, period, f):
    """The payroll summary row (CTC, base sales, margin, invoice, gross, take-home, headcount) of ONE month."""
    sql, params = build_sql("summary", mode, [period], f)
    df = run_query(sql, params)
    return (None if df.empty else df.iloc[0]), sql, params


def _two_periods(question, mode):
    """(earlier, later) month for 'from aug-26 to sep-26' / 'sep-26 vs aug-26' / 'in sep-26' (= vs month before)."""
    ps = parse_periods(question)
    if len(ps) >= 2:
        return ps[0], ps[-1]
    b = ps[0] if ps else latest_period(mode)
    return _prev_period(*b), b


def _fmt_metric(col, v):
    if v is None or pd.isna(v):
        return "–"
    return f"{int(round(v)):,}" if col == "Employees" else f"{v:.2f}%" if col == "MarginPct" else inr(v)


TOTAL_METRICS = [   # (words, summary column, label) - margin % before margin, the first match per words wins
    (r"head\s*-?\s*count|\bemployees?\b|\bstaff\b|\bpeople\b|manpower|\bhc\b", "Employees", "Headcount"),
    (r"margin\s*(?:%|percent\w*|pct)|%\s*margin", "MarginPct", "Margin %"),
    (r"\bmargin|\bprofit", "Margin", "Margin"),
    (r"base\s*sales?|\bsales\b|revenue|turnover|billing|invoice", "BaseSalesValue", "Base sales value"),
    (r"\bgross\b", "GrossAmount", "Gross"),
    (r"takehome|\bnet\b", "TakeHome", "Take-home"),
    (r"\bctc\b|\bcost\b|salary|salaries|payroll", "CTC", "CTC"),
]


def _asked_metrics(question):
    out, q = [], question
    for rx, col, lab in TOTAL_METRICS:
        if re.search(rx, q, re.IGNORECASE):
            out.append((col, lab))
            q = re.sub(rx, " ", q, flags=re.IGNORECASE)
    return out


CHANGE_RX = re.compile(r"\b(?:growth|grow[sn]?|grew|chang(?:e|ed|es)|difference|diff|variance|increas(?:e|ed|es)|"
                       r"decreas(?:e|ed|es)|went\s+(?:up|down)|go(?:ne|es)?\s+(?:up|down)|rose|rise[sn]?|fell|fall(?:en)?|"
                       r"dropped|drop|declined?|badh\w*|ghat\w*)\b", re.IGNORECASE)
WHY_RX = re.compile(r"\bwhy\b|\breasons?\b|\bwhat\s+(?:caused|drove|led\s+to)\b|\bdrivers?\b|\bkyu+n?\b|\bkyon\b", re.I)
DIM_RX = re.compile(r"work\s*-?\s*orders?|workorders?|\bwos?\b|departments?|\bdepts?\b|categor|\bwise\b|\beach\b|"
                    r"\bwhich\b|\btop\s+\d|\bemployee\s+(?:name|code)|\blist\b", re.IGNORECASE)


def change_answer(question, mode, notes, t0):
    """'ctc growth from aug-26 to sep-26 in percent', 'headcount change between aug and sep', 'did margin go up in
    sep-26' -> the change of the month's totals (amount and %), no AI."""
    if not CHANGE_RX.search(question) or WHY_RX.search(question) or DIM_RX.search(question) \
            or wants_projection(question) or WHATIF_RX.search(question):
        return None
    asked = _asked_metrics(question)
    if not asked:
        return None
    (ma, ya), (mb, yb) = _two_periods(question, mode)
    f = _scope(question)
    t1 = time.time()
    ra, sql, params = _month_row(mode, (ma, ya), f)
    rb, _, _ = _month_row(mode, (mb, yb), f)
    la, lb = period_label(ma, ya), period_label(mb, yb)
    if ra is None or rb is None:
        return clarify(question, f"There is no data for **{la if ra is None else lb}** to compare with. "
                                 "Try two months that have data, e.g.:", [f"ctc growth from {la.lower()} to {lb.lower()}"])
    rows, lines = [], []
    for col, lab in asked:
        a, b = float(ra[col]), float(rb[col])
        d = b - a
        pct = d * 100 / a if a else None
        rows.append({"Figure": lab, la: a, lb: b, "Change": d, "Change %": None if col == "MarginPct" else
                     (round(pct, 2) if pct is not None else None)})
        word = "went up" if d > 0 else "went down" if d < 0 else "did not change"
        if col == "MarginPct":
            amt = f"{abs(d):.2f} percentage points ({a:.2f}% → {b:.2f}%)"
        else:
            amt = (f"by {_fmt_metric(col, abs(d))}" + (f" ({pct:+.2f}%)" if pct is not None else "")
                   + f" — {_fmt_metric(col, a)} → {_fmt_metric(col, b)}")
        lines.append(f"**{lab} {word}** {amt}" if d else f"**{lab} did not change** ({_fmt_metric(col, a)})")
    scope = " · ".join(v for v in (f["department"], f["category"]) if v)
    text = (f"{lines[0]} from {la} to {lb} ({PERIOD_COL.get(mode, 'Salary month').lower()})"
            + (f" — {scope}" if scope else "") + "."
            + ("".join(f"\n- {x}" for x in lines[1:]) if len(lines) > 1 else ""))
    notes = ["Understood as: **Change between two months** (month totals)."] + notes + [
        f"Filtered by **{mode_label(mode)}**. Margin % = margin ÷ CTC.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": pd.DataFrame(rows), "report": "change", "sql": sql,
            "params": params, "intent": {"report": "change"}}


def _wo_month(mode, period, f):
    sql, params = build_sql("workorder_summary", mode, [period], {**f, "sort": "desc", "top": None})
    return run_query(sql, params).set_index("WorkOrderNo"), sql, params


def why_answer_wanted(question):
    """'why did margin go up', 'reason for ctc drop' (a why about margin / CTC / base sales, not about anything else)."""
    return bool(WHY_RX.search(question) and any(c in ("Margin", "MarginPct", "CTC", "BaseSalesValue")
                                                for c, _ in _asked_metrics(question)))


def why_answer(question, mode, notes, t0):
    """'why did margin increase in sep-26', 'reason for ctc drop' -> month vs the month before, split into new work
    orders, work orders that stopped and work orders in both months (biggest movers). Exact, no AI."""
    if not why_answer_wanted(question):
        return None
    asked = [(c, l) for c, l in _asked_metrics(question) if c in ("Margin", "MarginPct", "CTC", "BaseSalesValue")]
    col, lab = asked[0] if asked else ("Margin", "Margin")
    col, lab = ("Margin", "Margin") if col == "MarginPct" else (col, lab)
    (ma, ya), (mb, yb) = _two_periods(question, mode)
    la, lb = period_label(ma, ya), period_label(mb, yb)
    f = _scope(question)
    t1 = time.time()
    wa, _, _ = _wo_month(mode, (ma, ya), f)
    wb, sql, params = _wo_month(mode, (mb, yb), f)
    if wa.empty or wb.empty:
        return clarify(question, f"There is no data for **{la if wa.empty else lb}** to explain the change against.")
    j = wa[[col]].join(wb[[col]], how="outer", lsuffix="_a", rsuffix="_b")
    j["Status"] = ["Both months" if pd.notna(a) and pd.notna(b) else f"New in {lb}" if pd.notna(b) else f"Stopped after {la}"
                   for a, b in zip(j[f"{col}_a"], j[f"{col}_b"])]
    j["Change"] = j[f"{col}_b"].fillna(0) - j[f"{col}_a"].fillna(0)
    ra, rb = _month_row(mode, (ma, ya), f)[0], _month_row(mode, (mb, yb), f)[0]   # same totals as the summary answer
    ta, tb = float(ra[col]), float(rb[col])
    bs = (float(rb["BaseSalesValue"] - ra["BaseSalesValue"]), float(rb["CTC"] - ra["CTC"]))
    d = tb - ta
    new, stop, both = (j[j["Status"] == s] for s in (f"New in {lb}", f"Stopped after {la}", "Both months"))
    up, dn = both["Change"].idxmax(), both["Change"].idxmin()
    word = "went up" if d > 0 else "went down"
    text = (f"**{lab} {word} by {inr(abs(d))}** ({d * 100 / ta:+.2f}%) from {la} to {lb} "
            f"({PERIOD_COL.get(mode, 'Salary month').lower()}): {inr(ta)} → {inr(tb)}.")
    if col == "Margin":
        pa, pb = ta * 100 / float(ra["CTC"]), tb * 100 / float(rb["CTC"])
        text += (f" Base sales value {'rose' if bs[0] >= 0 else 'fell'} {inr(abs(bs[0]))} while CTC "
                 f"{'rose' if bs[1] >= 0 else 'fell'} {inr(abs(bs[1]))}; margin % moved {pa:.2f}% → {pb:.2f}%.")
    text += ("\n\n**Where the change came from:**"
             f"\n- **{len(new)} new work order(s)** (in {lb}, not in {la}): {inr(new['Change'].sum())}"
             f"\n- **{len(stop)} work order(s) that stopped** (in {la}, not in {lb}): {inr(stop['Change'].sum())}"
             f"\n- **{len(both)} work order(s) in both months**: {inr(both['Change'].sum())}"
             + (f" — biggest rise **{up}** ({inr(both.loc[up, 'Change'])}), biggest fall **{dn}** "
                f"({inr(both.loc[dn, 'Change'])})" if len(both) else ""))
    out = j.reset_index().rename(columns={"index": "WorkOrderNo", f"{col}_a": f"{lab} {la}", f"{col}_b": f"{lab} {lb}"})
    out = out.reindex(out["Change"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    notes = [f"Understood as: **Why {lab.lower()} changed** — {lb} vs {la}, work order by work order."] + notes + [
        f"Filtered by **{mode_label(mode)}**. The three parts add up to the total change.",
        f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": "why", "sql": sql, "params": params,
            "intent": {"report": "why"}}


RATIO_RX = re.compile(r"\b(cost|ctc|salary|salaries|payroll|revenue|sales|base\s*sales?|billing|turnover)\s*"
                      r"(?:-|to|/|vs\.?|versus|as\s+(?:a\s+)?(?:%|percent\w*|share)\s+of|ratio\s+(?:to|with)|of)\s*"
                      r"(revenue|sales|base\s*sales?|billing|turnover|cost|ctc|salary|payroll)\b|\bcost\s+ratio\b", re.I)


def ratio_answer(question, mode, notes, t0):
    """'cost to revenue ratio for sep-26', 'ctc as % of revenue', 'revenue to cost' (one row per month)."""
    m = RATIO_RX.search(question)
    if not m:
        return None
    cost_w = r"cost|ctc|salar|payroll"
    a, b = (m.group(1) or "cost"), (m.group(2) or "revenue")
    if bool(re.match(cost_w, a, re.I)) == bool(re.match(cost_w, b, re.I)):
        return None                                      # "cost of ctc" / "sales to billing": not a cost/revenue ratio
    rev_first = not re.match(cost_w, a, re.I)
    periods = parse_periods(question) or [latest_period(mode)]
    f = _scope(question)
    t1 = time.time()
    rows, sql, params = [], "", []
    for p in periods:
        r, sql, params = _month_row(mode, p, f)
        if r is None or not r["BaseSalesValue"] or not r["CTC"]:
            continue
        rows.append({"Month": period_label(*p), "CTC": float(r["CTC"]), "BaseSalesValue": float(r["BaseSalesValue"]),
                     "Cost to revenue %": round(r["CTC"] * 100 / r["BaseSalesValue"], 2),
                     "Revenue per ₹100 CTC": round(r["BaseSalesValue"] * 100 / r["CTC"], 2),
                     "Margin % on revenue": round((r["BaseSalesValue"] - r["CTC"]) * 100 / r["BaseSalesValue"], 2)})
    if not rows:
        return None
    r = rows[-1]
    ptxt = f"{r['Month']} ({PERIOD_COL.get(mode, 'Salary month').lower()})"
    if rev_first:
        text = (f"**Revenue to cost for {ptxt}: {r['Revenue per ₹100 CTC'] / 100:.2f}×** — every ₹100 of CTC brings "
                f"₹{r['Revenue per ₹100 CTC']:.2f} of base sales value ({inr(r['BaseSalesValue'])} ÷ {inr(r['CTC'])}).")
    else:
        text = (f"**Cost to revenue ratio for {ptxt}: {r['Cost to revenue %']:.2f}%** — CTC {inr(r['CTC'])} is "
                f"{r['Cost to revenue %']:.2f}% of base sales value {inr(r['BaseSalesValue'])}. Out of every ₹100 of "
                f"revenue, ₹{r['Cost to revenue %']:.2f} is salary cost and ₹{r['Margin % on revenue']:.2f} is margin.")
    if len(rows) > 1:
        text += "\n" + "".join(f"\n- {x['Month']}: {x['Cost to revenue %']:.2f}%" for x in rows)
    notes = ["Understood as: **Cost to revenue ratio** = CTC ÷ base sales value (invoice × 100/118)."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": pd.DataFrame(rows), "report": "ratio", "sql": sql,
            "params": params, "intent": {"report": "ratio"}}


WHATIF_RX = re.compile(r"\b(?:what\s+if|if|suppose|assum\w*|agar|scenario|in\s+case)\b", re.IGNORECASE)
_WI_METRIC = r"(ctc|salary\s+cost|salaries|salary|cost|payroll|revenue|base\s*sales?|sales|billing|invoice|rates?)"
_WI_DIR = (r"(increas\w*|rise[sn]?|rises|rising|hike[sd]?|goes\s+up|go\s+up|grows?|grew|raised?|badh\w*|up|"
           r"decreas\w*|reduc\w*|cut|cuts|drops?|dropped|falls?|goes\s+down|go\s+down|ghat\w*|down)")
_WI_NUM = r"(?:rs\.?|₹)?\s*(\d[\d,]*(?:\.\d+)?)\s*(%|percent|pct|lakhs?|lacs?|crores?|cr\b|k\b|thousand)?"


def whatif_answer(question, mode, notes, t0):
    """'if ctc increases by 10% what will be margin for sep-26', 'what if revenue drops 5 lakh' -> the month's margin
    after the change, and how far CTC can rise before margin turns negative. Exact, no AI."""
    if not (WHATIF_RX.search(question) or re.search(r"\bwhat\s+(?:will|would|happens?)\b|\bimpact\b|\beffect\b", question, re.I)):
        return None
    changes = []
    for rx, order in ((rf"{_WI_METRIC}\s+(?:\w+\s+){{0,3}}?{_WI_DIR}\s+(?:by\s+)?{_WI_NUM}", "mdn"),
                      (rf"{_WI_NUM}\s+{_WI_DIR}\s+(?:in|of|on)\s+(?:the\s+)?{_WI_METRIC}", "ndm"),
                      (rf"{_WI_DIR}\s+(?:the\s+)?{_WI_METRIC}\s+(?:by\s+)?{_WI_NUM}", "dmn")):
        for m in re.finditer(rx, question, re.IGNORECASE):
            g = m.groups()
            met, dr, num, unit = ((g[0], g[1], g[2], g[3]) if order == "mdn" else (g[3], g[2], g[0], g[1])
                                  if order == "ndm" else (g[1], g[0], g[2], g[3]))
            col = "CTC" if re.match(r"ctc|salar|cost|payroll", met, re.I) else "BaseSalesValue"
            sign = -1 if re.match(r"decreas|reduc|cut|drop|fall|goes\s+down|go\s+down|ghat|down", dr, re.I) else 1
            if col not in [c[0] for c in changes]:
                changes.append((col, sign, num, unit))
    if not changes:
        return None
    periods = parse_periods(question) or [latest_period(mode)]
    p = periods[-1]
    f = _scope(question)
    t1 = time.time()
    r, sql, params = _month_row(mode, p, f)
    if r is None:
        return None
    now = {"CTC": float(r["CTC"]), "BaseSalesValue": float(r["BaseSalesValue"])}
    new, said, read_as = dict(now), [], []
    for col, sign, num, unit in changes:
        v = float(num.replace(",", ""))
        if unit and unit.lower() in ("%", "percent", "pct"):
            new[col] = now[col] * (1 + sign * v / 100)
            said.append(f"{'CTC' if col == 'CTC' else 'base sales value'} {'+' if sign > 0 else '−'}{v:g}%")
        elif unit or v > 100:
            amt = _num(num, unit)
            new[col] = now[col] + sign * amt
            said.append(f"{'CTC' if col == 'CTC' else 'base sales value'} {'+' if sign > 0 else '−'}{inr(amt)}")
        else:                                            # "increase by 10" -> 10 %
            new[col] = now[col] * (1 + sign * v / 100)
            said.append(f"{'CTC' if col == 'CTC' else 'base sales value'} {'+' if sign > 0 else '−'}{v:g}%")
            read_as.append(f"Read “{v:g}” as {v:g}%.")
    m0, m1 = now["BaseSalesValue"] - now["CTC"], new["BaseSalesValue"] - new["CTC"]
    pct0, pct1 = m0 * 100 / now["CTC"], (m1 * 100 / new["CTC"] if new["CTC"] else float("nan"))
    ptxt = f"{period_label(*p)} ({PERIOD_COL.get(mode, 'Salary month').lower()})"
    text = (f"**If {' and '.join(said)}, margin for {ptxt} would be {inr(m1)} ({pct1:.2f}% on CTC)** — "
            f"{'down' if m1 < m0 else 'up'} {inr(abs(m1 - m0))} from today's {inr(m0)} ({pct0:.2f}%).\n\n"
            f"- CTC {inr(now['CTC'])} → {inr(new['CTC'])}\n- Base sales value {inr(now['BaseSalesValue'])} → "
            f"{inr(new['BaseSalesValue'])}\n- Break-even: CTC can rise **{m0 * 100 / now['CTC']:.2f}%** "
            f"(or base sales can fall **{m0 * 100 / now['BaseSalesValue']:.2f}%**) before margin becomes zero.")
    out = pd.DataFrame([{"Item": "CTC", "Now": now["CTC"], "After change": new["CTC"]},
                        {"Item": "Base sales value", "Now": now["BaseSalesValue"], "After change": new["BaseSalesValue"]},
                        {"Item": "Margin", "Now": m0, "After change": m1}])
    out["Change"] = out["After change"] - out["Now"]
    notes = ["Understood as: **What-if** on the month's totals (everything else unchanged)."] + read_as + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": "whatif", "sql": sql, "params": params,
            "intent": {"report": "whatif"}}


COUNT_DIM_RX = re.compile(r"\b(?:how\s+many|number\s+of|no\.?\s+of|count\s+of|total(?:\s+no\.?\s+of|\s+number\s+of)?|"
                          r"kitne|kitni)\s+(?:active\s+|running\s+|live\s+)?(work\s*-?\s*orders?|workorders?|wos|"
                          r"departments?|depts|categor(?:y|ies))\b|\b(work\s*-?\s*orders?|workorders?|departments?)\s+count\b",
                          re.IGNORECASE)
COUNT_SKIP = re.compile(r"margin|profit|loss|negative|\b(?:above|below|over|under)\b|more\s+than|less\s+than|greater|[<>=]|"
                        r"\bwith\b|\bhaving\b|\bwithout\b|\bwhere\b|\bwise\b|\bper\b|\beach\b|\bctc\s+(?:above|below)", re.I)


def count_answer(question, mode, notes, t0):
    """'how many work orders in sep-26', 'number of clients in aug-26' -> distinct work orders / departments
    (and categories, employees) that have salary in the month, split by category. No AI."""
    m = COUNT_DIM_RX.search(question)
    if not m or COUNT_SKIP.search(question) or NEW_WO_RX.search(question) or EXPIRED_RX.search(question) \
            or WO_END_RX.search(question):
        return None
    what = (m.group(1) or m.group(2) or "").lower()
    col = "Department" if what.startswith("dep") else "Category" if what.startswith("categ") else "WorkOrderNo"
    periods = parse_periods(question) or [latest_period(mode)]
    f = _scope(question)
    Y, M = period_exprs(mode)
    where, params = build_where(mode, periods, f)
    sql = f"""SELECT {Y} AS [Year], {M} AS [Month], [Category], COUNT(DISTINCT [WorkOrderNo]) AS [Work orders],
       COUNT(DISTINCT [Department]) AS [Departments], COUNT(DISTINCT [Employee]) AS [Employees], SUM([CTC]) AS [CTC]
FROM {T} {where} GROUP BY {Y}, {M}, [Category] ORDER BY {Y}, {M}, [Work orders] DESC"""
    tot_sql = f"""SELECT {Y} AS [Year], {M} AS [Month], COUNT(DISTINCT [WorkOrderNo]) AS [Work orders],
       COUNT(DISTINCT [Department]) AS [Departments], COUNT(DISTINCT [Category]) AS [Categories],
       COUNT(DISTINCT [Employee]) AS [Employees] FROM {T} {where} GROUP BY {Y}, {M} ORDER BY {Y}, {M}"""
    t1 = time.time()
    df, tot = run_query(sql, params), run_query(tot_sql, params)
    if tot.empty:
        return None
    key = {"WorkOrderNo": "Work orders", "Department": "Departments", "Category": "Categories"}[col]
    name = {"WorkOrderNo": "work orders", "Department": "departments (clients)", "Category": "categories"}[col]
    mt = PERIOD_COL.get(mode, "Salary month").lower()
    parts = [f"**{int(r[key]):,} {name}** in {period_label(r['Month'], r['Year'])} ({mt}) · "
             f"{int(r['Work orders']):,} work orders · {int(r['Departments']):,} departments · "
             f"{int(r['Employees']):,} employees" for _, r in tot.iterrows()]
    text = parts[0] if len(parts) == 1 else "\n".join("- " + p_ for p_ in parts)
    df.insert(0, "Period", [period_label(mm, yy) for mm, yy in zip(df["Month"], df["Year"])])
    df = df.drop(columns=["Year", "Month"])
    notes = [f"Understood as: **Count of {name}** that have salary in the month (by category below)."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "count", "sql": tot_sql, "params": params,
            "intent": {"report": "count"}}


PAID_DAYS_RX = re.compile(r"\b(?:paid|working|payable|present|attendance)\s*-?\s*days?\b|\bdays?\s+paid\b|\bpaiddays\b|"
                          r"\b(?:lop|loss\s+of\s+pay|leave\s+without\s+pay)\b", re.IGNORECASE)


def paid_days_answer(question, mode, notes, t0):
    """'total / average paid days for sep-26', 'employees with less than 30 paid days', 'who has loss of pay'.
    Paid days are per salary record (one employee can have 2 salary months in one liable month)."""
    if not PAID_DAYS_RX.search(question):
        return None
    periods = parse_periods(question) or [latest_period(mode)]
    f = _scope(question)
    where, params = build_where(mode, periods, f)
    sql = (f"SELECT [Employee], [Name], [Department], [WorkOrderNo], [Year], [Month], [PaidDays] FROM {T} {where} "
           f"ORDER BY [PaidDays], [Name]")
    t1 = time.time()
    df = run_query(sql, params)
    if df.empty:
        return None
    pdays = pd.to_numeric(df["PaidDays"], errors="coerce")
    ptxt = _period_text(periods, mode)
    cond = re.search(r"(less\s+than|fewer\s+than|below|under|<|more\s+than|above|over|>|equal\s+to|exactly|=)\s*(\d{1,2})"
                     r"\s*(?:paid\s+)?days?", question, re.IGNORECASE)
    lop = re.search(r"\b(lop|loss\s+of\s+pay|leave\s+without\s+pay|not\s+full|partial|less\s+than\s+full|absent)\b",
                    question, re.IGNORECASE)
    zero = re.search(r"\b(zero|no|0)\s+paid\s+days?\b", question, re.IGNORECASE)
    if cond or lop or zero:
        if lop:
            full = [_month_end(int(y), int(m)).day for y, m in zip(df["Year"], df["Month"])]
            sel, how = pdays < pd.Series(full, index=df.index), "fewer paid days than days in their salary month"
        elif zero:
            sel, how = pdays == 0, "0 paid days"
        else:
            op, n = cond.group(1).lower(), int(cond.group(2))
            if re.match(r"less|fewer|below|under|<", op):
                sel, how = pdays < n, f"less than {n} paid days"
            elif re.match(r"more|above|over|>", op):
                sel, how = pdays > n, f"more than {n} paid days"
            else:
                sel, how = pdays == n, f"exactly {n} paid days"
        out = df[sel].copy()
        out.insert(4, "Salary month", [period_label(m, y) for m, y in zip(out["Month"], out["Year"])])
        out = out.drop(columns=["Year", "Month"]).reset_index(drop=True)
        text = (f"**{out['Employee'].nunique():,} employee(s) have {how}** in {ptxt} "
                f"({len(out):,} salary record(s) of {len(df):,}).")
    else:
        out = pd.DataFrame([{"Total paid days": int(pdays.sum()), "Average paid days": round(pdays.mean(), 2),
                             "Salary records": len(df), "Employees": int(df["Employee"].nunique()),
                             "Lowest": int(pdays.min()), "Highest": int(pdays.max())}])
        avg = bool(re.search(r"\b(average|avg|mean)\b", question, re.IGNORECASE))
        text = (f"**{'Average' if avg else 'Total'} paid days for {ptxt}: "
                f"{f'{pdays.mean():.2f}' if avg else f'{int(pdays.sum()):,}'}**"
                f" · {'total' if avg else 'average'} {f'{int(pdays.sum()):,}' if avg else f'{pdays.mean():.2f}'} · "
                f"{len(df):,} salary records · {df['Employee'].nunique():,} employees.")
    notes = ["Understood as: **Paid days** (per salary record)."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": "paid_days", "sql": sql, "params": params,
            "intent": {"report": "paid_days"}}


EMP_VALUE_COLS = [   # (words, column, label) - per-employee amounts that "employees with zero tds" can ask about
    (r"employer\s+(?:pf|contribution|share)|pf\s+employer", "PF_Employer", "PF (employer)"),
    (r"\btds\b", "IncomeTax", "TDS"),
    (r"\bpf\b", "PF_Employee", "PF (employee)"),
    (r"\bpt\b", "PT", "PT"),
    (r"\besic?\b", "ESIC_Employee", "ESIC (employee)"),
    (r"\bnps\b", "CORP_NPS", "Corporate NPS"),
    (r"\barrears?\b", "Arrears", "Arrears"),
    (r"\bbonus\b", "Bonus", "Bonus"),
    (r"\bgross\b", "GrossAmount", "Gross"),
    (r"takehome", "TakeHome", "Take-home"),
]


def emp_value_answer(question, mode, notes, t0):
    """'how many employees have zero tds in sep-26', 'employees without pf', 'staff with tds above 10000'
    -> each employee's month total of that amount, then the condition. No AI."""
    if not re.search(r"\b(employees?|staff|people|persons?|emps?|who|whom|log)\b", question, re.IGNORECASE):
        return None
    hit = next(((rx, c, lab) for rx, c, lab in EMP_VALUE_COLS if re.search(rx, question, re.IGNORECASE)), None)
    if not hit:
        return None
    rx, col, lab = hit
    q = question
    if re.search(rf"(?:\bzero|\bno|\bnil|\b0)\s+(?:{rx})|without\s+(?:any\s+)?(?:{rx})|(?:{rx})\s+(?:is\s+|=\s*|of\s+)?"
                 rf"(?:zero|nil|0)\b|(?:{rx})\s+(?:not|never)\s+(?:deducted|paid)|not\s+(?:paying|getting|having)\s+(?:{rx})",
                 q, re.IGNORECASE):
        op, val, how = "=", 0.0, f"zero {lab}"
    else:
        m = re.search(rf"(?:{rx})\s*(?:is|of|amount|deducted)?\s*(above|over|more\s+than|greater\s+than|>|below|under|"
                      rf"less\s+than|<)\s*" + NUM_RX, q, re.IGNORECASE)
        if m:
            op = ">" if re.match(r"above|over|more|greater|>", m.group(1), re.I) else "<"
            val = _num(m.group(2), m.group(3))
            how = f"{lab} {'above' if op == '>' else 'below'} {inr(val)}"
        elif re.search(rf"(?:with|having|getting|paying|deducted)\s+(?:{rx})|(?:{rx})\s+(?:is\s+)?deducted", q, re.I):
            op, val, how = ">", 0.0, f"{lab} deducted"
        else:
            return None
    periods = parse_periods(question) or [latest_period(mode)]
    f = _scope(question)
    where, params = build_where(mode, periods, f)
    expr = TAKEHOME_EXPR if col == "TakeHome" else f"ISNULL([{col}], 0)"
    sql = (f"SELECT [Employee], MAX([Name]) AS [Name], MAX([Department]) AS [Department], SUM({expr}) AS [{lab}] "
           f"FROM {T} {where} GROUP BY [Employee]")
    t1 = time.time()
    df = run_query(sql, params)
    if df.empty:
        return None
    s = pd.to_numeric(df[lab], errors="coerce")
    if s.isna().all():
        return None                                      # column hidden for this login
    sel = (s == val) if op == "=" else (s > val) if op == ">" else (s < val)
    out = df[sel].sort_values([lab, "Name"], ascending=[op == "<", True]).reset_index(drop=True)
    text = (f"**{len(out):,} of {len(df):,} employees have {how}** in {_period_text(periods, mode)}"
            + (f" · total {lab} of these: {inr(pd.to_numeric(out[lab]).sum())}" if op != "=" else "") + ".")
    notes = [f"Understood as: **Employees by {lab}** ({lab} added up per employee for the month)."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": "emp_value", "sql": sql, "params": params,
            "intent": {"report": "emp_value"}}


INVOICE_NO_RX = re.compile(r"\binvoice\s*(?:no\.?s?|nos|numbers?|num|#|ids?|list)\b|\bbill\s*(?:no\.?s?|numbers?)\b", re.I)


def invoice_no_answer(question, mode, notes, t0):
    """'invoice numbers for work order M244725124', 'invoice nos for sep-26' -> each invoice once (per work order and
    salary month) with its amount and date."""
    if not INVOICE_NO_RX.search(question):
        return None
    wo, rest = find_workorder(question)
    periods = parse_periods(rest)
    if not periods and not wo:
        periods = [latest_period(mode)]
    f = {**_scope(rest), "workorder": wo if valid_wo(wo) else None}
    where, params = build_where(mode, periods, f)
    sql = f"""SELECT [WorkOrderNo], [InvoiceNo], MIN([Year] * 100 + [Month]) AS [_sm], MAX([month_year]) AS [month_year],
       MAX([InvoiceDate]) AS [Invoice date], MAX([InvoiceAmount]) AS [Invoice amount], COUNT(DISTINCT [Employee]) AS [Employees]
FROM {T} {where} GROUP BY [WorkOrderNo], [InvoiceNo] ORDER BY [WorkOrderNo], MIN([Year] * 100 + [Month]), [InvoiceNo]"""
    t1 = time.time()
    df = run_query(sql, params)
    if df.empty:
        return None
    df.insert(2, "Salary month", [period_label(int(v) % 100, int(v) // 100) for v in df["_sm"]])
    df = add_payment_date(df.drop(columns=["_sm"]))
    nos = [f"{n} ({sm} salary)" for n, sm in zip(df["InvoiceNo"], df["Salary month"]) if pd.notna(n)]
    where_txt = f"work order {f['workorder']}" if f["workorder"] else _period_text(periods, mode)
    text = (f"**{len(nos):,} invoice number(s) for {where_txt}:** " + ", ".join(nos[:12])
            + (f" … and {len(nos) - 12:,} more (see table)" if len(nos) > 12 else "") + ".")
    notes = ["Understood as: **Invoice numbers** (each invoice once, with its salary month)."] + notes + [
        f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "invoice_no", "sql": sql, "params": params,
            "intent": {"report": "invoice_no"}}


BREAKUP_RX = re.compile(r"\b(?:old|older|previous|earlier|past|pending|arrear)\s+(?:salary\s+)?months?\b|"
                        r"\bsalary\s+months?\s+(?:are\s+|is\s+)?(?:in|inside|included|covered|paid|booked|under|within)\b|"
                        r"\bwhich\s+salary\s+months?\b|\b(?:breakup|break\s*-?\s*up|split)\s+(?:of\s+)?(?:the\s+)?"
                        r"(?:liable|payment)\s+month|\bliable\s+month\s+(?:breakup|break\s*-?\s*up|split)", re.IGNORECASE)


def liable_breakup_answer(question, mode, notes, t0):
    """'how many old salary months are paid in liable month sep-26' -> the salary months (Month/Year) inside one
    liable (or payment) month: current month vs older months (arrears), records, people, CTC."""
    if not BREAKUP_RX.search(question):
        return None
    mode = "payment" if month_mode_of(question, None) == "payment" else "month_year"
    periods = parse_periods(question) or [latest_period(mode)]
    f = _scope(question)
    Y, M = period_exprs(mode)
    where, params = build_where(mode, periods, f)
    sql = f"""SELECT {Y} AS [PY], {M} AS [PM], [Year], [Month], COUNT(*) AS [Salary records],
       COUNT(DISTINCT [Employee]) AS [Employees], COUNT(DISTINCT [WorkOrderNo]) AS [Work orders],
       SUM([CTC]) AS [CTC], SUM([_Base]) AS [BaseSalesValue]
FROM {base_rows(Y, M, where)} GROUP BY {Y}, {M}, [Year], [Month] ORDER BY {Y}, {M}, [Year], [Month]"""
    t1 = time.time()
    df = run_query(sql, params)
    if df.empty:
        return None
    lines = []
    for (py, pm), g in df.groupby(["PY", "PM"], sort=True):
        cy, cm = (int(py), int(pm)) if mode == "month_year" else _prev_period(int(pm), int(py))[::-1]
        cur = g[(g["Year"] == cy) & (g["Month"] == cm)]
        old = g[(g["Year"] * 100 + g["Month"]) < cy * 100 + cm]
        lines.append(f"**{PERIOD_COL[mode]} {period_label(pm, py)} holds {len(g)} salary month(s):** "
                     f"{int(cur['Salary records'].sum()):,} record(s) for {period_label(cm, cy)} itself · "
                     f"**{int(old['Salary records'].sum()):,} record(s) from {len(old)} older salary month(s)** "
                     + (f"({', '.join(f'{period_label(m, y)} {int(n)}' for y, m, n in zip(old['Year'], old['Month'], old['Salary records']))})"
                        f" · CTC of older months {inr(old['CTC'].sum())}" if len(old) else ""))
    df.insert(0, PERIOD_COL[mode], [period_label(m, y) for m, y in zip(df["PM"], df["PY"])])
    df.insert(1, "Salary month", [period_label(m, y) for m, y in zip(df["Month"], df["Year"])])
    df = df.drop(columns=["PY", "PM", "Year", "Month"])
    notes = ["Understood as: **Salary months inside a liable month** (older months = arrears / late salary)."] + notes + [
        f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": "\n\n".join(lines), "note": " ".join(notes), "df": df, "report": "breakup", "sql": sql,
            "params": params, "intent": {"report": "breakup"}}


BUT_NOT_RX = re.compile(r"\bbut\s+not\s+(?:in\s+)?|\band\s+not\s+in\b|\bnot\s+in\b|\bmissing\s+in\b|\babsent\s+in\b", re.I)


def in_not_in_answer(question, mode, notes, t0):
    """'which work orders were in aug-26 but not in sep-26', 'employees in sep-26 but not in aug-26', clients too ->
    the list, with what they had in the month they were in. No AI."""
    m = BUT_NOT_RX.search(question)
    if not m:
        return None
    dm = re.search(r"work\s*-?\s*orders?|workorders?|\bwos\b|employees?|\bstaff\b|\bpeople\b|departments?|\bdepts?\b",
                   question, re.IGNORECASE)
    pin, pout = parse_periods(question[:m.start()]), parse_periods(question[m.end():])
    if not dm or len(pin) != 1 or len(pout) != 1:
        return None
    col = ("WorkOrderNo" if re.match(r"w", dm.group(0), re.I) else "Department" if re.match(r"dep", dm.group(0), re.I)
           else "Employee")
    f = _scope(question[:dm.start()] + question[dm.end():])
    t1 = time.time()
    wi, pi = build_where(mode, pin, f)
    wo_, po = build_where(mode, pout, f)
    name = "MAX([Name]) AS [Name], MAX([Department]) AS [Department], " if col == "Employee" else \
        "MAX([Department]) AS [Department], MAX([Category]) AS [Category], " if col == "WorkOrderNo" else ""
    sql = (f"SELECT [{col}], {name}COUNT(DISTINCT [Employee]) AS [Employees], SUM([CTC]) AS [CTC] FROM {T} {wi} "
           f"AND [{col}] IS NOT NULL AND [{col}] NOT IN (SELECT [{col}] FROM {T} {wo_} AND [{col}] IS NOT NULL) "
           f"GROUP BY [{col}] ORDER BY SUM([CTC]) DESC")
    params = pi + po
    df = run_query(sql, params)
    n_people = int(run_query(f"SELECT COUNT(DISTINCT [Employee]) AS n FROM {T} {wi} AND [{col}] NOT IN "
                             f"(SELECT [{col}] FROM {T} {wo_} AND [{col}] IS NOT NULL)", params)["n"].iloc[0])
    la, lb = period_label(*pin[0]), period_label(*pout[0])
    what = {"WorkOrderNo": "work order(s)", "Department": "department(s) (clients)", "Employee": "employee(s)"}[col]
    mt = PERIOD_COL.get(mode, "Salary month").lower()
    text = (f"**{len(df):,} {what} in {la} but not in {lb}** ({mt})"
            + ((f" · they had {n_people:,} employee(s) and CTC {inr(df['CTC'].sum())} in {la}" if col != "Employee"
                else f" · their CTC in {la}: {inr(df['CTC'].sum())}") if len(df) else "")
            + ".")
    if col == "Employee" and "Employees" in df:
        df = df.drop(columns=["Employees"])
    notes = [f"Understood as: **{what.capitalize()} in {la}, not in {lb}**."] + notes + [
        f"Filtered by **{mode_label(mode)}**.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "in_not_in", "sql": sql, "params": params,
            "intent": {"report": "in_not_in"}}


NEW_WO_RX = re.compile(r"\b(?:new|newly|fresh|started|starting|start(?:s|ed)?\s+in|began|begin\w*|commenc\w*|added|"
                       r"receiv\w*|got\s+new|awarded|won|mil[ae]|aaye?|aay[ai]|"
                       r"latest\s+added|naye?|nayi|first\s+time|for\s+the\s+first|"
                       r"start\s*date\s+(?:is|was|in|of|on|=|for))\b", re.I)
FIRST_TIME_RX = re.compile(r"\bfirst\s+time\b|\bfor\s+the\s+first\b", re.I)
START_DATE_RX = re.compile(r"\bstart\s*date\b|\bstarted\s+(?:in|on)\b|\bstart(?:s|ed)?\s+in\b", re.I)


def new_wo_answer(question, mode, notes, t0):
    """'new work orders received in sep-26' (no month type named) = work orders whose START DATE (WOStartDate) is in
    sep-26. 'new work orders in liable month / month_year / salary month / paid in sep-26' = work orders that appear
    in that month type for the FIRST time (in sep-26, never before). 'first time' / 'start date' words always win.
    Their staff, CTC, base sales and margin for the month."""
    if not (WO_WORDS.search(question) or re.search(r"\bworkorders?\b", question, re.I)) or not NEW_WO_RX.search(question):
        return None
    if WO_END_RX.search(question) or re.search(r"\bstart\s*date\s+(?:before|after|between|from|till|until|upto|>|<)",
                                                 question, re.I):
        return None
    named_type = month_mode_of(question, None) is not None   # "liable month" / "month_year" / "salary month" / "paid in"
    only_first = bool(FIRST_TIME_RX.search(question)) or (named_type and not START_DATE_RX.search(question))
    only_started = not only_first
    periods = parse_periods(question) or [latest_period(mode)]
    (m1, y1), (m2, y2) = periods[0], periods[-1]
    d1, d2 = datetime.date(y1, m1, 1), _month_end(y2, m2) + datetime.timedelta(days=1)
    yms = [y * 100 + m for m, y in periods]
    Y, M = period_exprs(mode)
    pk = f"({Y}*100+{M})"                                   # the month in the question's month type
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where(mode, [], f)
    inn = ",".join("?" * len(yms))
    cond = {"started": "(MAX([WOStartDate]) >= ? AND MAX([WOStartDate]) < ?)", "first": f"MIN({pk}) IN ({inn})"}
    having = (cond["first"] if only_first else cond["started"] if only_started
              else f"{cond['started']} OR {cond['first']}")
    hparams = ([] if only_first else [d1, d2]) + ([] if only_started else yms)
    sql = f"""SELECT [WorkOrderNo], MAX([Category]) AS Category, MAX([Department]) AS Department,
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate, MIN({pk}) AS [_first],
       COUNT(DISTINCT CASE WHEN {pk} IN ({inn}) THEN [Employee] END) AS Employees,
       SUM(CASE WHEN {pk} IN ({inn}) THEN [CTC] ELSE 0 END) AS CTC,
       SUM(CASE WHEN {pk} IN ({inn}) THEN [_Base] ELSE 0 END) AS BaseSalesValue
FROM (SELECT *, MAX({BASE_EXPR}) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) * 1.0
            / COUNT(*) OVER (PARTITION BY [WorkOrderNo], [Year], [Month]) AS [_Base] FROM {T} {where}) AS r
GROUP BY [WorkOrderNo]
HAVING {having}
ORDER BY MAX([WOStartDate]), [WorkOrderNo]"""
    params = yms * 3 + list(params) + hparams               # same order as the ? in the SQL text
    t1 = time.time()
    df = run_query(sql, params)
    ptxt = period_label(m1, y1) if (m1, y1) == (m2, y2) else f"{period_label(m1, y1)} to {period_label(m2, y2)}"
    mtxt = PERIOD_COL.get(mode, "Salary month").lower()     # "liable month" / "salary month" / "payment month"
    what = "work orders paid for the first time" if only_first else "work orders received (start date)"
    if df.empty:
        return {"text": f"**No {what} in {ptxt}** ({mtxt}).",
                "note": f"Understood as: **New work orders** · {mode_label(mode)}.", "df": None, "report": "new_wo",
                "sql": sql, "params": params, "intent": {"report": "new_wo"}}
    start = pd.to_datetime(df["WOStartDate"], errors="coerce")
    started = (start >= pd.Timestamp(d1)) & (start < pd.Timestamp(d2))
    first_lbl = {"month_year": "First liable month", "payment": "First payment month"}.get(mode, "First salary month")
    df.insert(1, "Why new", ["Started in " + ptxt if s else
                             f"{first_lbl} {period_label(int(fm) % 100, int(fm) // 100)} (started "
                             f"{st.strftime('%d-%b-%Y') if pd.notna(st) else '?'})"
                             for s, fm, st in zip(started, df["_first"], start)])
    df.insert(6, first_lbl, [period_label(int(v) % 100, int(v) // 100) for v in df["_first"]])
    df = df.drop(columns=["_first"])
    df["Margin"] = df["BaseSalesValue"] - df["CTC"]
    df["MarginPct"] = (df["Margin"] * 100 / df["CTC"].where(df["CTC"] != 0)).round(2)
    ct, bs = df["CTC"].sum(), df["BaseSalesValue"].sum()
    pw, pp = build_where(mode, periods, f)               # distinct people: one person can work on 2 work orders
    pairs = run_query(f"SELECT DISTINCT [WorkOrderNo], [Employee] FROM {T} {pw}", pp)
    n_emp = pairs.loc[pairs["WorkOrderNo"].isin(df["WorkOrderNo"]), "Employee"].nunique()
    if only_first:
        head = (f"**{len(df)} new work order(s) in {mtxt} {ptxt}** (in {mtxt} {ptxt} for the first time, never "
                f"in an earlier {mtxt})")
    else:
        head = f"**{len(df)} work order(s) received in {ptxt}** (work order start date in {ptxt})"
    text = (head + f" · {n_emp:,} employees · CTC {inr(ct)} · base sales value {inr(bs)}"
            f" · margin {inr(bs - ct)}" + (f" ({(bs - ct) * 100 / ct:.2f}% on CTC)" if ct else "") + ".")
    text += ("\n\nℹ️ Want work orders by **start date** instead? Ask *work orders received in " + ptxt.lower() + "*."
             if only_first else "\n\nℹ️ Want work orders **paid for the first time** instead? Ask *new work orders in "
             "liable month " + ptxt.lower() + "*.")
    notes = [f"Understood as: **{what.capitalize()}** (no month type named = start date (WOStartDate); liable / "
             f"salary / payment month named = first month the work order appears in that month type)."] + notes + [
        f"Employees / CTC / base sales are for {mode_label(mode)} {ptxt}.", f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "new_wo", "sql": sql, "params": params,
            "intent": {"report": "new_wo"}}


EXPIRED_RX = re.compile(r"\b(?:expired|expire|expiring|expiry|ended|closed|completed|finished|lapsed)\b", re.I)


def expired_wo_answer(question, mode, notes, t0):
    """'expired work orders for sep-26': work orders whose WOEndDate fell in the month, and - flagged - work orders that
    had ALREADY expired before the month but were still paid salary in it."""
    if not (WO_WORDS.search(question) or re.search(r"\bworkorders?\b", question, re.I)) or not EXPIRED_RX.search(question):
        return None
    if re.search(r"\b(before|after|by|till|until|upto|up\s+to|next|coming|upcoming|within|later\s+than|earlier\s+than)\b",
                 question, re.I):
        return None                                      # "ending before dec-26", "next 3 months" -> wo_end_answer
    periods = parse_periods(question) or [latest_period(mode)]
    (m1, y1), (m2, y2) = periods[0], periods[-1]
    d1, d2 = datetime.date(y1, m1, 1), _month_end(y2, m2) + datetime.timedelta(days=1)
    yms = [y * 100 + m for m, y in periods]
    inn = ",".join("?" * len(yms))
    Y, M = period_exprs(mode)
    pk = f"({Y}*100+{M})"                                   # the month in the question's month type
    # "paid after expiry" = a row paid in the month whose SALARY month (Month+Year) is after the WO end month:
    # Aug-26 salary paid in liable / payment month Sep-26 for a WO ending in Aug-26 is fine, and so are arrears
    late = f"{pk} IN ({inn}) AND [Year]*100+[Month] > YEAR([WOEndDate])*100+MONTH([WOEndDate])"
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, wparams = build_where(mode, [], f)
    sql = f"""SELECT [WorkOrderNo], MAX([Category]) AS Category, MAX([Department]) AS Department,
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate,
       MAX({pk}) AS [_last],
       COUNT(DISTINCT CASE WHEN {pk} IN ({inn}) THEN [Employee] END) AS [Employees paid in month],
       SUM(CASE WHEN {pk} IN ({inn}) THEN [CTC] ELSE 0 END) AS [CTC paid in month],
       SUM(CASE WHEN {late} THEN [CTC] ELSE 0 END) AS [_late_ctc]
FROM {T} {where}
GROUP BY [WorkOrderNo]
HAVING (MAX([WOEndDate]) >= ? AND MAX([WOEndDate]) < ?)
    OR (MAX([WOEndDate]) < ? AND SUM(CASE WHEN {late} THEN 1 ELSE 0 END) > 0)
ORDER BY MAX([WOEndDate]), [WorkOrderNo]"""
    params = yms * 3 + list(wparams) + [d1, d2, d1] + yms
    t1 = time.time()
    df = run_query(sql, params)
    ptxt = period_label(m1, y1) if (m1, y1) == (m2, y2) else f"{period_label(m1, y1)} to {period_label(m2, y2)}"
    if df.empty:
        return {"text": f"**No work orders expired in {ptxt}**, and none were paid after their end date.",
                "note": "Understood as: **Expired work orders** (WOEndDate).", "df": None, "report": "expired_wo",
                "sql": sql, "params": params, "intent": {"report": "expired_wo"}}
    end = pd.to_datetime(df["WOEndDate"], errors="coerce")
    in_month = (end >= pd.Timestamp(d1)) & (end < pd.Timestamp(d2))
    df.insert(1, "Status", ["Expired in " + ptxt if x else "⚠️ Expired earlier, still paid in " + ptxt for x in in_month])
    last_lbl = {"month_year": "Last liable month", "payment": "Last payment month"}.get(mode, "Last salary month")
    df.insert(7, last_lbl, [period_label(int(v) % 100, int(v) // 100) for v in df["_last"]])
    n_in, n_late = int(in_month.sum()), int((~in_month).sum())
    late_ctc = df.loc[~in_month, "_late_ctc"].sum()
    df = df.drop(columns=["_last", "_late_ctc"])
    text = f"**{n_in} work order(s) expired in {ptxt}**"
    if n_in:
        first = df[in_month].iloc[0]
        text += f" (first: {first['WorkOrderNo']} on {pd.Timestamp(first['WOEndDate']).strftime('%d-%b-%Y')})"
    text += "."
    if n_late:
        text += (f"\n\n⚠️ **{n_late} work order(s) had already expired before {period_label(m1, y1)} but were still paid** "
                 f"for salary months after their end date: CTC {inr(late_ctc)} in {ptxt} — please check whether they "
                 f"were extended.")
    notes = ["Understood as: **Expired work orders** (WOEndDate in the month, plus work orders paid in the month for a "
             "salary month (Month/Year) after their end month - salary of the end month paid a month later, and "
             "arrears, are not flagged)."
             ] + notes + [f"Paid-in-month figures use **{mode_label(mode)}** {ptxt}.",
                          f"⏱️ No AI · SQL: {time.time() - t1:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "expired_wo", "sql": sql, "params": params,
            "intent": {"report": "expired_wo"}}


WO_END_RX = re.compile(r"\b(?:end|ends|ending|ended|expire|expires|expiring|expired|expiry|close|closes|closing|"
                       r"finish|finishes|finishing|complete|completes|completing)\b", re.I)


def wo_end_answer(question, mode, notes, t0):
    """'which work orders end before dec-26', 'work orders ending in next 3 months', 'work orders expiring in nov-26'."""
    if not (WO_WORDS.search(question) or re.search(r"\bworkorders?\b", question, re.I)) or not WO_END_RX.search(question):
        return None
    q = question.lower()
    today = datetime.date.today()
    nm = re.search(r"\b(?:next|coming|upcoming)\s+(\d{1,2})\s+months?\b", q)
    periods = parse_periods(question)
    if nm:
        k = int(nm.group(1))
        y, m = today.year + (today.month - 1 + k) // 12, (today.month - 1 + k) % 12 + 1
        lo, hi = today, _month_end(y, m)
        desc = f"between today and {period_label(m, y)}"
    elif re.search(r"\bthis\s+month\b|\bnext\s+month\b", q):
        k = 1 if "next month" in q else 0
        y, m = today.year + (today.month - 1 + k) // 12, (today.month - 1 + k) % 12 + 1
        lo, hi = datetime.date(y, m, 1), _month_end(y, m)
        desc = f"in {period_label(m, y)}"
    elif periods:
        (m1, y1), (m2, y2) = periods[0], periods[-1]
        start, end = datetime.date(y1, m1, 1), _month_end(y2, m2)
        if re.search(r"\b(before|earlier\s+than|less\s+than|prior\s+to)\b", q):
            lo, hi, desc = None, start - datetime.timedelta(days=1), f"before {period_label(m1, y1)}"
        elif re.search(r"\b(by|till|until|upto|up\s+to|on\s+or\s+before)\b", q):
            lo, hi, desc = None, end, f"by the end of {period_label(m2, y2)}"
        elif re.search(r"\b(after|later\s+than|greater\s+than|beyond)\b", q):
            lo, hi, desc = end + datetime.timedelta(days=1), None, f"after {period_label(m2, y2)}"
        else:
            lo, hi, desc = start, end, "in " + (period_label(m1, y1) if (m1, y1) == (m2, y2) else
                                               f"{period_label(m1, y1)} to {period_label(m2, y2)}")
    else:
        return None
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where(mode, [], f)
    lm, ly = latest_period(mode)                        # "running now" = paid in the latest month of the asked month type
    Y, M = period_exprs(mode)
    cond, hp = [], []
    if lo:
        cond.append("MAX([WOEndDate]) >= ?")
        hp.append(lo)
    if hi:
        cond.append("MAX([WOEndDate]) <= ?")
        hp.append(hi)
    running = not re.search(r"\ball\b|\bever\b|\bincluding\s+(old|closed|finished)", q)
    if running:
        cond.append(f"MAX(CASE WHEN {Y} = ? AND {M} = ? THEN 1 ELSE 0 END) = 1")
        hp += [ly, lm]
    sql = f"""SELECT [WorkOrderNo], MAX([Category]) AS Category, MAX([Department]) AS Department,
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate,
       COUNT(DISTINCT CASE WHEN {Y} = ? AND {M} = ? THEN [Employee] END) AS [Employees ({period_label(lm, ly)})],
       SUM(CASE WHEN {Y} = ? AND {M} = ? THEN [CTC] ELSE 0 END) AS [CTC ({period_label(lm, ly)})],
       SUM(CASE WHEN {Y} = ? AND {M} = ? THEN [_Base] ELSE 0 END) AS [Base sales ({period_label(lm, ly)})]
FROM {base_rows(Y, M, where)} GROUP BY [WorkOrderNo] HAVING {' AND '.join(cond)} ORDER BY MAX([WOEndDate]), [WorkOrderNo]"""
    params = [ly, lm, ly, lm, ly, lm] + list(params) + hp
    t1 = time.time()
    df = run_query(sql, params)
    scope = f" that are running now (paid in {mode_label(mode).split(' (')[0]} {period_label(lm, ly)})" if running else ""
    if df.empty:
        text = f"**No work orders{scope} end {desc}.**"
    else:
        first = df.iloc[0]
        text = (f"**{len(df):,} work order(s){scope} end {desc}.** Earliest: **{first['WorkOrderNo']}** on "
                f"{pd.Timestamp(first['WOEndDate']).strftime('%d-%b-%Y') if pd.notna(first['WOEndDate']) else '–'}.")
        if running:                                      # "revenue at risk": what these work orders bring in a month
            ct, bs = df[f"CTC ({period_label(lm, ly)})"].sum(), df[f"Base sales ({period_label(lm, ly)})"].sum()
            text += (f" In {period_label(lm, ly)} they brought **base sales value {inr(bs)}** (revenue at risk) · "
                     f"CTC {inr(ct)} · margin {inr(bs - ct)}.")
    notes = ["Understood as: **Work orders by end date** (WOEndDate)."] + notes
    if running:
        notes.append(f"Only work orders that still have salary in the latest {mode_label(mode)}. "
                     "Add *all* to include finished ones.")
    notes.append(f"⏱️ No AI · SQL: {time.time() - t1:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": df, "report": "wo_end", "sql": sql, "params": params,
            "intent": {"report": "wo_end"}}


def relative_months(question, mode):
    """'this month' = latest month in the data (month_year Aug-26 is paid in Sep), 'last month' = one before,
    'this year' / 'ytd' = April (financial year start) to the latest month, 'last 3 months'. Returns (new question, phrase)."""
    m0 = REL_MONTH.search(question)
    if not m0 or parse_periods(question):
        return None
    try:
        lm, ly = latest_period(mode)
    except Exception:
        return None
    w = m0.group(0).lower()
    if m0.group(2):                                   # this / last month
        m, y = _shift(lm, ly, 1) if re.match(r"(last|previous|prev|pichhle|pichla)", w) else (lm, ly)
        text = f"{MON[m - 1]}-{y}"
    elif m0.group(5):                                 # last N months
        k = max(1, min(24, int(m0.group(5))))
        m1, y1 = _shift(lm, ly, k - 1)
        text = f"{MON[m1 - 1]}-{y1} to {MON[lm - 1]}-{ly}"
    else:                                             # this / last financial year, ytd
        fy = ly if lm >= 4 else ly - 1
        if w.startswith(("last", "previous")):
            text = f"Apr-{fy - 1} to Mar-{fy}"
        else:
            text = f"Apr-{fy} to {MON[lm - 1]}-{ly}"
    return question[:m0.start()] + text + question[m0.end():], m0.group(0)


# ======================= PROJECTION (next / future month CTC and base sales) =======================
# Same rule as the user's "proposed data" Excel (columns N, O, P, Q). Base = SALARY month (Year/Month columns).
#   Days      = last day of base month - WOStartDate + 1
#   0         if WOEndDate is blank or ends before the projected month starts
#   daily     = value / Days                 if the work order started during the base month (Days < days in base month)
#             = value / days in projected month   otherwise
#   projected = ROUND(daily x MIN(days in projected month, WOEndDate - day before projected month), 0)
#   value     = CTC (sum of the work order's employees) and Base sales = ROUND(Invoice x 100/118, 2) once per WO
PROJ_WORDS = re.compile(r"\b(project\w*|forecast\w*|future|estimat\w*|expected|anticipat\w*|predict\w*|next\s+month|"
                        r"upcoming|coming\s+month)\b", re.IGNORECASE)


def wants_projection(question):
    return bool(PROJ_WORDS.search(question))


def _month_end(y, m):
    nxt = datetime.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    return nxt - datetime.timedelta(days=1)


def _proj_one(r, by, bm, ty, tm):
    """(days in base month, days running in target month, projected CTC, projected base sales) of one work order."""
    base_end = _month_end(by, bm)
    t_start, t_end = datetime.date(ty, tm, 1), _month_end(ty, tm)
    days_b, days_t = base_end.day, t_end.day
    start = pd.to_datetime(r.get("WOStartDate"), errors="coerce")
    end = pd.to_datetime(r.get("WOEndDate"), errors="coerce")
    days = (base_end - start.date()).days + 1 if pd.notna(start) else days_b + 1
    if pd.isna(end) or end.date() < t_start:
        return days, 0, 0.0, 0.0
    active = min(days_t, (end.date() - (t_start - datetime.timedelta(days=1))).days)
    div = max(days, 1) if days < days_b else days_t
    return (days, active, round(float(r["CTC"] or 0) / div * active),
            round(float(r["BaseSalesValue"] or 0) / div * active))


def project_rows(df, by, bm, ty, tm):
    """Adds the projected columns to work-order rows of base month bm/by for target month tm/ty."""
    out = [_proj_one(r, by, bm, ty, tm) for _, r in df.iterrows()]
    df = df.copy()
    df["Days"] = [o[0] for o in out]
    df["ProjDays"] = [o[1] for o in out]
    df["ProjCTC"] = [o[2] for o in out]
    df["ProjBaseSales"] = [o[3] for o in out]
    df["ProjMargin"] = df["ProjBaseSales"] - df["ProjCTC"]
    df["ProjMarginPct"] = (df["ProjMargin"] * 100 / df["ProjCTC"].where(df["ProjCTC"] != 0)).round(2)
    return df


FILTER_IGNORE = {"projection", "projections", "projected", "projectile", "forecast", "forecasted", "future", "estimate",
                 "estimated", "expected", "using", "use", "data", "based", "basis", "workorder", "workorders", "wise",
                 "next", "months", "month", "only", "upcoming", "coming", "predict", "prediction", "give", "show",
                 "and", "with", "from", "for", "the", "of", "all", "each", "every", "total", "please", "detail",
                 "details", "list", "report", "want", "need", "tell", "what", "how", "much", "will", "would",
                 "giv", "gimme", "plz", "pls", "me", "column", "columns", "last", "paid", "salary", "month_year",
                 "add", "also", "include", "which", "based", "on", "in", "is", "it"}


def _word_backs(value, words):
    """True if one of the not-understood words looks like (part of) the value, typos allowed:
    'openvm' -> OPENVEM, 'firm' -> Efirm. 'average' -> nothing."""
    import difflib
    vn = re.sub(r"[^a-z0-9]", "", str(value).lower())
    parts = [p_ for p_ in re.findall(r"[a-z0-9]+", str(value).lower()) if len(p_) >= 3] + [vn]
    for w in words:
        if len(w) < 3:
            continue
        if w in vn or any(difflib.SequenceMatcher(None, w, p_).ratio() >= 0.8 for p_ in parts):
            return True
    return False


def understand_filters(question, f, notes, use_ai=True, extra_ignore=frozenset()):
    """Makes sure every word of the question was used. If some words are not understood, Qwen reads the
    WHOLE question and picks department / category from the real lists (never invents a value)."""
    left = unknown_words(question, f) - FILTER_IGNORE - set(extra_ignore)
    for v in (f.get("category"), f.get("department")):     # "open vem" was matched to OPENVEM
        if v:
            vc = re.sub(r"[^a-z0-9]", "", str(v).lower())
            left = {w for w in left if not (len(w) >= 3 and w in vc)}
    left = {w for w in left if not re.fullmatch(r"\d+", w)}
    if not left or not use_ai:
        if left:
            notes.append(f"⚠️ Not understood: **{', '.join(sorted(left))}** (AI is off) - check the filter above.")
        return f
    try:
        cats, deps = distinct_values("Category"), distinct_values("Department")
        system = ("You read a payroll question and find which CATEGORY and/or DEPARTMENT it is limited to. "
                  "Reply ONLY with JSON: {\"category\": null or one value from the list, "
                  "\"department\": null or one value from the list}. Choose ONLY exact values from these lists; "
                  "if the question does not limit to one, use null. Words may be split or misspelt "
                  "(e.g. 'open vem' = OPENVEM).\n"
                  f"Categories: {', '.join(cats[:60])}\nDepartments: {', '.join(deps[:150])}")
        resp = ollama.chat(model=MODEL, messages=[{"role": "system", "content": system},
                                                  {"role": "user", "content": question}],
                           format="json", keep_alive="60m",
                           options={"temperature": 0, "num_predict": 80, "num_ctx": 4096, "num_thread": AI_THREADS})
        data = json.loads(resp["message"]["content"])
        got = []
        for key, col, vals in (("category", "Category", cats), ("department", "Department", deps)):
            v = _clean(data.get(key))
            hit = next((x for x in vals if v and x.lower() == str(v).lower()), None)
            if hit and not _word_backs(hit, left):
                continue                                  # the AI picked a value none of the user's words point to
            if hit and not f.get(key):
                f[key] = hit
                got.append(f"{col} = {hit}")
        if got:
            notes.append(f"🧠 Qwen read your full question: **{'; '.join(got)}**.")
        else:
            notes.append(f"⚠️ Not understood: **{', '.join(sorted(left))}** - no filter applied for it.")
    except Exception as e:
        notes.append(f"⚠️ Not understood: **{', '.join(sorted(left))}** (AI not reachable: {e.__class__.__name__}).")
    return f


BASE_HINT = re.compile(r"\b(?:using|use|based\s+on|on\s+(?:the\s+)?basis\s+of|basis\s+of|basis|with|from)\s+(?:the\s+)?"
                       r"((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[\s\-/'.]*\d{2,4})"
                       r"(?!\s*(?:to|till|until|-)\s*(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec))", re.IGNORECASE)
MAX_PROJ_MONTHS = 24
PROJ_LIABLE_MONTHS = 3   # a projection reads only the last 3 liable months before the projected month


def _add_months(m, y, k):
    n = y * 12 + m - 1 + k
    return n % 12 + 1, n // 12


def _dmy(v):
    v = pd.to_datetime(v, errors="coerce")
    return v.strftime("%d-%b-%y") if pd.notna(v) else "no date"


def proj_snapshot(first, f):
    """The data as it was BEFORE liable month `first` (= the day the projection is made, e.g. 31-Aug-26 for Sep-26):
    only the last PROJ_LIABLE_MONTHS liable months before it (Jun, Jul, Aug for Sep), one row per work order at its
    OWN last salary month (Year/Month) in them.
    Returns (rows, latest liable month in that data (m, y), sql, params, left_out, known) - rows is None when there is
    no data; left_out = running work orders last booked before those months; known = every work order before `first`."""
    tm, ty = first
    cut = datetime.date(ty, tm, 1)
    lm_, ly_ = _add_months(tm, ty, -PROJ_LIABLE_MONTHS)
    lo = datetime.date(ly_, lm_, 1)
    c = run_query(f"SELECT MAX([month_year]) AS c FROM {T} WHERE [month_year] >= ? AND [month_year] < ?",
                  [lo, cut])["c"].iloc[0]
    if pd.isna(c):
        return None, None, "", [], [], set()
    c = pd.to_datetime(c)
    where0, params0 = build_where("salary", [], f)
    w_all = (where0 + " AND " if where0 else "WHERE ") + "[month_year] < ?"
    allw = run_query(f"SELECT [WorkOrderNo], MAX([month_year]) AS LastMY, MAX([WOEndDate]) AS WOEndDate "
                     f"FROM {T} {w_all} GROUP BY [WorkOrderNo]", params0 + [cut])
    known = set(allw["WorkOrderNo"])
    old = allw[(pd.to_datetime(allw["LastMY"]) < pd.Timestamp(lo))
               & (pd.to_datetime(allw["WOEndDate"], errors="coerce") >= pd.Timestamp(cut))]
    left_out = sorted(map(str, old["WorkOrderNo"]))
    where = (where0 + " AND " if where0 else "WHERE ") + "[month_year] >= ? AND [month_year] < ?"
    params = params0 + [lo, cut]
    sql = f"""WITH s AS (SELECT * FROM {T} {where}),
l AS (SELECT [WorkOrderNo] AS W, MAX([Year]*100+[Month]) AS L FROM s GROUP BY [WorkOrderNo])
SELECT MAX(s.[Category]) AS Category, MAX(s.[Department]) AS Department, s.[WorkOrderNo],
       MAX(s.[WOStartDate]) AS WOStartDate, MAX(s.[WOEndDate]) AS WOEndDate,
       COUNT(DISTINCT s.[Employee]) AS Employees, SUM(s.[CTC]) AS CTC,
       ROUND(MAX(s.[InvoiceAmount] * 100.0 / {100 + GST_RATE}), 2) AS BaseSalesValue,
       FORMAT(MAX(s.[month_year]), 'MMM-yyyy') AS LastPaidMY, MAX(l.L) AS LastSalary
FROM s JOIN l ON s.[WorkOrderNo] = l.W AND s.[Year]*100+s.[Month] = l.L
GROUP BY s.[WorkOrderNo] ORDER BY s.[WorkOrderNo]"""
    return run_query(sql, params), (c.month, c.year), sql, params, left_out, known


def proj_catchup(raw, c, targets):
    """Projects every work order from its OWN last salary month. Salary months after it that are already over
    (up to liable month c) and still not paid are PENDING: they are added to the first projected month, with a
    remark ('2 months salary to pay: Aug-26 (pending) + Sep-26'). Rule per month = project_rows."""
    cm, cy = c
    (t1m, t1y) = targets[0]
    rows = []
    for _, r in raw.iterrows():
        L = int(r["LastSalary"])
        lm, ly = L % 100, L // 100
        pend, p_ctc, p_base = [], 0.0, 0.0
        m_, y_ = _add_months(lm, ly, 1)
        while (y_, m_) <= (cy, cm):
            _, act, pc, pb = _proj_one(r, ly, lm, y_, m_)
            if act > 0:
                pend.append(period_label(m_, y_))
                p_ctc, p_base = p_ctc + pc, p_base + pb
            m_, y_ = _add_months(m_, y_, 1)
        out = {"LastSalaryLabel": period_label(lm, ly), "Pending": pend, "PendingCTC": p_ctc, "PendingBase": p_base}
        any_ = bool(pend)
        for i, (tm, ty) in enumerate(targets):
            _, act, pc, pb = _proj_one(r, ly, lm, ty, tm)
            if i == 0:
                pc, pb = pc + p_ctc, pb + p_base
                out["T1Active"] = act
            out[(tm, ty)] = (act, pc, pb)
            any_ = any_ or act > 0
        if not any_:
            continue                                    # ended - nothing left to pay
        t1 = period_label(t1m, t1y)
        if pend and out["T1Active"]:
            out["Remark"] = (f"{len(pend) + 1} months salary to pay: " + " + ".join(f"{p_} (pending)" for p_ in pend)
                             + f" + {t1}")
        elif pend:
            out["Remark"] = (f"{len(pend)} month{'s' if len(pend) > 1 else ''} salary pending: " + " + ".join(pend)
                             + f" (work order ended {_dmy(r.get('WOEndDate'))})")
        elif 0 < out["T1Active"] < _month_end(t1y, t1m).day:
            out["Remark"] = f"Part month: {out['T1Active']} days (work order ends {_dmy(r.get('WOEndDate'))})"
        else:
            out["Remark"] = ""
        rows.append((r, out))
    return rows


def projection_answer(question, mode, notes, t0, use_ai=True):
    """'projection for sep-26', 'projection of sep-26 and oct-26 using aug-26 data', 'next 3 months projection'.
    One row per work order; every projected month gets its OWN columns (no extra rows)."""
    latest = run_query(f"SELECT MAX([Year]*100+[Month]) AS p FROM {T}", [])["p"].iloc[0]
    latest = (int(latest) % 100, int(latest) // 100) if pd.notna(latest) else None
    if not latest:
        return clarify(question, "There is no salary data to project from.")
    base = None
    hint = BASE_HINT.search(question)                       # "... using aug-26 data"
    if hint:
        bp = parse_periods(hint.group(1))
        base = bp[0] if bp else None
        q_targets = question[:hint.start()] + " " + question[hint.end():]
    else:
        q_targets = question
    targets = sorted(set(parse_periods(q_targets)), key=lambda p_: (p_[1], p_[0]))
    nm = re.search(r"\bnext\s+(\d{1,2})\s+months?\b", question, re.IGNORECASE)
    if not targets:                                         # "next month" / "next 3 months"
        if base:
            b = base
        else:                                               # after the latest LIABLE month
            lmy = pd.to_datetime(run_query(f"SELECT MAX([month_year]) AS p FROM {T}", [])["p"].iloc[0])
            b = (lmy.month, lmy.year) if pd.notna(lmy) else latest
        k = int(nm.group(1)) if nm else 1
        targets, (m_, y_) = [], b
        for _ in range(max(1, min(MAX_PROJ_MONTHS, k))):
            m_, y_ = (1, y_ + 1) if m_ == 12 else (m_ + 1, y_)
            targets.append((m_, y_))
    if not hint:                                            # as on the day before the month: pending months added
        return catchup_projection_answer(question, targets[:MAX_PROJ_MONTHS], notes, use_ai)
    if base is None:                                        # latest salary month before the first target
        tm, ty = targets[0]
        prev = run_query(f"SELECT MAX([Year]*100+[Month]) AS p FROM {T} WHERE [Year]*100+[Month] < ?",
                         [ty * 100 + tm])["p"].iloc[0]
        if pd.isna(prev):
            return clarify(question, f"There is no salary data before **{period_label(tm, ty)}** to project from.")
        base = (int(prev) % 100, int(prev) // 100)
    bm, by = base
    partial = None                                          # base month only partly loaded -> projection too low
    if not hint:
        cnt = run_query(f"SELECT [Year]*100+[Month] AS p, COUNT(DISTINCT [Employee]) AS n FROM {T} "
                        f"WHERE [Year]*100+[Month] <= ? GROUP BY [Year]*100+[Month] ORDER BY p DESC",
                        [by * 100 + bm]).head(4)
        if len(cnt) >= 2 and int(cnt["p"].iloc[0]) == by * 100 + bm:
            usual = float(cnt["n"].iloc[1:].median())
            if usual and cnt["n"].iloc[0] < 0.7 * usual:
                full = cnt[cnt["n"] >= 0.7 * usual]["p"]
                partial = (int(cnt["n"].iloc[0]), int(usual),
                           period_label(int(full.iloc[0]) % 100, int(full.iloc[0]) // 100) if len(full) else None)
    targets = [t_ for t_ in targets if (t_[1], t_[0]) > (by, bm)][:MAX_PROJ_MONTHS]
    if not targets:
        return clarify(question, "The projected month(s) must be after the month the projection is based on.")
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    f = understand_filters(question, f, notes, use_ai)        # nothing in the question is ignored
    where, params = build_where("salary", [(bm, by)], f)       # base month = SALARY month (Year/Month)
    cut = datetime.date(targets[0][1], targets[0][0], 1)       # only the last 3 liable months BEFORE the projected month
    lm_, ly_ = _add_months(targets[0][0], targets[0][1], -PROJ_LIABLE_MONTHS)
    where = (where + " AND " if where else "WHERE ") + "[month_year] >= ? AND [month_year] < ?"
    params = params + [datetime.date(ly_, lm_, 1), cut]
    sql = f"""SELECT MAX([Category]) AS Category, MAX([Department]) AS Department, [WorkOrderNo],
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate,
       COUNT(DISTINCT [Employee]) AS Employees, SUM([CTC]) AS CTC,
       ROUND(MAX([InvoiceAmount] * 100.0 / {100 + GST_RATE}), 2) AS BaseSalesValue,
       FORMAT(MAX([month_year]), 'MMM-yyyy') AS LastPaidMY
FROM {T} {where} GROUP BY [WorkOrderNo], [Year], [Month] ORDER BY [WorkOrderNo]"""
    t1 = time.time()
    raw = run_query(sql, params)
    t_sql = time.time() - t1
    if raw.empty:
        return clarify(question, f"No salary data found for **{period_label(bm, by)}** to project from.")
    bl = period_label(bm, by)
    df = raw.copy()
    totals, lines = [], []
    for tm, ty in targets:
        tl = period_label(tm, ty)
        pr = project_rows(raw, by, bm, ty, tm)
        if f"Days in {bl}" not in df:
            df[f"Days in {bl}"] = pr["Days"]
        df[f"Days in {tl}"] = pr["ProjDays"]
        df[f"Projected CTC {tl}"] = pr["ProjCTC"]
        df[f"Projected base sales {tl}"] = pr["ProjBaseSales"]
        df[f"Projected margin {tl}"] = pr["ProjMargin"]
        pc, pb = pr["ProjCTC"].sum(), pr["ProjBaseSales"].sum()
        totals.append((tl, pc, pb))
        running = int((pr["ProjCTC"] > 0).sum())
        part = int(((pr["ProjDays"] > 0) & (pr["ProjDays"] < _month_end(ty, tm).day)).sum())
        pct = f" ({(pb - pc) * 100 / pc:.2f}%)" if pc else ""
        lines.append(f"- **{tl}:** CTC {inr(pc)} · Base sales {inr(pb)} · Margin {inr(pb - pc)}{pct} · "
                     f"{running} work order(s) running" + (f", {part} part month" if part else ""))
    df = df.rename(columns={"CTC": f"CTC {bl}", "BaseSalesValue": f"Base sales {bl}",
                            "LastPaidMY": f"{bl} salary paid in (month_year)"})
    if f"{bl} salary paid in (month_year)" in df:           # right after the work order no.
        c_ = df.pop(f"{bl} salary paid in (month_year)")
        df.insert(list(df.columns).index("WorkOrderNo") + 1, f"{bl} salary paid in (month_year)", c_)
    mcols = [c for c in df.columns if str(c).startswith("Projected margin ")]
    if len(targets) > 1:                                    # totals of all projected months, as columns
        df["Total projected CTC"] = df[[c for c in df.columns if str(c).startswith("Projected CTC ")]].sum(axis=1)
        df["Total projected base sales"] = df[[c for c in df.columns
                                               if str(c).startswith("Projected base sales ")]].sum(axis=1)
        df["Total projected margin"] = df["Total projected base sales"] - df["Total projected CTC"]
        sort_col = "Total projected margin"
    else:
        sort_col = mcols[0]
    t = set(re.findall(r"[a-z]+", question.lower()))
    group = "Department" if t & {"department", "departments", "dept", "departmentwise"} else \
        "Category" if t & {"category", "categories", "categorywise"} else None
    if group:                                               # same columns, added up per department / category
        num = [c for c in df.columns if str(c).startswith(("CTC ", "Base sales ", "Projected ", "Total projected"))]
        df = df.groupby(group, as_index=False)[num].sum()
    df = df.sort_values(sort_col, ascending=False).reset_index(drop=True)
    all_c, all_b = sum(x[1] for x in totals), sum(x[2] for x in totals)
    tl_text = totals[0][0] if len(totals) == 1 else f"{totals[0][0]} to {totals[-1][0]}"
    pct = f" ({(all_b - all_c) * 100 / all_c:.2f}% on CTC)" if all_c else ""
    only = " · ".join(v for v in (f.get("category"), f.get("department")) if v) or \
        ", ".join(current_exclusions().get("include_Category", []) + current_exclusions().get("include_Department", []))
    exc = excluded_text()
    scope = (f" — {only} only" if only else "") + (f" — excluding {exc}" if exc else "")
    text = (f"**Projected for {tl_text}{scope} (from {bl} salary data):** CTC {inr(all_c)} · Base sales value {inr(all_b)} · "
            f"Margin {inr(all_b - all_c)}{pct}\n\n" + ("\n".join(lines) + "\n" if len(lines) > 1 else "- " + lines[0].split(" · ")[-1] + "\n") +
            f"- {bl} actual: CTC {inr(raw['CTC'].sum())} · Base sales value {inr(raw['BaseSalesValue'].sum())}")
    notes.insert(0, f"Understood as: **Projection** for {tl_text} using {bl} (salary month Year/Month), only data "
                    f"booked in liable {period_label(lm_, ly_)} to {period_label(*_add_months(targets[0][0], targets[0][1], -1))}"
                    + (f", {group}-wise" if group else ", work order-wise") + ". Each month has its own columns.")
    notes.append(f"**{bl} salary paid in (month_year)** = the latest month_year in which that work order's {bl} "
                 "salary was processed - the salary this projection is based on.")
    notes.append("Rule (same as your Excel): 0 if the work order ends before the month or has no end date; "
                 "daily rate = value ÷ days in that month, or ÷ days it ran in the base month if it "
                 "started in that month; × days it runs in that month.")
    notes.append(f"⏱️ No AI · SQL: {t_sql:.1f}s")
    intent = {"report": "projection", "base": bl, "targets": [x[0] for x in totals], "group": group,
              "department": f.get("department"), "category": f.get("category"),
              "excluded": excluded_text()}
    options = []
    if partial:
        text = (f"⚠️ **{bl} salary data looks incomplete** ({partial[0]:,} employees vs about {partial[1]:,} in a normal "
                f"month), so this projection is probably **too low**."
                + (f" Use the last complete month instead: *{q_targets.strip()} using {partial[2].lower()} data*." if partial[2] else "")
                + "\n\n" + text)
        if partial[2]:
            options = [f"{q_targets.strip()} using {partial[2].lower()} data"]
    return {"text": text, "note": " ".join(notes), "df": df, "report": "projection", "sql": sql,
            "params": params, "intent": intent, "proj_totals": (all_c, all_b), "options": options}


def _proj_scope_text(f):
    only = " · ".join(v for v in (f.get("category"), f.get("department")) if v) or \
        ", ".join(current_exclusions().get("include_Category", []) + current_exclusions().get("include_Department", []))
    exc = excluded_text()
    return (f" — {only} only" if only else "") + (f" — excluding {exc}" if exc else "")


def catchup_projection(targets, f):
    """Projection of liable month(s) `targets` as it would be made on the last day before the first one.
    Returns None or a dict with the per-work-order table and totals."""
    t1 = time.time()
    raw, c, sql, params, left_out, known = proj_snapshot(targets[0], f)
    if raw is None or raw.empty:
        return None
    rows = proj_catchup(raw, c, targets)
    labels = [period_label(m_, y_) for m_, y_ in targets]
    recs = []
    for r, o in rows:
        d = {"Category": r["Category"], "Department": r["Department"], "WorkOrderNo": r["WorkOrderNo"],
             "Last salary month": o["LastSalaryLabel"], "Last salary paid in (month_year)": r["LastPaidMY"],
             "WOStartDate": r["WOStartDate"], "WOEndDate": r["WOEndDate"], "Employees": r["Employees"],
             "CTC in last salary month": r["CTC"], "Base sales in last salary month": r["BaseSalesValue"],
             "Pending months": len(o["Pending"]), f"Pending months CTC (in {labels[0]})": o["PendingCTC"],
             f"Pending months base sales (in {labels[0]})": o["PendingBase"]}
        for (tm, ty), tl in zip(targets, labels):
            act, pc, pb = o[(tm, ty)]
            d[f"Days in {tl}"] = act
            d[f"Projected CTC {tl}"] = pc
            d[f"Projected base sales {tl}"] = pb
            d[f"Projected margin {tl}"] = pb - pc
        d["Remarks"] = o["Remark"]
        d["_PayMonths"] = o["Pending"] + ([labels[0]] if o["T1Active"] else [])   # salary months paid in labels[0]
        recs.append(d)
    df = pd.DataFrame(recs)
    w0 = _add_months(*targets[0], -PROJ_LIABLE_MONTHS)
    return {"df": df, "c": c, "labels": labels, "left_out": left_out, "sql": sql, "params": params,
            "t_sql": time.time() - t1, "known": known,
            "window": f"{period_label(*w0)} to {period_label(*_add_months(*targets[0], -1))}"}


def catchup_projection_answer(question, targets, notes, use_ai=True):
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    f = understand_filters(question, f, notes, use_ai)
    first = period_label(*targets[0])
    p = catchup_projection(targets, f)
    if not p or p["df"].empty:
        return clarify(question, f"There is no salary data before liable month **{first}** to project from.")
    df, labels, (cm, cy) = p["df"].drop(columns=["_PayMonths"]), p["labels"], p["c"]
    cl, as_on = period_label(cm, cy), _month_end(cy, cm).strftime("%d-%b-%y")
    totals, lines = [], []
    for tl in labels:
        pc, pb = df[f"Projected CTC {tl}"].sum(), df[f"Projected base sales {tl}"].sum()
        totals.append((tl, pc, pb))
        running = int((df[f"Days in {tl}"] > 0).sum())
        pct = f" ({(pb - pc) * 100 / pc:.2f}%)" if pc else ""
        lines.append(f"- **{tl}:** CTC {inr(pc)} · Base sales {inr(pb)} · Margin {inr(pb - pc)}{pct} · "
                     f"{running} work order(s) running")
    pend = df[df["Pending months"] > 0]
    if not pend.empty:
        by_n = pend.groupby("Pending months").size()
        lines.append(f"- **Pending salary added to {labels[0]}:** {len(pend)} work order(s) · CTC "
                     f"{inr(pend[f'Pending months CTC (in {labels[0]})'].sum())} · base sales "
                     f"{inr(pend[f'Pending months base sales (in {labels[0]})'].sum())} ("
                     + ", ".join(f"{int(v)} with {k} month{'s' if k > 1 else ''} pending" for k, v in by_n.items())
                     + ") - see **Remarks**")
    if p["left_out"]:
        lines.append(f"- Not projected: {len(p['left_out'])} running work order(s) not booked in liable "
                     f"{p['window']} ({', '.join(p['left_out'][:5])}" + (" …" if len(p["left_out"]) > 5 else "") + ")")
    if len(labels) > 1:
        df["Total projected CTC"] = df[[f"Projected CTC {tl}" for tl in labels]].sum(axis=1)
        df["Total projected base sales"] = df[[f"Projected base sales {tl}" for tl in labels]].sum(axis=1)
        df["Total projected margin"] = df["Total projected base sales"] - df["Total projected CTC"]
        sort_col = "Total projected margin"
    else:
        sort_col = f"Projected margin {labels[0]}"
    t = set(re.findall(r"[a-z]+", question.lower()))
    group = "Department" if t & {"department", "departments", "dept", "departmentwise"} else \
        "Category" if t & {"category", "categories", "categorywise"} else None
    if group:
        num = [c for c in df.columns if str(c).startswith(("CTC ", "Base sales ", "Projected ", "Total projected",
                                                            "Pending months "))]
        df = df.groupby(group, as_index=False)[num].sum()
    df = df.sort_values(sort_col, ascending=False).reset_index(drop=True)
    all_c, all_b = sum(x[1] for x in totals), sum(x[2] for x in totals)
    tl_text = labels[0] if len(labels) == 1 else f"{labels[0]} to {labels[-1]}"
    pct = f" ({(all_b - all_c) * 100 / all_c:.2f}% on CTC)" if all_c else ""
    text = (f"**Projected for liable month {tl_text}{_proj_scope_text(f)} (as on {as_on}, data of liable "
            f"{p['window']}):** CTC {inr(all_c)} · Base sales value {inr(all_b)} · Margin {inr(all_b - all_c)}{pct}\n\n"
            + "\n".join(lines))
    notes.insert(0, f"Understood as: **Projection** for liable month {tl_text}, made on {as_on}: only the last "
                    f"{PROJ_LIABLE_MONTHS} liable months ({p['window']}) are used, each work order from its own last "
                    f"salary month in them"
                    + (f", {group}-wise" if group else ", work order-wise") + ".")
    notes.append(f"**Pending** = a salary month after the work order's last salary month that is already over "
                 f"(up to {cl}) and not paid yet - it is added to {labels[0]} (paid together).")
    notes.append("Rule (same as your Excel): 0 if the work order ends before the month or has no end date; "
                 "daily rate = value ÷ days in that month, or ÷ days it ran in the base month if it "
                 "started in that month; × days it runs in that month.")
    notes.append(f"⏱️ No AI · SQL: {p['t_sql']:.1f}s")
    # partial liable month -> projection too low
    cnt = run_query(f"SELECT YEAR([month_year])*100+MONTH([month_year]) AS p, COUNT(DISTINCT [Employee]) AS n "
                    f"FROM {T} WHERE [month_year] < ? GROUP BY YEAR([month_year])*100+MONTH([month_year]) "
                    f"ORDER BY p DESC", [datetime.date(targets[0][1], targets[0][0], 1)]).head(4)
    if len(cnt) >= 2:
        usual = float(cnt["n"].iloc[1:].median())
        if usual and cnt["n"].iloc[0] < 0.7 * usual:
            text = (f"⚠️ **Liable {cl} data looks incomplete** ({int(cnt['n'].iloc[0]):,} employees vs about "
                    f"{int(usual):,} in a normal month): more work orders will show as pending.\n\n" + text)
    intent = {"report": "projection", "base": f"data of liable {p['window']}", "targets": labels, "group": group,
              "department": f.get("department"), "category": f.get("category"), "excluded": excluded_text()}
    return {"text": text, "note": " ".join(notes), "df": df, "report": "projection", "sql": p["sql"],
            "params": p["params"], "intent": intent, "proj_totals": (all_c, all_b), "options": []}


# ---------------- Projection vs actual: what was projected for a liable month and what was really paid ----------------
PVA_PROJ = re.compile(r"\b(project\w*|forecast\w*|estimat\w*|expected|budget\w*|plan(?:ned)?)\b", re.IGNORECASE)
PVA_ACT = re.compile(r"\b(actual\w*|paid|real|booked|processed|happened)\b", re.IGNORECASE)
PVA_CMP = re.compile(r"\b(diff\w*|vs\.?|versus|compar\w*|variance|gap|against|mismatch|reconcil\w*|deviation|"
                     r"between|antar|farak|fark)\b", re.IGNORECASE)


def wants_proj_vs_actual(question):
    return bool(PVA_PROJ.search(question) and PVA_ACT.search(question) and PVA_CMP.search(question))


def proj_vs_actual_answer(question, mode, notes, t0):
    """'difference between projection and actual paid salary of liable month sep-26': the projection made on the
    day before the month (data up to the month before) against what was really booked in that liable month,
    work order by work order - new work orders, pending months paid, arrears, increases and decreases."""
    if not wants_proj_vs_actual(question):
        return None
    t1 = time.time()
    ps = parse_periods(question)
    if ps:
        tm, ty = ps[-1]
    else:
        lmy = pd.to_datetime(run_query(f"SELECT MAX([month_year]) AS p FROM {T}", [])["p"].iloc[0])
        if pd.isna(lmy):
            return clarify(question, "There is no salary data.")
        tm, ty = lmy.month, lmy.year
    tl = period_label(tm, ty)
    f = _scope(question)
    p = catchup_projection([(tm, ty)], f)
    if not p:
        return clarify(question, f"There is no salary data before liable month **{tl}**, so there is no projection "
                                 f"to compare with.")
    cl, as_on = period_label(*p["c"]), _month_end(p["c"][1], p["c"][0]).strftime("%d-%b-%y")
    proj = p["df"]
    proj = proj[(proj[f"Projected CTC {tl}"] != 0) | (proj["Pending months"] > 0)]
    where, params = build_where("month_year", [(tm, ty)], f)
    sql = f"""SELECT [WorkOrderNo], [Year], [Month], MAX([Category]) AS Category, MAX([Department]) AS Department,
       MAX([WOEndDate]) AS WOEndDate, COUNT(DISTINCT [Employee]) AS Employees, SUM([CTC]) AS CTC,
       SUM([_Base]) AS BaseSales, SUM([Arrears]) AS Arrears
FROM {base_rows('[Year]', '[Month]', where)} GROUP BY [WorkOrderNo], [Year], [Month]"""
    act = run_query(sql, params)
    if act.empty:
        return clarify(question, f"Nothing is booked in liable month **{tl}** yet, so there is nothing to compare.")
    act["sm"] = act["Year"] * 100 + act["Month"]
    cur = ty * 100 + tm
    known = p["known"]                                       # in the data before the month (projected or not)
    a_rows = []
    for wo, g in act.groupby("WorkOrderNo"):
        old = g[g["sm"] < cur].sort_values("sm")
        a_rows.append({"WorkOrderNo": wo, "aCategory": g["Category"].iloc[0], "aDepartment": g["Department"].iloc[0],
                       "aEnd": g["WOEndDate"].max(), "Actual CTC": g["CTC"].sum(), "Actual base sales": g["BaseSales"].sum(),
                       "Actual employees": int(g.loc[g["sm"] == cur, "Employees"].sum()) if (g["sm"] == cur).any()
                       else int(g["Employees"].max()),
                       "Salary months paid": " + ".join(period_label(int(s) % 100, int(s) // 100)
                                                        for s in sorted(g["sm"].unique())),
                       "Older months CTC": old["CTC"].sum(),
                       "OldMonths": [period_label(int(s) % 100, int(s) // 100) for s in old["sm"].unique()],
                       "Arrears": g["Arrears"].fillna(0).sum()})
    a = pd.DataFrame(a_rows)
    pc_ = f"Projected CTC {tl}"
    pb_ = f"Projected base sales {tl}"
    m = proj[["WorkOrderNo", "Category", "Department", "WOEndDate", "Employees", pc_, pb_, "Pending months",
              "Remarks"]].merge(a, on="WorkOrderNo", how="outer")
    m["Category"] = m["Category"].fillna(m["aCategory"])
    m["Department"] = m["Department"].fillna(m["aDepartment"])
    for c_ in (pc_, pb_, "Actual CTC", "Actual base sales", "Older months CTC", "Arrears", "Pending months"):
        m[c_] = m[c_].fillna(0)
    m["Remarks"] = m["Remarks"].fillna("")
    m["OldMonths"] = [v if isinstance(v, list) else [] for v in m["OldMonths"]]
    m["CTC difference"] = m["Actual CTC"] - m[pc_]
    m["Base sales difference"] = m["Actual base sales"] - m[pb_]
    status, why = [], []
    for _, r in m.iterrows():
        in_p, in_a = r[pc_] != 0 or r["Pending months"] > 0, pd.notna(r["Salary months paid"])
        bits = []
        if in_a and not in_p:
            s = ("New work order (not in projection)" if r["WorkOrderNo"] not in known
                 else "Not in projection (end date before the month / no salary for long)")
            if r["OldMonths"]:
                bits.append(f"{len(r['OldMonths']) + (1 if str(tl) in r['Salary months paid'] else 0)} salary months "
                            f"paid together: {r['Salary months paid']}")
        elif in_p and not in_a:
            end = pd.to_datetime(r["WOEndDate"], errors="coerce")
            s = "Projected but not paid"
            bits.append("work order ended " + _dmy(end) if pd.notna(end) and end.date() < datetime.date(ty, tm, 1)
                        else f"salary not booked in {tl} - still pending")
            if r["Remarks"]:
                bits.append(f"projection: {r['Remarks']}")
        else:
            d = r["CTC difference"]
            s = "Higher than projected" if d >= 1 else "Lower than projected" if d <= -1 else "Same as projected"
            if r["Remarks"]:
                bits.append(f"projection: {r['Remarks']}")
            if r["OldMonths"]:
                bits.append(f"older salary paid: {' + '.join(r['OldMonths'])} ({inr(r['Older months CTC'])})")
            e0, e1 = r["Employees"], r["Actual employees"]
            if pd.notna(e0) and e1 and int(e0) != int(e1):
                bits.append(f"employees {int(e0)} → {int(e1)}")
            if r["Pending months"] > 0 and not r["OldMonths"]:
                bits.append("pending months projected but not paid")
        if r["Arrears"]:
            bits.append(f"arrears {inr(r['Arrears'])}")
        status.append(s)
        why.append("; ".join(bits))
    m["Status"], m["Reason"] = status, why
    m = m.rename(columns={"Employees": "Projected employees"})
    cols = ["Status", "WorkOrderNo", "Department", "Category", pc_, "Actual CTC", "CTC difference", pb_,
            "Actual base sales", "Base sales difference", "Projected employees", "Actual employees",
            "Salary months paid", "Older months CTC", "Arrears", "Reason"]
    df = m[cols].copy()
    df["_abs"] = df["CTC difference"].abs()
    order = {"New work order (not in projection)": 0, "Not in projection (end date before the month / no salary for long)": 1,
             "Projected but not paid": 2, "Higher than projected": 3, "Lower than projected": 4, "Same as projected": 5}
    df = df.sort_values(["Status", "_abs"], key=lambda s_: s_.map(order) if s_.name == "Status" else -s_) \
        .drop(columns="_abs").reset_index(drop=True)
    P, A = df[pc_].sum(), df["Actual CTC"].sum()
    PB, AB = df[pb_].sum(), df["Actual base sales"].sum()

    def part(st_):
        g = df[df["Status"] == st_]
        return len(g), g["CTC difference"].sum(), g["Base sales difference"].sum()

    def pct(a_, b_):
        return f" ({(a_ - b_) * 100 / b_:+.2f}%)" if b_ else ""

    lines = [f"- **CTC:** projected {inr(P)} · actual {inr(A)} · difference **{inr(A - P)}**{pct(A, P)}",
             f"- **Base sales value:** projected {inr(PB)} · actual {inr(AB)} · difference **{inr(AB - PB)}**{pct(AB, PB)}",
             f"- **Margin:** projected {inr(PB - P)} · actual {inr(AB - A)} · difference **{inr((AB - A) - (PB - P))}**",
             "", f"**Where the CTC difference comes from:**"]
    for st_, lab in (("New work order (not in projection)", "New work orders (not in projection)"),
                     ("Not in projection (end date before the month / no salary for long)",
                      "Paid but not projected (end date before the month / no salary for long)"),
                     ("Projected but not paid", "Projected but not paid"),
                     ("Higher than projected", "Paid more than projected"),
                     ("Lower than projected", "Paid less than projected")):
        n, dc, db = part(st_)
        if n:
            extra = ""
            if st_.startswith("New"):
                g = df[(df["Status"] == st_) & (df["Older months CTC"] > 0)]
                if len(g):
                    extra = f" - {len(g)} of them paid older months together ({inr(g['Older months CTC'].sum())})"
            lines.append(f"- {lab}: {n} work order(s) · CTC {inr(dc)} · base sales {inr(db)}{extra}")
    n_same = part("Same as projected")[0]
    if n_same:
        lines.append(f"- Same as projected: {n_same} work order(s)")
    oldm = act[act["sm"] < cur]
    if not oldm.empty:
        by_m = oldm.groupby("sm")["CTC"].sum()
        lines.append(f"- Older salary months booked in {tl}: " + ", ".join(
            f"{period_label(int(s) % 100, int(s) // 100)} {inr(v)}" for s, v in by_m.items()))
    arr = act["Arrears"].fillna(0)
    if arr.sum():
        lines.append(f"- Arrears paid in {tl}: {inr(arr.sum())} in {act.loc[arr != 0, 'WorkOrderNo'].nunique()} "
                     f"work order(s)")
    for st_, lab in (("Higher than projected", "Biggest increases"), ("Lower than projected", "Biggest decreases")):
        g = df[df["Status"] == st_].head(3)
        if len(g):
            lines.append(f"- {lab}: " + "; ".join(f"{r['WorkOrderNo']} {inr(r['CTC difference'])}"
                                                  + (f" ({r['Reason']})" if r["Reason"] else "") for _, r in g.iterrows()))
    text = (f"**Liable month {tl}: projection vs actual{_proj_scope_text(f)}** (projection made on {as_on} from data "
            f"of liable {p['window']})\n\n" + "\n".join(lines))
    t = set(re.findall(r"[a-z]+", question.lower()))
    group = "Department" if t & {"department", "departments", "dept", "departmentwise", "client", "clients"} else \
        "Category" if t & {"category", "categories", "categorywise"} else None
    if group:                                               # same columns added up; work-order count per status
        num = [pc_, "Actual CTC", "CTC difference", pb_, "Actual base sales", "Base sales difference",
               "Older months CTC", "Arrears"]
        wo_n = df.pivot_table(index=group, columns="Status", values="WorkOrderNo", aggfunc="count", fill_value=0)
        wo_n.columns = [f"WOs: {c_}" for c_ in wo_n.columns]
        df = df.groupby(group)[num].sum().join(wo_n).reset_index()
        df = df.reindex(df["CTC difference"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    notes.insert(0, f"Understood as: **Projection vs actual** for liable month {tl}"
                    + (f", {group}-wise" if group else ", work order-wise") + ". Projection = what was known on "
                    f"{as_on} (liable {p['window']} only, each work order from its last salary month, pending months "
                    f"added). Actual = everything booked in liable month {tl} (month_year).")
    notes.append(f"⏱️ No AI · {time.time() - t1:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": df, "report": "proj_vs_actual", "sql": sql,
            "params": params, "intent": {"report": "proj_vs_actual", "month": tl}}


# ---------------- Projection vs projection: why the projection of one liable month differs from another ----------------
def wants_proj_compare(question):
    return bool(PVA_PROJ.search(question) and (PVA_CMP.search(question) or WHY_RX.search(question))
                and not PVA_ACT.search(question) and not BASE_HINT.search(question)
                and len(set(parse_periods(question))) >= 2)


def proj_compare_answer(question, mode, notes, t0):
    """'why there is a diff in projection of sep-26 and oct-26': each month's projection as made on the day before it
    (Sep-26 on 31-Aug, Oct-26 on 30-Sep), work order by work order - expired, new, pending salary added, part month,
    employees / CTC changed, end date extended."""
    if not wants_proj_compare(question):
        return None
    t1 = time.time()
    ps = sorted(set(parse_periods(question)), key=lambda p_: (p_[1], p_[0]))
    (am, ay), (bm, by) = ps[0], ps[-1]
    al, bl = period_label(am, ay), period_label(bm, by)
    f = _scope(question)
    pa, pb = catchup_projection([(am, ay)], f), catchup_projection([(bm, by)], f)
    if not pa or not pb or pa["df"].empty or pb["df"].empty:
        return clarify(question, f"There is no salary data before liable month **{al if not pa else bl}** to project from.")
    b_start = datetime.date(by, bm, 1)

    def side(p, lbl, k):
        d = p["df"][["WorkOrderNo", "Category", "Department", "WOStartDate", "WOEndDate", "Employees", "Last salary month",
                     "CTC in last salary month", "Base sales in last salary month", "Pending months",
                     f"Pending months CTC (in {lbl})",
                     f"Days in {lbl}", f"Projected CTC {lbl}", f"Projected base sales {lbl}", "Remarks", "_PayMonths"]]
        return d.rename(columns={c_: f"{k}{c_}" for c_ in d.columns if c_ != "WorkOrderNo"})

    def pay_txt(months, lbl, later):
        """'in Sep-2026 we pay 2 months salary: Jul-2026 + Aug-2026' - only when more than one month is paid."""
        months = months if isinstance(months, list) else []
        if len(months) < 2:
            return None
        return f"in {lbl} we {'need to ' if later else ''}pay {len(months)} months salary: {' + '.join(months)}"

    m = side(pa, al, "A ").merge(side(pb, bl, "B "), on="WorkOrderNo", how="outer")
    ca, cb, ba, bb = f"A Projected CTC {al}", f"B Projected CTC {bl}", f"A Projected base sales {al}", \
        f"B Projected base sales {bl}"
    for c_ in (ca, cb, ba, bb, "A Pending months", "B Pending months", f"A Pending months CTC (in {al})",
               f"B Pending months CTC (in {bl})"):
        m[c_] = m[c_].fillna(0)
    m = m[(m[ca] != 0) | (m[cb] != 0)]
    m["Category"] = m["A Category"].fillna(m["B Category"])
    m["Department"] = m["A Department"].fillna(m["B Department"])
    # work orders not known when the first projection was made: when were they first booked, for which salary months
    fresh = [w_ for w_ in m.loc[(m[ca] == 0) & (m[cb] != 0), "WorkOrderNo"] if w_ not in pa["known"]]
    first_bk = {}
    if fresh:
        inn = ",".join("?" * len(fresh))
        fb = run_query(f"""WITH f AS (SELECT [WorkOrderNo] AS W, MIN([month_year]) AS F FROM {T}
               WHERE [WorkOrderNo] IN ({inn}) GROUP BY [WorkOrderNo])
SELECT f.W, f.F, t.[Year]*100+t.[Month] AS sm FROM f JOIN {T} t ON t.[WorkOrderNo] = f.W
 AND YEAR(t.[month_year]) = YEAR(f.F) AND MONTH(t.[month_year]) = MONTH(f.F)
GROUP BY f.W, f.F, t.[Year]*100+t.[Month]""", list(fresh))
        for w_, g in fb.groupby("W"):
            f_ = pd.to_datetime(g["F"].iloc[0])
            first_bk[w_] = (period_label(f_.month, f_.year),
                            [period_label(int(s) % 100, int(s) // 100) for s in sorted(g["sm"])])
    # ... and already paid for salary month al: a normal work order - its al salary is the al figure
    m[f"{al} figure"] = "Projection"
    A0, AB0 = m[ca].sum(), m[ba].sum()
    used = m["WorkOrderNo"].isin(fresh) & (m["B Last salary month"] == al)
    m.loc[used, ca] = m.loc[used, "B CTC in last salary month"].fillna(0)
    m.loc[used, ba] = m.loc[used, "B Base sales in last salary month"].fillna(0)
    m.loc[used, f"{al} figure"] = f"{al} salary (first booked after the projection)"
    m["CTC difference"] = m[cb] - m[ca]
    m["Base sales difference"] = m[bb] - m[ba]
    asa = _month_end(*reversed(pa["c"])).strftime("%d-%b-%y")
    asb = _month_end(*reversed(pb["c"])).strftime("%d-%b-%y")
    status, why = [], []
    for _, r in m.iterrows():
        in_a, in_b, bits = r[ca] != 0, r[cb] != 0, []
        end_a = pd.to_datetime(r["A WOEndDate"], errors="coerce")
        end_b = pd.to_datetime(r["B WOEndDate"], errors="coerce")
        end = end_b if pd.notna(end_b) else end_a
        pend_a, pend_b = r["A Pending months"], r["B Pending months"]
        if in_a and not in_b:
            if pd.notna(end) and end.date() < b_start:
                s = "Expired"
                bits.append(f"work order ended {_dmy(end)} - nothing to pay in {bl}")
            else:
                s = f"Not in {bl} projection"
                bits.append(f"no salary booked in the {PROJ_LIABLE_MONTHS} liable months before {bl}"
                            if pd.isna(r["B CTC in last salary month"]) else
                            ("no end date" if pd.isna(end) else f"ends {_dmy(end)}"))
            t_ = pay_txt(r["A _PayMonths"], al, False)
            if t_:
                bits.append(t_)
        elif in_b and not in_a:
            if r["WorkOrderNo"] not in pa["known"]:
                fl, sms = first_bk.get(r["WorkOrderNo"], ("?", []))
                s = f"First paid in liable {fl}"
                bits.append(f"first booked in liable {fl}" + (f" ({len(sms)} months salary paid together: "
                                                              f"{' + '.join(sms)})" if len(sms) > 1 else
                                                              f" (salary {sms[0]})" if sms else "")
                            + f" - not known on {asa}, so not in the {al} projection; a normal running work order "
                              f"for {bl}")
            else:
                s = f"Back in {bl} projection"
                if pd.notna(end_a) and pd.notna(end_b) and end_b > end_a:
                    bits.append(f"end date extended {_dmy(end_a)} → {_dmy(end_b)}")
                elif pd.isna(r["A CTC in last salary month"]):
                    bits.append(f"no salary booked in the {PROJ_LIABLE_MONTHS} liable months before {al}")
            t_ = pay_txt(r["B _PayMonths"], bl, True)
            if t_:
                bits.append(t_)
        else:
            d = r["CTC difference"]
            s = f"Higher in {bl}" if d >= 1 else f"Lower in {bl}" if d <= -1 else "Same"
            if r[f"{al} figure"] != "Projection":
                fl, sms = first_bk.get(r["WorkOrderNo"], ("?", []))
                bits.append(f"first booked in liable {fl}, after the {al} projection was made ({asa}) - "
                            f"its {al} salary is compared"
                            + (f"; in {fl} we paid {len(sms)} months salary: {' + '.join(sms)}" if len(sms) > 1 else ""))
                st_d = pd.to_datetime(r["B WOStartDate"], errors="coerce")
                if pd.notna(st_d) and (st_d.year, st_d.month) == (ay, am) and st_d.day > 1:
                    bits.append(f"started {_dmy(st_d)} - part month in {al}, full month in {bl}")
            for k_, lbl_, later in (("A", al, False), ("B", bl, True)):
                t_ = pay_txt(r[f"{k_} _PayMonths"], lbl_, later)
                if t_:
                    bits.append(t_ + f" ({inr(r[f'{k_} Pending months CTC (in {lbl_})'])} for the earlier month(s))")
            da, db = r[f"A Days in {al}"], r[f"B Days in {bl}"]
            if pd.notna(db) and 0 < db < _month_end(by, bm).day:
                bits.append(f"part month in {bl}: {int(db)} days (ends {_dmy(end)})")
            if pd.notna(da) and 0 < da < _month_end(ay, am).day:
                bits.append(f"part month in {al}: {int(da)} days")
            if pd.notna(end_a) and pd.notna(end_b) and end_b > end_a:
                bits.append(f"end date extended {_dmy(end_a)} → {_dmy(end_b)}")
            e0, e1 = r["A Employees"], r["B Employees"]
            if pd.notna(e0) and pd.notna(e1) and int(e0) != int(e1):
                bits.append(f"employees {int(e0)} → {int(e1)}")
            c0, c1 = r["A CTC in last salary month"], r["B CTC in last salary month"]
            if pd.notna(c0) and pd.notna(c1) and abs(c1 - c0) >= 1:
                bits.append(f"monthly CTC {inr(c0)} ({r['A Last salary month']}) → {inr(c1)} ({r['B Last salary month']})")
        status.append(s)
        why.append("; ".join(bits))
    m["Status"], m["Reason"] = status, why
    m = m.rename(columns={"A Employees": f"Employees ({al})", "B Employees": f"Employees ({bl})",
                          "A WOEndDate": f"WOEndDate (as on {al})", "B WOEndDate": f"WOEndDate (as on {bl})",
                          ca: f"CTC {al}", cb: f"Projected CTC {bl}",
                          ba: f"Base sales {al}", bb: f"Projected base sales {bl}"})
    ca, cb, ba, bb = f"CTC {al}", f"Projected CTC {bl}", f"Base sales {al}", f"Projected base sales {bl}"
    cols = ["Status", "WorkOrderNo", "Department", "Category", f"{al} figure", ca, cb, "CTC difference", ba, bb,
            "Base sales difference", f"Employees ({al})", f"Employees ({bl})", f"WOEndDate (as on {al})",
            f"WOEndDate (as on {bl})", "Reason"]
    df = m[cols].copy()
    order = {"Expired": 0, f"Not in {bl} projection": 1, f"Back in {bl} projection": 3,
             f"Lower in {bl}": 4, f"Higher in {bl}": 5, "Same": 6}
    is_first = df["Status"].str.startswith("First paid in liable")
    df["_o"], df["_abs"] = df["Status"].map(order).where(~is_first, 2), -df["CTC difference"].abs()
    df = df.sort_values(["_o", "_abs"]).drop(columns=["_o", "_abs"]).reset_index(drop=True)
    is_first = df["Status"].str.startswith("First paid in liable")
    A, B, AB_, BB_ = df[ca].sum(), df[cb].sum(), df[ba].sum(), df[bb].sum()

    def pct(x, y):
        return f" ({(x - y) * 100 / y:+.2f}%)" if y else ""

    n_used = int(used.sum())
    lines = [f"- **{al}** (as on {asa}, data of liable {pa['window']}): CTC {inr(A)} · base sales {inr(AB_)} · "
             f"margin {inr(AB_ - A)}" + (f" = projection {inr(A0)} + {n_used} work order(s) first booked in liable "
                                         f"{al} at their {al} salary {inr(A - A0)}" if n_used else ""),
             f"- **{bl}** (as on {asb}, data of liable {pb['window']}): CTC {inr(B)} · base sales {inr(BB_)} · "
             f"margin {inr(BB_ - B)}",
             f"- **Difference:** CTC **{inr(B - A)}**{pct(B, A)} · base sales **{inr(BB_ - AB_)}**{pct(BB_, AB_)}",
             "", "**Where the CTC difference comes from:**"]
    first_lab = (f"Work orders first booked after {asa} (not known for the {al} projection, normal running work "
                 f"orders for {bl})")
    for st_, lab in (("Expired", f"Expired before {bl} (paid in {al}, nothing in {bl})"),
                     (f"Not in {bl} projection", f"In {al} but not projected for {bl}"),
                     (None, first_lab),
                     (f"Back in {bl} projection", f"Not in {al} but projected for {bl}"),
                     (f"Lower in {bl}", f"Lower in {bl}"), (f"Higher in {bl}", f"Higher in {bl}")):
        g = df[is_first] if st_ is None else df[df["Status"] == st_]
        if len(g):
            lines.append(f"- {lab}: {len(g)} work order(s) · CTC {inr(g['CTC difference'].sum())} · base sales "
                         f"{inr(g['Base sales difference'].sum())}")
    for p, lbl, later in ((pa, al, False), (pb, bl, True)):    # more than one salary month paid in the month
        many = p["df"][p["df"]["_PayMonths"].map(len) > 1]
        if len(many):
            by_n = many["_PayMonths"].map(len).value_counts().sort_index()
            lines.append(f"- In {lbl} we {'need to ' if later else ''}pay more than one month's salary in "
                         f"{len(many)} work order(s) (" + ", ".join(f"{int(v)} with {k} months" for k, v in by_n.items())
                         + f"): {inr(many[f'Pending months CTC (in {lbl})'].sum())} for the earlier months")
    n_same = int((df["Status"] == "Same").sum())
    if n_same:
        lines.append(f"- Same in both: {n_same} work order(s)")
    for st_, lab in (("Expired", "expired"), (f"Lower in {bl}", f"lower in {bl}"), (f"Higher in {bl}", f"higher in {bl}"),
                     (None, f"first booked after {asa}")):
        g = (df[is_first] if st_ is None else df[df["Status"] == st_]).head(3)
        if len(g):
            lines.append(f"- Biggest {lab}: " + "; ".join(
                f"{r['WorkOrderNo']} {inr(r['CTC difference'])}" + (f" ({r['Reason']})" if r["Reason"] else "")
                for _, r in g.iterrows()))
    text = (f"**Projection {al} vs {bl}{_proj_scope_text(f)}: why they differ**\n\n" + "\n".join(lines))
    t = set(re.findall(r"[a-z]+", question.lower()))
    group = "Department" if t & {"department", "departments", "dept", "departmentwise", "client", "clients"} else \
        "Category" if t & {"category", "categories", "categorywise"} else None
    if group:
        num = [ca, cb, "CTC difference", ba, bb, "Base sales difference"]
        wo_n = df.pivot_table(index=group, columns="Status", values="WorkOrderNo", aggfunc="count", fill_value=0)
        wo_n.columns = [f"WOs: {c_}" for c_ in wo_n.columns]
        df = df.groupby(group)[num].sum().join(wo_n).reset_index()
        df = df.reindex(df["CTC difference"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    notes.insert(0, f"Understood as: **Projection vs projection** - {al} as made on {asa} and {bl} as made on {asb}"
                    + (f", {group}-wise" if group else ", work order-wise") + ". Each projection uses only the "
                    f"{PROJ_LIABLE_MONTHS} liable months before its own month, each work order from its last salary "
                    f"month, pending months added to the month.")
    notes.append(f"In *projection of {al.lower()} to {bl.lower()}* both months are projected as on {asa}, so the "
                 f"{bl} figure there differs from the {bl} projection made on {asb}.")
    notes.append(f"⏱️ No AI · {time.time() - t1:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": df, "report": "proj_compare", "sql": pb["sql"],
            "params": pb["params"], "intent": {"report": "proj_compare", "months": [al, bl]}}


# ------------------------------ Proposed TDS (FY 2026-27) ------------------------------
# The calculation lives in ONE place: the SQL procedure below (file "Payrollall for Employee wise
# TDS calculation F&F v16.sql"). The chatbot only runs it - fixed text, nothing from the question
# goes into the SQL - and then picks the employee(s) you asked about.
#   Gross = CTC - PF_Employer (arrears NOT added - their months have their own rows)
#   TDSOpt 2 = New regime, 1 = Old regime, blank = New regime (the default regime)
#   MonthlyTDS = (annual tax with cess - TDS already deducted this FY) / months left, never below 0
TDS_PROC = "dbo.usp_CalculateTDS_FY2627"
TDS_FY_TEXT = "FY 2026-27"

TDS_ASK = re.compile(
    r"\b(propos\w*|project\w*|estimat\w*|expected|remaining|balance|pending|future|calculat\w*|comput\w*)"
    r"\s+(?:monthly\s+|annual\s+|yearly\s+)?(?:tds|income\s*-?\s*tax|tax)\b"
    r"|\btds\s+(?:propos\w*|project\w*|calculat\w*|calc|plan|remaining|pending|balance|left|"
    r"to\s+(?:be\s+)?deduct\w*)\b"
    r"|\b(?:annual|yearly)\s+(?:tds|tax)\b|\b(?:old|new)\s+regime\b|\btds\s*opt\b|\bregime\s+wise\b",
    re.IGNORECASE)
TDS_MONTHLY = re.compile(r"\bmonthly\s+tds\b|\btds\s+(?:per|every|each)\s+month\b", re.IGNORECASE)
TDS_STOP = {
    "tds", "proposed", "propose", "proposal", "projected", "projection", "estimated", "estimate", "expected",
    "remaining", "balance", "pending", "future", "calculate", "calculated", "calculation", "calc", "compute",
    "computed", "monthly", "month", "months", "annual", "yearly", "tax", "income", "regime", "regimes", "old",
    "new", "opt", "option", "tdsopt", "of", "for", "the", "a", "an", "me", "my", "give", "show", "get", "tell",
    "what", "whats", "is", "are", "will", "be", "employee", "employees", "emp", "code", "name", "all", "every",
    "each", "list", "data", "details", "detail", "please", "pls", "kindly", "and", "to", "deduct", "deducted",
    "deduction", "fy", "financial", "year", "in", "on", "with", "how", "much", "salary", "wise", "employeewise",
    "per", "by", "can", "you", "which", "who", "whose", "should", "this", "next", "f", "ff", "fnf", "settlement",
    "final", "full", "report", "table", "sheet", "csv", "excel", "download", "his", "her", "their", "about",
    "from", "now", "mr", "mrs", "ms", "shri", "sir", "madam", "ki", "ka", "ke", "ko", "hai", "batao", "dikhao",
    "do", "de", "kitna", "kitni", "mujhe", "sab", "sabhi", "total", "both", "compare", "comparison", "vs",
    "versus", "under", "as", "at", "cut", "rs", "amount", "value", "current", "actual", "paid", "already",
    "left", "rest", "till", "upto", "up", "its", "it", "or", "if", "remain", "need", "needs", "deducting",
    "highest", "lowest", "top", "bottom", "max", "min", "most", "least", "only", "wala", "wale", "wali",
} | set(MONTH_WORDS)


def wants_proposed_tds(question):
    """'proposed tds of 1023', 'monthly tds for amit', 'old regime employees tds'."""
    if TDS_ASK.search(question):
        return True
    q = re.sub(r"\bfy\s*\d{2,4}\s*-\s*\d{2,4}\b", " ", question, flags=re.IGNORECASE)
    return bool(TDS_MONTHLY.search(q)) and not parse_periods(q)   # "monthly tds for aug-26" = actual TDS cut


@st.cache_data(ttl=300, show_spinner=False)
def tds_table():
    """Runs the TDS procedure once (fixed text, no user input) and keeps the result 5 minutes."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(f"SET NOCOUNT ON; EXEC {TDS_PROC};")
        while cur.description is None and cur.nextset():     # skip to the result table
            pass
        cols = [d[0] for d in cur.description or []]
        df = pd.DataFrame.from_records([tuple(r) for r in cur.fetchall()], columns=cols)
    finally:
        try:
            conn.rollback()
        finally:
            conn.close()
    for c in df.columns:
        if df[c].dtype == object and df[c].map(lambda v: isinstance(v, Decimal)).any():
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # works with the older v14 procedure too: pick the employee's own regime here
    opt = pd.to_numeric(df.get("TDSOpt"), errors="coerce")
    old = opt == 1
    if "Regime" not in df:
        df["Regime"] = old.map({True: "Old", False: "New"})
    if "AnnualTax" not in df:
        df["AnnualTax"] = df["AnnualTaxWithCess_Old"].where(old, df["AnnualTaxWithCess_New"])
    if "RemainingTDS" not in df:
        df["RemainingTDS"] = df["RemainingTDS_Old"].where(old, df["RemainingTDS_New"])
    if "MonthlyTDS" not in df:
        m = df["MonthlyTDS_Old"].where(old, df["MonthlyTDS_New"])
        df["MonthlyTDS"] = m.clip(lower=0).round(0)
    df["EmployeeName"] = df["EmployeeName"].fillna("").astype(str).str.strip()
    return df


def _tds_pick(df, question):
    """Finds the employee(s) asked about: by code, else by name words. Returns (rows, what_was_searched)."""
    t = question.lower()
    t = re.sub(r"\bfy\s*\d{2,4}\s*-\s*\d{2,4}\b", " ", t)
    codes = df["EmployeeCode"].astype(str).str.strip()
    nums = [n for n in re.findall(r"\b\d{2,}\b", t) if n in set(codes)]
    if nums:
        return df[codes.isin(nums)], ", ".join(nums)
    words = [w for w in re.findall(r"[a-z]+", t) if w not in TDS_STOP and len(w) >= 2]
    if not words:
        return df, None
    name_words = df["EmployeeName"].str.lower().map(lambda s: re.findall(r"[a-z]+", s))
    hit = name_words.map(lambda nw: all(any(x.startswith(w) for x in nw) for w in words))
    if not hit.any() and len(words) > 1:                      # e.g. extra word in the question
        hit = name_words.map(lambda nw: sum(any(x.startswith(w) for x in nw) for w in words) >= len(words) - 1)
    return (df[hit] if hit.any() else df.iloc[0:0]), " ".join(words)


def proposed_tds_answer(question, notes, t0):
    try:
        full = tds_table()
    except Exception as e:
        return {"text": "⚠️ I couldn't run the TDS calculation in SQL Server.\n\n"
                        f"Please run **Payrollall for Employee wise TDS calculation F&F v16.sql** once in SSMS "
                        f"(it creates `{TDS_PROC}`), then ask again.\n\nError: `{e}`",
                "note": "Understood as: **Proposed TDS**.", "df": None, "report": "proposed_tds",
                "sql": f"EXEC {TDS_PROC};", "params": [], "intent": {}}
    t_sql = time.time() - t0
    if _row_where():                                    # the TDS procedure reads all rows: keep only allowed employees
        ok = set(run_query(f"SELECT DISTINCT [Employee] AS e FROM {T}", [])["e"].astype(str))
        full = full[full["EmployeeCode"].astype(str).isin(ok)].reset_index(drop=True)
    df, searched = _tds_pick(full, question)
    ql = question.lower()
    has_old, has_new = bool(re.search(r"\bold\b", ql)), bool(re.search(r"\bnew\b", ql))
    if "regime" not in ql or (has_old and has_new):
        pass                                                  # no regime filter / comparing both
    elif has_old:
        df = df[df["Regime"] == "Old"]
        notes.append("Only **Old regime** employees (TDSOpt 1).")
    elif has_new:
        df = df[df["Regime"] == "New"]
        notes.append("Only **New regime** employees (TDSOpt 2 or blank).")
    if df.empty:
        if searched:
            import difflib
            names = list(dict.fromkeys(full["EmployeeName"]))
            close = difflib.get_close_matches(searched.title(), names, n=5, cutoff=0.5) or \
                [n for n in names if any(w[:3] in n.lower() for w in searched.split())][:5]
            if close:
                return clarify(question, f"I couldn't find an employee **{searched}** in the TDS list. Did you mean:",
                               [f"proposed tds of {n}" for n in close])
            return clarify(question, f"I couldn't find an employee **{searched}** in the {TDS_FY_TEXT} TDS list. "
                                     "Try the employee code, e.g. *proposed tds of 1023*.",
                           ["proposed tds of all employees"])
        return {"text": "No employees found for that.", "note": " ".join(notes), "df": None,
                "report": "proposed_tds", "sql": f"EXEC {TDS_PROC};", "params": [], "intent": {}}

    df = df.sort_values("MonthlyTDS", ascending=False, na_position="last").reset_index(drop=True)
    both = bool(re.search(r"\bboth\b|\bcompar\w*|\bvs\b|\bversus\b|\bold\s+(?:and|&|or|vs)\s+new\b|"
                          r"\bnew\s+(?:and|&|or|vs)\s+old\b", ql)) or len(df) == 1
    cols = {"EmployeeCode": "Employee", "EmployeeName": "Name", "Regime": "Regime", "TDSOpt": "TDSOpt",
            "WorkOrderNo": "WorkOrderNo", "MonthlySalary": "Monthly gross (CTC − PF employer)",
            "ActualSalaryPaid_FY2627": "Salary paid so far", "ProposedSalary_FY2627": "Proposed salary (rest of FY)",
            "TotalSalary_FY2627": "Total salary FY", "TDSAlreadyPaid_FY2627": "TDS already deducted",
            "AnnualTax": "Annual tax (with cess)", "RemainingTDS": "Remaining TDS", "PendingMonths": "Months left",
            "MonthlyTDS": "Monthly TDS"}
    if both:
        cols.update({"MonthlyTDS_New": "Monthly TDS if New regime", "MonthlyTDS_Old": "Monthly TDS if Old regime"})
    out = df[[c for c in cols if c in df]].rename(columns=cols)

    if len(df) == 1:
        r = df.iloc[0]
        rem = 0.0 if pd.isna(r["RemainingTDS"]) else float(r["RemainingTDS"])
        months = 0 if pd.isna(r["PendingMonths"]) else int(r["PendingMonths"])
        opt = "blank → default" if pd.isna(r.get("TDSOpt")) else f"{int(r['TDSOpt'])}"
        other = "Old" if r["Regime"] == "New" else "New"
        other_m = r.get(f"MonthlyTDS_{other}")
        lines = [f"**Monthly TDS for {r['EmployeeName']} (emp {r['EmployeeCode']}): {inr(r['MonthlyTDS'])}** "
                 f"· {r['Regime']} regime (TDSOpt {opt})",
                 f"- Total gross {TDS_FY_TEXT} (CTC − PF employer): {inr(r['TotalSalary_FY2627'])} "
                 f"(paid {inr(r['ActualSalaryPaid_FY2627'])} + proposed {inr(r['ProposedSalary_FY2627'])})",
                 f"- Annual tax with cess: {inr(r['AnnualTax'])} · already deducted: {inr(r['TDSAlreadyPaid_FY2627'])}",
                 (f"- Remaining {inr(rem)} ÷ {months} month(s) left = {inr(r['MonthlyTDS'])} a month"
                  if months and rem >= 0 else
                  f"- Already deducted {inr(-rem)} **more** than the yearly tax → no more TDS needed"
                  if rem < 0 else "- No months left in the year (or no salary expected) → no monthly figure")]
        if other_m is not None and not pd.isna(other_m):
            lines.append(f"- If {other} regime: {inr(max(float(other_m), 0))} a month")
        text = "\n".join(lines)
    else:
        top = df.iloc[0]
        n_old = int((df["Regime"] == "Old").sum())
        text = (f"**Proposed monthly TDS for {len(df):,} employees ({TDS_FY_TEXT}): {inr(df['MonthlyTDS'].sum())} "
                f"a month in total** · New regime: {len(df) - n_old:,} · Old regime: {n_old:,}\n"
                f"- Highest: {top['EmployeeName']} (emp {top['EmployeeCode']}) {inr(top['MonthlyTDS'])} a month\n"
                f"- Remaining TDS for the year: {inr(df['RemainingTDS'].clip(lower=0).sum())}")
        if searched:
            text = f"Employees matching **{searched}**:\n\n" + text
    notes = [f"Understood as: **Proposed TDS {TDS_FY_TEXT}** (each employee's own regime from TDSOpt: "
             "2 = New, 1 = Old, blank = New)."] + notes + [
        "Gross = **CTC − PF employer** each month (arrears not added). "
        "Annual tax = slab tax on (total salary − standard deduction ₹75,000 new / ₹50,000 old) + 4% cess; "
        "old regime does not include 80C/HRA declarations.",
        f"⏱️ No AI · SQL: {t_sql:.1f}s"]
    return {"text": text, "note": " ".join(notes), "df": out, "report": "proposed_tds",
            "sql": f"EXEC {TDS_PROC};", "params": [],
            "intent": {"report": "proposed_tds", "employee": searched, "rows": len(out)}}


# ------------------------------ Salary run (all employees, one month) ------------------------------
# "salary run for oct-26": one row per employee - the salary to pay for that month AND the TDS to cut.
#   Base = latest SALARY month (Year/Month) before it. Same day rule as the projection: nothing if the work
#   order ends before the month (or has no end date) or the employee resigned before it; part month if it
#   ends / they leave during it; a work order that started in the base month is scaled up to a full month.
#   Gross = base month gross WITHOUT arrears x that rule; PF / ESIC (employee) scale the same way;
#   PT and other deduction as in the base month; Take-home = Gross - PF - ESIC - PT - deduction - TDS
#   TDS, for the FY of the payroll month (Apr-27 -> FY 2027-28, worked out here - the SQL procedure is FY 2026-27 only):
#     total = salary paid this FY (CTC - PF employer, procedure's month rule) + gross of every month still to pay
#     TDS = (annual tax, own regime - TDS already deducted this FY) / months from the payroll month to March with salary
SALARY_RUN = re.compile(
    r"\bsalary\s+(?:run|sheet|payout|pay\s*-?\s*out|register|disburs\w*|payable|to\s+(?:be\s+)?pa(?:y|id))\b"
    r"|\b(?:pay\s*-?\s*out|payable|disburs\w*)\s+(?:salary|list|sheet)\b|\bpayroll\s+run\b"
    r"|\b(?:going\s+to|will|to\s+be|need\s+to)\s+pa(?:y|id)\b.*\bsalar", re.IGNORECASE)


def wants_salary_run(question):
    """'salary run for oct-26', 'salary to pay for oct-26', 'oct-26 payout list'."""
    return bool(SALARY_RUN.search(question))


def _emp_key(v):
    try:
        f_ = float(v)
        return str(int(f_)) if f_.is_integer() else str(v).strip()
    except (TypeError, ValueError):
        return str(v).strip()


def _run_days(r, by, bm, ty, tm):
    """(days paid in the target month, share of the base month amount, status) for one work-order row."""
    base_end, t_start, t_end = _month_end(by, bm), datetime.date(ty, tm, 1), _month_end(ty, tm)
    start = pd.to_datetime(r["WOStartDate"], errors="coerce")
    end = pd.to_datetime(r["WOEndDate"], errors="coerce")
    left = pd.to_datetime(r["resigndate"], errors="coerce")
    if pd.isna(end):
        return 0, 0.0, "Work order has no end date"
    if end.date() < t_start:
        return 0, 0.0, f"Work order ended {end:%d-%b-%y}"
    if pd.notna(left) and left.date() < t_start:
        return 0, 0.0, f"Resigned {left:%d-%b-%y}"
    last = min(t_end, end.date(), left.date() if pd.notna(left) else t_end)
    active = (last - t_start).days + 1
    days = (base_end - start.date()).days + 1 if pd.notna(start) else base_end.day + 1
    div = max(days, 1) if days < base_end.day else t_end.day
    status = ("Leaves " if pd.notna(left) and left.date() <= t_end else "Part month to ") + f"{last:%d-%b-%y}" \
        if active < t_end.day else ""
    return active, active / div, status


def salary_run_answer(question, notes, t0):
    latest = run_query(f"SELECT MAX([Year]*100+[Month]) AS p FROM {T}", [])["p"].iloc[0]
    if pd.isna(latest):
        return clarify(question, "There is no salary data to work from.")
    lm, ly = int(latest) % 100, int(latest) // 100
    periods = parse_periods(question)
    tm, ty = periods[0] if periods else ((1, ly + 1) if lm == 12 else (lm + 1, ly))
    tl = period_label(tm, ty)
    if ty * 100 + tm <= int(latest):                         # already paid: the real figures are in the data
        return clarify(question, f"**{tl}** salary is already in the data, so it is not a month still to pay. "
                                 f"Ask for the actual figures instead:",
                       [f"employee wise gross, tds and take home for {MON[tm - 1].lower()}-{str(ty)[-2:]}"])
    bm, by = lm, ly
    bl = period_label(bm, by)
    fy_start = ty if tm >= 4 else ty - 1                   # the FY comes from the month paid: Apr-27 -> FY 2027-28
    fy_text = f"FY {fy_start}-{str(fy_start + 1)[-2:]}"
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where("salary", [(bm, by)], f)       # base month = SALARY month (Year/Month)
    # Like the TDS procedure, the tax needs every figure even if a column is hidden for this login - so read the
    # table with only the login's row limits; hidden columns are removed from the result (by name) afterwards.
    src = f"(SELECT * FROM {RAW_T}{_row_where()}) AS {RAW_T}" if _row_where() else RAW_T
    sql = f"""SELECT [Employee], MAX([Name]) AS Name, MAX([Department]) AS Department, [WorkOrderNo],
       MAX([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate, MAX([resigndate]) AS resigndate,
       MAX([joindate]) AS joindate,
       SUM(ISNULL([GrossAmount],0) - ISNULL([Arrears],0)) AS Gross, SUM(ISNULL([PF_Employee],0)) AS PF,
       SUM(ISNULL([ESIC_Employee],0)) AS ESIC, MAX(ISNULL([PT],0)) AS PT, SUM(ISNULL([Deduction],0)) AS Deduction,
       SUM(ISNULL([IncomeTax],0)) AS TDS, MAX([tdsopt]) AS TDSOpt
FROM {src} {where} GROUP BY [Employee], [WorkOrderNo] ORDER BY [Employee]"""
    # FY so far, same rule as the TDS procedure: gross = CTC - PF employer; a row counts in the month it was
    # paid only if that is more than a month late (arrears), else in its salary month
    fy_sql = f"""SELECT [Employee], SUM(ISNULL([CTC],0) - ISNULL([PF_Employer],0)) AS FYPaid,
       SUM(ISNULL([IncomeTax],0)) AS FYTDS
FROM (SELECT [Employee], [CTC], [PF_Employer], [IncomeTax],
             CASE WHEN DATEDIFF(MONTH, DATEFROMPARTS([Year],[Month],1),
                                COALESCE(DATEADD(MONTH, 1, [month_year]), DATEFROMPARTS([Year],[Month],1))) > 1
                  THEN COALESCE(DATEADD(MONTH, 1, [month_year]), DATEFROMPARTS([Year],[Month],1))
                  ELSE DATEFROMPARTS([Year],[Month],1) END AS Eff
      FROM {src} WHERE [CTC] IS NOT NULL) s
WHERE s.Eff >= ? AND s.Eff <= ? GROUP BY [Employee]"""
    t1 = time.time()
    raw = run_query(sql, params)
    fy = run_query(fy_sql, [datetime.date(fy_start, 4, 1), datetime.date(fy_start + 1, 3, 31)])
    t_sql = time.time() - t1
    if raw.empty:
        return clarify(question, f"No salary data found for **{bl}** to work from.")
    for c in ("Gross", "PF", "ESIC", "PT", "Deduction", "TDS"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0)
    for c in ("WOStartDate", "WOEndDate", "joindate", "resigndate"):
        raw[c] = pd.to_datetime(raw[c], errors="coerce")
    rows = [_run_days(r, by, bm, ty, tm) for _, r in raw.iterrows()]
    raw["Days"], raw["Share"], raw["Status"] = [x[0] for x in rows], [x[1] for x in rows], [x[2] for x in rows]
    paid = raw["Days"] > 0
    for c in ("Gross", "PF", "ESIC"):
        raw["T_" + c] = (raw[c] * raw["Share"]).round(0)
    raw["T_PT"] = raw["PT"].where(paid, 0)
    raw["T_Deduction"] = raw["Deduction"].where(paid, 0)
    raw["Employee"] = raw["Employee"].map(_emp_key)
    # salary still to pay this FY: every month from the payroll month to March, same day rule as above
    months, (m_, y_) = [], (tm, ty)
    while (y_, m_) <= (fy_start + 1, 3):
        months.append((m_, y_))
        m_, y_ = (1, y_ + 1) if m_ == 12 else (m_ + 1, y_)
    per_month = pd.DataFrame({i: [round(r["Gross"] * _run_days(r, by, bm, y2, m2)[1]) for _, r in raw.iterrows()]
                              for i, (m2, y2) in enumerate(months)}, index=raw.index)
    per_month["Employee"] = raw["Employee"]
    per_month = per_month.groupby("Employee").sum()
    emp = raw.groupby("Employee", sort=False).agg(
        Name=("Name", "first"), Department=("Department", "first"),
        WorkOrderNo=("WorkOrderNo", lambda s: ", ".join(dict.fromkeys(str(x).strip() for x in s))),
        WOStartDate=("WOStartDate", "min"), WOEndDate=("WOEndDate", "max"),      # several WOs: first start, last end
        joindate=("joindate", "max"), resigndate=("resigndate", "max"),
        Status=("Status", lambda s: "; ".join(dict.fromkeys(x for x in s if x))),
        Days=("Days", "max"), Gross=("T_Gross", "sum"), PF=("T_PF", "sum"), ESIC=("T_ESIC", "sum"),
        PT=("T_PT", "max"), Deduction=("T_Deduction", "sum"), BaseTDS=("TDS", "sum"),
        TDSOpt=("TDSOpt", "max")).reset_index()
    pay = emp["Gross"] > 0
    emp.loc[pay & emp["Status"].str.startswith(("Work order", "Resigned")), "Status"] = ""   # another WO still runs

    # TDS for the FY of the payroll month: (annual tax - TDS already deducted this FY) / months left to pay
    k = emp["Employee"]
    fy["Employee"] = fy["Employee"].map(_emp_key)
    fy = fy.groupby("Employee")[["FYPaid", "FYTDS"]].sum()
    pm = per_month.reindex(k).fillna(0)
    emp["SalPaid"] = k.map(fy["FYPaid"]).fillna(0).round(0).values
    emp["SalProposed"] = pm.sum(axis=1).values
    emp["SalTotal"] = emp["SalPaid"] + emp["SalProposed"]
    emp["Already"] = k.map(fy["FYTDS"]).fillna(0).round(0).values
    emp["MonthsLeft"] = (pm > 0).sum(axis=1).values
    old = pd.to_numeric(emp["TDSOpt"], errors="coerce") == 1           # TDSOpt 1 = Old, 2 / blank = New
    emp["Regime"] = old.map({True: "Old", False: "New"})
    emp["Yearly"] = [round(tds_tax_old(t - TDS_STD_DEDUCTION["old"]) if o else tds_tax_new(t - TDS_STD_DEDUCTION["new"]))
                     for t, o in zip(emp["SalTotal"], old)]
    left = (emp["Yearly"] - emp["Already"]).clip(lower=0)
    emp["TDS"] = (left / emp["MonthsLeft"].where(emp["MonthsLeft"] > 0)).fillna(0).round(0).where(pay, 0)
    emp["TakeHome"] = emp["Gross"] - emp["PF"] - emp["ESIC"] - emp["PT"] - emp["Deduction"] - emp["TDS"]
    emp = pd.concat([emp[pay].sort_values("Employee", key=lambda s: pd.to_numeric(s, errors="coerce")),
                     emp[~pay].sort_values("Employee", key=lambda s: pd.to_numeric(s, errors="coerce"))])
    out = emp.rename(columns={
        "Days": f"Pay days {tl}", "Gross": f"Gross {tl}", "PF": f"PF_Employee {tl}", "ESIC": f"ESIC_Employee {tl}",
        "PT": f"PT {tl}", "Deduction": f"Deduction {tl}", "TDS": f"TDS {tl}", "TakeHome": f"Take-home {tl}",
        "Regime": "TDS regime", "BaseTDS": f"TDS {bl}", "Yearly": f"Annual TDS ({fy_text})",
        "Already": f"TDS already deducted ({fy_text})", "SalPaid": f"Salary paid till now (gross, {fy_text})",
        "SalProposed": "Proposed salary (gross, rest of FY)", "SalTotal": f"Total salary (gross, {fy_text})",
        "MonthsLeft": "TDS months left"})
    out = out.rename(columns={"WOStartDate": "WO start date", "WOEndDate": "WO end date", "joindate": "Joining date",
                              "resigndate": "Resign date"})
    out = out[["Employee", "Name", "Department", "WorkOrderNo", "WO start date", "WO end date", "Joining date",
               "Resign date", "Status", f"Pay days {tl}", f"Gross {tl}",
               f"PF_Employee {tl}", f"ESIC_Employee {tl}", f"PT {tl}", f"Deduction {tl}", f"TDS {tl}",
               f"Take-home {tl}", "TDS regime", f"Salary paid till now (gross, {fy_text})",
               "Proposed salary (gross, rest of FY)", f"Total salary (gross, {fy_text})",
               f"Annual TDS ({fy_text})", f"TDS already deducted ({fy_text})", "TDS months left",
               f"TDS {bl}"]].reset_index(drop=True)
    eff = current_eff()

    def vis(*keys):                                       # a figure in the text only if the login may see it
        return eff is None or all(eff.key_visible(k_) for k_ in keys)
    p_ = emp[pay]
    n_part = int((p_["Status"] != "").sum())
    stop = emp[~pay]["Status"]
    n_wo, n_res = int(stop.str.startswith("Work order").sum()), int(stop.str.startswith("Resigned").sum())
    only = " · ".join(v for v in (f.get("category"), f.get("department")) if v)
    head = [f"Gross {inr(p_['Gross'].sum())}" if vis("GrossAmount") else "",
            f"TDS {inr(p_['TDS'].sum())}" if vis("IncomeTax") else "",
            f"Take-home {inr(p_['TakeHome'].sum())}" if vis("TakeHome") else ""]
    ded = [f"PF (employee) {inr(p_['PF'].sum())}" if vis("PF_Employee") else "",
           f"ESIC (employee) {inr(p_['ESIC'].sum())}" if vis("ESIC_Employee") else "",
           f"PT {inr(p_['PT'].sum())}" if vis("PT") else "",
           f"Other deduction {inr(p_['Deduction'].sum())}" if vis("Deduction") else ""]
    text = (f"**Salary to pay for {tl}{f' — {only} only' if only else ''} (from {bl} salary data): "
            f"{len(p_):,} employees**" + "".join(" · " + x for x in head if x) + "\n"
            + ("- " + " · ".join(x for x in ded if x) + "\n" if any(ded) else "")
            + (f"- Part month: {n_part:,} employee(s) (work order ends or they leave during {tl})\n" if n_part else "")
            + (f"- **Not paid in {tl}: {len(stop):,}** — work order ended / no end date {n_wo:,}, resigned {n_res:,}. "
               "They are listed at the bottom with the reason; check them if a work order was extended.\n"
               if len(stop) else "")
            + (f"- {fy_text} salary (gross) of these {len(p_):,}: paid till now {inr(p_['SalPaid'].sum())} + "
               f"proposed {inr(p_['SalProposed'].sum())} = total {inr(p_['SalTotal'].sum())}\n"
               if vis("GrossAmount", "CTC") else "")
            + (f"- {fy_text} TDS of these {len(p_):,}: annual {inr(p_['Yearly'].sum())} · already deducted "
               f"{inr(p_['Already'].sum())} · {tl} {inr(p_['TDS'].sum())} (TDS {bl} was {inr(p_['BaseTDS'].sum())})"
               if vis("IncomeTax") else ""))
    notes.insert(0, f"Understood as: **Salary run for {tl}** ({fy_text}) — one row per employee, using {bl} "
                    "(salary month Year/Month) as the base.")
    notes.append("Rule: gross = base month gross without arrears, × days paid in the month (same day rule as the "
                 "projection); PF / ESIC scale the same way; PT and other deduction as in the base month; "
                 "take-home = gross − PF − ESIC − PT − deduction − TDS. "
                 f"**TDS** ({fy_text}, the FY of {tl}): salary paid till now = this FY's rows in the data "
                 f"(CTC − PF employer, as in the TDS procedure) + proposed = gross of every month from {tl} to "
                 f"Mar-{fy_start + 1} it will be paid (same day rule); annual tax on the total with the employee's own "
                 "regime (TDSOpt 1 = Old, else New; standard deduction, 87A, 4% cess, no surcharge); "
                 f"{tl} TDS = (annual tax − TDS already deducted this FY) ÷ TDS months left (months from {tl} to "
                 "March with salary), never below 0.")
    if fy_text != TDS_FY_TEXT:
        notes.append(f"⚠️ Tax slabs used are the {TDS_FY_TEXT} slabs — update them if the Budget changes them for {fy_text}.")
    notes.append(f"⏱️ No AI · SQL: {t_sql:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": out, "report": "salary_run", "sql": sql, "params": params,
            "intent": {"report": "salary_run", "target": tl, "base": bl, "rows": len(out),
                       "department": f.get("department"), "category": f.get("category")}}


# ---------------- Length of service: "which employee work more than 1 year" ----------------
TENURE = re.compile(
    r"\b(?:work(?:ed|ing|s)?|serv(?:ed|ing|ice)|tenure|completed?|experience|with\s+us|in\s+(?:the\s+)?company)\b"
    r"[^?.]*?\b(\d+(?:\.\d+)?)\s*\+?\s*(years?|yrs?|months?)\b"
    r"|\b(\d+(?:\.\d+)?)\s*\+?\s*(years?|yrs?|months?)\s+(?:of\s+)?(?:service|tenure|completed|experience)\b", re.I)
TENURE_LESS = re.compile(r"\b(?:less|under|below|within|fewer|not\s+(?:yet\s+)?completed?|up\s*to|upto)\b", re.I)
TENURE_MONEY = re.compile(r"\b(?:ctc|salary|salaries|gross|tds|margin|pf|esic|take.?home|invoice|payroll|paid)\b", re.I)


def wants_tenure(question):
    """'which employee work more than 1 year', 'employees with 2+ years service', 'less than 6 months'.
    Questions about money ('ctc of employees working 2 years') go to the normal reports."""
    return bool(TENURE.search(question)) and not TENURE_MONEY.search(question)


def tenure_answer(question, notes, t0):
    m_ = TENURE.search(question)
    num, unit = (m_.group(1), m_.group(2)) if m_.group(1) else (m_.group(3), m_.group(4))
    months_req = float(num) * (1 if unit.lower().startswith("m") else 12)
    less = bool(TENURE_LESS.search(question))
    ql = question.lower()
    only = "current" if re.search(r"\b(?:current|active|still|presently|now|working\s+employees?)\b", ql) and \
        not re.search(r"\bresign|\bleft\b", ql) else ("resigned" if re.search(r"\bresign\w*|\bleft\b|\bex[- ]", ql)
                                                       else "")
    periods = parse_periods(question)
    asof = _month_end(periods[0][1], periods[0][0]) if periods else datetime.date.today()
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where("salary", [], f)
    sql = f"""SELECT [Employee], MAX([Name]) AS Name, MAX([Department]) AS Department, MAX([Category]) AS Category,
       MAX([joindate]) AS joindate, MAX([resigndate]) AS resigndate,
       MIN([WOStartDate]) AS WOStartDate, MAX([WOEndDate]) AS WOEndDate,
       MAX(DATEFROMPARTS([Year],[Month],1)) AS LastSalary
FROM {T} {where} GROUP BY [Employee]"""
    t1 = time.time()
    raw = run_query(sql, params)
    t_sql = time.time() - t1
    if raw.empty:
        return clarify(question, "There are no employees in the data you can see.")
    for c in ("joindate", "resigndate", "WOStartDate", "WOEndDate", "LastSalary"):
        raw[c] = pd.to_datetime(raw[c], errors="coerce")
    if raw["joindate"].isna().all():
        return {"text": "🔒 I can't see joining dates for your login, so I can't work out length of service.",
                "note": "", "df": None, "report": "blocked", "sql": "", "params": [], "intent": {}}
    asof_ts = pd.Timestamp(asof)
    raw = raw[raw["joindate"].notna() & (raw["joindate"] <= asof_ts)].copy()
    gone = raw["resigndate"].notna() & (raw["resigndate"] <= asof_ts)
    end = raw["resigndate"].where(gone, asof_ts)
    raw["Days"] = (end - raw["joindate"]).dt.days + 1
    # whole months completed (calendar), then the remaining days
    mon = (end.dt.year - raw["joindate"].dt.year) * 12 + (end.dt.month - raw["joindate"].dt.month) \
        - (end.dt.day < raw["joindate"].dt.day).astype(int)
    raw["Months"] = mon.clip(lower=0)
    raw["Service"] = [f"{int(x) // 12} y {int(x) % 12} m" for x in raw["Months"]]
    raw["Years"] = (raw["Days"] / 365.25).round(2)
    raw["Status"] = gone.map({True: "Resigned", False: "Working"})
    stale = ~gone & (raw["LastSalary"] < raw["LastSalary"].max())      # no resign date, but no salary in the latest month
    raw.loc[stale, "Status"] = "Not paid since " + (raw.loc[stale, "LastSalary"] + pd.offsets.MonthBegin(1)).dt.strftime("%b-%Y")
    keep = raw["Months"] < months_req if less else raw["Months"] >= months_req   # "more than 1 year" = 1 year done
    if only:
        keep &= raw["Status"] == ("Working" if only == "current" else "Resigned")
    order = raw["Status"].map(lambda s: 0 if s == "Working" else 2 if s == "Resigned" else 1)
    out = raw[keep].assign(_o=order).sort_values(["_o", "Days"], ascending=[True, False]).drop(columns="_o")
    out["Employee"] = out["Employee"].map(_emp_key)
    out = out.rename(columns={"joindate": "Joining date", "resigndate": "Resign date", "WOStartDate": "WO start date",
                              "WOEndDate": "WO end date", "LastSalary": "Last salary month",
                              "Service": "Service (years, months)", "Years": "Service in years",
                              "Days": "Service days"})
    out["Last salary month"] = out["Last salary month"].dt.strftime("%b-%Y")
    out = out[["Employee", "Name", "Department", "Category", "Status", "Joining date", "Resign date",
               "Service (years, months)", "Service in years", "Service days", "WO start date", "WO end date",
               "Last salary month"]].reset_index(drop=True)
    want = (f"{num} {unit}").replace("yrs", "years").replace("yr", "year")
    cmp_ = "less than" if less else "at least"
    who = {"current": "working employees", "resigned": "resigned employees"}.get(only, "employees")
    sub = " · ".join(v for v in (f.get("category"), f.get("department")) if v)
    n_w, n_r = int((out["Status"] == "Working").sum()), int((out["Status"] == "Resigned").sum())
    n_s = len(out) - n_w - n_r
    text = (f"**{len(out):,} {who} with {cmp_} {want} of service** (as on {asof:%d-%b-%Y})"
            f"{f' — {sub} only' if sub else ''}"
            + (f"\n- Working: **{n_w:,}** · resigned: **{n_r:,}**"
               + (f" · no resign date but not paid in the latest month: **{n_s:,}**" if n_s else "") if not only else "")
            + (f"\n- Longest: {out.iloc[0]['Name']} ({out.iloc[0]['Employee']}) — {out.iloc[0]['Service (years, months)']}"
               if len(out) else ""))
    notes.insert(0, f"Understood as: **length of service {cmp_} {want}**.")
    notes.append("Rule: service = joining date → resign date (if they left on or before "
                 f"{asof:%d-%b-%Y}) or {asof:%d-%b-%Y}; counted in completed calendar months. "
                 "Employees without a joining date are left out. Ask *current employees working more than 1 year* "
                 "or *resigned …* to see only one group.")
    notes.append(f"⏱️ No AI · SQL: {t_sql:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": out, "report": "tenure", "sql": sql, "params": params,
            "intent": {"report": "tenure", "rows": len(out), "department": f.get("department"),
                       "category": f.get("category")}}


# ---------------- Work order history: which WOs an employee worked on, and breaks between them ----------------
_WO = r"work\s*-?\s*orders?"
WO_HIST = re.compile(rf"\b{_WO}\s+(?:history|timeline|journey|details?\s+of\s+(?:employee|emp))\b"
                     rf"|\bin\s+which\s+{_WO}\b|\bwhich\s+{_WO}\s+(?:did|does|do|has|have|is|was|were)\b[^?]*\bwork"
                     rf"|\b(?:all|every)\s+(?:the\s+)?{_WO}\s+(?:of|for)\s+(?:employee|emp|\d|[a-z]+\s+[a-z]+)", re.I)
WO_BREAK = re.compile(rf"\b(?:breaks?|bracks?|gaps?)\b[^?]*\b{_WO}\b|\b{_WO}\b[^?]*\b(?:breaks?|bracks?|gaps?)\b"
                      r"|\b(?:working|worked|works?)\s+(?:continuous\w*|consist[ae]nt\w*|without\s+(?:a\s+|any\s+)?break)"
                      r"|\b(?:continuous\w*|consist[ae]nt\w*)\s+(?:working|employees?|service)\b"
                      r"|\b(?:took|take|taken|taking|had|has)\s+(?:a\s+|any\s+)?(?:breaks?|bracks?)\b", re.I)
WO_NAME_STOP = {"in", "which", "workorder", "workorders", "work", "order", "orders", "wo", "history", "timeline", "did",
                "does", "do", "has", "have", "is", "was", "were", "employee", "emp", "of", "for", "the", "all", "every",
                "an", "a", "worked", "working", "works", "with", "start", "end", "date", "dates", "and", "his", "her",
                "their", "show", "me", "list", "give", "details", "detail", "journey", "who", "breaks", "break", "gap",
                "gaps", "between", "took", "take", "on", "what", "are", "be", "been"}


def wants_wo_history(question):
    """'in which workorder 5725 worked', 'work order history of amit yadav', 'who took a break between work
    orders', 'who is working continuously'. Money questions go to the normal reports."""
    return bool(WO_HIST.search(question) or WO_BREAK.search(question)) and not TENURE_MONEY.search(question)


def _wo_pick(emp, question):
    """Employee codes / names asked about (None = everyone). emp: one row per employee with Employee, Name."""
    t = question.lower()
    codes = set(emp["Employee"])
    nums = [n for n in re.findall(r"\b\d{2,}\b", t) if n in codes]
    if nums:
        return nums
    full = find_employee_name(question)
    if full:
        return list(emp.loc[emp["Name"].astype(str).str.strip().str.lower() == full.lower(), "Employee"])
    if not WO_HIST.search(question) or WO_BREAK.search(question):   # "who took a break" = everyone
        return None
    words = [w for w in re.findall(r"[a-z]+", t) if w not in WO_NAME_STOP and len(w) >= 3]
    if not words:
        return None
    nw = emp["Name"].astype(str).str.lower().map(lambda s: re.findall(r"[a-z]+", s))
    hit = nw.map(lambda x: all(any(p.startswith(w) for p in x) for w in words))
    return list(emp.loc[hit, "Employee"]) if hit.any() else []


def wo_history_answer(question, notes, t0):
    ql = question.lower()
    m_ = re.search(r"\b(\d+)\s*\+?\s*days?\b", ql)
    min_gap = 1                                            # "break of more than 30 days" = 31+, "30+ days" = 30+
    if m_ and re.search(r"\b(?:break|brack|gap)", ql):
        min_gap = int(m_.group(1)) + (1 if re.search(rf"\b(?:more|over|above|greater)\b[^0-9]*{m_.group(1)}", ql) else 0)
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    where, params = build_where("salary", [], f)
    sql = f"""SELECT [Employee], [WorkOrderNo], MAX([Name]) AS Name, MAX([Department]) AS Department,
       MAX([Category]) AS Category, MIN([WOStartDate]) AS WOStart, MAX([WOEndDate]) AS WOEnd,
       MIN(DATEFROMPARTS([Year],[Month],1)) AS FirstSal, MAX(DATEFROMPARTS([Year],[Month],1)) AS LastSal,
       COUNT(DISTINCT [Year]*100+[Month]) AS SalMonths, MAX([joindate]) AS joindate, MAX([resigndate]) AS resigndate
FROM {T} {where} GROUP BY [Employee], [WorkOrderNo]"""
    t1 = time.time()
    raw = run_query(sql, params)
    t_sql = time.time() - t1
    if raw.empty:
        return clarify(question, "There are no work orders in the data you can see.")
    for c in ("WOStart", "WOEnd", "FirstSal", "LastSal", "joindate", "resigndate"):
        raw[c] = pd.to_datetime(raw[c], errors="coerce")
    if raw["WOStart"].isna().all() and raw["WOEnd"].isna().all():
        return {"text": "🔒 I can't see work order dates for your login, so I can't show work order history.",
                "note": "", "df": None, "report": "blocked", "sql": "", "params": [], "intent": {}}
    raw["Employee"] = raw["Employee"].map(_emp_key)
    raw["WorkOrderNo"] = raw["WorkOrderNo"].astype(str).str.strip()
    # a missing WO date falls back to the salary months it was paid in
    raw["DateNote"] = ""
    ms, me_ = raw["WOStart"].isna(), raw["WOEnd"].isna()
    raw.loc[ms, "WOStart"] = raw.loc[ms, "FirstSal"]
    raw.loc[me_, "WOEnd"] = raw.loc[me_, "LastSal"] + pd.offsets.MonthEnd(0)
    raw.loc[ms | me_, "DateNote"] = "WO date missing - salary months used"
    raw = raw.sort_values(["Employee", "WOStart", "WOEnd"]).reset_index(drop=True)
    # break before each WO = days between the latest end so far and this start (overlap / back-to-back = 0)
    prev_end = raw.groupby("Employee")["WOEnd"].transform(lambda s: s.cummax().shift(1))
    gap = (raw["WOStart"] - prev_end).dt.days - 1
    raw["Gap"] = gap.where(prev_end.notna()).fillna(0).clip(lower=0).astype(int)
    raw["IsBreak"] = raw["Gap"] >= min_gap
    prev_wo = raw.groupby("Employee")["WorkOrderNo"].shift(1)
    raw["Note"] = ""
    raw.loc[prev_end.notna() & ~raw["IsBreak"], "Note"] = "Continuous after " + prev_wo
    raw.loc[raw["IsBreak"], "Note"] = "Break of " + raw["Gap"].astype(str) + " days after " + prev_wo
    raw.loc[prev_end.isna(), "Note"] = "First work order"
    raw["Note"] = raw["Note"] + raw["DateNote"].map(lambda x: f" ({x})" if x else "")
    latest = raw["LastSal"].max()
    emp = raw.groupby("Employee", sort=False).agg(
        Name=("Name", "last"), Department=("Department", "last"), Category=("Category", "last"),
        WOs=("WorkOrderNo", "nunique"), Breaks=("IsBreak", "sum"),
        BreakDays=("Gap", lambda s: int(s[raw.loc[s.index, "IsBreak"]].sum())),
        Longest=("Gap", lambda s: int(s[raw.loc[s.index, "IsBreak"]].max()) if raw.loc[s.index, "IsBreak"].any() else 0),
        First=("WOStart", "min"), Last=("WOEnd", "max"), joindate=("joindate", "max"), resigndate=("resigndate", "max"),
        LastSal=("LastSal", "max")).reset_index()
    emp["Status"] = (emp["Breaks"] > 0).map({True: "Took break", False: "Continuous"})
    now = emp["resigndate"].notna().map({True: "Resigned", False: ""})
    now = now.where(now != "", (emp["LastSal"] < latest).map({True: "Not paid in latest month", False: "Working"}))
    emp["Now"] = now
    pick = _wo_pick(emp, question)

    def d(x):
        return "" if pd.isna(x) else f"{x:%d-%b-%y}"
    if pick is not None:                                   # one / a few employees: one row per work order
        if not pick:
            return clarify(question, "I couldn't find that employee. Please give the employee code, "
                                     "e.g. *work order history of 5725*.")
        rows = raw[raw["Employee"].isin(pick)].copy()
        rows["No"] = rows.groupby("Employee").cumcount() + 1
        rows["WO days"] = (rows["WOEnd"] - rows["WOStart"]).dt.days + 1
        for c in ("FirstSal", "LastSal"):
            rows[c] = rows[c].dt.strftime("%b-%Y")
        out = rows.rename(columns={"WOStart": "WO start date", "WOEnd": "WO end date", "Gap": "Break before (days)",
                                   "FirstSal": "First salary month", "LastSal": "Last salary month",
                                   "SalMonths": "Salary months", "joindate": "Joining date",
                                   "resigndate": "Resign date"})[
            ["Employee", "Name", "No", "WorkOrderNo", "Department", "Category", "WO start date", "WO end date",
             "WO days", "Break before (days)", "Note", "First salary month", "Last salary month", "Salary months",
             "Joining date", "Resign date"]].reset_index(drop=True)
        lines = []
        for _, e in emp[emp["Employee"].isin(pick)].iterrows():
            lines.append(f"**{e['Name']} ({e['Employee']})** — {e['WOs']} work order(s), {d(e['First'])} to "
                         f"{d(e['Last'])}: " + ("**continuous**, no break" if not e["Breaks"] else
                                               f"**{int(e['Breaks'])} break(s)**, {e['BreakDays']:,} days in all "
                                               f"(longest {e['Longest']:,} days)") + f" · now: {e['Now']}")
        text = "\n".join("- " + x for x in lines[:20]) if len(lines) > 1 else lines[0]
        report = "wo_history"
    else:                                                  # everyone: one row per employee
        want_break = bool(re.search(r"\b(?:break|brack|gap)", ql))
        want_cont = bool(re.search(r"\bcontinu|\bconsist[ae]nt|without\s+(?:a\s+|any\s+)?break", ql))
        sel = emp
        if want_break and not want_cont:
            sel = emp[emp["Breaks"] > 0]
        elif want_cont and not want_break:
            sel = emp[emp["Breaks"] == 0]
        if re.search(r"\b(?:current|active|still|presently|working\s+now)\b", ql):
            sel = sel[sel["Now"] == "Working"]
        timeline = raw.groupby("Employee").apply(
            lambda g: " → ".join((f"[break {r.Gap} d] " if r.IsBreak else "") +
                                 f"{r.WorkOrderNo} ({d(r.WOStart)} to {d(r.WOEnd)})" for r in g.itertuples()))
        sel = sel.assign(Timeline=sel["Employee"].map(timeline))
        sel = sel.sort_values(["Breaks", "BreakDays", "Employee"], ascending=[False, False, True],
                              key=lambda s: pd.to_numeric(s, errors="coerce") if s.name == "Employee" else s)
        out = sel.rename(columns={"WOs": "Work orders", "Breaks": "Breaks", "BreakDays": "Total break days",
                                  "Longest": "Longest break (days)", "First": "First WO start", "Last": "Last WO end",
                                  "joindate": "Joining date", "resigndate": "Resign date", "Now": "Now",
                                  "Timeline": "Work orders (start to end)"})[
            ["Employee", "Name", "Department", "Category", "Status", "Now", "Work orders", "Breaks", "Total break days",
             "Longest break (days)", "First WO start", "Last WO end", "Joining date", "Resign date",
             "Work orders (start to end)"]].reset_index(drop=True)
        n_b, n_c = int((sel["Breaks"] > 0).sum()), int((sel["Breaks"] == 0).sum())
        multi = int((sel["WOs"] > 1).sum())
        sub = " · ".join(v for v in (f.get("category"), f.get("department")) if v)
        text = (f"**{len(sel):,} employees{f' — {sub} only' if sub else ''}: {n_c:,} worked continuously, "
                f"{n_b:,} took a break between work orders**"
                + (f" (a break = {min_gap}+ day(s) between one work order's end and the next one's start)")
                + f"\n- {multi:,} of them worked on more than one work order"
                + (f"\n- Longest break: {sel.iloc[0]['Name']} ({sel.iloc[0]['Employee']}) — "
                   f"{int(sel['Longest'].max()):,} days" if n_b else "")
                + "\n- Ask *work order history of 5725* for one employee's work orders with dates.")
        report = "wo_breaks"
    notes.insert(0, "Understood as: **work order history** — each work order's start and end date, and the break "
                    "between one work order and the next.")
    notes.append("Rule: work orders sorted by start date; break = days between the latest end date so far and the "
                 "next start date (back-to-back or overlapping = continuous). The time after the last work order "
                 "is not counted as a break — see *Now*. A missing WO date uses the salary months instead.")
    notes.append(f"⏱️ No AI · SQL: {t_sql:.1f}s")
    return {"text": text, "note": " ".join(notes), "df": out, "report": report, "sql": sql, "params": params,
            "intent": {"report": report, "rows": len(out), "department": f.get("department"),
                       "category": f.get("category")}}


# ------------------------------ TDS calculator (amounts typed by the user) ------------------------------
# Worked out HERE on this PC - nothing is sent outside. Same rules as the SQL functions
# dbo.fn_TDS_NewRegime_FY2627 / dbo.fn_TDS_OldRegime_FY2627 used by "proposed tds of ...":
#   Yearly TDS = TDS already paid + proposed TDS (remaining) = annual tax with cess
#   Monthly TDS = proposed TDS / pending months of the FY, never below 0
# Each signed-in user's past calculations are kept in tds_history.json (per user ID, deletable).
TDS_HISTORY_FILE = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "tds_history.json")
TDS_HISTORY_MAX = 500                 # per user; the oldest drop off
_TDS_LOCK = _threading.RLock()
TDS_STD_DEDUCTION = {"new": 75000, "old": 50000}


def tds_tax_new(taxable):
    """= dbo.fn_TDS_NewRegime_FY2627: slabs 4/8/12/16/20/24 lakh, 87A up to 12 lakh + marginal relief, 4% cess."""
    if not taxable or taxable <= 0:
        return 0.0
    t = round(taxable / 10.0) * 10                       # Sec 288A: nearest Rs 10
    tax = 0.0
    for lo, hi, rate in ((400000, 800000, .05), (800000, 1200000, .10), (1200000, 1600000, .15),
                         (1600000, 2000000, .20), (2000000, 2400000, .25), (2400000, None, .30)):
        if t > lo:
            tax += ((t if hi is None else min(t, hi)) - lo) * rate
    if t <= 1200000:
        tax = 0.0
    elif tax > t - 1200000:
        tax = t - 1200000
    tax = round(tax)
    return float(tax + round(tax * 0.04))


def tds_tax_old(taxable):
    """= dbo.fn_TDS_OldRegime_FY2627: slabs 2.5/5/10 lakh, 87A up to 5 lakh, 4% cess."""
    if taxable > 1000000:
        tax = 112500 + (taxable - 1000000) * 0.30
    elif taxable > 500000:
        tax = 12500 + (taxable - 500000) * 0.20
    elif taxable > 250000:
        tax = (taxable - 250000) * 0.05
    else:
        tax = 0.0
    if taxable <= 500000:
        tax = 0.0
    return round(tax * 1.04, 2)


def tds_fy(today=None):
    """FY of today, e.g. 'FY 2026-27', and the months left in it counting this month (Oct -> Oct..Mar = 6)."""
    d = today or datetime.date.today()
    start = d.year if d.month >= 4 else d.year - 1
    left = (16 - d.month) if d.month >= 4 else (4 - d.month)
    return f"FY {start}-{str(start + 1)[-2:]}", left


def tds_calculate(paid, proposed, regime, tds_paid=0, months=None):
    """paid + proposed salary of the FY -> yearly TDS (paid + proposed) and monthly TDS for the pending months.
    Both regimes are worked out so you can see which is lower; 'regime' is the one asked for."""
    regime = str(regime or "").strip().lower()
    if regime not in ("new", "old"):
        raise ValueError("Regime must be New or Old.")
    vals = {}
    for k, v in (("paid salary", paid), ("proposed salary", proposed), ("TDS already paid", tds_paid)):
        try:
            v = float(str(v if v not in (None, "") else 0).replace(",", "").strip())
        except ValueError:
            raise ValueError(f"Please type a number for {k}.")
        if v != v or v < 0 or v > 1e10:
            raise ValueError(f"{k.capitalize()} must be between 0 and 1,000 crore.")
        vals[k] = round(v, 2)
    fy, auto_months = tds_fy()
    if months in (None, ""):
        months = auto_months
    try:
        months = int(months)
    except (TypeError, ValueError):
        raise ValueError("Pending months must be a whole number from 0 to 12.")
    if not 0 <= months <= 12:
        raise ValueError("Pending months must be a whole number from 0 to 12.")
    total = vals["paid salary"] + vals["proposed salary"]
    out = {"fy": fy, "regime": regime, "paid": vals["paid salary"], "proposed": vals["proposed salary"],
           "total": total, "tds_paid": vals["TDS already paid"], "months": months}
    for r, fn in (("new", tds_tax_new), ("old", tds_tax_old)):
        taxable = max(total - TDS_STD_DEDUCTION[r], 0)
        yearly = round(fn(taxable))
        remaining = yearly - vals["TDS already paid"]
        out[r] = {"std_deduction": TDS_STD_DEDUCTION[r], "taxable": taxable, "yearly_tds": yearly,
                  "proposed_tds": max(remaining, 0), "excess_paid": max(-remaining, 0),
                  "monthly_tds": round(max(remaining, 0) / months) if months else None}
    out.update(out[regime])
    other = "old" if regime == "new" else "new"
    out["lower"] = regime if out[regime]["yearly_tds"] <= out[other]["yearly_tds"] else other
    out["surcharge_note"] = total > 5000000                  # the SQL rules have no surcharge either
    return out


def _tds_history_all():
    try:
        with open(TDS_HISTORY_FILE, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (FileNotFoundError, ValueError):
        return {}


def _tds_history_save(d):
    tmp = TDS_HISTORY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=1)
    _os.replace(tmp, TDS_HISTORY_FILE)


def tds_history(uid):
    """This user's saved calculations, newest first."""
    with _TDS_LOCK:
        return list(reversed(_tds_history_all().get(uid, [])))


def tds_history_add(uid, calc):
    import secrets as _secrets
    item = dict(calc, id=_secrets.token_hex(6), time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    with _TDS_LOCK:
        d = _tds_history_all()
        d[uid] = (d.get(uid, []) + [item])[-TDS_HISTORY_MAX:]
        _tds_history_save(d)
    return item


def tds_history_delete(uid, item_id=None):
    """Deletes one saved calculation (by id) or, with no id, all of this user's. Returns how many were removed."""
    with _TDS_LOCK:
        d = _tds_history_all()
        mine = d.get(uid, [])
        keep = [x for x in mine if item_id and x.get("id") != item_id]
        if keep:
            d[uid] = keep
        else:
            d.pop(uid, None)
        _tds_history_save(d)
    return len(mine) - len(keep)


# ======================= QUERY PLANNER (understands the WHOLE question) =======================
# Reads every part of a question - filters, months, grouping, requested columns, calculations,
# sorting, "which ... highest" sub-questions, comparisons - into ONE plan, then builds ONE safe
# read-only query from it. Business rules stay the same as everywhere else:
#   invoice / base sales counted ONCE per WorkOrderNo + salary Year + Month (spread over its rows),
#   CTC = sum of employee rows, employee count = distinct employees, margin % on CTC by default.
PLAN_DIMS = {   # dimension: (words that mean it)
    "WorkOrderNo": r"work\s*-?\s*orders?|workorders?|w\.?o\.?s?|wo(?:\s*no)?",
    "Department": r"departments?|depts?|clients?",
    "Category": r"categor(?:y|ies)",
    "Employee": r"employees?|emps?|persons?|staff",
    "Month": r"months?",
}
PLAN_COLS = [   # (pattern, column) - longest phrases first; the order the user writes them is kept
    (r"margin\s*(?:%|percent(?:age)?|pct|ratio)|%\s*margin|margin\s+rate", "MarginPct"),
    (r"average\s+ctc|avg\.?\s+ctc|ctc\s+per\s+(?:employee|head|person)|per\s+head\s+ctc", "AvgCTC"),
    (r"employee\s*count|employees?\s+count|emp\s*count|no\.?\s+of\s+employees?|number\s+of\s+employees?|"
     r"head\s*-?\s*count|\bhc\b|manpower|employees?\s+deployed|how\s+many\s+employees?|total\s+employees?|"
     r"count\s+of\s+(?:employees?|people|staff)|number\s+of\s+(?:people|staff)", "Employees"),
    (r"base\s*sales?\s*(?:value|amount)?|basic\s+sales?|sales\s+value|base\s+value|base\s+amount|\brevenue\b|"
     r"\bturnover\b", "BaseSalesValue"),
    (r"invoice\s*(?:amount|value)?|billing|billed", "InvoiceAmount"),
    (r"\bmargin\b|\bprofit\b", "Margin"),
    (r"\bctc\b|cost\s+to\s+company", "CTC"),
    (r"work\s*-?\s*order\s*(?:no\.?|number|num)|workorder\s*(?:no\.?|number)|\bwo\s*(?:no\.?|number)", "WorkOrderNo"),
    (r"\bdepartments?\b|\bdept\b|\bclient\s+name\b", "Department"),
    (r"\bcategor(?:y|ies)\b", "Category"),
    (r"(?:wo\s*|work\s*order\s*)?start\s*date", "WOStartDate"),
    (r"(?:wo\s*|work\s*order\s*)?end\s*date", "WOEndDate"),
    (r"\bmonths?\b", "Month"),
    (r"employee\s+(?:name|code)s?", "Employee"),
]
PLAN_METRICS = ["Employees", "CTC", "AvgCTC", "BaseSalesValue", "InvoiceAmount", "Margin", "MarginPct"]
PLAN_ATTRS = ["WOStartDate", "WOEndDate", "Department", "Category"]
PLAN_DEFAULT = ["Employees", "BaseSalesValue", "CTC", "Margin", "MarginPct"]
PLAN_COMPLEX = re.compile(
    r"\bcolumns?\b|\bfields?\b|\bcontain(?:s|ing)?\b|\binclud(?:e|ing)\b|\bhaving\b|\bwhere\b|"
    r"\b(?:above|below|over|under|greater|less|more|lower|higher)\s+than\b|\b(?:above|below|over|under)\s+\d|[<>]=?\s*\d|"
    r"\bstart(?:s|ing)?\s+with\b|\bbegin(?:s|ning)?\s+with\b|\bcontains?\b|\blike\b|"
    r"\balso\b|\bwhich\b.*\b(?:highest|lowest|most|least|max|min|top|biggest|smallest)\b|"
    r"\bhighlight\b|\bnegative\b|\bcompare\b|\bvs\.?\b|\bversus\b|\bmonth[\s\-]*on[\s\-]*month\b|\bmom\b|"
    r"\bgrowth\b|\bchange\b|\bdifference\b",
    re.IGNORECASE)
NUM_RX = r"(\d[\d,]*(?:\.\d+)?)\s*(lakhs?|lacs?|l\b|crores?|cr\b|k\b|thousand)?"


def _num(txt, unit):
    v = float(txt.replace(",", ""))
    u = (unit or "").lower()
    if u.startswith(("lakh", "lac")) or u == "l":
        v *= 1e5
    elif u.startswith("cr"):
        v *= 1e7
    elif u in ("k", "thousand"):
        v *= 1e3
    return v


def _wo_patterns(question):
    """Work-order filters written in the question: 'm%', 'starting with M', 'containing 4425', 'M244425094'."""
    pats = [t.upper() for t in re.findall(r"(?<![\w%])([A-Za-z0-9][A-Za-z0-9/\-_.()]*%|%[A-Za-z0-9][A-Za-z0-9/\-_.()]*%?)",
                                          question)]
    pats = [p_ for p_ in pats if not re.fullmatch(r"(MARGIN|PCT|PERCENT\w*|PROFIT|CTC)%|[\d.]+%", p_)]   # "10%" = a rate
    wo_ctx = r"(?:work\s*-?\s*orders?|workorders?|\bwos?\b)(?:\s+no\.?)?\s+(?:which\s+|that\s+|are\s+|is\s+)?"
    for m in re.finditer(wo_ctx + r"(?:start(?:s|ing)?|begin(?:s|ning)?)\s+with\s+['\"]?([A-Za-z0-9][A-Za-z0-9/\-]*)['\"]?",
                         question, re.IGNORECASE):
        pats.append(m.group(1).upper() + "%")
    for m in re.finditer(wo_ctx + r"(?:contain(?:s|ing)?|having|with)\s+['\"]?([A-Za-z0-9][A-Za-z0-9/\-]{1,})['\"]?",
                         question, re.IGNORECASE):
        w = m.group(1)
        if not re.fullmatch(r"(?:ctc|margin\w*|invoice\w*|base|sales?|department\w*|employee\w*|month\w*|the|a|an|"
                            r"highest|lowest|negative|positive|zero|no|nil|missing|blank|null|more|less|over|under|"
                            r"start\w*|end\w*|count|details?)", w, re.IGNORECASE):
            pats.append("%" + w.upper() + "%")
    stop = {"the", "all", "my", "our", "each", "every", "top", "which", "what", "how", "of", "for", "and", "per",
            "give", "show", "any", "these", "those", "total", "new", "old", "wise", "its", "his", "her", "list", "many",
            "much", "no", "by", "in", "on", "to", "me", "is", "are", "was", "has", "had", "with", "bottom", "a", "an",
            "first", "last", "next", "this", "that", "such", "both", "other", "than", "then", "only", "most", "least"}
    for m in re.finditer(r"(?<![\w%])([A-Za-z]{1,5})\s+(?:work\s*-?\s*orders?|workorders?|wos?)\b", question):
        w = m.group(1)
        if (len(w) == 1 or w.isupper()) and w.lower() not in stop and w.lower() not in MONTH_WORDS and not re.search(
                r"(?i)(without|except|excluding|other\s+th[ae]n)\s+$", question[:m.start()]):
            pats.append(w.upper() + "%")                 # "July M workorder CTC" -> work orders starting with M
    wo, _ = find_workorder(question)
    if wo and not is_pattern(wo):
        pats.append(wo.upper())
    out = []
    for p_ in pats:
        if re.fullmatch(r"[A-Z0-9/\-_.()% ]{1,80}", p_) and p_ not in out:
            out.append(p_)
    return out


def _plan_metrics():
    return PLAN_METRICS + list(_taught())              # v14: taught formulas are metrics too


def _pname(c):
    return PLAN_NAMES.get(c) or _taught().get(c, {}).get("label") or c


def _metric_of(text):
    tk = taught_in(text)                               # v14: a taught formula named here
    if tk:
        return tk[0]
    for pat, col in PLAN_COLS:
        if col in _plan_metrics() and re.search(pat, text, re.IGNORECASE):
            return col
    return None


MOVE_RX = re.compile(r"\b(?:dropped|drops?|fell|falls?|fallen|decreas\w*|declin\w*|reduc\w*|went\s+down|lower(?:ed)?\s+than\s+last|"
                     r"increas\w*|grew|grow[sn]?|rose|rises?|risen|improv\w*|went\s+up)\b")


def parse_plan(question):
    """Rules-only reading of the whole question. Returns a plan dict (no AI, instant)."""
    q = question
    plan = {"wo": _wo_patterns(q), "group": [], "cols": [], "having": [], "sort": None, "limit": None,
            "which": [], "compare": False, "highlight_negative": False, "how_many": []}
    # requested columns, in the order written (phrases that are conditions / sorting / "which" are not columns)
    qc = re.sub(r"\bwhich\s+(?:[\w%]+\s+){0,3}?(?:has|had|have|is|was|got|with)?\s*(?:the\s+)?"
                r"(?:highest|lowest|most|least|max\w*|min\w*|top|biggest|smallest|largest)\s+[a-z%\s]+?(?=\?|,|$|\band\b|\bin\b)",
                " ", q, flags=re.IGNORECASE)
    qc = re.sub(r"\b(?:sort(?:ed)?|order(?:ed)?|rank(?:ed)?)\b[^,.]*?(?=,|\.|$|\band\b)", " ", qc, flags=re.IGNORECASE)
    qc = re.sub(r"\bwhere\b[^,.]*?(?=,|\.|$|\s+(?:also|and\s+give|with|plus|along)\b)|"
                r"\b(?:margin|ctc|invoice|base\s*sales?|employees?)\s+(?:is|are)?\s*(?:negative|above|below|"
                r"over|under|more\s+than|less\s+than|greater\s+than|>|<)[^,.]*?(?=,|\.|$|\s+(?:also|and\s+give|with|plus|along)\b)",
                " ", qc, flags=re.IGNORECASE)
    qc = re.sub(r"\b((?:top|bottom)\s+\d+\s+(?:[\w%]+\s+){0,3}?)by\s+[a-z\s]+?(?=\s+in\b|\s+for\b|\s+and\b|,|$)",
                r"\1", qc, flags=re.IGNORECASE)
    found, taken = [], [False] * len(qc)
    for pat, col in PLAN_COLS:
        for m in re.finditer(pat, qc, re.IGNORECASE):
            if any(taken[m.start():m.end()]):
                continue
            if col == "Month" and re.search(r"(this|last|next|previous|current|\d+)\s+$", qc[:m.start()], re.IGNORECASE):
                continue                                   # "last 3 months" is a period, not a column
            for i in range(m.start(), m.end()):
                taken[i] = True
            found.append((m.start(), col))
    for _, col in sorted(found):
        if col not in plan["cols"]:
            plan["cols"].append(col)
    for k in taught_in(qc):                            # v14: taught formulas asked for as columns
        if k not in plan["cols"]:
            plan["cols"].append(k)
    # grouping: "workorder wise", "workorder and department wise", "by department", "per month", "month wise"
    dim_rx = "|".join(f"(?P<{k}>{v})" for k, v in PLAN_DIMS.items())
    for m in re.finditer(rf"((?:(?:{dim_rx})\s*(?:no\.?\s*)?(?:,|and|&|\+|/)?\s*)+)-?\s*wise\b", q, re.IGNORECASE):
        for k, v in PLAN_DIMS.items():
            if re.search(rf"\b(?:{v})\b", m.group(1), re.IGNORECASE) and k not in plan["group"]:
                plan["group"].append(k)
    for m in re.finditer(r"\b(?:group(?:ed)?\s+by|by|per|for\s+each|each)\s+(\w+(?:\s+(?:and|&)\s+\w+)?)", q, re.IGNORECASE):
        for k, v in PLAN_DIMS.items():
            if re.search(rf"\b(?:{v})\b", m.group(1), re.IGNORECASE) and k not in plan["group"]:
                if k == "Employee" and re.search(r"count|deployed", q[m.end():m.end() + 8], re.IGNORECASE):
                    continue
                if k == "Employee" and "AvgCTC" in plan["cols"] and re.match(r"per\b", m.group(0), re.IGNORECASE):
                    continue                               # "average ctc PER EMPLOYEE" = the divisor, not a grouping
                plan["group"].append(k)
    m = re.search(r"\b(?:show|list|give|get|display|find)\s+(?:me\s+)?(?:all\s+|the\s+)?(departments|categories|employees)\b",
                  q, re.IGNORECASE)
    if m and not plan["group"]:
        plan["group"].append({"departments": "Department", "categories": "Category", "employees": "Employee"}[m.group(1).lower()])
    for k in ("monthwise", "workorderwise", "departmentwise", "categorywise", "employeewise"):
        if k in q.lower():
            col = {"monthwise": "Month", "workorderwise": "WorkOrderNo", "departmentwise": "Department",
                   "categorywise": "Category", "employeewise": "Employee"}[k]
            if col not in plan["group"]:
                plan["group"].append(col)
    # conditions on numbers: "ctc above 2 lakh", "margin < 0", "employees more than 5"
    for m in re.finditer(r"(margin\s*%|margin\s+percent\w*|margin|ctc|base\s*sales?|invoice|employee\s*count|employees|"
                         r"head\s*count|average\s+ctc)\s*(?:is|are|of|was)?\s*"
                         r"(>=|<=|>|<|=|above|over|more\s+than|greater\s+than|higher\s+than|at\s+least|"
                         r"below|under|less\s+than|lower\s+than|at\s+most|equal\s+to)(?:\s+th[ae]n)?"   # "below than"
                         r"\s*(?:rs\.?|₹)?\s*" + NUM_RX,
                         q, re.IGNORECASE):
        op = m.group(2).lower()
        op = ">" if op in (">", "above", "over") or op.startswith(("more", "greater", "higher")) else \
            ">=" if op in (">=",) or op.startswith("at least") else "<=" if op in ("<=",) or op.startswith("at most") else \
            "=" if op in ("=",) or op.startswith("equal") else "<"
        metric = _metric_of(m.group(1)) or "CTC"
        plan["having"].append((metric, op, _num(m.group(3), m.group(4))))
    # number before the thing: "more than 50 employees", "above 2 lakh ctc", "less than 20 people"
    for m in re.finditer(r"\b(above|over|more\s+than|greater\s+than|at\s+least|below|under|less\s+than|fewer\s+than|"
                         r"at\s+most)\s*(?:rs\.?|₹)?\s*" + NUM_RX + r"\s*(employees?|people|staff|heads?|ctc|margin|base\s*sales?|invoice)\b",
                         q, re.IGNORECASE):
        op = m.group(1).lower()
        op = ">" if op in ("above", "over") or op.startswith(("more", "greater")) else \
            ">=" if op.startswith("at least") else "<=" if op.startswith("at most") else "<"
        metric = "Employees" if re.match(r"employee|people|staff|head", m.group(4), re.I) else (_metric_of(m.group(4)) or "CTC")
        h = (metric, op, _num(m.group(2), m.group(3)))
        if h not in plan["having"]:
            plan["having"].append(h)
    # "work orders with zero / no / nil invoice"
    for m in re.finditer(r"\b(?:zero|no|nil|missing|blank)\s+(invoice|base\s*sales?|ctc|margin)\b", q, re.IGNORECASE):
        h = (_metric_of(m.group(1)) or "InvoiceAmount", "=", 0.0)
        if h not in plan["having"]:
            plan["having"].append(h)
    # same, word order reversed: "below ctc 40000", "more than margin 1 lakh", "under 40k ctc"
    for m in re.finditer(r"\b(above|over|more\s+than|greater\s+than|higher\s+than|at\s+least|below|under|less\s+than|"
                         r"lower\s+than|at\s+most)(?:\s+th[ae]n)?\s+(?:the\s+|a\s+)?"
                         r"(margin\s*%|margin|ctc|base\s*sales?|invoice|average\s+ctc)\s*(?:of|is|=)?\s*(?:rs\.?|₹)?\s*"
                         + NUM_RX, q, re.IGNORECASE):
        op = m.group(1).lower()
        op = ">" if op in ("above", "over") or op.startswith(("more", "greater", "higher")) else \
            ">=" if op.startswith("at least") else "<=" if op.startswith("at most") else "<"
        h = (_metric_of(m.group(2)) or "CTC", op, _num(m.group(3), m.group(4)))
        if h not in plan["having"]:
            plan["having"].append(h)
    if re.search(r"\b(margin|profit)\s+(?:is\s+|are\s+|was\s+)?(negative|in\s+loss|below\s+zero)|\bnegative\s+margin|\bloss\b",
                 q, re.IGNORECASE):
        if re.search(r"\bhighlight|\bflag|\bmark", q, re.IGNORECASE):
            plan["highlight_negative"] = True
        else:
            plan["having"].append(("Margin", "<", 0.0))
    # sub-questions: "which workorder has the highest employee count", "tell me the highest margin one"
    for m in re.finditer(r"\bwhich\s+((?:[\w%]+\s+){0,2}?[\w%]+)\s+(?:has|had|have|is|was|got|with|gives?)?\s*(?:the\s+)?"
                         r"(highest|lowest|most|least|max\w*|min\w*|top|biggest|smallest|largest|best|worst)\s+"
                         r"([a-z%\s]+?)(?=\?|,|\.|$|\band\b|\bin\b|\bfor\b|\bof\b)", q, re.IGNORECASE):
        dim = next((k for k, v in PLAN_DIMS.items() if re.search(rf"\b(?:{v})\b", m.group(1), re.IGNORECASE)), None)
        metric = _metric_of(m.group(3)) or ("Employees" if re.search(r"employee|people|staff", m.group(3), re.I) else None)
        if metric:
            low = m.group(2).lower() in ("lowest", "least", "smallest", "worst") or m.group(2).lower().startswith("min")
            plan["which"].append({"dim": dim, "metric": metric, "low": low})
    # "department with the lowest average ctc", "work orders having highest margin %"
    for m in re.finditer(r"\b(work\s*-?\s*orders?|workorders?|departments?|depts?|categor(?:y|ies)|employees?)\s+"
                         r"(?:with|having|that\s+ha(?:s|ve)|which\s+ha(?:s|ve))\s+(?:the\s+)?"
                         r"(highest|lowest|most|least|max\w*|min\w*|biggest|smallest|best|worst)\s+"
                         r"([a-z%\s]{2,25}?)(?=\?|,|\.|$|\bin\b|\bfor\b|\bof\b)", q, re.IGNORECASE):
        dim = next((k for k, v in PLAN_DIMS.items() if re.search(rf"\b(?:{v})\b", m.group(1), re.IGNORECASE)), None)
        metric = _metric_of(m.group(3)) or ("Employees" if re.search(r"employee|people|staff|headcount", m.group(3), re.I) else None)
        if dim and metric and not plan["which"]:
            low = m.group(2).lower() in ("lowest", "least", "smallest", "worst") or m.group(2).lower().startswith("min")
            plan["which"].append({"dim": dim, "metric": metric, "low": low})
    for m in re.finditer(r"\b(?:the\s+)?(highest|lowest)\s+([a-z%\s]{2,25}?)\s+one\b", q, re.IGNORECASE):
        metric = _metric_of(m.group(2))
        if metric:
            plan["which"].append({"dim": None, "metric": metric, "low": m.group(1).lower() == "lowest"})
    for m in re.finditer(r"\bhow\s+many\s+(work\s*orders?|workorders?|departments?|categor\w+)", q, re.IGNORECASE):
        dim = next((k for k, v in PLAN_DIMS.items() if re.search(rf"\b(?:{v})\b", m.group(1), re.IGNORECASE)), None)
        if dim:
            plan["how_many"].append(dim)
    # sorting / top N
    top = re.search(r"\b(?:top|bottom|first|highest|lowest)\s+(\d{1,4})\b", q, re.IGNORECASE)
    if top:
        plan["limit"] = int(top.group(1))
    srt = re.search(r"\b(?:sort(?:ed)?|order(?:ed)?|rank(?:ed)?|arrange[d]?)\s+(?:it\s+|them\s+|by\s+|on\s+|as\s+per\s+)*"
                    r"(?:the\s+)?(highest|lowest|max\w*|min\w*|descending|ascending|desc|asc)?\s*([a-z%\s]{2,25}?)"
                    r"(?:\s+(first|on\s+top|desc\w*|asc\w*|high\s+to\s+low|low\s+to\s+high))?(?=\s*$|,|\.|\band\b|\bwith\b)",
                    q, re.IGNORECASE)
    if srt:
        metric = _metric_of(srt.group(2) or "")
        d = " ".join(x for x in (srt.group(1), srt.group(3)) if x).lower()
        if metric:
            plan["sort"] = (metric, bool(re.search(r"lowest|min|asc|low\s+to\s+high", d)))
    if not plan["sort"]:
        t = re.search(r"\b(?:top|highest|lowest|bottom|biggest|smallest)\s+(?:\d+\s+)?(?:[\w%]+\s+){0,3}?(?:by|on|in)\s+"
                      r"([a-z%\s]{2,25}?)(?=\s*$|,|\.|\band\b|\bfor\b|\bin\b|\bwith\b|\binclud)", q, re.IGNORECASE)
        if t and _metric_of(t.group(1)):
            plan["sort"] = (_metric_of(t.group(1)), bool(re.search(r"\b(lowest|bottom|smallest)\b", t.group(0), re.I)))
    plan["compare"] = bool(re.search(r"\bcompare|\bcomparison|\bvs\.?\b|\bversus\b|month[\s\-]*on[\s\-]*month|\bmom\b|"
                                     r"\bgrowth\b|\bchange\b|\bdifference\b|" + MOVE_RX.pattern, q, re.IGNORECASE))
    mv = MOVE_RX.search(q)                                 # "margin dropped from aug to sep" -> only the rows that fell
    plan["direction"] = None if not mv else ("down" if re.match(r"(drop|fell|fall|decreas|declin|reduc|went\s+down|lower)",
                                                                mv.group(0), re.I) else "up")
    if plan["compare"] and not plan["group"]:             # "compare sep and aug workorder margin %" = per work order
        found = [(m_.start(), d) for d, rx in (("WorkOrderNo", r"\bwork\s*-?\s*orders?\b|\bworkorders?\b"),
                                               ("Department", r"\bdepartments?\b|\bdepts?\b"),
                                               ("Category", r"\bcategor(?:y|ies)\b"))
                 for m_ in [re.search(rx, q, re.IGNORECASE)] if m_]
        if found:
            plan["group"] = [min(found)[1]]
    return plan


def wants_plan(question):
    """Use the planner for questions with more than one part (columns, conditions, patterns, sub-questions...)."""
    if not re.search(r"work\s*orders?|workorders?|\bwo\b|department|dept|categor|employee|ctc|margin|invoice|base\s*sales?",
                     question, re.IGNORECASE):
        return False
    plan = parse_plan(question)
    parts = 0
    parts += bool(plan["wo"] and any(c in _plan_metrics() or c in PLAN_ATTRS for c in plan["cols"]) or
                  plan["wo"] and re.search(r"wise|\bby\b|each|per\b", question, re.I) and
                  any("%" in p_ for p_ in plan["wo"]))
    parts += len(plan["group"]) >= 2
    parts += bool(plan["having"]) + bool(plan["which"]) + bool(plan["how_many"]) + plan["highlight_negative"]
    parts += len([c for c in plan["cols"] if c in _plan_metrics() or c in PLAN_ATTRS]) >= 4
    parts += bool(re.search(r"\bcolumns?\b|\bfields?\b|\bcontain", question, re.I) and plan["cols"])
    parts += bool(plan["compare"] and plan["group"] and len(parse_periods(question)) >= 2)
    if any(g_ in plan["group"] for g_ in ("Department", "Employee")):
        parts += 1                                         # department / employee wise = per salary month
    if current_prefs().get("preferred_columns") and plan["group"]:
        return True
    if re.search(r"\bresign\w*|\bjoin\w*|\bleft\b|\b(?:start|end)\s*date", question, re.IGNORECASE):
        return False                                       # date-column questions: leavers / joiners / AI analyst
    if "AvgCTC" in plan["cols"]:                           # "average ctc (per employee) (by category)" - exact, no AI
        return True
    if plan["which"] and plan["which"][0].get("dim") in ("Department", "Category"):   # "which category gives highest
        return True                                                                    # margin %" (work orders: own report)
    strong = bool(plan["wo"] or plan["having"] or len(plan["group"]) >= 2 or plan["highlight_negative"] or
                  any(g_ in plan["group"] for g_ in ("Department", "Employee")) or
                  len([c for c in plan["cols"] if c in _plan_metrics() or c in PLAN_ATTRS]) >= 4 or
                  re.search(r"\bcolumns?\b|\bfields?\b|\bcontain", question, re.I) or
                  (plan["compare"] and plan["group"]) or
                  (re.search(r"\b(?:its|their|along\s+with|together\s+with|include|including)\b", question, re.I)
                   and len([c for c in plan["cols"] if c in _plan_metrics()]) >= 2))
    return (parts >= 2 or (parts >= 1 and strong)) and bool(PLAN_COMPLEX.search(question) or strong)


def ai_plan_fill(question, plan, leftover):
    """Qwen reads the full question and fills only what the rules could not read. Every value is checked."""
    allowed = _plan_metrics() + PLAN_ATTRS + ["WorkOrderNo", "Month", "Employee"]
    system = ("You turn a payroll question into a JSON query plan. Reply ONLY with JSON: "
              '{"wo_like": [patterns like "M%"], "group_by": [dimensions], "columns": [columns], '
              '"having": [{"metric": m, "op": ">"|"<"|">="|"<="|"=", "value": number}], '
              '"sort": {"metric": m, "desc": true|false} or null, "limit": number or null, '
              '"which": [{"dim": dimension or null, "metric": m, "lowest": true|false}]}. '
              f"Dimensions: WorkOrderNo, Department, Category, Employee, Month. Columns/metrics: {', '.join(allowed)}. "
              "A work order filter like 'M%' or 'starting with M' is a FILTER (wo_like), never a column. "
              "Only include what the question asks for; leave the rest empty.")
    if taught_prompt():
        system += "\n" + taught_prompt()
    try:
        resp = ollama.chat(model=MODEL, messages=[{"role": "system", "content": system},
                                                  {"role": "user", "content": question}],
                           format="json", keep_alive="60m",
                           options={"temperature": 0, "num_predict": 220, "num_ctx": 2048, "num_thread": AI_THREADS})
        data = json.loads(resp["message"]["content"])
    except Exception as e:
        return plan, f"AI not reachable ({e.__class__.__name__})"
    got = []
    for p_ in data.get("wo_like") or []:
        p_ = str(p_).upper().strip()
        if re.fullmatch(r"[A-Z0-9/\-_.()% ]{1,80}", p_) and p_ not in plan["wo"] and "%" in p_:
            plan["wo"].append(p_)
            got.append(f"work order {p_}")
    for d in data.get("group_by") or []:
        if d in PLAN_DIMS and d not in plan["group"] and re.search(rf"\b(?:{PLAN_DIMS[d]})\b", question, re.I):   # never invented
            plan["group"].append(d)
            got.append(f"{d}-wise")
    col_rx = {col: pat for pat, col in PLAN_COLS}
    for c in data.get("columns") or []:                  # only columns the question actually names (never invented)
        if c in allowed and c not in plan["cols"] and (c in _taught() or (c in col_rx and re.search(col_rx[c], question, re.I))):
            plan["cols"].append(c)
            got.append(c)
    for h in data.get("having") or []:
        try:
            if h["metric"] in _plan_metrics() and h["op"] in (">", "<", ">=", "<=", "=") and \
                    not any(x[0] == h["metric"] for x in plan["having"]):          # the rules already read this one
                plan["having"].append((h["metric"], h["op"], float(h["value"])))
                got.append(f"{h['metric']} {h['op']} {h['value']}")
        except Exception:
            pass
    srt = data.get("sort") or {}
    if not plan["sort"] and isinstance(srt, dict) and srt.get("metric") in _plan_metrics():
        plan["sort"] = (srt["metric"], not srt.get("desc", True))
    if not plan["limit"] and str(data.get("limit") or "").isdigit():
        plan["limit"] = min(int(data["limit"]), ROW_LIMIT)
    for w in data.get("which") or []:
        if isinstance(w, dict) and w.get("metric") in _plan_metrics() and not plan["which"]:
            plan["which"].append({"dim": w.get("dim") if w.get("dim") in PLAN_DIMS else None,
                                  "metric": w["metric"], "low": bool(w.get("lowest"))})
    return plan, ("Qwen read: " + ", ".join(got)) if got else ""


PLAN_NAMES = {"WorkOrderNo": "Work order", "Department": "Department", "Category": "Category", "Employee": "Employee",
              "Month": "Month", "Employees": "Employee count", "CTC": "CTC", "AvgCTC": "Average CTC",
              "BaseSalesValue": "Base sales value", "InvoiceAmount": "Invoice amount", "Margin": "Margin",
              "MarginPct": "Margin %", "WOStartDate": "WO start date", "WOEndDate": "WO end date"}


def plan_answer(question, mode, notes, t0, use_ai=True):
    plan = parse_plan(question)
    # anything the rules could not place -> Qwen reads the full question (only then)
    vocab = {"give", "show", "me", "detail", "details", "data", "column", "columns", "contain", "contains", "containing",
             "include", "including", "field", "fields", "wise", "with", "and", "also", "tell", "which", "highest",
             "lowest", "sort", "sorted", "order", "first", "top", "where", "having", "above", "below", "more", "less",
             "than", "greater", "lakh", "lakhs", "crore", "start", "starting", "starts", "begin", "beginning", "number",
             "count", "value", "amount", "total", "deployed", "many", "how", "highlight", "negative", "compare",
             "change", "growth", "difference", "one", "its", "their", "what", "are", "were", "was", "the", "for",
             "workorder", "workorders", "work", "order", "orders", "employee", "employees", "month", "months",
             "department", "departments", "category", "per", "each", "group", "grouped", "list", "report", "only",
             "under", "percentage", "percent", "calculate", "calculated", "its", "had", "has", "have", "deploy",
             "mom", "versus", "comparison", "rows", "row", "flag", "mark", "highlighted", "descending", "ascending",
             "average", "avg", "mean", "head", "person", "liable", "payment", "paid", "making", "loss", "losses",
             "profitable", "gives", "giving", "earns", "brings", "running", "active", "zero", "nil", "no", "blank",
             "missing", "dropped", "fell", "grew", "increased", "decreased", "most", "least", "best", "worst",
             "performing"}
    left = {w for w in unknown_words(question, {}) - vocab - FILTER_IGNORE
            if not re.fullmatch(r"\d+|[a-z]\d*", w) and not any(w in p_.lower() for p_ in plan["wo"])}
    for v in (find_known(question, "Department"), find_known(question, "Category")):
        if v:
            vc = re.sub(r"[^a-z0-9]", "", v.lower())
            left = {w for w in left if w not in vc and w not in v.lower()}
    ai_note = ""
    if left and use_ai:
        plan, ai_note = ai_plan_fill(question, plan, left)
    explicit_cols = bool(plan["cols"])
    prefs = current_prefs()
    key2plan = {"EmployeeCount": "Employees", "Employee": "Employee"}
    if not explicit_cols and prefs.get("preferred_columns"):                # level 4: "only these columns"
        plan["cols"] = [key2plan.get(k, k) for k in prefs["preferred_columns"]
                        if key2plan.get(k, k) in _plan_metrics() + PLAN_ATTRS + ["WorkOrderNo", "Month", "Employee"]]
    for k in prefs.get("always_columns", []):
        k = key2plan.get(k, k)
        if k in _plan_metrics() + PLAN_ATTRS and k not in plan["cols"] and plan["cols"]:
            plan["cols"].append(k)
    f = {"department": find_known(question, "Department"), "category": find_known(question, "Category")}
    if left:
        fn = []
        f = understand_filters(question, f, fn, use_ai, extra_ignore=vocab | {p_.lower().strip("%") for p_ in plan["wo"]})
        notes += [n for n in fn if not (ai_note and n.startswith("⚠️"))]
    periods = parse_periods(question)
    if not periods:
        periods = [latest_period(mode)]
        notes.append("No month mentioned, showing the latest month in the data.")
    if len(periods) >= 2 and "Month" not in plan["group"] and not plan["compare"]:
        plan["group"].append("Month")                      # several months -> one row per month
    if "Month" in plan["cols"] and "Month" not in plan["group"] and not plan["compare"]:
        plan["group"].append("Month")
    if not plan["group"] and plan["which"] and plan["which"][0]["dim"] in PLAN_DIMS:
        plan["group"] = [plan["which"][0]["dim"]]
    if not plan["group"] and plan["having"]:              # a condition on "which X ..." is checked per X
        found = []
        for dim, rx in (("WorkOrderNo", r"\bwork\s*-?\s*orders?\b|\bworkorders?\b|\bwos\b"),
                        ("Department", r"\bdepartments?\b|\bdepts?\b|\bclients?\b"), ("Category", r"\bcategor(?:y|ies)\b"),
                        ("Employee", r"\b(?:employees?|emps?|staff|people)\b")):
            m_ = re.search(rx, question, re.I)
            if m_:
                found.append((m_.start(), dim))
        if found:                                          # the thing named FIRST is what is listed:
            dim = min(found)[1]                            # "employees ... also department column" = per employee
            plan["group"] = [dim]
            if re.search(r"\bhow\s+many\b|\bcount\b|\bnumber\s+of\b", question, re.I) and dim not in plan["how_many"]:
                plan["how_many"].append(dim)
    if not plan["group"]:
        plan["group"] = ["WorkOrderNo"] if (plan["wo"] or "WorkOrderNo" in plan["cols"]) else \
            ["Department"] if "Department" in plan["cols"] else []
    if plan["compare"] and len(periods) >= 2:
        plan["group"] = [g for g in plan["group"] if g != "Month"]
    elif "WorkOrderNo" in plan["group"] \
            and "Month" not in plan["group"] \
            and not (len(periods) == 1 and {c for c in plan["cols"] if c in _plan_metrics()} and
                     {c for c in plan["cols"] if c in _plan_metrics()} <= {"AvgCTC", "Employees"}) \
            and not (len(periods) == 1 and plan["which"]) \
            and not (len(periods) == 1 and plan["having"]):     # "margin % below 5": checked per whole work order
        plan["group"].append("Month")                     # "workorder wise" = per work order per salary month
    metrics = [c for c in plan["cols"] if c in _plan_metrics()] or (
        ["CTC"] if "Employee" in plan["group"] else list(PLAN_DEFAULT))   # per-employee lists: CTC, not WO margins
    for h in plan["having"]:
        if h[0] not in metrics:
            metrics.append(h[0])
        if h[0] in ("Margin", "MarginPct"):                 # a margin condition is read with what makes the margin
            metrics[:0] = [c for c in ("BaseSalesValue", "CTC") if c not in metrics]
    for w in plan["which"]:
        if w["metric"] not in metrics:
            metrics.append(w["metric"])
    if plan["sort"] and plan["sort"][0] not in metrics:
        metrics.append(plan["sort"][0])
    attrs = [c for c in plan["cols"] if c in PLAN_ATTRS and c not in plan["group"]]
    # ---------- SQL: rows -> invoice/base spread once per WO + salary month, then grouped ----------
    Y, M = period_exprs(mode)
    where, params = build_where(mode, periods, f)
    if plan["wo"]:
        cl = " OR ".join("[WorkOrderNo] LIKE ?" if "%" in p_ else "[WorkOrderNo] = ?" for p_ in plan["wo"])
        where = (where + " AND " if where else "WHERE ") + f"({cl})"
        params += plan["wo"]
    part = "PARTITION BY [WorkOrderNo], [Year], [Month]"
    sel, grp = [], []
    # "each month" = the month type the user filtered by (liable / payment / salary); only work order-wise lists keep
    # the SALARY month, because each salary month of a work order has its own invoice
    month_by_mode = "Month" in plan["group"] and "WorkOrderNo" not in plan["group"]
    month_col = PERIOD_COL.get(mode, "Salary month") if month_by_mode else "Salary month"
    for g_ in plan["group"]:
        if g_ == "Month" and month_by_mode:
            sel += [f"{Y} AS [PY]", f"{M} AS [PM]"]
            grp += [Y, M]
        elif g_ == "Month":                                # with work order / dept: Month = SALARY month (Year / Month)
            sel += ["[Year] AS [PY]", "[Month] AS [PM]"]
            grp += ["[Year]", "[Month]"]
        elif g_ == "Employee":
            sel += ["[Employee]", "MAX([Name]) AS [Name]"]
            grp.append("[Employee]")
        else:
            sel.append(f"[{g_}]")
            grp.append(f"[{g_}]")
    if plan["compare"] and len(periods) >= 2:              # compare the months you filtered on
        sel += [f"{Y} AS [FY]", f"{M} AS [FM]"]
        grp += [Y, M]
    for a in attrs:
        sel.append(f"MAX([{a}]) AS [{a}]")
    if "WorkOrderNo" in plan["cols"] and "WorkOrderNo" not in plan["group"]:
        sel.append("COUNT(DISTINCT [WorkOrderNo]) AS [WorkOrders]")
    sel += ["COUNT(DISTINCT [Employee]) AS [Employees]", "SUM([CTC]) AS [CTC]",
            "ROUND(SUM([_Base]), 2) AS [BaseSalesValue]", "ROUND(SUM([_Inv]), 2) AS [InvoiceAmount]"]
    tm = [c for c in metrics if c in _taught()]       # v14: taught formulas -> their group expression
    for k in tm:
        sel.append(f"ROUND({_taught()[k]['agg_expr']}, 2) AS [{k}]")
    pcols = (f",\n      MAX({BASE_EXPR}) OVER ({part}) * 1.0 / COUNT(*) OVER ({part}) AS [BaseSales], "
             f"DATEFROMPARTS([Year], [Month], 1) AS [SalaryMonth], "
             f"DATEFROMPARTS(YEAR([month_year]), MONTH([month_year]), 1) AS [MYMonth], {TAKEHOME_EXPR} AS [TakeHome]"
             ) if tm else ""
    sql = (f"SELECT {', '.join(sel)}\nFROM (SELECT *, MAX({BASE_EXPR}) OVER ({part}) * 1.0 / COUNT(*) OVER ({part}) AS [_Base],\n"
           f"      MAX([InvoiceAmount]) OVER ({part}) * 1.0 / COUNT(*) OVER ({part}) AS [_Inv]{pcols}\n"
           f"      FROM {T} {where}) AS r" + (f"\nGROUP BY {', '.join(grp)}" if grp else ""))
    t1 = time.time()
    df = run_query(sql, params)
    t_sql = time.time() - t1
    # ---------- calculations, conditions, sorting (in Python, exact) ----------
    sales = pct_basis(question) == "sales"
    df["Margin"] = df["BaseSalesValue"] - df["CTC"]
    den = df["BaseSalesValue"] if sales else df["CTC"]
    df["MarginPct"] = (df["Margin"] * 100 / den.where(den != 0)).round(2)
    df["AvgCTC"] = (df["CTC"] / df["Employees"].where(df["Employees"] != 0)).round(2)
    if "PY" in df:
        df["Month"] = [period_label(int(m_), int(y_)) for m_, y_ in zip(df["PM"], df["PY"])]
        df["_ym"] = df["PY"] * 100 + df["PM"]
    if "FY" in df:
        df["_fy"] = df["FY"] * 100 + df["FM"]
    lines = []
    # comparison: each month side by side + change
    if plan["compare"] and len(periods) >= 2 and "_fy" in df and df["_fy"].nunique() < 2:
        have = {int(x) for x in df["_fy"].unique()}
        miss = [period_label(m_, y_) for m_, y_ in periods if y_ * 100 + m_ not in have]
        notes.append(f"⚠️ No data for {', '.join(miss)} - nothing to compare, showing the month that has data.")
        plan["compare"] = False
        if "Month" not in plan["group"]:
            plan["group"].append("Month")
    if plan["compare"] and len(periods) >= 2 and "_fy" in df:
        keys = [("Name" if g_ == "Employee" else g_) for g_ in plan["group"]] or ["_all"]
        if keys == ["_all"]:
            df["_all"] = "Total"
        first, last = df["_fy"].min(), df["_fy"].max()
        fl, ll = period_label(first % 100, first // 100), period_label(last % 100, last // 100)
        wide = None
        for ym in sorted(df["_fy"].unique()):
            lab = period_label(ym % 100, ym // 100)
            part_ = df[df["_fy"] == ym].set_index(keys)[metrics].add_suffix(f" {lab}")
            wide = part_ if wide is None else wide.join(part_, how="outer")
        for mtr in metrics:
            a_, b_ = f"{mtr} {fl}", f"{mtr} {ll}"
            if a_ in wide and b_ in wide and mtr != "MarginPct":
                wide[f"{mtr} change"] = wide[b_].fillna(0) - wide[a_].fillna(0)
                wide[f"{mtr} change %"] = (wide[f"{mtr} change"] * 100 / wide[a_].where(wide[a_] != 0)).round(2)
            elif a_ in wide and b_ in wide:                        # margin %: change in percentage POINTS
                wide["Margin % change (pts)"] = (wide[b_] - wide[a_]).round(2)
                ch = wide["Margin % change (pts)"].dropna()
                if len(ch):
                    who = "work orders" if keys == ["WorkOrderNo"] else (keys[0].lower() + "s" if keys != ["_all"] else "")
                    up, dn = ch.idxmax(), ch.idxmin()
                    lab = lambda k_: k_ if not isinstance(k_, tuple) else " · ".join(map(str, k_))
                    lines.append(f"- **Margin % {fl} → {ll}:** {int((ch > 0).sum())} {who} improved, "
                                 f"{int((ch < 0).sum())} dropped · biggest rise **{lab(up)}** ({ch[up]:+.2f} pts) · "
                                 f"biggest drop **{lab(dn)}** ({ch[dn]:+.2f} pts)")
                    wide = wide.sort_values("Margin % change (pts)", na_position="last")
        df = wide.reset_index()
        if plan.get("direction") and keys != ["_all"]:     # "dropped" / "increased": keep only the rows that moved that way
            chg_col = next((c for c in (f"{metrics[0]} change", "Margin % change (pts)") if c in df), None)
            if chg_col:
                mtr0 = "MarginPct" if chg_col.startswith("Margin %") else metrics[0]
                a0, b0 = f"{mtr0} {fl}", f"{mtr0} {ll}"
                both, only_a, only_b = pd.Series(True, index=df.index), 0, 0
                if a0 in df and b0 in df:
                    both = df[a0].notna() & df[b0].notna()
                    only_a, only_b = int((df[a0].notna() & df[b0].isna()).sum()), int((df[a0].isna() & df[b0].notna()).sum())
                df = df[both]                              # a work order that stopped did not "go down" - it ended
                before_n = len(df)
                df = df[df[chg_col] < 0] if plan["direction"] == "down" else df[df[chg_col] > 0]
                df = df.sort_values(chg_col, ascending=plan["direction"] == "down")
                lines.insert(0, f"- **{len(df):,} of {before_n:,}** (in both months) "
                                f"{'went down' if plan['direction'] == 'down' else 'went up'} "
                                f"from {fl} to {ll} (by {chg_col.replace(' change', '').replace(' (pts)', '')}). "
                                f"Not counted: {only_a:,} only in {fl} (stopped), {only_b:,} only in {ll} (new).")
        for mtr in metrics:
            a_, b_ = f"{mtr} {fl}", f"{mtr} {ll}"
            if a_ in df and b_ in df and mtr != "MarginPct":
                ta, tb = df[a_].sum(), df[b_].sum()
                chg = f"{(tb - ta) * 100 / ta:+.1f}%" if ta else ""
                fmt = (lambda v: f"{int(v):,}") if mtr == "Employees" else inr
                lines.append(f"- **{_pname(mtr)}:** {fl} {fmt(ta)} → {ll} {fmt(tb)} ({chg})")
        out_cols = [c for c in df.columns if c != "_all"]
        sort_col = None
    else:
        for metric, op, val in plan["having"]:
            s_ = pd.to_numeric(df[metric], errors="coerce").fillna(0)    # no invoice (blank) counts as 0
            df = df[{">": s_ > val, "<": s_ < val, ">=": s_ >= val, "<=": s_ <= val, "=": s_ == val}[op]]
        if plan["highlight_negative"]:
            df["Flag"] = ["⚠️ negative margin" if m_ < 0 else "" for m_ in df["Margin"]]
        sort_col, asc = (plan["sort"] or (None, False))
        if not sort_col and plan["which"]:
            sort_col, asc = plan["which"][0]["metric"], plan["which"][0]["low"]
        if sort_col:
            df = df.sort_values(sort_col, ascending=asc, kind="stable")
        else:                                              # no order asked -> by month, then name / number
            keys_ = [k for k in [g_ for g_ in plan["group"] if g_ != "Month"] + (["_ym"] if "_ym" in df else [])
                     if k in df]
            if keys_:
                df = df.sort_values(keys_, kind="stable")
        if plan["limit"]:
            df = df.head(plan["limit"])
        dims = []
        for g_ in plan["group"]:
            dims += ["Employee", "Name"] if g_ == "Employee" else [g_]
        order = []
        for c in plan["cols"]:                                    # the user's own column order
            if c == "Employee":
                c = "Name"
            if c in df.columns and c not in order:
                order.append(c)
        for c in dims + (["WorkOrders"] if "WorkOrders" in df else []) + attrs + metrics:
            if c in df.columns and c not in order:
                order.insert(len([x for x in order if x in dims]) if c in dims else len(order), c)
        if "Flag" in df:
            order.append("Flag")
        out_cols = order
    df = df[[c for c in out_cols if c in df.columns]].reset_index(drop=True)
    # ---------- answers to the sub-questions ----------
    for w in plan["which"]:
        mtr = w["metric"]
        if mtr not in df or df.empty:
            continue
        r_ = df.loc[df[mtr].idxmin() if w["low"] else df[mtr].idxmax()]
        label = " · ".join(str(r_[c]) for c in df.columns if c in ("WorkOrderNo", "Department", "Category", "Name", "Month")
                           and pd.notna(r_[c])) or "Total"
        val = f"{r_[mtr]:.2f}%" if mtr == "MarginPct" else (f"{int(r_[mtr]):,}" if mtr == "Employees" else
                                                             f"{r_[mtr]:,.2f}" if mtr in _taught() else inr(r_[mtr]))
        extra = [f"{_pname(c)} {f'{int(r_[c]):,}' if c == 'Employees' else (f'{r_[c]:.2f}%' if c == 'MarginPct' else inr(r_[c]))}"
                 for c in metrics if c != mtr and c in df.columns and pd.notna(r_[c])]
        nm = _pname(mtr) if _pname(mtr).isupper() else _pname(mtr).lower()
        lines.insert(0, f"- **{'Lowest' if w['low'] else 'Highest'} {nm}: {label} — {val}**"
                        + (f" ({' · '.join(extra)})" if extra else ""))
    for dim in plan["how_many"]:
        if dim in df:
            lines.append(f"- **Number of {'work orders' if dim == 'WorkOrderNo' else _pname(dim).lower() + 's'}:** "
                         f"{df[dim].nunique():,}")
    if "Flag" in df:
        lines.append(f"- ⚠️ **{int((df['Flag'] != '').sum())}** row(s) with negative margin are flagged.")
    # ---------- what was understood (so a wrong filter is visible at once) ----------
    ptxt = ", ".join(period_label(m_, y_) for m_, y_ in periods)
    und = [f"{mode_label(mode)} = {ptxt}"]
    if plan["wo"]:
        und.append("work order " + " / ".join(("starts with " + p_[:-1]) if p_.endswith("%") and "%" not in p_[:-1] else
                                              ("like " + p_) if "%" in p_ else ("= " + p_) for p_ in plan["wo"]))
    for v in (f.get("category"), f.get("department")):
        if v:
            und.append(f"only {v}")
    inc = current_exclusions()
    if inc.get("include_Category") or inc.get("include_Department"):
        und.append("only " + ", ".join(inc.get("include_Category", []) + inc.get("include_Department", [])))
    if excluded_text():
        und.append("excluding " + excluded_text())
    for metric, op, val in plan["having"]:
        und.append(f"{_pname(metric)} {op} " + (inr(val) if metric not in ("Employees", "MarginPct") else f"{val:g}"))
    grp_txt = " + ".join(_pname(g_) for g_ in plan["group"]) or "overall total"
    head = f"{grp_txt}-wise for {ptxt}" if plan["group"] else f"Total for {ptxt}"
    title = f"**{head}" + (" — " + "; ".join(und[1:]) if len(und) > 1 else "") + ":**"
    tot = []
    if not plan["compare"] or len(periods) < 2:
        if "AvgCTC" in df and len(df) == 1:
            tot.append(f"Average CTC per employee {inr(df['AvgCTC'].iloc[0])}")
        if "CTC" in df:
            tot.append(f"CTC {inr(df['CTC'].sum())}")
        if "BaseSalesValue" in df:
            tot.append(f"Base sales value {inr(df['BaseSalesValue'].sum())}")
        if "InvoiceAmount" in df:
            tot.append(f"Invoice {inr(df['InvoiceAmount'].sum())}")
        if "Margin" in df and "BaseSalesValue" in df and "CTC" in df:
            bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
            dd = bs if sales else ct
            tot.append(f"Margin {inr(bs - ct)}" + (f" ({(bs - ct) * 100 / dd:.2f}%)" if dd else ""))
        elif "Margin" in df:                            # "department wise margin": margin column only
            tot.append(f"Margin {inr(pd.to_numeric(df['Margin'], errors='coerce').sum())}")
    text = f"{title} {len(df):,} row(s)" + (" · " + " · ".join(tot) if tot else "")
    if lines:
        text += "\n\n" + "\n".join(lines)
    if df.empty:
        text = f"{title} no matching rows."
    missing = [PLAN_NAMES.get(c, c) for c in plan["cols"] if c not in ("Employee",) and
               not any(str(x).startswith(c) or str(x) == c for x in df.columns) and not plan["compare"]]
    notes.insert(0, "Understood as: **Query plan** · Filters: " + "; ".join(und) + f" · Grouped by: {grp_txt}"
                 + f" · Columns: {', '.join(PLAN_NAMES.get(c, c) for c in df.columns if c in PLAN_NAMES) or '-'}"
                 + (f" · Sorted by {_pname(plan['sort'][0])} {'low→high' if plan['sort'][1] else 'high→low'}"
                    if plan["sort"] else "") + (f" · Top {plan['limit']}" if plan["limit"] else "") + ".")
    if ai_note:
        notes.append(f"🧠 {ai_note}.")
    if missing:
        notes.append(f"⚠️ Not available: {', '.join(missing)}.")
    if "Month" in plan["group"] and not month_by_mode:
        notes.append(f"**{month_col}** = Year/Month columns; rows filtered by {mode_label(mode)} = {ptxt}.")
    notes.append(f"Invoice / base sales counted once per work order per salary month; base sales = Invoice × 100 / "
                 f"{100 + GST_RATE}. ⏱️ SQL: {t_sql:.1f}s")
    df = df.rename(columns={"Employees": "Employee count", "AvgCTC": "Average CTC", "InvoiceAmount": "Invoice amount",
                            "WorkOrders": "Work orders", "WOStartDate": "WO start date", "WOEndDate": "WO end date",
                            "Month": month_col, **{k: _taught()[k]["label"] for k in tm}})
    if month_by_mode and month_col in df and plan["group"] == ["Month"]:   # month-wise table: the month first
        df.insert(0, month_col, df.pop(month_col))
    intent = {"report": "plan", "plan": {k: v for k, v in plan.items()}, "pct_basis": "sales" if sales else "ctc",
              "month_col": month_col,
              "explicit_cols": explicit_cols,
              "excluded": excluded_text()}
    return {"text": text, "note": " ".join(notes), "df": df, "report": "plan", "sql": sql, "params": params,
            "intent": intent}


# ======================= POLICY CONTEXT (who is asking, what they may see) =======================
POLICY_STORE = None          # set by the app (needs the app folder)
PREFS_SAVER = None           # set by the app: saves data preferences for the signed-in user only
USERS_LOADER = None          # set by the app: returns users.json
VOCAB = pe.Vocab()


def set_policy_context(uid, role, rec=None):
    """Loads levels 2-4 for this user for the current request (thread).
    Roles in users.json: "master" = full powers (permissions, teaching, user management) - inside this file that is
    the "admin" role; "admin" = sees all data and may use "/" SQL, but may NOT change permissions or teach for others
    (inside this file treated as "user" + sql_ok); "user" = normal person."""
    global VOCAB
    raw = role
    role = "admin" if raw == "master" else ("user" if raw == "admin" else raw)
    _POL.sql_ok, _POL.raw_role = raw in ("master", "admin"), raw
    try:
        VOCAB = pe.Vocab([c for c, _ in _table_columns()])
    except Exception:
        pass
    data = POLICY_STORE.load() if POLICY_STORE else {"global": dict(pe.DEFAULT_GLOBAL), "users": {}}
    taught, hidden = {}, set()
    if TEACH_STORE is not None:                          # v14: taught formulas in this user's scope
        try:
            taught = tg.taught_formulas(TEACH_STORE, uid, [c for c, _ in _table_columns()])
            hidden = tg.hidden_formula_keys(TEACH_STORE, uid)
        except Exception:
            taught, hidden = {}, set()
    eff = pe.Effective(data, uid, role, VOCAB, table=TABLE, store=POLICY_STORE, taught=taught, taught_hidden=hidden)
    _POL.uid, _POL.role, _POL.eff, _POL.data = uid, role, eff, data
    _POL.taught = {k: v for k, v in taught.items() if eff.key_visible(k)}   # hidden if a source column is hidden
    _POL.prefs = pe.prefs_of(rec or {})
    _POL.null_cols = eff.sql_null_columns([c for c, _ in _table_columns()])
    _POL.row_filter = eff.row_filter                    # admin limited this user to some categories / departments
    return eff


def current_eff():
    return getattr(_POL, "eff", None)


def current_prefs():
    return getattr(_POL, "prefs", None) or dict(pe.DEFAULT_PREFS)


def audit_policy_event(action, detail=""):
    try:
        audit(getattr(_POL, "uid", "?"), action, detail)
    except Exception:
        pass


def _policy_reply(text, report="policy", df=None, options=None, note=""):
    return {"text": text, "note": note, "df": df, "report": report, "sql": "", "params": [], "intent": {},
            "options": options or []}


GROUP_OF = {"Total PF": "PF", "PF_Employee": "PF", "PF_Employer": "PF", "TakeHome": "Take-home", "NetSalary": "Take-home",
            "ChequeAmount": "Take-home", "PT": "PT", "ESIC": "ESIC", "ESIC_Employee": "ESIC", "GrossAmount": "Gross",
            "IncomeTax": "TDS"}


def policy_command(question):
    """Runs FIRST for every question: confirmations, admin questions, admin policy changes, saved preferences.
    Returns a reply dict, or None for a normal data question."""
    uid, role = getattr(_POL, "uid", None), getattr(_POL, "role", "user")
    if uid is None:
        return None
    q = question.strip()
    try:
        ss = st.session_state
    except Exception:
        ss = {}
    pending = ss.get("pending_policy") if hasattr(ss, "get") else None
    # ---- confirm / cancel a pending admin change ----
    if pending and pe.CONFIRM_RX.match(q):
        ss["pending_policy"] = None
        users = USERS_LOADER() if USERS_LOADER else {}
        if users.get(uid, {}).get("role") != "master" or pending.get("by") != uid:       # re-checked, not trusted
            audit_policy_event("policy_change_denied", pending.get("question", ""))
            return _policy_reply("🔒 Only the master administrator can change data permissions.", "blocked")
        act = pending["action"]
        done = []
        for target in (act["users"] or [None]):
            a = dict(act, user=target, scope="USER" if target else "GLOBAL")
            done.append(POLICY_STORE.apply(a, uid, reason=f"chat: {act.get('question', '')[:120]}"))
            if a.get("op") == "SET_TEACH" and TEACH_STORE is not None:
                TEACH_STORE.log(uid, "right_removed" if a["value"] == "none" else "right_granted", "",
                                f"user:{target}", "", a["value"], "chat")
        audit_policy_event("policy_changed", "; ".join(done))
        set_policy_context(uid, role, users.get(uid))
        return _policy_reply("✅ **Policy saved.** " + " · ".join(done) +
                             "\n\nIt applies from the next question of every affected user.", "policy",
                             note="Logged in Settings → Administration → Policy audit log.")
    if pending and pe.CANCEL_RX.match(q):
        ss["pending_policy"] = None
        return _policy_reply("Cancelled - nothing was changed.", "policy")
    users = USERS_LOADER() if USERS_LOADER else {}
    # ---- admin policy change (natural language) -> needs confirmation ----
    act = pe.parse_admin_command(q, VOCAB, users)
    if act:
        if role != "admin":
            audit_policy_event("policy_change_denied", q[:200])
            return _policy_reply("🔒 Only an **administrator** can change data permissions. Your own *display* "
                                 "preferences can be changed with e.g. *from now on never show UniqueNo*.", "blocked")
        if act.get("error"):
            return _policy_reply(act["error"], "policy")
        act["question"] = q
        ss["pending_policy"] = {"action": act, "by": uid, "question": q}
        ss["pending_teach"] = None
        return _policy_reply(pe.describe_action(act, users), "policy_confirm", options=["Confirm", "Cancel"],
                             note="Nothing changes until you confirm.")
    # ---- admin information questions ----
    if pe.is_admin_question(q) and not pe.parse_pref_command(q, VOCAB):
        eff = current_eff()
        if not (eff and eff.admin_info):
            audit_policy_event("admin_info_denied", q[:200])
            return _policy_reply("🔒 I can't share user, access or system administration details. "
                                 "Please ask your administrator.", "blocked")
        logins = []
        try:
            log = pd.read_csv(AUDIT_FILE)
            t = pd.to_datetime(log["time"], errors="coerce")
            logins = list(log[(log["action"] == "login") & (t >= pd.Timestamp.now() - pd.Timedelta(hours=1))]["user"])
        except Exception:
            logins = None
        text, rows = pe.admin_answer(q, users, POLICY_STORE.load(), POLICY_STORE, VOCAB, logins)
        return _policy_reply(text, "admin_info", df=pd.DataFrame(rows) if rows else None,
                             note="Administrator information · only visible to admins.")
    # ---- the user's own persistent preferences (level 4) ----
    cmd = pe.parse_pref_command(q, VOCAB)
    if cmd:
        prefs = current_prefs()
        eff = current_eff()
        if cmd["type"] == "show_prefs":
            lines = [f"- **Hidden columns:** {', '.join(pe.label(c) for c in prefs['hidden_columns']) or 'none'}",
                     f"- **Only these columns:** {', '.join(pe.label(c) for c in prefs['preferred_columns']) or 'not set (normal columns)'}",
                     f"- **Always include:** {', '.join(pe.label(c) for c in prefs['always_columns']) or 'none'}",
                     f"- **Employee-level (raw) data:** {'shown when needed' if prefs['raw_data'] else 'only when I ask'}",
                     f"- **Charts:** {'on' if prefs['charts'] else 'off'}"]
            conflict = [c for c in prefs["preferred_columns"] + prefs["always_columns"] if eff and not eff.key_visible(c)]
            if conflict:
                lines.append(f"- 🔒 {', '.join(pe.label(c) for c in conflict)}: your administrator has restricted "
                             "this, so your preference cannot override it.")
            return _policy_reply("**Your saved data preferences**\n\n" + "\n".join(lines), "pref",
                                 note="Change them in Settings → My data preferences, or just tell me.")
        new, msg = pe.apply_pref(prefs, cmd)
        blocked = [c for c in cmd.get("cols", []) if eff and not eff.key_visible(c)]
        if blocked and cmd["type"] in ("show", "only"):
            msg += (f"\n\n🔒 **{', '.join(pe.label(c) for c in blocked)}**: your administrator has restricted this, "
                    "so your preference cannot override it.")
        if PREFS_SAVER:
            PREFS_SAVER(new)
        _POL.prefs = new
        # keep the older PF / TDS / take-home switches in step
        groups = {GROUP_OF[c] for c in cmd.get("cols", []) if c in GROUP_OF}
        if groups:
            try:
                cur = set(hidden_now())
                cur = (cur | groups) if cmd["type"] == "hide" else (cur - groups) if cmd["type"] in ("show", "only") else cur
                st.session_state.hide_cols = sorted(cur, key=DEFAULT_HIDE.index)
                if PREF_SAVER:
                    PREF_SAVER(st.session_state.hide_cols)
            except Exception:
                pass
        audit_policy_event("preference_changed", json.dumps(cmd)[:200])
        return _policy_reply(msg, "pref", note="Saved for your login only · see Settings → My data preferences.")
    return None


# ======================= v14 TEACHING (formulas, terms, response rules) + "/" SQL mode =======================
TEACH_STORE = None           # set by the app: tg.TeachingStore(APP_DIR)
_EMBED_STATE = {"ok": True, "retry": 0.0}


@functools.lru_cache(maxsize=4000)
def _embed_cached(text):
    r = ollama.embeddings(model=EMBED_MODEL, prompt=text, keep_alive="60m")
    v = r.get("embedding") if isinstance(r, dict) else getattr(r, "embedding", None)
    return tuple(v) if v else None


def _embed(text):
    """Embedding for rule matching (Ollama nomic-embed-text). None -> teaching.py falls back to text comparison."""
    if not EMBED_MODEL or (not _EMBED_STATE["ok"] and time.time() < _EMBED_STATE["retry"]):
        return None
    try:
        v = _embed_cached(str(text)[:500])
        _EMBED_STATE["ok"] = True
        return v
    except Exception:
        _EMBED_STATE.update(ok=False, retry=time.time() + 600)       # model missing / Ollama down: retry in 10 min
        return None


def _llm_yes_no(prompt):
    resp = ollama.chat(model=MODEL, messages=[{"role": "user", "content": prompt}], keep_alive="60m",
                       options={"temperature": 0, "num_predict": 3, "num_ctx": 1024, "num_thread": AI_THREADS})
    return resp["message"]["content"]


def _llm_json(prompt, num_predict=500):
    resp = ollama.chat(model=MODEL, messages=[{"role": "user", "content": prompt}], format="json", keep_alive="60m",
                       options={"temperature": 0, "num_predict": num_predict, "num_ctx": 4096,
                                "num_thread": AI_THREADS})
    return resp["message"]["content"]


def _taught():
    """This user's visible taught formulas for the current request."""
    return getattr(_POL, "taught", None) or {}


def taught_in(question):
    """Taught formula keys named in the question (name or alias), in the order written."""
    found = []
    q = " " + str(question).lower() + " "
    for k, info in _taught().items():
        names = [info.get("label", k), k] + list(info.get("aliases") or [])
        for n in sorted({str(x).lower().strip() for x in names if str(x).strip()}, key=len, reverse=True):
            rx = r"(?<![a-z0-9])" + re.escape(n).replace(r"\ ", r"[\s_\-]*").replace("%", r"\s*%") + r"(?![a-z0-9])"
            m = re.search(rx, q)
            if m:
                found.append((m.start(), k))
                break
    return [k for _, k in sorted(found)]


def taught_prompt():
    """'Taught rules' block for the AI prompts (data only)."""
    if TEACH_STORE is None or getattr(_POL, "uid", None) is None:
        return ""
    try:
        return tg.taught_rules_text(TEACH_STORE, _POL.uid, current_eff())
    except Exception:
        return ""


def taught_definition(question):
    """'formula of utilisation %' -> (title, text) for a taught formula, or None."""
    if not (STRONG_ASK.search(question) or EXPLAIN_ASK.search(question)):
        return None
    keys = taught_in(question)
    if not keys:
        return None
    info = _taught()[keys[0]]
    rec = TEACH_STORE.get(info["id"]) if TEACH_STORE else None
    who = f"taught by {rec.get('created_by')} on {str(rec.get('created_at', ''))[:10]}" if rec else "taught"
    txt = (f"{info.get('understood') or ''}\n\n"
           + (f"- **Per row:** `{info['sql_expr']}`\n" if info.get("sql_expr") else "")
           + f"- **For a group / total:** `{info['agg_expr']}`\n"
           + f"- Teaching **{info['id']}** ({who}). Days in month = days of the **salary month**.")
    return info.get("label", keys[0]) + " (taught formula)", txt


def _ss():
    try:
        return st.session_state
    except Exception:
        return {}


def _teach_reply(text, report="teach", df=None, options=None, note=""):
    return _policy_reply(text, report, df=df, options=options, note=note)


def _preview_formula(prop, mode):
    """Runs the new formula on the latest month (read-only): total + top 5 work orders."""
    lm, ly = latest_period(mode)
    pm = {"month_year": "MYMonth", "payment": "PayMonth"}.get(mode, "SalaryMonth")
    where = f"WHERE [{pm}] = '{ly}-{lm:02d}-01'"
    agg = prop["agg_expr"]
    for q_ in [f"SELECT TOP 1 ({prop['sql_expr']}) AS [v] FROM p {where}"] if prop.get("sql_expr") else []:
        run_ai_query(with_view(q_))                          # the row formula must also run
    tot = run_ai_query(with_view(f"SELECT ROUND({agg}, 2) AS [Total] FROM p {where}"))
    top = run_ai_query(with_view(f"SELECT TOP 5 [WorkOrderNo] AS [Work order], ROUND({agg}, 2) AS [{prop['name']}] "
                                 f"FROM p {where} GROUP BY [WorkOrderNo] ORDER BY 2 DESC"))
    total = tot.iloc[0, 0] if not tot.empty else None
    return period_label(lm, ly), total, top


def _scope_for(q, uid, role, right, users):
    """Where a new teaching applies: 'for everyone' -> global, 'for user1' (admin) -> user:user1, else default."""
    if re.search(r"\b(for\s+)?(every\s*one|everybody|all\s+users|global(ly)?|company[\s\-]*wide)\b", q, re.I):
        return "global"
    if role == "admin":
        who = [u for u in pe.find_users(q, users) if u != uid]
        m = re.search(r"\bfor\s+(?:user\s+)?([a-z0-9._\-]+)\s*$", q, re.I)
        if who and m and m.group(1).lower() in who:
            return f"user:{m.group(1).lower()}"
        return "global"
    return f"self:{uid}"


CORRECTION_RX = re.compile(r"^\s*(?:correction|correct\s+this|that'?s\s+wrong|that\s+is\s+wrong|this\s+is\s+wrong|"
                           r"wrong\s+answer|not\s+correct)\s*[:,\-.!]*\s*(.+)$", re.I | re.S)
T_SHOW = re.compile(r"^\s*(?:show|list|see|view)\s+(?:all\s+)?(my\s+)?teachings?(?:\s+(?:for|of)\s+(?:user\s+)?([a-z0-9._\-]+))?\s*[?.!]*$", re.I)
T_TODAY = re.compile(r"^\s*what\s+(?:did|have)\s+you\s+learn(?:ed|t)?\s+today\s*\??\s*$", re.I)
T_ACT = re.compile(r"^\s*(forget|delete|remove|disable|enable|restore)\s+(?:teaching\s+)?(T\d{1,6})(?:\s+(?:to\s+)?version\s+(\d+))?\s*[.!]*$", re.I)
T_UNDO = re.compile(r"^\s*undo\s+(?:the\s+|my\s+)?last\s+teaching\s*[.!]*$", re.I)
T_TEST = re.compile(r"^\s*test\s+as\s+([a-z0-9._\-]+)\s*:\s*(.+)$", re.I | re.S)
YES_RX = re.compile(r"^\s*(yes|y|ok|okay|save|save\s+it|confirm|go\s+ahead|do\s+it|run|run\s+it|yes,?\s+run\s+it)\s*[.!]*\s*$", re.I)
NO_RX = re.compile(r"^\s*(no|n|cancel|don'?t|do\s+not\s+save|stop|abort|never\s*mind)\s*[.!]*\s*$", re.I)


def _teach_table(recs):
    rows = [{"ID": r["id"], "Type": r["type"], "Scope": r.get("scope"),
             "What": r.get("name") or r.get("understood") or r.get("text_original"),
             "Active": "yes" if r.get("active") else ("deleted" if r.get("deleted") else "disabled"),
             "By": r.get("created_by"), "Created": str(r.get("created_at", ""))[:16], "Version": r.get("version")}
            for r in recs]
    return pd.DataFrame(rows) if rows else None


def teaching_command(question, mode="month_year"):
    """learn / correction / teaching admin commands / confirmations. Returns a reply dict or None."""
    uid, role = getattr(_POL, "uid", None), getattr(_POL, "role", "user")
    if uid is None or TEACH_STORE is None:
        return None
    q = question.strip()
    ss = _ss()
    eff = current_eff()
    right = eff.teach if eff is not None else "none"
    users = USERS_LOADER() if USERS_LOADER else {}
    # ---------- confirm / cancel a pending teaching or a fixed SQL ----------
    pend = ss.get("pending_teach") if hasattr(ss, "get") else None
    if pend and (YES_RX.match(q) or NO_RX.match(q)):
        ss["pending_teach"] = None
        if NO_RX.match(q):
            return _teach_reply("Cancelled - nothing was saved.")
        if pend.get("by") != uid:
            return _teach_reply("🔒 This confirmation belongs to another login.", "blocked")
        if pend["kind"] == "sql":
            return user_sql_answer("/" + pend["sql"], time.time(), fixed=True)
        rec = pend["rec"]
        why = tg.can_teach(right, role, uid, rec["type"], rec.get("scope", "global"))   # re-checked, not trusted
        if why:
            return _teach_reply(why, "blocked")
        try:
            saved = TEACH_STORE.add(rec, uid, reason=pend.get("reason", "chat"))
        except PermissionError as e:
            return _teach_reply(str(e), "blocked")
        set_policy_context(uid, role, users.get(uid))
        what = {"formula": "formula", "term": "rule", "response": "response rule"}[saved["type"]]
        return _teach_reply(f"✅ Saved {what} **{saved['id']}** ({saved.get('scope')}).",
                            note="Admins see every change in Teaching → Changes.")
    # ---------- test as user1: question ----------
    m = T_TEST.match(q)
    if m:
        if role != "admin":
            return _teach_reply("🔒 Only an administrator can test as another user.", "blocked")
        tu, tq = m.group(1).lower(), m.group(2).strip()
        if tu not in users:
            return _teach_reply(f"I couldn't find a user called **{tu}**.")
        lines = []
        for r in TEACH_STORE.response_rules():
            fired, score, why = tg.rule_matches(r, tu, tq, None, parse_periods, _embed, _llm_yes_no)
            if fired:
                lines.append(f"- 🎯 **{r['id']}** fires ({why}) → {r.get('action')}: “{tg.rule_reply(r) or r.get('redirect_question')}”")
        mine = [r for r in TEACH_STORE.for_user(tu)]
        TEACH_STORE.log(uid, "test", "", f"user:{tu}", "", tq[:200])
        txt = (f"**Test as {tu}:** “{tq}”\n\n" + ("\n".join(lines) if lines else "- No response rule fires - "
               "the question would be answered normally.") +
               f"\n\n{len(mine)} formula/term teaching(s) apply to {tu}.")
        return _teach_reply(txt)
    # ---------- list / today ----------
    m = T_SHOW.match(q)
    if m or T_TODAY.match(q):
        if T_TODAY.match(q):
            recs = TEACH_STORE.added_on()
            title = "What I learned today"
        elif m.group(2):
            if role != "admin":
                return _teach_reply("🔒 Only an administrator can see other users' teachings.", "blocked")
            tu = m.group(2).lower()
            recs = [r for r in TEACH_STORE.all() if tu in (r.get("created_by"),) or
                    r.get("scope") in (f"user:{tu}", f"self:{tu}") or tu in (r.get("targets") or [])]
            title = f"Teachings for {tu}"
        elif m.group(1) or role != "admin":
            recs = [r for r in TEACH_STORE.created_by(uid) if r["type"] != "response"]
            title = "Your teachings"
        else:
            recs = TEACH_STORE.all()
            title = "All teachings"
        if role != "admin":
            recs = [r for r in recs if r["type"] != "response" and r.get("created_by") == uid]
        return _teach_reply(f"**{title}:** {len(recs)}" + ("" if recs else " - nothing yet."), "teach",
                            df=_teach_table(recs))
    # ---------- forget / disable / enable / restore / undo ----------
    m = T_ACT.match(q)
    tid, act = None, None
    if m:
        act, tid = m.group(1).lower(), m.group(2).upper()
    elif T_UNDO.match(q):
        act, tid = "delete", TEACH_STORE.last_added(None if role == "admin" else uid)
        if not tid:
            return _teach_reply("There is no teaching to undo.")
    if act:
        rec = TEACH_STORE.get(tid)
        if not rec:
            return _teach_reply(f"There is no teaching **{tid}**.")
        if role != "admin" and (rec.get("created_by") != uid or rec["type"] == "response" or right == "none"):
            return _teach_reply("🔒 You can only change your own teachings.", "blocked")
        try:
            if act in ("forget", "delete", "remove"):
                TEACH_STORE.delete(tid, uid, "chat")
                txt = f"🗑️ **{tid}** is deleted (kept in the history - *restore {tid}* brings it back)."
            elif act == "disable":
                TEACH_STORE.disable(tid, uid, "chat")
                txt = f"⏸️ **{tid}** is disabled."
            elif act == "enable":
                TEACH_STORE.enable(tid, uid, "chat")
                txt = f"▶️ **{tid}** is active again."
            else:
                TEACH_STORE.restore(tid, uid, int(m.group(3)) if m and m.group(3) else None, "chat")
                txt = f"♻️ **{tid}** is restored" + (f" to version {m.group(3)}." if m and m.group(3) else ".")
        except (KeyError, ValueError) as e:
            return _teach_reply(str(e).strip("'\""))
        set_policy_context(uid, role, users.get(uid))
        return _teach_reply(txt)
    # ---------- learn ... / correction ----------
    kind, body = tg.parse_learn(q)
    origin = "learn"
    mc = CORRECTION_RX.match(q) if not kind else None
    if mc:
        body = mc.group(1).strip()
        kind = "formula" if "=" in body and re.search(r"[/*x×+\-]", body) else "term"
        origin = "correction"
    if not kind:
        return None
    if right == "none" and role != "admin":
        audit_policy_event("teach_denied", q[:200])
        return _teach_reply(tg.NOT_ENABLED, "blocked")
    if kind == "response" and role != "admin":
        return _teach_reply("🔒 Only an administrator can create response rules.", "blocked")
    scope = _scope_for(body, uid, role, right, users)
    why = tg.can_teach(right, role, uid, kind, scope)
    if why:
        return _teach_reply(why, "blocked")
    bad = tg.bypass_reason(body)
    if bad:
        audit_policy_event("teach_bypass_refused", body[:200])
        return _teach_reply(bad, "blocked")
    cols = [c for c, _ in _table_columns()]
    if kind == "formula":
        existing = cols + list(tg.P_EXTRA) + [v.get("label", k) for k, v in _taught().items()] + list(_taught())
        try:
            prop, err = tg.parse_formula_reply(_llm_json(tg.formula_prompt(body, cols)), body, cols, existing)
        except Exception as e:
            return _teach_reply(f"⚠️ The AI is not reachable ({e.__class__.__name__}), so I can't read the formula now.")
        if err:
            return _teach_reply(f"I can't save this formula: {err}")
        hidden = [c for c in prop["sources"] if eff is not None and not eff.key_visible(c)]
        if hidden:
            return _teach_reply("🔒 This formula uses data that is not available for your login, so it can't be saved.",
                                "blocked")
        try:
            lab, total, top = _preview_formula(prop, mode)
        except Exception as e:
            return _teach_reply(f"I tried the formula on the latest month and SQL Server refused it "
                                f"({str(e)[:160]}). Please write it differently.")
        prop.update(scope=scope, origin=origin)
        if origin == "correction" and ss.get("last_q"):
            prop["understood"] = (prop.get("understood") or "") + f" (correction after: “{ss.get('last_q')}”)"
        ss["pending_teach"] = {"kind": "formula", "rec": prop, "by": uid, "reason": origin}
        ss["pending_policy"] = None
        unsure = f"\n\n⚠️ **The AI is unsure:** {prop['unsure']}" if prop.get("unsure") else ""
        tot_txt = "-" if total is None or pd.isna(total) else f"{total:,.2f}"
        txt = (f"**I understood this formula** ({'correction' if origin == 'correction' else 'new'}, scope "
               f"**{scope}**):\n\n- **Name:** {prop['name']}  (column `{prop['key']}`)\n"
               f"- **Meaning:** {prop.get('understood') or '-'}\n"
               + (f"- **Per row:** `{prop['sql_expr']}`\n" if prop.get("sql_expr") else "- **Per row:** - (totals only)\n")
               + f"- **For a group / total:** `{prop['agg_expr']}`  (ratio of sums, not an average of row %)\n"
               f"- **Uses:** {', '.join(prop['sources']) or 'only months / counts'}\n"
               f"- **Days in month** = days of the **salary month**.\n\n"
               f"**Test on {lab}:** total = **{tot_txt}** · top 5 work orders below.{unsure}\n\n"
               "**Save this formula? (yes / no)**")
        return _teach_reply(txt, "teach_confirm", df=top, options=["Yes", "No"], note="Nothing is saved until you say yes.")
    if kind == "response":
        ids = [u for u in users if u != uid] or list(users)
        try:
            prop, err = tg.parse_response_reply(_llm_json(tg.response_prompt(body, ids), 700), body, ids)
        except Exception as e:
            return _teach_reply(f"⚠️ The AI is not reachable ({e.__class__.__name__}), so I can't read the rule now.")
        if err:
            return _teach_reply(f"I can't save this rule: {err}")
        prop.update(scope="global", origin=origin)
        rows = []
        tu = prop["targets"][0] if prop["targets"][0] != "all" else uid
        for ex, should in [(x, True) for x in prop["trigger"]["examples"][:5]] + \
                          [(x, False) for x in prop["trigger"]["non_examples"][:3]]:
            fired, score, _ = tg.rule_matches(dict(prop, active=True), tu, ex, None, parse_periods, _embed)
            rows.append({"Example question": ex, "Should fire": "yes" if should else "no",
                         "Fires now": "yes" if fired else "no", "Score": round(score, 2),
                         "OK": "✅" if fired == should else "⚠️"})
        ss["pending_teach"] = {"kind": "response", "rec": prop, "by": uid, "reason": origin}
        ss["pending_policy"] = None
        trig = prop["trigger"]
        months = ", ".join(period_label(m_, y_) for m_, y_ in trig["months"]) or "any month"
        action = {"reply": f"reply “{prop['reply']}”", "block": "say the information is not available",
                  "redirect": f"answer “{prop.get('redirect_question')}” instead"}[prop["action"]]
        unsure = f"\n\n⚠️ **The AI is unsure:** {prop['unsure']}" if prop.get("unsure") else ""
        txt = (f"**I understood this response rule:**\n\n- **For:** {', '.join(prop['targets'])}\n"
               f"- **When they ask:** {trig['meaning']}  (kind: {trig['kind']}, months: {months})\n"
               f"- **Then:** {action} - no data query runs.\n- **Match threshold:** {prop['threshold']:.2f}\n"
               "- The targeted user never sees that a rule exists. Data permissions are not changed.\n\n"
               f"Examples below (⚠️ = my matcher disagrees - rephrase the rule or change the threshold later on the "
               f"Teaching page).{unsure}\n\n**Save this rule? (yes / no)**")
        return _teach_reply(txt, "teach_confirm", df=pd.DataFrame(rows), options=["Yes", "No"],
                            note="Nothing is saved until you say yes.")
    # term / plain rule
    und = body if origin != "correction" or not ss.get("last_q") else \
        f"{body} (correction after: “{ss.get('last_q')}”)"
    rec = {"type": "term", "scope": scope, "text_original": body, "understood": und, "origin": origin}
    ss["pending_teach"] = {"kind": "term", "rec": rec, "by": uid, "reason": origin}
    ss["pending_policy"] = None
    return _teach_reply(f"**I will remember this rule** (scope **{scope}**):\n\n> {und}\n\nIt is given to the AI as a "
                        "definition when it reads questions and writes SQL. It never changes who may see what.\n\n"
                        "**Save this rule? (yes / no)**", "teach_confirm", options=["Yes", "No"],
                        note="Nothing is saved until you say yes.")


def response_rule_hit(question, previous=None, log=True):
    """Checks the admin's response rules for the signed-in user. Returns the rule or None."""
    uid = getattr(_POL, "uid", None)
    if uid is None or TEACH_STORE is None:
        return None
    try:
        rules = TEACH_STORE.response_rules()
    except Exception:
        return None
    for r in rules:
        try:
            fired, score, why = tg.rule_matches(r, uid, question, previous, parse_periods, _embed,
                                                _llm_yes_no if log else None)
        except Exception:
            continue
        if fired:
            if log:
                TEACH_STORE.log(uid, "rule_fired", r["id"], "global", "", question[:300], why)
            return r
    return None


def response_rule_answer(question, default_mode, use_ai, always_ai, depth=0):
    prev = _ss().get("last_q") if hasattr(_ss(), "get") else None
    r = response_rule_hit(question, prev)
    if not r:
        return None
    if r.get("action") == "redirect" and r.get("redirect_question") and depth == 0:
        res = answer(r["redirect_question"], default_mode, use_ai, always_ai, _rule_depth=1)
        return res
    return {"text": tg.rule_reply(r), "note": "", "df": None, "report": "chat", "sql": "", "params": [], "intent": {}}


# ---------------- "/" = the admin's own SQL ----------------
def _policy_table_sql(sql):
    """FROM payrollall / dbo.payrollall / [payrollall] -> the user's policy view (same table, restricted columns NULL)."""
    view = str(T)
    if view == RAW_T:
        return sql
    inner = view[:-len(f" AS {RAW_T}")]

    def rep(m):
        after = sql[m.end():m.end() + 40]
        has_alias = re.match(r"\s+(?:AS\s+)?(?!(?:WHERE|GROUP|ORDER|JOIN|ON|INNER|LEFT|RIGHT|FULL|CROSS|OUTER|UNION|"
                             r"HAVING|WITH)\b)[A-Za-z_]\w*", after, re.I)
        return m.group(1) + (inner if has_alias else view)
    return re.sub(rf"(?i)(\b(?:FROM|JOIN)\s+)(?:\[?dbo\]?\.)?\[?{re.escape(TABLE)}\]?(?![\w\]])", rep, sql)


def _ai_fix_sql(sql, error):
    cols = ", ".join(f"[{c}] {t}" for c, t in _table_columns())
    prompt = ("Fix this Microsoft SQL Server query. Keep the user's meaning exactly; change only what is needed. "
              f"Only table: {TABLE} with columns {cols}. [Month] and [Year] are INT (salary month); [month_year] is a "
              "date (processing month). One SELECT only, no ';', no comments.\n"
              f"Query: {sql}\nError: {error}\n"
              'Reply JSON only: {"sql": "...", "changes": ["what you changed and why"]}')
    data = json.loads(_llm_json(prompt, 500))
    fixed = re.sub(r"^```(?:sql)?|```$", "", str(data.get("sql") or "")).strip().rstrip(";").strip()
    return fixed, [str(x) for x in data.get("changes") or []]


PASTED_SQL = re.compile(r"^\s*(select|with)\b.*\bfrom\s+(\[?dbo\]?\.)?\[?" + re.escape(TABLE) + r"\b", re.I | re.S)


def is_user_sql(text):
    """'/select ...', or a pasted 'select ... from payrollall ...' (without the /)."""
    t = str(text or "").lstrip()
    return t.startswith("/") or bool(PASTED_SQL.match(t))


# "/" SQL mode: these words are refused ANYWHERE in the text (also inside quotes, [brackets] or comments)
USER_SQL_BANNED = re.compile(
    r"\b(DELETE|DROP|TRUNCATE|UPDATE|INSERT|MERGE|ALTER|CREATE|REPLACE|RENAME|EXEC|EXECUTE|GRANT|REVOKE|DENY|"
    r"SHUTDOWN|KILL|BACKUP|RESTORE|DBCC|BULK|INTO|OPENROWSET|OPENQUERY|OPENDATASOURCE|OPENXML|WAITFOR|"
    r"RECONFIGURE|DECLARE|SET|USE|GO|XP_\w*|SP_\w*)\b|;|--|/\*|\*/|@", re.IGNORECASE)


def user_sql_answer(question, t0, fixed=False):
    """'/select ...' - runs the admin's own query exactly as written (read-only, policy view, row limit)."""
    uid, role = getattr(_POL, "uid", None), getattr(_POL, "role", "user")
    sql = question[1:].strip().rstrip(";").strip()
    if role != "admin" and not getattr(_POL, "sql_ok", False):          # master or admin
        audit_policy_event("user_sql_denied", sql[:200])
        return _policy_reply("🔒 SQL mode ( / ) is only for administrators. Please ask your question in words.", "blocked")
    if not sql:
        return _policy_reply(f"Type your SELECT after the **/**, e.g. `/select Category, count(*) from {TABLE} group by Category`.",
                             "chat")
    bad = USER_SQL_BANNED.search(sql)                  # no delete / drop / update... type query, ever
    if bad:
        audit_policy_event("user_sql_blocked", sql[:200])
        return _policy_reply(f"🛡️ Blocked: SQL mode only runs **SELECT** queries. “{bad.group(0).strip()}” is not "
                             "allowed (delete, drop, truncate, update, insert, alter, create, exec and similar are "
                             "always refused - even inside quotes or comments). Nothing was sent to SQL Server.\n\n"
                             f"```sql\n{sql}\n```", "blocked")
    try:
        assert_ai_sql(sql)                                 # read-only + payroll table only (checked BEFORE anything)
    except PermissionError as e:
        return _policy_reply(f"🛡️ {e}\n\n```sql\n{sql}\n```", "blocked")
    eff = current_eff()
    if eff is not None and eff.sql_violations(sql):
        return _policy_reply("🔒 The query uses a column that is not available for your login.", "blocked")
    run_sql = _policy_table_sql(sql)
    t1 = time.time()
    try:
        df = prettify(run_ai_query(run_sql))
    except PermissionError as e:
        return _policy_reply(f"🛡️ {e}", "blocked")
    except Exception as e:
        err = re.sub(r"\[[^\]]*ODBC[^\]]*\]|\[SQL Server\]", "", str(e))[:300]
        if fixed or not (ANALYST_MODE and ai_ok()):
            return _policy_reply(f"❌ SQL Server error: {err}\n\n```sql\n{sql}\n```", "chat")
        try:
            new_sql, changes = _ai_fix_sql(sql, err)
            if USER_SQL_BANNED.search(new_sql):            # the AI's fix must obey the same rule
                raise PermissionError("the AI's fix was not a plain SELECT")
            assert_ai_sql(new_sql)
        except Exception as e2:
            return _policy_reply(f"❌ SQL Server error: {err}\n\nThe AI could not fix it ({e2.__class__.__name__}).\n\n"
                                 f"```sql\n{sql}\n```", "chat")
        ss = _ss()
        ss["pending_teach"] = {"kind": "sql", "sql": new_sql, "by": uid}
        ss["pending_policy"] = None
        ch = "\n".join(f"- {c}" for c in changes) or "- (no explanation given)"
        return _policy_reply(f"❌ SQL Server refused your query: *{err}*\n\n**Your query:**\n```sql\n{sql}\n```\n"
                             f"**AI fixed version:**\n```sql\n{new_sql}\n```\n**What changed:**\n{ch}\n\n"
                             "**Run the fixed query? (yes / no)**", "teach_confirm", options=["Yes", "No"],
                             note="Nothing runs until you say yes.")
    secs = time.time() - t1
    n = len(df)
    txt = f"**Your SQL** · {n:,} row(s)" + (f" (first {ROW_LIMIT:,} shown)" if n >= ROW_LIMIT else "") + \
          (" · ✨ AI-fixed version" if fixed else "")
    return {"text": txt, "note": f"Understood as: **your own SQL** (run exactly as written, read-only). ⏱️ SQL: {secs:.1f}s",
            "df": df, "report": "analysis", "sql": run_sql, "params": [],
            "intent": {"title": "Your SQL", "user_sql": True}}


# ======================= small tasks: reminders + "what can you do" =======================
REMINDER_STORE = None        # set by the app: rmd.ReminderStore(APP_DIR)
CAPABILITY_RX = re.compile(
    r"\b(?:what|which)\b.{0,25}\b(?:can|could|do)\s+(?:you|u)\s+(?:do|help|offer)\b|\bwhat\s+(?:you|u)\s+(?:can|could)\s+do\b|"
    r"\bwhat\s+(?:all\s+)?(?:can|could)\s+(?:you|u)\b|\byour\s+(?:features|capabilities|abilities|skills)\b|"
    r"\bhow\s+(?:can|do|should)\s+i\s+use\s+(?:you|u|this|the\s+(?:bot|chatbot))\b|^\s*help\s*[?!.]*\s*$|"
    r"\bwhat\s+(?:are|is)\s+you\s+(?:able|capable)\b|\bkya\s+kar\s+sakt[ae]\b", re.I)


def capabilities_reply():
    admin = bool(getattr(_POL, "sql_ok", False)) or getattr(_POL, "role", "") == "admin"   # master or admin: "/" SQL
    lm = latest_label().lower() if "latest_label" in globals() else "aug-26"
    txt = (
        "Here is what I can **really** do (everything reads your payroll table - I never change it):\n\n"
        "**📊 Payroll answers**\n"
        f"- Totals for a month: *salary for {lm}*\n"
        f"- Department / category / work order-wise: *department wise ctc and margin for {lm}*\n"
        f"- Filters: *openvem and gml other than wox% for {lm}*, *without efirm*\n"
        f"- Highest / lowest: *which workorder has highest margin in {lm}*, *top 10 employees by ctc*\n"
        "- Trends and comparisons: *salary trend*, *compare jul-26 and aug-26 workorder wise*\n"
        "- Projections: *projection of next 2 months workorder wise*\n"
        "- Proposed TDS (FY 2026-27): *proposed tds of 1023*\n"
        "- Salary + TDS to pay for every employee: *salary run for oct-26*\n"
        "- Which months you can see: *from which month to which month is data available*\n"
        "- Length of service: *which employee work more than 1 year*, *current employees with 2+ years service*\n"
        "- Work order history and breaks: *work order history of 5725*, *who took a break between work orders*\n"
        "- Every answer has a table, chart and a CSV download (also all rows)\n\n"
        "**📘 Explain**\n- *formula of margin %*, *what is month_year* · add *explain* or *why* for a written analysis\n\n"
        "**⚙️ Remember your choices**\n- *from now on never show UniqueNo* · categories on/off in Settings\n\n"
        "**⏰ Small tasks - reminders**\n- *remind me tomorrow to pay 114 workorder* · *remind me in 2 hours to check the "
        "invoice* · *remind me every month on 5th to run payroll*\n- *show my reminders* · *done 3* · *snooze 3 1 day*\n"
        "- Reminders appear here in the chatbot when they are due.\n\n")
    if admin:
        txt += ("**🛠️ Admin only**\n- Your own SQL: start with **/** (read-only, coloured while you type)\n"
                "- Permissions in chat: *block GrossSalary for everyone* · *give teaching right to user1*\n"
                "- Teach me: *learn formula: utilisation % = PaidDays / days in month x 100* · *learn: GIL means …*\n"
                "- Refresh data button in the left menu\n\n")
    txt += ("**🚫 What I can't do**\n- Change, add or delete payroll data · send e-mails, WhatsApp or messages · make "
            "payments · see anything outside the payroll table · pop up reminders when the chatbot is closed.")
    return {"text": txt, "note": "", "df": None, "report": "chat", "sql": "", "params": [], "intent": {}}


def _reminder_reply(text, options=None, df=None):
    return {"text": text, "note": "", "df": df, "report": "reminder", "sql": "", "params": [], "intent": {},
            "options": options or []}


WHEN_OPTIONS = ["Tomorrow 10 AM", "In 1 hour", "Next Monday 10 AM", "Cancel"]


def _save_reminder(uid, text, due, repeat):
    r = REMINDER_STORE.add(uid, text, due, repeat)
    rep_txt = {"daily": " (every day)", "weekly": " (every week)", "monthly": " (every month)"}.get(repeat, "")
    return _reminder_reply(f"⏰ Done - I'll remind you **{rmd.when_text(due)}**{rep_txt}:\n\n> {r['text']}\n\n"
                           f"Reminder **#{r['id']}**. It will show at the top of the chat when it is due "
                           "(when the chatbot is open). *show my reminders* lists them.")


def reminder_command(question):
    """Reminders: create / list / done / delete / snooze. Returns a reply dict or None."""
    uid = getattr(_POL, "uid", None)
    if uid is None or REMINDER_STORE is None:
        return None
    q = question.strip()
    ss = _ss()
    pend = ss.get("pending_reminder") if hasattr(ss, "get") else None
    if pend:
        if re.match(r"^\s*(cancel|no|stop|never\s*mind)\s*[.!]*\s*$", q, re.I):
            ss["pending_reminder"] = None
            return _reminder_reply("OK - no reminder was set.")
        r = rmd.parse_reminder("remind me " + q)
        if r and r["due"]:
            ss["pending_reminder"] = None
            return _save_reminder(uid, pend["text"], r["due"], r["repeat"])
        ss["pending_reminder"] = None                       # something else was asked -> carry on normally
    if rmd.LIST_RX.match(q):
        rows = [{"#": r["id"], "Reminder": r["text"], "When": rmd.when_text(dt_parse(r["due"])),
                 "Repeats": r.get("repeat") or "-",
                 "Status": "⏰ due now" if dt_parse(r["due"]) <= datetime.datetime.now() else "upcoming"}
                for r in REMINDER_STORE.mine(uid)]
        if not rows:
            return _reminder_reply("You have no pending reminders. Try *remind me tomorrow to pay 114 workorder*.")
        return _reminder_reply(f"**Your reminders:** {len(rows)}  ·  say *done 3*, *snooze 3 1 day* or "
                               "*delete reminder 3*", df=pd.DataFrame(rows))
    m = rmd.DONE_RX.match(q)
    if m:
        rid = m.group(1) or m.group(2)
        r = REMINDER_STORE.done(uid, rid)
        if not r:
            return _reminder_reply(f"I couldn't find your reminder #{rid}.")
        return _reminder_reply(f"✅ Reminder #{rid} done." + (f" Next time: **{rmd.when_text(dt_parse(r['due']))}**."
                                                             if r.get("repeat") else ""))
    m = rmd.DEL_RX.match(q)
    if m:
        okd = REMINDER_STORE.delete(uid, m.group(1))
        return _reminder_reply(f"🗑️ Reminder #{m.group(1)} deleted." if okd else f"I couldn't find your reminder #{m.group(1)}.")
    m = rmd.SNOOZE_RX.match(q)
    if m:
        r = REMINDER_STORE.snooze(uid, m.group(1), rmd.snooze_delta(m.group(2)))
        return _reminder_reply(f"😴 Reminder #{m.group(1)} moved to **{rmd.when_text(dt_parse(r['due']))}**." if r
                               else f"I couldn't find your reminder #{m.group(1)}.")
    r = rmd.parse_reminder(q)
    if not r:
        return None
    if r["due"] is None:
        ss["pending_reminder"] = {"text": r["text"]}
        return _reminder_reply(f"Sure - when should I remind you to **{r['text']}**? (e.g. *tomorrow 10 am*, "
                               "*in 2 hours*, *on 5-oct-26 at 3 pm*)", options=WHEN_OPTIONS)
    return _save_reminder(uid, r["text"], r["due"], r["repeat"])


def dt_parse(s_):
    return datetime.datetime.strptime(s_, "%Y-%m-%d %H:%M")


ADD_COL_RX = re.compile(
    r"^\s*(?:(?:please|pls|plz|can\s+you|could\s+you|kindly|also|i\s+mean|no|now|and)[\s,]+)*"
    r"(?:add|include|show|put|insert|give)\s+(?:me\s+)?(?:a\s+|an\s+|the\s+|one\s+more\s+|another\s+|extra\s+|new\s+)?"
    r"(.{2,80}?)\s+(?:column|col|field)s?\b", re.I)


def add_column_followup(question, default_mode, use_ai, always_ai, depth=0):
    """'add last paid salary month_year column' after an answer = show the previous answer again WITH that column.
    Only the answer on screen changes - nothing is written to SQL Server."""
    m = ADD_COL_RX.search(question)
    if not m or depth:
        return None
    if parse_periods(question) or len(m.group(1).split()) > 6:
        return None                                    # a full new question ("... in sep-26, also give me X column")
    ss = _ss()
    prev = ss.get("last_q") if hasattr(ss, "get") else None
    want = re.sub(r"(?i)\b(in|to|into|on)\s+(your|the|this|that|previous|last|above)\s+(provided\s+)?"
                  r"(data|answer|table|report|result)\b.*$", "", m.group(1)).strip()
    if not prev:
        return _policy_reply(f"Which answer should I add **{want}** to? Ask the question again and include it, "
                             "e.g. *projection of sep-26 with last paid month_year*. (Nothing in SQL Server is ever "
                             "changed - columns are only added to the answer you see.)", "chat")
    new_q = f"{prev} with {want} column"
    res = answer(new_q, default_mode, use_ai, always_ai, _rule_depth=1)
    res["_last_q"] = new_q                              # follow-ups continue from the question WITH the column
    cols = " ".join(map(str, res.get("df").columns)).lower() if res.get("df") is not None else ""
    words = [w for w in re.findall(r"[a-z_]+", want.lower()) if len(w) > 2 and w not in
             {"the", "and", "last", "paid", "salary", "column", "with", "data"}]
    found = bool(words) and all(w.replace("_", "") in cols.replace("_", "").replace(" ", "") or w in cols for w in words)
    extra = (f"✅ Same answer as before, with **{want}**." if found else
             f"⚠️ I re-ran your previous question, but this report has no column for **{want}**. "
             "Ask it in your question, or (admin) use your own SQL with **/**.")
    res["note"] = (extra + " Display only - SQL Server is not changed. " + (res.get("note") or "")).strip()
    return res


from ui_state import valid_until, write_valid_until, expired_answer   # expired -> no answers until the master renews


def answer(question, default_mode, use_ai, always_ai=False, _rule_depth=0):
    gone = expired_answer()
    if gone:
        return gone
    t0 = time.time()
    notes = []
    if question.strip().startswith("/"):                # v14: "/select ..." = the admin's own SQL, run as written
        return user_sql_answer(question.strip(), t0)
    if (getattr(_POL, "sql_ok", False) or getattr(_POL, "role", "") == "admin") and PASTED_SQL.match(question.strip()):
        return user_sql_answer("/" + question.strip(), t0)
    question = re.sub(r"\bbasic\s+(sale|sales)\b", r"base \1", question, flags=re.IGNORECASE)  # "basic sales" = base sales
    if len(question.split()) > 2:                       # "<did-you-mean question> yes" (reply to "did you mean ...?")
        question = re.sub(r"\s+(?:yes|yeah|yep|ok|okay|haan|han|correct|right|sahi|ji)\s*[.!]*$", "", question, flags=re.I)
    try:
        _latest = latest_period(month_mode_of(question, default_mode))
    except Exception:
        _latest = None
    question = expand_quarters(question, latest=_latest)   # "q2 of this fy" / "this quarter" / "h1" -> month range
    question = fix_typos(question)                      # "employe list", "salry", "catagory wise"
    notes += [f"✏️ Read “{a_}” as “{b_}”." for a_, b_ in getattr(_POL, "typo_fixes", [])]
    # "employees without pf" = people with zero PF (not "hide the PF column")
    question = re.sub(r"\b(employees?|staff|people|emps?)\s+(?:who\s+(?:are|have|has)\s+)?(?:without|with\s+no|having\s+no)\s+"
                      r"(pf|tds|esic|pt|nps|bonus|arrears?)\b", r"\1 with zero \2", question, flags=re.IGNORECASE)
    # "how is margin compared to last month" = latest month vs the month before
    question = re.sub(r"\b(?:compared?|vs\.?|versus|against)\s+(?:to\s+|with\s+)?(?:the\s+)?(?:last|previous|prev)\s+month\b",
                      " month-on-month change", question, flags=re.IGNORECASE)
    question = question.strip().strip("*").strip()
    dfc = data_from_command(question)                   # "from now on only show data from apr-26" / "which months"
    if dfc:
        return dfc
    cpc = category_pref_command(question)               # "from now on do not show gil data" (a category, not a column)
    if cpc:
        return cpc
    pol = policy_command(question)                      # admin / preference commands come first
    if pol:
        return pol
    tc = teaching_command(question, default_mode)       # v14: learn / correction / teaching commands / yes-no
    if tc:
        return tc
    rc = reminder_command(question)                     # v14: small tasks - reminders
    if rc:
        return rc
    if CAPABILITY_RX.search(question) and not parse_periods(question):   # "what can you do" - honest, fixed list
        return capabilities_reply()
    if _rule_depth == 0:                                # v14: admin response rules - BEFORE any SQL
        rr = response_rule_answer(question, default_mode, use_ai, always_ai)
        if rr:
            return rr
    question = pe.strip_one_time(question, VOCAB)     # "without CTC" -> applied to this answer by the result filter
    eff = current_eff()
    if eff is not None and not eff.payroll_access:
        return _policy_reply("🔒 Your login does not have access to payroll data. Please contact your administrator.",
                             "blocked")
    ac = add_column_followup(question, default_mode, use_ai, always_ai, _rule_depth)
    if ac:
        return ac
    if WRITE_REQUEST.search(question):
        return {"text": "🛡️ I can only **read** payroll data. I can't add, change or delete anything "
                        "in SQL Server. Please ask a question like *salary for aug-26*.",
                "note": "Read-only guard", "df": None, "report": "blocked", "sql": "", "params": [],
                "intent": {}}
    if AI_EXPLAIN_ASK.search(question) and parse_periods(question) and not why_answer_wanted(question):
        # "explain/why ... aug-26": get the data;
        question = re.sub(r"\s+", " ", AI_EXPLAIN_ASK.sub(" ", question)).strip()   # the AI explains it after
    pref = preference_command(question)                 # "from now on don't show pf, tds..."
    if pref:
        return pref
    if wants_salary_run(question):                      # "salary run for oct-26" - salary + TDS of everyone
        return salary_run_answer(question, notes, t0)
    if wants_wo_history(question):                      # "in which workorder 5725 worked", "who took a break"
        return wo_history_answer(question, notes, t0)
    if wants_tenure(question):                          # "which employee work more than 1 year"
        return tenure_answer(question, notes, t0)
    if wants_proposed_tds(question):                    # "proposed tds of 1023", "monthly tds for amit"
        return proposed_tds_answer(question, notes, t0)
    tdef = taught_definition(question)                  # v14: "formula of utilisation %" (taught formula)
    if tdef and not parse_periods(question):
        return {"text": f"📘 **{tdef[0]}**\n\n{tdef[1]}", "note": "Understood as: **Definition / formula** "
                        "(no data query).", "df": None, "report": "definition", "sql": "", "params": [], "intent": {}}
    ex = wants_explanation(question)                    # "formula of margin", "what is month_year"
    if ex and not parse_periods(question):
        return {"text": f"📘 **{ex[0]}**\n\n{ex[1]}", "note": "Understood as: **Definition / formula** "
                        "(no data query).", "df": None, "report": "definition", "sql": "", "params": [],
                "intent": {}}
    signal = has_payroll_signal(question) or wants_projection(question)
    if not signal and SMALLTALK.search(question):      # "great", "lets go", "thanks" -> no report
        return chat_reply(question)
    if not signal and CHAT_STYLE and (use_ai or always_ai) and len(question.split()) > 3:
        return chat_ai(question)                         # general question -> answer like ChatGPT

    mode = month_mode_of(question, default_mode)
    wo_grp = wo_group_scope(question)                    # "NORTH margin for sep-26", "SOUTH vs NORTH CTC"
    if wo_grp:
        question = wo_grp["question"]
    question, asked_out = split_exclusions(question)     # "without efirm" -> exclude EFIRM
    try:
        off = {"Category": list(st.session_state.get("cat_off", []))}   # category switches in Settings
    except Exception:
        off = {"Category": []}
    inc_cat = find_known(question, "Category")          # asked for a category by name -> show it anyway
    question, asked_in = split_includes(question)       # "openvem, gml" -> both
    inc_all = set(asked_in.get("Category", [])) | ({inc_cat} if inc_cat else set())
    excl = {"Category": sorted(set(asked_out["Category"]) | {c for c in off["Category"] if c not in inc_all}),
            "Department": sorted(set(asked_out["Department"])),
            "WorkOrderNo": sorted(set(asked_out.get("WorkOrderNo", [])))}
    for col, vals in asked_in.items():
        excl[f"include_{col}"] = vals
    if wo_grp and wo_grp["names"]:
        excl["wo_group"] = wo_grp["names"]
        notes.append("✅ Only work-order group " + ", ".join(
            f"**{g_['name']}** (work orders starting with {' / '.join(g_['prefixes'])})"
            for g_ in wo_grp["groups"] if g_["name"] in wo_grp["names"]) + ".")
    inc_text = ", ".join(sum(asked_in.values(), []))
    if inc_text:
        notes.append(f"✅ Only **{inc_text}**.")
    try:
        st.session_state["_excl"] = excl
    except Exception:
        pass
    ex_text = ", ".join(excl["Category"] + excl["Department"] + excl["WorkOrderNo"])
    if ex_text:
        why = []
        if asked_out["Category"] or asked_out["Department"] or asked_out.get("WorkOrderNo"):
            why.append("your question")
        if [c for c in off["Category"] if c not in inc_all]:
            why.append("Settings → categories")
        notes.append(f"🚫 Excluding **{ex_text}** ({' + '.join(why)}).")
    rel = relative_months(question, mode)               # "this month" / "last month" / "this year"
    if rel:
        question, said = rel
        notes.append(f"Read “{said}” as the latest months in the data.")
    if wo_grp and wo_grp["wise"]:                        # one row per work-order group
        return wo_group_answer(question, wo_grp, mode, notes, t0)

    # 0) common analyst / HOD / CFO questions with an exact answer: work orders by end date, share of total, statistics
    for direct in (proj_vs_actual_answer, proj_compare_answer, liable_breakup_answer, invoice_no_answer, in_not_in_answer, paid_days_answer, why_answer,
                   whatif_answer, ratio_answer,
                   change_answer, count_answer, emp_value_answer,
                   expired_wo_answer, wo_end_answer, new_wo_answer, share_answer, stats_answer, avg_wo_margin_answer):
        res = direct(question, mode, notes, t0)
        if res:
            return res
    # 1) employee ranking - fixed report, no AI needed
    if wants_projection(question):                      # "projection / future / next month" -> fixed rule
        return projection_answer(question, mode, notes, t0, use_ai or always_ai)
    if wants_plan(question) or taught_in(question):     # several parts / a taught formula -> one plan
        return plan_answer(question, mode, notes, t0, use_ai or always_ai)
    rk = wants_emp_ranking(question)
    if rk:
        return employee_ranking(question, mode, rk[0], rk[1], notes, t0)
    # 2) fully understood by the fixed rules -> no AI call (same answer, much faster)
    intent = fast_intent(question)
    ai_used = False
    force_trend = is_monthly_totals(question)            # "each month / month wise" totals -> fixed report
    if force_trend:
        intent = {**(intent or fast_intent_fallback(question)), "report": "trend", "confident": True}
    wo_list = not force_trend and wants_wo_list(question)
    if wo_list:
        intent = {**(intent or fast_intent_fallback(question)), "report": "workorder_monthly", "confident": True}
    if intent is None and not force_trend and _rule_depth == 0:
        dym = did_you_mean(question)                     # "revnue for sep-26": ask instead of guessing / slow AI
        if dym:
            q2, fixed = dym
            res = clarify(question, "Did you mean **" + q2 + "**? (" + ", ".join(f"“{a_}” → “{b_}”" for a_, b_ in fixed)
                          + ") Click it, or reply *yes*.", [q2] + [o for o in suggest_options(q2, n=3) if o != q2][:2])
            res["clarify_for"] = q2                      # a "yes" reply is added to this -> answers the corrected one
            return res
    # 3) clearly an analysis question -> skip the classify step, go straight to AI SQL
    direct_analysis = (intent is None and not force_trend and ANALYST_MODE and (use_ai or always_ai)
                       and bool(set(re.findall(r"[a-z]+", question.lower())) & ANALYSIS_WORDS)
                       and not wants_wo_summary(question))
    if direct_analysis:
        tried = list(notes)
        res = analyst_answer(question, mode, tried, t0)
        if res:
            return res
        why = tried[-1] if tried else ""
        return clarify(question, (f"{why} — so I don't have a reliable answer yet. " if why else
                                  "I don't have a reliable answer for this yet. ") +
                       "Could you rephrase it (which figure, which month, what to compare)? Or pick one of these:",
                       suggest_options(question))
    if intent is None:
        if use_ai or always_ai:
            try:
                intent = ai_intent(question)
                ai_used = True
            except Exception as e:
                if "timeout" in f"{e.__class__.__name__} {e}".lower() or "timed out" in str(e).lower():
                    notes.append(f"⚠️ The AI took longer than {AI_TIMEOUT_SEC}s - answered with the keyword rules.")
                    use_ai = always_ai = False           # don't make the user wait for a second slow AI call
                else:
                    notes.append(f"⚠️ AI not reachable ({e.__class__.__name__}) - used keyword rules. Is Ollama running?")
        if intent is None:
            intent = fast_intent(question) or fast_intent_fallback(question)
    t_ai = time.time() - t0

    if not valid_wo(intent.get("workorder")):
        intent["workorder"] = None
    no_filters = not any(intent.get(k) for k in ("department", "category", "employee", "workorder"))
    if not signal and no_filters:                      # nothing payroll-related -> just chat
        return chat_ai(question) if CHAT_STYLE and (use_ai or always_ai) else chat_reply(question)
    if intent.get("report") == "chitchat":             # AI said chat, but question has payroll words
        intent = fast_intent(question) or fast_intent_fallback(question)
    kw = fast_intent(question) or fast_intent_fallback(question)
    keyword_sure = fast_intent(question) is not None or force_trend or wo_list   # understood without AI
    if keyword_sure and not force_trend and not wo_list:   # fully understood by fixed rules -> no AI guessing
        intent = {**kw, "confident": True}
    # a report that needs a work order / employee can't run without one
    if intent.get("report") == "workorder_detail" and not intent.get("workorder"):
        intent["report"] = kw.get("report", "summary") if kw.get("report") != "workorder_detail" else "summary"
    if intent.get("report") == "employee_detail" and not intent.get("employee"):
        intent["report"] = kw.get("report", "summary") if kw.get("report") != "employee_detail" else "summary"
    if not intent.get("workorder") and kw.get("workorder"):          # regex found a work order no.
        intent["workorder"] = kw["workorder"]
    if intent.get("report", "summary") == "summary" and kw.get("report", "summary") != "summary":
        intent["report"] = kw["report"]                                # keywords are more specific
    for k in ("department", "category"):
        intent[k] = intent.get(k) or kw.get(k)
    no_filters = not any(intent.get(k) for k in ("department", "category", "employee", "workorder"))
    # --- anything the fixed reports can't do -> AI analyst (read-only SQL) ---
    unknown = unknown_words(question, intent)
    if (use_ai or always_ai) and not force_trend and wants_analysis(question, intent, unknown, keyword_sure):
        tried = list(notes)
        res = analyst_answer(question, mode, tried, t0)
        if res:
            return res
        # don't force a wrong report - say so and ask
        why = tried[-1] if tried else ""
        return clarify(question, (f"{why} — so I don't have a reliable answer yet. " if why else
                                  "I don't have a reliable answer for this yet. ") +
                       "Could you rephrase it (which figure, which month, what to compare)? Or pick one of these:",
                       suggest_options(question, kw.get("report")))
    # --- ask back instead of guessing ---
    if unknown and not keyword_sure:
        # words we could not match: look for them in department / category names (real data, not AI)
        hits = []
        for col in ("Category", "Department"):
            for w in sorted(unknown, key=len, reverse=True)[:3]:
                for v in similar_values(col, w, limit=4):
                    if v not in [h[1] for h in hits]:
                        hits.append((col, v))
        if hits:
            p = period_phrase(question)
            return clarify(question, f"I couldn't match **{', '.join(sorted(unknown))}** exactly. "
                                     "Did you mean one of these?",
                           [f"ctc and base sales for {v}" + (f" for {p}" if p else "") for _, v in hits[:6]])
    if ai_used and intent.get("confident") is False and not keyword_sure:
        # don't show the AI's own wording (it can invent things) - say plainly what was not understood
        what = f"I didn't understand **{', '.join(sorted(unknown))}**. " if unknown else \
            "I'm not sure which report you want. "
        return clarify(question, what + "Did you mean one of these?", suggest_options(question, intent.get("report")))
    if (intent.get("report", "summary") == "summary" and no_filters and unknown
            and not wants_wo_summary(question) and not keyword_sure):
        return clarify(question, f"I didn't understand **{', '.join(sorted(unknown))}**. "
                                 "Which of these did you mean?")
    intent["pct_basis"] = pct_basis(question)
    intent["excluded"] = ex_text
    intent["included"] = inc_text
    for k, col in (("category", "Category"), ("department", "Department")):
        if intent.get(k) and intent[k] in excl[col]:
            intent[k] = None                              # never "only X" and "without X" at once
    intent["focus"], intent["focus_label"] = focus_metric(question)
    intent["sort"] = intent.get("sort") or ranking(question)[0]
    intent["_q"] = question
    intent["ask_rank"] = bool(MONTH_RANK.search(question) or re.search(r"\b(highest|lowest|best|worst|max\w*|min\w*|most|least)\b",
                                                                       question, re.I))
    report = intent.get("report", "summary")
    if force_trend:
        report = "trend"
    elif wo_list:
        report = "workorder_monthly"
    elif wants_wo_summary(question):
        report = "workorder_summary"
        intent["sort"], intent["top"] = ranking(question)
    elif intent.get("workorder"):
        report = "workorder_detail"
    elif intent.get("employee") and report in ("summary", "employee_list"):
        report = "employee_detail"

    periods = parse_periods(question)
    if not periods and report not in ("trend", "employee_detail", "workorder_detail") \
            and not (report == "workorder_summary" and intent.get("workorder")):
        periods = [latest_period(mode)]
        notes.append("No month mentioned, showing the latest month in the data.")

    t1 = time.time()
    sql, params = build_sql(report, mode, periods, intent)
    df = run_query(sql, params)
    wo = intent.get("workorder")
    if df.empty and wo and not is_pattern(wo):   # exact no. not found -> try "contains"
        intent["workorder"] = f"%{wo}%"
        sql, params = build_sql(report, mode, periods, intent)
        df = run_query(sql, params)
        if not df.empty:
            notes.append(f"No exact work order {wo}; showing work orders containing it.")
    t_sql = time.time() - t1

    # --- nothing found: suggest instead of a dead end ---
    if df.empty and report == "employee_detail" and intent.get("employee") \
            and not str(intent["employee"]).isdigit():
        names = similar_values("Name", intent["employee"])
        if names:
            return clarify(question, f"I couldn't find an employee named **{intent['employee']}**. Did you mean:",
                           [f"salary of {n}" for n in names])
    if df.empty and intent.get("workorder"):
        wos = similar_values("WorkOrderNo", str(intent["workorder"]).strip("%"))
        if wos:
            return clarify(question, f"No work order matches **{wo}**. Did you mean one of these?",
                           [f"details of workorder {w}" for w in wos])
    if df.empty and periods and report not in ("joiners", "leavers"):
        lm, ly = latest_period(mode)
        if (lm, ly) not in periods:
            lp = period_label(lm, ly).lower()
            return clarify(question, f"There is no data for **{', '.join(period_label(m, y) for m, y in periods)}**. "
                                     f"The latest month available is **{period_label(lm, ly)}**. Show that instead?",
                           [re.sub(r"\b" + re.escape(period_phrase(question)) + r"\b", lp, question,
                                   flags=re.IGNORECASE) if period_phrase(question) else f"{question} {lp}",
                            f"salary summary for {lp}"])

    df = hide_columns(df, question, report)       # no PF / take-home / PT / ESIC / gross / TDS unless asked
    if "MarginPctSales" in df:                    # keep only the margin % the user asked for
        if intent.get("pct_basis") == "sales":
            df["MarginPct"] = df["MarginPctSales"]
        df = df.drop(columns=["MarginPctSales"])
    fc = intent.get("focus")
    if report == "workorder_summary" and fc == "CTC" and not re.search(r"\bctc\b|cost", question, re.IGNORECASE):
        intent["focus"], intent["focus_label"] = fc, _ = ("Margin", "Margin")   # work orders rank by margin
        fc = "Margin"
    if report in ("department", "category", "workorder_summary") and fc in df and not df.empty:
        asc = intent.get("sort") == "asc"
        keys = ([c for c in ("Year", "Month") if c in df] + [fc])
        df = df.sort_values(keys, ascending=[True] * (len(keys) - 1) + [asc], na_position="last")
        df = df.reset_index(drop=True)
        n_top = ranking(question)[1]
        if report in ("department", "category") and n_top and len(periods) == 1 and len(df) > n_top:
            intent["all_rows"] = {"n": len(df), **{c: df[c].sum() for c in ("CTC", "BaseSalesValue", "Margin") if c in df}}
            df = df.head(n_top)                           # "top 3 departments by margin %": only those 3
    if report == "trend" and not df.empty and {"Year", "Month"} <= set(df.columns):
        df = df.sort_values(["Year", "Month"])
    if report not in ("employee_detail", "workorder_detail") and {"Year", "Month"} <= set(df.columns) and not df.empty:
        df = df.dropna(subset=["Year", "Month"]).reset_index(drop=True)   # rows with a blank month_year have no month
        df.insert(0, "Period", [period_label(m, y) for m, y in zip(df["Month"], df["Year"])])
        df = df.drop(columns=["Year", "Month"])

    if report == "workorder_monthly" and "Period" in df:
        df = df.rename(columns={"Period": "Salary month"})       # Year/Month = the month the salary is for
    intent["period_col"] = PERIOD_COL.get(mode, "Salary month")   # table header for "Period"
    periods_text = ", ".join(period_label(m, y) for m, y in periods) if periods else "all months"
    if periods and report not in ("joiners", "leavers"):
        periods_text += f" ({intent['period_col'].lower()})"      # "Oct-2026 (payment month)"
    notes.insert(0, f"Understood as: **{REPORT_TITLES.get(report, report)}**.")
    if report in ("workorder_summary", "workorder_detail", "workorder_monthly"):
        notes.append(f"Base sales value = **Invoice × 100 / {100 + GST_RATE}** (excl. {GST_RATE}% GST).")
    if report not in ("joiners", "leavers"):
        notes.append(f"Filtered by **{mode_label(mode)}**.")
    notes.append(f"⏱️ {'AI' if ai_used else 'No AI'}: {t_ai:.1f}s · SQL: {t_sql:.1f}s")
    return {
        "text": (f"No records found for work order **{intent['workorder']}**."
                 if df.empty and intent.get("workorder") else make_answer(report, df, periods_text, intent)),
        "note": " ".join(notes),
        "df": df, "report": report, "sql": sql, "params": params, "intent": intent,
    }


def fast_intent_fallback(question):
    """Keyword-only intent when AI is off or unavailable."""
    wo, question = find_workorder(question)
    if wo:
        return {"report": "workorder_detail", "workorder": wo}
    t = question.lower()
    tokens = set(re.findall(r"[a-z_]+", t))
    intent = {"report": "summary", "department": find_known(question, "Department"),
              "category": find_known(question, "Category"), "employee": None}
    for report, words in RULES:
        if tokens & words:
            intent["report"] = report
            break
    return intent


# ------------------------------ Streamlit UI ------------------------------
import base64

st.set_page_config(page_title="Payroll AI Assistant · Workorder Dummy", page_icon="🤖", layout="wide",
                   initial_sidebar_state="expanded")


# ---------- theme (each user chooses; saved with their login) ----------
THEMES = {   # v14 look: deep-navy Workorder Dummy palette (sidebar is navy in both modes)
    "light": dict(bg="#f3f6fb", surface="#ffffff", surface2="#f5f8fd", line="#e0e7f2", text="#0f1b33",
                  muted="#5a6883", heading="#0b2a6f", metric="#0b2a6f", accent="#1f5fd6",
                  user="#ffffff", user_line="#d6e0f3", chip_bg="#fdeee6", chip_fg="#a8431a",
                  hover="#fff7f3", hero1="#ffffff", hero2="#eef3ff", tip="#eef3ff", chart="#1f5fd6",
                  side1="#0b1733", side2="#0b1733", code="#f3f6fb", topbar="#ffffff", glow="rgba(31,95,214,.10)"),
    "dark": dict(bg="#0a1330", surface="#0f1b3b", surface2="#0c1734", line="#22325e", text="#e6ecf8",
                 muted="#97a5c6", heading="#ffffff", metric="#ffffff", accent="#2f6fe4",
                 user="#111f45", user_line="#2a3d72", chip_bg="#3a1f14", chip_fg="#ffb08a",
                 hover="#16244c", hero1="#15275a", hero2="#0f1c40", tip="#142453", chart="#3b82f6",
                 side1="#0b1733", side2="#0b1733", code="#0a142e", topbar="#0b1733", glow="rgba(0,0,0,.35)"),
}
DARK = bool(st.session_state.get("ui_dark", DEFAULT_THEME == "dark"))
TH = THEMES["dark" if DARK else "light"]
CHART_MAIN = TH["chart"]

_VARS = (":root {--navy:#0b2a6f; --orange:#E65F26; --silver:#899AAA;"
         + "".join(f" --{k.replace('_', '-')}:{v};" for k, v in TH.items()) + "}")

CSS = r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap');
__VARS__
html, body, .stApp, [class*="st-"], button, input, textarea {font-family: 'Poppins', 'Segoe UI', sans-serif;}
[data-testid="stIconMaterial"], .material-symbols-rounded, span[class*="material"] {
     font-family: 'Material Symbols Rounded' !important;}
[data-testid="stChatInput"] textarea::placeholder, input::placeholder {color: var(--muted) !important; opacity: 1;}
footer, [data-testid="stAppDeployButton"], [data-testid="stToolbar"], #MainMenu {display: none !important;}
header[data-testid="stHeader"] {background: transparent; height: 0;}
.stApp {background: var(--bg);}
.block-container {padding: 1.2rem 2.2rem 7rem; max-width: 1320px;}
.stApp, .stMarkdown, [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
label, [data-testid="stWidgetLabel"] p, .stRadio label p, .stCheckbox label p {color: var(--text);}
h1, h2, h3, h4, h5 {color: var(--heading) !important; font-weight: 700;}
[data-testid="stCaptionContainer"], .stCaption {color: var(--muted) !important;}
code {background: var(--code) !important; color: var(--text) !important;}

/* ---------------- sidebar ---------------- */
section[data-testid="stSidebar"] {background: linear-gradient(180deg, var(--side1) 0%, var(--side2) 100%);
     border-right: none; min-width: 300px !important; max-width: 300px !important;}
section[data-testid="stSidebar"] * {color: #e8eefc;}
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {padding-top: .4rem;}
section[data-testid="stSidebar"] hr {border-color: rgba(255,255,255,.14); margin: .8rem 0;}
.sb-logo {display: flex; align-items: center; gap: 10px; margin: 6px 0 18px;}
.sb-logo .t1 {font-weight: 800; font-size: 1.28rem; letter-spacing: .02em; line-height: 1;}
.sb-logo .t1 b {color: var(--orange);}
.sb-logo .t2 {font-size: .64rem; letter-spacing: .32em; color: #b9c7e6 !important; margin-top: 3px;}
.sb-bot {display: flex; align-items: center; gap: 12px; margin-bottom: 10px;}
.sb-bot .av {width: 58px; height: 58px; border-radius: 50%; background: radial-gradient(circle at 30% 30%, #2f6fe4, #0b2a6f);
     display: flex; align-items: center; justify-content: center; box-shadow: 0 0 0 3px rgba(255,255,255,.15);}
.sb-bot .n {font-weight: 700; font-size: 1.05rem; line-height: 1.2;}
.sb-bot .n b {color: var(--orange);}
.sb-bot .s {font-size: .78rem; color: #c9d5ee !important;}
.sb-status {display: flex; flex-wrap: wrap; align-items: center; gap: 6px 10px; font-size: .68rem;
     color: #b9c7e6 !important; margin-bottom: 8px;}
.sb-online {background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.18); border-radius: 999px;
     padding: 3px 12px; font-size: .75rem; font-weight: 600;}
.sb-online i {display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; vertical-align: 1px;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button {background: transparent; border: none;
     justify-content: flex-start; padding: 10px 14px; border-radius: 10px; font-size: .95rem; font-weight: 500;
     color: #e8eefc; box-shadow: none; min-height: 44px;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button:hover {background: rgba(255,255,255,.08); color: #fff;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button p {font-size: .95rem; text-align: left;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button > div {justify-content: flex-start !important; width: 100%;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button [data-testid="stMarkdownContainer"] {text-align: left; width: 100%;}
section[data-testid="stSidebar"] label:has(input[role="switch"]) > div:not([data-testid]) {
     background: rgba(255,255,255,.30) !important;}
section[data-testid="stSidebar"] label:has(input[role="switch"]:checked) > div:not([data-testid]) {
     background: var(--orange) !important;}
section[data-testid="stSidebar"] label:has(input[role="switch"]) > div:not([data-testid]) > div {background: #fff !important;}
.sb-user {display: flex; align-items: center; gap: 12px; padding: 10px 6px;}
.sb-user .ic {width: 44px; height: 44px; border-radius: 50%; background: rgba(255,255,255,.12);
     display: flex; align-items: center; justify-content: center; font-size: 1.2rem;}
.sb-user .n {font-weight: 600; font-size: .95rem;}
.sb-user .r {font-size: .74rem; color: #b9c7e6 !important; text-transform: capitalize;}
section[data-testid="stSidebar"] div[class*="st-key-sb_logout"] button,
section[data-testid="stSidebar"] div[class*="st-key-sb_refresh"] button {background: rgba(255,255,255,.06);
     border: 1px solid rgba(255,255,255,.18); border-radius: 10px; font-size: .8rem; min-height: 34px;}
section[data-testid="stSidebar"] [data-testid="stToggle"] p, section[data-testid="stSidebar"] label p {color: #e8eefc !important;}
.sb-tag {margin-top: 26px; font-style: italic; font-size: 1.12rem; line-height: 1.35; color: #dbe5f8 !important;}
.sb-tag b {color: #f4a04a !important; font-weight: 700; font-size: 1.22rem;}
.sb-swoosh {height: 14px; margin-top: 4px; border-top: 3px solid var(--orange); border-radius: 60% 60% 0 0 / 100% 100% 0 0;
     width: 150px; transform: rotate(-8deg);}

/* ---------------- top bar ---------------- */
.topbar {display: flex; justify-content: flex-end; align-items: center; gap: 14px; color: var(--muted);
     font-size: .86rem; margin: 2px 4px 14px;}
.topbar b {letter-spacing: .06em; color: var(--muted); font-weight: 600;}
.topbar i {font-style: italic;} .topbar span.sep {opacity: .5;}

/* ---------------- hero ---------------- */
.hero2 {position: relative; overflow: hidden; border-radius: 22px; padding: 30px 38px; min-height: 238px;
     background: linear-gradient(110deg, var(--hero1) 0%, var(--hero2) 62%, var(--hero1) 100%);
     border: 1px solid var(--line); box-shadow: 0 10px 30px rgba(15,40,100,.07); margin-bottom: 26px;}
.hero2:after {content: ""; position: absolute; right: -60px; bottom: -118px; width: 520px; height: 170px;
     border-radius: 50%; border-top: 6px solid rgba(230,95,38,.75); transform: rotate(-8deg);}
.hero2 .txt {position: relative; z-index: 2; max-width: 64%;}
.hero2 .hi {font-size: 1.6rem; font-weight: 700; color: var(--heading);}
.hero2 .h1 {font-size: 2.35rem; margin: 2px 0 10px; line-height: 1.15; font-weight: 800; color: var(--heading);}
.hero2 .h1 span.ai {color: var(--orange);}
.hero2 p {font-size: 1.02rem; color: var(--text); margin: 0 0 18px; line-height: 1.55;}
.hero2 .art {position: absolute; right: 26px; top: 10px; z-index: 1;}
.bubble {position: absolute; right: 250px; top: 26px; z-index: 3; background: #1f5fd6; color: #fff !important;
     padding: 12px 18px; border-radius: 18px 18px 18px 4px; font-weight: 600; font-size: 1rem; line-height: 1.25;
     transform: rotate(-6deg); box-shadow: 0 8px 18px rgba(31,95,214,.35);}
.hero-compact {border-radius: 16px; padding: 14px 22px; margin-bottom: 18px; display: flex; align-items: center; gap: 14px;
     background: linear-gradient(110deg, var(--hero1), var(--hero2)); border: 1px solid var(--line);}
.hero-compact .t {font-weight: 800; font-size: 1.25rem; color: var(--heading); white-space: nowrap;}
.hero-compact .t span {color: var(--orange);}
.hero-compact .pl {padding: 5px 11px; font-size: .72rem;} .hero-compact .pills {gap: 6px; justify-content: flex-end;}
.pills {display: flex; flex-wrap: wrap; gap: 10px;}
.pl {display: inline-flex; align-items: center; gap: 7px; padding: 8px 15px; border-radius: 999px; font-size: .8rem;
     font-weight: 600; color: #fff !important; box-shadow: 0 4px 10px rgba(0,0,0,.08);}
.pl.g {background: #0f7a4f;} .pl.p {background: #3b3fb8;} .pl.o {background: var(--orange);}
.pl.b {background: #1f5fd6;} .pl.r {background: #c0392b;}
.pl i {width: 8px; height: 8px; border-radius: 50%; background: #7CFFB2; display: inline-block;}
.pl.r i {background: #ffd2cc;}
@media (max-width: 1100px) {.hero2 .art, .bubble {display: none;} .hero2 .txt {max-width: 100%;}}

/* ---------------- welcome cards ---------------- */
.sec-h {display: flex; align-items: center; gap: 12px; margin: 6px 0 2px;}
.sec-h .bulb {width: 38px; height: 38px; border-radius: 50%; background: var(--tip); display: flex;
     align-items: center; justify-content: center; font-size: 1.2rem;}
.sec-h .t {font-size: 1.45rem; font-weight: 700; color: var(--heading);}
.sec-sub {color: var(--muted); margin: 2px 0 16px; font-size: .95rem;}
div[class*="st-key-card_"] {position: relative;}
div[class*="st-key-card_"] div[class*="st-key-qbtn_"] {position: absolute !important; inset: 0; z-index: 5; margin: 0 !important;}
div[class*="st-key-qbtn_"] .stButton, div[class*="st-key-qbtn_"] .stButton > div {width: 100%; height: 100%;}
div[class*="st-key-qbtn_"] button {width: 100% !important; height: 100% !important; opacity: 0; cursor: pointer;}
.qcard {display: flex; align-items: center; gap: 16px; background: var(--surface); border: 1px solid var(--line);
     border-radius: 16px; padding: 18px 18px; min-height: 92px; box-shadow: 0 4px 14px rgba(15,40,100,.05);
     transition: all .15s ease;}
div[class*="st-key-card_"]:hover .qcard {border-color: var(--orange); transform: translateY(-2px);
     box-shadow: 0 10px 22px rgba(15,40,100,.10);}
.qcard .ic {min-width: 54px; height: 54px; border-radius: 50%; display: flex; align-items: center;
     justify-content: center; font-size: 1.45rem;}
.qcard .tt {font-weight: 600; color: var(--heading); font-size: 1rem;}
.qcard .ss {color: var(--muted); font-size: .84rem; margin-top: 3px;}
.qcard .ch {margin-left: auto; color: var(--muted); font-size: 1.3rem;}
.tipbar {background: var(--tip); border-radius: 12px; padding: 12px 18px; margin: 18px 0 8px; color: var(--text);
     font-size: .88rem; border: 1px solid var(--line);}
.tipbar b {color: var(--heading);}

/* ---------------- chat ---------------- */
div[data-testid="stChatMessage"] {background: var(--surface); border: 1px solid var(--line); border-radius: 16px;
     padding: 12px 16px; margin-bottom: 10px; box-shadow: 0 3px 10px rgba(15,40,100,.04);}
div[data-testid="stChatMessage"]:has([aria-label="Chat message from user"]),
div[data-testid="stChatMessage"]:has(div[data-testid="stChatMessageAvatarUser"]) {
     background: var(--user); border-color: var(--user-line); border-left: 4px solid var(--accent);}
[data-testid="stChatInput"] {border-radius: 999px !important; border: 1px solid var(--line) !important;
     background: var(--surface) !important; box-shadow: 0 8px 24px rgba(15,40,100,.10); padding: 4px 8px;}
[data-testid="stChatInput"] textarea {color: var(--text) !important; background: transparent !important; font-size: .98rem;}
[data-testid="stChatInput"] > div {background: transparent !important; border: none !important;}
[data-testid="stChatInputSubmitButton"] {background: #0b2a6f !important; border-radius: 50% !important;
     width: 42px; height: 42px; color: #fff !important;}
[data-testid="stChatInputSubmitButton"] svg {fill: #fff !important; color: #fff !important;}
[data-testid="stBottom"] > div, [data-testid="stBottomBlockContainer"] {background: var(--bg) !important;}
.note {color: var(--muted); font-size: .8rem; margin-top: 6px;}
.timechip {display: inline-block; background: var(--chip-bg); color: var(--chip-fg); border-radius: 999px;
     padding: 2px 10px; font-size: .75rem; margin-right: 6px; font-weight: 600;}

/* ---------------- widgets in both themes ---------------- */
div[data-testid="stMetric"] {background: var(--surface); border: 1px solid var(--line); border-radius: 14px;
     padding: 12px 16px; border-top: 4px solid var(--orange); box-shadow: 0 3px 10px rgba(15,40,100,.05);}
div[data-testid="stMetric"] label, div[data-testid="stMetric"] label p {color: var(--muted) !important; font-weight: 700;
     text-transform: uppercase; font-size: .72rem !important; letter-spacing: .05em;}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] div {color: var(--metric) !important; font-weight: 700;}
.stTabs [data-baseweb="tab"] {color: var(--muted);} .stTabs [aria-selected="true"] {color: var(--heading) !important;}
div[data-baseweb="tab-highlight"] {background-color: var(--orange) !important;}
div[data-baseweb="tab-border"] {background-color: var(--line) !important;}
[data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="select"] > div, [data-baseweb="textarea"],
.stTextInput input, .stTextArea textarea, .stNumberInput input {background: var(--surface) !important;
     color: var(--text) !important; border-color: var(--line) !important;}
[data-baseweb="popover"] ul, [data-baseweb="menu"] {background: var(--surface) !important;}
[data-baseweb="popover"] li {color: var(--text) !important;}
[data-testid="stExpander"] details, [data-testid="stExpander"] {background: var(--surface); border-color: var(--line) !important;
     border-radius: 12px;}
[data-testid="stExpander"] summary p {color: var(--text) !important;}
[data-testid="stForm"] {background: var(--surface); border-color: var(--line) !important; border-radius: 14px;}
.stButton button, .stDownloadButton button, .stFormSubmitButton button {border-radius: 10px; border-color: var(--line);
     background: var(--surface); color: var(--text);}
.stButton button:hover, .stDownloadButton button:hover {border-color: var(--orange); color: var(--orange);}
.stFormSubmitButton button[kind="primaryFormSubmit"], .stButton button[kind="primary"] {background: var(--orange);
     border-color: var(--orange); color: #fff !important;}
div[class*="st-key-opt_"] button {border-radius: 999px; border: 1px solid var(--orange); color: var(--orange);
     background: var(--surface); font-weight: 600;}
div[class*="st-key-opt_"] button:hover {background: var(--orange); color: #fff;}
[data-testid="stAlert"] {border-radius: 12px;}
.pagecard {background: var(--surface); border: 1px solid var(--line); border-radius: 16px; padding: 18px 22px;
     margin-bottom: 14px; box-shadow: 0 3px 10px rgba(15,40,100,.04);}
.pagecard h4 {margin: 0 0 6px 0 !important;}
.pagetitle {font-size: 1.6rem; font-weight: 800; color: var(--heading); margin: 4px 0 4px;}
.pagesub {color: var(--muted); margin-bottom: 16px;}
.footer {text-align: center; color: var(--muted); font-size: .75rem; margin-top: 26px;}

/* =================== v14 LOOK (navy + orange, Workorder Dummy) =================== */
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
html, body, .stApp, [class*="st-"], button, input, textarea {font-family: 'Plus Jakarta Sans', 'Segoe UI', sans-serif;}
.block-container {padding: .6rem 2.2rem 8rem; max-width: 1400px;}
/* sidebar */
section[data-testid="stSidebar"] {background: var(--side1) !important; border-right: 1px solid rgba(255,255,255,.07);}
.lk-brand {display: flex; align-items: center; gap: 14px; padding: 10px 4px 18px; border-bottom: 1px solid rgba(255,255,255,.08); margin-bottom: 16px;}
.lk-brand .mk {width: 50px; height: 50px; border-radius: 14px; background: linear-gradient(145deg,#ff7a3d,#e2531a);
     display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 1.6rem; color: #fff !important;
     box-shadow: 0 6px 16px rgba(230,95,38,.35);}
.lk-brand .t1 {font-weight: 800; font-size: 1.12rem; letter-spacing: .04em; color: #fff !important; line-height: 1.1;}
.lk-brand .t1 sup {color: var(--orange) !important; font-size: .7rem;}
.lk-brand .t2 {font-size: .7rem; letter-spacing: .2em; color: #aab7d6 !important; margin-top: 4px;}
.lk-bot {display: flex; align-items: center; gap: 12px; padding: 0 4px 6px;}
.lk-bot .av {position: relative; width: 46px; height: 46px; border-radius: 50%; background: #1b2b55;
     display: flex; align-items: center; justify-content: center;}
.lk-bot .av i {position: absolute; right: 0; bottom: 1px; width: 11px; height: 11px; border-radius: 50%; border: 2px solid var(--side1);}
.lk-bot .n {font-weight: 800; font-size: 1.02rem; color: #fff !important;}
.lk-bot .n b {color: var(--orange) !important;}
.lk-bot .s {font-size: .78rem; font-weight: 600;}
.lk-sub {font-size: .72rem; color: #8f9cbd !important; padding: 0 4px 14px; border-bottom: 1px solid rgba(255,255,255,.08); margin-bottom: 14px;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button {min-height: 50px; border-radius: 12px; font-size: 1rem;
     color: #dfe6f6 !important; padding: 10px 16px;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button p {font-size: 1rem !important; font-weight: 500;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button:hover {background: rgba(255,255,255,.06) !important;}
section[data-testid="stSidebar"] div[class*="st-key-nav_"] button {position: relative;}
section[data-testid="stSidebar"] div.st-key-nav_Chat button:after {content: "v14"; position: absolute; right: 14px; top: 50%;
     transform: translateY(-50%); font-family: 'JetBrains Mono', monospace; font-size: .7rem; padding: 3px 9px; border-radius: 999px;
     background: rgba(255,255,255,.14); color: #dfe6f6;}
.lk-user {display: flex; align-items: center; gap: 12px; padding: 12px 14px; border-radius: 14px;
     border: 1px solid rgba(255,255,255,.12); background: rgba(255,255,255,.03);}
.lk-user .ic {width: 40px; height: 40px; border-radius: 50%; background: #23345f; display: flex; align-items: center;
     justify-content: center; font-weight: 800; font-size: .85rem; color: #fff !important;}
.lk-user .n {font-weight: 700; font-size: .95rem; color: #fff !important;}
.lk-user .r {font-size: .76rem; color: var(--orange) !important; text-transform: capitalize;}
section[data-testid="stSidebar"] div[class*="st-key-sb_refresh"] {margin-top: 10px;}
section[data-testid="stSidebar"] div[class*="st-key-sb_refresh"] button {background: #13224a !important; border: 1px solid #2a3d72 !important;
     min-height: 44px; font-weight: 600; color: #fff !important;}
section[data-testid="stSidebar"] div[class*="st-key-sb_logout"] button {background: transparent !important; border: 1px solid rgba(255,255,255,.14) !important;
     min-height: 40px;}
.lk-sync {display: flex; justify-content: space-between; font-size: .74rem; color: #8f9cbd !important; margin: 8px 2px 0;}
.lk-sync b {color: #22c55e !important;}
/* top bar */
.lk-top {display: flex; align-items: center; gap: 10px; min-height: 52px;}
.lk-top {flex-wrap: wrap;}
.lk-top .co {font-weight: 800; letter-spacing: .08em; font-size: .9rem; color: var(--heading); white-space: nowrap;}
.lk-top .tg {font-style: italic; color: var(--muted); font-size: .88rem; white-space: nowrap;}
.lk-pills {display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; align-items: center; min-height: 52px;}
.lk-pill {display: inline-flex; align-items: center; gap: 7px; padding: 7px 14px; border-radius: 999px; font-size: .8rem; font-weight: 600;
     border: 1px solid; white-space: nowrap;}
.lk-pill i {width: 7px; height: 7px; border-radius: 50%; display: inline-block;}
.lk-pill.g {color: #16a34a; border-color: rgba(34,197,94,.45); background: rgba(34,197,94,.10);} .lk-pill.g i {background: #22c55e;}
.lk-pill.p {color: #8b5cf6; border-color: rgba(139,92,246,.45); background: rgba(139,92,246,.10);}
.lk-pill.o {color: #ea7a3c; border-color: rgba(230,95,38,.5); background: rgba(230,95,38,.10);}
.lk-pill.b {color: #3b82f6; border-color: rgba(59,130,246,.45); background: rgba(59,130,246,.10);}
.lk-pill.r {color: #ef4444; border-color: rgba(239,68,68,.45); background: rgba(239,68,68,.10);} .lk-pill.r i {background: #ef4444;}
div.st-key-lk_theme {display: flex; justify-content: flex-end; padding-top: 10px;}
div.st-key-lk_theme label p {font-weight: 600; color: var(--text) !important;}
/* hero */
.lk-hero {position: relative; overflow: hidden; border-radius: 22px; padding: 34px 38px; margin: 10px 0 24px;
     background: linear-gradient(120deg, var(--hero1) 0%, var(--hero2) 70%); border: 1px solid var(--line);
     box-shadow: 0 14px 34px var(--glow);}
.lk-hero .txt {position: relative; z-index: 2; max-width: 74%;}
.lk-hello {display: inline-flex; padding: 6px 14px; border-radius: 999px; font-weight: 700; font-size: .9rem;
     color: #ff8a50 !important; background: rgba(230,95,38,.12); border: 1px solid rgba(230,95,38,.35);}
.lk-hero h2 {font-size: 2.35rem !important; font-weight: 800; margin: 14px 0 10px; line-height: 1.12;}
.lk-hero h2 span.ai {color: var(--orange) !important;}
.lk-hero p {font-size: 1.04rem; color: var(--text); opacity: .9; line-height: 1.6; margin: 0 0 18px;}
.lk-chips {display: flex; flex-wrap: wrap; gap: 10px;}
.lk-chip {display: inline-flex; align-items: center; gap: 8px; padding: 7px 14px; border-radius: 8px; font-size: .86rem; font-weight: 600; border: 1px solid;}
.lk-chip i {width: 8px; height: 8px; border-radius: 50%; display: inline-block;}
.lk-chip.g {color: #22c55e; border-color: rgba(34,197,94,.45); background: rgba(34,197,94,.10);} .lk-chip.g i {background: #22c55e;}
.lk-chip.p {color: #a78bfa; border-color: rgba(139,92,246,.45); background: rgba(139,92,246,.12);}
.lk-chip.o {color: #f08a4b; border-color: rgba(230,95,38,.5); background: rgba(230,95,38,.12);}
.lk-chip.b {color: #60a5fa; border-color: rgba(59,130,246,.45); background: rgba(59,130,246,.10);}
.lk-chip.r {color: #ef4444; border-color: rgba(239,68,68,.45); background: rgba(239,68,68,.10);} .lk-chip.r i {background: #ef4444;}
.lk-hero .art {position: absolute; right: 28px; top: 50%; transform: translateY(-50%) scale(.72); transform-origin: right center; z-index: 1;}
.lk-hero .say {position: absolute; right: 168px; top: 26px; z-index: 3; background: #2f6fe4; color: #fff !important; font-size: .76rem;
     font-weight: 700; padding: 8px 12px; border-radius: 10px 10px 10px 2px; line-height: 1.25;}
@media (max-width: 1100px) {.lk-hero .art, .lk-hero .say {display: none;} .lk-hero .txt {max-width: 100%;}}
.lk-sec {display: flex; justify-content: space-between; align-items: center; margin: 4px 0 14px; gap: 10px; flex-wrap: wrap;}
.lk-sec .t {font-weight: 800; font-size: 1.15rem; color: var(--heading);}
.lk-sec .t:before {content: "💡"; margin-right: 8px;}
.lk-sec .s {color: var(--muted); font-size: .88rem;}
.qcard {background: var(--surface); border: 1px solid var(--line); border-radius: 16px; min-height: 82px; padding: 16px 18px;
     box-shadow: none;}
.qcard .ic {min-width: 46px; height: 46px; border-radius: 12px; font-size: 1.15rem;}
.qcard .tt {font-weight: 700; font-size: .98rem;}
/* user message: right-aligned pill */
div[class*="st-key-umsg_"] {display: flex; justify-content: flex-end; margin: 6px 0 14px;}
.lk-u {display: flex; align-items: center; gap: 14px; justify-content: flex-end;}
.lk-u .b {display: flex; align-items: center; gap: 12px; background: var(--user); border: 1px solid var(--user-line);
     border-radius: 18px; padding: 14px 20px; font-size: 1rem; color: var(--text); max-width: 760px;}
.lk-u .b .ic {width: 34px; height: 34px; border-radius: 10px; background: rgba(230,95,38,.15); display: flex; align-items: center;
     justify-content: center; flex: none;}
.lk-u .av {width: 40px; height: 40px; border-radius: 50%; background: #23345f; color: #fff !important; font-weight: 800;
     font-size: .82rem; display: flex; align-items: center; justify-content: center; flex: none; border: 2px solid var(--user-line);}
div[class*="st-key-usql_"] {margin-left: auto; max-width: 82%;}
/* answer card */
div[data-testid="stChatMessage"] {background: var(--surface); border: 1px solid var(--line); border-radius: 20px;
     padding: 20px 24px; box-shadow: 0 10px 28px var(--glow);}
div[data-testid="stChatMessage"] > div:first-child:not([data-testid="stChatMessageContent"]) {display: none;}
.lk-ans {display: flex; gap: 14px; align-items: flex-start;}
.lk-ans .ic {width: 42px; height: 42px; border-radius: 12px; background: rgba(47,111,228,.16); display: flex; align-items: center;
     justify-content: center; flex: none;}
.lk-ans .tx {font-size: 1rem; line-height: 1.6; color: var(--text);}
.lk-ans .tx b {color: var(--heading);}
.lk-live {display: inline-flex; align-items: center; gap: 6px; font-family: 'JetBrains Mono', monospace; font-size: .75rem;
     color: #22c55e !important; border: 1px solid rgba(34,197,94,.45); background: rgba(34,197,94,.08); border-radius: 8px;
     padding: 6px 10px; white-space: nowrap;}
div[class*="st-key-exp_"] button {min-height: 38px; font-weight: 600;}
.lk-kpis {display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin: 18px 0 8px;}
.lk-kpi {background: var(--surface2); border: 1px solid var(--line); border-radius: 16px; padding: 16px 18px;}
.lk-kpi .l {font-size: .74rem; font-weight: 800; letter-spacing: .06em; text-transform: uppercase; color: var(--muted);}
.lk-kpi .v {font-size: 1.85rem; font-weight: 800; color: var(--metric); margin: 10px 0 6px; font-variant-numeric: tabular-nums; line-height: 1.1;}
.lk-kpi .f {font-size: .76rem; color: var(--muted); font-variant-numeric: tabular-nums;}
.lk-kpi.pos .v {color: #22c55e;}
.lk-kpi.hi {border: 1.5px solid var(--orange); background: rgba(230,95,38,.07);}
.lk-kpi.hi .l {color: var(--orange);}
div[data-testid="stChatMessage"] .stTabs [data-baseweb="tab-list"] {gap: 22px;}
div[data-testid="stChatMessage"] .stTabs [data-baseweb="tab"] p {font-size: .95rem; font-weight: 600;}
div[data-testid="stChatMessage"] .stTabs [aria-selected="true"] p {color: var(--orange) !important;}
/* input */
[data-testid="stChatInput"] {border-radius: 16px !important; border: 1px solid var(--line) !important; background: var(--surface) !important;
     box-shadow: 0 10px 30px var(--glow); padding: 6px 8px 6px 40px; position: relative;}
[data-testid="stChatInput"]:before {content: ""; position: absolute; left: 16px; top: 50%; width: 18px; height: 18px; transform: translateY(-50%);
     background: no-repeat center/18px url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2397a5c6' stroke-width='2.2' stroke-linecap='round'><circle cx='11' cy='11' r='7'/><path d='m20 20-3.5-3.5'/></svg>");}
[data-testid="stChatInputSubmitButton"] {background: var(--orange) !important; border-radius: 12px !important; width: 46px; height: 46px;
     box-shadow: 0 6px 16px rgba(230,95,38,.35);}
.lk-foot {display: flex; justify-content: space-between; color: var(--muted); font-size: .75rem; margin-top: 26px; flex-wrap: wrap; gap: 8px;}
.lk-foot .sec {font-family: 'JetBrains Mono', monospace; color: #22c55e;}
.footer {display: none;}
__DARKEXTRA__

/* ---------------- login ---------------- */
.login-card {max-width: 540px; margin: 50px auto 18px; border-radius: 22px; padding: 26px 30px;
     background: linear-gradient(110deg, var(--hero1), var(--hero2)); border: 1px solid var(--line);
     box-shadow: 0 10px 30px rgba(15,40,100,.08); display: flex; gap: 18px; align-items: center;}
.login-card .tt {font-size: 1.6rem; font-weight: 800; color: var(--heading); line-height: 1.1;}
.login-card .tt span {color: var(--orange);}
.login-card .br {font-size: .72rem; letter-spacing: .22em; color: var(--muted); font-weight: 600; margin-bottom: 4px;}
.login-card p {margin: 6px 0 0; color: var(--text); font-size: .92rem;}
</style>
"""
DARK_EXTRA = """
/* dark mode: tables are drawn by the browser canvas -> invert them for a dark look */
[data-testid="stDataFrame"] {filter: invert(0.92) hue-rotate(180deg);}
[data-testid="stCode"] pre, .stCodeBlock pre {background: var(--code) !important;}
[data-testid="stCode"] code span {filter: brightness(1.4);}
"""
st.markdown(CSS.replace("__VARS__", _VARS).replace("__DARKEXTRA__", DARK_EXTRA if DARK else ""),
            unsafe_allow_html=True)


def logo_html(size=44):
    """Company logo: put your own logo.png next to this file to use it; otherwise a simple S mark."""
    for name in ("logo.png", "logo.jpg", "logo.svg"):
        f = Path(__file__).resolve().parent / name
        if f.exists():
            mime = "image/svg+xml" if name.endswith("svg") else "image/png"
            b64 = base64.b64encode(f.read_bytes()).decode()
            return f'<img src="data:{mime};base64,{b64}" style="height:{size}px">'
    return f"""<svg width="{size}" height="{size}" viewBox="0 0 64 64">
  <path d="M46 14 C38 6 18 8 18 20 C18 30 34 30 40 34" fill="none" stroke="#E65F26" stroke-width="9" stroke-linecap="round"/>
  <path d="M40 34 C48 38 48 52 34 54 C26 55 20 52 16 48" fill="none" stroke="#5b8def" stroke-width="9" stroke-linecap="round"/>
</svg>"""


BOT_FACE = """<svg width="40" height="40" viewBox="0 0 64 64">
  <rect x="10" y="14" width="44" height="38" rx="18" fill="#fff"/>
  <rect x="15" y="20" width="34" height="25" rx="12" fill="#0b1f4d"/>
  <path d="M22 34 q4 -5 8 0 M34 34 q4 -5 8 0" stroke="#5fd3ff" stroke-width="3" fill="none" stroke-linecap="round"/>
  <rect x="5" y="26" width="6" height="14" rx="3" fill="#5b8def"/><rect x="53" y="26" width="6" height="14" rx="3" fill="#5b8def"/>
</svg>"""

ROBOT_ART = """<svg width="330" height="228" viewBox="0 0 330 228">
  <defs><linearGradient id="bd" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffffff"/>
  <stop offset="1" stop-color="#dfe9fb"/></linearGradient></defs>
  <g transform="translate(232,18)">
    <rect width="86" height="112" rx="12" fill="#ffffff" stroke="#cfdcf5" stroke-width="2"/>
    <rect x="14" y="16" width="46" height="7" rx="3.5" fill="#c9d9fb"/>
    <rect x="14" y="30" width="58" height="6" rx="3" fill="#e1eafb"/>
    <rect x="14" y="42" width="40" height="6" rx="3" fill="#e1eafb"/>
    <circle cx="56" cy="80" r="20" fill="#eaf1ff" stroke="#5b8def" stroke-width="3"/>
    <text x="56" y="88" text-anchor="middle" font-size="22" font-weight="700" fill="#1f5fd6">₹</text>
  </g>
  <g transform="translate(262,150)"><ellipse cx="22" cy="30" rx="22" ry="7" fill="#9db8ea"/>
    <rect x="0" y="12" width="44" height="18" fill="#b8ccf2"/><ellipse cx="22" cy="12" rx="22" ry="7" fill="#d5e2fa"/></g>
  <line x1="150" y1="26" x2="150" y2="46" stroke="#9db8ea" stroke-width="4"/>
  <circle cx="150" cy="22" r="7" fill="#E65F26"/>
  <rect x="90" y="44" width="120" height="100" rx="46" fill="url(#bd)" stroke="#d3e0f7" stroke-width="2"/>
  <rect x="104" y="60" width="92" height="66" rx="30" fill="#0b1f4d"/>
  <path d="M124 96 q10 -13 20 0 M158 96 q10 -13 20 0" stroke="#5fd3ff" stroke-width="6" fill="none" stroke-linecap="round"/>
  <path d="M140 110 q10 7 20 0" stroke="#5fd3ff" stroke-width="3" fill="none" stroke-linecap="round"/>
  <rect x="78" y="78" width="16" height="36" rx="8" fill="#1f5fd6"/><rect x="206" y="78" width="16" height="36" rx="8" fill="#1f5fd6"/>
  <rect x="112" y="146" width="76" height="66" rx="24" fill="url(#bd)" stroke="#d3e0f7" stroke-width="2"/>
  <path d="M158 164 C152 158 140 160 141 168 C142 175 156 173 157 180 C158 188 145 190 139 185" fill="none"
        stroke="#E65F26" stroke-width="6" stroke-linecap="round"/>
  <path d="M112 160 C92 150 80 132 74 112" stroke="#dfe9fb" stroke-width="16" fill="none" stroke-linecap="round"/>
  <circle cx="72" cy="104" r="13" fill="#1f5fd6"/>
  <path d="M188 162 C200 172 204 186 200 198" stroke="#dfe9fb" stroke-width="16" fill="none" stroke-linecap="round"/>
</svg>"""


def altair_theme(chart):
    """Makes charts match light / dark mode."""
    return (chart.configure(background="transparent")
            .configure_axis(labelColor=TH["muted"], titleColor=TH["muted"], gridColor=TH["line"],
                            domainColor=TH["line"], tickColor=TH["line"])
            .configure_legend(labelColor=TH["text"], titleColor=TH["text"])
            .configure_view(stroke=None))


@st.cache_data(ttl=60, show_spinner=False)
def db_ok():
    try:
        get_conn().close()
        return True
    except Exception:
        return False


@st.cache_data(ttl=60, show_spinner=False)
def ai_ok():
    try:
        ollama.list()
        return True
    except Exception:
        return False


def inr_short(x):
    """Short money for cards: ₹12.35 L, ₹1.24 Cr."""
    x = float(x or 0)
    if abs(x) >= 1e7:
        return f"₹{x / 1e7:,.2f} Cr"
    if abs(x) >= 1e5:
        return f"₹{x / 1e5:,.2f} L"
    return inr(x)


# ================================ LOGIN ================================
import csv
import hashlib
import hmac
import os
import secrets
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
USERS_FILE = APP_DIR / "users.json"        # IDs + scrambled (hashed) passwords, never plain text
AUDIT_FILE = APP_DIR / "audit_log.csv"     # who logged in / what they asked
ID_RULE = re.compile(r"^[a-z0-9._-]{3,30}$")


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), 200_000).hex()


def load_users():
    try:
        return json.loads(USERS_FILE.read_text("utf-8"))
    except FileNotFoundError:
        return {}


@st.cache_resource
def _users_lock():
    import threading
    return threading.RLock()          # shared by all browsers: two people saving at once can't clash


def save_users(users):
    with _users_lock():
        tmp = USERS_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(users, indent=2), "utf-8")
        os.replace(tmp, USERS_FILE)


def update_user(uid, **fields):
    """Saves settings for ONE user only - re-reads the file first, so other users' changes are kept."""
    with _users_lock():
        users = load_users()
        if uid in users:
            users[uid].update(fields)
            save_users(users)


def set_password(rec, pw, must_change):
    rec["salt"] = secrets.token_hex(16)
    rec["hash"] = _hash(pw, rec["salt"])
    rec["must_change"] = must_change


def check_password(rec, pw):
    return hmac.compare_digest(rec.get("hash", ""), _hash(pw, rec.get("salt", "00")))


def pw_problem(pw):
    if len(pw) < 8:
        return "Password must be at least 8 characters."
    if not re.search(r"[A-Za-z]", pw) or not re.search(r"\d", pw):
        return "Password must contain letters and numbers."
    return None


def temp_password():
    return "Wod@" + secrets.token_hex(3) + str(secrets.randbelow(90) + 10)


REFRESH_FILE = APP_DIR / "refresh_status.json"


def _refresh_status():
    try:
        return json.loads(REFRESH_FILE.read_text("utf-8"))
    except Exception:
        return {}


def _save_refresh_status(user, old, new, secs, status):
    try:
        REFRESH_FILE.write_text(json.dumps({
            "at": datetime.datetime.now().strftime("%d-%b-%Y %H:%M"), "by": user, "rows_before": old,
            "rows_after": new, "seconds": round(secs), "status": status}, indent=2), "utf-8")
    except Exception:
        pass


def audit(user, action, detail=""):
    try:
        new = not AUDIT_FILE.exists()
        with AUDIT_FILE.open("a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["time", "user", "action", "detail"])
            w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user, action, detail[:300]])
    except Exception:
        pass


@st.cache_resource(show_spinner=False)
def _warm_up_ai():
    import threading

    def go():
        for m in {MODEL, explain_model()}:
            try:
                ollama.generate(model=m, prompt="hi", keep_alive="60m",
                                options={"num_predict": 1, "num_thread": AI_THREADS})
            except Exception:
                pass
    threading.Thread(target=go, daemon=True).start()
    return True


@st.cache_resource
def _login_fails():
    return {}          # id -> [wrong tries, time of first wrong try]  (shared by all browsers)


def is_locked(uid):
    rec = _login_fails().get(uid)
    if not rec:
        return False
    if time.time() - rec[1] > 15 * 60:
        _login_fails().pop(uid, None)
        return False
    return rec[0] >= MAX_LOGIN_TRIES


def note_fail(uid):
    rec = _login_fails().setdefault(uid, [0, time.time()])
    rec[0] += 1


def login_banner(subtitle):
    st.markdown(f"""<div class="login-card"><div>{logo_html(56)}</div><div>
  <div class="br">WORKORDER DUMMY</div>
  <div class="tt">Payroll <span>AI</span> Assistant</div><p>{subtitle}</p></div>
  <div style="margin-left:auto">{BOT_FACE}</div></div>""", unsafe_allow_html=True)


def setup_screen():
    login_banner("First-time setup: create the administrator account.")
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid.form("setup"):
        uid = st.text_input("Admin ID", value="admin").strip().lower()
        name = st.text_input("Your name")
        pw = st.text_input("Password", type="password")
        pw2 = st.text_input("Confirm password", type="password")
        if st.form_submit_button("Create admin account", use_container_width=True, type="primary"):
            err = (None if ID_RULE.match(uid) else "ID: 3-30 characters, letters/numbers/._- only.") \
                or pw_problem(pw) or (None if pw == pw2 else "Passwords do not match.")
            if err:
                st.error(err)
            else:
                rec = {"name": name or uid, "role": "admin", "active": True,
                       "created": datetime.date.today().isoformat()}
                set_password(rec, pw, must_change=False)
                save_users({uid: rec})
                audit(uid, "setup", "admin account created")
                st.session_state.user, st.session_state.last_seen = uid, time.time()
                st.rerun()


def login_screen(message=None):
    login_banner("Please sign in with the ID and password given to you.")
    _, mid, _ = st.columns([1, 1.3, 1])
    if message:
        mid.info(message)
    with mid.form("login"):
        uid = st.text_input("User ID").strip().lower()
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("🔐 Sign in", use_container_width=True, type="primary"):
            users = load_users()
            rec = users.get(uid)
            if is_locked(uid):
                st.error(f"Too many wrong tries. ID locked for 15 minutes.")
                audit(uid, "login_locked")
            elif not rec or not check_password(rec, pw):
                note_fail(uid)
                audit(uid or "(blank)", "login_failed")
                st.error("Wrong ID or password.")
            elif not rec.get("active", True):
                audit(uid, "login_disabled")
                st.error("This ID is disabled. Please contact the administrator.")
            else:
                _login_fails().pop(uid, None)
                load_prefs(rec)                                          # this user's own settings
                st.session_state._prefs_for = uid
                st.session_state.page = "Chat"
                st.session_state.user, st.session_state.last_seen = uid, time.time()
                st.session_state.history = []
                audit(uid, "login")
                st.rerun()
    mid.caption("🔒 Internal use only · all access is logged")


def change_password_form(uid, forced=False):
    users = load_users()
    with st.form(f"chpw_{'f' if forced else 's'}"):
        cur = st.text_input("Current password", type="password")
        new = st.text_input("New password", type="password")
        new2 = st.text_input("Confirm new password", type="password")
        if st.form_submit_button("Change password", use_container_width=True, type="primary"):
            rec = users.get(uid)
            err = (None if rec and check_password(rec, cur) else "Current password is wrong.") \
                or pw_problem(new) or (None if new == new2 else "New passwords do not match.") \
                or (None if new != cur else "New password must be different.")
            if err:
                st.error(err)
            else:
                set_password(rec, new, must_change=False)
                update_user(uid, salt=rec["salt"], hash=rec["hash"], must_change=False)
                audit(uid, "password_changed")
                st.success("Password changed.")
                if forced:
                    st.rerun()


USER_PREFS = {   # session key: (saved name in users.json, default)
    "cfg_mode": ("mode", "month_year"), "cfg_always_ai": ("always_ai", ALWAYS_AI), "cfg_use_ai": ("use_ai", True),
    "cfg_explain": ("explain", CHAT_STYLE), "cfg_ai_explain": ("ai_explain", AI_EXPLAIN),
    "hide_cols": ("hide_cols", DEFAULT_HIDE), "cat_off": ("cat_off", []),
}


def load_prefs(rec):
    """Puts one user's own saved settings into this browser session (nothing is shared between users)."""
    for key, (name, default) in USER_PREFS.items():
        v = rec.get(name, default)
        st.session_state[key] = list(v) if isinstance(v, list) else v
    st.session_state.ui_dark = rec.get("theme", DEFAULT_THEME) == "dark"
    st.session_state.pop("_excl", None)
    for k in [k for k in st.session_state.keys() if str(k).startswith("cat_") and k != "cat_off"]:
        st.session_state.pop(k, None)             # old switch positions from another user


def logout(message=None):
    uid = st.session_state.get("user")
    if uid:
        audit(uid, "logout", message or "")
    for k in ("user", "last_seen", "history", "clarify_for", "pending", "last_q", "_excl", "_prefs_for",
              *USER_PREFS):
        st.session_state.pop(k, None)
    st.session_state.login_message = message


def auth_gate():
    """Shows setup / login / change-password screens. Returns (id, user record) when signed in."""
    users = load_users()
    if not users:
        setup_screen()
        st.stop()
    uid = st.session_state.get("user")
    if uid and (uid not in users or not users[uid].get("active", True)):
        logout("Your access was removed.")
        uid = None
    if uid and SESSION_MINUTES and time.time() - st.session_state.get("last_seen", 0) > SESSION_MINUTES * 60:
        logout(f"Signed out after {SESSION_MINUTES} minutes of no activity.")
        uid = None
    if not uid:
        login_screen(st.session_state.pop("login_message", None))
        st.stop()
    st.session_state.last_seen = time.time()
    if users[uid].get("must_change"):
        login_banner(f"Welcome {users[uid].get('name', uid)}! Please set your own password to continue.")
        _, mid, _ = st.columns([1, 1.3, 1])
        with mid:
            change_password_form(uid, forced=True)
        st.stop()
    return uid, users[uid]


def admin_panel(me):
    users = load_users()
    with st.expander("👥 Manage users"):
        st.markdown("**Add a person**")
        with st.form("add_user", clear_on_submit=True):
            nid = st.text_input("User ID (e.g. priya.shah)").strip().lower()
            nname = st.text_input("Full name")
            nrole = st.selectbox("Role", ["user", "admin"],
                                 help="admin = can also add / remove users and see the activity log")
            npw = st.text_input("Temporary password (blank = make one for me)")
            if st.form_submit_button("➕ Add user", use_container_width=True):
                npw = npw or temp_password()
                err = (None if ID_RULE.match(nid) else "ID: 3-30 characters, letters/numbers/._- only.") \
                    or (None if nid not in users else "That ID already exists.") or pw_problem(npw)
                if err:
                    st.error(err)
                else:
                    rec = {"name": nname or nid, "role": nrole, "active": True,
                           "created": datetime.date.today().isoformat(), "created_by": me}
                    set_password(rec, npw, must_change=True)
                    with _users_lock():
                        fresh = load_users()
                        fresh[nid] = rec
                        save_users(fresh)
                    audit(me, "user_added", f"{nid} ({nrole})")
                    st.success(f"Added. Give them → ID: `{nid}` · Password: `{npw}` "
                               "(they must change it at first login)")

        st.markdown("**Existing users**")
        st.dataframe(pd.DataFrame([{"ID": k, "Name": v.get("name"), "Role": v.get("role"),
                                    "Active": "✅" if v.get("active", True) else "⛔",
                                    "Created": v.get("created", "")} for k, v in users.items()]),
                     hide_index=True, use_container_width=True)
        pick = st.selectbox("Choose a user", [u for u in users if u != me] or ["(no other users)"])
        if pick in users:
            c1, c2, c3 = st.columns(3)
            if c1.button("🔑 Reset", use_container_width=True, help="Give a new temporary password"):
                npw = temp_password()
                set_password(users[pick], npw, must_change=True)
                update_user(pick, salt=users[pick]["salt"], hash=users[pick]["hash"], must_change=True)
                audit(me, "password_reset", pick)
                st.success(f"New password for `{pick}`: `{npw}`")
            active = users[pick].get("active", True)
            if c2.button("⛔ Disable" if active else "✅ Enable", use_container_width=True):
                update_user(pick, active=not active)
                audit(me, "user_disabled" if active else "user_enabled", pick)
                st.rerun()
            new_role = st.selectbox("Role", ["user", "admin"], index=0 if users[pick].get("role") != "admin" else 1,
                                    key=f"role_{pick}")
            if new_role != users[pick].get("role") and st.button(f"Change role to {new_role}", key=f"rolebtn_{pick}"):
                old_role = users[pick].get("role")
                update_user(pick, role=new_role)
                audit(me, "role_changed", f"{pick}: {old_role} -> {new_role}")
                POLICY_STORE.log(me, "CHANGE_ROLE", "USER", pick, "role", old_role, new_role)
                st.rerun()
            if c3.button("🗑️ Delete", use_container_width=True):
                with _users_lock():
                    fresh = load_users()
                    fresh.pop(pick, None)
                    save_users(fresh)
                audit(me, "user_deleted", pick)
                st.rerun()

    data_admin_panel(me, users)
    with st.expander("📜 Activity log"):
        try:
            log = pd.read_csv(AUDIT_FILE).tail(200).iloc[::-1]
            st.dataframe(log, hide_index=True, use_container_width=True, height=300)
            st.download_button("⬇️ Download log", AUDIT_FILE.read_bytes(), "audit_log.csv", "text/csv")
        except Exception:
            st.caption("No activity yet.")


# ------------------------------ KPI cards & charts ------------------------------
def kpis(report, df, basis="ctc"):
    """Returns list of (label, value, full value) for the metric cards."""
    if df is None or df.empty:
        return []

    def tot(c):
        return df[c].sum() if c in df else None

    def money(label, c):
        v = tot(c)
        return None if v is None else (label, inr_short(v), inr(v))

    cards = []
    if report in ("summary", "trend", "department", "category"):
        if "Period" in df and df["Period"].nunique() > 1:
            cards.append(("Months", str(df["Period"].nunique()), None))
        elif "Employees" in df:
            cards.append(("Employees", f"{int(tot('Employees')):,}", None))
        cards += [money("CTC", "CTC"), money("Base sales value", "BaseSalesValue"), money("Margin", "Margin")]
        if "BaseSalesValue" in df:
            bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
            den = bs if basis == "sales" else ct
            if den:
                cards.append((f"Margin % (on {'base sales' if basis == 'sales' else 'CTC'})",
                              f"{(bs - ct) * 100 / den:.2f}%", None))
    elif report == "plan":
        cards.append(("Rows", f"{len(df):,}", None))
        for lab, c in (("CTC", "CTC"), ("Base sales value", "BaseSalesValue"), ("Margin", "Margin")):
            if c in df and pd.api.types.is_numeric_dtype(df[c]):
                cards.append(money(lab, c))
        if {"BaseSalesValue", "CTC"} <= set(df.columns):
            bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
            den = bs if basis == "sales" else ct
            if den:
                cards.append((f"Margin % (on {'base sales' if basis == 'sales' else 'CTC'})",
                              f"{(bs - ct) * 100 / den:.2f}%", None))
    elif report == "projection":
        pcs = [c for c in df.columns if str(c).startswith("Projected CTC ")]
        pbs = [c for c in df.columns if str(c).startswith("Projected base sales ")]
        if pcs and pbs:
            ct, bs = float(df[pcs].sum().sum()), float(df[pbs].sum().sum())
            many = f" ({len(pcs)} months)" if len(pcs) > 1 else ""
            cards += [("Projected CTC" + many, inr_short(ct), inr(ct)),
                      ("Projected base sales" + many, inr_short(bs), inr(bs))]
            cards.append(("Projected margin", inr_short(bs - ct), inr(bs - ct)))
            if ct:
                cards.append(("Projected margin % (on CTC)", f"{(bs - ct) * 100 / ct:.2f}%", None))
    elif report in ("workorder_summary", "workorder_monthly"):
        cards.append(("Work orders", f"{df['WorkOrderNo'].nunique():,}", None))
        cards += [money("Base sales value", "BaseSalesValue"), money("CTC", "CTC"), money("Total margin", "Margin")]
        bs, ct = df["BaseSalesValue"].sum(), df["CTC"].sum()
        den = bs if basis == "sales" else ct
        if den:
            cards.append((f"Margin % (on {'base sales' if basis == 'sales' else 'CTC'})",
                          f"{(bs - ct) * 100 / den:.2f}%", None))
    elif report == "deductions":
        cards += [money("TDS / Income tax", "IncomeTax"), money("Total PF", "TotalPF"),
                  money("PT", "PT"), money("ESIC (employee)", "ESIC_Employee")]
    elif report == "components":
        cards += [money("Basic", "Basic"), money("HRA", "HRA"), money("Bonus", "Bonus"),
                  money("Gross", "GrossAmount")]
    elif report in ("employee_detail", "workorder_detail", "employee_list"):
        cards.append(("Records", f"{len(df):,}", None))
        if report == "workorder_detail" and "WorkOrderNo" in df:
            cards.append(("Work orders", f"{df['WorkOrderNo'].nunique():,}", None))
        if "Employee" in df:
            cards.append(("Employees", f"{df['Employee'].nunique():,}", None))
        cards += [money("Gross salary", "GrossAmount"), money("Take-home", "TakeHome")]
    elif report == "proposed_tds":
        cards.append(("Employees", f"{len(df):,}", None))
        cards += [money("Monthly TDS", "Monthly TDS"), money("Remaining TDS", "Remaining TDS"),
                  money("TDS already deducted", "TDS already deducted")]
    elif report == "salary_run":
        g = next((c for c in df.columns if str(c).startswith("Gross ")), None)
        if g:
            tl = str(g)[6:]
            cards.append(("Employees to pay", f"{int((pd.to_numeric(df[g], errors='coerce') > 0).sum()):,}", None))
            cards += [money(f"Gross {tl}", g), money(f"TDS {tl}", f"TDS {tl}"),
                      money(f"Take-home {tl}", f"Take-home {tl}")]
    elif report == "tenure" and "Status" in df:
        cards += [("Employees", f"{len(df):,}", None), ("Working", f"{int((df['Status'] == 'Working').sum()):,}", None),
                  ("Resigned", f"{int((df['Status'] == 'Resigned').sum()):,}", None)]
    elif report in ("joiners", "leavers"):
        cards.append(("Joiners" if report == "joiners" else "Leavers", f"{len(df):,}", None))
        if "Department" in df:
            cards.append(("Departments", f"{df['Department'].nunique():,}", None))
    return [c for c in cards if c]


def show_kpis(report, df, basis="ctc"):
    cards = kpis(report, df, basis)
    if not cards:
        return
    out = []
    for n, (label, value, full) in enumerate(cards):
        cls = "hi" if "%" in label and n == len(cards) - 1 else \
            "pos" if "margin" in label.lower() and not str(value).startswith("₹-") else ""
        out.append(f'<div class="lk-kpi {cls}"><div class="l">{html.escape(label)}</div>'
                   f'<div class="v">{html.escape(str(value))}</div>'
                   f'<div class="f">{html.escape(full) if full else "&nbsp;"}</div></div>')
    st.markdown('<div class="lk-kpis">' + "".join(out) + "</div>", unsafe_allow_html=True)


def _md_bold(text):
    """Answer text -> safe HTML (escape, then **bold** and line breaks)."""
    t = html.escape(str(text or ""))
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"<i>\1</i>", t)
    return t.replace("\n", "<br>")


def show_alt(chart):
    try:
        st.altair_chart(altair_theme(chart), use_container_width=True, theme=None)
    except TypeError:
        st.altair_chart(altair_theme(chart), use_container_width=True)


def _chart(kind, data, color, title=None, asc=False, pct=False):
    """Bar chart of a pandas Series (index = labels), sorted by value, in theme colours."""
    import altair as alt
    d = data.reset_index()
    d.columns = ["Label", "Value"]
    d["Label"] = d["Label"].astype(str)
    show_alt(alt.Chart(d).mark_bar(cornerRadiusEnd=4, color=color).encode(
        y=alt.Y("Label:N", sort="x" if asc else "-x", title=None), x=alt.X("Value:Q", title=title),
        tooltip=["Label", alt.Tooltip("Value:Q", format=".2f" if pct else ",.0f", title=title or "Value")]
    ).properties(height=max(200, 30 * len(d))))


def draw_chart(report, df, intent=None):
    """Draws a chart when it helps. Returns True if something was drawn."""
    intent = intent or {}
    if report == "plan":
        try:
            plan = intent.get("plan") or {}
            dim = next((("Name" if g_ == "Employee" else g_) for g_ in plan.get("group", []) if g_ != "Month"), None)
            srt = plan.get("sort")
            names = {"Employees": "Employee count", "InvoiceAmount": "Invoice amount", "AvgCTC": "Average CTC"}
            cand = ([names.get(srt[0], srt[0])] if srt else []) + ["Margin", "CTC", "BaseSalesValue", "Employee count"]
            met = next((c for c in cand if c in df and pd.api.types.is_numeric_dtype(df[c])), None)
            if dim and dim in df and met:
                data = df.groupby(dim)[met].sum().sort_values(ascending=False).head(25)
                _chart("bar", data, "#E65F26" if met == "Margin" else CHART_MAIN, title=met, asc=False)
                return True
            mc = intent.get("month_col")
            if not dim and mc in df and df[mc].nunique() > 1:      # month-wise -> line chart in month order
                import altair as alt
                series = [c for c in ("CTC", "BaseSalesValue", "Margin") if c in df] or ([met] if met else [])
                if series:
                    data = df.melt(mc, series, var_name="Type", value_name="Amount")
                    show_alt(alt.Chart(data).mark_line(point=True, strokeWidth=3).encode(
                        x=alt.X(f"{mc}:N", sort=list(df[mc]), title=None, axis=alt.Axis(labelAngle=0)),
                        y=alt.Y("Amount:Q", title=None),
                        color=alt.Color("Type:N", scale=alt.Scale(range=[CHART_MAIN, "#E65F26", "#0f9d6b"]),
                                        legend=alt.Legend(orient="bottom", title=None)),
                        tooltip=[mc, "Type", alt.Tooltip("Amount:Q", format=",.0f")]).properties(height=340))
                    return True
        except Exception:
            pass
        return False
    fc, fl = intent.get("focus", "CTC"), intent.get("focus_label", "CTC")
    asc = intent.get("sort") == "asc"
    try:
        if "Period" in df and df["Period"].nunique() > 1 and "CTC" in df \
                and report in ("summary", "trend"):
            import altair as alt   # comes with Streamlit
            order = list(df["Period"])            # already in month order
            series = [c for c in ("CTC", "BaseSalesValue", "Margin") if c in df] or ["CTC"]
            data = df.melt("Period", series, var_name="Type", value_name="Amount")
            chart = alt.Chart(data).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("Period:N", sort=order, title=None, axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Amount:Q", title=None),
                color=alt.Color("Type:N", scale=alt.Scale(range=[CHART_MAIN, "#E65F26", "#0f9d6b"]),
                                legend=alt.Legend(orient="bottom", title=None)),
                tooltip=["Period", "Type", alt.Tooltip("Amount:Q", format=",.0f")],
            ).properties(height=340)
            show_alt(chart)
            return True
        if report == "analysis" and 2 <= len(df) <= 60:
            import altair as alt
            num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
            txt = [c for c in df.columns if c not in num]
            if num and txt:
                d = df[[txt[0], num[0]]].copy()
                d[txt[0]] = d[txt[0]].astype(str)
                show_alt(alt.Chart(d).mark_bar(cornerRadiusEnd=4, color=CHART_MAIN).encode(
                    y=alt.Y(f"{txt[0]}:N", sort="-x", title=None),
                    x=alt.X(f"{num[0]}:Q", title=num[0]),
                    tooltip=list(d.columns)).properties(height=max(220, 24 * len(d))))
                return True
        if report == "employee_ranking" and "Name" in df and len(df) >= 2:
            import altair as alt
            m = df.columns[-1]
            d = df[["Name", m]].copy()
            d["Name"] = d["Name"].astype(str)
            show_alt(alt.Chart(d).mark_bar(cornerRadiusEnd=4, color=CHART_MAIN).encode(
                y=alt.Y("Name:N", sort=list(d["Name"]), title=None), x=alt.X(f"{m}:Q", title=m),
                tooltip=["Name", alt.Tooltip(f"{m}:Q", format=",.0f")]).properties(height=max(220, 28 * len(d))))
            return True
        if report == "workorder_summary" and "Margin" in df:
            import altair as alt
            d = df.head(15)
            m = fc if fc in d and fc != "Employees" else "Margin"
            show_alt(alt.Chart(d).mark_bar(cornerRadiusEnd=4).encode(
                y=alt.Y("WorkOrderNo:N", sort=list(d["WorkOrderNo"]), title=None),
                x=alt.X(f"{m}:Q", title=(fl if m == fc else "Margin") + ("" if m == "MarginPct" else " (₹)")),
                color=alt.condition(getattr(alt.datum, m) < 0, alt.value("#E65F26"), alt.value(CHART_MAIN)),
                tooltip=["WorkOrderNo", alt.Tooltip("Margin:Q", format=",.0f"),
                         alt.Tooltip("MarginPct:Q", format=".2f", title="Margin %"),
                         alt.Tooltip("BaseSalesValue:Q", format=",.0f", title="Base sales value"),
                         alt.Tooltip("CTC:Q", format=",.0f")],
            ).properties(height=max(220, 28 * len(d))))
            return True
        if report in ("department", "category"):
            col = report.capitalize()
            g = df.groupby(col).sum(numeric_only=True)
            if fc == "MarginPct" and {"Margin", "CTC"} <= set(g.columns):
                den = g["BaseSalesValue"] if intent.get("pct_basis") == "sales" else g["CTC"]
                data = (g["Margin"] * 100 / den.replace(0, float("nan"))).dropna()
                title = f"Margin % (on {'base sales' if intent.get('pct_basis') == 'sales' else 'CTC'})"
            else:
                use = fc if fc in g else "CTC"
                data, title = g[use], (fl if use == fc else "CTC")
            data = data.sort_values(ascending=asc).head(25)
            _chart("bar", data, "#E65F26" if fc in ("Margin", "MarginPct") else CHART_MAIN, title=title,
                   asc=asc, pct=fc == "MarginPct")
            return True
        if report in ("deductions", "components"):
            skip = {"Employees", "Period"}
            sums = df[[c for c in df.columns if c not in skip]].select_dtypes("number").sum()
            sums = sums[sums != 0].sort_values(ascending=False)
            if not sums.empty:
                _chart("bar", sums.rename("Amount"), "#E65F26")
                return True
        if report in ("joiners", "leavers") and "Department" in df:
            _chart("bar", df["Department"].fillna("(blank)").value_counts().rename("People"),
                   CHART_MAIN)
            return True
    except Exception:
        pass
    return False


def view_df(df, intent=None):
    """Table / CSV view: 'Period' named after the month column used, month_year shown as Liable month + Payment date."""
    had_my = df is not None and "month_year" in df.columns
    df = add_payment_date(df)
    pc = (intent or {}).get("period_col")
    if pc == "Liable month" and had_my and "Period" in df.columns:
        df = df.drop(columns=["Period"])               # the Liable month (month_year) column already says it
    if pc and "Period" in df.columns and pc not in df.columns:
        df = df.rename(columns={"Period": pc})
    return df


def show_table(df, key, basis="ctc"):
    names = {"MarginPct": f"Margin % (on {'base sales' if basis == 'sales' else 'CTC'})",
             "BaseSalesValue": "Base sales value", "TakeHome": "Take-home", "GrossAmount": "Gross",
             "IncomeTax": "TDS", "TotalPF": "Total PF", "InvoiceWithGST": "Invoice (with GST)"}
    df = df.rename(columns={k: v for k, v in names.items() if k in df.columns})
    money_cols = {c for c in df.select_dtypes("number").columns
                  if c not in ("Employee", "Employees", "PaidDays", "Month", "Year")}
    try:
        cfg = {c: st.column_config.NumberColumn(c, format="localized") for c in money_cols}
        st.dataframe(df, use_container_width=True, hide_index=True, column_config=cfg,
                     height=min(420, 38 + 35 * len(df)))
    except Exception:   # older Streamlit without "localized" format
        st.dataframe(df, use_container_width=True, hide_index=True,
                     height=min(420, 38 + 35 * len(df)))


# ------------------------------ Chat rendering ------------------------------
def full_csv(msg):
    """Runs the same read-only query again without the 1,000-row limit and returns CSV bytes."""
    sql = msg.get("sql") or ""
    if response_rule_hit(msg.get("question", ""), None, log=True):  # v14: response rules apply to downloads too
        raise PermissionError("Not available for your login.")
    if msg.get("report") in ("proposed_tds", "salary_run", "tenure", "wo_history", "wo_breaks"):   # the full list
        _, d_ = _display(msg)
        if d_ is None:
            raise PermissionError("Not available for your login.")
        return d_.to_csv(index=False).encode("utf-8-sig"), len(d_)
    if msg.get("report") == "analysis":
        df = run_ai_query(sql, limit=CSV_LIMIT)
        df = prettify(df) if "prettify" in globals() else df
    else:
        df = _run_query_raw(re.sub(rf"(?i)\bTOP\s*\(\s*{ROW_LIMIT}\s*\)\s*", "", sql, count=1),
                            list(msg.get("params") or []))
    try:
        df = hide_columns(df, msg.get("question", ""), msg.get("report", ""))
    except Exception:
        pass
    eff = current_eff()
    if eff is not None:
        df, _, block = pe.secure_df(df, eff, current_prefs(), msg.get("question", ""), msg.get("report", ""),
                                    (msg.get("intent") or {}).get("explicit_cols"))
        if block or df is None:
            raise PermissionError("Not available for your login.")
    df = view_df(df, msg.get("intent"))
    return df.to_csv(index=False).encode("utf-8-sig"), len(df)      # utf-8-sig: opens fine in Excel


def secure_result(result, question):
    """Final security boundary for a NEW answer: permissions, then preferences, then text."""
    eff = current_eff()
    if eff is None or result.get("report") in ("policy", "policy_confirm", "pref", "blocked", "definition",
                                               "chat", "chat_ai", "clarify", "teach", "teach_confirm", "reminder"):
        return result
    if result.get("report") == "admin_info":
        if not eff.admin_info:
            return {**result, "df": None, "text": "🔒 I can't share administration details."}
        return result
    df, notes, block = pe.secure_df(result.get("df"), eff, current_prefs(), question, result.get("report", ""),
                                    (result.get("intent") or {}).get("explicit_cols"))
    if block:
        audit_policy_event("result_blocked", f"{result.get('report')}: {question[:120]}")
        return {**result, "df": None, "text": block, "explain": "", "explain_pending": False,
                "note": "Data permissions / your preferences were applied.", "options": []}
    out = dict(result, df=df)
    out["text"] = pe.secure_text(result.get("text", ""), eff, current_prefs(), question)
    if notes:
        n0 = result.get("note", "") or ""
        if "⏱️" in n0:                                   # keep the timer chip last
            a_, b_ = n0.split("⏱️", 1)
            out["note"] = f"{a_.strip()} {' '.join(notes)} ⏱️{b_}".strip()
        else:
            out["note"] = (n0 + " " + " ".join(notes)).strip()
    return out


def _display(msg):
    """Re-checks a stored answer with the CURRENT policy every time it is shown or downloaded."""
    eff = current_eff()
    df = msg.get("df")
    if msg.get("question") and msg.get("report") not in NO_EXPLAIN and TEACH_STORE is not None:
        r = response_rule_hit(msg["question"], None, log=False)       # v14: a rule added later also hides old answers
        if r:
            return tg.rule_reply(r) or "This information is not available.", None
    if eff is None or df is None or msg.get("report") in ("admin_info",):
        if msg.get("report") == "admin_info" and eff is not None and not eff.admin_info:
            return "🔒 I can't share administration details.", None
        return msg.get("text", ""), df
    df2, _, block = pe.secure_df(df, eff, current_prefs(), msg.get("question", ""), msg.get("report", ""),
                                 (msg.get("intent") or {}).get("explicit_cols"))
    if block:
        return block, None
    return pe.secure_text(msg.get("text", ""), eff, current_prefs(), msg.get("question", "")), df2


def _chain(first, gen):
    yield first
    yield from gen


def render(i, msg):
    if msg["role"] == "user":
        if IS_ADMIN and is_user_sql(msg["text"]):                  # v14: SQL mode -> coloured like SSMS
            q_ = str(msg["text"]).strip()
            q_ = "/" + (q_[1:].strip() if q_.startswith("/") else q_)   # copy button gives the query WITH "/"
            with st.container(key=f"usql_{i}"):
                try:
                    st.code(q_, language="sql", wrap_lines=True)
                except TypeError:                                          # older Streamlit
                    st.code(q_, language="sql")
        else:
            ini = "".join(w[0] for w in str(ME_REC.get("name") or ME).split()[:2]).upper()
            with st.container(key=f"umsg_{i}"):
                st.markdown(f"""<div class="lk-u"><div class="b"><span class="ic">💼</span>
<span>{html.escape(str(msg['text']))}</span></div><div class="av">{html.escape(ini)}</div></div>""",
                            unsafe_allow_html=True)
        return

    with st.chat_message("assistant", avatar="🤖"):
        if msg.get("explain"):
            eff_ = current_eff()
            st.markdown(pe.secure_text(msg["explain"], eff_, current_prefs(), msg.get("question", ""))
                        if eff_ is not None else msg["explain"])
        elif msg.get("explain_pending"):
            try:
                with st.spinner("✍️ Writing the answer..."):
                    gen = explain_stream(msg)
                    first = next(gen)                # wait for the first words, then type the rest
                msg["explain"] = st.write_stream(_chain(first, gen))
            except Exception as e:
                if msg.get("report") == "chat_ai":
                    msg["explain"] = (f"⚠️ The AI is not reachable right now ({e.__class__.__name__}). "
                                      "Please ask a payroll question like *salary for aug-26*.")
                    st.markdown(msg["explain"])
            msg["explain_pending"] = False
        shown_text, df = _display(msg)                   # re-checked with the CURRENT policy
        has_df = df is not None and not df.empty
        if msg.get("text"):
            if has_df:                                        # data answer: header row with icon, live chip, export
                h1, h2 = st.columns([7, 2.2], vertical_alignment="top")
                h1.markdown(f'<div class="lk-ans"><div class="ic">{BOT_FACE}</div><div class="tx">'
                            f'{_md_bold(shown_text)}</div></div>', unsafe_allow_html=True)
                with h2.container(key=f"exp_{i}"):
                    st.markdown('<div style="display:flex;justify-content:flex-end">'
                                '<span class="lk-live">✓ Verified Live</span></div>', unsafe_allow_html=True)
                    st.download_button(":material/download: Export",
                                       view_df(df, msg.get("intent")).to_csv(index=False).encode("utf-8-sig"),
                                       file_name=f"{msg['report']}.csv", mime="text/csv", key=f"dlx{i}",
                                       use_container_width=True)
            elif msg.get("explain") and df is not None:
                st.caption(re.sub(r"\*\*", "", shown_text))   # the exact figures line, small
            else:
                st.markdown(shown_text)
        if has_df:
            basis = (msg.get("intent") or {}).get("pct_basis", "ctc")
            show_kpis(msg["report"], df, basis)
            names = ["📊 Chart", "📋 Table", "</> SQL Query", "💡 Executive Insights"]
            if not current_prefs().get("charts", True):           # "never show me charts"
                names = ["📋 Table", "</> SQL Query", "💡 Executive Insights", "📊 Chart"]
            tabs_ = dict(zip(names, st.tabs(names)))
            tab_chart, tab_table, tab_sql = tabs_["📊 Chart"], tabs_["📋 Table"], tabs_["</> SQL Query"]
            with tabs_["💡 Executive Insights"]:
                ins = msg.get("explain") or ""
                if ins:
                    eff_ = current_eff()
                    st.markdown(pe.secure_text(ins, eff_, current_prefs(), msg.get("question", ""))
                                if eff_ is not None else ins)
                else:
                    st.caption("Add *explain* or *why* to your question for a written analysis.")
                if msg.get("note"):
                    st.caption("How I read your question: " + re.sub(r"\*\*", "", msg["note"].split("⏱️")[0]))
            with tab_chart:
                if not current_prefs().get("charts", True):
                    st.caption("Charts are off in your preferences.")
                elif not draw_chart(msg["report"], df, msg.get("intent")):
                    st.caption("No chart for this report. See the Table tab.")
            with tab_table:
                show_table(view_df(df, msg.get("intent")), i, (msg.get("intent") or {}).get("pct_basis", "ctc"))
                st.download_button("⬇️ Download CSV",
                                   view_df(df, msg.get("intent")).to_csv(index=False).encode("utf-8-sig"),
                                   file_name=f"{msg['report']}.csv", mime="text/csv", key=f"dl{i}")
                if len(df) >= ROW_LIMIT and msg.get("sql"):
                    st.caption(f"Only the first {ROW_LIMIT:,} rows are shown here.")
                    if msg.get("full_csv") is None:
                        if st.button("📦 Prepare CSV with ALL rows", key=f"full{i}"):
                            with st.spinner("Reading all rows..."):
                                try:
                                    msg["full_csv"], msg["full_n"] = full_csv(msg)
                                except Exception as e:
                                    st.error(f"Could not read all rows: {e}")
                            st.rerun()
                    else:
                        st.download_button(f"⬇️ Download ALL {msg['full_n']:,} rows (CSV)", msg["full_csv"],
                                           file_name=f"{msg['report']}_all.csv", mime="text/csv", key=f"dla{i}")
            with tab_sql:
                st.code(msg.get("sql", "") if IS_ADMIN else pe.redact_sql(msg.get("sql", ""), RAW_T), language="sql")
                st.write("Parameters:", [str(p) for p in msg.get("params", [])])
                st.write("Understood:", msg.get("intent"))
        elif msg.get("sql"):
            with st.expander("🧾 Show SQL", expanded=msg.get("report") == "analysis"):
                st.code(msg["sql"] if IS_ADMIN else pe.redact_sql(msg["sql"], RAW_T), language="sql")
        if msg.get("options"):
            cols = st.columns(min(len(msg["options"]), 3))
            for n, opt in enumerate(msg["options"]):
                if cols[n % len(cols)].button(opt, key=f"opt_{i}_{n}", use_container_width=True):
                    st.session_state.pending = opt
                    st.session_state.clarify_for = None
                    st.rerun()
        if msg.get("note"):
            parts = msg["note"].split("⏱️")
            chip = f'<span class="timechip">⏱️ {parts[1].strip()}</span>' if len(parts) > 1 else ""
            chip = f'<span class="timechip">⏱️ {html.escape(parts[1].strip())}</span>' if len(parts) > 1 else ""
            text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", html.escape(parts[0].strip()))
            st.markdown(f'<div class="note">{chip}{text}</div>', unsafe_allow_html=True)


ME, ME_REC = auth_gate()
IS_ADMIN = ME_REC.get("role") == "master"          # the old backup site: only the master gets its admin pages
POLICY_STORE = pe.PolicyStore(APP_DIR)        # policies.json + policy_audit.csv (admin only, separate from chats)
USERS_LOADER = load_users


def _save_data_prefs(new):
    update_user(ME, data_prefs=new)             # this user's own record only
    ME_REC["data_prefs"] = new


PREFS_SAVER = _save_data_prefs
TEACH_STORE = tg.TeachingStore(APP_DIR)       # v14: teachings.json + teaching_log.csv
REMINDER_STORE = rmd.ReminderStore(APP_DIR)   # v14: reminders.json (each login sees only its own)
EFF = set_policy_context(ME, ME_REC.get("role"), ME_REC)

if "history" not in st.session_state:
    st.session_state.history = []
st.session_state.setdefault("page", "Chat")
if st.session_state.get("_prefs_for") != ME:          # a different person signed in on this browser
    load_prefs(ME_REC)
    st.session_state._prefs_for = ME
try:
    _warm_up_ai()
except Exception:
    pass
default_mode = st.session_state.cfg_mode
always_ai = st.session_state.cfg_always_ai
use_ai = always_ai or st.session_state.cfg_use_ai


def _save_hide(hide):
    update_user(ME, hide_cols=list(hide))


PREF_SAVER = _save_hide
def _save_cat_off(off):
    update_user(ME, cat_off=list(off))


CAT_OFF_SAVER = _save_cat_off


def _save_theme():
    update_user(ME, theme="dark" if st.session_state.get("ui_dark") else "light")


def save_answer_prefs():
    """Saves the Answers settings of the signed-in user only."""
    new = {name: st.session_state.get(key) for key, (name, _) in USER_PREFS.items()
           if key.startswith("cfg_")}
    if any(ME_REC.get(k) != v for k, v in new.items()):
        update_user(ME, **new)
        ME_REC.update(new)


def go(page):
    st.session_state.page = page


# ------------------------------ Sidebar ------------------------------
ADMIN_BADGE_CSS = """section[data-testid="stSidebar"] div.st-key-nav_Settings button:after {content: "RBAC"; position: absolute; right: 14px; top: 50%;
     transform: translateY(-50%); font-size: .66rem; font-weight: 700; padding: 4px 8px; border-radius: 6px;
     border: 1px solid rgba(244,160,74,.5); color: #f4b26a; background: rgba(244,160,74,.08);}
"""
PAGES = [("Chat", ":material/chat:"), ("My Requests", ":material/history:"),
         ("Documents", ":material/description:"), ("FAQs", ":material/help:"), ("Settings", ":material/tune:")]
NAV_LABELS = {"Chat": "AI Chat & Assistant", "My Requests": "Query History", "Documents": "Documents",
              "FAQs": "Knowledgebase & FAQs", "Settings": "System Settings"}

with st.sidebar:
    online = db_ok() and ai_ok()
    _logo = logo_html(50)
    _mark = _logo if _logo.startswith("<img") else '<div class="mk">S</div>'
    st.markdown(f"""
<div class="lk-brand">{_mark}<div><div class="t1">WORKORDER</div><div class="t2">DUMMY</div></div></div>
<div class="lk-bot"><div class="av">{BOT_FACE}<i style="background:{'#22c55e' if online else '#f59e0b'}"></i></div>
  <div><div class="n">Payroll <b>AI</b> Assistant</div>
  <div class="s" style="color:{'#22c55e' if online else '#f59e0b'} !important">● {'Online &amp; Ready' if online else 'Limited - check SQL / AI'}</div></div></div>
<div class="lk-sub">Ask · Get Answers · Stay Updated</div>""", unsafe_allow_html=True)
    for name, icon in PAGES:
        st.button(f"{icon}&nbsp;&nbsp;{NAV_LABELS.get(name, name)}", key=f"nav_{name.replace(' ', '_')}",
                  use_container_width=True, on_click=go, args=(name,))
    st.markdown('<div style="height:18px"></div>', unsafe_allow_html=True)
    _ini = "".join(w[0] for w in str(ME_REC.get("name") or ME).split()[:2]).upper()
    u1, u2 = st.columns([5, 1.3], vertical_alignment="center")
    u1.markdown(f"""<div class="lk-user"><div class="ic">{html.escape(_ini)}</div>
  <div><div class="n">{html.escape(str(ME_REC.get('name', ME)))}</div>
  <div class="r">{'Finance Admin' if IS_ADMIN else html.escape(str(ME_REC.get('role', 'user')))}</div></div></div>""",
                unsafe_allow_html=True)
    if u2.button(":material/logout:", key="sb_logout", use_container_width=True, help="Log out"):
        logout("You have logged out.")
        st.rerun()
    if IS_ADMIN or ME_REC.get("role") == "admin":       # master and admins may refresh the data
        last = _refresh_status()
        if not st.session_state.get("confirm_refresh"):
            if st.button(":material/sync:&nbsp;&nbsp;Refresh Database", key="sb_refresh", use_container_width=True,
                         help=f"Master and admins only. Reloads {TABLE} from the source server."):
                st.session_state.confirm_refresh = True
                st.rerun()
        else:
            st.warning(f"This **replaces all rows** in {TABLE} with fresh data from {REFRESH_SOURCE} "
                       f"and runs {REFRESH_PROC}. Continue?")
            c1, c2 = st.columns(2)
            if c1.button("✅ Yes", key="sb_refresh_yes", use_container_width=True):
                st.session_state.confirm_refresh = False
                with st.spinner("Refreshing payroll data... (can take a few minutes)"):
                    try:
                        old, new, secs = refresh_payroll_data()
                        _save_refresh_status(ME, old, new, secs, "ok")
                        audit(ME, "data_refresh", f"{old} -> {new} rows in {secs:.0f}s")
                        st.cache_data.clear()            # answers now use the new data
                        st.session_state.refresh_msg = ("ok", f"Done: {new:,} rows (before: {old:,}) in {secs:.0f}s.")
                    except Exception as e:
                        _save_refresh_status(ME, None, None, 0, f"failed: {str(e)[:200]}")
                        audit(ME, "data_refresh_failed", str(e)[:300])
                        st.session_state.refresh_msg = ("err", f"Refresh failed, {TABLE} was NOT changed. ({str(e)[:300]})"
                                                        if REFRESH_ALL_OR_NOTHING else f"Refresh failed: {str(e)[:300]}")
                st.rerun()
            if c2.button("Cancel", key="sb_refresh_no", use_container_width=True):
                st.session_state.confirm_refresh = False
                st.rerun()
        msg = st.session_state.pop("refresh_msg", None)
        if msg:
            (st.success if msg[0] == "ok" else st.error)(msg[1])
        if last:
            ok_ = str(last.get("status", "")).startswith("ok")
            st.markdown(f"""<div class="lk-sync"><span>Synced: {html.escape(str(last.get('at', '?')))}</span>
  <b style="color:{'#22c55e' if ok_ else '#f59e0b'} !important">● {'OK' if ok_ else 'check'}</b></div>""",
                        unsafe_allow_html=True)
    st.markdown("""<div class="sb-tag">Together<br>We Build a<br><b>Brighter Future</b></div>
<div class="sb-swoosh"></div>""", unsafe_allow_html=True)

PAGE = st.session_state.page
st.markdown(f"""<style>
section[data-testid="stSidebar"] div.st-key-nav_{PAGE.replace(' ', '_')} button {{
  background: #2f6fe4 !important; color: #fff !important;
  box-shadow: 0 8px 20px rgba(47,111,228,.35) !important; font-weight: 600 !important;}}
section[data-testid="stSidebar"] div.st-key-nav_{PAGE.replace(' ', '_')} button p {{color: #fff !important; font-weight: 600 !important;}}
</style>""" + ("<style>" + ADMIN_BADGE_CSS + "</style>" if IS_ADMIN else ""), unsafe_allow_html=True)

# ------------------------------ Top bar ------------------------------
_t1, _t2, _t3 = st.columns([4.2, 4.2, 1.2], vertical_alignment="center")
_t1.markdown("""<div class="lk-top"><span class="co">WORKORDER DUMMY</span>
<span class="tg">| People · Process · Progress</span></div>""", unsafe_allow_html=True)
_d_ok, _a_ok = db_ok(), ai_ok()
_t2.markdown(f"""<div class="lk-pills">
  <span class="lk-pill {'g' if _d_ok else 'r'}"><i></i>SQL Server</span>
  <span class="lk-pill {'p' if _a_ok else 'r'}">Ollama AI (Local)</span>
  <span class="lk-pill o">Read-Only Data</span></div>""", unsafe_allow_html=True)
with _t3.container(key="lk_theme"):
    st.toggle("🌙 Dark", key="ui_dark", on_change=_save_theme,
              help="Your own choice - it doesn't change other people's screens")


def status_pills():
    d, a = db_ok(), ai_ok()
    return f"""<div class="pills">
  <span class="pl {'g' if d else 'r'}"><i></i>SQL Server</span>
  <span class="pl {'p' if a else 'r'}">🧠 Ollama AI</span>
  <span class="pl o">🛡️ Read-only · data cannot be changed</span>
  <span class="pl b">💻 Runs on this PC</span>
  <span class="pl b">💬 {'AI on every question' if always_ai else 'Fast mode'}</span></div>"""


def latest_label():
    try:
        m, y = latest_period(default_mode)
        return period_label(m, y)
    except Exception:
        return "Aug-26"


def prev_label(lbl):
    try:
        d = datetime.datetime.strptime("01-" + lbl, "%d-%b-%Y")
        p = (d.replace(day=1) - datetime.timedelta(days=1))
        return p.strftime("%b-%Y")
    except Exception:
        return lbl


def month_back(lbl, n):
    for _ in range(n):
        lbl = prev_label(lbl)
    return lbl


# ================================ PAGES ================================
def _col_options():
    return sorted(set(VOCAB.all_keys) - {"Year"}, key=str.lower)


def _policy_save(me, action, reason):
    msg = POLICY_STORE.apply(action, me, reason)
    audit(me, "policy_changed", msg[:200])
    set_policy_context(ME, ME_REC.get("role"), ME_REC)
    st.success("Saved and logged: " + msg)


def data_admin_panel(me, users):
    """Admin: data permissions (global + per user), user preferences, policy audit log."""
    data = POLICY_STORE.load()
    opts = _col_options()
    fmt = pe.label
    with st.expander("🔐 Data permissions — everyone (global)"):
        g = data["global"]
        st.caption("Global rules apply to **every** login (admins too). Users can never override them.")
        with st.form("pol_global"):
            blocked = st.multiselect("Blocked columns (nobody sees these)", opts, default=[c for c in g["blocked_columns"] if c in opts],
                                     format_func=fmt)
            allowed = st.multiselect("Allowed columns (empty = all columns allowed)", opts,
                                     default=[c for c in g["allowed_columns"] if c in opts], format_func=fmt)
            sens = st.multiselect("Sensitive columns (hidden from users without 'sensitive data' access)", opts,
                                  default=[c for c in g["sensitive_columns"] if c in opts], format_func=fmt)
            info = st.toggle("Normal users may see user / admin information", value=bool(g.get("admin_info_for_users")))
            reason = st.text_input("Reason (for the audit log)")
            ok = st.checkbox("I confirm this security change")
            if st.form_submit_button("💾 Save global policy", type="primary"):
                if not ok:
                    st.warning("Tick *I confirm this security change* first.")
                else:
                    add, rem = set(blocked) - set(g["blocked_columns"]), set(g["blocked_columns"]) - set(blocked)
                    if add:
                        _policy_save(me, {"op": "BLOCK_COLUMN", "scope": "GLOBAL", "columns": sorted(add)}, reason)
                    if rem:
                        _policy_save(me, {"op": "UNBLOCK_COLUMN", "scope": "GLOBAL", "columns": sorted(rem)}, reason)
                    if sorted(allowed) != sorted(g["allowed_columns"]):
                        _policy_save(me, {"op": "SET_ALLOWED", "scope": "GLOBAL", "columns": allowed}, reason)
                    if sorted(sens) != sorted(g["sensitive_columns"]):
                        _policy_save(me, {"op": "SET_SENSITIVE_LIST", "scope": "GLOBAL", "columns": sens}, reason)
                    if bool(info) != bool(g.get("admin_info_for_users")):
                        _policy_save(me, {"op": "SET_ADMIN_INFO", "scope": "GLOBAL", "value": info}, reason)
                    if not (add or rem) and sorted(allowed) == sorted(g["allowed_columns"]) and \
                            sorted(sens) == sorted(g["sensitive_columns"]) and bool(info) == bool(g.get("admin_info_for_users")):
                        st.info("No change.")
    with st.expander("👤 Data permissions — one user"):
        pick = st.selectbox("User", list(users), key="pol_user")
        if pick:
            up = POLICY_STORE.user(data, pick)
            with st.form("pol_user_form"):
                ub = st.multiselect("Blocked columns for this user", opts, default=[c for c in up["blocked_columns"] if c in opts],
                                    format_func=fmt)
                ua = st.multiselect("Allowed columns for this user (empty = all)", opts,
                                    default=[c for c in up["allowed_columns"] if c in opts], format_func=fmt)
                c1, c2, c3 = st.columns(3)
                pa = c1.toggle("Payroll access", value=up["payroll_access"])
                ag = c2.toggle("Aggregated (summary) data only", value=up["aggregated_only"])
                se = c3.toggle("Sensitive data", value=up["sensitive_access"])
                tb = st.text_input("Blocked tables / views (comma separated)", ", ".join(up.get("blocked_tables", [])))
                reason = st.text_input("Reason (for the audit log)", key="pol_user_reason")
                ok = st.checkbox("I confirm this security change", key="pol_user_ok")
                cA, cB = st.columns(2)
                save = cA.form_submit_button("💾 Save user policy", type="primary")
                reset = cB.form_submit_button("↩️ Reset to default access")
                if (save or reset) and not ok:
                    st.warning("Tick *I confirm this security change* first.")
                elif reset:
                    _policy_save(me, {"op": "RESET_USER_POLICY", "scope": "USER", "user": pick}, reason)
                elif save:
                    base = {"scope": "USER", "user": pick}
                    add, rem = set(ub) - set(up["blocked_columns"]), set(up["blocked_columns"]) - set(ub)
                    if add:
                        _policy_save(me, dict(base, op="BLOCK_COLUMN", columns=sorted(add)), reason)
                    if rem:
                        _policy_save(me, dict(base, op="UNBLOCK_COLUMN", columns=sorted(rem)), reason)
                    if sorted(ua) != sorted(up["allowed_columns"]):
                        _policy_save(me, dict(base, op="SET_ALLOWED", columns=ua), reason)
                    for op, key, val in (("SET_PAYROLL", "payroll_access", pa), ("SET_AGG_ONLY", "aggregated_only", ag),
                                         ("SET_SENSITIVE", "sensitive_access", se)):
                        if bool(val) != bool(up[key]):
                            _policy_save(me, dict(base, op=op, value=val), reason)
                    tables = [t.strip() for t in tb.split(",") if t.strip()]
                    if sorted(tables) != sorted(up.get("blocked_tables", [])):
                        _policy_save(me, dict(base, op="SET_TABLES", columns=tables), reason)
            e = pe.Effective(POLICY_STORE.load(), pick, users[pick].get("role"), VOCAB, table=TABLE, store=POLICY_STORE)
            st.markdown("**What this user can see right now (global + user rules):**")
            st.json(e.summary())
    with st.expander("🧠 User preferences (what each user asked me to remember)"):
        rows = []
        for u, r in users.items():
            pr = pe.prefs_of(r)
            rows.append({"ID": u, "Hidden columns": ", ".join(map(fmt, pr["hidden_columns"])) or "-",
                         "Only these columns": ", ".join(map(fmt, pr["preferred_columns"])) or "-",
                         "Always include": ", ".join(map(fmt, pr["always_columns"])) or "-",
                         "Raw data": "when needed" if pr["raw_data"] else "only when asked",
                         "Charts": "on" if pr["charts"] else "off"})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.caption("Preferences = what a user WANTS to see. Permissions above = what they are ALLOWED to see.")
    with st.expander("🧾 Policy audit log"):
        log = POLICY_STORE.read_log()
        if log:
            st.dataframe(pd.DataFrame(log).iloc[::-1], hide_index=True, use_container_width=True, height=300)
            st.download_button("⬇️ Download policy log", POLICY_STORE.audit_path.read_bytes(), "policy_audit.csv", "text/csv")
        else:
            st.caption("No policy changes yet.")


def my_prefs_panel():
    """Every user: their own saved data preferences (level 4)."""
    prefs = current_prefs()
    eff = current_eff()
    opts = [c for c in _col_options() if eff is None or eff.key_visible(c)
            or c in prefs["hidden_columns"] + prefs["preferred_columns"] + prefs["always_columns"]]
    with st.container(border=True, key="setcard_cols"):
        _set_head("🧠", "Which columns do you want?", "What <b>you want</b> to see. You can also just tell me in the "
                  "chat, e.g. <i>from now on never show UniqueNo</i>.")
        with st.form("my_prefs", border=False):
            hid = st.multiselect("🚫 Never show these columns", opts,
                                 default=[c for c in prefs["hidden_columns"] if c in opts], format_func=pe.label,
                                 placeholder="None - show everything normal",
                                 help="Removed from every answer, unless you name the column in your question.")
            only = st.multiselect("🎯 Only give me these columns (in this order)", opts,
                                  default=[c for c in prefs["preferred_columns"] if c in opts], format_func=pe.label,
                                  placeholder="Empty = normal columns",
                                  help="For work order / department / category-wise answers: just these, in this order.")
            alw = st.multiselect("➕ Always add these columns", opts,
                                 default=[c for c in prefs["always_columns"] if c in opts], format_func=pe.label,
                                 placeholder="None", help="Added to answers even when you don't ask for them.")
            c1, c2 = st.columns(2)
            raw = c1.toggle("👥 Employee-level rows when a question needs them", value=prefs["raw_data"],
                            help="Off = totals only, unless you ask for employee-wise data.")
            ch = c2.toggle("📊 Show charts", value=prefs["charts"])
            st.caption("Your administrator's data permissions always come first - a preference can't show a "
                       "restricted column.")
            submitted = st.form_submit_button("💾 Save my column preferences", type="primary")
        if submitted:
            new = dict(prefs, hidden_columns=hid, preferred_columns=only, always_columns=alw, raw_data=raw, charts=ch)
            _save_data_prefs(new)
            _POL.prefs = new
            audit(ME, "preference_changed", "settings page")
            st.success("✅ Saved for your login.")
    conflict = [c for c in prefs["preferred_columns"] + prefs["always_columns"] if eff is not None and not eff.key_visible(c)]
    if conflict:
        st.warning(f"🔒 {', '.join(map(pe.label, conflict))}: your administrator has restricted this field, "
                   "so your preference cannot override it.")


SQL_HL_JS = """<script type="text/plain" id="pbhl">
(function () {
  const P = window, D = document, VER = "3";
  if (P.__pbHLtimer && P.__pbHLver === VER) return;
  if (P.__pbHLtimer) { P.clearInterval(P.__pbHLtimer); D.querySelectorAll(".pb-sql-hl").forEach(e => e.remove()); }
  P.__pbHLver = VER;
  const SQL_RX = /^\\s*(select|with)\\b[\\s\\S]*\\bfrom\\s+(\\[?dbo\\]?\\.)?\\[?payrollall1?\\b/i;
  const isSql = v => v.startsWith("/") || SQL_RX.test(v);
  const KW = new Set(("select from where group by order having join inner left right full outer cross on as and or not "
    + "in is null like between case when then else end top distinct with union all asc desc over partition "
    + "exists any some percent ties offset fetch next rows only table").split(" "));
  const FN = new Set(("count sum avg min max round format datefromparts eomonth isnull nullif cast convert coalesce "
    + "year month day datediff dateadd getdate left right len ltrim rtrim upper lower replace substring charindex "
    + "row_number rank dense_rank lag lead abs floor ceiling string_agg concat iif try_cast").split(" "));
  const BAD = new Set(("delete drop truncate update insert merge alter create rename exec execute grant revoke deny "
    + "shutdown kill backup restore dbcc bulk into openrowset openquery opendatasource openxml waitfor reconfigure "
    + "declare set use go").split(" "));
  const esc = t => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  function paint(text) {
    const C = P.__pbHL, cols = new Set((C.cols || []).map(c => c.toLowerCase()));
    const sp = (c, t, extra) => `<span style="color:${c};${extra || ""}">${esc(t)}</span>`;
    const slash = text.startsWith("/");
    let out = slash ? sp(C.slash, "/") : "";
    const re = /(--[^\\n]*|\\/\\*[\\s\\S]*?(?:\\*\\/|$))|('(?:[^']|'')*'?)|(\\[[^\\]\\n]*\\]?)|(\\b\\d+(?:\\.\\d+)?\\b)|([A-Za-z_][A-Za-z0-9_]*)|(\\s+)|([^\\sA-Za-z0-9_'\\[]+)/g;
    let m, body = slash ? text.slice(1) : text;
    while ((m = re.exec(body)) !== null) {
      const [t, com, str, br, num, word, ws, op] = m;
      if (com || (op && /;|@/.test(op))) out += sp(C.bad, t, "text-decoration:underline wavy");
      else if (str) out += sp(C.str, t);
      else if (br) out += cols.has(br.replace(/[\\[\\]]/g, "").toLowerCase()) ? sp(C.col, t) : sp(C.text, t);
      else if (num) out += sp(C.num, t);
      else if (word) {
        const w = word.toLowerCase();
        if (BAD.has(w)) out += sp(C.bad, t, "text-decoration:underline wavy");
        else if (KW.has(w)) out += sp(C.kw, t);
        else if (FN.has(w) && (/^\\s*\\(/.test(body.slice(re.lastIndex)) || !cols.has(w))) out += sp(C.fn, t);
        else if (cols.has(w) || /^payrollall1?$/.test(w)) out += sp(/^payrollall1?$/.test(w) ? C.tbl : C.col, t);
        else out += sp(C.text, t);
      } else if (ws) out += esc(t);
      else out += sp(C.op, t);
    }
    return out + "\\n";
  }
  function sync(ta) {
    const C = P.__pbHL;
    const box = ta.parentElement;
    let ov = box.querySelector(":scope > .pb-sql-hl");          // ONE overlay per chat box (never an old one)
    if (!isSql(ta.value)) {
      if (ov) { ov.style.display = "none"; ov.innerHTML = ""; ov.__src = null; }
      ta.style.removeProperty("color"); ta.style.removeProperty("-webkit-text-fill-color");
      ta.spellcheck = true;
      return;
    }
    if (!ov) {
      ov = D.createElement("div"); ov.className = "pb-sql-hl";
      if (P.getComputedStyle(box).position === "static") box.style.position = "relative";
      box.insertBefore(ov, ta);
    }
    ta.spellcheck = false;                                         // no red spelling lines under column names
    const cs = P.getComputedStyle(ta);
    ["fontFamily", "fontSize", "fontWeight", "lineHeight", "letterSpacing", "paddingTop", "paddingRight",
     "paddingBottom", "paddingLeft", "borderTopWidth", "borderLeftWidth", "textIndent", "tabSize"]
      .forEach(k => ov.style[k] = cs[k]);
    Object.assign(ov.style, {display: "block", position: "absolute", pointerEvents: "none", margin: "0",
      top: ta.offsetTop + "px", left: ta.offsetLeft + "px", width: ta.clientWidth + "px", height: ta.clientHeight + "px",
      whiteSpace: "pre-wrap", overflowWrap: "break-word", wordBreak: "break-word", overflow: "hidden",
      boxSizing: cs.boxSizing, borderStyle: "solid", borderColor: "transparent", zIndex: "0"});
    ta.style.setProperty("color", "transparent", "important");
    ta.style.setProperty("-webkit-text-fill-color", "transparent", "important");
    ta.style.setProperty("caret-color", C.caret, "important");
    ta.style.position = "relative"; ta.style.zIndex = "1"; ta.style.background = "transparent";
    if (ov.__src !== ta.value) { ov.innerHTML = paint(ta.value); ov.__src = ta.value; }
    ov.scrollTop = ta.scrollTop;
  }
  function tick() {
    const tas = [...D.querySelectorAll('[data-testid="stChatInput"] textarea')];
    D.querySelectorAll(".pb-sql-hl").forEach(o => {                // remove overlays left from an old chat box
      if (!tas.some(t => t.parentElement === o.parentElement)) o.remove();
    });
    tas.forEach(ta => {
      if (!ta.__pbBound) {
        ta.__pbBound = true;
        ["input", "scroll", "keyup", "focus", "paste", "cut"].forEach(e => ta.addEventListener(e, () => sync(ta)));
      }
      sync(ta);
    });
  }
  P.__pbHLtimer = P.setInterval(tick, 250);
  tick();
})();
</script>
<script>
(function () {
  const P = window.parent, D = P.document;
  P.__pbHL = P.JSON.parse(__CFG_STR__);             // newest colours / columns (copied into the main page)
  if (P.__pbHLtimer && P.__pbHLver === "3") return; // installed once per browser tab (new version replaces old)
  const code = document.getElementById("pbhl").textContent;
  const s = D.createElement("script");              // run inside the main page, so it survives reruns
  s.textContent = code;
  D.head.appendChild(s);
})();
</script>"""


def chat_sql_highlighter():
    """Admin: colours a '/' SQL query while it is typed in the chat box (keywords, functions, columns, text,
    numbers; delete / drop / update... in red). Only changes how the text LOOKS - what is sent is unchanged."""
    try:
        import streamlit.components.v1 as components
    except Exception:
        return
    cfg = {"cols": [c for c, _ in _table_columns()],
           "slash": "#E65F26", "kw": "#5b8def" if DARK else "#1f3fbf", "fn": "#c678dd" if DARK else "#8a2be2",
           "col": "#4ec9b0" if DARK else "#0f7a6c", "tbl": "#e5c07b" if DARK else "#a35a00",
           "str": "#ce9178" if DARK else "#b3261e", "num": "#b5cea8" if DARK else "#0b7a2f",
           "op": TH["muted"], "text": TH["text"], "bad": "#ff4d4f", "caret": TH["text"]}
    components.html(SQL_HL_JS.replace("__CFG_STR__", json.dumps(json.dumps(cfg))), height=0)


def sql_editor():
    """Admin SQL editor: colours the query while typing (streamlit-ace if installed), Run = '/' + query."""
    with st.expander("🧾 SQL editor (admin) - or type / at the start of a chat message", expanded=False):
        try:
            from streamlit_ace import st_ace
        except Exception:
            st_ace = None
        if st_ace is not None:
            code = st_ace(placeholder=f"select Category, count(distinct Employee) from {TABLE} where [Year] = 2026 "
                                      "and [Month] = 8 group by Category", language="sql",
                          theme="tomorrow_night" if DARK else "sqlserver", key="sql_ace", height=170, font_size=14,
                          wrap=True, auto_update=True, show_gutter=True)
        else:
            code = st.text_area("SQL", key="sql_txt", height=140, label_visibility="collapsed",
                                placeholder=f"select Category, count(*) from {TABLE} group by Category")
            if code:
                st.code(code, language="sql")
            st.caption("Tip: run `pip install streamlit-ace` on this PC to get colours while you type.")
        if st.button("▶ Run SQL", key="sql_run", disabled=not str(code or "").strip()):
            st.session_state.pending = "/" + str(code).strip()
            st.rerun()


def due_reminders_banner():
    """Shows this login's due reminders at the top of the chat, with Done / Snooze buttons."""
    if REMINDER_STORE is None:
        return
    try:
        due = REMINDER_STORE.due(ME)
    except Exception:
        return
    if not due:
        return
    shown = st.session_state.setdefault("_toasted", set())
    for r in due:
        if r["id"] not in shown:
            shown.add(r["id"])
            try:
                st.toast(f"⏰ Reminder: {r['text']}")
            except Exception:
                pass
        with st.container(border=True, key=f"remind_{r['id']}"):
            c1, c2, c3, c4 = st.columns([6, 1.2, 1.2, 1.4])
            c1.markdown(f"⏰ **Reminder #{r['id']}:** {html.escape(r['text'])}  \n"
                        f"<span style='color:var(--muted);font-size:.82rem'>was due "
                        f"{html.escape(rmd.when_text(dt_parse(r['due'])))}</span>", unsafe_allow_html=True)
            if c2.button("✅ Done", key=f"rdone_{r['id']}", use_container_width=True):
                REMINDER_STORE.done(ME, r["id"])
                st.rerun()
            if c3.button("1 hour", key=f"rh_{r['id']}", use_container_width=True, help="Snooze 1 hour"):
                REMINDER_STORE.snooze(ME, r["id"], datetime.timedelta(hours=1))
                st.rerun()
            if c4.button("Tomorrow", key=f"rt_{r['id']}", use_container_width=True, help="Snooze to tomorrow 10 AM"):
                tmr = datetime.datetime.combine(datetime.date.today() + datetime.timedelta(days=1), datetime.time(10))
                REMINDER_STORE.snooze(ME, r["id"], tmr - datetime.datetime.now())
                st.rerun()


def page_chat():
    due_reminders_banner()
    history = st.session_state.history
    first = ME_REC.get("name", ME).split()[0] if ME_REC.get("name") else ""
    d_ok, a_ok = db_ok(), ai_ok()
    st.markdown(f"""<div class="lk-hero">
  <div class="txt"><span class="lk-hello">Hello{(' ' + html.escape(first)) if first else ''}! 👋</span>
    <h2>I'm your Payroll <span class="ai">AI Assistant</span></h2>
    <p>Ask me anything about salaries, statutory deductions, departments, work orders, CTC margins and
    payroll trends in plain English.</p>
    <div class="lk-chips">
      <span class="lk-chip {'g' if d_ok else 'r'}"><i></i>{'SQL Server Connected' if d_ok else 'SQL Server not reachable'}</span>
      <span class="lk-chip {'p' if a_ok else 'r'}">{'Ollama LLM' if a_ok else 'Ollama AI off'}</span>
      <span class="lk-chip o">🔒 Read-only · Data cannot be altered</span>
      <span class="lk-chip b">🖥 Runs on On-Prem Host</span>
      <span class="lk-chip b">{'🧠 AI on every question' if always_ai else '⚡ Fast mode'}</span>
    </div></div>
  <div class="say">How can I<br>help you today?</div>
  <div class="art">{ROBOT_ART}</div></div>""", unsafe_allow_html=True)

    lm = latest_label()
    dark_ = DARK
    cards = [
        ("₹", "rgba(34,197,94,.14)", "#22c55e", f"Give me salary for {lm}", "View executive payroll overview",
         f"give me salary for {lm.lower()}"),
        ("🏢", "rgba(59,130,246,.14)", "#3b82f6", f"Department wise salary for {prev_label(lm)}",
         "Breakdown by department", f"department wise salary for {prev_label(lm).lower()}"),
        ("🧾", "rgba(139,92,246,.16)", "#8b5cf6", f"PF, TDS and ESIC for {lm}", "Review statutory deductions",
         f"pf, tds and esic for {lm.lower()}"),
        ("📈", "rgba(230,95,38,.14)", "#E65F26", f"Salary trend from {month_back(lm, 4)} to {lm}",
         "Check salary trend and comparison", f"salary trend from {month_back(lm, 4).lower()} to {lm.lower()}"),
        ("👥", "rgba(13,148,136,.16)", "#14b8a6", f"List all employees salary for {lm}", "Complete employee salary list",
         f"list all employees salary for {lm.lower()}"),
        ("🏆", "rgba(59,130,246,.14)", "#3b82f6", f"Highest margin work orders in {lm}", "Base sales value, CTC and margin",
         f"which workorder has highest margin in {lm.lower()}"),
    ]
    if history or "pending" in st.session_state:
        cards = cards[:3]                                   # during a chat: one row of examples
    s1, s2 = st.columns([6, 1.2], vertical_alignment="center")
    s1.markdown("""<div class="lk-sec"><span class="t">What would you like to know?</span>
<span class="s">Click an example or type your own question below</span></div>""", unsafe_allow_html=True)
    if history and s2.button(":material/add_comment: New chat", use_container_width=True):
        st.session_state.history = []
        st.session_state.clarify_for = None
        st.session_state.last_q = None
        st.rerun()
    cols = st.columns(3, gap="medium")
    for n, (ic, bg, fg, title, sub, q) in enumerate(cards):
        with cols[n % 3]:
            with st.container(key=f"card_{n}"):
                st.markdown(f"""<div class="qcard"><div class="ic" style="background:{bg};color:{fg}">{ic}</div>
  <div><div class="tt">{title}</div><div class="ss">{sub}</div></div><div class="ch">›</div></div>""",
                            unsafe_allow_html=True)
                if st.button(title, key=f"qbtn_{n}"):
                    st.session_state.pending = q
                    st.rerun()
            st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
    if not history:
        st.markdown("""<div class="tipbar">✦ <b>Tip:</b> You can ask anything, e.g. <i>which workorder end date is
 greater than jan-28</i> or <i>average ctc per employee by category for aug-26</i>.
 Say <i>salary month</i>, <i>liable month</i> or <i>paid in oct-26</i> to choose the month column.</div>""",
                    unsafe_allow_html=True)

    start = max(0, len(history) - SHOW_LAST_MESSAGES)
    if start:
        st.caption(f"({start} older messages hidden to keep the page fast)")
    for i in range(start, len(history)):
        render(i, history[i])

    if IS_ADMIN:
        sql_editor()
    typed = st.chat_input("Ask e.g. give me salary for aug-26" + ("   ·   start with / to run your own SQL" if IS_ADMIN else ""))
    if IS_ADMIN:
        chat_sql_highlighter()                         # colours "/" SQL while typing
    question = typed or st.session_state.pop("pending", None)
    if question:
        history.append({"role": "user", "text": question})
        render(len(history) - 1, history[-1])
        audit(ME, "question", question)
        run_q = question
        prev = st.session_state.get("clarify_for")
        is_sql = is_user_sql(question) if IS_ADMIN else question.lstrip().startswith("/")
        if typed and prev and len(typed.split()) <= 4 and not is_sql:   # short reply to "what did you mean?"
            run_q = f"{prev} {typed}"
        elif typed and not is_sql:
            fu = follow_up(question, st.session_state.get("last_q"))   # "and jul-26?" / "what about IT?"
            if fu:
                run_q = fu
        st.session_state.clarify_for = None
        with st.spinner("🤖 Thinking..."):
            try:
                result = answer(run_q, default_mode, use_ai, always_ai)
                result = secure_result(result, run_q)        # final security boundary (permissions + preferences)
                if run_q != question and result.get("note") is not None and result["report"] != "clarify":
                    result["note"] = f"Read as: “{run_q}”. " + result.get("note", "")
                if result.get("report") == "clarify":
                    st.session_state.clarify_for = result.get("clarify_for")
                if result.get("report") not in NO_EXPLAIN and not is_sql:
                    st.session_state.last_q = result.get("_last_q") or run_q   # remembered for follow-ups
                result["question"] = question
                d = result.get("df")
                if result.get("report") not in NO_EXPLAIN and d is not None and not d.empty:
                    try:
                        result["explain"] = quick_explain(result)        # instant, no AI
                    except Exception:
                        result["explain"] = ""
                want_ai = result.get("report") == "chat_ai" or (
                    st.session_state.get("cfg_explain", CHAT_STYLE) and
                    (st.session_state.get("cfg_ai_explain", AI_EXPLAIN) == "always" or AI_EXPLAIN_ASK.search(question))
                    and result.get("report") not in NO_EXPLAIN and d is not None and not d.empty
                    and not (result.get("report") in ("summary", "employee_detail", "workorder_detail",
                                                      "employee_list", "joiners", "leavers")
                             and not AI_EXPLAIN_ASK.search(question)))
                if (use_ai or always_ai) and want_ai:
                    result["explain"] = ""
                    result["explain_pending"] = True
                history.append({"role": "assistant", **result})
            except Exception as e:
                history.append({"role": "assistant", "text": f"❌ Error: {e}",
                                 "note": "Check the CONFIG values at the top of the file."})
        render(len(history) - 1, history[-1])


def _hidden_for(uid):
    rec = load_users().get(uid, {})
    return set(rec.get("hidden_requests", [])), rec.get("requests_cleared_at", "")


def _hide_request(uid, q=None, all_=False):
    """Hides questions from MY list only - the admin activity log keeps everything (security)."""
    with _users_lock():
        _hide_request_locked(uid, q, all_)


def _hide_request_locked(uid, q=None, all_=False):
    users = load_users()
    if uid not in users:
        return
    if all_:
        users[uid]["requests_cleared_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        users[uid]["hidden_requests"] = []
    else:
        h = set(users[uid].get("hidden_requests", []))
        h.add(q.strip().lower())
        users[uid]["hidden_requests"] = sorted(h)[-500:]
    save_users(users)


def page_requests():
    st.markdown('<div class="pagetitle">My Requests</div>'
                '<div class="pagesub">Questions you have asked. Ask one again in a fresh chat, or delete it '
                'from this list.</div>', unsafe_allow_html=True)
    try:
        log = pd.read_csv(AUDIT_FILE)
        mine = log[(log["user"] == ME) & (log["action"] == "question")].iloc[::-1]
    except Exception:
        mine = pd.DataFrame()
    hidden, cleared = _hidden_for(ME)
    if not mine.empty and cleared:
        mine = mine[mine["time"].astype(str) > cleared]
    rows, seen = [], set()
    for _, r in mine.iterrows():
        q = str(r["detail"]).strip()
        k = q.lower()
        if k in seen or k in hidden:
            continue
        seen.add(k)
        rows.append((q, r["time"]))
        if len(rows) >= 50:
            break
    if not rows:
        st.info("No questions here yet. Go to **Chat** and ask something!")
        return
    _, cr = st.columns([6, 1.4])
    if st.session_state.get("confirm_del_all"):
        cr.warning("Delete all?")
        y, n_ = cr.columns(2)
        if y.button("Yes", key="del_all_yes", use_container_width=True):
            _hide_request(ME, all_=True)
            st.session_state.confirm_del_all = False
            st.rerun()
        if n_.button("No", key="del_all_no", use_container_width=True):
            st.session_state.confirm_del_all = False
            st.rerun()
    elif cr.button(":material/delete_sweep: Delete all", key="del_all", use_container_width=True):
        st.session_state.confirm_del_all = True
        st.rerun()
    for n, (q, when) in enumerate(rows):
        key = hashlib.md5(q.lower().encode()).hexdigest()[:10]
        c1, c2 = st.columns([6, 1.4])
        c1.markdown(f'<div class="pagecard" style="padding:12px 18px;margin-bottom:8px"><b>{q}</b>'
                    f'<div class="note">{when}</div></div>', unsafe_allow_html=True)
        if c2.button(":material/replay: Ask again", key=f"again_{key}", use_container_width=True):
            st.session_state.history = []            # fresh chat: only this question
            st.session_state.clarify_for = None
            st.session_state.pending = q
            st.session_state.page = "Chat"
            st.rerun()
        if c2.button(":material/delete: Delete", key=f"del_{key}", use_container_width=True):
            _hide_request(ME, q)
            st.rerun()


def page_documents():
    st.markdown('<div class="pagetitle">Documents</div>'
                '<div class="pagesub">Download the results from this chat session.</div>', unsafe_allow_html=True)
    history = st.session_state.history
    items = []
    for i, m in enumerate(history):
        df = _display(m)[1] if m.get("role") == "assistant" else None
        if m.get("role") == "assistant" and df is not None and not df.empty:
            q = history[i - 1]["text"] if i and history[i - 1].get("role") == "user" else m.get("report", "")
            items.append((i, q, m, df))
    if not items:
        st.info("No results yet. Answers with tables from **Chat** will appear here for download.")
        return
    for i, q, m, df in reversed(items):
        c1, c2 = st.columns([7, 2])
        c1.markdown(f'<div class="pagecard" style="padding:12px 18px;margin-bottom:8px"><b>{q}</b>'
                    f'<div class="note">{REPORT_TITLES.get(m["report"], m["report"])} · {len(df):,} row(s)</div></div>',
                    unsafe_allow_html=True)
        safe = re.sub(r"[^A-Za-z0-9]+", "_", q)[:40].strip("_") or "result"
        c2.download_button("⬇️ CSV", df.to_csv(index=False).encode("utf-8"), f"{safe}.csv", "text/csv",
                           key=f"doc_{i}", use_container_width=True)


FAQS = [
    ("What can I ask?",
     "Anything about the payroll data: monthly salary totals, department / category wise, PF · TDS · ESIC, salary "
     "breakup, employee details, work orders, base sales value and margin, trends, joiners, leavers, and custom "
     "questions like *which workorder end date is greater than jan-28*."),
    ("How do I write the month?",
     "Any of these work: *aug-26*, *Aug 2026*, *august'26*, *08/26*, *2026-08*. Ranges: *apr-26 to aug-26*. "
     "If you don't give a month, the latest month in the data is used."),
    ("What is 'base sales value'?",
     f"Invoice amount without GST: **Invoice × 100 / {100 + GST_RATE}**. The invoice is repeated on every employee row "
     "of a work order, so it is counted **once per work order + salary month**. **Margin = Base sales value − CTC**, "
     "**Margin % = Margin ÷ CTC × 100** (ask *margin % on base sales* for the sales basis)."),
    ("Salary month vs liable month vs payment month?",
     "**Salary month** = Month + Year columns (the month the salary is for). **Liable month** = month_year "
     "(payment date − 1 month). **Payment date** = month_year + 1 month (calculated, not stored in SQL). "
     "Pick the default in **Settings**, or say *salary month*, *liable month* or *paid in oct-26* in a question."),
    ("What does 'AI analysis' mean?",
     "For questions no fixed report covers, the AI writes a SQL query. It is checked (read-only, payroll table "
     "only, time limit) before it runs. Always look at the **SQL** tab for important numbers."),
    ("Can the chatbot change data?", "No. It can only read. Requests to add, change or delete anything are refused, "
     "every query is checked, and the database connection is read-only."),
    ("Who can see my questions?", "Your chat is private to your login. Administrators can see an activity log "
     "(who logged in and what was asked) for security."),
    ("Can I ask follow-up questions?",
     "Yes. After *salary for aug-26* you can type *and jul-26?* or *what about IT department?* - it reuses your "
     "previous question. Click **New chat** to start fresh."),
    ("Can it explain formulas?",
     "Yes. Ask e.g. *formula of margin*, *how is base sales calculated*, *what is month_year*, *explain CTC*. "
     "Add a month (e.g. *margin for aug-26*) to get the actual numbers instead."),
    ("Why is an answer slow?",
     "Normal questions (salary, department, deductions, work orders, top employees) answer in seconds. "
     "**AI analysis** questions need the AI to write a query, which can take a minute on a normal PC. "
     "While it is thinking, please wait - clicking elsewhere stops the question."),
    ("I forgot my password", "Ask the administrator to reset it in **Settings → Manage users**. You will choose a new "
     "one at your next login."),
]


def page_faqs():
    st.markdown('<div class="pagetitle">FAQs</div><div class="pagesub">Quick help for using the Payroll AI Assistant.'
                '</div>', unsafe_allow_html=True)
    for q, a in FAQS:
        with st.expander(q):
            st.markdown(a)


SETTINGS_CSS = """<style>
.set-hero {display:flex; align-items:center; gap:18px; padding:18px 22px; border-radius:18px; margin:6px 0 14px;
  background: linear-gradient(120deg, var(--hero1), var(--hero2)); border:1px solid var(--line);}
.set-hero .av {width:58px; height:58px; border-radius:50%; background:var(--navy); color:#fff; font-weight:800;
  font-size:1.35rem; display:flex; align-items:center; justify-content:center; flex:none;
  box-shadow:0 4px 14px rgba(11,42,111,.25);}
.set-hero .nm {font-size:1.25rem; font-weight:800; color:var(--heading);}
.set-hero .id {color:var(--muted); font-size:.88rem;}
.set-hero .badge {display:inline-block; margin-left:8px; padding:2px 10px; border-radius:999px; font-size:.72rem;
  font-weight:700; background:var(--orange); color:#fff; vertical-align:middle; text-transform:uppercase; letter-spacing:.04em;}
.set-chips {display:flex; flex-wrap:wrap; gap:8px; margin-top:10px;}
.set-chip {background:var(--surface); border:1px solid var(--line); color:var(--text); border-radius:999px;
  padding:4px 12px; font-size:.8rem; white-space:nowrap;}
.set-chip b {color:var(--heading);}
.set-h {display:flex; gap:12px; align-items:flex-start; margin:2px 0 10px;}
.set-h .ic {width:38px; height:38px; border-radius:12px; display:flex; align-items:center; justify-content:center;
  font-size:1.15rem; background:var(--tip); flex:none;}
.set-h .t {font-weight:800; color:var(--heading); font-size:1.02rem; line-height:1.2;}
.set-h .s {color:var(--muted); font-size:.85rem; margin-top:2px; line-height:1.35;}
.set-ex {background:var(--tip); border-radius:10px; padding:9px 13px; color:var(--text); font-size:.85rem; margin-top:6px;}
.set-ex b {color:var(--heading);}
.set-count {display:flex; gap:10px; margin:2px 0 10px;}
[class*="st-key-setcard"] {background:var(--surface) !important; border:1px solid var(--line) !important;
  border-radius:16px !important; padding:16px 18px !important; box-shadow:0 2px 10px rgba(11,42,111,.06);}
.set-count > div {flex:1; background:var(--surface2); border:1px solid var(--line); border-radius:12px; padding:8px 12px;}
.set-count .n {font-size:1.3rem; font-weight:800; color:var(--heading);} .set-count .l {font-size:.78rem; color:var(--muted);}
</style>"""


def _set_head(icon, title, sub):
    st.markdown(f'<div class="set-h"><div class="ic">{icon}</div><div><div class="t">{html.escape(title)}</div>'
                f'<div class="s">{sub}</div></div></div>', unsafe_allow_html=True)


AI_LEVELS = [("fast", "⚡ Fast", "Quick keyword rules; the AI helps only when a question is unclear (recommended)."),
             ("smart", "🧠 Smart", "The AI reads every question first. Understands odd wording better, but slower."),
             ("off", "🚫 Off", "Keyword rules only - fastest, but unusual questions may not be understood.")]


def _ai_level():
    if st.session_state.cfg_always_ai:
        return "smart"
    return "fast" if st.session_state.cfg_use_ai else "off"


def page_settings():
    st.markdown(SETTINGS_CSS, unsafe_allow_html=True)
    name = ME_REC.get("name") or ME
    initials = "".join(w[0] for w in str(name).split()[:2]).upper() or "?"
    role = ME_REC.get("role", "user")
    try:
        cats_all = sorted(distinct_values("Category"))
    except Exception:
        cats_all = []
    off_now = [c for c in st.session_state.cat_off if c in cats_all]
    hidden = st.session_state.hide_cols
    chips = [("📅 Months", {"month_year": "liable month", "payment": "payment month"}.get(st.session_state.cfg_mode,
                                                                                          "salary month")),
             ("🤖 AI", dict((k, l.split(" ", 1)[1]) for k, l, _ in AI_LEVELS)[_ai_level()]),
             ("✍️ Explanations", ("always" if st.session_state.cfg_ai_explain == "always" else "when I ask")
              if st.session_state.cfg_explain else "off"),
             ("🙈 Hidden", ", ".join(hidden) if hidden else "nothing"),
             ("🏷️ Categories", "all" if not off_now else f"{len(cats_all) - len(off_now)} of {len(cats_all)}"),
             ("🎨 Theme", "dark" if DARK else "light")]
    st.markdown(f"""<div class="set-hero"><div class="av">{html.escape(initials)}</div><div style="flex:1">
<div class="nm">{html.escape(str(name))}<span class="badge">{html.escape(role)}</span></div>
<div class="id">Login ID: {html.escape(ME)} · these settings are saved for your login only</div>
<div class="set-chips">{''.join(f'<span class="set-chip">{a}: <b>{html.escape(str(b))}</b></span>' for a, b in chips)}</div>
</div></div>""", unsafe_allow_html=True)

    names = ["💬 Answers", "📋 My columns", "🏷️ Categories", "🔑 Account"] + (["🛠️ Administration"] if IS_ADMIN else [])
    tabs = st.tabs(names)

    # ---------------- Answers ----------------
    with tabs[0]:
        a1, a2 = st.columns(2, gap="large")
        with a1:
            with st.container(border=True, key="setcard_1"):
                _set_head("📅", "Which month should I use?", "When you ask for <b>Aug-26</b>, which rows should count?")
                opts = {"month_year": "Liable month (month_year)", "salary": "Salary month (Month / Year columns)",
                        "payment": "Payment month (month_year + 1 month)"}
                if st.session_state.get("set_mode") not in opts.values():      # old label from before v15
                    st.session_state.pop("set_mode", None)
                st.session_state.setdefault("set_mode", opts.get(st.session_state.cfg_mode, opts["month_year"]))
                v = st.radio("Month basis", list(opts.values()), key="set_mode", label_visibility="collapsed")
                st.session_state.cfg_mode = next(k for k, t in opts.items() if t == v)
                if st.session_state.cfg_mode == "month_year":
                    st.markdown('<div class="set-ex"><b>Aug-26</b> = everything <b>liable in August</b> '
                                '(payment date − 1 month, e.g. paid 07-Sep-26 → 07-Aug-26). It can include arrears of '
                                'older salary months.</div>', unsafe_allow_html=True)
                elif st.session_state.cfg_mode == "payment":
                    st.markdown('<div class="set-ex"><b>Aug-26</b> = everything <b>paid in August</b> '
                                '(= liable month Jul-26).</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="set-ex"><b>Aug-26</b> = salary <b>for</b> August (Month = 8, Year = '
                                '2026), even if it was paid or billed later.</div>', unsafe_allow_html=True)
                st.caption("One question only? Type *liable month*, *salary month* or *paid in* in it.")
            with st.container(border=True, key="setcard_2"):
                _set_head("🤖", "How much AI?", "The AI understands your words. More AI = smarter, but slower.")
                st.session_state.setdefault("set_ai_level", _ai_level())
                lv = st.radio("AI level", [k for k, _, _ in AI_LEVELS], key="set_ai_level",
                              format_func=lambda k: dict((a, b) for a, b, _ in AI_LEVELS)[k], horizontal=True,
                              label_visibility="collapsed")
                st.session_state.cfg_always_ai = lv == "smart"
                st.session_state.cfg_use_ai = lv != "off"
                st.markdown(f'<div class="set-ex">{dict((a, c) for a, _, c in AI_LEVELS)[lv]}</div>',
                            unsafe_allow_html=True)
        with a2:
            with st.container(border=True, key="setcard_3"):
                _set_head("✍️", "Written explanation", "A short answer in plain words above the table.")
                st.session_state.cfg_explain = st.toggle("Explain answers in plain words",
                                                         value=st.session_state.cfg_explain)
                if st.session_state.cfg_explain:
                    ch = ["Only when I ask (fast)", "Always let the AI explain (slow)"]
                    st.session_state.setdefault("set_explain", ch[1] if st.session_state.cfg_ai_explain == "always" else ch[0])
                    v = st.radio("When", ch, key="set_explain", horizontal=True, label_visibility="collapsed")
                    st.session_state.cfg_ai_explain = "always" if v.startswith("Always") else "ask"
                    st.caption("“When I ask” = add *explain*, *why* or *insights* to your question.")
            with st.container(border=True, key="setcard_4"):
                _set_head("🙈", "Hide these amounts", "Left out of answers <b>unless your question asks for them</b>.")
                hide = st.multiselect("Hidden amounts", DEFAULT_HIDE,
                                      default=[g for g in DEFAULT_HIDE if g in st.session_state.hide_cols],
                                      label_visibility="collapsed", placeholder="Nothing hidden")
                if hide != st.session_state.hide_cols:
                    st.session_state.hide_cols = hide
                    _save_hide(hide)
                st.markdown('<div class="set-ex">Example: with <b>PF</b> hidden, <i>salary for aug-26</i> shows no PF, '
                            'but <i>pf for aug-26</i> still shows it.</div>', unsafe_allow_html=True)
        save_answer_prefs()

    # ---------------- My columns ----------------
    with tabs[1]:
        my_prefs_panel()

    # ---------------- Categories ----------------
    with tabs[2]:
        with st.container(border=True, key="setcard_5"):
            _set_head("🏷️", "Categories in my answers", "Switch a category <b>off</b> to leave it out of every answer.")
            if not cats_all:
                st.info("Category names are not available right now (database not reachable).")
            else:
                inc = len(cats_all) - len(off_now)
                st.markdown(f'<div class="set-count"><div><div class="n">{inc}</div><div class="l">included</div></div>'
                            f'<div><div class="n">{len(off_now)}</div><div class="l">left out</div></div>'
                            f'<div><div class="n">{len(cats_all)}</div><div class="l">categories in data</div></div></div>',
                            unsafe_allow_html=True)
                all_on = st.toggle("Include all categories", value=not off_now, key="cat_all")
                if all_on:
                    new_off = []
                else:
                    cols = st.columns(3)
                    new_off = [c for n, c in enumerate(cats_all)
                               if not cols[n % 3].toggle(c, value=c not in off_now, key=f"cat_{c}")]
                    if len(new_off) == len(cats_all):
                        st.warning("All categories are off - answers will be empty. Turn at least one on.")
                if new_off != off_now:
                    st.session_state.cat_off = new_off
                    _save_cat_off(new_off)
                st.markdown('<div class="set-ex">Typing <i>without efirm</i> leaves a category out for one question. '
                            'Naming one (<i>efirm margin for aug-26</i>) shows it even if it is switched off here.</div>',
                            unsafe_allow_html=True)

    # ---------------- Account ----------------
    with tabs[3]:
        b1, b2 = st.columns(2, gap="large")
        with b1:
            with st.container(border=True, key="setcard_6"):
                _set_head("🔑", "Change password", "At least 8 characters, with letters and numbers. Don't share it.")
                change_password_form(ME)
        with b2:
            with st.container(border=True, key="setcard_7"):
                _set_head("👤", "My account", "Who you are signed in as.")
                st.markdown(f"- **Name:** {name}\n- **Login ID:** `{ME}`\n- **Role:** {role}"
                            + (f"\n- **Auto sign-out:** after {SESSION_MINUTES} idle minutes" if SESSION_MINUTES else ""))
            with st.container(border=True, key="setcard_8"):
                _set_head("🎨", "Light / dark look", "Use the switch at the bottom of the left menu.")
                st.caption(f"Now: **{'Dark' if DARK else 'Light'}** · saved for your login.")

    # ---------------- Administration ----------------
    if IS_ADMIN:
        with tabs[4]:
            st.markdown('<div class="set-ex">💡 You can also manage permissions in the chat, e.g. <i>block GrossSalary '
                        'for everyone</i> or <i>give teaching right to user1</i> - you will be asked to confirm.</div>',
                        unsafe_allow_html=True)
            admin_panel(ME)
            st.caption(f"Model: `{MODEL}` · DB: `{DATABASE}` · Table: `{TABLE}` · AI analyst: "
                       f"{'on' if ANALYST_MODE else 'off'}")


{"Chat": page_chat, "My Requests": page_requests, "Documents": page_documents,
 "FAQs": page_faqs, "Settings": page_settings}.get(PAGE, page_chat)()

st.markdown('<div class="lk-foot"><span>Workorder Dummy · Internal use only · Read-only access to '
            'payroll data</span><span class="sec">● 100% On-Premise Secure</span></div>', unsafe_allow_html=True)
