"""
Payroll AI Assistant - web server for the "Blue Panther" UI (web/index.html).

Same brain as the Streamlit app: this file loads the logic of payroll_chatbot.py (questions -> read-only SQL,
logins in users.json, data permissions, audit_log.csv) and serves it as a small JSON API + the HTML page.
Nothing in SQL Server is changed, except by the admin-only "Refresh database" button (same as before).

Run:   python web_server.py          ->  http://localhost:8080   (other PCs: http://<this-PC-IP>:8080)
"""
import ast
import asyncio
import csv
import datetime
import decimal
import json
import logging
import math
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import warnings
from pathlib import Path

import pandas as pd
import uvicorn
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse, Response
from starlette.routing import Route

# ======================= CONFIG =======================
HOST = "0.0.0.0"            # listen on the office network (other PCs + your laptop on the same network/VPN)
PORT = 8080
SESSION_IDLE_MIN = 30       # signed out after this many minutes without using the page
SESSION_MAX_HOURS = 10      # and always after this many hours, however active (sign in again each working day)
COOKIE = "pb_session"
# ======================================================

APP_DIR = Path(__file__).resolve().parent
WEB_DIR = APP_DIR / "web"
warnings.filterwarnings("ignore")
logging.getLogger("streamlit").setLevel(logging.ERROR)
for _n in ("streamlit.runtime.caching.cache_data_api", "streamlit.runtime.scriptrunner_utils.script_run_context",
           "streamlit.runtime.state.session_state_proxy"):
    logging.getLogger(_n).setLevel(logging.CRITICAL)


# ---------------------------------------------------------------------------------------------------------
# 1) Load the chatbot logic from payroll_chatbot.py (everything before the Streamlit page, plus the login /
#    permission / table helpers after it). One source of truth: fixing an answer there fixes it here too.
# ---------------------------------------------------------------------------------------------------------
_EXTRA = {"inr_short", "APP_DIR", "USERS_FILE", "AUDIT_FILE", "ID_RULE", "_hash", "load_users", "_users_lock",
          "save_users", "update_user", "set_password", "check_password", "pw_problem", "temp_password",
          "REFRESH_FILE", "_refresh_status", "_save_refresh_status", "audit", "_login_fails", "is_locked",
          "note_fail", "USER_PREFS", "kpis", "view_df", "full_csv", "secure_result", "_display"}


def _load_bot():
    sys.path.insert(0, str(APP_DIR))
    src_path = APP_DIR / "payroll_chatbot.py"
    tree = ast.parse(src_path.read_text("utf-8"), str(src_path))
    cut = next(n.lineno for n in tree.body if isinstance(n, ast.Expr) and isinstance(n.value, ast.Call)
               and ast.unparse(n.value.func) == "st.set_page_config")

    def names(n):
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            return {n.name}
        if isinstance(n, ast.Assign):
            return {t.id for t in n.targets if isinstance(t, ast.Name)}
        return set()
    body = [n for n in tree.body if n.lineno < cut]
    body += [n for n in tree.body if n.lineno > cut and (isinstance(n, (ast.Import, ast.ImportFrom)) or names(n) & _EXTRA)]
    ns = {"__name__": "payroll_bot", "__file__": str(src_path)}
    exec(compile(ast.Module(body=body, type_ignores=[]), str(src_path), "exec"), ns)
    missing = _EXTRA - set(ns)
    if missing:
        raise RuntimeError(f"payroll_chatbot.py is missing: {sorted(missing)}")
    return ns


B = _load_bot()
pe, tg, rmd = B["pe"], B["tg"], B["rmd"]
import display_rules as dr                                # column names / formats set in chat by the admin
RULES = dr.DisplayRules(APP_DIR)
import workorder_groups as wg                            # "treat work orders starting M24 as NORTH" (master / admin)
WO_GROUPS = wg.WorkOrderGroups(APP_DIR)
CODE_COLUMNS = {"employee", "employee code", "employeecode", "emp code", "uniqueno", "uniqueno2", "invoiceno",
                "finaluniqcode", "workorderno", "unique no", "invoice no"}     # codes are never numbers with commas
_TL = threading.local()       # per request: the signed-in user's id and chat state


class _SessionState:
    """Stands in for st.session_state: each signed-in browser session has its own dict (thread-local per request)."""
    def _d(self):
        d = getattr(_TL, "state", None)
        if d is None:
            d = _TL.state = {}
        return d

    def __getattr__(self, k):
        try:
            return self._d()[k]
        except KeyError:
            raise AttributeError(k)

    def __setattr__(self, k, v):
        self._d()[k] = v

    def __getitem__(self, k):
        return self._d()[k]

    def __setitem__(self, k, v):
        self._d()[k] = v

    def __delitem__(self, k):
        del self._d()[k]

    def __contains__(self, k):
        return k in self._d()

    def get(self, k, default=None):
        return self._d().get(k, default)

    def setdefault(self, k, default=None):
        return self._d().setdefault(k, default)

    def pop(self, k, *default):
        return self._d().pop(k, *default)

    def keys(self):
        return self._d().keys()


class _StProxy:
    """Real streamlit for caching etc., but session_state is per user."""
    session_state = _SessionState()

    def __init__(self, real):
        object.__setattr__(self, "_real", real)

    def __getattr__(self, k):
        return getattr(self._real, k)


B["st"] = _StProxy(B["st"])
B["POLICY_STORE"] = pe.PolicyStore(APP_DIR)
B["TEACH_STORE"] = tg.TeachingStore(APP_DIR)
B["REMINDER_STORE"] = rmd.ReminderStore(APP_DIR)
B["USERS_LOADER"] = B["load_users"]
B["PREF_SAVER"] = lambda hide: B["update_user"](_TL.uid, hide_cols=list(hide))           # always the asking user
B["PREFS_SAVER"] = lambda new: B["update_user"](_TL.uid, data_prefs=new)
B["CAT_OFF_SAVER"] = lambda off: B["update_user"](_TL.uid, cat_off=list(off))
audit = B["audit"]

# ---------------------------------------------------------------------------------------------------------
# 2) Sessions (in memory: a server restart signs everyone out)
# ---------------------------------------------------------------------------------------------------------
SESSIONS = {}
_S_LOCK = threading.Lock()
_REFRESH_LOCK = threading.Lock()
MAX_ANSWERS = 60              # answers kept per session (for CSV downloads / insights)


def _new_state(rec):
    s = {}
    for key, (name, default) in B["USER_PREFS"].items():
        v = rec.get(name, default)
        s[key] = list(v) if isinstance(v, list) else v
    return s


def _client_ip(request):
    return request.client.host if request.client else "?"


def _session(request):
    tok = request.cookies.get(COOKIE)
    if not tok:
        return None
    with _S_LOCK:
        s = SESSIONS.get(tok)
        if not s:
            return None
        if time.time() - s["last"] > SESSION_IDLE_MIN * 60 or time.time() - s.get("born", s["last"]) > SESSION_MAX_HOURS * 3600:
            SESSIONS.pop(tok, None)
            audit(s["uid"], "logout", "signed out automatically (idle / time limit)")
            return None
        s["last"] = time.time()
    rec = B["load_users"]().get(s["uid"])
    if not rec or not rec.get("active", True):          # removed / disabled while signed in
        with _S_LOCK:
            SESSIONS.pop(tok, None)
        return None
    s["rec"] = rec
    return s


class _UserContext:
    """Everything the chatbot logic needs to know about the person asking - for THIS request only."""
    def __init__(self, s, job=None):
        self.s, self.job = s, job

    def __enter__(self):
        _TL.uid, _TL.state, _TL.job = self.s["uid"], self.s["state"], self.job
        B["set_policy_context"](self.s["uid"], self.s["rec"].get("role"), self.s["rec"])
        return self.s

    def __exit__(self, *a):
        _TL.uid, _TL.state, _TL.job = None, None, None


# ---------------------------------------------------------------------------------------------------------
# Stop button: each question (and its AI write-up) runs as a "job" the browser can stop by its id (rid).
# Stopping closes the job's own Ollama connection (Ollama then stops generating), cancels its SQL cursors,
# and the answer that may still come back is thrown away.
# ---------------------------------------------------------------------------------------------------------
class _Stopped(Exception):
    """The person pressed Stop: whatever the question was still doing is abandoned."""


class _Job:
    def __init__(self, s, rid, q, stage="question"):
        self.s, self.rid, self.q, self.stage, self.t0 = s, rid, q, stage, time.time()
        self.cancelled, self.cursors, self.ai = False, [], None
        self.lock = threading.Lock()

    def check(self):
        if self.cancelled:
            raise _Stopped()

    def stop(self):
        with self.lock:
            self.cancelled = True
            cursors, ai = list(self.cursors), self.ai
        for cur in cursors:
            try:
                cur.cancel()                   # SQL Server stops the running query
            except Exception:
                pass
        if ai is not None:
            try:
                ai._client.close()             # the waiting AI call fails at once; Ollama drops the request
            except Exception:
                pass


JOBS = {}                     # rid -> running job (or one stopped before the server got to it)
_J_LOCK = threading.Lock()


def _job_start(s, rid, q, stage="question"):
    """The job for this request, or None if Stop was pressed before it got here."""
    with _J_LOCK:
        job = JOBS.get(rid)
        if job is not None and job.s is s and job.cancelled:
            JOBS.pop(rid, None)
            audit(s["uid"], "question_stopped", f"{q[:250]} (before it started)")
            return None
        job = JOBS[rid] = _Job(s, rid, q, stage)
    return job


def _job_end(job):
    with _J_LOCK:
        if JOBS.get(job.rid) is job:
            JOBS.pop(job.rid, None)


_OLLAMA = B["ollama"]


class _OllamaProxy:
    """B["ollama"] for the chatbot logic: inside a job every AI call uses the job's own client, so Stop can close it."""
    def __getattr__(self, k):
        job = getattr(_TL, "job", None)
        if job is None:
            return getattr(_OLLAMA, k)
        with job.lock:
            job.check()
            if job.ai is None:
                job.ai = B["_ollama_lib"].Client(host=str(_OLLAMA._client.base_url), timeout=B["AI_TIMEOUT_SEC"])
            return getattr(job.ai, k)


class _JobConn:
    """A SQL connection whose cursors Stop can cancel."""
    def __init__(self, conn, job):
        self._conn, self._job = conn, job

    def cursor(self):
        cur = self._conn.cursor()
        with self._job.lock:
            self._job.cursors.append(cur)
        self._job.check()
        return cur

    def __getattr__(self, k):
        return getattr(self._conn, k)


_GET_CONN = B["get_conn"]


def _get_conn():
    job = getattr(_TL, "job", None)
    if job is None:
        return _GET_CONN()
    job.check()
    return _JobConn(_GET_CONN(), job)


B["ollama"] = _OllamaProxy()
B["get_conn"] = _get_conn


def _json(data, status=200):
    return JSONResponse(data, status_code=status, headers={"Cache-Control": "no-store"})


def _err(msg, status=400):
    return _json({"error": msg}, status)


ROLE_RANK = {"user": 0, "admin": 1, "master": 2}    # master: everything · admin: all data + add users · user


def role_of(s):
    return s["rec"].get("role", "user") if s else "user"


def need_login(admin=False, master=False):
    import inspect
    from starlette.concurrency import run_in_threadpool

    def deco(fn):
        async def wrapper(request):
            s = _session(request)
            if not s:
                return _err("Please sign in.", 401)
            if s["rec"].get("must_change") and fn.__name__ not in ("api_password", "api_me", "api_logout"):
                return _err("Please set your own password first.", 403)
            if master and role_of(s) != "master":
                return _err("Only the master administrator can do this.", 403)
            if admin and ROLE_RANK.get(role_of(s), 0) < 1:
                return _err("Administrators only.", 403)
            if inspect.iscoroutinefunction(fn):
                return await fn(request, s)
            return await run_in_threadpool(fn, request, s)       # slow work (SQL / AI) off the event loop
        wrapper.__name__ = fn.__name__
        return wrapper
    return deco


async def _body(request):
    try:
        return await request.json()
    except Exception:
        return {}


# ---------------------------------------------------------------------------------------------------------
# 3) Turning a result (pandas table) into JSON for the page
# ---------------------------------------------------------------------------------------------------------
NICE = {"BaseSalesValue": "Base sales value", "MarginPct": "Margin %", "TakeHome": "Take-home", "GrossAmount": "Gross",
        "IncomeTax": "TDS", "TotalPF": "Total PF", "InvoiceWithGST": "Invoice (with GST)", "Employees": "Employees"}


def _cell(v):
    if v is None or v is pd.NaT:                         # blank date
        return None
    if isinstance(v, decimal.Decimal):
        v = float(v)
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        return v.strftime("%d-%b-%Y")
    if isinstance(v, datetime.date):
        return v.strftime("%d-%b-%Y")
    if hasattr(v, "item"):                               # numpy number
        v = v.item()
    if isinstance(v, float):
        return None if math.isnan(v) or math.isinf(v) else round(v, 2)
    if isinstance(v, (int, bool, str)):
        return v
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return str(v)


def _kind(name, series):
    n = str(name)
    if re.search(r"pct|%|percent", n, re.I):
        return "pct"
    if pd.api.types.is_numeric_dtype(series):
        if B["MONEY_COL"].search(n) and not re.search(r"count|employees|days|month|year|^employee$|work ?orders|rows",
                                                       n, re.I):
            return "money"
        return "num"
    return "text"


def _table(df, intent, limit=1000):
    v = B["view_df"](df, intent)
    names = {"MarginPct": f"Margin % (on {'base sales' if (intent or {}).get('pct_basis') == 'sales' else 'CTC'})",
             "BaseSalesValue": "Base sales value", "TakeHome": "Take-home", "GrossAmount": "Gross",
             "IncomeTax": "TDS", "TotalPF": "Total PF", "InvoiceWithGST": "Invoice (with GST)"}
    rules = RULES.load()
    text_cols = {c.lower() for c in rules["text_columns"]} | CODE_COLUMNS
    rename = {k.lower(): n for k, n in rules["rename"].items()}
    kinds = ["text" if str(c).lower() in text_cols else _kind(c, v[c]) for c in v.columns]
    cols = [rename.get(str(c).lower()) or names.get(str(c), str(c)) for c in v.columns]
    rows = [[(_plain(x) if k == "text" else _cell(x)) for x, k in zip(r, kinds)]
            for r in v.head(limit).itertuples(index=False, name=None)]
    return {"cols": cols, "kinds": kinds, "rows": rows, "total": int(len(v))}


def _plain(v):
    """A code shown as text: 104576 not 1,04,645 / 104576.0."""
    v = _cell(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return None if v is None else str(v)


def _rename_csv(data):
    """Applies the display rules (names, plain codes) to a full CSV download as well."""
    import io
    rules = RULES.load()
    rename = {k.lower(): n for k, n in rules["rename"].items()}
    text = data.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return data
    text_idx = [i for i, c in enumerate(rows[0]) if c.lower() in ({c.lower() for c in rules["text_columns"]} | CODE_COLUMNS)]
    rows[0] = [rename.get(c.lower(), c) for c in rows[0]]
    for r in rows[1:]:
        for i in text_idx:
            if i < len(r) and re.fullmatch(r"\d+\.0+", r[i] or ""):
                r[i] = r[i].split(".")[0]
    out = io.StringIO()
    csv.writer(out, lineterminator="\r\n").writerows(rows)
    return out.getvalue().encode("utf-8-sig")


def _num(df, c):
    return pd.to_numeric(df[c], errors="coerce").fillna(0)


def _bar(labels, values, title, signed=False, color=None):
    return {"type": "bar", "labels": [str(x) for x in labels], "values": [_cell(float(x)) for x in values],
            "title": title, "signed": signed, "color": color}


def _line(df, x, cols):
    return {"type": "line", "labels": [str(v) for v in df[x]],
            "series": [{"name": NICE.get(c, c), "values": [_cell(float(v)) for v in _num(df, c)]} for c in cols]}


def chart_spec(report, df, intent):
    """Same choices as draw_chart() in the Streamlit app, as data for the page to draw."""
    intent = intent or {}
    if df is None or df.empty:
        return None
    try:
        fc, fl = intent.get("focus") or "CTC", intent.get("focus_label") or "CTC"
        asc = intent.get("sort") == "asc"
        if report == "plan":
            plan = intent.get("plan") or {}
            dim = next((("Name" if g_ == "Employee" else g_) for g_ in plan.get("group", []) if g_ != "Month"), None)
            srt = plan.get("sort")
            names = {"Employees": "Employee count", "InvoiceAmount": "Invoice amount", "AvgCTC": "Average CTC"}
            cand = ([names.get(srt[0], srt[0])] if srt else []) + ["Margin", "CTC", "BaseSalesValue", "Average CTC",
                                                                    "Employee count"]
            met = next((c for c in cand if c in df and pd.api.types.is_numeric_dtype(df[c])), None)
            if dim and dim in df and met:
                g = df.groupby(dim)[met].sum().sort_values(ascending=False).head(25)
                return _bar(g.index, g.values, NICE.get(met, met), signed=met == "Margin")
            mc = intent.get("month_col")
            if not dim and mc in df and df[mc].nunique() > 1:
                cols = [c for c in ("CTC", "BaseSalesValue", "Margin") if c in df] or ([met] if met else [])
                return _line(df, mc, cols) if cols else None
            return None
        if report == "share" and "Share %" in df:
            d = df.head(15)
            return _bar(d[d.columns[0]], d["Share %"], "Share of total %")
        if report == "expired_wo" and "CTC paid in month" in df and _num(df, "CTC paid in month").sum():
            d = df.sort_values("CTC paid in month", ascending=False).head(20)
            return _bar(d["WorkOrderNo"], _num(d, "CTC paid in month"), "CTC paid in the month (₹)")
        if report == "new_wo" and "CTC" in df and len(df):
            d = df.sort_values("CTC", ascending=False).head(20)
            return _bar(d["WorkOrderNo"], _num(d, "CTC"), "CTC this month (₹)")
        if report == "bands":
            return _bar(df["CTC band"], df["Employees"], "Employees")
        if report == "wo_end" and "WOEndDate" in df and len(df):
            ctc = [c for c in df.columns if str(c).startswith("CTC (")]
            if ctc:
                d = df.head(20)
                return _bar(d["WorkOrderNo"].astype(str) + " · " + pd.to_datetime(d["WOEndDate"]).dt.strftime("%d-%b-%y"),
                            _num(d, ctc[0]), str(ctc[0]) + " (₹)")
        if report in ("department", "category") and "Period" in df and df["Period"].nunique() > 1:
            col = report.capitalize()                        # trend by category / department: one line each
            top = df.groupby(col)["CTC"].sum().sort_values(ascending=False).head(6).index if "CTC" in df else []
            if len(top):
                piv = df[df[col].isin(top)].pivot_table(index="Period", columns=col, values="CTC", aggfunc="sum", sort=False)
                piv = piv.reindex(list(dict.fromkeys(df["Period"])))
                return {"type": "line", "labels": [str(x) for x in piv.index],
                        "series": [{"name": str(c), "values": [_cell(float(v)) if pd.notna(v) else None for v in piv[c]]}
                                   for c in piv.columns]}
        if report in ("summary", "trend") and "Period" in df and "CTC" in df:
            cols = [c for c in ("CTC", "BaseSalesValue", "Margin") if c in df]
            if df["Period"].nunique() > 1:
                return _line(df, "Period", cols)
            r = df.iloc[0]
            return _bar([NICE.get(c, c) for c in cols], [r[c] for c in cols], "Amount (₹)", signed=True)
        if report == "analysis" and 2 <= len(df) <= 60:
            num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
            txt = [c for c in df.columns if c not in num]
            if num and txt:
                d = df[[txt[0], num[0]]].dropna()
                return _bar(d[txt[0]], d[num[0]], str(num[0]), signed=True)
        if report == "employee_ranking" and "Name" in df and len(df) >= 2:
            m = df.columns[-1]
            return _bar(df["Name"], _num(df, m), str(m))
        if report == "workorder_summary" and "Margin" in df:
            d = df.head(15)
            m = fc if fc in d and fc != "Employees" else "Margin"
            return _bar(d["WorkOrderNo"], _num(d, m), (fl if m == fc else "Margin") + ("" if m == "MarginPct" else " (₹)"),
                        signed=True)
        if report == "workorder_monthly" and "Margin" in df:
            g = df.groupby("WorkOrderNo")["Margin"].sum().sort_values(ascending=asc).head(15)
            return _bar(g.index, g.values, "Margin (₹)", signed=True)
        if report in ("department", "category"):
            col = report.capitalize()
            g = df.groupby(col).sum(numeric_only=True)
            if fc == "MarginPct" and {"Margin", "CTC"} <= set(g.columns):
                den = g["BaseSalesValue"] if intent.get("pct_basis") == "sales" else g["CTC"]
                data = (g["Margin"] * 100 / den.replace(0, float("nan"))).dropna()
                title = "Margin %"
            else:
                use = fc if fc in g else "CTC"
                data, title = g[use], NICE.get(use, use)
            data = data.sort_values(ascending=asc).head(25)
            return _bar(data.index, data.values, title, signed=True)
        if report in ("deductions", "components"):
            skip = {"Employees", "Period"}
            sums = df[[c for c in df.columns if c not in skip]].select_dtypes("number").sum()
            sums = sums[sums != 0].sort_values(ascending=False)
            if not sums.empty:
                return _bar([NICE.get(k, k) for k in sums.index], sums.values, "Amount (₹)", color="purple")
        if report in ("joiners", "leavers") and "Department" in df:
            vc = df["Department"].fillna("(blank)").value_counts().head(25)
            return _bar(vc.index, vc.values, "People")
        if report == "projection":
            pm = [c for c in df.columns if str(c).startswith("Projected margin ")]
            if pm and "WorkOrderNo" in df:
                d = df.assign(_m=_num(df, pm[0])).sort_values("_m", ascending=False)
                d = d[d["_m"] != 0].head(15)
                if not d.empty:
                    return _bar(d["WorkOrderNo"], d["_m"], str(pm[0]) + " (₹)", signed=True)
        if report in ("proj_vs_actual", "proj_compare") and "CTC difference" in df:
            key = "WorkOrderNo" if "WorkOrderNo" in df else df.columns[0]
            d = df.assign(_d=_num(df, "CTC difference"))
            d = d[d["_d"] != 0].reindex(d["_d"].abs().sort_values(ascending=False).index).head(15)
            if not d.empty:
                months = (intent or {}).get("months") or ["first", "second"]
                title = ("Actual − projected CTC (₹)" if report == "proj_vs_actual"
                         else f"Projected CTC {months[-1]} − {months[0]} (₹)")
                return _bar(d[key], d["_d"], title, signed=True)
    except Exception:
        return None
    return None


def _md_strip(t):
    return re.sub(r"\*\*", "", str(t or ""))


def build_payload(aid, result, s):
    admin = ROLE_RANK.get(role_of(s), 0) >= 1
    df, intent, report = result.get("df"), result.get("intent") or {}, result.get("report")
    has = df is not None and not df.empty
    note = result.get("note") or ""
    note_txt, _, timer = note.partition("⏱️")
    sql = result.get("sql") or ""
    out = {"id": aid, "question": result.get("question"), "report": report, "text": result.get("text") or "",
           "note": note_txt.strip(), "time": timer.strip(), "options": result.get("options") or [],
           "sql": sql if admin else pe.redact_sql(sql, B["RAW_T"]), "params": [str(p) for p in result.get("params") or []],
           "explain": result.get("explain") or "", "explain_pending": bool(result.get("explain_pending")),
           "kpis": [], "table": None, "chart": None}
    if has:
        try:
            out["kpis"] = [list(k) for k in B["kpis"](report, df, intent.get("pct_basis", "ctc"))]
        except Exception:
            pass
        out["table"] = _table(df, intent)
        out["chart"] = chart_spec(report, df, intent)
    return out


# ---------------------------------------------------------------------------------------------------------
# 4) API
# ---------------------------------------------------------------------------------------------------------
_STATUS = {"at": 0, "db": False, "ai": False}


def _health():
    if time.time() - _STATUS["at"] > 60:
        try:
            B["get_conn"]().close()
            _STATUS["db"] = True
        except Exception:
            _STATUS["db"] = False
        try:
            B["ollama"].list()
            _STATUS["ai"] = True
        except Exception:
            _STATUS["ai"] = False
        _STATUS["at"] = time.time()
    return _STATUS["db"], _STATUS["ai"]


IP_FAILS = {}                 # ip -> [wrong tries, first try time]: stops one computer guessing many IDs
IP_MAX_FAILS = 20


def _ip_blocked(ip):
    rec = IP_FAILS.get(ip)
    if rec and time.time() - rec[1] > 15 * 60:
        IP_FAILS.pop(ip, None)
        return False
    return bool(rec) and rec[0] >= IP_MAX_FAILS


async def api_login(request):
    d = await _body(request)
    uid, pw = str(d.get("uid", "")).strip().lower()[:60], str(d.get("pw", ""))[:200]
    ip = _client_ip(request)
    if _ip_blocked(ip):
        audit(uid or "(blank)", "login_ip_blocked", f"from {ip}")
        return _err("Too many wrong tries from this computer. Try again in 15 minutes.", 429)
    users = B["load_users"]()
    rec = users.get(uid)
    if B["is_locked"](uid):
        audit(uid, "login_locked", f"from {ip}")
        return _err("Too many wrong tries. This ID is locked for 15 minutes.", 429)
    if not rec or not B["check_password"](rec, pw):
        B["note_fail"](uid)
        IP_FAILS.setdefault(ip, [0, time.time()])[0] += 1
        audit(uid or "(blank)", "login_failed", f"from {ip}")
        return _err("Wrong ID or password.", 401)
    if not rec.get("active", True):
        audit(uid, "login_disabled", f"from {ip}")
        return _err("This ID is disabled. Please contact the administrator.", 403)
    B["_login_fails"]().pop(uid, None)
    tok = secrets.token_urlsafe(32)
    with _S_LOCK:
        SESSIONS[tok] = {"uid": uid, "rec": rec, "state": _new_state(rec), "answers": {}, "order": [],
                         "last": time.time(), "born": time.time(), "ip": ip}
    audit(uid, "login", f"web from {ip}")
    resp = _json({"ok": True, "must_change": bool(rec.get("must_change"))})
    resp.set_cookie(COOKIE, tok, httponly=True, samesite="strict", path="/")   # no max_age: gone when the browser closes
    return resp


async def api_logout(request):
    tok = request.cookies.get(COOKIE)
    with _S_LOCK:
        s = SESSIONS.pop(tok, None) if tok else None
    if s:
        audit(s["uid"], "logout", "web")
    resp = _json({"ok": True})
    resp.delete_cookie(COOKIE, path="/")
    return resp


@need_login()
def api_me(request, s):
    rec, st_ = s["rec"], s["state"]
    db, ai = _health()
    mode = st_.get("cfg_mode", "month_year")
    latest, rng, whole = "", [None, None], [None, None]
    admin = ROLE_RANK.get(role_of(s), 0) >= 1
    if db:
        try:
            with _UserContext(s):
                m, y = B["latest_period"](mode)
                latest = B["period_label"](m, y)
                rng = list(B["data_range"]())
                if admin:
                    whole = list(B["data_range"](whole=True))
        except Exception:
            pass
    rs = B["_refresh_status"]()
    df_ = B["data_from"]()
    return _json({
        "uid": s["uid"], "name": rec.get("name") or s["uid"], "role": rec.get("role", "user"),
        "must_change": bool(rec.get("must_change")), "theme": rec.get("theme", "dark"),
        "prefs": {"mode": mode, "always_ai": bool(st_.get("cfg_always_ai")), "hide": bool(st_.get("hide_cols"))},
        "status": {"db": db, "ai": ai, "model": B["MODEL"], "table": B["TABLE"], "latest": latest,
                   "dummy": B["DUMMY_DATA"], "dummy_msg": B["DUMMY_DATA_MSG"],
                   "refreshed": rs.get("at", ""), "refresh_by": rs.get("by", ""), "refresh_status": rs.get("status", ""),
                   "range": rng, "whole": whole, "data_from": df_.get("from", "")[:7],
                   "data_from_label": (datetime.date.fromisoformat(df_["from"]).strftime("%b-%Y") if df_.get("from") else ""),
                   "data_from_by": df_.get("by", ""), "data_from_at": df_.get("at", "")},
        "expiry": _expiry_warning(s),
    })


@need_login()
async def api_password(request, s):
    d = await _body(request)
    cur, new, new2 = str(d.get("current", "")), str(d.get("new", "")), str(d.get("confirm", ""))
    rec = B["load_users"]().get(s["uid"])
    err = (None if rec and B["check_password"](rec, cur) else "Current password is wrong.") \
        or B["pw_problem"](new) or (None if new == new2 else "New passwords do not match.") \
        or (None if new != cur else "New password must be different.")
    if err:
        return _err(err)
    B["set_password"](rec, new, must_change=False)
    B["update_user"](s["uid"], salt=rec["salt"], hash=rec["hash"], must_change=False)
    audit(s["uid"], "password_changed", "web")
    return _json({"ok": True})


def _validity():
    until = B["valid_until"]()
    left = (until - datetime.date.today()).days if until else None
    return {"until": until.strftime("%d-%m-%Y") if until else None, "iso": until.isoformat() if until else "",
            "days_left": left, "expired": until is None or left < 0}


@need_login(master=True)
def api_validity(request, s):
    return _json(_validity())


@need_login(master=True)
async def api_validity_set(request, s):
    """Master only, and the master password is asked again: sets the date the chatbot works until."""
    d = await _body(request)
    rec = B["load_users"]().get(s["uid"])
    if not rec or not B["check_password"](rec, str(d.get("password", ""))):
        audit(s["uid"], "validity_denied", "wrong password")
        return _err("Password is wrong. Nothing changed.")
    try:
        day = datetime.date.fromisoformat(str(d.get("date", "")))
    except ValueError:
        return _err("Please pick a date.")
    if day < datetime.date.today():
        return _err("That date is already over - pick a later date.")
    old = _validity()["until"] or "not set"
    B["write_valid_until"](day)
    audit(s["uid"], "validity_set", f"valid till {day.strftime('%d-%m-%Y')} (was {old})")
    return _json({"ok": True, **_validity()})


def _ai_up():
    try:
        _OLLAMA.list()
        return True
    except Exception:
        return False


def _ai_state():
    up = _ai_up()
    _STATUS.update(ai=up, at=time.time())                  # the status line under the chat follows at once
    return {"running": up, "model": B["MODEL"]}


def _ollama_exe():
    """The Ollama tray app if installed (it starts the AI server itself), else the ollama command."""
    cli = shutil.which("ollama") or os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Ollama", "ollama.exe")
    app = os.path.join(os.path.dirname(cli), "ollama app.exe")
    return [app] if os.path.exists(app) else [cli, "serve"]


@need_login()
def api_ai(request, s):
    return _json(_ai_state())


@need_login()
async def api_ai_set(request, s):
    """Anyone signed in (the AI on/off button on the chat page): switches the AI (Ollama + qwen) on or off for
    everyone. Off = questions that need AI say 'AI not reachable'; the quick answers without AI keep working."""
    act = str((await _body(request)).get("action", ""))
    if act == "start":
        if not _ai_up():
            try:
                subprocess.Popen(_ollama_exe(), cwd=str(APP_DIR), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            except OSError:
                return _err("Could not start Ollama - is it installed on this PC?")
            for _ in range(30):                              # wait up to 30 s for it to answer
                await asyncio.sleep(1)
                if _ai_up():
                    break
            else:
                return _err("Ollama was started but is not answering yet. Try again in a minute.")
            threading.Thread(target=_warm_up, daemon=True).start()
        audit(s["uid"], "ai_start", B["MODEL"])
    elif act == "stop":
        for exe in ("ollama app.exe", "ollama.exe"):
            subprocess.run(["taskkill", "/F", "/IM", exe], capture_output=True,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        await asyncio.sleep(1)
        audit(s["uid"], "ai_stop", B["MODEL"])
    else:
        return _err("Unknown action.")
    return _json({"ok": True, **_ai_state()})


@need_login()
async def api_prefs(request, s):
    d = await _body(request)
    st_, upd = s["state"], {}
    if d.get("mode") in ("month_year", "salary", "payment"):
        st_["cfg_mode"] = upd["mode"] = d["mode"]
    if isinstance(d.get("always_ai"), bool):
        st_["cfg_always_ai"] = upd["always_ai"] = d["always_ai"]
    if isinstance(d.get("hide"), bool):
        st_["hide_cols"] = upd["hide_cols"] = list(B["DEFAULT_HIDE"]) if d["hide"] else []
    if d.get("theme") in ("dark", "light"):
        upd["theme"] = d["theme"]
    if upd:
        B["update_user"](s["uid"], **upd)
    return _json({"ok": True})


def _wo_reply(question, text, options=None):
    return {"text": text, "note": "Work-order groups: say e.g. *list work order rules* or *remove the SOUTH rule*.",
            "df": None, "report": "pref", "sql": "", "params": [], "intent": {}, "question": question,
            "options": options or []}


def _wo_group_command(s, question):
    """'treat work orders starting M24 as NORTH' -> asks 'Save this rule?' first; only the master / admins may change
    the groups. Returns a reply, or None for a normal question."""
    st_ = s["state"]
    pending = st_.pop("pending_wo_groups", None)          # a yes / no answers it; any other message drops it
    if pending and wg.YES_RX.match(question):
        if ROLE_RANK.get(role_of(s), 0) < 1 or pending.get("by") != s["uid"]:       # re-checked, not trusted
            audit(s["uid"], "wo_group_denied", question[:250])
            return _wo_reply(question, "🔒 Only the master or an admin can change work-order groups. Nothing was saved.")
        lines = WO_GROUPS.apply(pending["actions"])
        audit(s["uid"], "wo_group", "; ".join(re.sub(r"\*\*", "", x) for x in lines)[:280])
        return _wo_reply(question, "✅ **Saved for everyone:**\n" + "\n".join("- " + x for x in lines) +
                         "\n\nNow ask e.g. *NORTH margin for sep-26* or *work order group wise base sales*.")
    if pending and wg.NO_RX.match(question):
        return _wo_reply(question, "Cancelled - nothing was saved.")
    acts = wg.parse(question, WO_GROUPS.names())
    if not acts:
        return None
    if acts == [("list",)]:
        return _wo_reply(question, "**Work-order groups (apply to everyone):**\n" + WO_GROUPS.describe())
    if ROLE_RANK.get(role_of(s), 0) < 1:
        audit(s["uid"], "wo_group_denied", question[:250])
        return _wo_reply(question, "🔒 Only the master or an admin can create, change or remove work-order groups. "
                                   "I didn't change anything.")
    if acts[0][0] == "error":
        return _wo_reply(question, acts[0][1] + "\n\n**Work-order groups now:**\n" + WO_GROUPS.describe())
    st_["pending_wo_groups"] = {"actions": acts, "by": s["uid"]}
    return _wo_reply(question, "You're about to change the **work-order groups** (for everyone):\n" +
                     "\n".join("- " + x for x in WO_GROUPS.describe_actions(acts)) +
                     "\n\n**Save this rule? (yes / no)**", ["Yes", "No"])


def _run_question(s, question, from_option=False):
    """Same steps as the Streamlit chat page: follow-ups, 'did you mean' replies, permissions, quick insight."""
    st_, rec = s["state"], s["rec"]
    admin = ROLE_RANK.get(role_of(s), 0) >= 1                 # master or admin may use "/" SQL
    is_sql = B["is_user_sql"](question) if admin else question.lstrip().startswith("/")
    rule_msg = ""
    if not is_sql:                                        # "treat workorder no start with M24% as NORTH ..."
        wo = _wo_group_command(s, question)
        if wo:
            return wo
    if not is_sql:                                        # "rename Employee column to Employee code", "no commas in ..."
        acts, rest = dr.parse(question)
        if acts:
            store = B["POLICY_STORE"]
            allowed = role_of(s) == "master" or bool(store.user(store.load(), s["uid"]).get("display_rules"))
            if not allowed:
                rule_msg = ("🔒 Only the master administrator (or people the master allows) can change how columns are "
                            "named or shown. I didn't change anything.")
                audit(s["uid"], "display_rule_denied", question[:250])
            elif acts == [("show",)]:
                rule_msg = "**Display rules (apply to everyone):**\n" + RULES.describe()
            else:
                lines = RULES.apply(acts)
                audit(s["uid"], "display_rule", "; ".join(re.sub(r"\*\*", "", x) for x in lines)[:280])
                rule_msg = "✅ **Saved for everyone:**\n" + "\n".join("- " + x for x in lines)
            if not rest or len(rest.split()) < 3:          # nothing else asked
                return {"text": rule_msg, "note": "Display rules: say e.g. *show display rules* or *reset display rules*.",
                        "df": None, "report": "pref", "sql": "", "params": [], "intent": {}, "question": question}
            question = rest                                # answer the rest of the message too
    run_q = question
    prev = st_.get("clarify_for")
    if prev and not from_option and len(question.split()) <= 4 and not is_sql:   # short reply to "what did you mean?"
        run_q = f"{prev} {question}"
    elif not from_option and not is_sql:
        fu = B["follow_up"](question, st_.get("last_q"))       # "and jul-26?" / "what about IT?"
        if fu:
            run_q = fu
    st_["clarify_for"] = None
    always_ai = bool(st_.get("cfg_always_ai", B["ALWAYS_AI"]))
    use_ai = always_ai or bool(st_.get("cfg_use_ai", True))
    result = B["answer"](run_q, st_.get("cfg_mode", "month_year"), use_ai, always_ai)
    result = B["secure_result"](result, run_q)
    if run_q != question and result.get("note") is not None and result.get("report") != "clarify":
        result["note"] = f"Read as: “{run_q}”. " + result.get("note", "")
    if result.get("report") == "clarify":
        st_["clarify_for"] = result.get("clarify_for")
    if result.get("report") not in B["NO_EXPLAIN"] and not is_sql:
        st_["last_q"] = result.get("_last_q") or run_q
    result["question"] = question
    d = result.get("df")
    if result.get("report") not in B["NO_EXPLAIN"] and d is not None and not d.empty:
        try:
            result["explain"] = B["quick_explain"](result)
        except Exception:
            result["explain"] = ""
    ask_ai = B["AI_EXPLAIN_ASK"].search(question)
    want_ai = result.get("report") == "chat_ai" or (
        st_.get("cfg_explain", B["CHAT_STYLE"])
        and (st_.get("cfg_ai_explain", B["AI_EXPLAIN"]) == "always" or ask_ai)
        and result.get("report") not in B["NO_EXPLAIN"] and d is not None and not d.empty
        and not (result.get("report") in ("summary", "employee_detail", "workorder_detail", "employee_list",
                                          "joiners", "leavers") and not ask_ai))
    if use_ai and want_ai:
        result["explain_pending"] = True
    if rule_msg:                                          # rule saved + the question answered in one reply
        result["text"] = rule_msg + "\n\n" + (result.get("text") or "")
    warn = _expiry_warning(s)
    if warn and result.get("report") != "expired":
        result["note"] = "⏳ " + warn + " " + (result.get("note") or "")
    return result


def _expiry_warning(s):
    """The master is reminded 15 days before the validity date, everyone else 5 days before."""
    until = B["valid_until"]()
    if not until:
        return ""
    left = (until - datetime.date.today()).days
    master = role_of(s) == "master"
    if left < 0 or left > (15 if master else 5):
        return ""
    when = "today" if left == 0 else f"in {left} day{'s' if left != 1 else ''}"
    return (f"The chatbot expires {when} ({until.strftime('%d-%m-%Y')}). "
            + ("Renew it in Settings → Chatbot validity." if master else "Please ask the administrator to renew it."))


@need_login()
async def api_ask(request, s):
    d = await _body(request)
    q = str(d.get("q", "")).strip()[:2000]
    if not q:
        return _err("Please type a question.")
    audit(s["uid"], "question", q)
    from starlette.concurrency import run_in_threadpool
    rid = str(d.get("rid") or "")[:40] or secrets.token_hex(8)
    job = _job_start(s, rid, q)
    if job is None:
        return _json({"stopped": True})
    s["qgen"] = gen = s.get("qgen", 0) + 1
    keep = {k: s["state"].get(k) for k in ("last_q", "clarify_for")}

    def work():
        with _UserContext(s, job):
            try:
                result = _run_question(s, q, bool(d.get("from_option")))
            except Exception as e:
                if not job.cancelled:
                    logging.exception("question failed")
                    return {"id": None, "text": f"❌ Something went wrong while answering: {str(e)[:300]}",
                            "note": "The error was logged. Try rephrasing, or ask the administrator.", "report": "error"}
            if job.cancelled:                  # stopped: throw the answer away, follow-ups carry on from before
                if s.get("qgen") == gen:
                    s["state"].update(keep)
                return {"stopped": True}
            aid = secrets.token_hex(6)
            s["answers"][aid] = result
            s["order"].append(aid)
            while len(s["order"]) > MAX_ANSWERS:
                s["answers"].pop(s["order"].pop(0), None)
            return build_payload(aid, result, s)
    try:
        return _json(await run_in_threadpool(work))
    finally:
        _job_end(job)


@need_login()
async def api_stop(request, s):
    from starlette.concurrency import run_in_threadpool
    d = await _body(request)
    rid = str(d.get("rid") or "")[:40]
    if not rid:
        return _err("Nothing to stop.")
    with _J_LOCK:
        for k in [k for k, j in JOBS.items() if j.cancelled and time.time() - j.t0 > 600]:
            JOBS.pop(k, None)                   # old "stopped before it started" marks
        job = JOBS.get(rid)
        if job is None:                         # Stop beat the question to the server
            JOBS[rid] = _Job(s, rid, "")
            JOBS[rid].cancelled = True
            return _json({"ok": True})
    if job.s is not s:
        return _err("Nothing to stop.", 404)
    if not job.cancelled:
        await run_in_threadpool(job.stop)
        where = "" if job.stage == "question" else f", during the {job.stage}"
        audit(s["uid"], "question_stopped", f"{job.q[:250]} (after {time.time() - job.t0:.0f}s{where})")
    return _json({"ok": True})


@need_login()
def api_explain(request, s):
    aid = request.path_params["aid"]
    msg = s["answers"].get(aid)
    if not msg:
        return _err("That answer has expired. Please ask again.", 404)
    if msg.get("explain_ai"):
        return _json({"text": msg["explain_ai"]})
    rid = request.query_params.get("rid", "")[:40]
    job = _job_start(s, rid, msg.get("question", ""), "AI write-up") if rid else None
    if rid and job is None:
        return _json({"stopped": True})
    try:
        with _UserContext(s, job):
            try:
                text = "".join(B["explain_stream"](msg)).strip()
                eff = B["current_eff"]()
                if eff is not None:
                    text = pe.secure_text(text, eff, B["current_prefs"](), msg.get("question", ""))
            except Exception as e:
                if job and job.cancelled:
                    return _json({"stopped": True})
                return _json({"text": "", "error": f"The AI is not reachable right now ({e.__class__.__name__})."})
        if job and job.cancelled:              # finished just as Stop was pressed: not kept
            return _json({"stopped": True})
    finally:
        if job:
            _job_end(job)
    msg["explain_ai"] = text
    return _json({"text": text})


@need_login()
def api_csv(request, s):
    aid = request.path_params["aid"]
    msg = s["answers"].get(aid)
    if not msg or msg.get("df") is None:
        return _err("That answer has expired. Please ask again.", 404)
    with _UserContext(s):
        try:
            data, n = B["full_csv"](msg)
        except PermissionError as e:
            return _err(str(e), 403)
        except Exception as e:
            return _err(f"Could not read all rows: {str(e)[:200]}", 500)
    data = _rename_csv(data)
    safe = re.sub(r"[^A-Za-z0-9]+", "_", str(msg.get("question") or "result"))[:40].strip("_") or "result"
    audit(s["uid"], "csv_download", f"{safe} ({n} rows)")
    return Response(data, media_type="text/csv",
                    headers={"Content-Disposition": f'attachment; filename="{safe}.csv"', "Cache-Control": "no-store"})


def _audit_rows():
    try:
        with B["AUDIT_FILE"].open(encoding="utf-8", newline="") as f:
            return list(csv.DictReader(f))
    except FileNotFoundError:
        return []


@need_login()
def api_history(request, s):
    rec = s["rec"]
    hidden, cleared = set(rec.get("hidden_requests", [])), rec.get("requests_cleared_at", "")
    seen, out = set(), []
    for r in reversed(_audit_rows()):
        if r.get("user") != s["uid"] or r.get("action") != "question":
            continue
        q = (r.get("detail") or "").strip()
        k = q.lower()
        if not q or k in seen or k in hidden or (cleared and r.get("time", "") <= cleared):
            continue
        seen.add(k)
        out.append({"q": q, "time": r.get("time", "")})
        if len(out) >= 150:
            break
    return _json({"items": out})


@need_login()
async def api_history_hide(request, s):
    d = await _body(request)
    with B["_users_lock"]():
        users = B["load_users"]()
        u = users.get(s["uid"])
        if u is not None:
            if d.get("all"):
                u["requests_cleared_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                u["hidden_requests"] = []
            elif d.get("q"):
                h = set(u.get("hidden_requests", []))
                h.add(str(d["q"]).strip().lower())
                u["hidden_requests"] = sorted(h)[-500:]
            B["save_users"](users)
    return _json({"ok": True})


# ---------------- TDS calculator (worked out on this PC, nothing sent outside) ----------------
@need_login()
async def api_tds_calc(request, s):
    d = await _body(request)
    gone = B["expired_answer"]()
    if gone:
        return _err(re.sub(r"\*\*", "", gone["text"]))
    try:
        calc = B["tds_calculate"](d.get("paid"), d.get("proposed"), d.get("regime"), d.get("tds_paid"), d.get("months"))
    except ValueError as e:
        return _err(str(e))
    return _json({"ok": True, "item": B["tds_history_add"](s["uid"], calc)})


@need_login()
def api_tds_history(request, s):
    fy, months = B["tds_fy"]()
    return _json({"items": B["tds_history"](s["uid"]), "fy": fy, "months": months})


@need_login()
async def api_tds_delete(request, s):
    d = await _body(request)
    if d.get("all"):
        n = B["tds_history_delete"](s["uid"])
    elif d.get("id"):
        n = B["tds_history_delete"](s["uid"], str(d["id"])[:40])
    else:
        return _err("Nothing to delete.")
    return _json({"ok": True, "deleted": n})


# ---------------- admin ----------------
@need_login(admin=True)
async def api_data_from(request, s):
    """Master / admin: {"month": "2026-04"} -> everyone sees data from Apr-2026 on; {"clear": true} -> all months."""
    from starlette.concurrency import run_in_threadpool
    d = await _body(request)
    if d.get("clear"):
        B["set_data_from"](None, s["uid"])
        audit(s["uid"], "data_from_cleared", "show all months again (web)")
        return _json({"ok": True, "data_from": ""})
    try:
        day = datetime.date.fromisoformat(str(d.get("month", ""))[:7] + "-01")
    except ValueError:
        return _err("Please pick a month.")

    def last_month():
        with _UserContext(s):
            return B["data_range"](whole=True)[1]
    try:
        last = await run_in_threadpool(last_month)
    except Exception as e:
        return _err(f"Could not read the table: {str(e)[:200]}", 500)
    if last and day > datetime.datetime.strptime(last, "%b-%Y").date():
        return _err(f"The table only has data up to {last}, so nobody would see anything. Nothing changed.")
    B["set_data_from"](day, s["uid"])
    label = B["period_label"](day.month, day.year)
    audit(s["uid"], "data_from_set", f"show data from {label} (web)")
    return _json({"ok": True, "data_from": day.isoformat()[:7], "label": label})


@need_login(admin=True)
async def api_refresh(request, s):
    from starlette.concurrency import run_in_threadpool
    if B["DUMMY_DATA"]:                                     # dummy data: nothing to refresh from
        audit(s["uid"], "data_refresh_skipped", "dummy data - no data update system (web)")
        return _json({"ok": True, "dummy": True, "message": B["DUMMY_DATA_MSG"]})
    if not _REFRESH_LOCK.acquire(blocking=False):
        return _err("A refresh is already running. Please wait.", 409)
    try:
        def work():
            try:
                old, new, secs = B["refresh_payroll_data"]()
            except Exception as e:
                B["_save_refresh_status"](s["uid"], None, None, 0, f"failed: {str(e)[:200]}")
                audit(s["uid"], "data_refresh_failed", str(e)[:300])
                return {"ok": False, "error": f"Refresh failed, {B['TABLE']} was NOT changed. ({str(e)[:300]})"}
            B["_save_refresh_status"](s["uid"], old, new, secs, "ok")
            audit(s["uid"], "data_refresh", f"{old} -> {new} rows in {secs:.0f}s (web)")
            B["st"].cache_data.clear()                          # answers now use the new data
            return {"ok": True, "old": old, "new": new, "seconds": round(secs)}
        res = await run_in_threadpool(work)
        return _json(res, 200 if res.get("ok") else 500)
    finally:
        _REFRESH_LOCK.release()


def _active_admins(users, but=None):
    return [k for k, v in users.items() if v.get("role") == "master" and v.get("active", True) and k != but]


def _access_summary(pol, role):
    if role in ("admin", "master"):
        return f"Full ({role})"
    if not pol.get("payroll_access", True):
        return "No data access"
    bits = []
    if pol.get("aggregated_only"):
        bits.append("totals only")
    if not pol.get("sensitive_access", True):
        bits.append("no salary details")
    if pol.get("blocked_columns"):
        bits.append(f"{len(pol['blocked_columns'])} hidden column(s)")
    if pol.get("allowed_columns"):
        bits.append("only listed columns")
    if pol.get("row_categories"):
        bits.append("categories: " + ", ".join(pol["row_categories"]))
    if pol.get("row_departments"):
        bits.append(f"{len(pol['row_departments'])} department(s)")
    extra = " · can set display rules" if pol.get("display_rules") else ""
    return ("Full" if not bits else "; ".join(bits)) + extra


@need_login(admin=True)
def api_users(request, s):
    users = B["load_users"]()
    store = B["POLICY_STORE"]
    data = store.load()
    online = {x["uid"] for x in SESSIONS.values() if time.time() - x["last"] < 15 * 60}
    master = role_of(s) == "master"                       # an admin never sees the master ID or its details
    return _json({"me": s["uid"], "my_role": role_of(s), "users": [
        {"id": k, "name": v.get("name", k), "role": v.get("role", "user"), "active": v.get("active", True),
         "created": v.get("created", ""), "must_change": bool(v.get("must_change")), "online": k in online,
         "access": _access_summary(store.user(data, k), v.get("role", "user"))}
        for k, v in sorted(users.items()) if master or v.get("role") != "master"]})


# columns an admin can hide per person (key -> label); Year/Month/WorkOrderNo stay (needed to read any answer)
ACCESS_COLUMNS = ["Name", "Employee", "CTC", "BaseSalesValue", "InvoiceAmount", "Margin", "MarginPct", "AvgCTC",
                  "GrossAmount", "TakeHome", "Salary", "Basic", "HRA", "Bonus", "Arrears", "IncomeTax", "Total PF",
                  "PT", "ESIC", "PaidDays", "joindate", "resigndate", "WOStartDate", "WOEndDate", "Remarks"]


@need_login(master=True)
def api_access_get(request, s):
    uid = request.path_params["uid"]
    users = B["load_users"]()
    if uid not in users:
        return _err("No such user.", 404)
    store = B["POLICY_STORE"]
    pol = store.user(store.load(), uid)
    with _UserContext(s):
        cats = [str(v) for v in B["run_query"](f"SELECT DISTINCT [Category] AS v FROM {B['RAW_T']} "
                                               "WHERE [Category] IS NOT NULL ORDER BY 1", [])["v"]]
        deps = [str(v) for v in B["run_query"](f"SELECT DISTINCT [Department] AS v FROM {B['RAW_T']} "
                                               "WHERE [Department] IS NOT NULL ORDER BY 1", [])["v"]]
    return _json({
        "id": uid, "name": users[uid].get("name", uid), "role": users[uid].get("role", "user"),
        "policy": {"payroll_access": bool(pol.get("payroll_access", True)),
                   "aggregated_only": bool(pol.get("aggregated_only", False)),
                   "sensitive_access": bool(pol.get("sensitive_access", True)),
                   "display_rules": bool(pol.get("display_rules", False)),
                   "blocked_columns": list(pol.get("blocked_columns", [])),
                   "categories": list(pol.get("row_categories", [])), "departments": list(pol.get("row_departments", []))},
        "options": {"columns": [{"key": k, "label": pe.label(k)} for k in ACCESS_COLUMNS],
                    "categories": cats, "departments": deps,
                    "sensitive": [pe.label(k) for k in store.load()["global"].get("sensitive_columns", [])]}})


@need_login(master=True)
async def api_access_set(request, s):
    d = await _body(request)
    uid = str(d.get("id", "")).strip().lower()
    users = B["load_users"]()
    if uid not in users:
        return _err("No such user.", 404)
    if users[uid].get("role") in ("admin", "master"):
        return _err("Admins and the master always have full access. Make them a user first to restrict them.")
    store, me, reason = B["POLICY_STORE"], s["uid"], "set in web admin"
    if d.get("reset"):
        store.apply({"op": "RESET_USER_POLICY", "scope": "USER", "user": uid}, me, reason)
        audit(me, "access_changed", f"{uid}: back to full access")
        return _json({"ok": True, "message": f"{uid} has full access again."})
    pol = store.user(store.load(), uid)
    valid_cols = set(ACCESS_COLUMNS)
    changes = []
    for key, op in (("payroll_access", "SET_PAYROLL"), ("aggregated_only", "SET_AGG_ONLY"),
                    ("sensitive_access", "SET_SENSITIVE"), ("display_rules", "SET_DISPLAY")):
        if isinstance(d.get(key), bool) and d[key] != bool(pol.get(key, key not in ("aggregated_only", "display_rules"))):
            changes.append(store.apply({"op": op, "scope": "USER", "user": uid, "value": d[key]}, me, reason))
    if isinstance(d.get("blocked_columns"), list):
        want = {c for c in d["blocked_columns"] if c in valid_cols}
        cur = set(pol.get("blocked_columns", []))
        if want - cur:
            changes.append(store.apply({"op": "BLOCK_COLUMN", "scope": "USER", "user": uid,
                                        "columns": sorted(want - cur)}, me, reason))
        if (cur - want) & valid_cols:
            changes.append(store.apply({"op": "UNBLOCK_COLUMN", "scope": "USER", "user": uid,
                                        "columns": sorted((cur - want) & valid_cols)}, me, reason))
    for col, key, pkey in (("Category", "categories", "row_categories"), ("Department", "departments", "row_departments")):
        if isinstance(d.get(key), list):
            vals = sorted({str(v)[:200] for v in d[key] if str(v).strip()})
            if vals != sorted(pol.get(pkey, [])):
                changes.append(store.apply({"op": "SET_ROWS", "scope": "USER", "user": uid, "column": col,
                                            "columns": vals}, me, reason))
    if changes:
        audit(me, "access_changed", f"{uid}: " + "; ".join(changes)[:280])
    return _json({"ok": True, "message": (f"Saved for {uid}: " + "; ".join(changes)) if changes else "No changes."})


@need_login(admin=True)
async def api_users_action(request, s):
    d = await _body(request)
    act, uid = d.get("action"), str(d.get("id", "")).strip().lower()
    me, my_role = s["uid"], role_of(s)
    with B["_users_lock"]():
        users = B["load_users"]()
        if act == "add":                                   # admin AND master may add people
            name, role = str(d.get("name", "")).strip()[:80], d.get("role", "user")
            if my_role != "master" and role != "user":
                return _err("Only the master administrator can create admins. You can add users.", 403)
            pw = B["temp_password"]()                      # always made here, never typed (browsers autofilled it)
            err = (None if B["ID_RULE"].match(uid) else "ID: 3-30 characters, letters / numbers / . _ - only.") \
                or (None if uid not in users else "That ID already exists." if my_role == "master"
                    or users[uid].get("role") != "master" else "That ID can't be used. Please choose another.") \
                or (None if role in ("user", "admin") else "Role must be user or admin.") or B["pw_problem"](pw)
            if err:
                return _err(err)
            rec = {"name": name or uid, "role": role, "active": True,
                   "created": datetime.date.today().isoformat(), "created_by": me}
            B["set_password"](rec, pw, must_change=True)
            users[uid] = rec
            B["save_users"](users)
            audit(me, "user_added", f"{uid} ({role})")
            return _json({"ok": True, "message": f"Added {uid}. Temporary password: {pw} (they must change it at first sign-in).",
                          "id": uid, "password": pw, "name": rec["name"], "kind": "added"})
        if my_role != "master":                            # everything below: master only
            audit(me, "user_action_denied", f"{act} {uid}")
            return _err("Only the master administrator can reset passwords, disable, delete, sign out or change people. "
                        "You can add new users.", 403)
        if uid not in users:
            return _err("No such user.", 404)
        if uid == me and act in ("toggle", "role", "delete", "signout", "reset"):
            return _err("You can't change your own access here (change your own password in Settings).")
        if users[uid].get("role") == "master":
            return _err("The master administrator can't be changed here.")
        if act == "rename":                                # names are set by the master only
            name = re.sub(r"\s+", " ", str(d.get("name", ""))).strip()[:80]
            if len(name) < 2:
                return _err("Please type a name.")
            old = users[uid].get("name", uid)
            users[uid]["name"] = name
            B["save_users"](users)
            audit(me, "user_renamed", f"{uid}: {old} -> {name}")
            return _json({"ok": True, "message": f"{uid} is now named {name}."})
        if act == "signout":
            _drop_sessions(uid)
            audit(me, "user_signed_out", uid)
            return _json({"ok": True, "message": f"{uid} has been signed out everywhere."})
        if act == "reset":
            pw = B["temp_password"]()
            B["set_password"](users[uid], pw, must_change=True)
            B["save_users"](users)
            audit(me, "password_reset", uid)
            _drop_sessions(uid)                                # their old sign-in stops working at once
            return _json({"ok": True, "message": f"New temporary password for {uid}: {pw}",
                          "id": uid, "password": pw, "name": users[uid].get("name", uid), "kind": "reset"})
        if act == "toggle":
            active = users[uid].get("active", True)
            if active and users[uid].get("role") == "admin" and not _active_admins(users, but=uid):
                return _err("At least one active administrator is needed.")
            users[uid]["active"] = not active
            B["save_users"](users)
            audit(me, "user_disabled" if active else "user_enabled", uid)
            _drop_sessions(uid)
            return _json({"ok": True, "message": f"{uid} {'disabled' if active else 'enabled'}."})
        if act == "role":
            new = d.get("role")
            if new not in ("user", "admin"):
                return _err("Role must be user or admin.")
            old = users[uid].get("role")
            if old == "admin" and new != "admin" and not _active_admins(users, but=uid):
                return _err("At least one active administrator is needed.")
            users[uid]["role"] = new
            B["save_users"](users)
            audit(me, "role_changed", f"{uid}: {old} -> {new}")
            if new == "admin" and uid in B["POLICY_STORE"].load()["users"]:      # admins have full access
                B["POLICY_STORE"].apply({"op": "RESET_USER_POLICY", "scope": "USER", "user": uid}, me, "made admin")
            try:
                B["POLICY_STORE"].log(me, "CHANGE_ROLE", "USER", uid, "role", old, new)
            except Exception:
                pass
            _drop_sessions(uid)
            return _json({"ok": True, "message": f"{uid} is now {new}."})
        if act == "delete":
            if users[uid].get("role") == "admin" and not _active_admins(users, but=uid):
                return _err("At least one active administrator is needed.")
            users.pop(uid)
            B["save_users"](users)
            audit(me, "user_deleted", uid)
            if uid in B["POLICY_STORE"].load()["users"]:              # a new person with this ID starts clean
                B["POLICY_STORE"].apply({"op": "RESET_USER_POLICY", "scope": "USER", "user": uid}, me, "user deleted")
            _drop_sessions(uid)
            return _json({"ok": True, "message": f"{uid} deleted."})
    return _err("Unknown action.")


def _drop_sessions(uid):
    with _S_LOCK:
        for k in [k for k, v in SESSIONS.items() if v["uid"] == uid]:
            SESSIONS.pop(k, None)


@need_login(master=True)
def api_activity(request, s):
    items = [{"time": r.get("time"), "user": r.get("user"), "action": r.get("action"), "detail": r.get("detail")}
             for r in _audit_rows()[-400:]]
    for r in B["POLICY_STORE"].read_log(200):          # access / restriction changes (policy_audit.csv)
        items.append({"time": r.get("time"), "user": r.get("admin"), "action": "access: " + str(r.get("action", "")),
                      "detail": f"{r.get('target')}: {r.get('item')} {r.get('old')} → {r.get('new')}"})
    items.sort(key=lambda x: x["time"] or "", reverse=True)
    return _json({"items": items[:500]})


def index(request):
    return FileResponse(WEB_DIR / "index.html", headers={"Cache-Control": "no-cache"})


class _Security(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.method == "POST" and request.url.path.startswith("/api/"):
            if request.headers.get("x-requested-with") != "payroll-ui":     # blocks cross-site form posts
                return _err("Bad request.", 400)
        resp = await call_next(request)
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["X-Content-Type-Options"] = "nosniff"
        resp.headers["Referrer-Policy"] = "no-referrer"
        resp.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline' "
            "https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com data:; img-src 'self' data:; "
            "connect-src 'self'; frame-ancestors 'none'")
        return resp


routes = [
    Route("/", index),
    Route("/api/login", api_login, methods=["POST"]),
    Route("/api/logout", api_logout, methods=["POST"]),
    Route("/api/me", api_me),
    Route("/api/password", api_password, methods=["POST"]),
    Route("/api/prefs", api_prefs, methods=["POST"]),
    Route("/api/ask", api_ask, methods=["POST"]),
    Route("/api/stop", api_stop, methods=["POST"]),
    Route("/api/explain/{aid}", api_explain),
    Route("/api/csv/{aid}", api_csv),
    Route("/api/history", api_history),
    Route("/api/history/hide", api_history_hide, methods=["POST"]),
    Route("/api/tds", api_tds_history),
    Route("/api/tds", api_tds_calc, methods=["POST"]),
    Route("/api/tds/delete", api_tds_delete, methods=["POST"]),
    Route("/api/admin/refresh", api_refresh, methods=["POST"]),
    Route("/api/admin/data-from", api_data_from, methods=["POST"]),
    Route("/api/admin/validity", api_validity),
    Route("/api/admin/validity", api_validity_set, methods=["POST"]),
    Route("/api/ai", api_ai),
    Route("/api/ai", api_ai_set, methods=["POST"]),
    Route("/api/admin/users", api_users),
    Route("/api/admin/users", api_users_action, methods=["POST"]),
    Route("/api/admin/activity", api_activity),
    Route("/api/admin/access/{uid}", api_access_get),
    Route("/api/admin/access", api_access_set, methods=["POST"]),
]
app = Starlette(routes=routes, middleware=[Middleware(_Security)])


def _warm_up():
    try:
        B["ollama"].generate(model=B["MODEL"], prompt="hi", keep_alive="60m", options={"num_predict": 1})
    except Exception:
        pass


if __name__ == "__main__":
    if "--port" in sys.argv:
        PORT = int(sys.argv[sys.argv.index("--port") + 1])
    threading.Thread(target=_warm_up, daemon=True).start()
    print(f"Payroll AI Assistant (web) on http://localhost:{PORT}  - other PCs: http://<this-PC-IP>:{PORT}")
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")
