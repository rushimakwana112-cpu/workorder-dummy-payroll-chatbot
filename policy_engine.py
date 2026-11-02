"""
Policy engine for the Payroll AI chatbot  (pure Python - no Streamlit, so it can be tested on its own)

    LLM            = understands the question
    POLICY ENGINE  = decides what each user is ALLOWED to see        (this file)
    SQL ENGINE     = reads only through a per-user policy view        (payroll_chatbot.py)
    RESULT FILTER  = last security boundary before anything is shown  (secure_df / secure_text here)

Priority (a lower level can never override a higher one):
    1 SYSTEM SECURITY        read-only, payroll table only, query checks
    2 ADMIN GLOBAL POLICY    policies.json -> "global"
    3 ADMIN USER POLICY      policies.json -> "users" -> <id>
    4 USER PREFERENCES       users.json -> <id> -> "data_prefs"   (what the user WANTS to see)
    5 CURRENT REQUEST        the question being asked now

PERMISSION (levels 1-3) = what a user is ALLOWED to see.  PREFERENCE (4-5) = what they WANT to see.
"""
import csv
import datetime
import json
import os
import re
import threading
from pathlib import Path

# ----------------------------------------------------------------------------------------------
# Columns
# ----------------------------------------------------------------------------------------------
KNOWN_COLUMNS = [   # payrollall columns (extended at runtime from the real table)
    "Category", "Month", "Year", "WorkOrderNo", "Employee", "Name", "PaidDays", "Salary", "CTC", "GrossAmount",
    "IncomeTax", "Total PF", "PT", "NetSalary", "ChequeAmount", "InvoiceAmount", "InvoiceDate", "UniqueNo",
    "UniqueNo2", "Arrears", "Remarks", "Department", "WOStartDate", "WOEndDate", "Deduction", "PFTax", "Basic",
    "HRA", "ConveyanceAllowance", "SpecialAllowance", "TransportAllowance", "IncomeTax_2", "PF_Employee",
    "PF_Employer", "ESIC", "Bonus", "ESIC_Employee", "ChildEduAllownce", "MedicalAllownce", "ConveyanceReinburse",
    "MedicalReimburse", "CORP_NPS", "joindate", "resigndate", "tdsopt", "FinaluniqCode", "month_year"]
# needed inside queries for correct maths (invoice once per work order + salary month, distinct employees,
# month filters). They are kept INSIDE SQL even when hidden, and removed from the OUTPUT by the result filter.
TECHNICAL = {"Year", "Month", "month_year", "WorkOrderNo", "Employee"}
# calculated answer columns -> the table columns they are made from
DERIVED = {
    "BaseSalesValue": ["InvoiceAmount"], "Margin": ["InvoiceAmount", "CTC"], "MarginPct": ["InvoiceAmount", "CTC"],
    "AvgCTC": ["CTC"], "TakeHome": ["GrossAmount", "IncomeTax", "PF_Employee", "PT", "ESIC_Employee", "Deduction"],
    "EmployeeCount": [],
}
SENSITIVE_DEFAULT = ["GrossAmount", "NetSalary", "ChequeAmount", "Salary", "Basic", "HRA", "IncomeTax",
                     "IncomeTax_2", "Total PF", "PF_Employee", "PF_Employer", "ESIC", "ESIC_Employee", "PT", "PFTax",
                     "CORP_NPS", "tdsopt", "Bonus", "Arrears", "Deduction", "TakeHome", "BankAccount"]
EMPLOYEE_KEYS = {"Employee", "Name"}          # columns that make a result "employee-level / raw"
EMPLOYEE_REPORTS = {"employee_list", "employee_detail", "workorder_detail", "employee_ranking", "joiners",
                    "leavers", "proposed_tds", "salary_run", "tenure", "wo_history", "wo_breaks"}
# words people use -> key   (keys are table columns or the calculated names above)
ALIASES = {
    "GrossAmount": ["gross salary", "gross amount", "gross pay", "gross", "grosssalary", "grossamount"],
    "UniqueNo": ["unique no", "unique number", "uniqueno", "unique1", "unique 1", "unique no 1"],
    "UniqueNo2": ["unique no 2", "unique number 2", "uniqueno2", "unique2", "unique 2", "unique no2"],
    "Employee": ["employee code", "emp code", "employee id", "emp id", "employee no", "employee number",
                 "employee codes"],
    "Name": ["employee name", "employee names", "emp name", "names"],
    "BankAccount": ["employee bank account", "bank account", "bank accounts", "bank details", "account number",
                    "bank account number", "bank"],
    "CTC": ["ctc", "cost to company"],
    "BaseSalesValue": ["base sales value", "base sales amount", "base sales", "base sale", "basesales",
                       "basesalesvalue", "base value", "base sale value"],
    "InvoiceAmount": ["invoice amount", "invoice value", "invoice", "invoiceamount", "invoice with gst"],
    "MarginPct": ["margin %", "margin percent", "margin percentage", "margin pct", "marginpct"],
    "Margin": ["margin amount", "margin", "profit"],
    "EmployeeCount": ["employee count", "employees count", "no of employees", "number of employees",
                      "head count", "headcount", "emp count", "employees"],
    "AvgCTC": ["average ctc", "avg ctc", "ctc per employee"],
    "TakeHome": ["take home", "take-home", "takehome", "in hand", "inhand"],
    "NetSalary": ["net salary", "net pay", "netsalary"],
    "ChequeAmount": ["cheque amount", "cheque", "check amount"],
    "IncomeTax": ["income tax", "tds", "incometax"],
    "Total PF": ["total pf", "pf", "provident fund", "totalpf"],
    "PT": ["professional tax", "pt"],
    "ESIC": ["esic", "esi"],
    "WorkOrderNo": ["work order no", "work order number", "workorder no", "workorder number", "work order",
                    "workorder", "wo no", "workorderno"],
    "Department": ["department", "departments", "dept", "client"],
    "Category": ["category", "categories"],
    "Month": ["salary month", "month", "period"],
    "month_year": ["month_year", "month year", "payment month"],
    "WOStartDate": ["wo start date", "work order start date", "start date", "wostartdate"],
    "WOEndDate": ["wo end date", "work order end date", "end date", "woenddate"],
    "PaidDays": ["paid days", "paiddays"],
    "Salary": ["salary rate", "monthly salary", "salary column"],
    "FinaluniqCode": ["final unique code", "finaluniqcode", "final uniq code"],
    "InvoiceDate": ["invoice date"],
    "Basic": ["basic salary", "basic"],
    "HRA": ["hra"],
    "Bonus": ["bonus"],
    "Arrears": ["arrears", "arrear"],
    "Remarks": ["remarks", "remark"],
    "joindate": ["join date", "joining date", "joindate"],
    "resigndate": ["resign date", "resignation date", "resigndate"],
}
LABELS = {   # how a figure is written in answer text (used to strip restricted figures from text)
    "CTC": [r"\bctc\b"], "BaseSalesValue": [r"\bbase\s+sales?\b"], "InvoiceAmount": [r"\binvoice\b"],
    "Margin": [r"\bmargin\b"], "MarginPct": [r"\bmargin\s*%"], "GrossAmount": [r"\bgross\b"],
    "TakeHome": [r"\btake[\s\-]?home\b"], "IncomeTax": [r"\btds\b", r"\bincome\s+tax\b"], "Total PF": [r"\bpf\b"],
    "PT": [r"\bpt\b"], "ESIC": [r"\besic\b"], "EmployeeCount": [r"\bemployees?\b(?!\s*code)", r"\bemployee\s+count\b"],
    "NetSalary": [r"\bnet\s+(salary|pay)\b"], "AvgCTC": [r"\baverage\s+ctc\b"], "Name": [], "Employee": [],
}
MONTH_SUFFIX = re.compile(r"\s*\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-]*\d{2,4}\b", re.I)
OUT_PREFIX = re.compile(r"^(total\s+projected|projected|total|monthly|annual|remaining|actual|avg\.?|sum\s+of)\s+", re.I)
OUT_SUFFIX = re.compile(r"\s*(\(.*?\)|change\s*%|change|%?\s*change|already\s+deducted)$", re.I)
OUT_EXACT = {  # result column names used by the app -> key
    "employees": "EmployeeCount", "employeecount": "EmployeeCount", "salarymonth": "Month", "period": "Month",
    "basesalesvalue": "BaseSalesValue", "basesales": "BaseSalesValue", "invoiceamount": "InvoiceAmount",
    "invoicewithgst": "InvoiceAmount", "invoice": "InvoiceAmount", "margin": "Margin", "marginpct": "MarginPct",
    "marginpctsales": "MarginPct", "margin%": "MarginPct", "projectedmargin%": "MarginPct", "takehome": "TakeHome",
    "totalpf": "Total PF", "averagectc": "AvgCTC", "avgctc": "AvgCTC", "wostartdate": "WOStartDate",
    "woenddate": "WOEndDate", "name": "Name", "employee": "Employee", "tds": "IncomeTax", "gross": "GrossAmount",
    "netpay": "NetSalary",
}


def norm(s):
    return re.sub(r"[^a-z0-9%]", "", str(s).lower())


class Vocab:
    """Turns words in a question into column keys ("gross salary" -> GrossAmount)."""

    def __init__(self, columns=None):
        self.columns = list(dict.fromkeys(list(columns or []) + KNOWN_COLUMNS))
        self.phrases = []                                 # (regex, key), longest phrase first
        for key, words in ALIASES.items():
            for w in words:
                self.phrases.append((w, key))
        generic = {"salary", "year", "month", "name", "category", "remarks", "employee", "department"}
        for c in self.columns:
            if c.lower() in generic:                      # everyday words - only their aliases count
                continue
            self.phrases.append((c.lower().replace("_", " "), c))
            self.phrases.append((c.lower(), c))
        self.phrases = sorted(set(self.phrases), key=lambda x: -len(x[0]))
        self.compiled = [(re.compile(r"(?<![a-z0-9])" + re.escape(p).replace(r"\ ", r"[\s_\-]*") + r"(?![a-z0-9])",
                                     re.I), k) for p, k in self.phrases]
        self.by_norm = {norm(c): c for c in self.columns}
        self.all_keys = list(dict.fromkeys(self.columns + list(DERIVED) + ["BankAccount"]))

    def resolve(self, text):
        """Keys mentioned in the text, in the order written."""
        text = str(text)
        taken = [False] * len(text)
        found = []
        for rx, key in self.compiled:
            for m in rx.finditer(text):
                if any(taken[m.start():m.end()]):
                    continue
                for i in range(m.start(), m.end()):
                    taken[i] = True
                found.append((m.start(), key))
        out = []
        for _, k in sorted(found):
            if k not in out:
                out.append(k)
        return out

    def out_key(self, colname):
        """Which key a RESULT column belongs to ('Projected CTC Sep-2026' -> CTC). None = calculated/unknown."""
        c = str(colname)
        n = norm(c)
        if n in OUT_EXACT:
            return OUT_EXACT[n]
        if n in self.by_norm:
            return self.by_norm[n]
        c2 = MONTH_SUFFIX.sub("", c).strip()
        for _ in range(2):
            c2 = OUT_SUFFIX.sub("", OUT_PREFIX.sub("", c2)).strip()
        n2 = norm(c2)
        if n2 in OUT_EXACT:
            return OUT_EXACT[n2]
        if n2 in self.by_norm:
            return self.by_norm[n2]
        keys = self.resolve(c2)
        return keys[0] if len(keys) == 1 else None

    def keys_in_name(self, colname):
        """Every key whose words appear in a column name (conservative check for restricted data)."""
        return set(self.resolve(MONTH_SUFFIX.sub(" ", str(colname))))


def label(key):
    names = {"GrossAmount": "Gross salary", "UniqueNo": "UniqueNo", "UniqueNo2": "UniqueNo2", "Employee": "Employee code",
             "Name": "Employee name", "BankAccount": "Bank account", "BaseSalesValue": "Base sales value",
             "InvoiceAmount": "Invoice amount", "MarginPct": "Margin %", "EmployeeCount": "Employee count",
             "AvgCTC": "Average CTC", "TakeHome": "Take-home", "IncomeTax": "TDS / Income tax", "Total PF": "PF",
             "WorkOrderNo": "Work order no", "Month": "Month", "month_year": "month_year"}
    return names.get(key, key)


# ----------------------------------------------------------------------------------------------
# Admin policies (policies.json) + policy audit log
# ----------------------------------------------------------------------------------------------
DEFAULT_GLOBAL = {"blocked_columns": [], "allowed_columns": [], "blocked_tables": [],
                  "sensitive_columns": list(SENSITIVE_DEFAULT), "admin_info_for_users": False}
DEFAULT_USER = {"blocked_columns": [], "allowed_columns": [], "aggregated_only": False, "sensitive_access": True,
                "payroll_access": True, "blocked_tables": [], "teach": "none",
                "row_categories": [], "row_departments": [],     # empty = all rows; else ONLY these values
                "display_rules": False}                          # may change column names / formats for everyone
TEACH_LEVELS = ("none", "self", "global")     # v14: teaching right per user (admins always "global")
AUDIT_FIELDS = ["time", "admin", "action", "scope", "target", "item", "old", "new", "reason"]


class PolicyStore:
    _lock = threading.RLock()

    def __init__(self, folder):
        self.path = Path(folder) / "policies.json"
        self.audit_path = Path(folder) / "policy_audit.csv"

    def load(self):
        with self._lock:
            try:
                data = json.loads(self.path.read_text("utf-8"))
            except FileNotFoundError:
                data = {}
            except Exception:
                data = {}
        data.setdefault("version", 1)
        g = data.setdefault("global", {})
        for k, v in DEFAULT_GLOBAL.items():
            g.setdefault(k, list(v) if isinstance(v, list) else v)
        data.setdefault("users", {})
        return data

    def save(self, data):
        with self._lock:
            data["version"] = int(data.get("version", 1)) + 1
            tmp = self.path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, indent=2), "utf-8")
            os.replace(tmp, self.path)

    def user(self, data, uid):
        u = dict(DEFAULT_USER)
        u.update({k: v for k, v in (data.get("users", {}).get(uid) or {}).items()})
        return u

    def log(self, admin, action, scope, target, item, old, new, reason=""):
        with self._lock:
            new_file = not self.audit_path.exists()
            with self.audit_path.open("a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if new_file:
                    w.writerow(AUDIT_FIELDS)
                w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), admin, action, scope, target,
                            item, json.dumps(old) if not isinstance(old, str) else old,
                            json.dumps(new) if not isinstance(new, str) else new, reason])

    def read_log(self, last=500):
        try:
            with self.audit_path.open(encoding="utf-8") as f:
                return list(csv.DictReader(f))[-last:]
        except FileNotFoundError:
            return []

    # -------- the only way policies change: an action, applied + logged under the lock --------
    def apply(self, action, admin, reason=""):
        """action: {"op", "scope": "GLOBAL"|"USER", "user", "columns", "value"} -> short description."""
        with self._lock:
            data = self.load()
            op, scope = action["op"], action.get("scope", "GLOBAL")
            target = "GLOBAL" if scope == "GLOBAL" else action["user"]
            if scope == "GLOBAL":
                pol = data["global"]
            else:
                pol = data["users"].setdefault(action["user"], {})
                for k, v in DEFAULT_USER.items():
                    pol.setdefault(k, list(v) if isinstance(v, list) else v)
            done = []
            if op in ("BLOCK_COLUMN", "UNBLOCK_COLUMN"):
                cur = list(pol.get("blocked_columns", []))
                for c in action["columns"]:
                    old = "BLOCKED" if c in cur else "ALLOWED"
                    if op == "BLOCK_COLUMN" and c not in cur:
                        cur.append(c)
                    if op == "UNBLOCK_COLUMN" and c in cur:
                        cur.remove(c)
                    new = "BLOCKED" if c in cur else "ALLOWED"
                    self.log(admin, op, scope, target, c, old, new, reason)
                    done.append(f"{label(c)}: {old} → {new}")
                pol["blocked_columns"] = cur
                if op == "UNBLOCK_COLUMN" and scope == "GLOBAL":   # also lift a "sensitive" listing? no - separate
                    pass
            elif op == "SET_ALLOWED":
                old = list(pol.get("allowed_columns", []))
                pol["allowed_columns"] = list(action["columns"])
                self.log(admin, op, scope, target, "allowed_columns", old, pol["allowed_columns"], reason)
                done.append("Allowed columns: " + (", ".join(map(label, action["columns"])) or "ALL"))
            elif op in ("SET_AGG_ONLY", "SET_SENSITIVE", "SET_PAYROLL", "SET_ADMIN_INFO", "SET_DISPLAY"):
                key = {"SET_AGG_ONLY": "aggregated_only", "SET_SENSITIVE": "sensitive_access",
                       "SET_PAYROLL": "payroll_access", "SET_ADMIN_INFO": "admin_info_for_users",
                       "SET_DISPLAY": "display_rules"}[op]
                old = pol.get(key)
                pol[key] = bool(action["value"])
                self.log(admin, op, scope, target, key, old, pol[key], reason)
                done.append(f"{key}: {old} → {pol[key]}")
            elif op == "SET_SENSITIVE_LIST":
                old = list(pol.get("sensitive_columns", []))
                pol["sensitive_columns"] = list(action["columns"])
                self.log(admin, op, scope, target, "sensitive_columns", old, pol["sensitive_columns"], reason)
                done.append("Sensitive columns updated")
            elif op == "SET_TABLES":
                old = list(pol.get("blocked_tables", []))
                pol["blocked_tables"] = list(action["columns"])
                self.log(admin, op, scope, target, "blocked_tables", old, pol["blocked_tables"], reason)
                done.append("Blocked tables: " + (", ".join(action["columns"]) or "none"))
            elif op == "SET_TEACH":
                if action["value"] not in TEACH_LEVELS:
                    raise ValueError("Teaching right must be none, self or global")
                old = pol.get("teach", "none")
                pol["teach"] = action["value"]
                self.log(admin, op, scope, target, "teach", old, pol["teach"], reason)
                done.append(f"Teaching right: {old} → {pol['teach']}")
            elif op == "SET_ROWS":                       # limit a user to some categories / departments
                key = {"Category": "row_categories", "Department": "row_departments"}[action["column"]]
                old = list(pol.get(key, []))
                pol[key] = sorted({str(v) for v in action["columns"]})
                self.log(admin, op, scope, target, key, old, pol[key], reason)
                done.append(f"{action['column']} limited to: " + (", ".join(pol[key]) or "ALL"))
            elif op == "RESET_USER_POLICY":
                old = data["users"].pop(action["user"], None)
                self.log(admin, op, scope, target, "all", old or {}, "DEFAULT", reason)
                done.append("Back to default access")
            else:
                raise ValueError(f"Unknown policy action {op}")
            self.save(data)
            return "; ".join(done)


# ----------------------------------------------------------------------------------------------
# Effective policy = levels 1-3 combined for one user
# ----------------------------------------------------------------------------------------------
class Effective:
    def __init__(self, data, uid, role, vocab, table="payrollall", store=None, taught=None, taught_hidden=None):
        """taught = {key: {"sources": [...], "label": "..."}} for the user's IN-SCOPE taught formulas (v14);
        taught_hidden = keys of taught formulas that are NOT in this user's scope (never visible)."""
        g = data.get("global", DEFAULT_GLOBAL)
        u = (store.user(data, uid) if store else dict(DEFAULT_USER, **(data.get("users", {}).get(uid) or {})))
        self.uid, self.role, self.vocab = uid, role, vocab
        self.blocked = set(g.get("blocked_columns", [])) | set(u.get("blocked_columns", []))
        if not u.get("sensitive_access", True):
            self.blocked |= set(g.get("sensitive_columns", SENSITIVE_DEFAULT))
        allowed = [set(x) for x in (g.get("allowed_columns") or [], u.get("allowed_columns") or []) if x]
        self.allowed = set.intersection(*allowed) if allowed else None      # None = every column
        self.aggregated_only = bool(u.get("aggregated_only", False))
        self.row_filter = {c: list(u.get(k) or []) for c, k in (("Category", "row_categories"),
                                                                ("Department", "row_departments")) if u.get(k)}
        tables = set(g.get("blocked_tables", [])) | set(u.get("blocked_tables", []))
        self.payroll_access = bool(u.get("payroll_access", True)) and table.lower() not in {t.lower() for t in tables}
        self.admin_info = role == "admin" or bool(g.get("admin_info_for_users", False))
        self.version = data.get("version", 1)
        t = str(u.get("teach", "none") or "none")
        self.teach = "global" if role == "admin" else (t if t in TEACH_LEVELS else "none")
        # a taught formula is just another calculated value: hidden when ANY of its source columns is hidden
        self.derived = dict(DERIVED)
        self.taught_labels = {}
        for k, info in (taught or {}).items():
            self.derived[k] = list(info.get("sources", []))
            self.taught_labels[norm(k)] = k
            if info.get("label"):
                self.taught_labels[norm(info["label"])] = k
        self.taught_hidden = set(taught_hidden or ()) - set(taught or {})

    def _key_of(self, colname):
        n = norm(MONTH_SUFFIX.sub("", str(colname)))
        if n in self.taught_labels:
            return self.taught_labels[n]
        return self.vocab.out_key(colname)

    # ---- is a key (column or calculated value) visible? ----
    def key_visible(self, key):
        if key is None:
            return True
        if key in self.blocked or key in self.taught_hidden:
            return False
        src = self.derived.get(key)
        if src and any(s in self.blocked for s in src):
            return False
        if self.allowed is not None:
            if key in self.allowed:
                return True
            if src:
                return all(s in self.allowed for s in src)
            if key == "EmployeeCount":
                return "Employee" in self.allowed
            return False
        return True

    def col_visible(self, colname):
        key = self._key_of(colname)
        if not self.key_visible(key):
            return False
        # conservative: a restricted word anywhere in the column name hides it ("Annual gross", "Remaining TDS")
        for k in self.vocab.keys_in_name(colname):
            if k in self.blocked or (self.derived.get(k) and any(s in self.blocked for s in self.derived[k])):
                return False
        return True

    # ---- SQL side ----
    def sql_null_columns(self, columns):
        """Table columns the policy view returns as NULL for this user."""
        needed = set()
        for k, src in self.derived.items():
            if self.key_visible(k):
                needed |= set(src)
        out = set()
        for c in columns:
            if c in TECHNICAL:
                continue                                  # kept for correct maths, removed from output later
            if c in self.blocked:
                out.add(c)
            elif self.allowed is not None and c not in self.allowed and c not in needed:
                out.add(c)
            elif not self.key_visible(c) and c not in needed:
                out.add(c)
        return out

    def sql_violations(self, sql):
        """Restricted columns referenced by an (AI-written) query - checked BEFORE it runs."""
        s = re.sub(r"'[^']*'", "''", str(sql))
        hits = set()
        names = set(re.findall(r"\[([^\]]+)\]", s)) | set(re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\b", s))
        pview = {"BaseSales": "BaseSalesValue", "TakeHome": "TakeHome"}
        for n in names:
            key = pview.get(n) or self.taught_labels.get(norm(n)) or self.vocab.by_norm.get(norm(n))
            if key and key not in TECHNICAL and not self.key_visible(key):
                hits.add(key)
        return sorted(hits)

    def restricted_keys(self):
        return {k for k in self.vocab.all_keys if not self.key_visible(k)}

    def summary(self):
        return {"blocked": sorted(self.blocked), "allowed": sorted(self.allowed) if self.allowed is not None else "ALL",
                "aggregated_only": self.aggregated_only, "payroll_access": self.payroll_access,
                "rows": self.row_filter or "ALL"}


# ----------------------------------------------------------------------------------------------
# User preferences (level 4)
# ----------------------------------------------------------------------------------------------
DEFAULT_PREFS = {"hidden_columns": [], "preferred_columns": [], "always_columns": [], "raw_data": True,
                 "charts": True, "default_group": None}


def prefs_of(rec):
    p = dict(DEFAULT_PREFS)
    p.update((rec or {}).get("data_prefs") or {})
    for k in ("hidden_columns", "preferred_columns", "always_columns"):
        p[k] = list(p.get(k) or [])
    return p


RAW_ASK = re.compile(r"\b(employee[\s\-]*(wise|level|list|details?|names?|codes?)|each\s+employee|per\s+employee|"
                     r"individual|raw|list\s+(of\s+)?(all\s+)?employees|employees?\s+list|employee\s+\d+|who\s+(is|are|has|have)|"
                     r"top\s+\d+\s+employees?|employees?\s+(whose|with|who)|names?|tds\s+of|salary\s+of)\b", re.I)
NEG_BEFORE = re.compile(r"\b(without|except|excluding|exclude|hide|remove|skip|don'?t\s+(show|include|give)|"
                        r"do\s+not\s+(show|include|give)|no)\s+(the\s+)?$", re.I)


ONCE_NEG = re.compile(r"(without|except|excluding|exclude|hide|remove|skip|minus|don'?t\s+(show|include|give)|"
                      r"do\s+not\s+(show|include|give)|no)\s+(the\s+)?((?:[\w%\-]+\s*(?:,|and|&)?\s*){0,6})$", re.I)


def one_time_hidden(question, vocab):
    """'... without UniqueNo and CTC' / "don't show CTC in this answer" -> hidden in THIS answer only."""
    out = set()
    q = str(question)
    for rx, key in vocab.compiled:
        for m in rx.finditer(q):
            m2 = ONCE_NEG.search(q[max(0, m.start() - 70):m.start()])
            if m2 and all(vocab.resolve(w) or w.lower() in ("and", "the", "&", ",") for w in
                          re.findall(r"[\w%\-]+", m2.group(5) or "")):
                out.add(key)
    return out


def strip_one_time(question, vocab):
    """Removes one-time wishes ('without CTC', "don't show UniqueNo in this answer") from the text used to
    pick the report - they are applied to the answer by secure_df."""
    q = str(question)
    once = one_time_hidden(q, vocab)
    if not once:
        return q
    for rx, key in vocab.compiled:
        if key in once:
            q = rx.sub(" ", q)
    q = re.sub(r"\b(without|except|excluding|exclude|hide|remove|skip|minus|don'?t\s+(show|include|give)|"
               r"do\s+not\s+(show|include|give))\b(\s*(the|,|and|&))*\s*(?=[.,;]|\bin\s+this|$|\s+(for|in|of)\b)",
               " ", q, flags=re.I)
    q = ONE_TIME.sub(" ", q)
    q = re.sub(r"^\s*[.,;:]+|\s+[.,;:]+\s*$", " ", q)
    return re.sub(r"\s+", " ", q).strip(" .,;") or str(question)


# ----------------------------------------------------------------------------------------------
# Result filter (enforcement point 2) - permissions first, then preferences
# ----------------------------------------------------------------------------------------------
def secure_df(df, eff, prefs, question="", report="", explicit_cols=None):
    """Returns (df or None, notes, block_message). Never trusts the LLM or the SQL to have done this."""
    notes = []
    if df is None:
        return None, notes, None
    if not eff.payroll_access:
        return None, notes, "🔒 Your login does not have access to payroll data. Please contact your administrator."
    vocab = eff.vocab
    asked = set(vocab.resolve(question))
    cols = list(df.columns)
    keys = {c: vocab.out_key(c) for c in cols}
    emp_level = report in EMPLOYEE_REPORTS or any(keys[c] in EMPLOYEE_KEYS for c in cols)
    # --- levels 2/3: raw (employee-level) data ---
    if emp_level and eff.aggregated_only:
        return None, notes, ("🔒 Your administrator allows only **summary (aggregated) data** for your login, so "
                             "employee-level rows can't be shown. Try a work order-wise or department-wise total.")
    # --- levels 2/3: columns ---
    drop, refused = [], []
    for c in cols:
        if not eff.col_visible(c):
            drop.append(c)
            k = keys[c] or next(iter(vocab.keys_in_name(c)), None)
            if k and (k in asked or any(s in asked for s in DERIVED.get(k, []))):
                refused.append(label(k))
    for k in asked:                         # asked for something that is not in the result because it is blocked
        if not eff.key_visible(k) and label(k) not in refused:
            refused.append(label(k))
    if refused:
        notes.append(f"🔒 Not available for your login: **{', '.join(sorted(set(refused)))}** "
                     "(restricted by your administrator).")
    df = df.drop(columns=drop)
    # --- level 4: raw-data preference ---
    if emp_level and prefs.get("raw_data") is False and not RAW_ASK.search(question or ""):
        return None, notes, ("You asked me not to show **employee-level data unless you ask for it**. "
                             "Ask e.g. *employee wise ctc for aug-26*, or change this in Settings → My data preferences.")
    # --- level 5 one-time hides, level 4 hidden columns (the current request can show a hidden column) ---
    once = one_time_hidden(question or "", vocab)
    hidden = set(prefs.get("hidden_columns", [])) - set(prefs.get("always_columns", []))
    drop2 = []
    for c in df.columns:
        k = keys.get(c) or vocab.out_key(c)
        if k in once:
            drop2.append(c)
        elif k in hidden and k not in (asked - once):
            drop2.append(c)
    if drop2:
        df = df.drop(columns=drop2)
    # --- level 4: "only these columns" (unless this question lists its own columns) ---
    pref_cols = list(prefs.get("preferred_columns", []))
    explicit = explicit_cols if explicit_cols is not None else (
        len([k for k in asked if k not in ("Month", "WorkOrderNo", "Department", "Category")]) >= 2 or
        bool(re.search(r"\bcolumns?\b|\bfields?\b", question or "", re.I)))
    if pref_cols and not explicit and len(df.columns):
        want = pref_cols + [k for k in prefs.get("always_columns", []) if k not in pref_cols]
        keep = [c for k in want for c in df.columns if (keys.get(c) or vocab.out_key(c)) == k]
        keep = list(dict.fromkeys(keep))
        if keep:
            df = df[keep]
            notes.append("Columns from **your saved column list** (change it in Settings → My data preferences).")
    if len(df.columns) == 0:
        return None, notes, "🔒 Nothing in this answer can be shown with your current access."
    return df, notes, None


def secure_text(text, eff, prefs=None, question=""):
    """Removes figures of restricted (or preference-hidden) values from answer text."""
    if not text:
        return text
    vocab = eff.vocab
    asked = set(vocab.resolve(question))
    hide = {k for k in LABELS if not eff.key_visible(k)}
    if prefs:
        hide |= {k for k in prefs.get("hidden_columns", []) if k in LABELS and k not in asked}
    if not hide:
        return text
    pats = [re.compile(p, re.I) for k in hide for p in LABELS.get(k, [])]
    if not pats:
        return text
    out = []
    for line in str(text).split("\n"):
        parts = re.split(r"(\s·\s)", line)
        keep = [p for p in parts[::2] if not any(px.search(p) for px in pats)]
        new = " · ".join(x.strip() for x in keep if x.strip())
        if new.count("**") % 2:
            new = new.replace("**", "")
        if line.strip() and not new.strip():
            continue
        out.append(new if line.strip() else line)
    res = "\n".join(out).strip()
    return res or "Here is the result (some figures are hidden for your login)."


def redact_sql(sql, raw_table="[payrollall]"):
    """Shows the policy view as the plain table name (users don't need to see restricted column names)."""
    rt = re.escape(raw_table)
    return re.sub(r"\(SELECT\s+(?:[^()]|\([^()]*\))*?FROM\s+" + rt + r"(?:\s+WHERE\s+(?:[^()]|\([^()]*\))*?)?\)\s+AS\s+" + rt,
                  raw_table, str(sql or ""), flags=re.I | re.S)


# ----------------------------------------------------------------------------------------------
# Understanding instructions (natural language -> structured actions)
# ----------------------------------------------------------------------------------------------
PERSIST = re.compile(r"\b(from\s+now(\s+on(wards)?)?|always|never|by\s+default|every\s*time|permanently|in\s+(the\s+)?future|"
                     r"going\s+forward|here\s*after|aage\s+se|hamesha|forever|all\s+the\s+time|from\s+today|"
                     r"in\s+all\s+(my\s+)?(answers|results|reports)|remember\s+(that|to|this))\b", re.I)
ONE_TIME = re.compile(r"\b(in\s+this\s+(answer|reply|result|report|table|one)|this\s+time|for\s+now|only\s+for\s+this|"
                      r"just\s+this\s+(once|time)|right\s+now|for\s+this\s+(one|question|answer))\b", re.I)
HIDE_RX = re.compile(r"\b(never|don'?t|dont|do\s+not|stop|no\s+need\s+to|not)\s+(?:to\s+)?(show|showing|give|giving|include|"
                     r"including|display|displaying|put|send|provide|add|return)\b|\b(hide|remove|exclude|skip|drop)\b", re.I)
SHOW_RX = re.compile(r"\b(un-?hide|stop\s+hiding|bring\s+back|start\s+showing|show\s+.{1,60}?again|include\s+.{1,60}?again|"
                     r"i\s+want\s+.{1,60}?again|give\s+.{1,60}?again|always\s+(show|include|give|display)|"
                     r"you\s+can\s+(show|include|give)|ok\s+to\s+show|allow\s+me\s+to\s+see)\b", re.I)
ONLY_RX = re.compile(r"\bonly\s+(give|show|include|display|return|want|need)|\b(give|show|display|return)\s+(me\s+)?only\b|"
                     r"\bonly\s+these\s+columns\b|\bjust\s+(give|show)\s+me\b|\bonly\s*:", re.I)
CLEAR_ONLY = re.compile(r"\b(all\s+columns\s+again|stop\s+(using\s+)?(only|my)\s+(these\s+)?columns|clear\s+(my\s+)?(column\s+list|"
                        r"preferred\s+columns)|forget\s+(the\s+)?column\s+list|normal\s+columns\s+again)\b", re.I)
RESET_RX = re.compile(r"\b(reset|clear|forget|delete|remove|wipe)\s+(all\s+)?(of\s+)?(my\s+)?(data\s+)?(preferences|memory|"
                      r"saved\s+settings|instructions)\b", re.I)
SHOWPREF_RX = re.compile(r"\b(what|show|list|tell|see|view|display)\b.{0,30}\bmy\s+(data\s+)?(preferences|memory|saved\s+"
                         r"settings|instructions)\b|\bwhat\s+(do\s+you|did\s+you)\s+remember\b|\bmy\s+preferences\??$", re.I)
RAW_RX = re.compile(r"\b(raw|employee[\s\-]*(level|wise)|individual|row[\s\-]*level|person[\s\-]*wise|per\s+employee|"
                    r"employee\s+data)\b", re.I)
CHART_RX = re.compile(r"\b(charts?|graphs?)\b", re.I)


# a question that asks for data of a month ("... in sep-26") is a data question, never a settings change
DATA_PERIOD = re.compile(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sept?|oct|nov|dec)[a-z]*[\s\-/'.]*\d{2,4}\b|"
                         r"\b\d{1,2}[/\-]\d{2,4}\b|\b20\d\d[/\-]\d{1,2}\b", re.I)
# the words right after "don't show / hide / show again" are the columns meant; stop at the next clause
OBJ_END = re.compile(r"\s(?:in|into|to|for|of|as|from|unless|because|when|while|also|then)\s|[.;!?\"'()]|\band\s+(?:do|don'?t|"
                     r"also|from|give|show|make|never|always)\b", re.I)


def _cols_after(rx, q, vocab):
    """Columns named right after a hide/show verb ('never show CTC and PF' -> CTC, PF), not anywhere in the text."""
    out = []
    for m in rx.finditer(q):
        rest = q[m.end():m.end() + 120]
        stop = OBJ_END.search(rest)
        for k in vocab.resolve(rest[:stop.start()] if stop else rest):
            if k not in out:
                out.append(k)
    return out


def parse_pref_command(question, vocab):
    """Recognises PERSISTENT instructions only. Returns an action dict or None (None = normal question)."""
    q = str(question).strip()
    if RESET_RX.search(q):
        return {"type": "reset"}
    if SHOWPREF_RX.search(q):
        return {"type": "show_prefs"}
    persistent = bool(PERSIST.search(q)) or bool(SHOW_RX.search(q)) or bool(CLEAR_ONLY.search(q))
    if not persistent or (ONE_TIME.search(q) and not PERSIST.search(q)):
        return None                                       # one-time wishes are handled in the answer itself
    if DATA_PERIOD.search(q):
        return None                                       # "give me ... for sep-26 ... from now on ..." = a data question
    if CLEAR_ONLY.search(q):
        return {"type": "clear_only"}
    hide = bool(HIDE_RX.search(q))
    show = bool(SHOW_RX.search(q))
    if RAW_RX.search(q) and (hide or show or re.search(r"\bunless\b", q, re.I)):
        return {"type": "raw", "value": bool(show and not hide)}
    if CHART_RX.search(q) and (hide or show):
        return {"type": "charts", "value": bool(show and not hide)}
    if not vocab.resolve(q):
        return None
    if ONLY_RX.search(q):
        cols = _cols_after(ONLY_RX, q, vocab) or vocab.resolve(q)
        return {"type": "only", "cols": cols}
    if show and not (hide and not re.search(r"again|un-?hide|stop\s+hiding", q, re.I)):
        cols = _cols_after(SHOW_RX, q, vocab) or vocab.resolve(q)
        return {"type": "show", "cols": cols, "always": bool(re.search(r"\balways\s+(show|include|give)", q, re.I))}
    if hide:
        cols = _cols_after(HIDE_RX, q, vocab)
        return {"type": "hide", "cols": cols} if cols else None
    return None


def apply_pref(prefs, action):
    """Returns (new prefs, message). Pure - the caller saves it for THIS user only."""
    p = {k: (list(v) if isinstance(v, list) else v) for k, v in prefs.items()}
    t = action["type"]
    if t == "reset":
        return dict(DEFAULT_PREFS, hidden_columns=[], preferred_columns=[], always_columns=[]), \
            "✅ Your data preferences are reset to normal."
    if t == "clear_only":
        p["preferred_columns"] = []
        return p, "✅ Done - I'll use the normal columns again."
    if t == "raw":
        p["raw_data"] = action["value"]
        return p, ("✅ Saved. I'll show employee-level data when a question needs it." if action["value"] else
                   "✅ Saved. From now on I **won't show employee-level (raw) data unless you ask for it** "
                   "(e.g. *employee wise …*).")
    if t == "charts":
        p["charts"] = action["value"]
        return p, "✅ Saved. " + ("Charts are on." if action["value"] else "From now on I won't show charts.")
    cols = action.get("cols", [])
    names = ", ".join(label(c) for c in cols)
    if t == "hide":
        p["hidden_columns"] = list(dict.fromkeys(p["hidden_columns"] + cols))
        p["always_columns"] = [c for c in p["always_columns"] if c not in cols]
        p["preferred_columns"] = [c for c in p["preferred_columns"] if c not in cols]
        return p, f"✅ Saved. From now on I **won't show {names}** in your answers (unless you ask for it by name)."
    if t == "show":
        p["hidden_columns"] = [c for c in p["hidden_columns"] if c not in cols]
        if action.get("always"):
            p["always_columns"] = list(dict.fromkeys(p["always_columns"] + cols))
        return p, f"✅ Saved. I'll show **{names}** again" + (" in every answer where it fits." if action.get("always") else ".")
    if t == "only":
        p["preferred_columns"] = cols
        p["hidden_columns"] = [c for c in p["hidden_columns"] if c not in cols]
        return p, (f"✅ Saved. From now on I'll give you **only these columns**: {names}. "
                   "Say *show all columns again* to go back to normal.")
    return p, ""


# ---------------------------------------- admin ----------------------------------------
ADMIN_BLOCK = re.compile(r"\b(block|forbid|restrict|deny|disallow|ban|lock|revoke)\b|\b(never|don'?t|dont|do\s+not|not)\s+"
                         r"(allow|let|give|show|permit)\b|\bnot\s+allowed\b|\bcan(no|')?t\s+see\b|\bcannot\s+see\b|"
                         r"\bshould\s+not\s+see\b", re.I)
ADMIN_UNBLOCK = re.compile(r"\b(unblock|unrestrict|unlock|lift|remove\s+(the\s+)?(restriction|block)s?(\s+on)?|allow\s+.{0,40}?again|"
                           r"re-?allow|give\s+.{0,40}?access\s+(again|back)|restore)\b", re.I)
ADMIN_ALLOW = re.compile(r"\b(allow|permit|let|grant)\b", re.I)
GLOBAL_RX = re.compile(r"\b(every\s*one|every\s*body|all\s+(the\s+)?users|every\s+user|any\s+users?|anyone|any\s+one|"
                       r"globally|global|for\s+all|company[\s\-]*wide|all\s+logins)\b", re.I)
AGG_RX = re.compile(r"\b(aggregat\w*|summary|summari[sz]ed|totals?)\b", re.I)
ONLY_SEE_RX = re.compile(r"\b(can|may|should|will)\s+only\s+(see|access|view|get)\b|\ballow\s+only\b|\bonly\s+allow\b|"
                         r"\bonly\s+(see|access)\b", re.I)
SENS_RX = re.compile(r"\bsensitive\b", re.I)
PAYROLL_RX = re.compile(r"\bpayroll\s+(data\s+)?access\b|\baccess\s+to\s+payroll\b|\baccess\s+to\s+(the\s+)?(chatbot|data)\b", re.I)
ADMIN_INFO_RX = re.compile(r"\b(admin|user)\s+(info|information|details)\b", re.I)
CONFIRM_RX = re.compile(r"^\s*(yes[,\s]*)?(confirm(ed)?|yes|ok(ay)?\s+confirm|go\s+ahead|do\s+it|apply)\s*[.!]*\s*$", re.I)
CANCEL_RX = re.compile(r"^\s*(cancel|no|stop|don'?t|abort|never\s*mind)\s*[.!]*\s*$", re.I)


def find_users(question, users):
    """User IDs named in the text (ID, full name or first name)."""
    q = " " + str(question).lower() + " "
    out = []
    for uid, rec in users.items():
        names = {uid.lower(), str(rec.get("name", "")).lower()}
        first = str(rec.get("name", "")).split()[:1]
        if first and len(first[0]) >= 3:
            names.add(first[0].lower())
        if any(n and re.search(r"(?<![a-z0-9._])" + re.escape(n) + r"(?![a-z0-9])", q) for n in names):
            out.append(uid)
    return out


TEACH_RX = re.compile(r"\bteach(ing)?\s*(right|rights|permission|access)\b|\bright\s+to\s+teach\b|"
                      r"\b(allow|let)\s+.{1,40}?\s+to\s+teach\b", re.I)
TEACH_REMOVE_RX = re.compile(r"\b(remove|revoke|take\s+away|withdraw|cancel|stop|disable|no\s+more|don'?t\s+allow|"
                             r"do\s+not\s+allow|block)\b", re.I)


def parse_teach_command(question, users):
    """'give teaching right to user1' / 'give global teaching right to user1' / 'remove teaching right from user1'."""
    q = str(question)
    if not TEACH_RX.search(q):
        return None
    who = find_users(q, users)
    if not who:
        m = re.search(r"\b(?:to|from|for)\s+(?:user\s+)?([a-z0-9._\-]+)\s*[.!?]*$", q, re.I)
        name = m.group(1) if m else "that user"
        return {"error": f"I couldn't find a user called **{name}**. Check the ID in Settings → Manage users."}
    if TEACH_REMOVE_RX.search(q):
        value = "none"
    elif re.search(r"\b(global|for\s+everyone|company[\s\-]*wide|all\s+users)\b", q, re.I):
        value = "global"
    else:
        value = "self"
    return {"op": "SET_TEACH", "scope": "USER", "users": who, "value": value, "question": q}


def parse_admin_command(question, vocab, users):
    """Recognises an ADMIN policy change. Returns {"op", ...} or {"error": ...} or None (not a policy command)."""
    q = str(question)
    t = parse_teach_command(q, users)
    if t is not None:
        return t
    if not (ADMIN_BLOCK.search(q) or ADMIN_UNBLOCK.search(q) or ONLY_SEE_RX.search(q) or
            (ADMIN_ALLOW.search(q) and (RAW_RX.search(q) or SENS_RX.search(q) or PAYROLL_RX.search(q)))):
        return None
    glob = bool(GLOBAL_RX.search(q))
    who = find_users(q, users)
    if not glob and not who:
        if re.search(r"\b(user|login|person)\s+[a-z0-9._\-]+\b", q, re.I) and (ADMIN_BLOCK.search(q) or ONLY_SEE_RX.search(q)):
            m = re.search(r"\b(?:user|login|person)\s+([a-z0-9._\-]+)\b", q, re.I)
            return {"error": f"I couldn't find a user called **{m.group(1)}**. Check the ID in Settings → Manage users."}
        return None
    scope = "GLOBAL" if glob and not who else "USER"
    cols = [c for c in vocab.resolve(q) if c not in ("Month",)]
    base = {"scope": scope, "users": who if scope == "USER" else [], "question": q}
    unblock = bool(ADMIN_UNBLOCK.search(q)) or (ADMIN_ALLOW.search(q) and not ADMIN_BLOCK.search(q) and not ONLY_SEE_RX.search(q))
    if AGG_RX.search(q) and (ONLY_SEE_RX.search(q) or re.search(r"\bonly\b", q, re.I)) and scope == "USER":
        return dict(base, op="SET_AGG_ONLY", value=True)
    if RAW_RX.search(q) and scope == "USER":
        return dict(base, op="SET_AGG_ONLY", value=not unblock)
    if SENS_RX.search(q):
        return dict(base, op="SET_SENSITIVE", value=bool(unblock), scope="USER") if who else None
    if PAYROLL_RX.search(q) and who:
        return dict(base, op="SET_PAYROLL", value=bool(unblock))
    if ONLY_SEE_RX.search(q) and cols:
        return dict(base, op="SET_ALLOWED", columns=cols)
    if not cols:
        return None
    return dict(base, op="UNBLOCK_COLUMN" if unblock else "BLOCK_COLUMN", columns=cols)


def describe_action(action, users):
    op = action["op"]
    tgt = "ALL users (100%)" if action["scope"] == "GLOBAL" else ", ".join(
        f"{u} ({users.get(u, {}).get('name', u)})" for u in action["users"])
    n = len(users) or 1
    pct = "100%" if action["scope"] == "GLOBAL" else f"{len(action['users']) * 100 // n}%"
    what = {"BLOCK_COLUMN": "BLOCK", "UNBLOCK_COLUMN": "ALLOW (remove restriction)", "SET_ALLOWED": "ALLOW ONLY these columns",
            "SET_AGG_ONLY": "Aggregated (summary) data ONLY" if action.get("value") else "Allow employee-level data",
            "SET_SENSITIVE": "Allow sensitive data" if action.get("value") else "BLOCK sensitive data",
            "SET_PAYROLL": "Give payroll access" if action.get("value") else "REMOVE payroll access",
            "SET_TEACH": {"none": "REMOVE teaching right", "self": "Give teaching right (own teachings only)",
                          "global": "Give GLOBAL teaching right (teachings for everyone)"}.get(action.get("value"), "")}[op]
    cols = ", ".join(label(c) for c in action.get("columns", [])) or "-"
    return (f"You're about to change data permissions:\n\n- **{'User' if action['scope'] == 'USER' else 'Scope'}:** {tgt}\n"
            f"- **Column(s):** {cols}\n- **Action:** {what}\n- **Affected users:** {pct}\n\n"
            "This is a security change. **Confirm?**")


ADMIN_Q = re.compile(r"\b(how\s+many|number\s+of|count\s+of|list|show|which|who|what|tell|give)\b.{0,40}\b(users?|logins?|"
                     r"admins?|administrators?|accounts?\s+in\s+(the\s+)?(chatbot|system))\b|"
                     r"\b(access\s+polic\w*|data\s+polic\w*|polic(y|ies)\s+(for|of)|restricted\s+columns?|globally\s+restricted|"
                     r"custom\s+restrictions?|policy\s+audit|audit\s+log|activity\s+log|user\s+permissions?|"
                     r"who\s+(has|have|is|are)\s+admin\w*|admin\s+access)\b", re.I)


def is_admin_question(question):
    return bool(ADMIN_Q.search(str(question)))


def admin_answer(question, users, data, store, vocab, logins=None):
    """Answers user-management questions. Returns (text, rows: list of dicts)."""
    q = str(question).lower()
    act = {u: r for u, r in users.items() if r.get("active", True)}
    rows = []
    def eff(u):
        return Effective(data, u, "admin" if users[u].get("role") == "master" else "user", vocab, store=store)
    if re.search(r"\bcolumns?\b.*\b(globally|global|everyone|all\s+users)\b|\bglobally\s+restricted\b|\bglobal\s+(block|restrict)", q):
        g = data["global"]
        txt = ("**Globally restricted columns:** " + (", ".join(label(c) for c in g["blocked_columns"]) or "none") +
               ("\n\n**Global allowed list:** " + ", ".join(label(c) for c in g["allowed_columns"]) if g["allowed_columns"] else ""))
        return txt, []
    m = re.search(r"\b(cannot|can'?t|can\s+not|not\s+allowed\s+to|blocked\s+from|restricted\s+from)\s+(see|view|access)?\s*(.+)", q)
    if m and vocab.resolve(m.group(3)):
        k = vocab.resolve(m.group(3))[0]
        who = [u for u in users if not eff(u).key_visible(k)]
        return (f"**Users who cannot see {label(k)}:** " + (", ".join(who) or "none"),
                [{"ID": u, "Name": users[u].get("name"), "Role": users[u].get("role")} for u in who])
    m = re.search(r"\b(can|who\s+can)\s+(see|view|access)\s+(.+)", q)
    if m and vocab.resolve(m.group(3)) and "only" not in q:
        k = vocab.resolve(m.group(3))[0]
        who = [u for u in users if eff(u).key_visible(k)]
        return f"**Users who can see {label(k)}:** " + (", ".join(who) or "none"), []
    named = find_users(q, users)
    if named and re.search(r"\b(polic\w*|access|permissions?|restrictions?|preferences?)\b", q):
        for u in named:
            e = eff(u)
            up = store.user(data, u)
            rows.append({"ID": u, "Name": users[u].get("name"), "Role": users[u].get("role"),
                         "Active": users[u].get("active", True), "Payroll access": e.payroll_access,
                         "Aggregated only": e.aggregated_only, "Sensitive data": up.get("sensitive_access", True),
                         "Blocked columns": ", ".join(label(c) for c in sorted(e.blocked)) or "-",
                         "Allowed columns": "ALL" if e.allowed is None else ", ".join(label(c) for c in sorted(e.allowed)),
                         "Own hidden columns": ", ".join(label(c) for c in prefs_of(users[u])["hidden_columns"]) or "-"})
        return f"**Access policy for {', '.join(named)}** (admin global rules + user rules + their own preferences):", rows
    if re.search(r"\bcustom\s+restrictions?\b|\buser[\s\-]*specific\b", q):
        who = [u for u in data.get("users", {}) if u in users]
        return "**Users with custom restrictions:** " + (", ".join(who) or "none"), \
            [{"ID": u, **{k: (", ".join(v) if isinstance(v, list) else v) for k, v in data["users"][u].items()}} for u in who]
    if re.search(r"\badmin\w*\b", q) and re.search(r"\bwho|which|list|how\s+many\b", q):
        ad = [f"{u} ({r.get('role')})" for u, r in users.items() if r.get("role") in ("admin", "master")]
        return f"**{len(ad)} admin(s):** " + ", ".join(ad), []
    if re.search(r"\bpayroll\s+access\b|\bcan\s+use\b", q):
        who = [u for u in act if eff(u).payroll_access]
        return f"**{len(who)} active user(s) with payroll access:** " + ", ".join(who), []
    if re.search(r"\b(currently\s+)?(active|online|logged\s*in)\b", q):
        recent = sorted(set(logins or []))
        txt = f"**{len(act)} active (enabled) user(s)** out of {len(users)}: " + ", ".join(act)
        if logins is not None:
            txt += f"\n\n**Logged in during the last hour:** " + (", ".join(recent) or "nobody")
        return txt, []
    if re.search(r"\bhow\s+many\b|\bnumber\s+of\b|\bcount\b", q):
        ad = sum(1 for r in users.values() if r.get("role") in ("admin", "master"))
        return (f"You have **{len(users)} user(s)**: {len(act)} active, {len(users) - len(act)} disabled · "
                f"{ad} admin(s), {len(users) - ad} normal user(s)."), []
    for u, r in users.items():
        e = eff(u)
        rows.append({"ID": u, "Name": r.get("name"), "Role": r.get("role"), "Active": r.get("active", True),
                     "Payroll access": e.payroll_access, "Aggregated only": e.aggregated_only,
                     "Blocked columns": ", ".join(label(c) for c in sorted(e.blocked)) or "-"})
    return f"**{len(users)} user(s)** - roles, status and data access:", rows
