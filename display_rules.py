"""
Display rules for the Payroll chatbot (pure Python, testable on its own).

How columns LOOK in every answer, for everyone: "don't treat Employee as a number", "no commas in employee code",
"rename Employee column to Employee code". Only the admin (or people the admin allows) can change them; the
web server checks that. Stored in display_rules.json next to this file:

    {"text_columns": ["Employee"], "rename": {"Employee": "Employee code"}}

text_columns = shown as plain text (no 1,04,645 commas, no rounding). rename = column header shown instead.
Keys are the chatbot's column names (Employee, WorkOrderNo, CTC ...), matched case-insensitively.
"""
import difflib
import json
import os
import re
import threading
from pathlib import Path

DEFAULT = {"text_columns": [], "rename": {}}
_LOCK = threading.RLock()

# words people use for a column -> the chatbot's column name (output columns of answers)
COLUMN_WORDS = {
    "Employee": ["employee code", "emp code", "employee id", "emp id", "employee no", "employee number", "employee",
                 "emp", "empcode", "employeecode"],
    "Name": ["employee name", "emp name", "name"],
    "WorkOrderNo": ["work order no", "work order number", "workorder no", "workorder number", "work order", "workorder",
                    "wo no", "wo"],
    "UniqueNo": ["unique no", "unique number", "uniqueno"], "UniqueNo2": ["unique no 2", "uniqueno2"],
    "InvoiceNo": ["invoice no", "invoice number", "invoiceno"], "FinaluniqCode": ["final unique code", "finaluniqcode"],
    "Department": ["department", "dept", "client"], "Category": ["category"],
    "CTC": ["ctc"], "BaseSalesValue": ["base sales value", "base sales", "base value"],
    "InvoiceAmount": ["invoice amount", "invoice value", "invoice"], "Margin": ["margin amount", "margin"],
    "MarginPct": ["margin %", "margin percent", "margin percentage"], "Employees": ["employee count", "employees",
                                                                                   "head count", "headcount"],
    "PaidDays": ["paid days"], "GrossAmount": ["gross", "gross amount", "gross salary"], "TakeHome": ["take home",
                                                                                                       "take-home"],
    "IncomeTax": ["tds", "income tax"], "Month": ["month"], "Year": ["year"],
}
_PHRASES = sorted(((w, k) for k, ws in COLUMN_WORDS.items() for w in ws), key=lambda x: -len(x[0]))


def col_of(text):
    """'emplyee code' / 'EMPLOYEE' / 'work order no' -> column name, or None. Typos allowed."""
    t = re.sub(r"[^a-z0-9% ]", " ", str(text).lower())
    t = re.sub(r"\b(the|column|columns|field|value|values|col)\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return None
    for w, k in _PHRASES:
        if t == w:
            return k
    for w, k in _PHRASES:
        if re.search(rf"(?<![a-z]){re.escape(w)}(?![a-z])", t):
            return k
    close = difflib.get_close_matches(t, [w for w, _ in _PHRASES], n=1, cutoff=0.82)
    return dict(_PHRASES).get(close[0]) if close else None


_COL = r"(?P<col>[A-Za-z][\w %.\-]{0,40}?)"
_NEW = r"[\"']?(?P<new>[A-Za-z][\w %().\-]{0,40}?)[\"']?"
_END = r"(?=\s+(?:and|also|then|from\s+now|please|plz)\b|[.;,!?]|$)"
TEXT_RULES = [
    re.compile(r"\b(?:do\s*n[o']?t|dont|never|stop)\s+(?:treat(?:ing)?|show(?:ing)?|format(?:ting)?|consider(?:ing)?|"
               r"read(?:ing)?)\s+" + _COL + r"\s+as\s+(?:a\s+)?numbers?\b", re.I),
    re.compile(r"\b(?:treat|show|keep|format|consider)\s+" + _COL + r"\s+as\s+(?:a\s+)?(?:plain\s+)?(?:text|code|string)\b",
               re.I),
    re.compile(r"\b(?:do\s*n[o']?t|dont|never|no|without|remove|stop\s+adding)\s+(?:add(?:ing)?|put(?:ting)?|use|using|show(?:ing)?)?"
               r"\s*(?:\"\s*,\s*\"|'\s*,\s*'|\(\s*,\s*\)|,|commas?|comma\s+separators?|thousand\s+separators?)\s+"
               r"(?:in|for|on|from)\s+(?:the\s+)?" + _COL + _END, re.I),
]
NUMBER_RULES = [
    re.compile(r"\b(?:treat|show|format)\s+" + _COL + r"\s+as\s+(?:a\s+)?numbers?\b", re.I),
    re.compile(r"\b(?:add|put|use|show)\s+(?:\"\s*,\s*\"|commas?)\s+(?:in|for|on)\s+(?:the\s+)?" + _COL + _END, re.I),
]
RENAME_RULES = [
    re.compile(r"\b(?:rename|re-name)\s+(?:the\s+)?" + _COL + r"\s+(?:column\s+)?(?:to|as|into)\s+" + _NEW + _END, re.I),
    re.compile(r"\b(?:call|label|name|title|show)\s+(?:the\s+)?" + _COL + r"\s+column\s+(?:to|as)\s+" + _NEW + _END, re.I),
    re.compile(r"\b" + _COL + r"\s+column(?:'?s)?\s+name\s+(?:is|should\s+be|will\s+be|=|must\s+be|to\s+be)?\s*"
               r"(?:make\s+as|made\s+as|make|changed?\s+to|change\s+as|set\s+as|set\s+to|as|to)?\s*" + _NEW + _END, re.I),
    re.compile(r"\bchange\s+(?:the\s+)?" + _COL + r"\s+column\s+name\s+(?:to|as)\s+" + _NEW + _END, re.I),
]
RESET_RX = re.compile(r"\b(reset|clear|remove|delete)\s+(all\s+)?(the\s+)?display\s+(rules|settings|names)\b", re.I)
SHOW_RX = re.compile(r"\b(show|list|what\s+are)\s+(the\s+|my\s+)?display\s+(rules|settings)\b", re.I)
UNDO_RENAME_RX = re.compile(r"\b(?:undo|remove|reset)\s+(?:the\s+)?(?:rename|new\s+name|column\s+name)\s+(?:of|for)\s+"
                            + _COL + _END, re.I)


def _clean_new(s):
    s = re.sub(r"\s+", " ", str(s or "")).strip(" \"'.")
    s = re.sub(r"\b(please|plz|from now on|always)\b.*$", "", s, flags=re.I).strip()
    return s.title() if s.isupper() or s.islower() else s


def parse(text):
    """Finds display instructions in a message. Returns (actions, leftover question text).
    actions: [("text", col) | ("number", col) | ("rename", col, new) | ("unrename", col) | ("reset",) | ("show",)]"""
    q = str(text)
    actions, spans = [], []
    if RESET_RX.search(q):
        return [("reset",)], ""
    if SHOW_RX.search(q):
        return [("show",)], ""

    def take(rx, kind):
        for m in rx.finditer(q):
            start = m.start()
            if m.start("col") == m.start():                # "<col> column name is ..." - the col is the LAST few words
                words = list(re.finditer(r"\S+", m.group("col")))
                col = None
                for n in (2, 3, 1):
                    if len(words) >= n:
                        c = col_of(" ".join(w.group() for w in words[-n:]))
                        if c:
                            col, start = c, m.start("col") + words[-n].start()
                            break
            else:
                col = col_of(m.group("col"))
            if not col or any(a < m.end() and start < b for a, b in spans):
                continue
            if kind == "rename":
                new = _clean_new(m.group("new"))
                if not new or col_of(new) == col and new.lower() == col.lower():
                    continue
                actions.append(("rename", col, new))
            elif (kind, col) not in actions:
                actions.append((kind, col))
            spans.append((start, m.end()))
    for rx in TEXT_RULES:
        take(rx, "text")
    for rx in RENAME_RULES:
        take(rx, "rename")
    for rx in NUMBER_RULES:
        take(rx, "number")
    for m in UNDO_RENAME_RX.finditer(q):
        col = col_of(m.group("col"))
        if col:
            actions.append(("unrename", col))
            spans.append((m.start(), m.end()))
    if not actions:
        return [], q
    rest = q
    for a, b in sorted(spans, reverse=True):
        rest = rest[:a] + " " + rest[b:]
    rest = re.sub(r"\b(from\s+now(\s+on(wards)?)?|always|and|also|then|please|plz)\b(\s*\b(and|also|then)\b)*\s*(?=[.,;]|$)",
                  " ", rest, flags=re.I)
    rest = re.sub(r"\s*\b(and|also|then)\s*$", "", re.sub(r"\s+", " ", rest).strip(" ,.;"), flags=re.I).strip()
    rest = re.sub(r"\b(from\s+now(\s+on(wards)?)?)\b", " ", rest, flags=re.I)
    rest = re.sub(r"\s+(and|also)(\s+(and|also))+\b", r" \1", rest, flags=re.I)
    rest = re.sub(r"\s+", " ", rest).strip(" ,.;")
    rest = re.sub(r"\s+\b(and|also|then)$", "", rest, flags=re.I).strip()
    return actions, rest


class DisplayRules:
    def __init__(self, folder):
        self.path = Path(folder) / "display_rules.json"

    def load(self):
        with _LOCK:
            try:
                d = json.loads(self.path.read_text("utf-8"))
            except Exception:
                d = {}
        return {"text_columns": list(d.get("text_columns", [])), "rename": dict(d.get("rename", {}))}

    def save(self, d):
        with _LOCK:
            tmp = self.path.with_suffix(".tmp")
            tmp.write_text(json.dumps(d, indent=2), "utf-8")
            os.replace(tmp, self.path)

    def apply(self, actions):
        """Applies actions, returns a list of plain-language lines describing what changed."""
        with _LOCK:
            d, done = self.load(), []
            for a in actions:
                if a[0] == "reset":
                    d = {"text_columns": [], "rename": {}}
                    done.append("All display rules removed - columns look as normal again.")
                elif a[0] == "text":
                    if a[1] not in d["text_columns"]:
                        d["text_columns"].append(a[1])
                    done.append(f"**{d['rename'].get(a[1], a[1])}** is shown as plain text (no commas, not as a number).")
                elif a[0] == "number":
                    d["text_columns"] = [c for c in d["text_columns"] if c != a[1]]
                    done.append(f"**{d['rename'].get(a[1], a[1])}** is shown as a number again.")
                elif a[0] == "rename":
                    d["rename"][a[1]] = a[2]
                    done.append(f"Column **{a[1]}** is now called **{a[2]}**.")
                elif a[0] == "unrename":
                    if d["rename"].pop(a[1], None):
                        done.append(f"Column **{a[1]}** has its normal name again.")
            self.save(d)
            return done

    def describe(self):
        d = self.load()
        lines = [f"- **{c}** → shown as **{n}**" for c, n in d["rename"].items()]
        lines += [f"- **{d['rename'].get(c, c)}** → plain text (no commas)" for c in d["text_columns"]]
        return "\n".join(lines) or "- (none - all columns look as normal)"


def matches(colname, key):
    """Does an answer column ('Employee', 'employee', 'Employee code') belong to rule key 'Employee'?"""
    return str(colname).strip().lower() == key.lower()
