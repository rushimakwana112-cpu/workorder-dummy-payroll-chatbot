"""
Work-order groups for the Payroll chatbot (pure Python, testable on its own).

"from now on treat workorder no start with m% as north workorder and workorder start with WOX as south workorder"
-> work orders starting M are "NORTH", starting WOX are "SOUTH". Only the master / admins can change them (the web
server checks that and asks "Save this rule?" first). Every user can then ask "NORTH margin for sep-26",
"SOUTH vs NORTH CTC", "work order group wise base sales". Stored in workorder_groups.json next to this file:

    {"groups": [{"name": "NORTH", "prefixes": ["M"]}, {"name": "SOUTH", "prefixes": ["WOX"]}]}

SQL is built from the saved rules only, as escaped literals (names / prefixes are checked to safe characters first).
"""
import json
import os
import re
import threading
from pathlib import Path

_LOCK = threading.RLock()
NAME_OK = re.compile(r"[A-Za-z][A-Za-z0-9_\-]{1,19}")
PREFIX_OK = re.compile(r"[A-Za-z0-9/\-_.]{1,30}")
# words that can't be a group name (they are ordinary question words)
NOT_NAME = {"a", "an", "the", "of", "per", "on", "well", "is", "are", "be", "list", "text", "number", "numbers", "code",
            "string", "work", "order", "orders", "workorder", "workorders", "wo", "group", "groups", "all", "total",
            "salary", "margin", "ctc", "data", "details", "detail", "same", "new", "other", "others", "one", "ones",
            "this", "that", "those", "these", "it", "them", "and", "or", "to", "for", "in", "by", "wise", "rule", "rules",
            "employee", "employees", "month", "year", "summary", "invoice", "sales", "base", "follows", "below"}

_WO = r"(?:work\s*-?\s*orders?|workorders?|w\.?o\.?s?)"
_WO_NO = _WO + r"(?:\s*(?:no|nos|number|numbers|code|codes)\.?)?"
_START = r"(?:which\s+|that\s+|whose\s+(?:no\.?|number)\s+)?(?:start(?:s|ing)?|begin(?:s|ning)?|starts?\s+from)\s+(?:with\s+)?"
_PFX = r"[\"']?[A-Za-z0-9/\-_.]{1,30}%?[\"']?"
_PFXS = r"(?P<pfx>" + _PFX + r"(?:\s*(?:,|/|\bor\b|\band\b(?!\s+(?:" + _WO + r"|treat|call|consider)\b))\s*" + _PFX + r")*)"
_NAME = r"[\"']?(?P<name>[A-Za-z][A-Za-z0-9_\-]{1,19})[\"']?"
_TAIL = r"(?:\s+(?:" + _WO + r"|group|type))?"
DEFINE_RULES = [
    # "(treat) workorder no start with m% as north workorder"
    re.compile(r"\b" + _WO_NO + r"\s+" + _START + _PFXS + r"\s+(?:as|is|are|=|means?|should\s+be(?:\s+(?:treated|called)\s+as)?|"
               r"will\s+be|belongs?\s+to|in)\s+(?:an?\s+|the\s+)?" + _NAME + _TAIL, re.I),
    # "treat m% as north (work orders)" - a % makes it clearly a work-order pattern
    re.compile(r"\b(?:treat|consider|call|count|mark|tag|label|take)\s+[\"']?(?P<pfx>[A-Za-z0-9/\-_.]{1,30}%)[\"']?\s+"
               r"(?:" + _WO + r"\s+)?as\s+(?:an?\s+|the\s+)?" + _NAME + _TAIL, re.I),
    # "north work orders are (the ones) starting with M"
    re.compile(r"\b" + _NAME + r"\s+(?:" + _WO + r"|group)\s+(?:are|=|means?|is|will\s+be)\s+(?:the\s+)?(?:ones?\s+|those\s+)?"
               r"(?:" + _WO_NO + r"\s+)?" + _START + _PFXS, re.I),
]
# words that make a define sentence an instruction (and not a data question that happens to say "... as ...")
INSTRUCTION_RX = re.compile(r"\b(treat|consider|call|count|mark|tag|label|define|classify|group|from\s+now(\s+on)?|always|"
                            r"onwards?|going\s+forward|means?|remember|save|rule|are\s+(the\s+)?(ones?|those))\b|=", re.I)
DATA_PERIOD = re.compile(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sept?|oct|nov|dec)[a-z]*[\s\-/'.]*\d{2,4}\b|"
                         r"\b\d{1,2}[/\-]\d{2,4}\b|\b20\d\d[/\-]\d{1,2}\b", re.I)
_RULE_WORDS = r"(?:" + _WO + r"[\s\-]*)?(?:group|rule|groups|rules)"
LIST_RX = re.compile(r"\b(show|list|what\s+are|view|see|display|tell\s+me)\b.{0,25}?" + _WO + r"[\s\-]*(group|rule)s?\b|"
                     r"\b(show|list|what\s+are|view|see)\b.{0,10}\bgroup\s+rules\b", re.I)
RESET_RX = re.compile(r"\b(remove|delete|clear|reset|drop|forget)\s+(all\s+)?(the\s+|my\s+)?" + _WO + r"[\s\-]*(group|rule)s\b", re.I)
REMOVE_RX = re.compile(r"\b(?:remove|delete|drop|forget|cancel|clear)\s+(?:the\s+)?" + _NAME + r"\s+" + _RULE_WORDS + r"\b", re.I)
RENAME_RX = re.compile(r"\brename\s+(?:the\s+)?[\"']?(?P<old>[A-Za-z][A-Za-z0-9_\-]{1,19})[\"']?\s+(?:" + _RULE_WORDS +
                       r"\s+)?(?:to|as)\s+[\"']?(?P<new>[A-Za-z][A-Za-z0-9_\-]{1,19})[\"']?\s*[.!]*\s*$", re.I)
CHANGE_RX = re.compile(r"\b(?:change|update|edit|modify|set)\s+(?:the\s+)?" + _NAME + r"\s+(?:" + _RULE_WORDS + r"|" + _WO +
                       r")?\s*(?:to|so\s+(?:that\s+)?(?:it\s+)?|:)?\s*(?:(?:" + _WO_NO + r"\s+)?" + _START + r"|=\s*)" + _PFXS,
                       re.I)
CHANGE_PCT_RX = re.compile(r"\b(?:change|update|edit|modify|set)\s+(?:the\s+)?" + _NAME + r"\s+(?:" + _RULE_WORDS + r"\s+)?"
                           r"(?:to|as|=)\s*(?P<pfx>[\"']?[A-Za-z0-9/\-_.]{1,30}%[\"']?(?:\s*(?:,|\bor\b)\s*[\"']?"
                           r"[A-Za-z0-9/\-_.]{1,30}%[\"']?)*)", re.I)      # "change SOUTH to WOY%" (the % says: a prefix)
YES_RX = re.compile(r"^\s*(yes|y|yeah|yep|ok(ay)?|confirm(ed)?|save(\s+it)?|yes[,\s]+save(\s+it)?|go\s+ahead|do\s+it)\s*[.!]*\s*$",
                    re.I)
NO_RX = re.compile(r"^\s*(no|n|nope|cancel|don'?t|do\s+not\s+save|stop|never\s*mind)\s*[.!]*\s*$", re.I)


def _prefixes(text):
    out = []
    for p in re.split(r"\s*(?:,|/(?=\s)|\bor\b|\band\b)\s*", str(text)):
        p = p.strip(" \"'").rstrip("%").upper()
        if p and PREFIX_OK.fullmatch(p) and p not in out:
            out.append(p)
    return out


def _name(s):
    s = str(s or "").strip(" \"'")
    return s.upper() if NAME_OK.fullmatch(s) and s.lower() not in NOT_NAME else None


def parse(text, existing=()):
    """Finds work-order group instructions in a message. Returns a list of actions (empty = a normal question):
    ("set", NAME, [PREFIXES]) | ("remove", NAME) | ("rename", OLD, NEW) | ("reset",) | ("list",) | ("error", message).
    existing = names of the saved groups (needed for remove / change / rename)."""
    q = str(text).strip()
    have = {str(n).upper() for n in existing}
    if LIST_RX.search(q) and not DATA_PERIOD.search(q):
        return [("list",)]
    if RESET_RX.search(q):
        return [("reset",)]
    m = RENAME_RX.search(q)
    if m and m.group("old").upper() in have:
        new = _name(m.group("new"))
        return [("rename", m.group("old").upper(), new)] if new else [("error", f"“{m.group('new')}” can't be a group name.")]
    m = REMOVE_RX.search(q)
    if m:
        n = m.group("name").upper()
        return [("remove", n)] if n in have else [("error", f"There is no work-order group called **{n}**.")]
    if DATA_PERIOD.search(q):                    # "... for sep-26" = a data question, never a rule change
        return []
    m = CHANGE_RX.search(q) or CHANGE_PCT_RX.search(q)
    if m and m.group("name").upper() in have:
        pf = _prefixes(m.group("pfx"))
        return [("set", m.group("name").upper(), pf)] if pf else []
    if not INSTRUCTION_RX.search(q):
        return []
    acts, spans = [], []
    for rx in DEFINE_RULES:
        for m in rx.finditer(q):
            if any(a < m.end() and m.start() < b for a, b in spans):
                continue
            name, pf = _name(m.group("name")), _prefixes(m.group("pfx"))
            if not name or not pf:
                continue
            for a in acts:                       # "... M as NORTH and MS as NORTH" -> one group, both prefixes
                if a[1] == name:
                    a[2].extend(p for p in pf if p not in a[2])
                    break
            else:
                acts.append(("set", name, pf))
            spans.append((m.start(), m.end()))
    return acts


def find_names(text, groups):
    """Group names said in a question, longest first, as whole words only ('east' never matches inside 'northeast').
    Returns (names in the order said, question with the names taken out)."""
    q, found = str(text), []
    for g in sorted(groups, key=lambda g: -len(g["name"])):
        rx = re.compile(r"(?<![A-Za-z0-9_\-])" + re.escape(g["name"]) + r"(?:'s)?(?![A-Za-z0-9_\-])", re.I)
        m = rx.search(q)
        if m:
            found.append((m.start(), g["name"]))
            q = rx.sub(" ", q)
    q = re.sub(r"\s+", " ", q).strip()
    return [n for _, n in sorted(found)], q


def _lit(s):
    return "N'" + str(s).replace("'", "''") + "'"


def _like(prefix):
    """Prefix -> LIKE pattern with the wildcards [ % _ taken literally."""
    p = re.sub(r"([\[%_])", r"[\1]", str(prefix))
    return _lit(p + "%")


def case_sql(groups, col="[WorkOrderNo]", other="Other"):
    """CASE WHEN [WorkOrderNo] LIKE N'WOX%' THEN N'SOUTH' WHEN ... ELSE N'Other' END - longest prefix first, so a more
    specific prefix wins whichever group it is in."""
    pairs = [(p, g["name"]) for g in groups for p in g["prefixes"]
             if NAME_OK.fullmatch(g["name"]) and PREFIX_OK.fullmatch(p)]
    if not pairs:
        return _lit(other)
    whens = " ".join(f"WHEN {col} LIKE {_like(p)} THEN {_lit(n)}" for p, n in sorted(pairs, key=lambda x: -len(x[0])))
    return f"CASE {whens} ELSE {_lit(other)} END"


def filter_sql(groups, names, col="[WorkOrderNo]"):
    """'(CASE ... END) IN (N'SOUTH')' - only work orders of these groups (same CASE, so a short prefix never picks up a longer group's rows)."""
    ok = [n for n in names if any(g["name"] == n for g in groups)]
    if not ok:
        return None
    return f"({case_sql(groups, col)}) IN ({', '.join(_lit(n) for n in ok)})"


class WorkOrderGroups:
    def __init__(self, folder):
        self.path = Path(folder) / "workorder_groups.json"

    def load(self):
        with _LOCK:
            try:
                d = json.loads(self.path.read_text("utf-8"))
            except Exception:
                d = {}
        out = []
        for g in d.get("groups", []):
            name, pf = str(g.get("name", "")), [str(p).upper() for p in g.get("prefixes", [])]
            if NAME_OK.fullmatch(name):
                out.append({"name": name.upper(), "prefixes": [p for p in pf if PREFIX_OK.fullmatch(p)]})
        return out

    def names(self):
        return [g["name"] for g in self.load()]

    def save(self, groups):
        with _LOCK:
            tmp = self.path.with_suffix(".tmp")
            tmp.write_text(json.dumps({"groups": groups}, indent=2), "utf-8")
            os.replace(tmp, self.path)

    def describe_actions(self, actions):
        """Plain-language lines for the 'Save this rule?' question (nothing is saved here)."""
        cur, lines = {g["name"]: g["prefixes"] for g in self.load()}, []
        for a in actions:
            if a[0] == "set":
                pf = " or ".join(f"**{p}**" for p in a[2])
                old = cur.get(a[1])
                lines.append(f"Work orders starting with {pf} → **{a[1]}**" +
                             (f" (was: {', '.join(old) or 'none'})" if old is not None else ""))
                for n, ps in cur.items():
                    taken = [p for p in ps if p in a[2]]
                    if n != a[1] and taken:
                        lines.append(f"{', '.join(taken)} moves out of **{n}**")
            elif a[0] == "remove":
                lines.append(f"Remove the **{a[1]}** group (its work orders become *Other*)")
            elif a[0] == "rename":
                lines.append(f"Rename the **{a[1]}** group to **{a[2]}**")
            elif a[0] == "reset":
                lines.append("Remove **all** work-order groups")
        return lines

    def apply(self, actions):
        """Applies actions, returns a list of plain-language lines describing what changed."""
        with _LOCK:
            groups, done = self.load(), []
            for a in actions:
                if a[0] == "reset":
                    groups = []
                    done.append("All work-order groups removed.")
                elif a[0] == "set":
                    for g in groups:                   # a prefix belongs to one group only
                        if g["name"] != a[1]:
                            g["prefixes"] = [p for p in g["prefixes"] if p not in a[2]]
                    g = next((g for g in groups if g["name"] == a[1]), None)
                    if g:
                        g["prefixes"] = list(a[2])
                    else:
                        groups.append({"name": a[1], "prefixes": list(a[2])})
                    done.append(f"Work orders starting with {' or '.join(a[2])} are **{a[1]}**.")
                elif a[0] == "remove":
                    if any(g["name"] == a[1] for g in groups):
                        groups = [g for g in groups if g["name"] != a[1]]
                        done.append(f"The **{a[1]}** group is removed.")
                elif a[0] == "rename":
                    for g in groups:
                        if g["name"] == a[1]:
                            g["name"] = a[2]
                            done.append(f"The **{a[1]}** group is now called **{a[2]}**.")
            groups = [g for g in groups if g["prefixes"]]
            self.save(groups)
            return done

    def describe(self):
        lines = [f"- **{g['name']}** → work orders starting with {' or '.join(g['prefixes'])}" for g in self.load()]
        return "\n".join(lines) or "- (none yet)"
