"""
Personal reminders for the Payroll AI chatbot (v14) - pure Python, no Streamlit, so it can be tested on its own.

    "i have to pay 114 workorder tomorrow, remind me"      -> reminder tomorrow 10:00
    "remind me in 2 hours to check the invoice"
    "remind me on 5-oct-26 at 3 pm to send the TDS file"
    "remind me every month on 5th to run payroll"
    "show my reminders" / "done 3" / "delete reminder 3" / "snooze 3 1 day"

Stored in reminders.json next to the app (one list, each reminder belongs to ONE login).
Reminders are shown inside the chatbot when they are due - there is no e-mail / phone message.
"""
import calendar
import datetime as dt
import json
import os
import re
import threading
from pathlib import Path

TRIGGER = re.compile(r"\b(remind(?:er|ers)?(?:\s+me)?|set\s+(?:a\s+)?reminder|don'?t\s+let\s+me\s+forget|"
                     r"yaad\s+dila\w*|alert\s+me)\b", re.I)
LIST_RX = re.compile(r"^\s*(?:show|list|see|view|what\s+are)?\s*(?:all\s+)?(?:my\s+)?(?:pending\s+)?reminders?\s*\??\s*$|"
                     r"^\s*(?:show|list|see|view)\s+(?:all\s+)?(?:my\s+)?(?:pending\s+)?reminders?\b", re.I)
DONE_RX = re.compile(r"^\s*(?:mark\s+)?(?:reminder\s+)?(?:#?(\d+)\s+(?:is\s+)?(?:done|completed?|finished)|"
                     r"(?:done|complete|completed|finished)\s+(?:reminder\s+)?#?(\d+))\s*[.!]*\s*$", re.I)
DEL_RX = re.compile(r"^\s*(?:delete|remove|cancel|forget)\s+(?:my\s+)?reminder\s*#?(\d+)\s*[.!]*\s*$", re.I)
SNOOZE_RX = re.compile(r"^\s*snooze\s+(?:reminder\s+)?#?(\d+)(?:\s+(?:by|for))?\s*(.*)$", re.I)

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
DAYS = {d.lower(): i for i, d in enumerate(calendar.day_name)}
DAYS.update({d.lower(): i for i, d in enumerate(calendar.day_abbr)})
TOMORROW = r"(?:tomm?or+ow|tomm?or+ow|tmrw|tmr|tomorow|tommorrow|tomorrow|kal)"
DEFAULT_HOUR = 10                           # a date without a time -> 10:00 AM


def _time_in(text):
    """(hour, minute, matched text) or None.  '3 pm', 'at 10', '10:30am', '15:45', 'morning', 'evening'."""
    m = re.search(r"\b(?:at|by|@)?\s*(\d{1,2})(?::|\.)(\d{2})\s*(am|pm|a\.m\.|p\.m\.)?\b", text, re.I)
    if not m:
        m = re.search(r"\b(?:at|by|@)?\s*(\d{1,2})\s*(am|pm|a\.m\.|p\.m\.)\b", text, re.I)
        if m:
            h, mi, ap = int(m.group(1)), 0, m.group(2)
        else:
            m = re.search(r"\b(?:at|by)\s+(\d{1,2})\b(?!\s*(?:-|/|st|nd|rd|th|days?|hours?|hrs?|min))", text, re.I)
            if m:
                h, mi, ap = int(m.group(1)), 0, None
                if 1 <= h <= 7:
                    h += 12                 # "at 3" during office hours = 3 pm
            else:
                for word, hh in (("morning", 10), ("noon", 12), ("afternoon", 15), ("evening", 18), ("tonight", 20),
                                 ("night", 20), ("subah", 10), ("shaam", 18)):
                    mm = re.search(rf"\b{word}\b", text, re.I)
                    if mm:
                        return hh, 0, mm.group(0)
                return None
    else:
        h, mi, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap:
        ap = ap.lower().replace(".", "")
        if ap == "pm" and h < 12:
            h += 12
        if ap == "am" and h == 12:
            h = 0
    if not (0 <= h <= 23 and 0 <= mi <= 59):
        return None
    return h, mi, m.group(0)


def _date_in(text, now):
    """(date, matched text, repeat) or None."""
    t = text.lower()
    rep = None
    m = re.search(r"\b(every\s*day|daily|roz)\b", t)
    if m:
        return now.date(), m.group(0), "daily"
    m = re.search(r"\bevery\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b|\bweekly\b", t)
    if m:
        if m.group(1):
            wd = DAYS[m.group(1)]
            d = now.date() + dt.timedelta(days=(wd - now.weekday()) % 7)
        else:
            d = now.date()
        return d, m.group(0), "weekly"
    m = re.search(r"\b(?:every\s+month|monthly)(?:\s+on)?(?:\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?)?\b", t)
    if m:
        day = int(m.group(1)) if m.group(1) else now.day
        d = _month_day(now.year, now.month, day)
        if d < now.date():
            y, mo = (now.year + (now.month // 12), now.month % 12 + 1)
            d = _month_day(y, mo, day)
        return d, m.group(0), "monthly"
    m = re.search(r"\bday\s+after\s+" + TOMORROW + r"\b|\bparso(n)?\b", t)
    if m:
        return now.date() + dt.timedelta(days=2), m.group(0), rep
    m = re.search(r"\b" + TOMORROW + r"\b", t)
    if m:
        return now.date() + dt.timedelta(days=1), m.group(0), rep
    m = re.search(r"\b(today|aaj|tonight|this\s+evening)\b", t)
    if m:
        return now.date(), m.group(0), rep
    m = re.search(r"\b(next|this|on|coming)?\s*(monday|tuesday|wednesday|thursday|friday|saturday|sunday|mon|tue|wed|"
                  r"thu|fri|sat|sun)\b", t)
    if m:
        wd = DAYS[m.group(2)]
        ahead = (wd - now.weekday()) % 7
        if ahead == 0 or (m.group(1) or "").strip() == "next" and ahead < 7 and ahead == 0:
            ahead = 7
        return now.date() + dt.timedelta(days=ahead), m.group(0), rep
    # 5-oct-26, 5 oct 2026, 5th october, oct 5
    m = re.search(r"\b(\d{1,2})(?:st|nd|rd|th)?[\s\-/.]*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*"
                  r"(?:[\s\-/.,']*(\d{2,4}))?\b", t)
    if m:
        d = _mk(int(m.group(3)) if m.group(3) else None, MONTHS[m.group(2)], int(m.group(1)), now)
        if d:
            return d, m.group(0), rep
    m = re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-]*(\d{1,2})(?:st|nd|rd|th)?"
                  r"(?:[\s,]+(\d{4}))?\b", t)
    if m:
        d = _mk(int(m.group(3)) if m.group(3) else None, MONTHS[m.group(1)], int(m.group(2)), now)
        if d:
            return d, m.group(0), rep
    m = re.search(r"\b(\d{1,2})[/\-.](\d{1,2})(?:[/\-.](\d{2,4}))?\b", t)         # 05/10/2026 = 5 Oct (dd/mm)
    if m and int(m.group(2)) <= 12:
        d = _mk(int(m.group(3)) if m.group(3) else None, int(m.group(2)), int(m.group(1)), now)
        if d:
            return d, m.group(0), rep
    m = re.search(r"\bon\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)\b", t)
    if m:
        day = int(m.group(1))
        d = _month_day(now.year, now.month, day)
        if d < now.date():
            y, mo = (now.year + (now.month // 12), now.month % 12 + 1)
            d = _month_day(y, mo, day)
        return d, m.group(0), rep
    return None


def _month_day(y, m, day):
    return dt.date(y, m, min(day, calendar.monthrange(y, m)[1]))


def _mk(y, m, d, now):
    if y is not None and y < 100:
        y += 2000
    try:
        res = dt.date(y or now.year, m, d)
    except ValueError:
        return None
    if y is None and res < now.date():
        res = dt.date(now.year + 1, m, d)
    return res


def parse_reminder(text, now=None):
    """'remind me tomorrow to pay 114 workorder' -> {"due": datetime, "text": ..., "repeat": None|"daily"|...}.
    {"due": None, "text": ...} when no time was given (ask the user).  None = not a reminder request."""
    now = now or dt.datetime.now()
    q = str(text or "").strip()
    if not TRIGGER.search(q) or LIST_RX.match(q):
        return None
    rel = re.search(r"\b(?:in|after)\s+(\d+|an?|one|two|three|half\s+an?)\s*(min(?:ute)?s?|hours?|hrs?|h|days?|weeks?)\b",
                    q, re.I)
    used = []
    due, repeat = None, None
    if rel:
        n_txt = rel.group(1).lower()
        n = {"a": 1, "an": 1, "one": 1, "two": 2, "three": 3}.get(n_txt, None)
        if n is None:
            n = 0.5 if n_txt.startswith("half") else int(n_txt)
        unit = rel.group(2).lower()
        delta = (dt.timedelta(minutes=n) if unit.startswith("min") else dt.timedelta(hours=n) if unit[0] == "h"
                 else dt.timedelta(days=n) if unit.startswith("day") else dt.timedelta(weeks=n))
        due = (now + delta).replace(second=0, microsecond=0)
        used.append(rel.group(0))
    else:
        d = _date_in(q, now)
        tm = _time_in(q)
        if d:
            used.append(d[1])
            repeat = d[2]
        if tm:
            used.append(tm[2])
        if d or tm:
            day = d[0] if d else now.date()
            h, mi = (tm[0], tm[1]) if tm else (DEFAULT_HOUR, 0)
            due = dt.datetime.combine(day, dt.time(h, mi))
            if due <= now and not d:                        # "at 9 am" already passed -> tomorrow
                due += dt.timedelta(days=1)
            if due <= now and repeat:
                due = next_due(due, repeat, now)
    body = q
    for u in used:
        body = body.replace(u, " ")
    body = TRIGGER.sub(" ", body)
    body = re.sub(r"(?i)\b(can|could|would|will)\s+(you|u)\b|\bplease\b|\bpls\b|\bplz\b|\bkindly\b|\babout\b(?=\s*$)",
                  " ", body)
    for _ in range(3):
        body = re.sub(r"(?i)^\s*(me|to|that|for|about|on|at|by)\b", " ", body.strip())
    body = re.sub(r"(?i)\b(me|to|that|on|at|by)\s*[?.!,]*\s*$", " ", body.strip())
    body = re.sub(r"\s+", " ", body).strip(" ,.?!-:;")
    return {"due": due, "text": body or "(reminder)", "repeat": repeat}


def next_due(due, repeat, now=None):
    now = now or dt.datetime.now()
    while due <= now:
        if repeat == "daily":
            due += dt.timedelta(days=1)
        elif repeat == "weekly":
            due += dt.timedelta(weeks=1)
        elif repeat == "monthly":
            y, m = (due.year + (due.month // 12), due.month % 12 + 1)
            due = dt.datetime.combine(_month_day(y, m, due.day), due.time())
        else:
            break
    return due


def when_text(due, now=None):
    now = now or dt.datetime.now()
    d = due.date()
    day = ("today" if d == now.date() else "tomorrow" if d == now.date() + dt.timedelta(days=1)
           else due.strftime("%a %d-%b-%Y"))
    return f"{day} at {due.strftime('%I:%M %p').lstrip('0')}"


class ReminderStore:
    _lock = threading.RLock()

    def __init__(self, folder):
        self.path = Path(folder) / "reminders.json"

    def _load(self):
        try:
            data = json.loads(self.path.read_text("utf-8"))
        except Exception:
            data = {}
        data.setdefault("next_id", 1)
        data.setdefault("items", [])
        return data

    def _save(self, data):
        tmp = self.path.with_name(self.path.name + ".tmp")
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), "utf-8")
        os.replace(tmp, self.path)

    def add(self, user, text, due, repeat=None):
        with self._lock:
            data = self._load()
            r = {"id": data["next_id"], "user": user, "text": str(text)[:300], "due": due.strftime("%Y-%m-%d %H:%M"),
                 "repeat": repeat, "done": False, "created": dt.datetime.now().strftime("%Y-%m-%d %H:%M")}
            data["next_id"] += 1
            data["items"].append(r)
            self._save(data)
            return r

    def mine(self, user, include_done=False):
        return sorted([r for r in self._load()["items"] if r["user"] == user and (include_done or not r["done"])],
                      key=lambda r: r["due"])

    def due(self, user, now=None):
        now = now or dt.datetime.now()
        return [r for r in self.mine(user) if dt.datetime.strptime(r["due"], "%Y-%m-%d %H:%M") <= now]

    def _change(self, user, rid, fn):
        with self._lock:
            data = self._load()
            for r in data["items"]:
                if r["id"] == int(rid) and r["user"] == user:       # only your own reminders
                    fn(r)
                    self._save(data)
                    return r
        return None

    def done(self, user, rid, now=None):
        def fn(r):
            if r.get("repeat"):                                    # repeating -> next time
                due = dt.datetime.strptime(r["due"], "%Y-%m-%d %H:%M")
                r["due"] = next_due(due, r["repeat"], now or dt.datetime.now()).strftime("%Y-%m-%d %H:%M")
            else:
                r["done"] = True
        return self._change(user, rid, fn)

    def delete(self, user, rid):
        with self._lock:
            data = self._load()
            keep = [r for r in data["items"] if not (r["id"] == int(rid) and r["user"] == user)]
            if len(keep) == len(data["items"]):
                return False
            data["items"] = keep
            self._save(data)
            return True

    def snooze(self, user, rid, delta, now=None):
        now = now or dt.datetime.now()
        return self._change(user, rid, lambda r: r.update(due=(now + delta).strftime("%Y-%m-%d %H:%M"), done=False))


def snooze_delta(text):
    m = re.search(r"(\d+)\s*(min(?:ute)?s?|hours?|hrs?|h|days?|weeks?)", str(text or ""), re.I)
    if not m:
        return dt.timedelta(hours=1)
    n, u = int(m.group(1)), m.group(2).lower()
    return (dt.timedelta(minutes=n) if u.startswith("min") else dt.timedelta(hours=n) if u[0] == "h"
            else dt.timedelta(days=n) if u.startswith("day") else dt.timedelta(weeks=n))
