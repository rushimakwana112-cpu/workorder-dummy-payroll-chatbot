"""Validity date: after it the chatbot answers nothing until the master renews it.
The date sits in a hidden file next to this one, signed with this install's own key (hidden file .pb_ui_key, created
on first run and never shared) - typing a new date into the file by hand (or deleting it) makes the chatbot expired.
Only the master renews it: Settings → Chatbot validity (asks the master password again).
A brand-new install (no key yet) starts valid for FIRST_RUN_DAYS days."""
import datetime
import hashlib
import hmac
import json
import os
import secrets

_UI_STATE = ".pb_ui_state"
_UI_KEY = ".pb_ui_key"
FIRST_RUN_DAYS = 365


def _here(name):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def _hide(p):
    if os.name == "nt":
        os.system(f'attrib +h +s "{p}" >nul 2>&1')


def _ui_seed():
    p = _here(_UI_KEY)
    try:
        with open(p, encoding="utf-8") as fh:
            return bytes.fromhex(fh.read().strip())
    except (OSError, ValueError):
        pass
    key = secrets.token_bytes(32)                       # first run: a new key for this install only
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(key.hex())
    _hide(p)
    write_valid_until(datetime.date.today() + datetime.timedelta(days=FIRST_RUN_DAYS), key)
    return key


def _ui_state_path():
    return _here(_UI_STATE)


def _ui_sign(day, key=None):
    return hmac.new(key or _ui_seed(), ("valid-until:" + day).encode("utf-8"), hashlib.sha256).hexdigest()


def valid_until():
    """The date the chatbot works until, or None if the file is missing or was changed by hand."""
    if not os.path.exists(_here(_UI_KEY)):
        _ui_seed()                                      # brand-new install: make its key and first validity date
    try:
        with open(_ui_state_path(), encoding="utf-8") as fh:
            d = json.load(fh)
        day = str(d.get("u", ""))
        if not hmac.compare_digest(str(d.get("s", "")), _ui_sign(day)):
            return None
        return datetime.date.fromisoformat(day)
    except Exception:
        return None


def write_valid_until(day, key=None):
    """Saves a new validity date (datetime.date), signed; the file is hidden."""
    p = _ui_state_path()
    if os.path.exists(p) and os.name == "nt":
        os.system(f'attrib -h -s "{p}" >nul 2>&1')         # a hidden file cannot be overwritten on Windows
    with open(p, "w", encoding="utf-8") as fh:
        json.dump({"u": day.isoformat(), "s": _ui_sign(day.isoformat(), key)}, fh)
    _hide(p)


def expired_answer():
    """None while the chatbot is valid; else the 'expired' reply."""
    until = valid_until()
    if until and datetime.date.today() <= until:
        return None
    when = f" on {until.strftime('%d-%m-%Y')}" if until else ""
    return {"text": f"⛔ **This chatbot has expired{when}.** It needs to be renewed before it can answer again. "
                    "Please contact the administrator.",
            "note": "", "df": None, "report": "expired", "sql": "", "params": [], "intent": {}}
