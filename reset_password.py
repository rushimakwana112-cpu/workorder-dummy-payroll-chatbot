"""
Set a new password for a chatbot ID (e.g. when the admin forgot it). Run ON THIS PC only:
    python reset_password.py
It asks for the ID and the new password (typing is hidden), saves it scrambled in users.json,
unlocks the ID, and writes an entry in audit_log.csv. The password is never shown or stored as text.
If there are no IDs yet (no users.json), it creates the master administrator instead.
"""
import csv
import datetime
import getpass
import hashlib
import json
import os
import re
import secrets
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
USERS_FILE = APP_DIR / "users.json"
AUDIT_FILE = APP_DIR / "audit_log.csv"


def _hash(pw, salt):                                    # same as payroll_chatbot.py
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), 200_000).hex()


def pw_problem(pw):
    if len(pw) < 8:
        return "Password must be at least 8 characters."
    if not re.search(r"[A-Za-z]", pw) or not re.search(r"\d", pw):
        return "Password must contain letters and numbers."
    return None


def main():
    users = json.loads(USERS_FILE.read_text("utf-8")) if USERS_FILE.exists() else {}
    if users:
        print("Chatbot IDs:", ", ".join(f"{k} ({v.get('role', 'user')})" for k, v in users.items()))
        uid = input("ID to reset [admin]: ").strip().lower() or "admin"
        if uid not in users:
            print(f"No ID '{uid}'. Nothing changed.")
            return
        action = "password_reset_local"
    else:
        # first ID is the master (full powers); admins can't create a master later
        print("No chatbot IDs yet - creating the master administrator.")
        uid = input("New master ID [admin]: ").strip().lower() or "admin"
        if not re.match(r"^[a-z0-9._-]{3,30}$", uid):
            print("ID: 3-30 characters, letters/numbers/._- only. Nothing changed.")
            return
        name = input("Your name: ").strip()
        users[uid] = {"name": name or uid, "role": "master", "active": True,
                      "created": datetime.date.today().isoformat()}
        action = "setup_local"
    while True:
        pw = getpass.getpass("New password (hidden while typing): ")
        err = pw_problem(pw)
        if err:
            print(err)
            continue
        if getpass.getpass("Type it again: ") != pw:
            print("They don't match. Try again.")
            continue
        break
    rec = users[uid]
    rec["salt"] = secrets.token_hex(16)
    rec["hash"] = _hash(pw, rec["salt"])
    rec["must_change"] = False
    rec["active"] = True
    tmp = USERS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(users, indent=2), "utf-8")
    os.replace(tmp, USERS_FILE)
    new = not AUDIT_FILE.exists()
    with AUDIT_FILE.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "user", "action", "detail"])
        w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), uid, action,
                    f"reset_password.py on {os.environ.get('COMPUTERNAME', '?')} by {os.environ.get('USERNAME', '?')}"])
    print(f"\nDone. Sign in with ID '{uid}' and your new password.")
    print("If it still says 'locked', wait 15 minutes (or run Stop + Start Chatbot).")


if __name__ == "__main__":
    main()
    input("\nPress Enter to close...")
