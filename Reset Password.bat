@echo off
rem Sets a new password for a chatbot ID (e.g. admin). Works only on this PC. Typing is hidden.
cd /d "%~dp0"
python reset_password.py
