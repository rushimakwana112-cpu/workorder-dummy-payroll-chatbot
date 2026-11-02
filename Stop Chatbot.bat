@echo off
rem Stops the Payroll Chatbot running in the background: new site (port 8080) and old Streamlit site (port 8501).
cd /d "%~dp0"
echo stop> stop.flag
for %%p in (8080 8501) do (
  for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":%%p " ^| findstr "LISTENING"') do taskkill /F /PID %%a >nul 2>&1
)
echo Payroll Chatbot stopped.
timeout /t 2 >nul
