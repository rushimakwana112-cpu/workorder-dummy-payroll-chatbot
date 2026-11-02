@echo off
rem Opens ports 8080 (new site) and 8501 (old site) in Windows Firewall so other office PCs / your laptop
rem on the office network can reach the Payroll Chatbot. Needs administrator rights - it will ask for permission.
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Asking for administrator permission...
  powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
set FAILED=0
for %%p in (8080 8501) do (
  netsh advfirewall firewall delete rule name="Payroll Chatbot %%p" >nul 2>&1
  netsh advfirewall firewall add rule name="Payroll Chatbot %%p" dir=in action=allow protocol=TCP localport=%%p profile=domain,private
  if errorlevel 1 set FAILED=1
)
if %FAILED%==1 (
  echo.
  echo Could not add the firewall rules. Your company may control the firewall - please ask IT
  echo to allow incoming TCP ports 8080 and 8501 on this PC.
) else (
  echo.
  echo Done. Ports 8080 and 8501 are now open for office networks ^(Domain / Private^).
)
echo.
pause
