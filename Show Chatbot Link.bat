@echo off
rem Shows the link other people should open, and checks the chatbot is running.
echo.
echo  Share one of these links with your colleagues (new Blue Panther site):
echo  ---------------------------------------------------------------------
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
  for /f "tokens=* delims= " %%b in ("%%a") do echo    http://%%b:8080
)
echo    http://%COMPUTERNAME%:8080
echo.
echo  (Old Streamlit site, backup: http://localhost:8501 - opens on THIS PC only)
echo.
netstat -an | findstr ":8080 " | findstr "LISTENING" >nul
if %errorlevel%==0 (echo  New site status: RUNNING) else (echo  New site status: NOT RUNNING - double-click "Start Chatbot (no window).vbs")
netstat -an | findstr ":8501 " | findstr "LISTENING" >nul
if %errorlevel%==0 (echo  Old site status: RUNNING) else (echo  Old site status: NOT RUNNING)
echo.
echo  Network type (must be DomainAuthenticated or Private, not Public):
powershell -NoProfile -Command "Get-NetConnectionProfile | ForEach-Object { '   ' + $_.Name + ' : ' + $_.NetworkCategory }"
echo.
pause
