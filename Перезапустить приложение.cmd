@echo off
cd /d "%~dp0"
echo Restarting Supplier Docs...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":3213 .*LISTENING"') do taskkill /PID %%P /T /F >nul 2>&1
timeout /t 1 /nobreak >nul
call "%~dp0Start_Supplier_Docs.cmd"
