@echo off
cd /d "%~dp0"
set "NODE_EXE=C:\Users\e.savko\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
  echo Windows PowerShell was not found.
  echo.
  echo Please send a screenshot of this window to Codex.
  pause
  exit /b 1
)
echo Starting Supplier Docs...
start "Supplier Docs Server" cmd.exe /k ""%NODE_EXE%" server.mjs"
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:3213"
echo The application was opened in your browser.
echo Keep the Supplier Docs Server window open while working. You can minimize it.
timeout /t 3 /nobreak >nul
