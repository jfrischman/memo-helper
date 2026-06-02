@echo off
setlocal
set "APPDIR=%~dp0"
set "PYTHON=C:\Users\jfrischman\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "URL=http://127.0.0.1:8799/"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%URL%' -TimeoutSec 2; if ($r.StatusCode -eq 200) { Start-Process '%URL%'; exit 0 } } catch {}"

start "" /b "%PYTHON%" "%APPDIR%app.py"
timeout /t 2 /nobreak >nul
start "" "%URL%"
