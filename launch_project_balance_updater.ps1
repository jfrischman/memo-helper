$ErrorActionPreference = 'Stop'

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$url = 'http://127.0.0.1:8799/'

# Always (re)start so the latest code is loaded. Python imports modules once at
# process start and does NOT hot-reload edited files, so any running instance must
# be stopped first or code changes won't take effect.
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -match 'app\.py' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Milliseconds 600

Start-Process -WindowStyle Hidden -FilePath 'py' -ArgumentList 'app.py' -WorkingDirectory $appDir
Start-Sleep -Seconds 2
Start-Process $url
