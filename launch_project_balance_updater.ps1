$ErrorActionPreference = 'Stop'

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = 'py'
$url = 'http://127.0.0.1:8799/'

try {
  $response = Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 2
  if ($response.StatusCode -eq 200) {
    Start-Process $url
    exit 0
  }
} catch {
  # The app is not running yet, so start it below.
}

Start-Process -WindowStyle Hidden -FilePath $python -ArgumentList 'app.py' -WorkingDirectory $appDir
Start-Sleep -Seconds 2
Start-Process $url
