$ErrorActionPreference = "Continue"
$project = "C:\Users\DELL\Desktop\pdf-editor-tool"
$python = "C:\Python314\python.exe"
if (!(Test-Path $python)) { $python = "python" }

Set-Location $project

$logDir = Join-Path $project "logs"
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

while ($true) {
  $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -Path (Join-Path $logDir "server.log") -Value "[$stamp] Starting uvicorn on 127.0.0.1:8000"

  & $python -m uvicorn app:app --host 127.0.0.1 --port 8000 2>&1 | Tee-Object -FilePath (Join-Path $logDir "server.log") -Append
  $exitCode = $LASTEXITCODE

  $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -Path (Join-Path $logDir "server.log") -Value "[$stamp] Uvicorn exited with code $exitCode"
  Start-Sleep -Seconds 2
}

