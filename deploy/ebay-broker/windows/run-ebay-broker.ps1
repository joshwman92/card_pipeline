$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
Set-Location $RepoRoot
$LogDir = Join-Path $RepoRoot "work\logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogPath = Join-Path $LogDir "ebay-broker.log"

if (-not $env:LUCAS_EBAY_BROKER_PUBLIC_URL) {
    $env:LUCAS_EBAY_BROKER_PUBLIC_URL = "https://lucas.mikeyscards.com/ebay"
}
if (-not $env:LUCAS_EBAY_ALLOWED_CALLBACK_HOSTS) {
    $env:LUCAS_EBAY_ALLOWED_CALLBACK_HOSTS = "lucas.mikeyscards.com,team-lucas.mikeyscards.com"
}
if (-not $env:LUCAS_EBAY_BROKER_STORE_PATH) {
    $env:LUCAS_EBAY_BROKER_STORE_PATH = "C:\LUCAS\ebay_broker_connections.json"
}
if (-not $env:HOST) {
    $env:HOST = "127.0.0.1"
}
if (-not $env:PORT) {
    $env:PORT = "8788"
}

New-Item -ItemType Directory -Force -Path (Split-Path $env:LUCAS_EBAY_BROKER_STORE_PATH) | Out-Null

$Python = "python"
if (Test-Path ".\.venv\Scripts\python.exe") {
    $Python = ".\.venv\Scripts\python.exe"
}

$message = "Starting LUCAS eBay broker from $RepoRoot on http://$($env:HOST):$($env:PORT)"
Write-Host $message
"[$(Get-Date -Format s)] $message" | Out-File -FilePath $LogPath -Append -Encoding utf8
& $Python ".\ebay_broker_server.py" *>&1 | Tee-Object -FilePath $LogPath -Append
