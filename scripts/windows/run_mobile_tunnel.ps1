param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("team", "personal")]
    [string]$Profile
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$logDir = Join-Path $repoRoot "work\logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir "mobile-tunnel-$Profile.log"

function Find-Cloudflared {
    $command = Get-Command "cloudflared.exe" -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }
    $candidates = @(
        (Join-Path $env:ProgramFiles "cloudflared\cloudflared.exe"),
        (Join-Path $env:ProgramFiles "Cloudflare\cloudflared.exe"),
        (Join-Path $env:ProgramFiles "Cloudflare\cloudflared\cloudflared.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "cloudflared\cloudflared.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Cloudflare\cloudflared.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Cloudflare\cloudflared\cloudflared.exe"),
        "C:\Cloudflared\bin\cloudflared.exe",
        "C:\Cloudflared\cloudflared.exe",
        (Join-Path $env:USERPROFILE "cloudflared.exe"),
        (Join-Path $repoRoot "cloudflared.exe"),
        (Join-Path $repoRoot "bin\cloudflared.exe")
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    return $null
}

$cloudflared = Find-Cloudflared
if (!$cloudflared) {
    "[$(Get-Date -Format s)] cloudflared.exe not found. Install it, then run scripts\windows\setup_mobile_tunnels.ps1." | Out-File -FilePath $logPath -Append -Encoding utf8
    throw "cloudflared.exe not found."
}

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "setup_mobile_tunnels.ps1")

$cloudflaredDir = Join-Path $env:USERPROFILE ".cloudflared"
if ($Profile -eq "team") {
    $configPath = Join-Path $cloudflaredDir "lucas-team.yml"
    $name = "lucas-team"
} else {
    $configPath = Join-Path $cloudflaredDir "lucas-personal.yml"
    $name = "lucas-personal"
}

"[$(Get-Date -Format s)] starting $name with $cloudflared --config $configPath tunnel run $name" | Out-File -FilePath $logPath -Append -Encoding utf8
$ErrorActionPreference = "Continue"
& $cloudflared --config $configPath tunnel run $name *>&1 | Tee-Object -FilePath $logPath -Append
exit $LASTEXITCODE
