param(
    [switch]$InstallService
)

$ErrorActionPreference = "Stop"

$cloudflaredDir = Join-Path $env:USERPROFILE ".cloudflared"
$teamTunnelId = "3b34592e-77d8-4976-8eac-26e771289bee"
$personalTunnelId = "789db1ce-bcba-479c-9fb5-f7b374e63fe3"

New-Item -ItemType Directory -Force -Path $cloudflaredDir | Out-Null

function Write-TunnelConfig {
    param(
        [string]$Name,
        [string]$TunnelId,
        [string]$Hostname,
        [int]$Port,
        [string[]]$ExtraIngress = @()
    )

    $credentialsFile = Join-Path $cloudflaredDir "$TunnelId.json"
    if (!(Test-Path $credentialsFile)) {
        Write-Warning "Missing credential file: $credentialsFile"
        Write-Warning "Copy this JSON from the Mac ~/.cloudflared folder or recreate/login the tunnel on Windows."
    }

    $configPath = Join-Path $cloudflaredDir "$Name.yml"
    $ingressLines = @()
    $ingressLines += "ingress:"
    $ingressLines += $ExtraIngress
    $ingressLines += "  - hostname: $Hostname"
    $ingressLines += "    service: http://127.0.0.1:$Port"
    $ingressLines += "  - service: http_status:404"

    $config = @(
@"
tunnel: $TunnelId
credentials-file: $credentialsFile
"@
    )
    $config += $ingressLines
    Set-Content -Path $configPath -Value $config -Encoding ascii
    Write-Host "Wrote $configPath -> http://127.0.0.1:$Port"
}

Write-TunnelConfig -Name "lucas-team" -TunnelId $teamTunnelId -Hostname "team-lucas.mikeyscards.com" -Port 8765
Write-TunnelConfig -Name "lucas-personal" -TunnelId $personalTunnelId -Hostname "lucas.mikeyscards.com" -Port 8766 -ExtraIngress @(
    "  - hostname: lucas.mikeyscards.com",
    "    path: /ebay*",
    "    service: http://127.0.0.1:8788"
)

if ($InstallService) {
    Write-Host ""
    Write-Host "Service install is intentionally manual because this repo runs two tunnels."
    Write-Host "Use Task Scheduler or the VBS startup launcher instead:"
    Write-Host "  Run LUCAS Mobile Stack.vbs"
}

Write-Host ""
Write-Host "Next:"
Write-Host "  1. Make sure cloudflared.exe is installed and available in PATH."
Write-Host "  2. Start servers and tunnels with Run LUCAS Mobile Stack.vbs."
Write-Host "  3. Test:"
Write-Host "     https://team-lucas.mikeyscards.com/mobile/team"
Write-Host "     https://lucas.mikeyscards.com/mobile/personal"
