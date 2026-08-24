# Windows Mobile Public Setup

This moves the existing LUCAS mobile public URLs from Mac to the always-on Windows laptop.

## Public URLs

- Team: `https://team-lucas.mikeyscards.com/mobile/team`
- Personal: `https://lucas.mikeyscards.com/mobile/personal`

## One-Time Cloudflare Files

Cloudflare tunnel credential JSON files are secrets and are not committed to Git.

Copy these files from the old Mac `~/.cloudflared` folder into the Windows folder:

```text
C:\Users\user\.cloudflared\3b34592e-77d8-4976-8eac-26e771289bee.json
C:\Users\user\.cloudflared\789db1ce-bcba-479c-9fb5-f7b374e63fe3.json
```

Tunnel mapping:

- `lucas-team`: `3b34592e-77d8-4976-8eac-26e771289bee` -> `http://127.0.0.1:8765`
- `lucas-personal`: `789db1ce-bcba-479c-9fb5-f7b374e63fe3` -> `/ebay*` to `http://127.0.0.1:8788`, everything else to `http://127.0.0.1:8766`

## Setup

Run once in PowerShell from the repo folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\windows\setup_mobile_tunnels.ps1
```

This writes:

```text
C:\Users\user\.cloudflared\lucas-team.yml
C:\Users\user\.cloudflared\lucas-personal.yml
```

## Start Everything

Double-click:

```text
Run LUCAS Mobile Stack.vbs
```

That starts:

- Team mobile server on port `8765`
- Personal mobile server on port `8766`
- eBay broker on port `8788`
- Team Cloudflare tunnel
- Personal Cloudflare tunnel

Logs are written to:

```text
work\logs\mobile-tunnel-team.log
work\logs\mobile-tunnel-personal.log
```

## Startup

Put a shortcut to `Run LUCAS Mobile Stack.vbs` in:

```text
shell:startup
```

After reboot/login, wait 30-60 seconds, then test the public URLs above from LTE.

## Mac Cutover

After Windows public URLs work, stop the old Mac tunnel LaunchAgents so only Windows is the canonical mobile host.
