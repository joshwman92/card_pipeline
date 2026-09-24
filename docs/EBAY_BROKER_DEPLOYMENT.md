# LUCAS eBay Broker

The same broker contract supports the desktop trial and the future hosted web app. Start in eBay Sandbox. End users should only click **Connect eBay**, sign in at eBay, and return to LUCAS ready to prepare a listing.

## Correct Hosted Flow

1. Hosted LUCAS opens `https://lucas.mikeyscards.com/ebay/connect`.
2. The eBay broker redirects the user to official eBay OAuth.
3. eBay redirects back to the broker callback registered in the eBay developer app.
4. The broker validates and consumes the persisted OAuth nonce, then stores the eBay refresh token encrypted for the Windows service account.
5. The broker redirects to LUCAS at `https://lucas.mikeyscards.com/mobile/ebay/broker/callback` with a short-lived, one-time exchange code. No reusable credential is placed in the browser URL.
6. The LUCAS backend exchanges that code directly with `/ebay/connection/exchange` and saves only an encrypted LUCAS connection token in its account store.
7. When listing cards, LUCAS asks the broker for short-lived eBay access tokens. Disconnect deletes both the broker connection and the local token.

The broker owns `/ebay*`. LUCAS owns `/mobile/ebay/broker/callback`. Do not route the LUCAS callback through `/ebay*`.

## Required Windows Environment

Set these on the Windows host account that runs the broker. Use Sandbox for development and the desktop trial:

```powershell
setx EBAY_ENV sandbox
setx EBAY_CLIENT_ID "..."
setx EBAY_CLIENT_SECRET "..."
setx EBAY_RUNAME "..."
setx EBAY_OAUTH_SCOPES "https://api.ebay.com/oauth/api_scope https://api.ebay.com/oauth/api_scope/sell.account https://api.ebay.com/oauth/api_scope/sell.inventory"
setx LUCAS_EBAY_BROKER_PUBLIC_URL "https://lucas.mikeyscards.com/ebay"
setx LUCAS_EBAY_ALLOWED_CALLBACK_HOSTS "lucas.mikeyscards.com,team-lucas.mikeyscards.com"
setx LUCAS_EBAY_BROKER_STATE_SECRET "replace-with-long-random-secret"
setx LUCAS_EBAY_BROKER_STORE_PATH "C:\LUCAS\ebay_broker_connections.json"
```

The broker and desktop app must run under the same Windows account that created their encrypted records; Windows DPAPI intentionally prevents another account or machine from decrypting them. Back up configuration, not token-store files.

For a deliberate production rollout, change `EBAY_ENV` to `production` and set `EBAY_ALLOW_PRODUCTION_LISTINGS=1` only in the trusted listing service after the listing has been reviewed. Leave that flag unset for Sandbox and for any read-only production checks.

Restart PowerShell after `setx`, or set the same values in the current shell with `$env:NAME = "value"` before testing.

`EBAY_RUNAME` must be the redirect name/URL registered in the shared LUCAS eBay developer app for the broker callback. It is not a per-user value.

## Update The Windows Checkout

From the Windows LUCAS repo:

```powershell
cd C:\Users\user\Documents\card_pipeline
git fetch origin
git checkout -B feature/ebay-broker-windows-host origin/feature/ebay-broker-windows-host
git rev-parse HEAD
dir .\ebay_broker_server.py
```

The file `ebay_broker_server.py` must exist in that folder. If `git rev-parse --show-toplevel` points somewhere else, run these commands from that top-level repo folder instead.

## Run The Broker On Windows

From the Windows repo:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\ebay-broker\windows\run-ebay-broker.ps1
```

Verify locally:

```powershell
curl http://127.0.0.1:8788/health
```

Expected JSON includes:

```json
{"ok": true, "service": "lucas-ebay-broker"}
```

To install it as a startup task:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\ebay-broker\windows\install-ebay-broker-task.ps1
```

## Cloudflare Routing On Windows

The tunnel must route only `/ebay*` to the broker:

```yaml
tunnel: YOUR_TUNNEL_ID
credentials-file: C:\Users\user\.cloudflared\YOUR_TUNNEL_ID.json

ingress:
  - hostname: lucas.mikeyscards.com
    path: /ebay*
    service: http://127.0.0.1:8788
  - hostname: lucas.mikeyscards.com
    service: http://127.0.0.1:8766
  - hostname: team-lucas.mikeyscards.com
    service: http://127.0.0.1:8765
  - service: http_status:404
```

The same template is saved at `deploy/ebay-broker/windows/cloudflared-windows.example.yml`.

Verify public routing:

```powershell
curl https://lucas.mikeyscards.com/ebay/health
curl https://lucas.mikeyscards.com/mobile/api/config
```

The first response should say `service: lucas-ebay-broker`. The second should still be the normal LUCAS mobile API. If `/ebay/health` says `lucas-mobile`, Cloudflare is still routing eBay to the app instead of the broker.

## Desktop/Hosted LUCAS Configuration

Hosted Windows LUCAS should have a public mobile URL configured, for example:

```powershell
setx LUCAS_PERSONAL_MOBILE_PUBLIC_URL "https://lucas.mikeyscards.com/mobile/personal"
setx LUCAS_TEAM_MOBILE_PUBLIC_URL "https://team-lucas.mikeyscards.com/mobile/team"
setx LUCAS_EBAY_BROKER_URL "https://lucas.mikeyscards.com/ebay"
```

The desktop callback may instead be `http://127.0.0.1:<port>/ebay/broker/callback`; the broker allowlist accepts only that loopback callback or the configured HTTPS hosted callback. The callback handler must always exchange the one-time code against its configured broker URL rather than trusting a broker origin supplied by the browser.

Production users should not set `LUCAS_EBAY_USE_LOCAL_OAUTH`. That flag is only for local developer testing.

## Verification Checklist

1. Confirm `/ebay/health` reports the intended `mode` (`sandbox` for the trial).
2. Connect once and verify that a repeated callback or repeated `/connection/exchange` returns an expired/replayed error.
3. Confirm the callback URL contains `exchange_code` but never `connection_token`, `refresh_token`, or `access_token`.
4. Prepare a Sandbox offer and verify it reaches `READY` before approving publish.
5. Retry the prepare and publish operations; the stable SKU and saved offer ID must prevent duplicate offers/listings.
6. Disconnect and verify subsequent `/token` requests for the former connection return `unknown connection`.
