# Desktop implementation pattern

Use this only when Lucas is a desktop application. Keep the listing domain service portable so the web implementation can reuse its draft schema, validation, idempotency, mappings, and state machine.

## Trust boundary

The installed UI is not an appropriate home for a shared eBay client secret or seller refresh token. Use this flow:

1. Desktop opens the configured broker `/connect` URL with account/profile and an allow-listed callback.
2. Broker creates a high-entropy state, persists a hashed nonce with account, callback, creation time, and short TTL, then redirects to eBay.
3. Broker callback validates signature, callback, TTL, and persisted nonce and consumes the nonce before exchanging the code.
4. Broker encrypts the refresh token and indexes the connection by a hash of a new opaque connection credential.
5. Broker redirects to desktop with a separate short-lived one-time exchange code and non-secret account/marketplace hints, never the reusable credential. The exchange response is authoritative for account, marketplace, and Sandbox/Production environment.
6. Desktop loopback callback ignores broker origins and environment values from the browser request, posts the code only to its configured broker origin, receives the opaque credential and authoritative metadata, protects the credential with the OS credential facility, and shows a no-store completion page.
7. Desktop requests short-lived access tokens from the broker when a background operation starts.

Reject signed-but-unpersisted state, callback reuse, missing/expired exchange codes, non-loopback HTTP callbacks, hosted callbacks outside the allow-list, and broker-origin overrides.

## Windows credential storage

Current-user DPAPI is suitable for a Windows-only broker trial when:

- the scheduled task/service runs as the same Windows account that created the store;
- the data is never copied to another machine or account as if it were portable;
- tests cover protect/unprotect through the actual launcher identity;
- broker connection lookup keys are hashes, not reusable plaintext tokens;
- token values never enter shared `CARD_PIPELINE` data.

If any condition does not hold, use Windows Credential Manager or a managed server-side secret store. Do not invent portable encryption with a key stored beside the ciphertext.

## Desktop persistence

Keep two stores:

- a machine-local protected connection store containing only broker details and protected credentials;
- a shared or application-data listing store containing non-secret inventory ID, account ID, marketplace, SKU, offer ID, listing ID, ownership type, status, retained complete-replacement draft, fee result, operation state, attempt count, timestamps, and sanitized errors.

Use the application's existing atomic-write and shared-lock primitives. Do not store access tokens, refresh tokens, reusable broker credentials, authorization codes, or raw callback URLs in the listing/audit store.

## UI and threading

- Add **List on eBay** only for one eligible active inventory record.
- Require a stable inventory ID, positive price, at least one image, leaf category, aspects, condition data, policies, and location before network work.
- Run OAuth exchange, seller setup discovery, image upload, Inventory calls, fee lookup, publish, withdraw, republish, and reconciliation off the UI thread.
- Treat Prepare and Publish as separate UI states. The second step shows environment, price, offer ID, and fee estimate or fee-unavailable warning.
- Keep Sandbox visually explicit. A production environment plus a separate operator gate is required before enabling a live Publish confirmation.
- Persist the offer ID before the confirmation dialog so app restarts and ambiguous failures resume rather than duplicate.
- On startup or manual reconciliation, query by saved offer ID; during prepare retry, also search by stable SKU.

## Desktop exit criteria

Do not call the desktop slice production-ready until a real Sandbox seller has completed connection, policy/location discovery, graded and ungraded publication, duplicate-click retry, restart after Prepare, withdrawal, republish, token expiry/reauthorization, disconnect, and reconciliation tests.
