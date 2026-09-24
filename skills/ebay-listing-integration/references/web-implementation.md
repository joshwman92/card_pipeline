# Web implementation contract

Use this for the intended Lucas web app. Reuse the runtime-neutral listing service and schemas, but replace every desktop transport and persistence adapter.

## Required boundaries

- Browser: forms, previews, user confirmation, progress, and sanitized errors only. It never receives eBay application credentials, authorization codes beyond the backend callback, refresh/access tokens, broker credentials, raw eBay error bodies, or direct privileged eBay endpoints.
- API/controller: authenticates Lucas session and tenant, validates ownership and CSRF, accepts idempotency key, calls the domain service, and returns a sanitized operation resource.
- Domain service: draft validation, inventory mapping, complete-replacement payload retention, state transitions, duplicate prevention, eBay sequencing, retry classification, and reconciliation rules.
- Infrastructure adapters: eBay HTTP client, KMS-backed credential repository, database repositories, durable image service, queue, scheduler, webhook verifier, clock, and observability.

Do not copy localhost callbacks, DPAPI, JSON-file repositories, Tkinter event queues, or desktop broker connection tokens into the web app.

## Minimum data model

Use database constraints as well as application checks.

`ebay_connections`:

- tenant/user ID, environment, marketplace, immutable eBay account ID when available
- encrypted refresh-token envelope and KMS key/version metadata
- granted scopes and status
- selected policy IDs and merchant location per marketplace
- connected/refreshed/revoked/last-success timestamps and safe error code
- unique active connection constraint for the supported account model

`ebay_listing_links`:

- inventory ID, connection ID, marketplace, stable SKU, offer ID, listing ID
- ownership/API type, local status, remote status, quantity, price, sold quantity
- retained draft/inventory-item payload required for complete replacement
- last operation ID, attempt count, last sync time, safe error summary, timestamps
- unique `(inventory_id, connection_id, marketplace)` and appropriate SKU/offer/listing constraints

`ebay_operations`:

- UUID/correlation ID, idempotency key, actor, tenant, action, state, attempt count
- inventory/listing references, request fingerprint, safe eBay IDs/status, retry-after, timestamps
- sanitized structured error; no credentials, buyer data, full authorization response, or unrestricted raw body
- unique `(tenant, action, idempotency_key)`

`ebay_oauth_states` and webhook-deduplication tables must enforce expiry and one-time consumption.

## Endpoint behavior

Suggested endpoints, adapted to repository conventions:

- `POST /integrations/ebay/connect-intents` returns an authorization URL after persisting state.
- `GET /integrations/ebay/callback` validates and consumes state before code exchange.
- `GET /integrations/ebay/connection` returns non-secret status/onboarding readiness.
- `DELETE /integrations/ebay/connection` revokes/deletes credentials after explicit confirmation.
- `GET /integrations/ebay/setup-options` returns current compatible policies and locations.
- `POST /inventory/{id}/ebay-drafts` creates/updates a locally validated draft.
- `POST /ebay-drafts/{id}/prepare` uploads images, replaces inventory item, creates/updates offer, stores offer ID, and returns fee estimate.
- `POST /ebay-drafts/{id}/publish` requires explicit confirmation/idempotency and publishes the retained offer.
- `PATCH /ebay-listings/{id}` revises via retained complete payload and `updateOffer`.
- `POST /ebay-listings/{id}/withdraw` and `/republish` require confirmation.
- `POST /internal/ebay/reconcile` is queue/scheduler-only, not a public user endpoint.

Every inventory/listing endpoint rechecks tenant ownership and current inventory eligibility. Do not trust IDs or tenant/account fields supplied by the browser.

## AWS mapping

Use only services present in the actual repository, but the expected shape is:

- Secrets Manager for application credentials; KMS envelope encryption for per-seller refresh tokens
- relational transaction or DynamoDB conditional write for duplicate prevention and operation state
- S3/Media API adapter for durable images, never expiring presigned URLs in listing payloads
- SQS for publish/reconcile work with message deduplication where available
- EventBridge Scheduler for periodic reconciliation
- API Gateway/Lambda or existing service for OAuth callback and authenticated endpoints
- CloudWatch structured metrics/alarms with centralized redaction

## Web-specific tests

- CSRF/state binding to the initiating user/tenant, TTL, one-time consume, and replay
- tenant isolation for connections, drafts, operations, listing links, images, and management actions
- duplicate HTTP requests and duplicate queue delivery converge on one offer/listing
- eBay success followed by database failure reconciles by SKU/offer without duplication
- optimistic/conditional lock rejects two publishers for the same inventory record
- KMS denied/decrypt failure and key rotation behavior
- queue retry exhaustion and operator-attention transition
- webhook public-key retrieval/cache, signature verification, replay/deduplication, fast acknowledgement, and authoritative refresh
- SSRF controls for any server-fetched image URL
- redaction snapshot proving tokens and buyer data do not enter logs or error responses
