---
name: ebay-listing-integration
description: Plan, implement, test, or review an eBay seller integration for a Lucas web or desktop application, including OAuth, inventory-card listing, listing management, and optional order sync. Use for Lucas-to-eBay engineering; do not use for unrelated marketplaces or general eBay shopping research.
---

# Lucas eBay Listing Integration

Build a secure, observable eBay integration that lets an authorized Lucas user connect an eBay seller account, publish inventory cards, and manage those listings from Lucas.

Treat this document as an execution playbook. Inspect the repository and preserve its established frontend, backend, deployment, database, authentication, logging, and test conventions rather than introducing a parallel architecture.

## Runtime routing

Choose the runtime before designing OAuth or persistence:

- **Web app:** the browser is untrusted. All OAuth code exchange, refresh, eBay API calls, credentials, idempotency, and operation state belong in the authenticated backend. Read [references/web-implementation.md](references/web-implementation.md).
- **Desktop app:** do not ship the eBay client secret or seller refresh token in the desktop process. Use a trusted broker for OAuth and access-token minting. The browser callback may carry only a short-lived, single-use exchange code; the desktop exchanges it server-to-server for an opaque broker connection credential stored with OS credential protection. Read [references/desktop-implementation.md](references/desktop-implementation.md).
- **Shared API details:** before mapping payloads, read [references/verified-api-contracts.md](references/verified-api-contracts.md). Treat live Taxonomy and Metadata API responses as authoritative when they differ.

Keep domain mapping, draft validation, Inventory API sequencing, state transitions, idempotency rules, and sanitized error handling runtime-neutral. Put transport adapters, secret storage, authentication, persistence, and UI integration behind runtime-specific boundaries.

## Required outcome

The minimum production outcome is:

1. A Lucas user can connect and disconnect an eBay seller account through eBay OAuth. Lucas never receives the user's eBay password.
2. The user can select an eligible card in Lucas inventory, complete eBay-specific fields, validate it, preview expected fees when available, and publish it to eBay.
3. Lucas stores the eBay SKU, offer ID, listing ID, status, and synchronization state against the inventory item.
4. An eBay Listings tab shows Lucas-created listings and supports permitted edits, price/quantity changes, withdrawal, and republishing.
5. Sold or unavailable cards cannot remain accidentally available through Lucas.
6. Tokens and application credentials remain server-side and encrypted.
7. Sandbox tests, production safeguards, logging, retry behavior, and operator documentation are complete.

The default MVP is one eBay marketplace (`EBAY_US`), fixed-price, single-card listings, quantity one, graded and ungraded cards, Lucas-created listings only, and Lucas as the source of truth. Do not silently expand beyond this scope.

## Non-negotiable constraints

- Use eBay's OAuth 2.0 Authorization Code Grant for seller-owned resources. Use a cryptographically random `state` value, bind it to the Lucas session, validate it once on callback, and reject missing, mismatched, expired, or replayed values.
- Never collect, proxy, log, or store eBay usernames or passwords.
- Keep the eBay client secret, access tokens, and refresh tokens out of browser code, URLs, analytics, logs, and error responses.
- Perform all eBay mutations from a trusted backend or broker. Browser code and desktop UI widgets call Lucas-owned services, never privileged eBay endpoints directly.
- Use the eBay Sandbox until the end-to-end test suite passes. A production listing must require explicit user authorization.
- Make listing publication, webhook handling, retries, and background reconciliation idempotent.
- Treat eBay category, aspect, condition, policy, and marketplace metadata as changeable. Retrieve and cache it with an expiry instead of permanently hard-coding it.
- Do not infer that an eBay success response means the Lucas database update also succeeded. Persist a recoverable operation state and reconcile partial failures.
- Listings created with the Inventory API must be managed through the Inventory API. Do not attempt to revise them through Trading API calls.
- Use complete-replacement Inventory API operations carefully: retrieve or retain every field that must survive before replacing an inventory item.
- Preserve unrelated user changes in the repository. Use migrations that are backward compatible with the current application whenever practical.
- Do not activate Production credentials, create live listings, withdraw live listings, subscribe a production webhook, or change a seller's business policies without the user's explicit authorization for that external action.
- Default configuration to Sandbox. Merely selecting `production` must not unlock publication: require a separate operator-controlled production-publish gate plus an in-product confirmation for each initial live trial.
- Never place an access token, refresh token, reusable broker connection token, or client secret in a browser callback URL. Use a short-lived single-use exchange code and bind it to the original OAuth state and callback.
- Encrypt credentials at rest with the platform facility. For the Windows-hosted desktop broker, current-user DPAPI is acceptable only when the broker always runs as that same Windows identity. For a web deployment, use the existing KMS-backed secret store and IAM boundary; do not copy the DPAPI design.

## Step 1: Inspect the application and record decisions

Before editing code:

1. Read repository instructions, the main README, deployment configuration, environment examples, schema/migrations, authentication code, inventory model, image-storage implementation, background-job system, and existing tests.
2. Identify the actual runtime. For web, record API Gateway/Lambda, ECS, Amplify, Cognito, RDS, DynamoDB, S3, SQS, EventBridge, Secrets Manager, and KMS components that exist. For desktop, record the local callback server, broker host, shared-data directory, OS credential facility, threading/event model, and packaging/runtime identity. Reuse only components already present or clearly justified by the feature.
3. Trace an inventory card from database to frontend. Record its stable ID, SKU behavior, title, description, sport/game, year, set, player/character, card number, variation, grading company, grade, certification number, condition, quantity, price, and image fields.
4. Determine whether images already have durable HTTPS URLs accessible to eBay. Expiring presigned URLs are not durable listing image identifiers.
5. Run the existing test suite and record the baseline before changing code.
6. Write a short implementation note that records the selected marketplace, seller model, listing formats, listing ownership strategy, image strategy, storage strategy, and whether orders are in scope.
7. Run a baseline secret scan of existing eBay-related files. A prior spike may already persist plaintext tokens or place reusable credentials in URLs; migrate or invalidate those records rather than building on them.

Resolve these decision gates before implementation when the repository does not answer them:

- Is Lucas connecting one company-owned eBay account or separate accounts for multiple Lucas users?
- Must Lucas manage listings created in Seller Hub or another platform, or only listings Lucas creates?
- Are auctions required in the first release, or only fixed-price listings?
- Are order retrieval, shipment tracking, refunds, offers, messages, or promoted listings part of this feature?
- Which eBay marketplaces and currencies are required?

If the default MVP satisfies the request, proceed with it and document the assumption. If an unanswered decision would materially change security, data ownership, or API selection, ask the user one concise question before implementing that branch.

## Step 2: Confirm the eBay application prerequisites

1. Confirm that the organization has an eBay Developers Program account.
2. Create or identify separate Sandbox and Production keysets.
3. Configure a Sandbox RuName/redirect URL with Lucas OAuth accept and decline URLs.
4. Configure the required privacy-policy URL and display name.
5. Create an eBay Sandbox seller test user.
6. Confirm that the seller account is eligible to sell and is opted into business policies.
7. For Production, confirm marketplace-account-deletion compliance is configured or formally opted out where allowed.
8. Determine whether restricted scopes or higher limits require an Application Growth Check. Do not promise an approval date; track it as an external dependency.

Use only the scopes required by implemented features. The expected starting scopes are:

- `https://api.ebay.com/oauth/api_scope/sell.inventory`
- `https://api.ebay.com/oauth/api_scope/sell.account`
- `https://api.ebay.com/oauth/api_scope/sell.fulfillment` only when orders, shipping, or refunds are implemented
- Any listing-notification read scope only when the relevant Notification topic is available to the application

Authoritative references:

- [Authorization](https://developer.ebay.com/develop/guides/sell/authorization)
- [Get started with eBay APIs](https://developer.ebay.com/develop/guides/sell/get-started-with-ebay-apis)
- [API call limits](https://developer.ebay.com/develop/get-started/api-call-limits)

## Step 3: Add configuration and secret handling

1. Add environment-aware configuration for eBay environment, client ID, client secret, RuName, OAuth callback, marketplace ID, locale, and API/auth base URLs.
2. Use Sandbox hosts outside Production:
   - Authorization: `https://auth.sandbox.ebay.com/oauth2/authorize`
   - API/token: `https://api.sandbox.ebay.com`
3. Use Production hosts only in Production:
   - Authorization: `https://auth.ebay.com/oauth2/authorize`
   - API/token: `https://api.ebay.com`
4. Store application credentials in the runtime's established server-side secret facility. Web deployments use KMS-backed or equivalent managed encryption. A Windows broker may use current-user DPAPI, but must document and test the service/scheduled-task identity. Store only an opaque protected broker credential in the desktop client.
5. Add safe configuration placeholders to the environment example. Never commit real credentials or tokens.
6. Add startup validation that fails with a clear operator error when required server configuration is missing.
7. Persist the credential's Sandbox/Production environment and reject token use when it differs from the current runtime. Never let a Sandbox broker credential be paired with Production API hosts, or vice versa.

## Step 4: Add the persistence model

Use names consistent with the repository, but preserve these concepts.

### eBay connection

Store:

- Lucas user or tenant ID
- eBay immutable user/account identifier when available
- marketplace and environment
- encrypted refresh token
- granted scopes
- token expiry metadata if useful
- connection status: connected, reauthorization required, disconnected, or error
- connected, refreshed, revoked, and last-success timestamps
- last safe error code; never persist raw tokens

Enforce one active connection per intended Lucas user/tenant and environment unless the product explicitly supports multiple eBay accounts.

### eBay listing link

Store:

- Lucas inventory ID
- eBay connection ID
- seller-defined SKU
- Inventory API inventory item key
- offer ID
- listing ID
- marketplace, currency, and listing format
- ownership type: Lucas Inventory API, migrated Inventory API, or Trading API
- local status and last known eBay status
- listed quantity, price, and sold quantity when applicable
- last synchronized time
- last operation ID, operation state, attempt count, and safe error summary
- created and updated timestamps

Add unique constraints that prevent duplicate publication of the same card for the same eBay account and marketplace unless explicit multi-listing behavior is designed.

### Operation/audit record

Record publish, revise, quantity update, withdraw, republish, reconnect, and reconciliation attempts with correlation IDs and sanitized results. This is required for diagnosing partial failures without exposing secrets or buyer data.

## Step 5: Implement the OAuth connection

Add backend endpoints equivalent to:

- `GET /integrations/ebay/connect`
- `GET /integrations/ebay/callback`
- `GET /integrations/ebay/status`
- `POST /integrations/ebay/disconnect`

Execution sequence:

1. Require an authenticated Lucas session on web. For desktop, bind the connection to the local Lucas profile/account and accept the callback only on loopback or an allow-listed hosted Lucas callback.
2. Generate a high-entropy, short-lived, single-use `state` value tied to the Lucas user, intended return location, and environment.
3. Redirect to eBay with `client_id`, RuName as `redirect_uri`, `response_type=code`, URL-encoded scopes, locale, and `state`.
4. On callback, validate and consume `state` before exchanging the code.
5. Exchange the authorization code server-side at `/identity/v1/oauth2/token` using HTTP Basic authentication with the client ID and secret.
6. Encrypt and store the refresh token. Keep the access token only in an encrypted short-lived cache or memory where practical.
7. Refresh access tokens server-side using `grant_type=refresh_token`. Coalesce concurrent refresh attempts so one expired token does not cause a refresh storm.
8. If eBay returns an invalid or revoked grant, mark the connection as requiring reauthorization and present a Connect Again action.
9. Disconnect by revoking tokens where supported, deleting or rendering stored credentials unusable, and retaining only the audit data allowed by policy.
10. For brokered desktop OAuth, store pending state server-side (or in an equivalent one-time store), enforce a short TTL, and consume it before code exchange. After exchange, redirect with a separate one-time connection-exchange code. Hash reusable connection credentials for broker lookup so a broker-store disclosure does not expose them.

Test consent accepted, consent declined, invalid state, replayed callback, expired code, expired access token, revoked refresh token, missing scopes, concurrent refresh, and Sandbox/Production isolation.

Do not treat a signed but stateless `state` token as single-use. Signature validation proves integrity, not freshness or replay prevention.

## Step 6: Implement seller onboarding

Before enabling Publish:

1. Use Account API to check seller privileges, relevant program enrollment, and existing payment, fulfillment, and return policies for the marketplace.
2. Use Inventory API to retrieve inventory locations.
3. Let the user select existing policies and a location. Creating or modifying business policies is a separate, explicit seller action.
4. If no usable location exists, collect the minimum shipping-origin data and create one after confirmation.
5. If policies are absent, guide the user to create them in eBay or provide an explicitly confirmed Lucas setup flow.
6. Save selected policy IDs and `merchantLocationKey` per eBay connection and marketplace.
7. Revalidate saved IDs when eBay reports they are missing, incompatible, or inactive.

Publishing requires a payment policy, fulfillment policy, return policy, inventory location, and category-compatible values.

References:

- [Account API](https://developer.ebay.com/develop/api/sell/account_api_v1)
- [Business policies](https://developer.ebay.com/api-docs/sell/static/seller-accounts/business-policies.html)
- [Inventory locations](https://developer.ebay.com/api-docs/sell/static/inventory/managing-inventory-locations.html)

## Step 7: Build category, aspect, and card-condition metadata services

1. Use Taxonomy API to obtain the default category tree for the marketplace.
2. Suggest categories using a normalized card title, but require or preserve a confirmed leaf category.
3. Retrieve item aspects for the selected leaf category.
4. Render required, recommended, and optional aspects from metadata. Do not assume every trading-card category has identical fields.
5. Use Metadata API `getItemConditionPolicies` for valid condition IDs and condition descriptors.
6. Support the three principal single-card categories when appropriate: sports trading cards, CCG individual cards, and non-sport trading card singles. Treat live metadata as authoritative over remembered IDs.
7. Map graded cards to the current grader, grade, and optional certification-number descriptors.
8. Map ungraded cards to the current ungraded-condition descriptor values.
9. Cache taxonomy and metadata with timestamps and invalidate it on relevant eBay validation errors.
10. Store the chosen eBay category and normalized aspect mapping with the Lucas listing draft so users do not repeatedly re-enter unchanged data.
11. For the three principal single-card categories, use current condition metadata. As a verified fallback for draft prefill only: graded cards use Inventory `condition=LIKE_NEW` with grader (`27501`) and grade (`27502`) descriptors plus optional certification (`27503`); ungraded cards use `condition=USED_VERY_GOOD` with card-condition descriptor `40001`. Revalidate the chosen descriptor values against live metadata before production publication.

References:

- [Taxonomy API](https://developer.ebay.com/develop/api/buy/taxonomy_api)
- [Listing Metadata guide](https://developer.ebay.com/develop/guides/sell/listing-metadata-guide)
- [Trading-card condition descriptors](https://developer.ebay.com/api-docs/user-guides/static/mip-user-guide/mip-enum-condition-descriptor-ids-for-trading-cards.html)

## Step 8: Implement durable image handling

Choose one strategy and record it:

### Preferred when supported: eBay Media API

1. Upload the Lucas image file or submit a durable source URL.
2. Capture the image resource URI or image ID from the response headers.
3. Retrieve and store the eBay Picture Services URL and expiration metadata.
4. Use the returned image URL in the inventory item.

For local files, use `POST /commerce/media/v1_beta/image/create_image_from_file` with multipart field name `image`; capture both the response `imageUrl` and `Location` image-resource URI. Do not hand-build multipart without tests for boundary formatting, filenames, MIME type, empty response bodies, and HTTP errors.

### Acceptable alternative: Lucas-hosted images

1. Provide stable, publicly reachable HTTPS URLs.
2. Do not use short-lived S3 presigned URLs.
3. Ensure eBay can fetch the object without cookies, Lucas authentication, custom headers, or IP allow-listing.
4. Preserve the image for the listing's lifetime and applicable retention period.

Validate file type, size, dimensions, ordering, missing objects, and inaccessible URLs. At least one image is required. Front and back images should remain in deterministic order.

Reference: [Managing images](https://developer.ebay.com/api-docs/sell/static/inventory/managing-image-media.html)

## Step 9: Implement listing drafts and validation

Create a Lucas listing draft rather than publishing immediately from an inventory row.

The draft must contain or derive:

- inventory ID and stable unique SKU
- marketplace and currency
- fixed-price format for the MVP
- listing duration
- quantity, defaulting to one only when correct
- title and description
- leaf category ID
- required aspects
- graded or ungraded condition plus descriptors
- ordered image URLs
- price
- policy IDs
- merchant location key
- optional best-offer terms only if explicitly supported

Add two validation layers:

1. Local validation for required fields, lengths, types, state transitions, ownership, and obvious category/condition incompatibilities.
2. eBay validation/preflight using current metadata and `getListingFees` when appropriate. Display fees as estimates, not guarantees.

Normalize eBay errors into user-actionable field messages while retaining the original error ID and sanitized response for operators. Do not expose raw backend payloads, credentials, or stack traces.

## Step 10: Publish a single-card listing

Use the Inventory API sequence:

1. Lock or claim the Lucas inventory item for publication using the application's established concurrency mechanism.
2. Create a durable operation record with an idempotency/correlation key.
3. Verify there is no existing active or in-progress listing link for the same inventory item, connection, and marketplace.
4. `PUT /sell/inventory/v1/inventory_item/{sku}` with quantity, condition, condition descriptors, product title, description, aspects, and image URLs.
5. `POST /sell/inventory/v1/offer` with SKU, marketplace, format, category, quantity, price, duration, listing policies, and merchant location.
6. Persist the returned offer ID before publishing.
7. Optionally call the listing-fee endpoint for the unpublished offer and show the estimate before the user confirms.
8. `POST /sell/inventory/v1/offer/{offerId}/publish` only after explicit confirmation.
9. Persist the returned listing ID and mark the Lucas link active.
10. If eBay succeeded but local persistence failed, leave a recoverable operation record and reconcile by SKU/offer instead of creating another listing.
11. If publishing failed, preserve the unpublished offer ID and present retryable errors without duplicating the inventory item or offer.
12. Before creating an offer on retry, query offers by the stable SKU and marketplace and reuse a matching unpublished offer. A local database lookup alone is insufficient after an eBay-success/local-write-failure split.
13. Separate **prepare** from **publish**. Prepare uploads images, replaces the inventory item, creates or updates the offer, persists the offer ID, and attempts `getListingFees`. Publish is a second action after the user reviews the draft and fee estimate or an explicit fee-unavailable warning.
14. Before retrying `publish`, retrieve the saved offer. If its remote status is already `PUBLISHED`, recover the listing ID and repair local state instead of calling `publish` again. This closes the eBay-success/local-write-failure window.

Do not decrement the Lucas item merely because a listing was published. Reserve it from other sales channels according to the application's inventory policy, then mark it sold or unavailable only after a confirmed sale or explicit operator action.

Reference: [Inventory item to offer](https://developer.ebay.com/api-docs/sell/static/inventory/inventory-item-to-offer.html)

## Step 11: Add the Lucas user interface

### Inventory menu action

Add a **List on eBay** action only for eligible, unsold inventory. The flow should:

1. Check connection and seller-onboarding status.
2. Open a draft form prefilled from the card record.
3. Show card images, category, item specifics, grading/condition, title, description, quantity, price, duration, policies, and location.
4. Show validation failures beside the relevant fields.
5. Provide a review step with the estimated fees and a clear statement that publishing may incur eBay charges.
6. Require an explicit Publish confirmation.
7. Show progress, final listing ID/link, warnings, and recoverable failures.
8. Prevent double submission while still allowing safe retry after an ambiguous network result.
9. Make the environment unmistakable in the UI. Sandbox and Production must not look identical at the confirmation boundary.

### eBay Listings tab

At minimum show:

- image and card identity
- SKU and listing ID
- marketplace and format
- price, available quantity, and sold quantity when known
- local/eBay status
- last synchronization time
- warning or failed-operation state
- link to the live eBay page

Support filters for status and search by title, SKU, listing ID, player/character, set, or certification number when those fields exist.

For Lucas Inventory API listings, support:

- edit product details through complete inventory-item replacement
- edit offer details through `updateOffer`
- optimized price/quantity updates through `bulkUpdatePriceQuantity` where useful
- withdraw while retaining the offer for possible republishing
- republish an unpublished offer after revalidation

Require confirmation for withdrawal and other externally destructive operations.

## Step 12: Synchronize state and process notifications

1. Add a background reconciliation job that retrieves relevant Inventory API items/offers and repairs stale local state.
2. Reconcile by stable identifiers: connection, marketplace, SKU, offer ID, and listing ID.
3. Use bounded retries with exponential backoff and jitter for transient errors and rate limits. Do not retry validation, authorization, or policy errors indefinitely.
4. Add a dead-letter or operator-attention state after the bounded retry policy is exhausted.
5. If available to the application, configure Notification API subscriptions for listing events and order confirmation events.
6. Implement the public HTTPS destination challenge/verification required by eBay.
7. Validate `X-EBAY-SIGNATURE` using eBay's public-key mechanism before accepting a notification.
8. Deduplicate by notification ID and make processing idempotent. Acknowledge promptly and move slow work to the existing queue system.
9. Treat webhook events as prompts to refresh authoritative state, not as the only permanent record.
10. Monitor subscription health, signature failures, delivery lag, reconciliation drift, token failures, and repeated API errors.

Reference: [Notification API and topics](https://developer.ebay.com/develop/api/buy/notification_events)

## Step 13: Branch only if existing non-Lucas listings are required

Do not add this branch for the default MVP.

1. Inventory API can retrieve and manage Inventory API records, not automatically every listing created in Seller Hub or another integration.
2. Evaluate `bulkMigrateListing` for eligible active listings that should become Lucas/Inventory API managed.
3. For listings that cannot or should not be migrated, use Trading API retrieval and mutation calls such as `GetSellerList`, `GetItem`, `ReviseItem`, `ReviseFixedPriceItem`, `EndItem`, and `EndFixedPriceItem`.
4. Persist the ownership/API type and dispatch every mutation to the correct API family.
5. Never revise an Inventory API-created listing through Trading API Revise/Relist calls.
6. Test Seller Hub-created, Trading API-created, migrated, Inventory API-created, auction, fixed-price, ended, and sold cases separately.

Explain to the user that Inventory API-created listings are managed through the Inventory API and therefore Lucas becomes their operational source of truth.

References:

- [Listing Management guide](https://developer.ebay.com/develop/guides/sell/listing-management)
- [Traditional Listing APIs](https://developer.ebay.com/develop/api/sell/traditional_listing_apis)

## Step 14: Branch only if orders and fulfillment are required

1. Use Fulfillment API `getOrders` and `getOrder` for completed-checkout orders.
2. Map eBay order line items back to Lucas inventory through listing ID and SKU.
3. Deduplicate order and line-item processing.
4. Mark inventory sold atomically and prevent overselling across other Lucas channels.
5. Support shipment tracking with `createShippingFulfillment` only after explicit shipment action or an established automated fulfillment rule.
6. Add refunds or disputes only when explicitly requested; treat them as separate high-impact workflows with confirmations and audit trails.
7. Observe the documented order-history/filter window and maintain Lucas's own permitted operational history.

Reference: [Fulfillment API](https://developer.ebay.com/develop/api/sell/fulfillment_api)

## Step 15: Test the integration

Implement automated tests at the repository's established levels.

### Unit tests

- OAuth state generation, expiry, one-time use, and validation
- scope and authorization URL construction
- token encryption/decryption boundaries and refresh coalescing
- Lucas-card to eBay inventory/offer mapping
- title, aspect, grading, condition, image, quantity, and price validation
- eBay error normalization
- listing state machine and ownership-based API dispatch
- idempotency and deduplication
- webhook signature adapter and event deduplication

### Contract/integration tests

- token exchange and refresh with mocked eBay responses
- category/aspect/condition metadata caching and invalidation
- inventory item, offer, publish, update, withdraw, and republish sequences
- partial success where eBay succeeds and Lucas persistence initially fails
- rate limiting, timeouts, 401 refresh-and-retry, 403 reauthorization, validation failures, and 5xx retry exhaustion

### Sandbox end-to-end tests

Test at least:

1. Connect and reconnect a Sandbox seller.
2. Publish one graded card with grader, grade, and certification number.
3. Publish one ungraded card with the correct card-condition descriptor.
4. Reject a draft missing a required aspect or accessible image.
5. Change price and quantity.
6. Withdraw and republish.
7. Confirm a duplicate click or retried job does not create a duplicate listing.
8. Exercise an expired access token and a revoked refresh token.
9. Receive or simulate a valid and invalid notification if notifications are in scope.
10. Reconcile a deliberately stale local status.

Run the full existing test suite after the feature tests. Report pre-existing failures separately from regressions.

For desktop implementations, also test OS credential round-trip under the actual launcher identity, loopback callback routing, broker origin pinning, single-use exchange codes, app restart recovery from a prepared offer, and UI-thread isolation for every network call. For web implementations, add authorization-boundary tests proving one tenant cannot read or mutate another tenant's connection, draft, operation, or listing.

## Step 16: Production readiness and rollout

Before Production:

1. Complete a security review covering secrets, token lifecycle, CSRF, authorization boundaries, logging redaction, webhook verification, SSRF risk from image URLs, and least-privilege IAM.
2. Confirm marketplace-account-deletion compliance and any required Growth Check.
3. Confirm Production RuName, privacy policy, callback URLs, scopes, keyset, seller policies, inventory location, and alarms.
4. Add dashboards or alerts for publish failures, token refresh failures, reconciliation drift, queue backlog, webhook signature failures, and elevated eBay error rates.
5. Add a feature flag and an emergency switch that stops new publishes while preserving read/reconciliation access.
6. Deploy schema changes and disabled code paths first when the current architecture supports staged rollout.
7. Connect a designated Production test seller.
8. Publish one explicitly approved low-risk listing, verify it on eBay, revise it, and withdraw it if that is part of the approved test.
9. Expand to a small pilot group and monitor before general availability.
10. Document operator recovery for revoked tokens, invalid policies, missing images, partial publication, duplicate-risk investigation, webhook failure, and eBay outage.

## Definition of done

Do not call the integration complete until:

- The agreed MVP scope is implemented end to end.
- All credentials and tokens are server-side, encrypted, redacted, and environment-isolated.
- OAuth acceptance, decline, expiry, revocation, and reconnection work.
- A graded and an ungraded card publish successfully in Sandbox.
- Required taxonomy, aspects, conditions, policies, location, and images are validated.
- Duplicate publication is prevented across browser retries, worker retries, and partial failures.
- The Listings tab accurately retrieves and manages Lucas-created listings.
- Withdrawal and republishing use explicit confirmation and correct state transitions.
- Reconciliation repairs stale or ambiguous local records.
- Notifications are authenticated and idempotent if enabled.
- Automated tests and the repository's existing tests pass, aside from explicitly documented pre-existing failures.
- Production prerequisites, monitoring, feature flags, audit records, and operator documentation exist.
- No live eBay mutation has occurred without explicit authorization.

## Expected implementation range

Use these estimates only for planning after inspecting the actual application:

- Technical spike and Sandbox proof: 2–3 weeks total from a bare integration surface
- Fixed-price, Lucas-created-listings MVP: 5–7 weeks
- Production-ready integration: 8–12 weeks for one experienced full-stack engineer
- Existing external listings: add roughly 2–4 weeks
- Auctions: add roughly 1–2 weeks
- Orders, tracking, and refunds: add roughly 2–3 weeks
- Multi-tenant seller accounts: add roughly 3–6 weeks

Update the estimate after Step 1 using evidence from the repository, inventory-data quality, existing AWS services, and the confirmed scope.
