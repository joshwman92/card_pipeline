# eBay Desktop Trial

The desktop trial creates and manages **Sandbox**, fixed-price, quantity-one listings from active L.U.C.A.S inventory. It uses eBay's Inventory API, uploads local card photos to eBay Picture Services with the Media API, and keeps eBay credentials out of the Tkinter process UI.

## Safety defaults

- `EBAY_ENV` defaults to `sandbox`.
- Production publication is blocked unless `EBAY_ALLOW_PRODUCTION_LISTINGS=1` is explicitly set.
- The browser callback carries a short-lived, single-use exchange code, never an access token, refresh token, or reusable broker connection token.
- The Windows broker protects refresh tokens with current-user DPAPI. Desktop connection tokens are also DPAPI-protected. Run the broker task as the same Windows account that created its credential store.
- L.U.C.A.S asks for confirmation after the offer is prepared and fee estimation has been attempted. Preparing an offer does not publish it.

## Prerequisites

1. Create an eBay Developers Program Sandbox keyset and Sandbox seller.
2. Opt the Sandbox seller into business policies and create payment, fulfillment, and return policies.
3. Create an enabled Inventory API location in the Sandbox seller account.
4. Configure the keyset's RuName/redirect to the broker callback. For a local browser trial, the callback route is `http://127.0.0.1:8788/ebay/callback` when eBay permits that registered URL. Otherwise expose only `/ebay*` through the documented HTTPS tunnel.
5. Copy the eBay block from `.env.example` into `.env` and fill in `EBAY_CLIENT_ID`, `EBAY_CLIENT_SECRET`, `EBAY_RUNAME`, and a long random `LUCAS_EBAY_BROKER_STATE_SECRET`.

Keep these values for the first trial:

```env
EBAY_ENV=sandbox
LUCAS_EBAY_BROKER_URL=http://127.0.0.1:8788/ebay
LUCAS_EBAY_BROKER_PUBLIC_URL=http://127.0.0.1:8788/ebay
EBAY_ALLOW_PRODUCTION_LISTINGS=0
```

## Start and connect

Start the broker in one PowerShell window:

```powershell
.\.venv\Scripts\python.exe .\ebay_broker_server.py
```

Start L.U.C.A.S normally in another window. Open the **eBay** tab and:

1. Click **Connect eBay** and approve the Sandbox seller consent page.
2. Return to L.U.C.A.S and click **Refresh**. The status should show `SANDBOX` and `credentials: windows-dpapi`.
3. Click **Load Seller Setup**, choose the existing payment, fulfillment, and return policies plus the enabled inventory location, then save.

## Prepare and publish a Sandbox listing

1. Ensure an active Inventory card has at least one linked local photo and a positive inventory/comp value.
2. Right-click the card and choose **List on eBay...**.
3. Review the title, price, leaf category, policy IDs, location, description, and item-specifics JSON.
4. Click **Prepare & Review Fees**. L.U.C.A.S uploads photos, creates or replaces the inventory item, reuses an existing unpublished offer by SKU when present, and creates an offer only when needed.
5. Review the returned fee estimate or warning. Click **Yes** only to create the Sandbox listing.

The **eBay** tab stores the SKU, offer ID, listing ID, state, price, timestamps, sanitized error, and the draft used for complete-replacement operations in `CARD_PIPELINE/ebay_listings.json`. It never stores access or refresh tokens there.

## Recovery and management

- **Refresh** reloads local listing state.
- **Reconcile** retrieves each saved offer by offer ID and repairs local published/unpublished status.
- **Withdraw** ends the listing while retaining the offer.
- **Republish** publishes a retained withdrawn offer after confirmation.
- Re-running **Prepare & Review Fees** for the same card/account/marketplace reuses its saved offer or finds the existing eBay offer by stable SKU; it does not intentionally create a duplicate.
- **Disconnect** invalidates the L.U.C.A.S broker connection and removes the local protected connection record. Listing/audit history remains.

## Trial limitations

- The draft pre-fills the three principal single-card categories and known trading-card condition descriptor IDs, but live eBay category/aspect/condition metadata remains authoritative. Review item-specifics and eBay validation errors before publishing.
- This trial does not import Seller Hub/Trading API listings, process orders, configure notifications, or automatically mark inventory sold.
- Production use still requires the security, metadata-cache, webhook/reconciliation scheduling, monitoring, account-deletion compliance, and rollout gates described by the reusable eBay integration skill.
