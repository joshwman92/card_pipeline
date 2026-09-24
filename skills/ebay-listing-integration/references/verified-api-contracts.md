# Verified eBay contracts

Use these contracts as implementation starting points, then confirm them against the current official eBay reference and live Sandbox metadata. They were verified against official documentation in September 2026.

## Hosts and scopes

| Purpose | Sandbox | Production |
|---|---|---|
| Authorization | `https://auth.sandbox.ebay.com/oauth2/authorize` | `https://auth.ebay.com/oauth2/authorize` |
| Token/API root | `https://api.sandbox.ebay.com` | `https://api.ebay.com` |

Start with the base scope plus `sell.account` and `sell.inventory`. Add fulfillment or notification scopes only when those features exist.

## Fixed-price single-card sequence

1. Upload each local image with `POST /commerce/media/v1_beta/image/create_image_from_file`, `multipart/form-data`, field `image`. Capture response `imageUrl`, expiration, and the `Location` image-resource URI.
2. `PUT /sell/inventory/v1/inventory_item/{url_encoded_sku}`. This is a complete replacement.
3. Search `GET /sell/inventory/v1/offer?sku={url_encoded_sku}` before creating on retry.
4. Create with `POST /sell/inventory/v1/offer` or update the retained payload with `PUT /sell/inventory/v1/offer/{offerId}`.
5. Persist `offerId` before any publication attempt.
6. Attempt `POST /sell/inventory/v1/offer/get_listing_fees` with `{"offers":[{"offerId":"..."}]}`. Display estimates as non-guaranteed and allow a clear fee-unavailable state.
7. After separate user confirmation, retrieve the offer once. If it is already `PUBLISHED`, recover its `listing.listingId`; otherwise call `POST /sell/inventory/v1/offer/{offerId}/publish` and persist the returned `listingId`. This makes an ambiguous publish retry converge instead of duplicating work.
8. Withdraw with `POST /sell/inventory/v1/offer/{offerId}/withdraw`; republish the retained unpublished offer with the publish endpoint.

Minimum inventory item payload:

```json
{
  "availability": {"shipToLocationAvailability": {"quantity": 1}},
  "condition": "LIKE_NEW",
  "conditionDescriptors": [],
  "product": {
    "title": "...",
    "description": "...",
    "aspects": {"Name": ["Value"]},
    "imageUrls": ["https://i.ebayimg.com/..."]
  }
}
```

Minimum offer payload:

```json
{
  "sku": "LUCAS-...",
  "marketplaceId": "EBAY_US",
  "format": "FIXED_PRICE",
  "availableQuantity": 1,
  "categoryId": "261328",
  "listingDescription": "...",
  "listingDuration": "GTC",
  "merchantLocationKey": "...",
  "pricingSummary": {"price": {"currency": "USD", "value": "49.99"}},
  "listingPolicies": {
    "paymentPolicyId": "...",
    "fulfillmentPolicyId": "...",
    "returnPolicyId": "..."
  }
}
```

## Seller onboarding reads

Use the seller access token and marketplace query:

- `GET /sell/account/v1/payment_policy?marketplace_id=EBAY_US`
- `GET /sell/account/v1/fulfillment_policy?marketplace_id=EBAY_US`
- `GET /sell/account/v1/return_policy?marketplace_id=EBAY_US`
- `GET /sell/inventory/v1/location?limit=200`

Do not silently create or change policies. Let the seller choose existing compatible records; creation is a separately confirmed workflow.

## Trading-card categories and condition descriptors

Principal US leaf categories:

- Sports Trading Card Singles: `261328`
- CCG Individual Cards: `183454`
- Non-Sport Trading Card Singles: `183050`

Graded cards use Inventory condition `LIKE_NEW`, descriptor name `27501` for professional grader, `27502` for grade, and optional open-text certification descriptor `27503`. Common grader values include PSA `275010`, BGS `275013`, CGC `275015`, and SGC `275016`. Grade 10 is `275020`, 9.5 is `275021`, and 9 is `275022`; retrieve the remaining values from current metadata or the official condition-descriptor table.

Ungraded cards use Inventory condition `USED_VERY_GOOD`, descriptor name `40001`, and a category-supported value such as Near Mint or Better `400010`. CCG categories have different played-condition values from sports/non-sport categories. Never reuse one category's ungraded descriptor value blindly.

Use Metadata API `GET /sell/metadata/v1/marketplace/{marketplace_id}/get_item_condition_policies?filter=categoryIds:{category_id}` and Taxonomy API category-aspect metadata as the production authority.

## Error and response handling

- Successful replace/update calls can return `204` with no JSON body.
- Preserve eBay error IDs, parameters, and HTTP status in sanitized operation data; render field-specific user messages.
- A `401` may justify one coalesced refresh and retry. Do not repeatedly retry authorization, validation, category, aspect, policy, or condition errors.
- Treat `429` and transient `5xx` with bounded exponential backoff and jitter.
- Never log request authorization headers, OAuth bodies, callback query strings, raw token-store records, or broker exchange responses.
