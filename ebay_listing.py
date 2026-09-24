from __future__ import annotations

import hashlib
import json
import os
import re
import time
import urllib.parse
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable

from ebay_api import (
    EbayConfig,
    EbayOAuthError,
    ebay_access_token_for_account,
    ebay_account_request,
    ebay_inventory_request,
    ebay_token_store_path,
    ebay_upload_image_file,
)
from shared_state import atomic_write_json, read_json, shared_lock


TRADING_CARD_CATEGORIES = {"183050", "183454", "261328"}
GRADER_DESCRIPTOR_IDS = {
    "PSA": "275010",
    "BCCG": "275011",
    "BVG": "275012",
    "BGS": "275013",
    "CSG": "275014",
    "CGC": "275015",
    "SGC": "275016",
}
GRADE_DESCRIPTOR_IDS = {
    "10": "275020",
    "9.5": "275021",
    "9": "275022",
    "8.5": "275023",
    "8": "275024",
    "7.5": "275025",
    "7": "275026",
    "6.5": "275027",
    "6": "275028",
    "5.5": "275029",
    "5": "2750210",
    "4.5": "2750211",
    "4": "2750212",
    "3.5": "2750213",
    "3": "2750214",
    "2.5": "2750215",
    "2": "2750216",
    "1.5": "2750217",
    "1": "2750218",
    "AUTHENTIC": "2750219",
}


class EbayListingError(RuntimeError):
    pass


def _now() -> int:
    return int(time.time())


def _money(value: object) -> float | None:
    try:
        amount = float(str(value or "").replace("$", "").replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return round(amount, 2) if amount > 0 else None


def sanitize_ebay_error(error: object) -> str:
    """Keep actionable diagnostics without persisting OAuth credentials."""
    message = str(error or "")[:4000]
    patterns = (
        r"(?i)(authorization\s*:\s*bearer\s+)[^\s\"']+",
        r'(?i)("?(?:access_token|refresh_token|connection_token|client_secret)"?\s*[:=]\s*"?)[^\s\",}]+',
    )
    for pattern in patterns:
        message = re.sub(pattern, r"\1[REDACTED]", message)
    return message[:1000]


def ebay_listing_store_path(data_root: object = None) -> Path:
    configured = str(os.environ.get("EBAY_LISTING_STORE_PATH") or "").strip()
    if configured:
        return Path(configured).expanduser()
    root = Path(str(data_root or "")).expanduser() if str(data_root or "").strip() else Path(__file__).resolve().parent / "work"
    return root / "ebay_listings.json"


def stable_ebay_sku(record: dict[str, object]) -> str:
    stable = str(record.get("inventory_key") or record.get("item_id") or record.get("cert_number") or "").strip()
    if not stable:
        raise EbayListingError("The inventory card needs a stable inventory ID before it can be listed.")
    readable = re.sub(r"[^A-Za-z0-9_-]+", "-", stable).strip("-").upper() or "CARD"
    digest = hashlib.sha256(stable.encode("utf-8")).hexdigest()[:12].upper()
    prefix = f"LUCAS-{readable}"
    return prefix if len(prefix) <= 50 else f"{prefix[:37].rstrip('-')}-{digest}"


def guess_ebay_category(record: dict[str, object]) -> str:
    sport = str(record.get("sport") or "").strip().casefold()
    title = str(record.get("card_title") or "").casefold()
    ccg_terms = ("pokemon", "pokémon", "magic", "yugioh", "yu-gi-oh", "lorcana", "one piece")
    if any(term in sport or term in title for term in ccg_terms):
        return "183454"
    sports = ("baseball", "basketball", "football", "hockey", "soccer", "wrestling", "golf", "racing")
    return "261328" if any(term in sport or term in title for term in sports) else "183050"


def _grade_from_title(record: dict[str, object]) -> str:
    title = str(record.get("card_title") or "").upper()
    grader = str(record.get("grader") or "").strip().upper()
    if not grader:
        return ""
    match = re.search(rf"\b{re.escape(grader)}\s+(10|9\.5|9|8\.5|8|7\.5|7|6\.5|6|5\.5|5|4\.5|4|3\.5|3|2\.5|2|1\.5|1|AUTHENTIC)\b", title)
    return match.group(1) if match else ""


def default_condition(record: dict[str, object], category_id: str) -> tuple[str, list[dict[str, object]]]:
    grader = str(record.get("grader") or "").strip().upper()
    grade = _grade_from_title(record)
    cert = str(record.get("cert_number") or "").strip()[:30]
    if category_id in TRADING_CARD_CATEGORIES and grader and grade:
        grader_id = GRADER_DESCRIPTOR_IDS.get(grader, "2750123")
        grade_id = GRADE_DESCRIPTOR_IDS.get(grade)
        if grade_id:
            descriptors: list[dict[str, object]] = [
                {"name": "27501", "values": [grader_id]},
                {"name": "27502", "values": [grade_id]},
            ]
            if cert:
                descriptors.append({"name": "27503", "additionalInfo": cert})
            return "LIKE_NEW", descriptors
    if category_id in TRADING_CARD_CATEGORIES:
        return "USED_VERY_GOOD", [{"name": "40001", "values": ["400010"]}]
    return "USED_EXCELLENT", []


def default_aspects(record: dict[str, object]) -> dict[str, list[str]]:
    aspects: dict[str, list[str]] = {}
    sport = str(record.get("sport") or "").strip()
    grader = str(record.get("grader") or "").strip()
    if sport:
        aspects["Sport"] = [sport.title()]
    if grader:
        aspects["Professional Grader"] = [grader.upper()]
    return aspects


@dataclass
class EbayListingDraft:
    inventory_id: str
    sku: str
    title: str
    description: str
    category_id: str
    price: float
    payment_policy_id: str
    fulfillment_policy_id: str
    return_policy_id: str
    merchant_location_key: str
    condition: str
    condition_descriptors: list[dict[str, object]] = field(default_factory=list)
    aspects: dict[str, list[str]] = field(default_factory=dict)
    image_paths: list[str] = field(default_factory=list)
    image_urls: list[str] = field(default_factory=list)
    marketplace_id: str = "EBAY_US"
    currency: str = "USD"
    listing_duration: str = "GTC"
    account: str = "default"

    @classmethod
    def from_inventory(cls, record: dict[str, object], defaults: dict[str, object] | None = None) -> "EbayListingDraft":
        defaults = defaults or {}
        category_id = str(defaults.get("category_id") or guess_ebay_category(record)).strip()
        condition, descriptors = default_condition(record, category_id)
        title = str(record.get("card_title") or record.get("cert_number") or "Trading Card").strip()[:80]
        price = _money(defaults.get("price")) or _money(record.get("inventory_value")) or _money(record.get("card_ladder_comps_average")) or _money(record.get("card_ladder_value")) or 0.0
        return cls(
            inventory_id=str(record.get("inventory_key") or record.get("item_id") or record.get("cert_number") or "").strip(),
            sku=stable_ebay_sku(record),
            title=title,
            description=str(defaults.get("description") or f"{title}\n\nListed from L.U.C.A.S inventory.").strip(),
            category_id=category_id,
            price=price,
            payment_policy_id=str(defaults.get("payment_policy_id") or "").strip(),
            fulfillment_policy_id=str(defaults.get("fulfillment_policy_id") or "").strip(),
            return_policy_id=str(defaults.get("return_policy_id") or "").strip(),
            merchant_location_key=str(defaults.get("merchant_location_key") or "").strip(),
            condition=str(defaults.get("condition") or condition).strip(),
            condition_descriptors=list(defaults.get("condition_descriptors") or descriptors),
            aspects=dict(defaults.get("aspects") or default_aspects(record)),
            image_paths=[str(value) for value in defaults.get("image_paths") or record.get("photo_paths") or []],
            image_urls=[str(value) for value in defaults.get("image_urls") or []],
            marketplace_id=str(defaults.get("marketplace_id") or "EBAY_US").strip(),
            currency=str(defaults.get("currency") or "USD").strip(),
            account=str(defaults.get("account") or "default").strip() or "default",
        )

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not self.inventory_id:
            errors.append("Inventory ID is required.")
        if not self.sku or len(self.sku) > 50:
            errors.append("SKU is required and must be at most 50 characters.")
        if not self.title or len(self.title) > 80:
            errors.append("Title is required and must be at most 80 characters.")
        if not self.description:
            errors.append("Description is required.")
        if not self.category_id.isdigit():
            errors.append("A numeric leaf category ID is required.")
        if self.price <= 0:
            errors.append("Price must be greater than zero.")
        if not self.aspects:
            errors.append("At least one item-specific aspect is required; confirm live category metadata before publishing.")
        if not (self.image_paths or self.image_urls):
            errors.append("At least one local image or durable HTTPS image URL is required.")
        if any(url and not url.lower().startswith("https://") for url in self.image_urls):
            errors.append("Every hosted image URL must use HTTPS.")
        for label, value in (
            ("Payment policy", self.payment_policy_id),
            ("Fulfillment policy", self.fulfillment_policy_id),
            ("Return policy", self.return_policy_id),
            ("Merchant location", self.merchant_location_key),
        ):
            if not value:
                errors.append(f"{label} is required.")
        if self.category_id in TRADING_CARD_CATEGORIES and not self.condition_descriptors:
            errors.append("Trading-card condition descriptors are required.")
        return errors


class EbayListingStore:
    def __init__(self, path: Path, owner: dict[str, str] | None = None) -> None:
        self.path = Path(path)
        self.owner = owner or {}

    def load(self) -> dict[str, object]:
        data = read_json(self.path, {"version": 1, "listings": {}, "operations": []})
        if not isinstance(data, dict):
            data = {"version": 1, "listings": {}, "operations": []}
        if not isinstance(data.get("listings"), dict):
            data["listings"] = {}
        if not isinstance(data.get("operations"), list):
            data["operations"] = []
        return data

    @staticmethod
    def key(inventory_id: str, account: str, marketplace_id: str) -> str:
        raw = f"{inventory_id}\n{account}\n{marketplace_id}".encode("utf-8")
        return hashlib.sha256(raw).hexdigest()

    def get(self, inventory_id: str, account: str, marketplace_id: str) -> dict[str, object]:
        record = self.load()["listings"].get(self.key(inventory_id, account, marketplace_id), {})
        return dict(record) if isinstance(record, dict) else {}

    def save_listing(self, record: dict[str, object], action: str, state: str, error: str = "") -> dict[str, object]:
        key = self.key(str(record.get("inventory_id") or ""), str(record.get("account") or "default"), str(record.get("marketplace_id") or "EBAY_US"))
        operation = {
            "id": uuid.uuid4().hex,
            "action": action,
            "state": state,
            "listing_key": key,
            "created_at": _now(),
            "error": str(error or "")[:1000],
        }
        with shared_lock(self.path.parent, f"ebay-{self.path.name}", self.owner):
            data = self.load()
            listings = data["listings"]
            previous = listings.get(key) if isinstance(listings.get(key), dict) else {}
            saved = {**previous, **record, "listing_key": key, "updated_at": _now(), "last_operation": operation}
            saved.setdefault("created_at", _now())
            listings[key] = saved
            operations = data["operations"]
            operations.append(operation)
            data["operations"] = operations[-500:]
            atomic_write_json(self.path, data)
        return saved

    def all(self) -> list[dict[str, object]]:
        listings = self.load().get("listings", {})
        values = listings.values() if isinstance(listings, dict) else []
        return sorted((dict(value) for value in values if isinstance(value, dict)), key=lambda row: int(row.get("updated_at") or 0), reverse=True)


class EbayListingService:
    def __init__(
        self,
        store: EbayListingStore,
        token_store: Path | None = None,
        config: EbayConfig | None = None,
        image_resolver: Callable[[str], Path] | None = None,
    ) -> None:
        self.store = store
        self.token_store = token_store or ebay_token_store_path()
        self.config = config or EbayConfig.from_env()
        self.image_resolver = image_resolver or (lambda value: Path(value).expanduser())

    def _token(self, account: str) -> str:
        return ebay_access_token_for_account(self.token_store, self.config, account)

    def discover_seller_setup(self, account: str = "default", marketplace_id: str = "EBAY_US") -> dict[str, object]:
        token = self._token(account)
        query = urllib.parse.urlencode({"marketplace_id": marketplace_id})
        result = {
            "payment_policies": ebay_account_request(self.config, token, "GET", f"payment_policy?{query}").get("paymentPolicies", []),
            "fulfillment_policies": ebay_account_request(self.config, token, "GET", f"fulfillment_policy?{query}").get("fulfillmentPolicies", []),
            "return_policies": ebay_account_request(self.config, token, "GET", f"return_policy?{query}").get("returnPolicies", []),
            "locations": ebay_inventory_request(self.config, token, "GET", "location?limit=200", marketplace_id=marketplace_id).get("locations", []),
        }
        return result

    def prepare(self, draft: EbayListingDraft) -> dict[str, object]:
        errors = draft.validate()
        if errors:
            raise EbayListingError("\n".join(errors))
        existing = self.store.get(draft.inventory_id, draft.account, draft.marketplace_id)
        if str(existing.get("status") or "").upper() == "ACTIVE":
            raise EbayListingError(f"This inventory card is already listed as {existing.get('listing_id') or existing.get('offer_id')}.")
        token = self._token(draft.account)
        record = {
            **existing,
            "inventory_id": draft.inventory_id,
            "account": draft.account,
            "marketplace_id": draft.marketplace_id,
            "sku": draft.sku,
            "title": draft.title,
            "price": draft.price,
            "currency": draft.currency,
            "draft": asdict(draft),
            "status": "PREPARING",
            "operation_state": "PREPARING",
            "attempt_count": int(existing.get("attempt_count") or 0) + 1,
        }
        self.store.save_listing(record, "prepare", "started")
        try:
            image_urls = list(draft.image_urls)
            if not image_urls:
                for raw_path in draft.image_paths:
                    result = ebay_upload_image_file(self.config, token, self.image_resolver(raw_path), draft.marketplace_id)
                    image_urls.append(str(result.get("imageUrl") or ""))
            inventory_payload: dict[str, object] = {
                "availability": {"shipToLocationAvailability": {"quantity": 1}},
                "condition": draft.condition,
                "product": {
                    "title": draft.title,
                    "description": draft.description,
                    "aspects": draft.aspects,
                    "imageUrls": image_urls,
                },
            }
            if draft.condition_descriptors:
                inventory_payload["conditionDescriptors"] = draft.condition_descriptors
            ebay_inventory_request(
                self.config,
                token,
                "PUT",
                "inventory_item/" + urllib.parse.quote(draft.sku, safe=""),
                inventory_payload,
                draft.marketplace_id,
            )
            offer_id = str(existing.get("offer_id") or "").strip()
            if not offer_id:
                offers_result = ebay_inventory_request(
                    self.config,
                    token,
                    "GET",
                    "offer?" + urllib.parse.urlencode({"sku": draft.sku}),
                    marketplace_id=draft.marketplace_id,
                )
                for offer in offers_result.get("offers", []) if isinstance(offers_result, dict) else []:
                    if isinstance(offer, dict) and str(offer.get("marketplaceId") or "") == draft.marketplace_id:
                        offer_id = str(offer.get("offerId") or "")
                        if offer_id:
                            break
            offer_payload = {
                "sku": draft.sku,
                "marketplaceId": draft.marketplace_id,
                "format": "FIXED_PRICE",
                "availableQuantity": 1,
                "categoryId": draft.category_id,
                "listingDescription": draft.description,
                "listingDuration": draft.listing_duration,
                "merchantLocationKey": draft.merchant_location_key,
                "pricingSummary": {"price": {"currency": draft.currency, "value": f"{draft.price:.2f}"}},
                "listingPolicies": {
                    "paymentPolicyId": draft.payment_policy_id,
                    "fulfillmentPolicyId": draft.fulfillment_policy_id,
                    "returnPolicyId": draft.return_policy_id,
                },
                "includeCatalogProductDetails": False,
            }
            if offer_id:
                ebay_inventory_request(self.config, token, "PUT", f"offer/{urllib.parse.quote(offer_id, safe='')}", offer_payload, draft.marketplace_id)
            else:
                created = ebay_inventory_request(self.config, token, "POST", "offer", offer_payload, draft.marketplace_id)
                offer_id = str(created.get("offerId") or "").strip()
                if not offer_id:
                    raise EbayListingError("eBay created the inventory item but did not return an offer ID.")
            fees: dict[str, object] = {}
            fee_error = ""
            try:
                fees = ebay_inventory_request(
                    self.config,
                    token,
                    "POST",
                    "offer/get_listing_fees",
                    {"offers": [{"offerId": offer_id}]},
                    draft.marketplace_id,
                )
            except EbayOAuthError as error:
                fee_error = sanitize_ebay_error(error)
            record.update(
                {
                    "offer_id": offer_id,
                    "image_urls": image_urls,
                    "fees": fees,
                    "fee_error": fee_error,
                    "status": "READY",
                    "operation_state": "PREPARED",
                    "last_error": "",
                }
            )
            return self.store.save_listing(record, "prepare", "succeeded")
        except Exception as error:
            safe_error = sanitize_ebay_error(error)
            record.update({"status": "ERROR", "operation_state": "PREPARE_FAILED", "last_error": safe_error})
            self.store.save_listing(record, "prepare", "failed", safe_error)
            raise

    def publish(self, inventory_id: str, account: str = "default", marketplace_id: str = "EBAY_US") -> dict[str, object]:
        record = self.store.get(inventory_id, account, marketplace_id)
        offer_id = str(record.get("offer_id") or "").strip()
        if not offer_id:
            raise EbayListingError("Prepare the eBay offer before publishing it.")
        if self.config.env == "production" and not self.config.production_publish_allowed():
            raise EbayListingError("Production publishing is locked. Set EBAY_ALLOW_PRODUCTION_LISTINGS=1 only for an explicitly approved live listing.")
        if str(record.get("status") or "").upper() == "ACTIVE" and record.get("listing_id"):
            return record
        token = self._token(account)
        self.store.save_listing({**record, "status": "PUBLISHING", "operation_state": "PUBLISHING"}, "publish", "started")
        try:
            remote = ebay_inventory_request(
                self.config,
                token,
                "GET",
                f"offer/{urllib.parse.quote(offer_id, safe='')}",
                marketplace_id=marketplace_id,
            )
            remote_listing = remote.get("listing") if isinstance(remote.get("listing"), dict) else {}
            remote_status = str(remote.get("status") or remote_listing.get("listingStatus") or "").upper()
            listing_id = str(remote_listing.get("listingId") or record.get("listing_id") or "").strip()
            if remote_status != "PUBLISHED" or not listing_id:
                result = ebay_inventory_request(self.config, token, "POST", f"offer/{urllib.parse.quote(offer_id, safe='')}/publish", marketplace_id=marketplace_id)
                listing_id = str(result.get("listingId") or listing_id).strip()
            if not listing_id:
                raise EbayListingError("eBay did not return a listing ID after publishing.")
            listing_host = "sandbox.ebay.com" if self.config.env == "sandbox" else "www.ebay.com"
            updated = {
                **record,
                "listing_id": listing_id,
                "listing_url": f"https://{listing_host}/itm/{listing_id}",
                "status": "ACTIVE",
                "ebay_status": "PUBLISHED",
                "operation_state": "PUBLISHED",
                "last_synced_at": _now(),
                "last_error": "",
            }
            return self.store.save_listing(updated, "publish", "succeeded")
        except Exception as error:
            safe_error = sanitize_ebay_error(error)
            self.store.save_listing({**record, "status": "ERROR", "operation_state": "PUBLISH_FAILED", "last_error": safe_error}, "publish", "failed", safe_error)
            raise

    def withdraw(self, inventory_id: str, account: str = "default", marketplace_id: str = "EBAY_US") -> dict[str, object]:
        record = self.store.get(inventory_id, account, marketplace_id)
        offer_id = str(record.get("offer_id") or "").strip()
        if not offer_id:
            raise EbayListingError("No eBay offer is linked to this inventory card.")
        token = self._token(account)
        ebay_inventory_request(self.config, token, "POST", f"offer/{urllib.parse.quote(offer_id, safe='')}/withdraw", marketplace_id=marketplace_id)
        updated = {**record, "status": "WITHDRAWN", "ebay_status": "UNPUBLISHED", "operation_state": "WITHDRAWN", "last_synced_at": _now(), "last_error": ""}
        return self.store.save_listing(updated, "withdraw", "succeeded")

    def republish(self, inventory_id: str, account: str = "default", marketplace_id: str = "EBAY_US") -> dict[str, object]:
        record = self.store.get(inventory_id, account, marketplace_id)
        if str(record.get("status") or "").upper() != "WITHDRAWN":
            raise EbayListingError("Only a withdrawn eBay offer can be republished.")
        return self.publish(inventory_id, account, marketplace_id)

    def reconcile(self, account: str = "default") -> list[dict[str, object]]:
        token = self._token(account)
        updated_rows: list[dict[str, object]] = []
        for record in self.store.all():
            if str(record.get("account") or "default") != account or not record.get("offer_id"):
                continue
            marketplace_id = str(record.get("marketplace_id") or "EBAY_US")
            try:
                remote = ebay_inventory_request(self.config, token, "GET", f"offer/{urllib.parse.quote(str(record['offer_id']), safe='')}", marketplace_id=marketplace_id)
                listing = remote.get("listing") if isinstance(remote.get("listing"), dict) else {}
                listing_id = str(listing.get("listingId") or record.get("listing_id") or "")
                remote_status = str(remote.get("status") or listing.get("listingStatus") or "UNKNOWN")
                status = "ACTIVE" if remote_status.upper() == "PUBLISHED" else "WITHDRAWN" if remote_status.upper() == "UNPUBLISHED" else str(record.get("status") or remote_status)
                record.update({"listing_id": listing_id, "ebay_status": remote_status, "status": status, "last_synced_at": _now(), "last_error": ""})
                updated_rows.append(self.store.save_listing(record, "reconcile", "succeeded"))
            except Exception as error:
                safe_error = sanitize_ebay_error(error)
                record.update({"last_error": safe_error, "operation_state": "RECONCILE_FAILED"})
                updated_rows.append(self.store.save_listing(record, "reconcile", "failed", safe_error))
        return updated_rows
