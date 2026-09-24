from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import ebay_api
import ebay_broker_server
import ebay_listing


class EbayListingTests(unittest.TestCase):
    def test_windows_dpapi_round_trip(self) -> None:
        if os.name != "nt":
            self.skipTest("Windows DPAPI test")
        protected = ebay_api.protect_secret("seller-secret")
        self.assertTrue(protected.startswith("dpapi:"))
        self.assertNotIn("seller-secret", protected)
        self.assertEqual(ebay_api.unprotect_secret(protected), "seller-secret")

    def test_stable_sku_is_repeatable_and_within_ebay_limit(self) -> None:
        record = {"inventory_key": "cert:" + "1234567890" * 10}
        first = ebay_listing.stable_ebay_sku(record)
        second = ebay_listing.stable_ebay_sku(record)
        self.assertEqual(first, second)
        self.assertLessEqual(len(first), 50)

    def test_graded_card_uses_ebay_trading_card_descriptors(self) -> None:
        condition, descriptors = ebay_listing.default_condition(
            {"grader": "PSA", "cert_number": "12345678", "card_title": "Test Card PSA 10"},
            "261328",
        )
        self.assertEqual(condition, "LIKE_NEW")
        self.assertEqual(descriptors[0], {"name": "27501", "values": ["275010"]})
        self.assertEqual(descriptors[1], {"name": "27502", "values": ["275020"]})
        self.assertEqual(descriptors[2], {"name": "27503", "additionalInfo": "12345678"})

    def test_ungraded_card_uses_near_mint_descriptor(self) -> None:
        condition, descriptors = ebay_listing.default_condition({"card_title": "Raw Pokemon Card"}, "183454")
        self.assertEqual(condition, "USED_VERY_GOOD")
        self.assertEqual(descriptors, [{"name": "40001", "values": ["400010"]}])

    def test_draft_validation_requires_publish_prerequisites(self) -> None:
        draft = ebay_listing.EbayListingDraft.from_inventory(
            {"inventory_key": "cert:1", "card_title": "Test PSA 10", "grader": "PSA", "cert_number": "1"}
        )
        errors = draft.validate()
        self.assertTrue(any("Price" in error for error in errors))
        self.assertTrue(any("Payment policy" in error for error in errors))
        self.assertTrue(any("image" in error.lower() for error in errors))

    def test_prepare_publish_and_duplicate_publish_are_idempotent(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            image = root / "front.jpg"
            image.write_bytes(b"not-a-real-image")
            store = ebay_listing.EbayListingStore(root / "listings.json")
            config = ebay_api.EbayConfig(env="sandbox", client_id="id", client_secret="secret", runame="runame")
            service = ebay_listing.EbayListingService(store, root / "tokens.json", config)
            draft = ebay_listing.EbayListingDraft(
                inventory_id="cert:123",
                sku="LUCAS-CERT-123",
                title="Test Card PSA 10",
                description="Test card",
                category_id="261328",
                price=49.99,
                payment_policy_id="PAY",
                fulfillment_policy_id="SHIP",
                return_policy_id="RETURN",
                merchant_location_key="MAIN",
                condition="LIKE_NEW",
                condition_descriptors=[{"name": "27501", "values": ["275010"]}, {"name": "27502", "values": ["275020"]}],
                aspects={"Sport": ["Baseball"]},
                image_paths=[str(image)],
            )

            calls = []

            def inventory_request(_config, _token, method, path, payload=None, marketplace_id="EBAY_US", timeout=45):
                calls.append((method, path, payload, marketplace_id))
                if method == "GET" and path.startswith("offer?"):
                    return {"offers": []}
                if method == "POST" and path == "offer":
                    return {"offerId": "OFFER-1"}
                if path == "offer/get_listing_fees":
                    return {"fees": [{"feeType": "INSERTION_FEE", "amount": {"value": "0.00", "currency": "USD"}}]}
                if path.endswith("/publish"):
                    return {"listingId": "LISTING-1"}
                return {}

            with (
                patch.object(service, "_token", return_value="access"),
                patch.object(ebay_listing, "ebay_upload_image_file", return_value={"imageUrl": "https://i.ebayimg.com/test.jpg"}),
                patch.object(ebay_listing, "ebay_inventory_request", side_effect=inventory_request),
            ):
                prepared = service.prepare(draft)
                self.assertEqual(prepared["status"], "READY")
                self.assertEqual(prepared["offer_id"], "OFFER-1")
                published = service.publish("cert:123")
                self.assertEqual(published["listing_id"], "LISTING-1")
                publish_call_count = sum(path.endswith("/publish") for _method, path, _payload, _market in calls)
                repeated = service.publish("cert:123")
                self.assertEqual(repeated["listing_id"], "LISTING-1")
                self.assertEqual(sum(path.endswith("/publish") for _method, path, _payload, _market in calls), publish_call_count)

            persisted = store.get("cert:123", "default", "EBAY_US")
            self.assertEqual(persisted["status"], "ACTIVE")
            self.assertNotIn("access", (root / "listings.json").read_text(encoding="utf-8"))

    def test_production_publish_requires_explicit_gate(self) -> None:
        with TemporaryDirectory() as tmp:
            store = ebay_listing.EbayListingStore(Path(tmp) / "listings.json")
            store.save_listing(
                {"inventory_id": "one", "account": "default", "marketplace_id": "EBAY_US", "offer_id": "offer", "status": "READY"},
                "prepare",
                "succeeded",
            )
            service = ebay_listing.EbayListingService(
                store,
                Path(tmp) / "tokens.json",
                ebay_api.EbayConfig(env="production", client_id="id", client_secret="secret", runame="runame"),
            )
            with patch.dict(os.environ, {}, clear=False):
                os.environ.pop("EBAY_ALLOW_PRODUCTION_LISTINGS", None)
                with self.assertRaisesRegex(ebay_listing.EbayListingError, "Production publishing is locked"):
                    service.publish("one")

    def test_publish_recovers_remote_success_without_republishing(self) -> None:
        with TemporaryDirectory() as tmp:
            store = ebay_listing.EbayListingStore(Path(tmp) / "listings.json")
            store.save_listing(
                {"inventory_id": "one", "account": "default", "marketplace_id": "EBAY_US", "offer_id": "offer", "status": "ERROR"},
                "publish",
                "failed",
            )
            service = ebay_listing.EbayListingService(
                store,
                Path(tmp) / "tokens.json",
                ebay_api.EbayConfig(env="sandbox", client_id="id", client_secret="secret", runame="runame"),
            )
            remote = {"status": "PUBLISHED", "listing": {"listingId": "LISTING-RECOVERED"}}
            with (
                patch.object(service, "_token", return_value="access"),
                patch.object(ebay_listing, "ebay_inventory_request", return_value=remote) as request,
            ):
                result = service.publish("one")
            self.assertEqual(result["listing_id"], "LISTING-RECOVERED")
            self.assertEqual(result["status"], "ACTIVE")
            request.assert_called_once()

    def test_error_sanitizer_redacts_credentials(self) -> None:
        safe = ebay_listing.sanitize_ebay_error(
            'Authorization: Bearer abc123 {"access_token":"secret-a","refresh_token":"secret-r"}'
        )
        self.assertNotIn("abc123", safe)
        self.assertNotIn("secret-a", safe)
        self.assertNotIn("secret-r", safe)
        self.assertIn("[REDACTED]", safe)

    def test_refreshed_direct_access_token_is_returned_unprotected(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "accounts.json"
            config = ebay_api.EbayConfig(env="sandbox", client_id="id", client_secret="secret", runame="runame")
            ebay_api.save_ebay_account_token(
                path,
                "default",
                config,
                {"refresh_token": "refresh", "access_token": "expired", "expires_in": 0},
            )
            with patch.object(
                ebay_api,
                "refresh_access_token",
                return_value={"access_token": "new-access", "expires_in": 7200},
            ):
                access = ebay_api.ebay_access_token_for_account(path, config)
            self.assertEqual(access, "new-access")
            self.assertNotIn("new-access", path.read_text(encoding="utf-8"))

    def test_account_environment_must_match_runtime(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "accounts.json"
            sandbox = ebay_api.EbayConfig(env="sandbox", client_id="id", client_secret="secret", runame="runame")
            production = ebay_api.EbayConfig(env="production", client_id="id", client_secret="secret", runame="runame")
            ebay_api.save_ebay_account_token(path, "default", sandbox, {"refresh_token": "refresh", "access_token": "access", "expires_in": 7200})
            with self.assertRaisesRegex(ebay_api.EbayOAuthError, "connected to sandbox"):
                ebay_api.ebay_access_token_for_account(path, production)

    def test_broker_ephemeral_values_expire(self) -> None:
        data = {
            "pending_states": {"old": {"created_at": 1}, "new": {"created_at": 999}},
            "exchange_codes": {"old": {"created_at": 1}, "new": {"created_at": 999}},
        }
        ebay_broker_server._prune_ephemeral(data, now=1000)
        self.assertEqual(set(data["pending_states"]), {"new"})
        self.assertEqual(set(data["exchange_codes"]), {"new"})


if __name__ == "__main__":
    unittest.main()
