from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes
import json
import mimetypes
import os
import re
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path


PRODUCTION_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
SANDBOX_TOKEN_URL = "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
PRODUCTION_AUTHORIZE_URL = "https://auth.ebay.com/oauth2/authorize"
SANDBOX_AUTHORIZE_URL = "https://auth.sandbox.ebay.com/oauth2/authorize"
PRODUCTION_INVENTORY_URL = "https://api.ebay.com/sell/inventory/v1"
SANDBOX_INVENTORY_URL = "https://api.sandbox.ebay.com/sell/inventory/v1"
PRODUCTION_ACCOUNT_URL = "https://api.ebay.com/sell/account/v1"
SANDBOX_ACCOUNT_URL = "https://api.sandbox.ebay.com/sell/account/v1"
PRODUCTION_MEDIA_URL = "https://api.ebay.com/commerce/media/v1_beta"
SANDBOX_MEDIA_URL = "https://api.sandbox.ebay.com/commerce/media/v1_beta"
CONNECT_STATE_KIND = "lucas_ebay_connect"
DEFAULT_BROKER_URL = "https://lucas.mikeyscards.com/ebay"
DEFAULT_SCOPES = (
    "https://api.ebay.com/oauth/api_scope",
    "https://api.ebay.com/oauth/api_scope/sell.account",
    "https://api.ebay.com/oauth/api_scope/sell.inventory",
)


class EbayOAuthError(RuntimeError):
    pass


def ebay_broker_url() -> str:
    return str(os.environ.get("LUCAS_EBAY_BROKER_URL") or DEFAULT_BROKER_URL).strip().rstrip("/")


@dataclass(frozen=True)
class EbayConfig:
    env: str
    client_id: str
    client_secret: str
    runame: str
    scopes: tuple[str, ...] = ()

    @property
    def token_url(self) -> str:
        return SANDBOX_TOKEN_URL if self.env.lower() == "sandbox" else PRODUCTION_TOKEN_URL

    @property
    def authorize_url(self) -> str:
        return SANDBOX_AUTHORIZE_URL if self.env.lower() == "sandbox" else PRODUCTION_AUTHORIZE_URL

    @property
    def inventory_url(self) -> str:
        return SANDBOX_INVENTORY_URL if self.env.lower() == "sandbox" else PRODUCTION_INVENTORY_URL

    @property
    def account_url(self) -> str:
        return SANDBOX_ACCOUNT_URL if self.env.lower() == "sandbox" else PRODUCTION_ACCOUNT_URL

    @property
    def media_url(self) -> str:
        return SANDBOX_MEDIA_URL if self.env.lower() == "sandbox" else PRODUCTION_MEDIA_URL

    @classmethod
    def from_env(cls) -> "EbayConfig":
        scopes = tuple(str(os.environ.get("EBAY_OAUTH_SCOPES") or "").split()) or DEFAULT_SCOPES
        return cls(
            env=str(os.environ.get("EBAY_ENV") or "sandbox").strip().lower() or "sandbox",
            client_id=str(os.environ.get("EBAY_CLIENT_ID") or "").strip(),
            client_secret=str(os.environ.get("EBAY_CLIENT_SECRET") or "").strip(),
            runame=str(os.environ.get("EBAY_RUNAME") or "").strip(),
            scopes=scopes,
        )

    def validate(self) -> None:
        missing = []
        if not self.client_id:
            missing.append("EBAY_CLIENT_ID")
        if not self.client_secret:
            missing.append("EBAY_CLIENT_SECRET")
        if not self.runame:
            missing.append("EBAY_RUNAME")
        if missing:
            raise EbayOAuthError(f"Missing eBay config value(s): {', '.join(missing)}")

    def production_publish_allowed(self) -> bool:
        return self.env != "production" or str(os.environ.get("EBAY_ALLOW_PRODUCTION_LISTINGS") or "").strip() == "1"


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob(value: bytes) -> tuple[_DataBlob, object]:
    buffer = ctypes.create_string_buffer(value)
    return _DataBlob(len(value), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def protect_secret(value: object) -> str:
    """Protect a local credential with Windows DPAPI for the current user."""
    secret = str(value or "")
    if not secret:
        return ""
    if os.name != "nt":
        raise EbayOAuthError(
            "Secure local eBay credential storage requires Windows DPAPI. "
            "Web/server implementations must use their managed secret store."
        )
    source, source_buffer = _blob(secret.encode("utf-8"))
    output = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptProtectData(ctypes.byref(source), None, None, None, None, 0x1, ctypes.byref(output)):
        raise EbayOAuthError("Windows could not protect the eBay credential.")
    try:
        encrypted = ctypes.string_at(output.pbData, output.cbData)
    finally:
        kernel32.LocalFree(output.pbData)
        del source_buffer
    return "dpapi:" + base64.urlsafe_b64encode(encrypted).decode("ascii")


def unprotect_secret(value: object) -> str:
    protected = str(value or "").strip()
    if not protected:
        return ""
    if not protected.startswith("dpapi:"):
        # Read-only compatibility for pre-encryption trial records. Callers rewrite
        # these records with protect_secret after the next successful refresh.
        return protected
    if os.name != "nt":
        raise EbayOAuthError("This eBay credential is protected for a Windows user and cannot be opened here.")
    try:
        encrypted = base64.urlsafe_b64decode(protected[6:].encode("ascii"))
    except Exception as error:
        raise EbayOAuthError("The stored eBay credential is invalid.") from error
    source, source_buffer = _blob(encrypted)
    output = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptUnprotectData(ctypes.byref(source), None, None, None, None, 0x1, ctypes.byref(output)):
        raise EbayOAuthError("Windows could not unlock the stored eBay credential for this user.")
    try:
        clear = ctypes.string_at(output.pbData, output.cbData).decode("utf-8")
    finally:
        kernel32.LocalFree(output.pbData)
        del source_buffer
    return clear


def mask_token(value: object) -> str:
    token = str(value or "").strip()
    if len(token) <= 16:
        return "*" * len(token)
    return f"{token[:8]}...{token[-6:]}"


def _basic_auth_header(config: EbayConfig) -> str:
    raw = f"{config.client_id}:{config.client_secret}".encode("utf-8")
    return "Basic " + base64.b64encode(raw).decode("ascii")


def _oauth_post(config: EbayConfig, payload: dict[str, object], timeout: int = 45) -> dict[str, object]:
    config.validate()
    data = urllib.parse.urlencode(payload).encode("utf-8")
    request = urllib.request.Request(
        config.token_url,
        data=data,
        headers={
            "Authorization": _basic_auth_header(config),
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(body or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        return json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"eBay returned invalid JSON: {raw[:200]}") from error


def exchange_authorization_code(config: EbayConfig, authorization_code: str) -> dict[str, object]:
    code = str(authorization_code or "").strip()
    if not code:
        raise EbayOAuthError("Missing eBay authorization code.")
    return _oauth_post(
        config,
        {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": config.runame,
        },
    )


def encode_connect_state(account: str = "", profile: str = "", local_callback: str = "") -> str:
    payload = {
        "kind": CONNECT_STATE_KIND,
        "account": str(account or "default").strip() or "default",
        "profile": str(profile or "").strip().lower(),
        "nonce": secrets.token_urlsafe(16),
        "created_at": int(time.time()),
    }
    callback = str(local_callback or "").strip()
    if callback:
        payload["local_callback"] = callback
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_connect_state(value: str) -> dict[str, object]:
    token = str(value or "").strip()
    if not token:
        return {}
    padding = "=" * (-len(token) % 4)
    try:
        payload = json.loads(base64.urlsafe_b64decode((token + padding).encode("ascii")).decode("utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict) or payload.get("kind") != CONNECT_STATE_KIND:
        return {}
    return payload


def build_authorization_url(config: EbayConfig, state: str, scopes: tuple[str, ...] | None = None) -> str:
    config.validate()
    scope_values = scopes if scopes is not None else config.scopes
    if not scope_values:
        raise EbayOAuthError("Missing eBay OAuth scopes.")
    return config.authorize_url + "?" + urllib.parse.urlencode(
        {
            "client_id": config.client_id,
            "redirect_uri": config.runame,
            "response_type": "code",
            "scope": " ".join(scope_values),
            "state": state,
        }
    )


def refresh_access_token(config: EbayConfig, refresh_token: str, scopes: tuple[str, ...] | None = None) -> dict[str, object]:
    token = str(refresh_token or "").strip()
    if not token:
        raise EbayOAuthError("Missing EBAY_REFRESH_TOKEN.")
    payload: dict[str, object] = {
        "grant_type": "refresh_token",
        "refresh_token": token,
    }
    scope_values = scopes if scopes is not None else config.scopes
    if scope_values:
        payload["scope"] = " ".join(scope_values)
    return _oauth_post(config, payload)


def ebay_inventory_request(
    config: EbayConfig,
    access_token: str,
    method: str,
    path: str,
    payload: dict[str, object] | None = None,
    marketplace_id: str = "EBAY_US",
    timeout: int = 45,
) -> dict[str, object]:
    token = str(access_token or "").strip()
    if not token:
        raise EbayOAuthError("Missing eBay access token.")
    url = config.inventory_url.rstrip("/") + "/" + path.lstrip("/")
    body = json.dumps(payload or {}).encode("utf-8") if payload is not None else None
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Content-Language": "en-US",
            "X-EBAY-C-MARKETPLACE-ID": marketplace_id,
        },
        method=method.upper(),
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        raw_error = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(raw_error or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        return json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"eBay returned invalid JSON: {raw[:200]}") from error


def ebay_account_request(
    config: EbayConfig,
    access_token: str,
    method: str,
    path: str,
    payload: dict[str, object] | None = None,
    timeout: int = 45,
) -> dict[str, object]:
    token = str(access_token or "").strip()
    if not token:
        raise EbayOAuthError("Missing eBay access token.")
    url = config.account_url.rstrip("/") + "/" + path.lstrip("/")
    body = json.dumps(payload or {}).encode("utf-8") if payload is not None else None
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Content-Language": "en-US",
        },
        method=method.upper(),
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        raw_error = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(raw_error or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        return json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"eBay returned invalid JSON: {raw[:200]}") from error


def exchange_broker_code(broker_url: str, exchange_code: str, timeout: int = 45) -> dict[str, object]:
    broker = str(broker_url or ebay_broker_url()).strip().rstrip("/")
    code = str(exchange_code or "").strip()
    if not broker or not code:
        raise EbayOAuthError("Missing LUCAS eBay broker exchange details.")
    request = urllib.request.Request(
        broker + "/connection/exchange",
        data=json.dumps({"exchange_code": code}).encode("utf-8"),
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(body or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        payload = json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"LUCAS eBay broker returned invalid JSON: {raw[:200]}") from error
    if not isinstance(payload, dict) or not str(payload.get("connection_token") or "").strip():
        raise EbayOAuthError(str(payload.get("error") or "The eBay broker did not return a connection token."))
    return payload


def ebay_upload_image_file(
    config: EbayConfig,
    access_token: str,
    path: Path,
    marketplace_id: str = "EBAY_US",
    timeout: int = 90,
) -> dict[str, object]:
    """Upload one local image to EPS using Media API multipart/form-data."""
    token = str(access_token or "").strip()
    image_path = Path(path)
    if not token:
        raise EbayOAuthError("Missing eBay access token.")
    if not image_path.is_file():
        raise EbayOAuthError(f"eBay image does not exist: {image_path}")
    if image_path.stat().st_size > 12 * 1024 * 1024:
        raise EbayOAuthError(f"eBay image exceeds the 12 MB trial limit: {image_path.name}")
    boundary = "----LUCAS-eBay-" + secrets.token_hex(16)
    mime_type = mimetypes.guess_type(image_path.name)[0] or "application/octet-stream"
    safe_filename = re.sub(r"[^A-Za-z0-9._-]+", "_", image_path.name).strip("._") or "card-image"
    body = b"".join(
        (
            f"--{boundary}\r\n".encode("ascii"),
            f'Content-Disposition: form-data; name="image"; filename="{safe_filename}"\r\n'.encode("utf-8"),
            f"Content-Type: {mime_type}\r\n\r\n".encode("ascii"),
            image_path.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode("ascii"),
        )
    )
    request = urllib.request.Request(
        config.media_url.rstrip("/") + "/image/create_image_from_file",
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "Accept": "application/json",
            "X-EBAY-C-MARKETPLACE-ID": marketplace_id,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            location = str(response.headers.get("Location") or "")
    except urllib.error.HTTPError as error:
        raw_error = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(raw_error or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        payload = json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"eBay Media API returned invalid JSON: {raw[:200]}") from error
    if not isinstance(payload, dict):
        payload = {}
    payload["location"] = location
    image_url = str(payload.get("imageUrl") or "").strip()
    if not image_url:
        raise EbayOAuthError("eBay accepted the image but did not return an EPS image URL.")
    return payload


def ebay_token_store_path(data_root: object = None) -> Path:
    configured = str(os.environ.get("EBAY_TOKEN_STORE_PATH") or "").strip()
    if configured:
        return Path(configured).expanduser()
    root = Path(str(data_root or "")).expanduser() if str(data_root or "").strip() else Path(__file__).resolve().parent / "work"
    return root / "ebay_accounts.json"


def load_ebay_accounts(path: Path) -> dict[str, object]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"accounts": {}}
    if not isinstance(data, dict):
        return {"accounts": {}}
    accounts = data.get("accounts")
    if not isinstance(accounts, dict):
        data["accounts"] = {}
    return data


def save_ebay_account_token(
    path: Path,
    account: str,
    config: EbayConfig,
    token_result: dict[str, object],
    scopes: tuple[str, ...] | None = None,
) -> dict[str, object]:
    refresh_token = str(token_result.get("refresh_token") or "").strip()
    if not refresh_token:
        raise EbayOAuthError("eBay accepted the authorization code but did not return a refresh_token.")
    access_token = str(token_result.get("access_token") or "").strip()
    account_key = str(account or "default").strip() or "default"
    now = int(time.time())
    access_expires_in = token_result.get("expires_in")
    try:
        access_expires_at = now + int(access_expires_in)
    except (TypeError, ValueError):
        access_expires_at = None
    data = load_ebay_accounts(path)
    accounts = data.setdefault("accounts", {})
    if not isinstance(accounts, dict):
        accounts = {}
        data["accounts"] = accounts
    previous = accounts.get(account_key) if isinstance(accounts.get(account_key), dict) else {}
    record = {
        **previous,
        "account": account_key,
        "env": config.env,
        "client_id": config.client_id,
        "runame": config.runame,
        "scopes": list(scopes if scopes is not None else config.scopes),
        "refresh_token": protect_secret(refresh_token),
        "refresh_token_expires_in": token_result.get("refresh_token_expires_in"),
        "access_token": protect_secret(access_token) if access_token else "",
        "access_token_expires_in": token_result.get("expires_in"),
        "access_token_expires_at": access_expires_at,
        "connected_at": previous.get("connected_at") or now,
        "updated_at": now,
    }
    accounts[account_key] = record
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)
    return record


def save_ebay_broker_account(
    path: Path,
    account: str,
    broker_url: str,
    connection_token: str,
    seller_username: str = "",
    marketplace_id: str = "EBAY_US",
    environment: str = "",
) -> dict[str, object]:
    token = str(connection_token or "").strip()
    if not token:
        raise EbayOAuthError("Missing LUCAS eBay connection token.")
    account_key = str(account or "default").strip() or "default"
    now = int(time.time())
    data = load_ebay_accounts(path)
    accounts = data.setdefault("accounts", {})
    if not isinstance(accounts, dict):
        accounts = {}
        data["accounts"] = accounts
    previous = accounts.get(account_key) if isinstance(accounts.get(account_key), dict) else {}
    normalized_environment = str(environment or os.environ.get("EBAY_ENV") or "sandbox").strip().lower()
    if normalized_environment not in {"sandbox", "production"}:
        raise EbayOAuthError("The eBay broker returned an unknown environment.")
    record = {
        **previous,
        "account": account_key,
        "env": normalized_environment,
        "connection_mode": "broker",
        "broker_url": str(broker_url or ebay_broker_url()).strip().rstrip("/"),
        "connection_token": protect_secret(token),
        "seller_username": str(seller_username or "").strip(),
        "marketplace_id": str(marketplace_id or "EBAY_US").strip() or "EBAY_US",
        "connected_at": previous.get("connected_at") or now,
        "updated_at": now,
    }
    accounts[account_key] = record
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)
    return record


def ebay_account_record(path: Path, account: str = "default") -> dict[str, object]:
    data = load_ebay_accounts(path)
    accounts = data.get("accounts") if isinstance(data, dict) else {}
    if not isinstance(accounts, dict):
        return {}
    key = str(account or "default").strip() or "default"
    record = accounts.get(key)
    return dict(record) if isinstance(record, dict) else {}


def ebay_broker_access_token(record: dict[str, object], timeout: int = 45) -> str:
    broker = str(record.get("broker_url") or ebay_broker_url()).strip().rstrip("/")
    connection_token = unprotect_secret(record.get("connection_token"))
    if not broker or not connection_token:
        raise EbayOAuthError("Connect eBay first so LUCAS has a seller connection token.")
    request = urllib.request.Request(
        broker + "/token",
        data=json.dumps({"connection_token": connection_token}).encode("utf-8"),
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(body or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        payload = json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"LUCAS eBay broker returned invalid JSON: {raw[:200]}") from error
    access_token = str(payload.get("access_token") or "").strip()
    if not access_token:
        raise EbayOAuthError("LUCAS eBay broker did not return an access token.")
    return access_token


def ebay_access_token_for_account(path: Path, config: EbayConfig, account: str = "default", allow_env_fallback: bool = False) -> str:
    record = ebay_account_record(path, account)
    record_env = str(record.get("env") or "").strip().lower()
    if record_env and record_env != config.env.lower():
        raise EbayOAuthError(
            f"The saved eBay account is connected to {record_env}, but LUCAS is configured for {config.env.lower()}. "
            "Switch EBAY_ENV or reconnect the matching account."
        )
    if str(record.get("connection_mode") or "").strip().lower() == "broker":
        return ebay_broker_access_token(record)
    stored_refresh = record.get("refresh_token")
    refresh_token = unprotect_secret(stored_refresh) if stored_refresh else str((os.environ.get("EBAY_REFRESH_TOKEN") if allow_env_fallback else "") or "").strip()
    if not refresh_token:
        raise EbayOAuthError("Connect eBay first so LUCAS has a seller refresh token.")
    now = int(time.time())
    stored_access = record.get("access_token")
    access_token = unprotect_secret(stored_access) if stored_access else str((os.environ.get("EBAY_ACCESS_TOKEN") if allow_env_fallback else "") or "").strip()
    expires_at = record.get("access_token_expires_at")
    try:
        if access_token and int(expires_at) > now + 300:
            return access_token
    except (TypeError, ValueError):
        pass
    token_result = refresh_access_token(config, refresh_token)
    updated = save_ebay_account_token(path, str(record.get("account") or account or "default"), config, {**token_result, "refresh_token": refresh_token})
    return unprotect_secret(updated.get("access_token"))


def ebay_account_status(path: Path) -> dict[str, object]:
    data = load_ebay_accounts(path)
    accounts = data.get("accounts") if isinstance(data, dict) else {}
    result: list[dict[str, object]] = []
    if isinstance(accounts, dict):
        for key, record in sorted(accounts.items()):
            if not isinstance(record, dict):
                continue
            result.append(
                {
                    "account": key,
                    "env": record.get("env", ""),
                    "client_id": record.get("client_id", ""),
                    "connected_at": record.get("connected_at"),
                    "updated_at": record.get("updated_at"),
                    "connection_mode": record.get("connection_mode") or "direct",
                    "seller_username": record.get("seller_username", ""),
                    "credential_storage": "windows-dpapi"
                    if str(record.get("refresh_token") or record.get("connection_token") or "").startswith("dpapi:")
                    else "legacy-plaintext" if record.get("refresh_token") or record.get("connection_token") else "none",
                    "refresh_token": "stored" if record.get("refresh_token") else "",
                    "connection_token": "stored" if record.get("connection_token") else "",
                    "scopes": record.get("scopes", []),
                }
            )
    return {"ok": True, "accounts": result}


def disconnect_ebay_account(path: Path, account: str = "default", timeout: int = 45) -> bool:
    data = load_ebay_accounts(path)
    accounts = data.get("accounts") if isinstance(data.get("accounts"), dict) else {}
    key = str(account or "default").strip() or "default"
    record = accounts.get(key) if isinstance(accounts, dict) else None
    if not isinstance(record, dict):
        return False
    if str(record.get("connection_mode") or "").lower() == "broker":
        broker = str(record.get("broker_url") or ebay_broker_url()).strip().rstrip("/")
        connection_token = unprotect_secret(record.get("connection_token"))
        request = urllib.request.Request(
            broker + "/connection/disconnect",
            data=json.dumps({"connection_token": connection_token}).encode("utf-8"),
            headers={"Content-Type": "application/json", "Accept": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                response.read()
        except urllib.error.HTTPError as error:
            body = error.read().decode("utf-8", errors="replace")
            raise EbayOAuthError(body or str(error)) from error
        except urllib.error.URLError as error:
            raise EbayOAuthError(str(error)) from error
    accounts.pop(key, None)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)
    return True


def update_env_values(env_path: Path, values: dict[str, object]) -> None:
    existing = env_path.read_text(encoding="utf-8").splitlines() if env_path.exists() else []
    remaining = {key: str(value or "") for key, value in values.items() if str(value or "").strip()}
    updated: list[str] = []
    for line in existing:
        stripped = line.strip()
        key = stripped.split("=", 1)[0] if "=" in stripped and not stripped.startswith("#") else ""
        if key in remaining:
            updated.append(f"{key}={remaining.pop(key)}")
        else:
            updated.append(line)
    for key, value in remaining.items():
        updated.append(f"{key}={value}")
    env_path.write_text("\n".join(updated).rstrip() + "\n", encoding="utf-8")
