from __future__ import annotations

import importlib.util
import os
import re
import time
from collections import Counter
from dataclasses import dataclass
from typing import Callable

try:
    from .cy_runtime import ensure_courtyard_runtime
except ImportError:
    from cy_runtime import ensure_courtyard_runtime


COURTYARD_PACKAGE = "io.courtyard.app"
SUPPORTED_GRADERS = {"PSA", "BGS", "CGC"}
NOT_BUYING_MESSAGE = "Courtyard marked this card NOT BUYING"
RETRYABLE_LOOKUP_MESSAGES = {
    "Courtyard search timed out",
    "Courtyard search timed out or returned no new result",
}
RESULT_XPATH = "//android.widget.Button[contains(@content-desc, 'CONFIDENCE')]"
TIMEOUT_XPATH = "//*[contains(@text, 'Search timed out') or contains(@content-desc, 'Search timed out')]"
RESULT_PATTERN = re.compile(
    r"^(?P<details>.*),\s*(?P<grader>PSA|BGS|CGC)\s+(?P<grade>\d+(?:\.\d+)?),\s*"
    r"(?:(?P<not_buying>NOT\s+BUYING),\s*)?"
    r"\$(?P<value>[\d,]+(?:\.\d+)?),\s*CONFIDENCE,\s*(?P<confidence>\d+)\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class CourtyardResult:
    description: str
    card_name: str
    card_set: str
    grader: str
    grade: str
    value: float
    confidence: int
    card_number: str
    not_buying: bool = False


def appium_client_available() -> bool:
    try:
        return importlib.util.find_spec("appium") is not None
    except (ImportError, ValueError):
        return False


def parse_result_description(value: object) -> CourtyardResult | None:
    description = re.sub(r"\s+", " ", str(value or "")).strip()
    match = RESULT_PATTERN.match(description)
    if not match:
        return None
    details = match.group("details").strip()
    detail_parts = [part.strip() for part in details.split(",")]
    card_name = detail_parts[0]
    card_set = ", ".join(detail_parts[1:]).strip()
    card_number = extract_card_number(details)
    try:
        price = float(match.group("value").replace(",", ""))
        confidence = int(match.group("confidence"))
    except ValueError:
        return None
    return CourtyardResult(
        description=description,
        card_name=card_name,
        card_set=card_set,
        grader=match.group("grader").upper(),
        grade=normalize_grade(match.group("grade")),
        value=price,
        confidence=confidence,
        card_number=card_number,
        not_buying=bool(match.group("not_buying")),
    )


def normalize_grade(value: object) -> str:
    text = str(value or "").strip()
    try:
        number = float(text)
    except ValueError:
        return text.upper()
    return str(int(number)) if number.is_integer() else f"{number:g}"


def normalize_card_number(value: object) -> str:
    token = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if token.isdigit():
        return str(int(token)) if token else ""
    return token


def extract_card_number(value: object) -> str:
    text = str(value or "")
    match = re.search(r"#\s*([A-Z0-9][A-Z0-9./-]*)", text, flags=re.I)
    return normalize_card_number(match.group(1)) if match else ""


def expected_profile_grade(profile_title: object, grader: str) -> str:
    matches = re.findall(
        rf"\b{re.escape(grader)}\s*(\d+(?:\.\d+)?)\b",
        str(profile_title or ""),
        flags=re.I,
    )
    return normalize_grade(matches[-1]) if matches else ""


def profile_contains_card_number(profile_title: object, card_number: str) -> bool:
    target = normalize_card_number(card_number)
    if not target:
        return False
    profile = re.sub(r"\s*/\s*", "/", str(profile_title or "").upper())
    tokens = re.findall(r"[A-Z0-9][A-Z0-9./-]*", profile)
    return target in {normalize_card_number(token) for token in tokens}


_NAME_STOP_WORDS = {
    "psa", "bgs", "cgc", "gem", "mint", "card", "cards", "rookie", "rc",
    "auto", "autograph", "refractor", "prism", "prizm", "parallel", "holo",
    "silver", "gold", "blue", "red", "green", "orange", "purple", "black",
    "white", "yellow", "pink", "wave", "laser", "shimmer", "choice", "optic",
}


def card_name_tokens(value: object) -> set[str]:
    text = re.sub(r"#\s*[A-Z0-9][A-Z0-9./-]*", " ", str(value or ""), flags=re.I).lower()
    tokens = re.findall(r"[^\W_]+", text, flags=re.UNICODE)
    return {token for token in tokens if len(token) >= 3 and token not in _NAME_STOP_WORDS and not token.isdigit()}


def _name_token_signature(value: str) -> str:
    token = str(value or "").lower()
    if len(token) < 4:
        return token
    # Marketplace/profile names sometimes omit vowels (for example Zkrm. for
    # Zekrom). Preserve the leading character and consonants so those compact
    # spellings compare without introducing broad edit-distance matches.
    return token[0] + re.sub(r"[aeiouy]", "", token[1:])


def card_name_matches(profile_title: object, result_name: object) -> bool:
    result_tokens = card_name_tokens(result_name)
    if not result_tokens:
        return False
    profile_tokens = card_name_tokens(profile_title)
    profile_signatures = {_name_token_signature(token) for token in profile_tokens}
    overlap = {
        token
        for token in result_tokens
        if token in profile_tokens or _name_token_signature(token) in profile_signatures
    }
    ratio = len(overlap) / len(result_tokens)
    # Two shared distinctive words plus an exact number/set identity is enough
    # for expanded names such as "Shadow Rider Calyrex VMAX" versus Card
    # Ladder's abbreviated "SR Calyrex VMAX".
    return ratio >= 0.6 or (len(overlap) >= 2 and ratio >= 0.5)


def card_set_matches(profile_title: object, result_set: object) -> bool:
    result_tokens = card_name_tokens(result_set)
    if not result_tokens:
        return False
    profile_tokens = card_name_tokens(profile_title)
    overlap = result_tokens & profile_tokens
    return len(overlap) / len(result_tokens) >= 0.6


def validate_result_profile(
    result: CourtyardResult,
    profile_title: object,
    expected_grader: str,
    *,
    allow_missing_card_number: bool = False,
) -> tuple[bool, str]:
    expected_grader = str(expected_grader or "").strip().upper()
    if result.grader != expected_grader:
        return False, f"grader mismatch: expected {expected_grader}, got {result.grader}"
    expected_grade = expected_profile_grade(profile_title, expected_grader)
    if not expected_grade:
        return False, "Card Ladder profile did not contain a grade"
    if result.grade != expected_grade:
        return False, f"grade mismatch: expected {expected_grade}, got {result.grade}"
    if not card_name_matches(profile_title, result.card_name):
        return False, f"card name {result.card_name!r} did not match the Card Ladder profile"
    if result.card_number and not profile_contains_card_number(profile_title, result.card_number):
        return False, f"card number #{result.card_number} did not match the Card Ladder profile"
    if not result.card_number:
        if not allow_missing_card_number:
            return False, "Courtyard result did not contain a card number"
        if not card_set_matches(profile_title, result.card_set):
            return False, f"card set {result.card_set!r} did not match the Card Ladder profile"
    return True, ""


def changed_result_descriptions(before: list[str], after: list[str]) -> list[str]:
    remaining = Counter(after) - Counter(before)
    changed: list[str] = []
    for description in after:
        if remaining[description] > 0:
            changed.append(description)
            remaining[description] -= 1
    if changed or before == after:
        return changed
    return [
        description
        for index, description in enumerate(after)
        if index >= len(before) or description != before[index]
    ]


class CourtyardAndroidAdapter:
    def __init__(
        self,
        driver_factory: Callable | None = None,
        monotonic: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
        runtime_status: Callable[[str], None] | None = None,
    ) -> None:
        self.driver_factory = driver_factory
        self.monotonic = monotonic
        self.sleep = sleep
        self.runtime_status = runtime_status
        self.driver = None
        self.unavailable_error = ""
        self.package = os.environ.get("LUCAS_CY_APP_PACKAGE", COURTYARD_PACKAGE).strip() or COURTYARD_PACKAGE
        self.server_url = os.environ.get("LUCAS_CY_APPIUM_URL", "http://127.0.0.1:4723").strip()
        self.activity = os.environ.get("LUCAS_CY_APP_ACTIVITY", ".MainActivity").strip() or ".MainActivity"
        self.udid = os.environ.get("LUCAS_CY_ANDROID_UDID", "").strip()
        self.timeout = max(5.0, float(os.environ.get("LUCAS_CY_LOOKUP_TIMEOUT", "30") or 30))

    def _create_driver(self):
        if self.driver_factory is not None:
            return self.driver_factory()
        runtime = ensure_courtyard_runtime(self.server_url, self.udid, status=self.runtime_status)
        if not self.udid:
            self.udid = runtime.device_serial
        try:
            from appium import webdriver
            from appium.options.android import UiAutomator2Options
        except ImportError as error:
            raise RuntimeError("Appium-Python-Client is not installed. Run install_dependencies.bat.") from error
        capabilities = {
            "platformName": "Android",
            "automationName": "UiAutomator2",
            "appPackage": self.package,
            "appActivity": self.activity,
            "noReset": True,
            "newCommandTimeout": 600,
        }
        if self.udid:
            capabilities["udid"] = self.udid
        options = UiAutomator2Options().load_capabilities(capabilities)
        try:
            return webdriver.Remote(self.server_url, options=options)
        except Exception as error:
            details = str(error)
            if "adbexec" in details.lower() and "timed out" in details.lower():
                raise RuntimeError(
                    "The Android emulator is connected but is not responding to ADB commands. "
                    "Close and restart the emulator, wait for Android to finish booting, then retry. "
                    f"Details: {details}"
                ) from error
            raise RuntimeError(
                f"Could not start the Courtyard Appium session at {self.server_url}. "
                f"Start the emulator and Appium server, then try again. Details: {details}"
            ) from error

    def _ensure_driver(self):
        if self.unavailable_error:
            raise RuntimeError(self.unavailable_error)
        if self.driver is None:
            try:
                self.driver = self._create_driver()
            except Exception as error:
                self.unavailable_error = str(error)
                raise
        return self.driver

    def _discard_driver(self) -> None:
        driver, self.driver = self.driver, None
        if driver is None:
            return
        try:
            driver.quit()
        except Exception:
            pass

    @staticmethod
    def _appium_by():
        try:
            from appium.webdriver.common.appiumby import AppiumBy
        except ImportError as error:
            raise RuntimeError("Appium-Python-Client is not installed. Run install_dependencies.bat.") from error
        return AppiumBy

    def _find_accessibility(self, label: str):
        driver = self._ensure_driver()
        by = self._appium_by().ACCESSIBILITY_ID
        find_many = getattr(driver, "find_elements", None)
        if callable(find_many):
            elements = find_many(by, label)
            if elements:
                return elements[0]
            raise LookupError(f"Courtyard did not show {label!r} yet.")
        # Retain compatibility with lightweight test/fake drivers that only
        # implement find_element. Real Appium sessions use the non-throwing
        # find_elements path above so expected polling misses do not flood the
        # Appium server log with W3C NoSuchElementError entries.
        return driver.find_element(by, label)

    def _wait_for_accessibility(self, label: str, timeout: float = 8.0):
        deadline = self.monotonic() + timeout
        last_error: Exception | None = None
        while self.monotonic() < deadline:
            try:
                return self._find_accessibility(label)
            except Exception as error:
                last_error = error
                self.sleep(0.25)
        raise RuntimeError(f"Courtyard did not show {label!r}.") from last_error

    def _wait_for_enabled_accessibility(self, label: str, timeout: float = 5.0):
        deadline = self.monotonic() + timeout
        element = None
        while self.monotonic() < deadline:
            element = self._wait_for_accessibility(label, timeout=min(1.0, max(0.25, deadline - self.monotonic())))
            enabled_check = getattr(element, "is_enabled", None)
            if not callable(enabled_check) or enabled_check():
                return element
            self.sleep(0.2)
        raise RuntimeError(f"Courtyard showed {label!r}, but it did not become enabled.")

    def _open_manual_entry(self) -> None:
        driver = self._ensure_driver()
        try:
            driver.activate_app(self.package)
        except Exception:
            pass
        # The Manual Entry control is part of the scanner page and disappears
        # while its drawer is open. Recognize the open drawer by its input
        # instead of trying to navigate away from an already-correct screen.
        try:
            self._find_accessibility("Certificate number")
            return
        except Exception:
            pass
        try:
            self._find_accessibility("Manual Entry").click()
        except Exception:
            self._wait_for_accessibility("User menu").click()
            self._wait_for_accessibility("Scanner").click()
            self._wait_for_accessibility("Manual Entry").click()
        self._wait_for_accessibility("Certificate number")

    @staticmethod
    def _content_description(element) -> str:
        for attribute in ("content-desc", "contentDescription"):
            try:
                value = element.get_attribute(attribute)
            except Exception:
                value = ""
            if value:
                return re.sub(r"\s+", " ", str(value)).strip()
        return ""

    def _result_descriptions(self) -> list[str]:
        driver = self._ensure_driver()
        elements = driver.find_elements(self._appium_by().XPATH, RESULT_XPATH)
        return [description for element in elements if (description := self._content_description(element))]

    def _search_timed_out(self) -> bool:
        try:
            return bool(self._ensure_driver().find_elements(self._appium_by().XPATH, TIMEOUT_XPATH))
        except Exception:
            return False

    def lookup(self, cert_number: str, slab_type: str, profile_title: str) -> tuple[float | None, object | None, str]:
        slab_type = str(slab_type or "").strip().upper()
        if slab_type not in SUPPORTED_GRADERS:
            return None, None, f"{slab_type or 'unknown'} is not supported by the Courtyard Android Scanner"
        if not str(profile_title or "").strip():
            return None, None, "Card Ladder profile was blank; Courtyard result could not be validated"
        for attempt in range(2):
            try:
                result = self._lookup_once(cert_number, slab_type, profile_title)
            except Exception:
                if attempt or self.driver is None:
                    raise
                # Retain the existing one-time recovery for a broken Appium
                # session. The replacement session is then the retry.
                self._discard_driver()
                continue
            if attempt or result[2] not in RETRYABLE_LOOKUP_MESSAGES:
                return result
            if callable(self.runtime_status):
                self.runtime_status(
                    f"Retrying CourtYard cert {str(cert_number or '').strip()} once after: {result[2]}"
                )
        return None, None, "Courtyard lookup failed after retry"

    def _lookup_once(self, cert_number: str, slab_type: str, profile_title: str) -> tuple[float | None, object | None, str]:
        self._open_manual_entry()
        self._wait_for_accessibility(slab_type).click()
        cert_field = self._wait_for_accessibility("Certificate number")
        cert_field.clear()
        cert_field.send_keys(str(cert_number or "").strip())
        try:
            before = self._result_descriptions()
        except Exception:
            before = []
        self._wait_for_enabled_accessibility("SEARCH").click()

        deadline = self.monotonic() + self.timeout
        mismatch = ""
        while self.monotonic() < deadline:
            try:
                after = self._result_descriptions()
            except Exception:
                after = []
            # CourtYard sometimes updates/reuses an existing Recent Scans row
            # instead of adding a new one. The profile validation below is the
            # stale-result guard, so inspect every visible result after submit.
            changed = changed_result_descriptions(before, after)
            reasons_by_description: dict[str, str] = {}
            for description in after:
                result = parse_result_description(description)
                if result is None:
                    continue
                valid, reason = validate_result_profile(
                    result,
                    profile_title,
                    slab_type,
                    # Some CourtYard rows omit the card number from their
                    # accessibility description. Only allow the stronger
                    # name + set + grader + grade fallback for a row proven to
                    # have appeared or changed after the current submission.
                    allow_missing_card_number=description in changed,
                )
                if valid:
                    if result.not_buying:
                        return None, None, NOT_BUYING_MESSAGE
                    return result.value, result.confidence, ""
                mismatch = reason
                reasons_by_description[description] = reason
            # A result that appeared or changed after this submit belongs to
            # the current search. If it is parseable but fails identity
            # validation, waiting the entire lookup timeout cannot make that
            # completed result valid and only makes the batch look hung.
            changed_reasons = [reasons_by_description[item] for item in changed if item in reasons_by_description]
            if changed_reasons:
                return None, None, f"Courtyard result rejected: {changed_reasons[-1]}"
            if self._search_timed_out():
                return None, None, "Courtyard search timed out"
            self.sleep(0.5)
        if mismatch:
            return None, None, f"Courtyard result rejected: {mismatch}"
        return None, None, "Courtyard search timed out or returned no new result"

    def close_app(self) -> None:
        self._discard_driver()
