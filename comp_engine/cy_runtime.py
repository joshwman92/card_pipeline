from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Mapping
from urllib.parse import urlparse


DEFAULT_AVD = "LUCAS_Courtyard"
PROJECT_ROOT = Path(__file__).resolve().parent.parent
APPIUM_LOG = PROJECT_ROOT / "work" / "cardladder-bridge" / "appium-autostart.log"


@dataclass(frozen=True)
class CourtyardRuntimeInfo:
    device_serial: str
    emulator_started: bool = False
    appium_started: bool = False


class CourtyardRuntimeManager:
    def __init__(
        self,
        env: Mapping[str, str] | None = None,
        platform: str | None = None,
        run_command: Callable | None = None,
        start_process: Callable | None = None,
        http_ready: Callable[[str], bool] | None = None,
        monotonic: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
        status: Callable[[str], None] | None = None,
    ) -> None:
        self.env = env if env is not None else os.environ
        self.platform = platform or os.sys.platform
        self.run_command = run_command or subprocess.run
        self.start_process = start_process or subprocess.Popen
        self.http_ready = http_ready or self._default_http_ready
        self.monotonic = monotonic
        self.sleep = sleep
        self.status = status or (lambda _message: None)
        self.started_processes: list[object] = []

    def ensure_ready(self, server_url: str, requested_udid: str = "") -> CourtyardRuntimeInfo:
        if self.platform != "win32":
            return CourtyardRuntimeInfo(device_serial=requested_udid)

        auto_start = self._env_bool("LUCAS_CY_AUTO_START", True)
        sdk = self._android_sdk()
        adb = sdk / "platform-tools" / "adb.exe"
        if not adb.exists():
            raise RuntimeError(
                f"Android SDK platform tools were not found at {adb}. "
                "Install Android Studio or set LUCAS_CY_ANDROID_SDK."
            )

        device_serial, emulator_started = self._ensure_emulator(adb, sdk, requested_udid, auto_start)
        appium_started = self._ensure_appium(server_url, auto_start)
        return CourtyardRuntimeInfo(
            device_serial=device_serial,
            emulator_started=emulator_started,
            appium_started=appium_started,
        )

    def _env_bool(self, name: str, default: bool) -> bool:
        value = str(self.env.get(name, "")).strip().lower()
        if not value:
            return default
        return value in {"1", "true", "yes", "on"}

    def _android_sdk(self) -> Path:
        candidates = [
            str(self.env.get("LUCAS_CY_ANDROID_SDK", "")).strip(),
            str(self.env.get("ANDROID_HOME", "")).strip(),
            str(self.env.get("ANDROID_SDK_ROOT", "")).strip(),
        ]
        local_app_data = str(self.env.get("LOCALAPPDATA", "")).strip()
        if local_app_data:
            candidates.append(str(Path(local_app_data) / "Android" / "Sdk"))
        for candidate in candidates:
            if candidate and Path(candidate).expanduser().is_dir():
                return Path(candidate).expanduser()
        raise RuntimeError(
            "Android SDK was not found. Install Android Studio or set "
            "LUCAS_CY_ANDROID_SDK to the SDK folder."
        )

    @staticmethod
    def _windows_creation_flags() -> int:
        return int(getattr(subprocess, "CREATE_NO_WINDOW", 0)) | int(
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        )

    def _run(self, args: list[str], timeout: float = 10.0):
        return self.run_command(
            args,
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=self._windows_creation_flags(),
        )

    def _device_states(self, adb: Path) -> dict[str, str]:
        try:
            result = self._run([str(adb), "devices"], timeout=10)
        except (OSError, subprocess.SubprocessError):
            return {}
        states: dict[str, str] = {}
        for line in str(getattr(result, "stdout", "") or "").splitlines()[1:]:
            parts = line.strip().split()
            if len(parts) >= 2:
                states[parts[0]] = parts[1]
        return states

    def _boot_state(self, adb: Path, serial: str) -> str:
        try:
            result = self._run(
                [str(adb), "-s", serial, "shell", "getprop", "sys.boot_completed"],
                timeout=10,
            )
        except subprocess.TimeoutExpired:
            return "unresponsive"
        except (OSError, subprocess.SubprocessError):
            return "booting"
        if getattr(result, "returncode", 1) != 0:
            return "booting"
        return "ready" if str(getattr(result, "stdout", "") or "").strip() == "1" else "booting"

    def _select_device(self, states: dict[str, str], requested_udid: str) -> str:
        if requested_udid:
            return requested_udid if requested_udid in states else ""
        ready = [serial for serial, state in states.items() if state == "device" and serial.startswith("emulator-")]
        return ready[0] if ready else ""

    def _ensure_emulator(
        self,
        adb: Path,
        sdk: Path,
        requested_udid: str,
        auto_start: bool,
    ) -> tuple[str, bool]:
        timeout = max(30.0, float(self.env.get("LUCAS_CY_EMULATOR_START_TIMEOUT", "180") or 180))
        states = self._device_states(adb)
        serial = self._select_device(states, requested_udid)
        if serial:
            boot_state = self._boot_state(adb, serial)
            if boot_state == "ready":
                self.status(f"CourtYard emulator ready: {serial}")
                return serial, False
            if boot_state == "unresponsive":
                if not auto_start:
                    raise RuntimeError(
                        f"Android emulator {serial} is connected but not responding to ADB and "
                        "LUCAS_CY_AUTO_START is disabled."
                    )
                recovered = self._recover_unresponsive_emulator(adb, serial)
                if recovered:
                    self.status(f"CourtYard emulator recovered: {serial}")
                    return serial, False
                serial = ""
            if serial:
                self.status(f"Waiting for Android to finish booting on {serial}...")
                deadline = self.monotonic() + timeout
                while self.monotonic() < deadline:
                    boot_state = self._boot_state(adb, serial)
                    if boot_state == "ready":
                        self.status(f"CourtYard emulator ready: {serial}")
                        return serial, False
                    if boot_state == "unresponsive":
                        raise RuntimeError(f"Android emulator {serial} stopped responding while it was starting.")
                    self.sleep(1.0)
                raise RuntimeError(f"Android emulator {serial} did not finish booting within {timeout:.0f} seconds.")

        if not auto_start:
            raise RuntimeError("CourtYard emulator is not running and LUCAS_CY_AUTO_START is disabled.")

        emulator = sdk / "emulator" / "emulator.exe"
        if not emulator.exists():
            raise RuntimeError(f"Android emulator executable was not found at {emulator}.")
        avd = str(self.env.get("LUCAS_CY_ANDROID_AVD", DEFAULT_AVD)).strip() or DEFAULT_AVD
        try:
            listed = self._run([str(emulator), "-list-avds"], timeout=15)
            avds = {line.strip() for line in str(getattr(listed, "stdout", "") or "").splitlines() if line.strip()}
        except (OSError, subprocess.SubprocessError) as error:
            raise RuntimeError(f"Could not list Android emulator profiles: {error}") from error
        if avd not in avds:
            available = ", ".join(sorted(avds)) or "none"
            raise RuntimeError(
                f"Android emulator profile {avd!r} was not found. Available profiles: {available}. "
                "Set LUCAS_CY_ANDROID_AVD to the correct profile name."
            )

        self.status(f"Starting CourtYard Android emulator {avd}...")
        try:
            process = self.start_process(
                [str(emulator), "-avd", avd],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=self._windows_creation_flags(),
            )
        except OSError as error:
            raise RuntimeError(f"Could not launch Android emulator {avd}: {error}") from error
        self.started_processes.append(process)

        deadline = self.monotonic() + timeout
        while self.monotonic() < deadline:
            if callable(getattr(process, "poll", None)) and process.poll() is not None:
                raise RuntimeError(f"Android emulator {avd} exited before Android finished booting.")
            states = self._device_states(adb)
            serial = self._select_device(states, requested_udid)
            if serial:
                boot_state = self._boot_state(adb, serial)
                if boot_state == "ready":
                    try:
                        self._run([str(adb), "-s", serial, "shell", "input", "keyevent", "82"], timeout=5)
                    except Exception:
                        pass
                    self.status(f"CourtYard emulator ready: {serial}")
                    return serial, True
                if boot_state == "unresponsive":
                    raise RuntimeError(f"Android emulator {serial} stopped responding while it was starting.")
            self.sleep(1.0)
        raise RuntimeError(f"Android emulator {avd} did not finish booting within {timeout:.0f} seconds.")

    def _recover_unresponsive_emulator(self, adb: Path, serial: str) -> bool:
        self.status(f"Recovering unresponsive CourtYard emulator {serial}...")
        try:
            self._run([str(adb), "kill-server"], timeout=10)
            self._run([str(adb), "start-server"], timeout=15)
        except (OSError, subprocess.SubprocessError):
            pass

        states = self._device_states(adb)
        if serial in states:
            boot_state = self._boot_state(adb, serial)
            if boot_state == "ready":
                return True
            if boot_state == "booting":
                deadline = self.monotonic() + 30.0
                while self.monotonic() < deadline:
                    boot_state = self._boot_state(adb, serial)
                    if boot_state == "ready":
                        return True
                    if boot_state == "unresponsive":
                        break
                    self.sleep(1.0)

        self.status(f"Restarting frozen CourtYard emulator {serial}...")
        try:
            result = self._run([str(adb), "-s", serial, "emu", "kill"], timeout=15)
        except (OSError, subprocess.SubprocessError) as error:
            raise RuntimeError(
                f"Android emulator {serial} is frozen and LUCAS could not stop it automatically: {error}. "
                "Use Cold Boot Now in Android Studio."
            ) from error
        if getattr(result, "returncode", 1) != 0:
            detail = str(getattr(result, "stderr", "") or getattr(result, "stdout", "") or "unknown error").strip()
            raise RuntimeError(
                f"Android emulator {serial} is frozen and rejected the automatic restart: {detail}. "
                "Use Cold Boot Now in Android Studio."
            )

        deadline = self.monotonic() + 30.0
        while self.monotonic() < deadline:
            if serial not in self._device_states(adb):
                return False
            self.sleep(0.5)
        raise RuntimeError(
            f"Android emulator {serial} did not stop after LUCAS requested a restart. "
            "Use Cold Boot Now in Android Studio."
        )

    @staticmethod
    def _default_http_ready(server_url: str) -> bool:
        status_url = server_url.rstrip("/") + "/status"
        try:
            with urllib.request.urlopen(status_url, timeout=2) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception:
            return False
        value = payload.get("value", {}) if isinstance(payload, dict) else {}
        return bool(value.get("ready"))

    def _appium_command(self) -> str:
        configured = str(self.env.get("LUCAS_CY_APPIUM_COMMAND", "")).strip().strip('"')
        if configured:
            return configured
        discovered = shutil.which("appium.cmd") or shutil.which("appium")
        if discovered:
            return discovered
        app_data = str(self.env.get("APPDATA", "")).strip()
        if app_data:
            candidate = Path(app_data) / "npm" / "appium.cmd"
            if candidate.exists():
                return str(candidate)
        return ""

    def _ensure_appium(self, server_url: str, auto_start: bool) -> bool:
        if self.http_ready(server_url):
            self.status(f"Appium ready: {server_url}")
            return False
        parsed = urlparse(server_url)
        if parsed.hostname not in {"127.0.0.1", "localhost", "::1"}:
            raise RuntimeError(f"Remote Appium server is unavailable: {server_url}")
        if not auto_start:
            raise RuntimeError("Appium is not running and LUCAS_CY_AUTO_START is disabled.")
        command = self._appium_command()
        if not command or not Path(command).exists():
            raise RuntimeError(
                "Appium was not found. Install it with npm or set LUCAS_CY_APPIUM_COMMAND "
                "to appium.cmd."
            )

        APPIUM_LOG.parent.mkdir(parents=True, exist_ok=True)
        self.status(f"Starting Appium at {server_url}...")
        comspec = str(self.env.get("COMSPEC", "")).strip() or "cmd.exe"
        try:
            with APPIUM_LOG.open("a", encoding="utf-8") as log:
                process = self.start_process(
                    [comspec, "/d", "/s", "/c", command],
                    stdin=subprocess.DEVNULL,
                    stdout=log,
                    stderr=subprocess.STDOUT,
                    creationflags=self._windows_creation_flags(),
                )
        except OSError as error:
            raise RuntimeError(f"Could not launch Appium from {command}: {error}") from error
        self.started_processes.append(process)

        timeout = max(10.0, float(self.env.get("LUCAS_CY_APPIUM_START_TIMEOUT", "45") or 45))
        deadline = self.monotonic() + timeout
        while self.monotonic() < deadline:
            if self.http_ready(server_url):
                self.status(f"Appium ready: {server_url}")
                return True
            if callable(getattr(process, "poll", None)) and process.poll() is not None:
                raise RuntimeError(
                    f"Appium exited before becoming ready. Review {APPIUM_LOG} for details."
                )
            self.sleep(0.5)
        raise RuntimeError(
            f"Appium did not become ready at {server_url} within {timeout:.0f} seconds. "
            f"Review {APPIUM_LOG} for details."
        )


def ensure_courtyard_runtime(
    server_url: str,
    requested_udid: str = "",
    status: Callable[[str], None] | None = None,
) -> CourtyardRuntimeInfo:
    return CourtyardRuntimeManager(status=status).ensure_ready(server_url, requested_udid)
