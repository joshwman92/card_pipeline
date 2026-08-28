from __future__ import annotations

import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace

from comp_engine.cy_runtime import CourtyardRuntimeManager


class FakeProcess:
    def poll(self):
        return None


class CourtyardRuntimeTests(unittest.TestCase):
    def _sdk(self, root: Path) -> Path:
        sdk = root / "Android" / "Sdk"
        (sdk / "platform-tools").mkdir(parents=True)
        (sdk / "emulator").mkdir(parents=True)
        (sdk / "platform-tools" / "adb.exe").touch()
        (sdk / "emulator" / "emulator.exe").touch()
        return sdk

    def test_existing_emulator_and_appium_are_reused(self) -> None:
        with TemporaryDirectory() as folder:
            sdk = self._sdk(Path(folder))
            starts = []

            def run(args, **_kwargs):
                if args[-1] == "devices":
                    return SimpleNamespace(returncode=0, stdout="List of devices attached\nemulator-5554\tdevice\n")
                if "getprop" in args:
                    return SimpleNamespace(returncode=0, stdout="1\n")
                raise AssertionError(args)

            manager = CourtyardRuntimeManager(
                env={"LUCAS_CY_ANDROID_SDK": str(sdk)},
                platform="win32",
                run_command=run,
                start_process=lambda *args, **kwargs: starts.append((args, kwargs)),
                http_ready=lambda _url: True,
            )

            info = manager.ensure_ready("http://127.0.0.1:4723")

            self.assertEqual(info.device_serial, "emulator-5554")
            self.assertFalse(info.emulator_started)
            self.assertFalse(info.appium_started)
            self.assertEqual(starts, [])

    def test_missing_emulator_and_appium_are_started_and_waited_for(self) -> None:
        with TemporaryDirectory() as folder:
            root = Path(folder)
            sdk = self._sdk(root)
            appium = root / "appium.cmd"
            appium.touch()
            starts = []
            device_checks = 0
            appium_checks = 0

            def run(args, **_kwargs):
                nonlocal device_checks
                if args[-1] == "devices":
                    device_checks += 1
                    body = "" if device_checks == 1 else "emulator-5554\tdevice\n"
                    return SimpleNamespace(returncode=0, stdout=f"List of devices attached\n{body}")
                if args[-1] == "-list-avds":
                    return SimpleNamespace(returncode=0, stdout="LUCAS_Courtyard\n")
                if "getprop" in args:
                    return SimpleNamespace(returncode=0, stdout="1\n")
                if "keyevent" in args:
                    return SimpleNamespace(returncode=0, stdout="")
                raise AssertionError(args)

            def start(args, **kwargs):
                starts.append((args, kwargs))
                return FakeProcess()

            def http_ready(_url):
                nonlocal appium_checks
                appium_checks += 1
                return appium_checks >= 2

            manager = CourtyardRuntimeManager(
                env={
                    "LUCAS_CY_ANDROID_SDK": str(sdk),
                    "LUCAS_CY_ANDROID_AVD": "LUCAS_Courtyard",
                    "LUCAS_CY_APPIUM_COMMAND": str(appium),
                    "COMSPEC": "cmd.exe",
                },
                platform="win32",
                run_command=run,
                start_process=start,
                http_ready=http_ready,
                sleep=lambda _seconds: None,
            )

            info = manager.ensure_ready("http://127.0.0.1:4723")

            self.assertEqual(info.device_serial, "emulator-5554")
            self.assertTrue(info.emulator_started)
            self.assertTrue(info.appium_started)
            self.assertEqual(len(starts), 2)
            self.assertEqual(starts[0][0][-2:], ["-avd", "LUCAS_Courtyard"])
            self.assertEqual(starts[1][0][-1], str(appium))

    def test_autostart_can_be_disabled(self) -> None:
        with TemporaryDirectory() as folder:
            sdk = self._sdk(Path(folder))

            manager = CourtyardRuntimeManager(
                env={"LUCAS_CY_ANDROID_SDK": str(sdk), "LUCAS_CY_AUTO_START": "0"},
                platform="win32",
                run_command=lambda _args, **_kwargs: SimpleNamespace(
                    returncode=0, stdout="List of devices attached\n"
                ),
                http_ready=lambda _url: False,
            )

            with self.assertRaisesRegex(RuntimeError, "AUTO_START is disabled"):
                manager.ensure_ready("http://127.0.0.1:4723")

    def test_unresponsive_existing_emulator_is_reported_when_autostart_is_disabled(self) -> None:
        with TemporaryDirectory() as folder:
            sdk = self._sdk(Path(folder))
            starts = []

            def run(args, **_kwargs):
                if args[-1] == "devices":
                    return SimpleNamespace(returncode=0, stdout="List of devices attached\nemulator-5554\tdevice\n")
                if "getprop" in args:
                    raise subprocess.TimeoutExpired(args, 10)
                raise AssertionError(args)

            manager = CourtyardRuntimeManager(
                env={"LUCAS_CY_ANDROID_SDK": str(sdk), "LUCAS_CY_AUTO_START": "0"},
                platform="win32",
                run_command=run,
                start_process=lambda *args, **kwargs: starts.append((args, kwargs)),
                http_ready=lambda _url: True,
            )

            with self.assertRaisesRegex(RuntimeError, "not responding to ADB"):
                manager.ensure_ready("http://127.0.0.1:4723")
            self.assertEqual(starts, [])

    def test_unresponsive_existing_emulator_is_restarted_automatically(self) -> None:
        with TemporaryDirectory() as folder:
            sdk = self._sdk(Path(folder))
            starts = []
            stopped = False
            launched = False

            def run(args, **_kwargs):
                nonlocal stopped
                if args[-1] == "devices":
                    connected = not stopped or launched
                    body = "emulator-5554\tdevice\n" if connected else ""
                    return SimpleNamespace(returncode=0, stdout=f"List of devices attached\n{body}")
                if args[-1] in {"kill-server", "start-server"}:
                    return SimpleNamespace(returncode=0, stdout="")
                if "getprop" in args:
                    if launched:
                        return SimpleNamespace(returncode=0, stdout="1\n")
                    raise subprocess.TimeoutExpired(args, 10)
                if args[-2:] == ["emu", "kill"]:
                    stopped = True
                    return SimpleNamespace(returncode=0, stdout="OK\n")
                if args[-1] == "-list-avds":
                    return SimpleNamespace(returncode=0, stdout="LUCAS_Courtyard\n")
                if "keyevent" in args:
                    return SimpleNamespace(returncode=0, stdout="")
                raise AssertionError(args)

            def start(args, **kwargs):
                nonlocal launched
                launched = True
                starts.append((args, kwargs))
                return FakeProcess()

            manager = CourtyardRuntimeManager(
                env={"LUCAS_CY_ANDROID_SDK": str(sdk)},
                platform="win32",
                run_command=run,
                start_process=start,
                http_ready=lambda _url: True,
                sleep=lambda _seconds: None,
            )

            info = manager.ensure_ready("http://127.0.0.1:4723")

            self.assertEqual(info.device_serial, "emulator-5554")
            self.assertTrue(info.emulator_started)
            self.assertTrue(stopped)
            self.assertEqual(len(starts), 1)

    def test_booting_existing_emulator_is_waited_for_without_starting_duplicate(self) -> None:
        with TemporaryDirectory() as folder:
            sdk = self._sdk(Path(folder))
            starts = []
            boot_checks = 0

            def run(args, **_kwargs):
                nonlocal boot_checks
                if args[-1] == "devices":
                    return SimpleNamespace(returncode=0, stdout="List of devices attached\nemulator-5554\tdevice\n")
                if "getprop" in args:
                    boot_checks += 1
                    return SimpleNamespace(returncode=0, stdout="1\n" if boot_checks >= 2 else "\n")
                raise AssertionError(args)

            manager = CourtyardRuntimeManager(
                env={"LUCAS_CY_ANDROID_SDK": str(sdk)},
                platform="win32",
                run_command=run,
                start_process=lambda *args, **kwargs: starts.append((args, kwargs)),
                http_ready=lambda _url: True,
                sleep=lambda _seconds: None,
            )

            info = manager.ensure_ready("http://127.0.0.1:4723")

            self.assertEqual(info.device_serial, "emulator-5554")
            self.assertFalse(info.emulator_started)
            self.assertEqual(starts, [])


if __name__ == "__main__":
    unittest.main()
