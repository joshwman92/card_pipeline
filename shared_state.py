from __future__ import annotations

import json
import os
import socket
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


LOCK_TIMEOUT_SECONDS = 45
LOCK_POLL_SECONDS = 0.25


def local_identity(settings_path: Path) -> dict[str, str]:
    identity_path = settings_path.with_name("lucas_user_identity.json")
    identity = read_json(identity_path, {})
    if not isinstance(identity, dict):
        identity = {}
    changed = False
    if not str(identity.get("user_id") or "").strip():
        identity["user_id"] = uuid.uuid4().hex
        changed = True
    if not str(identity.get("machine") or "").strip():
        identity["machine"] = socket.gethostname()
        changed = True
    if not str(identity.get("display_name") or "").strip():
        identity["display_name"] = os.environ.get("USERNAME") or identity["machine"]
        changed = True
    if changed:
        atomic_write_json(identity_path, identity)
    return {key: str(identity.get(key) or "") for key in ("user_id", "machine", "display_name")}


def read_json(path: Path, default):
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default
    return raw


def atomic_write_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    temp_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    temp_path.replace(path)


@contextmanager
def shared_lock(root: Path, name: str, owner: dict[str, str] | None = None, timeout: float = LOCK_TIMEOUT_SECONDS) -> Iterator[Path]:
    lock_dir = root / ".locks"
    lock_dir.mkdir(parents=True, exist_ok=True)
    lock_path = lock_dir / f"{safe_lock_name(name)}.lock"
    owner = owner or {}
    token = uuid.uuid4().hex
    started = time.time()
    payload = {
        "token": token,
        "name": name,
        "pid": os.getpid(),
        "user": owner.get("display_name") or "",
        "machine": owner.get("machine") or "",
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, indent=2, sort_keys=True)
            break
        except FileExistsError:
            if is_abandoned_local_lock(lock_path) or is_stale_lock(lock_path, timeout):
                try:
                    lock_path.unlink()
                    continue
                except OSError:
                    pass
            if time.time() - started >= timeout:
                detail = read_json(lock_path, {})
                holder = " ".join(part for part in (str(detail.get("user") or ""), str(detail.get("machine") or "")) if part).strip()
                raise TimeoutError(f"Timed out waiting for shared lock '{name}'" + (f" held by {holder}" if holder else ""))
            time.sleep(LOCK_POLL_SECONDS)
    try:
        yield lock_path
    finally:
        try:
            current = read_json(lock_path, {})
            if current.get("token") == token:
                lock_path.unlink()
        except OSError:
            pass


def is_stale_lock(path: Path, timeout: float) -> bool:
    try:
        age = time.time() - path.stat().st_mtime
    except OSError:
        return False
    return age > max(timeout * 3, 120)


def is_abandoned_local_lock(path: Path) -> bool:
    detail = read_json(path, {})
    if not isinstance(detail, dict):
        return False
    machine = str(detail.get("machine") or "").strip().casefold()
    local_machine = socket.gethostname().strip().casefold()
    if not machine or machine != local_machine:
        return False
    try:
        pid = int(detail.get("pid") or 0)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    except OSError:
        return True
    return False


def safe_lock_name(name: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in name)
    return cleaned.strip("_") or "shared"
