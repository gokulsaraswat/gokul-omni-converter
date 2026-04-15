from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any


def state_backup_dir(state_path: Path | str) -> Path:
    state_file = Path(state_path).expanduser()
    return state_file.parent / "backups"


def list_state_backups(state_path: Path | str, *, limit: int = 50) -> list[Path]:
    root = state_backup_dir(state_path)
    if not root.exists():
        return []
    backups = [path for path in root.glob("*.json") if path.is_file()]
    backups.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    return backups[: max(1, int(limit))]


def latest_state_backup(state_path: Path | str) -> Path | None:
    backups = list_state_backups(state_path, limit=1)
    return backups[0] if backups else None


def backup_state_file(state_path: Path | str, *, keep: int = 12) -> Path | None:
    source = Path(state_path).expanduser()
    if not source.exists() or not source.is_file():
        return None
    if source.stat().st_size <= 0:
        return None

    keep = max(3, int(keep))
    backup_root = state_backup_dir(source)
    backup_root.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    destination = backup_root / f"{source.stem}_{timestamp}.json"
    shutil.copy2(source, destination)

    backups = list_state_backups(source, limit=500)
    for stale in backups[keep:]:
        try:
            stale.unlink()
        except OSError:
            pass
    return destination


def restore_state_backup(backup_path: Path | str, state_path: Path | str) -> Path:
    backup_file = Path(backup_path).expanduser()
    if not backup_file.exists() or not backup_file.is_file():
        raise FileNotFoundError(f"Backup file not found: {backup_file}")
    destination = Path(state_path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(backup_file, destination)
    return destination


def load_latest_backup_json(state_path: Path | str) -> dict[str, Any] | None:
    for backup in list_state_backups(state_path, limit=12):
        try:
            data = json.loads(backup.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data, dict):
            return data
    return None
