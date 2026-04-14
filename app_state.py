from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from automation_core import add_fingerprints, normalize_preset_record, normalize_watch_config

APP_NAME = "Gokul Omni Convert Lite"
APP_STATE_DIR = Path.home() / ".gokul_omni_convert_lite"
APP_STATE_PATH = APP_STATE_DIR / "app_state.json"
MAX_RECENT_JOBS = 30
MAX_AUTOMATION_EVENTS = 120


DEFAULT_STATE: dict[str, Any] = {
    "theme": "dark",
    "output_dir": str(Path.cwd() / "converted_output"),
    "recursive_scan": True,
    "conversion_engine": "pure_python",
    "soffice_path": "",
    "tesseract_path": "",
    "ocr_language": "eng",
    "ocr_dpi": 220,
    "ocr_psm": 6,
    "ocr_output_dir": "",
    "smtp_settings": {},
    "organizer_last_pdf": "",
    "recent_jobs": [],
    "presets": [],
    "watch_config": {},
    "watch_seen_fingerprints": [],
    "automation_events": [],
}


class AppStateStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or APP_STATE_PATH
        self.state: dict[str, Any] = dict(DEFAULT_STATE)
        self.load()

    def load(self) -> dict[str, Any]:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            self.state = dict(DEFAULT_STATE)
            self.save()
            return self.state

        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self.state = dict(DEFAULT_STATE)
            self.save()
            return self.state

        merged = dict(DEFAULT_STATE)
        merged.update(data)
        if not isinstance(merged.get("recent_jobs"), list):
            merged["recent_jobs"] = []
        if not isinstance(merged.get("presets"), list):
            merged["presets"] = []
        if not isinstance(merged.get("watch_seen_fingerprints"), list):
            merged["watch_seen_fingerprints"] = []
        if not isinstance(merged.get("automation_events"), list):
            merged["automation_events"] = []
        if not isinstance(merged.get("watch_config"), dict):
            merged["watch_config"] = {}
        if not isinstance(merged.get("smtp_settings"), dict):
            merged["smtp_settings"] = {}
        self.state = merged
        return self.state

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.state, indent=2), encoding="utf-8")

    def get(self, key: str, default: Any = None) -> Any:
        return self.state.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.state[key] = value
        self.save()

    def update(self, **kwargs: Any) -> None:
        self.state.update(kwargs)
        self.save()

    def recent_jobs(self) -> list[dict[str, Any]]:
        jobs = self.state.get("recent_jobs", [])
        return [item for item in jobs if isinstance(item, dict)]

    def add_recent_job(self, job: dict[str, Any]) -> None:
        record = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            **job,
        }
        jobs = self.recent_jobs()
        jobs.insert(0, record)
        self.state["recent_jobs"] = jobs[:MAX_RECENT_JOBS]
        self.save()

    def clear_recent_jobs(self) -> None:
        self.state["recent_jobs"] = []
        self.save()

    def presets(self) -> list[dict[str, Any]]:
        raw = self.state.get("presets", [])
        normalized: list[dict[str, Any]] = []
        for item in raw:
            if isinstance(item, dict):
                normalized.append(normalize_preset_record(item).to_dict())
        return normalized

    def save_preset(self, preset: dict[str, Any]) -> None:
        normalized = normalize_preset_record(preset).to_dict()
        name = str(normalized.get("name", "")).strip().lower()
        presets = self.presets()
        updated: list[dict[str, Any]] = []
        replaced = False
        for item in presets:
            if str(item.get("name", "")).strip().lower() == name:
                updated.append(normalized)
                replaced = True
            else:
                updated.append(item)
        if not replaced:
            updated.insert(0, normalized)
        self.state["presets"] = updated[:50]
        self.save()

    def replace_presets(self, presets: Iterable[dict[str, Any]]) -> None:
        normalized: list[dict[str, Any]] = []
        seen: set[str] = set()
        for item in presets:
            record = normalize_preset_record(item).to_dict()
            name = str(record.get("name", "")).strip().lower()
            if not name or name in seen:
                continue
            seen.add(name)
            normalized.append(record)
        self.state["presets"] = normalized[:50]
        self.save()

    def delete_preset(self, name: str) -> None:
        target = str(name).strip().lower()
        if not target:
            return
        self.state["presets"] = [
            item for item in self.presets()
            if str(item.get("name", "")).strip().lower() != target
        ]
        self.save()

    def watch_config(self) -> dict[str, Any]:
        return normalize_watch_config(self.state.get("watch_config", {})).to_dict()

    def set_watch_config(self, config: dict[str, Any]) -> None:
        self.state["watch_config"] = normalize_watch_config(config).to_dict()
        self.save()

    def watch_seen_fingerprints(self) -> list[str]:
        items = self.state.get("watch_seen_fingerprints", [])
        return [str(item).strip() for item in items if str(item).strip()]

    def add_watch_seen(self, fingerprints: Iterable[str]) -> None:
        self.state["watch_seen_fingerprints"] = add_fingerprints(self.watch_seen_fingerprints(), fingerprints)
        self.save()

    def clear_watch_seen(self) -> None:
        self.state["watch_seen_fingerprints"] = []
        self.save()

    def automation_events(self) -> list[dict[str, Any]]:
        events = self.state.get("automation_events", [])
        return [item for item in events if isinstance(item, dict)]

    def add_automation_event(self, event: dict[str, Any]) -> None:
        record = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            **event,
        }
        events = self.automation_events()
        events.insert(0, record)
        self.state["automation_events"] = events[:MAX_AUTOMATION_EVENTS]
        self.save()
