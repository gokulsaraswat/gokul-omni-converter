from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from automation_core import add_fingerprints, normalize_preset_record, normalize_watch_config
from engagement_core import ensure_install_date, parse_datetime

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
    "performance_mode": "balanced",
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
    "install_date": "",
    "login_popup_dismissed": False,
    "login_popup_completed": False,
    "login_popup_last_shown": "",
    "login_popup_enabled": True,
    "splash_enabled": True,
    "splash_seen": False,
    "splash_gif_path": "assets/gokul_splash.gif",
    "link_cache_dir": "",
    "link_timeout": 25,
    "link_keep_downloads": True,
    "link_cache_max_age_days": 30,
    "link_cache_max_size_mb": 512,
    "recent_links": [],
    "auto_open_output_folder": False,
    "restore_last_session": True,
    "cleanup_temp_on_exit": True,
    "recent_outputs": [],
    "failed_jobs": [],
    "session_snapshot": {},
    "update_checker_enabled": False,
    "last_update_check": "",
    "window_geometry": "",
    "last_page": "home",
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
            ensure_install_date(self.state)
            self.save()
            return self.state

        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self.state = dict(DEFAULT_STATE)
            ensure_install_date(self.state)
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
        if parse_datetime(merged.get("install_date")) is None:
            ensure_install_date(merged)
        for key in ("login_popup_dismissed", "login_popup_completed", "login_popup_enabled", "splash_enabled", "splash_seen"):
            merged[key] = bool(merged.get(key, DEFAULT_STATE[key]))
        if not isinstance(merged.get("login_popup_last_shown"), str):
            merged["login_popup_last_shown"] = ""
        if not isinstance(merged.get("splash_gif_path"), str):
            merged["splash_gif_path"] = DEFAULT_STATE["splash_gif_path"]
        if not isinstance(merged.get("link_cache_dir"), str):
            merged["link_cache_dir"] = DEFAULT_STATE["link_cache_dir"]
        if not isinstance(merged.get("performance_mode"), str):
            merged["performance_mode"] = DEFAULT_STATE["performance_mode"]
        merged["performance_mode"] = str(merged.get("performance_mode", DEFAULT_STATE["performance_mode"])).strip().lower() or DEFAULT_STATE["performance_mode"]
        if merged["performance_mode"] not in {"eco", "balanced", "quality"}:
            merged["performance_mode"] = DEFAULT_STATE["performance_mode"]
        try:
            merged["link_timeout"] = int(merged.get("link_timeout", DEFAULT_STATE["link_timeout"]))
        except Exception:
            merged["link_timeout"] = DEFAULT_STATE["link_timeout"]
        merged["link_keep_downloads"] = bool(merged.get("link_keep_downloads", DEFAULT_STATE["link_keep_downloads"]))
        try:
            merged["link_cache_max_age_days"] = max(0, int(merged.get("link_cache_max_age_days", DEFAULT_STATE["link_cache_max_age_days"])))
        except Exception:
            merged["link_cache_max_age_days"] = DEFAULT_STATE["link_cache_max_age_days"]
        try:
            merged["link_cache_max_size_mb"] = max(32, int(merged.get("link_cache_max_size_mb", DEFAULT_STATE["link_cache_max_size_mb"])))
        except Exception:
            merged["link_cache_max_size_mb"] = DEFAULT_STATE["link_cache_max_size_mb"]
        if not isinstance(merged.get("recent_links"), list):
            merged["recent_links"] = []
        else:
            merged["recent_links"] = [str(item).strip() for item in merged["recent_links"] if str(item).strip()][:50]
        merged["auto_open_output_folder"] = bool(merged.get("auto_open_output_folder", DEFAULT_STATE["auto_open_output_folder"]))
        merged["restore_last_session"] = bool(merged.get("restore_last_session", DEFAULT_STATE["restore_last_session"]))
        merged["cleanup_temp_on_exit"] = bool(merged.get("cleanup_temp_on_exit", DEFAULT_STATE["cleanup_temp_on_exit"]))
        merged["update_checker_enabled"] = bool(merged.get("update_checker_enabled", DEFAULT_STATE["update_checker_enabled"]))
        if not isinstance(merged.get("last_update_check"), str):
            merged["last_update_check"] = DEFAULT_STATE["last_update_check"]
        if not isinstance(merged.get("window_geometry"), str):
            merged["window_geometry"] = DEFAULT_STATE["window_geometry"]
        if not isinstance(merged.get("last_page"), str):
            merged["last_page"] = DEFAULT_STATE["last_page"]
        if not isinstance(merged.get("recent_outputs"), list):
            merged["recent_outputs"] = []
        else:
            merged["recent_outputs"] = [str(item).strip() for item in merged["recent_outputs"] if str(item).strip()][:120]
        if not isinstance(merged.get("failed_jobs"), list):
            merged["failed_jobs"] = []
        else:
            merged["failed_jobs"] = [item for item in merged["failed_jobs"] if isinstance(item, dict)][:60]
        if not isinstance(merged.get("session_snapshot"), dict):
            merged["session_snapshot"] = {}
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

    def favorite_presets(self) -> list[dict[str, Any]]:
        favorites = [item for item in self.presets() if bool(item.get("favorite", False))]
        favorites.sort(key=lambda item: str(item.get("name", "")).lower())
        return favorites

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

    def set_preset_favorite(self, name: str, favorite: bool) -> None:
        target = str(name).strip().lower()
        if not target:
            return
        updated: list[dict[str, Any]] = []
        changed = False
        for item in self.presets():
            normalized = normalize_preset_record(item).to_dict()
            if str(normalized.get("name", "")).strip().lower() == target:
                normalized["favorite"] = bool(favorite)
                changed = True
            updated.append(normalized)
        if changed:
            self.state["presets"] = updated
            self.save()

    def recent_links(self) -> list[str]:
        items = self.state.get("recent_links", [])
        return [str(item).strip() for item in items if str(item).strip()]

    def remember_links(self, links: Iterable[str]) -> None:
        existing = self.recent_links()
        seen = {item.lower() for item in existing}
        merged = list(existing)
        for link in links:
            value = str(link).strip()
            if not value:
                continue
            key = value.lower()
            if key in seen:
                continue
            seen.add(key)
            merged.insert(0, value)
        self.state["recent_links"] = merged[:50]
        self.save()

    def clear_recent_links(self) -> None:
        self.state["recent_links"] = []
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


    def recent_outputs(self) -> list[str]:
        items = self.state.get("recent_outputs", [])
        return [str(item).strip() for item in items if str(item).strip()]

    def remember_outputs(self, paths: Iterable[str | Path]) -> None:
        existing = self.recent_outputs()
        seen = {item.lower() for item in existing}
        merged = list(existing)
        for raw_path in paths:
            value = str(raw_path).strip()
            if not value:
                continue
            key = value.lower()
            if key in seen:
                continue
            seen.add(key)
            merged.insert(0, value)
        self.state["recent_outputs"] = merged[:120]
        self.save()

    def clear_recent_outputs(self) -> None:
        self.state["recent_outputs"] = []
        self.save()

    def failed_jobs(self) -> list[dict[str, Any]]:
        items = self.state.get("failed_jobs", [])
        return [item for item in items if isinstance(item, dict)]

    def add_failed_job(self, job: dict[str, Any]) -> str:
        job_id = str(job.get("id", "")).strip() or datetime.now().strftime("failed_%Y%m%d_%H%M%S_%f")
        record = {"id": job_id, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), **job, "id": job_id}
        jobs = self.failed_jobs()
        jobs.insert(0, record)
        self.state["failed_jobs"] = jobs[:60]
        self.save()
        return job_id

    def remove_failed_job(self, job_id: str) -> None:
        target = str(job_id).strip()
        if not target:
            return
        self.state["failed_jobs"] = [item for item in self.failed_jobs() if str(item.get("id", "")).strip() != target]
        self.save()

    def clear_failed_jobs(self) -> None:
        self.state["failed_jobs"] = []
        self.save()

    def session_snapshot(self) -> dict[str, Any]:
        snapshot = self.state.get("session_snapshot", {})
        return snapshot if isinstance(snapshot, dict) else {}

    def set_session_snapshot(self, snapshot: dict[str, Any]) -> None:
        self.state["session_snapshot"] = snapshot if isinstance(snapshot, dict) else {}
        self.save()
