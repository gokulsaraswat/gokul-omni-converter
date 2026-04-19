from __future__ import annotations

import hashlib
import json
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

ASSET_CONFIG_PATH = Path(__file__).with_name("remote_assets.json")

DEFAULT_ASSET_CONFIG: dict[str, Any] = {
    "remote_enabled": False,
    "cache_dir": "",
    "timeout": 15,
    "refresh_hours": 24,
    "header_gif_path": "assets/gokul_header.gif",
    "splash_gif_path": "assets/gokul_splash.gif",
    "header_gif_url": "",
    "splash_gif_url": "",
    "about_image_url": "",
    "profile_json_url": "",
}


def load_asset_config(path: Path | None = None) -> dict[str, Any]:
    config_path = path or ASSET_CONFIG_PATH
    config_path.parent.mkdir(parents=True, exist_ok=True)
    if not config_path.exists():
        config_path.write_text(json.dumps(DEFAULT_ASSET_CONFIG, indent=2), encoding="utf-8")
        return dict(DEFAULT_ASSET_CONFIG)

    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        payload = {}

    config = dict(DEFAULT_ASSET_CONFIG)
    if isinstance(payload, dict):
        config.update({key: value for key, value in payload.items() if key in DEFAULT_ASSET_CONFIG})
    save_asset_config(config, config_path)
    return config


def save_asset_config(config: dict[str, Any], path: Path | None = None) -> Path:
    config_path = path or ASSET_CONFIG_PATH
    config_path.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(DEFAULT_ASSET_CONFIG)
    if isinstance(config, dict):
        payload.update({key: value for key, value in config.items() if key in DEFAULT_ASSET_CONFIG})
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return config_path


def normalize_remote_url(url: str) -> str:
    value = str(url or "").strip()
    if not value:
        return ""
    parsed = urllib.parse.urlparse(value)
    if parsed.scheme not in {"http", "https", "file"}:
        return value
    if parsed.scheme in {"http", "https"} and parsed.netloc.lower() == "github.com" and "/blob/" in parsed.path:
        parts = [part for part in parsed.path.split("/") if part]
        if len(parts) >= 5 and parts[2] == "blob":
            user = parts[0]
            repo = parts[1]
            branch = parts[3]
            rest = "/".join(parts[4:])
            return f"https://raw.githubusercontent.com/{user}/{repo}/{branch}/{rest}"
    return value


def is_remote_reference(value: str) -> bool:
    parsed = urllib.parse.urlparse(str(value or "").strip())
    return parsed.scheme in {"http", "https", "file"}


def asset_cache_root(config: dict[str, Any] | None = None, override: str | Path | None = None) -> Path:
    if override:
        return Path(override).expanduser()
    if isinstance(config, dict) and str(config.get("cache_dir", "")).strip():
        return Path(str(config.get("cache_dir", "")).strip()).expanduser()
    return Path.home() / ".gokul_omni_convert_lite" / "asset_cache"


def cached_asset_path(url: str, cache_dir: str | Path) -> Path:
    normalized = normalize_remote_url(url)
    parsed = urllib.parse.urlparse(normalized)
    raw_name = Path(urllib.parse.unquote(parsed.path)).name or "asset"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", raw_name).strip("._") or "asset"
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:14]
    return Path(cache_dir).expanduser() / f"{digest}_{safe_name}"


def is_cache_fresh(path: Path, refresh_hours: int) -> bool:
    if not path.exists():
        return False
    if refresh_hours <= 0:
        return True
    modified = datetime.fromtimestamp(path.stat().st_mtime)
    return datetime.now() - modified <= timedelta(hours=max(1, refresh_hours))


def download_binary_asset(
    url: str,
    cache_dir: str | Path,
    *,
    timeout: int = 15,
) -> Path:
    normalized = normalize_remote_url(url)
    if not normalized:
        raise ValueError("No asset URL was provided.")
    destination = cached_asset_path(normalized, cache_dir)
    destination.parent.mkdir(parents=True, exist_ok=True)

    parsed = urllib.parse.urlparse(normalized)
    if parsed.scheme == "file":
        source_path = Path(urllib.request.url2pathname(parsed.path))
        if not source_path.exists():
            raise FileNotFoundError(f"Asset not found: {source_path}")
        data = source_path.read_bytes()
    else:
        request = urllib.request.Request(normalized, headers={"User-Agent": "GokulOmniConvertLite/2.2"})
        with urllib.request.urlopen(request, timeout=max(3, int(timeout))) as response:
            data = response.read()

    if not data:
        raise ValueError("Downloaded asset was empty.")

    destination.write_bytes(data)
    return destination


def download_text_file(
    url: str,
    destination: str | Path,
    *,
    timeout: int = 15,
) -> Path:
    normalized = normalize_remote_url(url)
    if not normalized:
        raise ValueError("No source URL was provided.")
    target = Path(destination).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)

    parsed = urllib.parse.urlparse(normalized)
    if parsed.scheme == "file":
        source_path = Path(urllib.request.url2pathname(parsed.path))
        if not source_path.exists():
            raise FileNotFoundError(f"Text source not found: {source_path}")
        text = source_path.read_text(encoding="utf-8")
    else:
        request = urllib.request.Request(normalized, headers={"User-Agent": "GokulOmniConvertLite/2.2"})
        with urllib.request.urlopen(request, timeout=max(3, int(timeout))) as response:
            text = response.read().decode("utf-8")

    target.write_text(text, encoding="utf-8")
    return target


def resolve_local_or_remote_asset(
    local_value: str,
    remote_value: str,
    *,
    base_dir: str | Path,
    fallback_value: str | Path | None = None,
    config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    root = Path(base_dir).expanduser()
    fallback_text = str(fallback_value) if fallback_value is not None else ""
    remote_enabled = bool((config or {}).get("remote_enabled", False))
    timeout = max(3, int((config or {}).get("timeout", 15) or 15))
    refresh_hours = max(1, int((config or {}).get("refresh_hours", 24) or 24))
    cache_dir = asset_cache_root(config)

    normalized_remote = normalize_remote_url(remote_value)
    remote_error = ""
    if remote_enabled and normalized_remote:
        cached = cached_asset_path(normalized_remote, cache_dir)
        if cached.exists() and is_cache_fresh(cached, refresh_hours):
            return {
                "path": cached,
                "source": "remote-cache",
                "url": normalized_remote,
                "message": f"Using cached asset: {cached.name}",
            }
        try:
            downloaded = download_binary_asset(normalized_remote, cache_dir, timeout=timeout)
            return {
                "path": downloaded,
                "source": "remote-download",
                "url": normalized_remote,
                "message": f"Downloaded asset from {normalized_remote}",
            }
        except Exception as exc:  # pragma: no cover - network dependent
            remote_error = str(exc)
            if cached.exists():
                return {
                    "path": cached,
                    "source": "remote-stale-cache",
                    "url": normalized_remote,
                    "message": f"Remote refresh failed, reusing cached asset: {exc}",
                }

    local_text = str(local_value or "").strip()
    if local_text:
        local_path = Path(local_text).expanduser()
        if not local_path.is_absolute():
            local_path = root / local_path
        if local_path.exists():
            return {
                "path": local_path,
                "source": "local",
                "url": "",
                "message": f"Using local asset: {local_path.name}",
            }

    if fallback_text:
        fallback_path = Path(fallback_text).expanduser()
        if not fallback_path.is_absolute():
            fallback_path = root / fallback_path
        if fallback_path.exists():
            message = f"Using fallback asset: {fallback_path.name}"
            if remote_error:
                message += f" (remote asset failed: {remote_error})"
            return {
                "path": fallback_path,
                "source": "fallback",
                "url": "",
                "message": message,
            }
        return {
            "path": fallback_path,
            "source": "missing",
            "url": "",
            "message": remote_error or f"Asset not found: {fallback_path}",
        }

    return {
        "path": root,
        "source": "missing",
        "url": normalized_remote,
        "message": remote_error or "No asset was configured.",
    }


def clear_asset_cache(config: dict[str, Any] | None = None, override: str | Path | None = None) -> Path:
    root = asset_cache_root(config, override)
    if root.exists():
        for path in sorted(root.rglob("*"), reverse=True):
            if path.is_file():
                path.unlink(missing_ok=True)
            elif path.is_dir():
                try:
                    path.rmdir()
                except OSError:
                    pass
    root.mkdir(parents=True, exist_ok=True)
    return root


def cache_summary(config: dict[str, Any] | None = None, override: str | Path | None = None) -> dict[str, Any]:
    root = asset_cache_root(config, override)
    count = 0
    total = 0
    if root.exists():
        for path in root.rglob("*"):
            if path.is_file():
                count += 1
                total += path.stat().st_size
    return {"path": str(root), "count": count, "bytes": total}
