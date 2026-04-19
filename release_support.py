from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any


def version_key(value: str) -> tuple[int, ...]:
    text = str(value or "").strip().lower()
    if not text:
        return (0,)
    digits: list[int] = []
    token = ""
    for char in text:
        if char.isdigit():
            token += char
        else:
            if token:
                digits.append(int(token))
                token = ""
    if token:
        digits.append(int(token))
    return tuple(digits or [0])


def compare_versions(current_version: str, latest_version: str) -> int:
    current_key = version_key(current_version)
    latest_key = version_key(latest_version)
    size = max(len(current_key), len(latest_key))
    current_key += (0,) * (size - len(current_key))
    latest_key += (0,) * (size - len(latest_key))
    if current_key < latest_key:
        return -1
    if current_key > latest_key:
        return 1
    return 0


def _read_text_from_source(source: str, timeout: int = 8) -> str:
    value = str(source or "").strip()
    if not value:
        raise ValueError("No update manifest source was provided.")
    parsed = urllib.parse.urlparse(value)
    if parsed.scheme in {"http", "https"}:
        request = urllib.request.Request(value, headers={"User-Agent": "GokulOmniConvertLite/1.0"})
        with urllib.request.urlopen(request, timeout=max(3, int(timeout))) as response:
            return response.read().decode("utf-8")
    if parsed.scheme == "file":
        return Path(urllib.request.url2pathname(parsed.path)).read_text(encoding="utf-8")
    return Path(value).expanduser().read_text(encoding="utf-8")


def load_update_manifest(source: str, timeout: int = 8) -> dict[str, Any]:
    raw_text = _read_text_from_source(source, timeout=timeout)
    payload = json.loads(raw_text)
    if not isinstance(payload, dict):
        raise ValueError("Update manifest must be a JSON object.")
    return payload


def check_for_updates(current_version: str, source: str, *, timeout: int = 8) -> dict[str, Any]:
    checked_at = datetime.now().isoformat(timespec="seconds")
    cleaned_source = str(source or "").strip()
    result: dict[str, Any] = {
        "checked_at": checked_at,
        "source": cleaned_source,
        "current_version": str(current_version or "").strip(),
        "latest_version": "",
        "has_update": False,
        "status": "error",
        "message": "",
        "notes": "",
        "download_url": "",
    }
    if not cleaned_source:
        result["status"] = "missing_source"
        result["message"] = "Choose a local update manifest JSON or provide an HTTP/HTTPS URL first."
        return result
    try:
        manifest = load_update_manifest(cleaned_source, timeout=timeout)
    except FileNotFoundError:
        result["status"] = "error"
        result["message"] = f"Update manifest not found: {cleaned_source}"
        return result
    except urllib.error.URLError as exc:
        result["status"] = "error"
        result["message"] = f"Could not reach the update manifest: {exc}"
        return result
    except Exception as exc:
        result["status"] = "error"
        result["message"] = f"Could not parse the update manifest: {exc}"
        return result

    latest_version = str(manifest.get("version", "")).strip() or str(current_version or "").strip()
    notes = str(manifest.get("notes", "")).strip()
    download_url = str(manifest.get("download_url", "")).strip()
    compare = compare_versions(str(current_version or "").strip(), latest_version)

    result["latest_version"] = latest_version
    result["notes"] = notes
    result["download_url"] = download_url
    result["status"] = "ok"
    if compare < 0:
        result["has_update"] = True
        result["message"] = f"Update available: {latest_version} is newer than {current_version}."
    elif compare > 0:
        result["message"] = f"Current version {current_version} is newer than manifest version {latest_version}."
    else:
        result["message"] = f"You are on the latest version: {current_version}."
    return result


def build_example_update_manifest(destination: str | Path, current_version: str) -> Path:
    path = Path(destination).expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": str(current_version).strip(),
        "published_at": datetime.now().isoformat(timespec="seconds"),
        "notes": "Replace this JSON with your own release feed later. Keep the version string ahead of the app build when you want the checker to surface an update.",
        "download_url": "",
        "changelog_url": "",
        "min_supported_version": "",
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def export_workspace_bundle(
    destination_zip: str | Path,
    *,
    state_path: str | Path,
    notes_path: str | Path,
    about_profile_path: str | Path,
    static_about_profile_path: str | Path,
    installer_dir: str | Path,
    asset_config_path: str | Path | None = None,
    extra_files: list[str | Path] | None = None,
) -> Path:
    destination = Path(destination_zip).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    files_to_pack: list[Path] = []
    seed_files: list[str | Path | None] = [state_path, notes_path, about_profile_path, static_about_profile_path, asset_config_path]
    for item in seed_files:
        if item is None:
            continue
        path = Path(item).expanduser()
        if path.exists() and path.is_file():
            files_to_pack.append(path)
    installer_root = Path(installer_dir).expanduser()
    if installer_root.exists():
        for item in sorted(installer_root.rglob("*")):
            if item.is_file():
                files_to_pack.append(item)
    for item in extra_files or []:
        path = Path(item).expanduser()
        if path.exists() and path.is_file():
            files_to_pack.append(path)

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "file_count": len(files_to_pack),
        "files": [path.name for path in files_to_pack],
    }
    seen_names: set[str] = set()
    with zipfile.ZipFile(destination, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("workspace_manifest.json", json.dumps(manifest, indent=2))
        for index, path in enumerate(files_to_pack, start=1):
            if installer_root in path.parents:
                arcname = str(Path("installer") / path.relative_to(installer_root))
            else:
                arcname = path.name
            while arcname in seen_names:
                arc_path = Path(arcname)
                arcname = str(arc_path.with_name(f"{arc_path.stem}_{index}{arc_path.suffix}"))
            seen_names.add(arcname)
            archive.write(path, arcname=arcname)
    return destination


def import_workspace_bundle(bundle_path: str | Path, target_root: str | Path) -> dict[str, Any]:
    bundle = Path(bundle_path).expanduser()
    destination_root = Path(target_root).expanduser()
    destination_root.mkdir(parents=True, exist_ok=True)
    extracted: list[str] = []
    with zipfile.ZipFile(bundle) as archive:
        for member in archive.infolist():
            member_path = Path(member.filename)
            if member.is_dir():
                continue
            if member_path.is_absolute() or ".." in member_path.parts:
                continue
            target = destination_root / member_path
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(member) as source_handle, target.open("wb") as target_handle:
                target_handle.write(source_handle.read())
            extracted.append(str(target))
    return {
        "bundle": str(bundle),
        "target_root": str(destination_root),
        "extracted_count": len(extracted),
        "extracted": extracted,
    }
