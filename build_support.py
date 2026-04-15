from __future__ import annotations

import json
import platform
import sys
from datetime import datetime
from importlib import metadata
from pathlib import Path
from typing import Any, Iterable


PACKAGE_NAMES = [
    "Pillow",
    "PyMuPDF",
    "pdfplumber",
    "python-docx",
    "openpyxl",
    "pypdf",
    "reportlab",
    "python-pptx",
    "xlrd",
    "pytesseract",
]



def collect_package_versions(package_names: Iterable[str] = PACKAGE_NAMES) -> dict[str, str]:
    versions: dict[str, str] = {}
    for name in package_names:
        try:
            versions[name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            versions[name] = "not installed"
        except Exception:
            versions[name] = "unknown"
    return versions



def collect_installer_assets(installer_dir: Path) -> list[str]:
    root = Path(installer_dir)
    if not root.exists():
        return []
    assets: list[str] = []
    for path in sorted(root.rglob("*")):
        if path.is_file():
            assets.append(str(path.relative_to(root)))
    return assets



def export_json(path: Path, payload: dict[str, Any]) -> Path:
    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return destination


def export_text_file(path: Path, content: str) -> Path:
    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(str(content), encoding="utf-8")
    return destination



def export_diagnostics_report(
    path: Path,
    *,
    app_name: str,
    app_version: str,
    state_path: Path,
    about_profile_path: Path,
    notes_path: Path,
    installer_dir: Path,
    output_dir: Path,
    selected_files: list[str],
    last_outputs: list[str],
    dependency_status: dict[str, Any],
    smtp_summary: dict[str, Any],
    extra: dict[str, Any] | None = None,
) -> Path:
    payload: dict[str, Any] = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "app": {
            "name": app_name,
            "version": app_version,
        },
        "system": {
            "platform": platform.platform(),
            "python": sys.version,
            "python_executable": sys.executable,
            "cwd": str(Path.cwd()),
        },
        "paths": {
            "state_path": str(state_path),
            "about_profile_path": str(about_profile_path),
            "notes_path": str(notes_path),
            "installer_dir": str(installer_dir),
            "output_dir": str(output_dir),
        },
        "packages": collect_package_versions(),
        "installer_assets": collect_installer_assets(installer_dir),
        "dependencies": dependency_status,
        "smtp": smtp_summary,
        "current_context": {
            "selected_files": selected_files,
            "last_outputs": last_outputs,
        },
    }
    if extra:
        payload["extra"] = extra
    return export_json(path, payload)



def export_state_snapshot(path: Path, state: dict[str, Any]) -> Path:
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "state": state,
    }
    return export_json(path, payload)



def import_state_snapshot(path: Path) -> dict[str, Any]:
    source = Path(path).expanduser()
    data = json.loads(source.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Invalid state snapshot format.")
    payload = data.get("state", data)
    if not isinstance(payload, dict):
        raise ValueError("Invalid state snapshot payload.")
    return payload
