from __future__ import annotations

import json
import os
import platform
import shutil
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from app_state import APP_NAME

EXCLUDED_DIR_NAMES = {
    "__pycache__",
    ".git",
    ".idea",
    ".vscode",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    ".venv",
    "venv",
    "env",
    "build",
    "dist",
    "smoke_test_output_example",
}
EXCLUDED_FILE_SUFFIXES = {".pyc", ".pyo", ".pyd"}
DEFAULT_RELEASE_DIRNAME = "release_output"


@dataclass(slots=True)
class PackagingPaths:
    project_dir: Path
    installer_dir: Path
    spec_path: Path
    build_notes_path: Path
    icon_png_path: Path
    icon_ico_path: Path
    inno_script_path: Path
    windows_build_script: Path
    linux_build_script: Path
    macos_build_script: Path
    installer_build_script: Path
    build_requirements_path: Path
    build_release_bundle_path: Path
    release_notes_template_path: Path
    about_profile_path: Path



def packaging_paths(project_dir: Path) -> PackagingPaths:
    project_dir = Path(project_dir).resolve()
    installer_dir = project_dir / "installer"
    return PackagingPaths(
        project_dir=project_dir,
        installer_dir=installer_dir,
        spec_path=installer_dir / "gokul_omni_convert_lite.spec",
        build_notes_path=installer_dir / "BUILDING.md",
        icon_png_path=project_dir / "assets" / "gokul_omni_convert_lite.png",
        icon_ico_path=project_dir / "assets" / "gokul_omni_convert_lite.ico",
        inno_script_path=installer_dir / "GokulOmniConvertLite.iss",
        windows_build_script=installer_dir / "build_windows.bat",
        linux_build_script=installer_dir / "build_linux.sh",
        macos_build_script=installer_dir / "build_macos.sh",
        installer_build_script=installer_dir / "build_installer_windows.bat",
        build_requirements_path=installer_dir / "requirements-build.txt",
        build_release_bundle_path=installer_dir / "build_release_bundle.py",
        release_notes_template_path=installer_dir / "release_notes_template.md",
        about_profile_path=project_dir / "about_profile.json",
    )



def _check_python_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False



def build_packaging_report(project_dir: Path, app_version: str = "") -> dict[str, Any]:
    paths = packaging_paths(project_dir)
    pyinstaller_cli = shutil.which("pyinstaller") or shutil.which("pyinstaller.exe")
    inno_cli = shutil.which("iscc") or shutil.which("iscc.exe")

    file_map = {
        "spec": paths.spec_path,
        "build_notes": paths.build_notes_path,
        "icon_png": paths.icon_png_path,
        "icon_ico": paths.icon_ico_path,
        "inno_script": paths.inno_script_path,
        "build_windows": paths.windows_build_script,
        "build_linux": paths.linux_build_script,
        "build_macos": paths.macos_build_script,
        "build_installer_windows": paths.installer_build_script,
        "build_requirements": paths.build_requirements_path,
        "build_release_bundle": paths.build_release_bundle_path,
        "release_notes_template": paths.release_notes_template_path,
        "about_profile": paths.about_profile_path,
    }

    files: dict[str, dict[str, Any]] = {}
    missing: list[str] = []
    for key, path in file_map.items():
        exists = path.exists()
        files[key] = {
            "path": str(path),
            "exists": exists,
            "size": path.stat().st_size if exists and path.is_file() else None,
        }
        if not exists:
            missing.append(key)

    return {
        "app_name": APP_NAME,
        "app_version": app_version,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "python": sys.version.split()[0],
        },
        "tools": {
            "pyinstaller_cli": pyinstaller_cli or "",
            "pyinstaller_module": _check_python_module("PyInstaller"),
            "inno_setup_cli": inno_cli or "",
        },
        "paths": {
            "project_dir": str(paths.project_dir),
            "installer_dir": str(paths.installer_dir),
        },
        "files": files,
        "missing": missing,
        "recommended_release_dir": str(paths.project_dir / DEFAULT_RELEASE_DIRNAME),
    }



def render_packaging_report(report: dict[str, Any]) -> str:
    tools = report.get("tools", {}) if isinstance(report.get("tools"), dict) else {}
    files = report.get("files", {}) if isinstance(report.get("files"), dict) else {}
    missing = report.get("missing", []) if isinstance(report.get("missing"), list) else []

    lines = [
        f"{report.get('app_name', APP_NAME)} packaging report",
        f"Version: {report.get('app_version', '').strip() or 'Unspecified'}",
        f"Generated: {report.get('generated_at', '')}",
        "",
        "Tooling",
        f"- PyInstaller CLI: {'available' if tools.get('pyinstaller_cli') else 'not found on PATH'}",
        f"- PyInstaller module: {'available' if tools.get('pyinstaller_module') else 'not installed in this Python environment'}",
        f"- Inno Setup compiler: {'available' if tools.get('inno_setup_cli') else 'not found on PATH'}",
        "",
        "Packaging assets",
    ]
    for label, data in files.items():
        if not isinstance(data, dict):
            continue
        status = "OK" if data.get("exists") else "MISSING"
        lines.append(f"- {label}: {status} -> {data.get('path', '')}")
    lines.extend([
        "",
        f"Missing assets: {', '.join(missing) if missing else 'none'}",
        f"Recommended release folder: {report.get('recommended_release_dir', '')}",
        "",
        "Typical next steps",
        "1. Fill the About page with your real profile image and social links.",
        "2. Run the PyInstaller build helper from installer/.",
        "3. If you are on Windows, wrap dist/GokulOmniConvertLite with the Inno Setup script.",
        "4. Keep a source ZIP or portable bundle for backup and handoff.",
    ])
    return "\n".join(lines)



def export_packaging_manifest(report: dict[str, Any], output_path: Path) -> Path:
    output_path = Path(output_path).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return output_path



def _should_skip(path: Path, project_dir: Path, release_dir: Path | None = None) -> bool:
    rel = path.relative_to(project_dir)
    if rel.parts and rel.parts[0] in EXCLUDED_DIR_NAMES:
        return True
    if release_dir is not None:
        try:
            path.relative_to(release_dir)
            return True
        except Exception:
            pass
    if path.suffix.lower() in EXCLUDED_FILE_SUFFIXES:
        return True
    if path.name in {".DS_Store"}:
        return True
    if path.name.startswith("gokul_omni_convert_lite_patch") and path.suffix.lower() == ".zip":
        return True
    return False



def iter_project_files(project_dir: Path, release_dir: Path | None = None) -> Iterable[Path]:
    project_dir = Path(project_dir).resolve()
    for path in sorted(project_dir.rglob("*")):
        if path.is_dir():
            continue
        if _should_skip(path, project_dir, release_dir):
            continue
        yield path



def create_portable_source_bundle(project_dir: Path, destination_dir: Path, release_name: str = "") -> Path:
    project_dir = Path(project_dir).resolve()
    destination_dir = Path(destination_dir).expanduser().resolve()
    destination_dir.mkdir(parents=True, exist_ok=True)

    safe_name = release_name.strip() or f"gokul_omni_convert_lite_source_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    archive_path = destination_dir / f"{safe_name}.zip"
    base_folder_name = safe_name

    with zipfile.ZipFile(archive_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in iter_project_files(project_dir, destination_dir):
            rel = file_path.relative_to(project_dir)
            zf.write(file_path, arcname=str(Path(base_folder_name) / rel))
    return archive_path



def create_portable_layout(project_dir: Path, destination_dir: Path, release_name: str = "") -> Path:
    project_dir = Path(project_dir).resolve()
    destination_dir = Path(destination_dir).expanduser().resolve()
    destination_dir.mkdir(parents=True, exist_ok=True)

    folder_name = release_name.strip() or f"gokul_omni_convert_lite_portable_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    release_root = destination_dir / folder_name
    if release_root.exists():
        shutil.rmtree(release_root)
    release_root.mkdir(parents=True, exist_ok=True)

    for file_path in iter_project_files(project_dir, destination_dir):
        rel = file_path.relative_to(project_dir)
        target = release_root / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(file_path, target)
    return release_root
