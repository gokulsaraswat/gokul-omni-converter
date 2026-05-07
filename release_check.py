from __future__ import annotations

import importlib.util
import json
import platform
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


REQUIRED_RUNTIME_FILES = [
    "app.py",
    "converter_core.py",
    "ocr_core.py",
    "image_folder_pdf.py",
    "app_state.py",
    "README.md",
    "requirements.txt",
    "remote_assets.json",
    "about_profile.json",
    "footer_notes.md",
]

RECOMMENDED_INSTALLER_FILES = [
    "installer/TESSERACT_BUNDLE.md",
    "installer/PYINSTALLER_BUILD.md",
    "installer/GokulOmniConvertLite.spec",
    "installer/update_manifest.example.json",
]

REQUIRED_PACKAGES = [
    "PIL",
    "fitz",
    "pdfplumber",
    "docx",
    "openpyxl",
    "pypdf",
    "reportlab",
    "pptx",
    "xlrd",
    "pytesseract",
]

OPTIONAL_RUNTIME_FOLDERS = ["assets", "installer", "tesseract", "Tesseract-OCR"]


@dataclass(slots=True)
class ReleaseCheckItem:
    name: str
    status: str
    detail: str = ""


@dataclass(slots=True)
class ReleaseCheckResult:
    app_name: str
    app_version: str
    created_at: str
    root: str
    platform: str
    python: str
    status: str
    errors: int
    warnings: int
    items: list[ReleaseCheckItem]

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["items"] = [asdict(item) for item in self.items]
        return payload


def _item(name: str, status: str, detail: str = "") -> ReleaseCheckItem:
    return ReleaseCheckItem(name=name, status=status, detail=detail)


def _check_files(root: Path, relatives: Iterable[str], *, required: bool) -> list[ReleaseCheckItem]:
    items: list[ReleaseCheckItem] = []
    missing_status = "error" if required else "warning"
    for relative in relatives:
        path = root / relative
        if path.exists():
            detail = f"present ({path.stat().st_size} bytes)" if path.is_file() else "present"
            items.append(_item(f"file:{relative}", "ok", detail))
        else:
            items.append(_item(f"file:{relative}", missing_status, "missing"))
    return items


def _check_imports() -> list[ReleaseCheckItem]:
    items: list[ReleaseCheckItem] = []
    for package in REQUIRED_PACKAGES:
        if importlib.util.find_spec(package) is not None:
            items.append(_item(f"package:{package}", "ok", "importable"))
        else:
            items.append(_item(f"package:{package}", "warning", "not importable in this environment"))
    return items


def _check_json_file(path: Path, label: str) -> ReleaseCheckItem:
    if not path.exists():
        return _item(label, "warning", "file missing")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return _item(label, "error", f"invalid JSON: {exc}")
    if not isinstance(payload, dict):
        return _item(label, "error", "JSON root must be an object")
    return _item(label, "ok", f"{len(payload)} top-level keys")


def _check_tesseract_layout(root: Path) -> list[ReleaseCheckItem]:
    items: list[ReleaseCheckItem] = []
    found_any = False
    for folder_name in ("tesseract", "Tesseract-OCR"):
        folder = root / folder_name
        binary_candidates = [folder / "tesseract.exe", folder / "tesseract"]
        trained = folder / "tessdata" / "eng.traineddata"
        if not folder.exists():
            items.append(_item(f"tesseract-bundle:{folder_name}", "warning", "optional bundle folder not present"))
            continue
        found_any = True
        binary = next((candidate for candidate in binary_candidates if candidate.exists()), None)
        items.append(_item(f"tesseract-bundle:{folder_name}:binary", "ok" if binary else "error", str(binary) if binary else "missing"))
        items.append(_item(f"tesseract-bundle:{folder_name}:tessdata", "ok" if trained.exists() else "warning", str(trained) if trained.exists() else "eng.traineddata not bundled"))
    if not found_any:
        items.append(_item("tesseract-bundle:summary", "warning", "OCR bundle is optional and not included in this source package"))
    return items


def run_release_check(root: str | Path, *, app_name: str, app_version: str) -> ReleaseCheckResult:
    root_path = Path(root).expanduser().resolve()
    items: list[ReleaseCheckItem] = []
    items.append(_item("root", "ok" if root_path.exists() else "error", str(root_path)))
    items.extend(_check_files(root_path, REQUIRED_RUNTIME_FILES, required=True))
    items.extend(_check_files(root_path, RECOMMENDED_INSTALLER_FILES, required=False))
    items.extend(_check_imports())
    items.append(_check_json_file(root_path / "remote_assets.json", "json:remote_assets"))
    items.append(_check_json_file(root_path / "about_profile.json", "json:about_profile"))
    items.extend(_check_tesseract_layout(root_path))
    for folder_name in OPTIONAL_RUNTIME_FOLDERS:
        path = root_path / folder_name
        items.append(_item(f"folder:{folder_name}", "ok" if path.exists() else "warning", "present" if path.exists() else "optional folder missing"))
    errors = sum(1 for item in items if item.status == "error")
    warnings = sum(1 for item in items if item.status == "warning")
    return ReleaseCheckResult(
        app_name=app_name,
        app_version=app_version,
        created_at=datetime.now().isoformat(timespec="seconds"),
        root=str(root_path),
        platform=platform.platform(),
        python=sys.version,
        status="failed" if errors else ("warning" if warnings else "ok"),
        errors=errors,
        warnings=warnings,
        items=items,
    )


def render_release_check_markdown(result: ReleaseCheckResult) -> str:
    lines = [
        f"# {result.app_name} Release Check",
        "",
        f"- Version: `{result.app_version}`",
        f"- Created: `{result.created_at}`",
        f"- Root: `{result.root}`",
        f"- Status: **{result.status.upper()}**",
        f"- Errors: **{result.errors}**",
        f"- Warnings: **{result.warnings}**",
        "",
        "## Checks",
        "",
        "| Status | Check | Detail |",
        "|---|---|---|",
    ]
    for item in result.items:
        detail = str(item.detail or "").replace("|", "\\|").replace("\n", " ")
        lines.append(f"| {item.status} | `{item.name}` | {detail} |")
    lines.append("")
    return "\n".join(lines)


def write_release_check_report(result: ReleaseCheckResult, destination: str | Path) -> Path:
    target = Path(destination).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.suffix.lower() == ".json":
        target.write_text(json.dumps(result.to_dict(), indent=2), encoding="utf-8")
    else:
        target.write_text(render_release_check_markdown(result), encoding="utf-8")
    return target


if __name__ == "__main__":
    result = run_release_check(Path.cwd(), app_name="Gokul Omni Convert Lite", app_version="development")
    print(render_release_check_markdown(result))
    raise SystemExit(1 if result.errors else 0)
