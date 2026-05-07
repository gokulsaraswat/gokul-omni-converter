from __future__ import annotations

import json
import py_compile
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

from PIL import Image


CORE_COMPILE_TARGETS = [
    "app.py",
    "converter_core.py",
    "ocr_core.py",
    "image_folder_pdf.py",
    "app_state.py",
    "responsive_ui.py",
    "release_support.py",
    "release_validation.py",
]

REQUIRED_RELEASE_FILES = [
    "README.md",
    "requirements.txt",
    "about_profile.json",
    "remote_assets.json",
    "assets/gokul_header.gif",
    "assets/gokul_splash.gif",
    "assets/gokul_profile_placeholder.png",
    "installer/TESSERACT_BUNDLE.md",
    "installer/PYINSTALLER_BUILD.md",
    "installer/GokulOmniConvertLite.spec",
    "installer/update_manifest.example.json",
    "installer/about_static.json",
    "RELEASE_CHECKLIST.md",
    "PATCH_37_RELEASE_NOTES.md",
    "build_windows_release.bat",
    "build_linux_release.sh",
]


def _ok(name: str, detail: str = "") -> dict[str, Any]:
    return {"name": name, "status": "passed", "detail": detail}


def _fail(name: str, detail: str) -> dict[str, Any]:
    return {"name": name, "status": "failed", "detail": detail}


def _read_app_version(project_root: Path) -> str:
    text = (project_root / "app.py").read_text(encoding="utf-8")
    match = re.search(r'^APP_VERSION\s*=\s*["\']([^"\']+)["\']', text, flags=re.MULTILINE)
    return match.group(1) if match else ""


def _compile_files(project_root: Path) -> dict[str, Any]:
    missing = [name for name in CORE_COMPILE_TARGETS if not (project_root / name).exists()]
    if missing:
        return _fail("Python compile", f"Missing compile target(s): {', '.join(missing)}")
    try:
        for name in CORE_COMPILE_TARGETS:
            py_compile.compile(str(project_root / name), doraise=True)
    except Exception as exc:
        return _fail("Python compile", str(exc))
    return _ok("Python compile", f"Compiled {len(CORE_COMPILE_TARGETS)} core files.")


def _required_files(project_root: Path) -> dict[str, Any]:
    missing = [name for name in REQUIRED_RELEASE_FILES if not (project_root / name).exists()]
    if missing:
        return _fail("Release files", f"Missing: {', '.join(missing)}")
    return _ok("Release files", f"Found {len(REQUIRED_RELEASE_FILES)} release/build files.")


def _image_pdf_test(project_root: Path) -> dict[str, Any]:
    sys.path.insert(0, str(project_root))
    try:
        from converter_core import convert_images_to_single_pdf
        from image_folder_pdf import (
            IMAGE_FOLDER_SCOPE_ALL,
            IMAGE_FOLDER_SCOPE_PER_FOLDER,
            ImageFolderPdfConfig,
            build_image_folder_pdfs,
            summarize_image_folder,
        )
    finally:
        try:
            sys.path.remove(str(project_root))
        except ValueError:
            pass

    with tempfile.TemporaryDirectory(prefix="gokul_patch37_image_pdf_") as temp_dir:
        root = Path(temp_dir)
        source = root / "source"
        nested = source / "chapter_1"
        output_dir = root / "output"
        nested.mkdir(parents=True)
        output_dir.mkdir()
        images = []
        for index, folder in enumerate((source, source, nested), start=1):
            image_path = folder / f"page_{index:02d}.png"
            Image.new("RGBA", (170 + index * 20, 120 + index * 15), (40 * index, 90, 180, 160)).save(image_path)
            images.append(image_path)

        direct = convert_images_to_single_pdf(images, output_dir / "direct_images.pdf")
        if not direct.exists() or direct.stat().st_size <= 0:
            return _fail("Images -> PDF", "Direct Images -> PDF output was not created or is empty.")

        summary = summarize_image_folder(source, recursive=True)
        if int(summary.get("image_count", 0)) != 3:
            return _fail("Image folder discovery", f"Expected 3 recursive images, got {summary.get('image_count')}")

        all_outputs = build_image_folder_pdfs(
            ImageFolderPdfConfig(
                source_dir=source,
                output_dir=output_dir,
                output_name="combined_recursive",
                recursive=True,
                scope=IMAGE_FOLDER_SCOPE_ALL,
            )
        )
        per_folder_outputs = build_image_folder_pdfs(
            ImageFolderPdfConfig(
                source_dir=source,
                output_dir=output_dir,
                output_name="per_folder",
                recursive=True,
                scope=IMAGE_FOLDER_SCOPE_PER_FOLDER,
            )
        )
        for path in all_outputs + per_folder_outputs:
            if not path.exists() or path.stat().st_size <= 0:
                return _fail("Image folder PDF", f"Output missing or empty: {path}")

    return _ok("Image folder PDF", "Direct, all-images, and per-folder image PDF workflows passed.")


def _ocr_detection_test(project_root: Path) -> dict[str, Any]:
    sys.path.insert(0, str(project_root))
    try:
        from ocr_core import detect_tesseract_status
    finally:
        try:
            sys.path.remove(str(project_root))
        except ValueError:
            pass

    with tempfile.TemporaryDirectory(prefix="gokul_patch37_ocr_") as temp_dir:
        fake_wrapper = Path(temp_dir) / "pytesseract.exe"
        fake_wrapper.write_text("not a real binary", encoding="utf-8")
        invalid = detect_tesseract_status(fake_wrapper, language="eng")
        message = str(invalid.get("message", "")).lower()
        if bool(invalid.get("available")) or "python wrapper" not in message:
            return _fail("OCR detection", "pytesseract.exe wrapper was not rejected clearly.")
        general = detect_tesseract_status(language="eng")
        for key in ("available", "path", "source", "tessdata", "message"):
            if key not in general:
                return _fail("OCR detection", f"Missing status key: {key}")
    return _ok("OCR detection", "Status shape is complete and pytesseract.exe is rejected.")


def _version_consistency(project_root: Path, app_version: str) -> dict[str, Any]:
    if not app_version:
        return _fail("Version consistency", "APP_VERSION not found in app.py")
    required_mentions = [
        "README.md",
        "PATCH_37_RELEASE_NOTES.md",
        "installer/INSTALLER_RELEASE_NOTES.md",
        "installer/update_manifest.example.json",
        "installer/about_static.json",
    ]
    mismatches: list[str] = []
    for relative in required_mentions:
        path = project_root / relative
        if not path.exists():
            mismatches.append(f"{relative}: missing")
            continue
        if app_version not in path.read_text(encoding="utf-8", errors="ignore"):
            mismatches.append(relative)
    if mismatches:
        return _fail("Version consistency", f"Version {app_version} missing from: {', '.join(mismatches)}")
    return _ok("Version consistency", f"Release files mention {app_version}.")


def run_release_checks(project_root: str | Path | None = None) -> dict[str, Any]:
    root = Path(project_root or Path(__file__).resolve().parent).resolve()
    app_version = _read_app_version(root)
    checks = [
        _compile_files(root),
        _required_files(root),
        _image_pdf_test(root),
        _ocr_detection_test(root),
        _version_consistency(root, app_version),
    ]
    failed = [item for item in checks if item.get("status") != "passed"]
    return {
        "status": "failed" if failed else "passed",
        "app": "Gokul Omni Convert Lite",
        "version": app_version,
        "project_root": str(root),
        "checks": checks,
    }


def main() -> None:
    result = run_release_checks()
    print(json.dumps(result, indent=2))
    if result["status"] != "passed":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
