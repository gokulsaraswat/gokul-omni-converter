from __future__ import annotations

import json
import shutil
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

from converter_core import ENGINE_AUTO, MODE_ANY_TO_PDF, supported_extensions_for_mode


MAX_SEEN_FINGERPRINTS = 5000


@dataclass(slots=True)
class PresetRecord:
    name: str
    mode: str = MODE_ANY_TO_PDF
    output_dir: str = ""
    merge_to_one_pdf: bool = False
    merged_output_name: str = ""
    image_format: str = "png"
    image_scale: float = 2.0
    engine_mode: str = ENGINE_AUTO
    recursive: bool = True
    created_at: str = ""
    updated_at: str = ""

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(slots=True)
class WatchFolderConfig:
    source_dir: str = ""
    output_dir: str = ""
    mode: str = MODE_ANY_TO_PDF
    merge_to_one_pdf: bool = False
    merged_output_name: str = "watch_output"
    recursive: bool = True
    interval_seconds: int = 15
    engine_mode: str = ENGINE_AUTO
    archive_processed: bool = False
    archive_dir: str = ""
    create_zip_bundle: bool = False
    create_report: bool = True
    open_mail_draft: bool = False
    skip_existing_on_start: bool = True

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def normalize_preset_record(data: dict[str, object] | PresetRecord | None) -> PresetRecord:
    if isinstance(data, PresetRecord):
        return data
    payload = data or {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    name = str(payload.get("name", "New preset")).strip() or "New preset"
    created_at = str(payload.get("created_at", now)).strip() or now
    updated_at = str(payload.get("updated_at", now)).strip() or now
    image_scale_raw = payload.get("image_scale", 2.0)
    try:
        image_scale = float(image_scale_raw)
    except (TypeError, ValueError):
        image_scale = 2.0
    return PresetRecord(
        name=name,
        mode=str(payload.get("mode", MODE_ANY_TO_PDF)).strip() or MODE_ANY_TO_PDF,
        output_dir=str(payload.get("output_dir", "")).strip(),
        merge_to_one_pdf=bool(payload.get("merge_to_one_pdf", False)),
        merged_output_name=str(payload.get("merged_output_name", "")).strip(),
        image_format=str(payload.get("image_format", "png")).strip() or "png",
        image_scale=image_scale,
        engine_mode=str(payload.get("engine_mode", ENGINE_AUTO)).strip().lower() or ENGINE_AUTO,
        recursive=bool(payload.get("recursive", True)),
        created_at=created_at,
        updated_at=updated_at,
    )


def normalize_watch_config(data: dict[str, object] | WatchFolderConfig | None) -> WatchFolderConfig:
    if isinstance(data, WatchFolderConfig):
        return data
    payload = data or {}
    interval_raw = payload.get("interval_seconds", 15)
    try:
        interval_seconds = max(2, int(interval_raw))
    except (TypeError, ValueError):
        interval_seconds = 15
    return WatchFolderConfig(
        source_dir=str(payload.get("source_dir", "")).strip(),
        output_dir=str(payload.get("output_dir", "")).strip(),
        mode=str(payload.get("mode", MODE_ANY_TO_PDF)).strip() or MODE_ANY_TO_PDF,
        merge_to_one_pdf=bool(payload.get("merge_to_one_pdf", False)),
        merged_output_name=str(payload.get("merged_output_name", "watch_output")).strip() or "watch_output",
        recursive=bool(payload.get("recursive", True)),
        interval_seconds=interval_seconds,
        engine_mode=str(payload.get("engine_mode", ENGINE_AUTO)).strip().lower() or ENGINE_AUTO,
        archive_processed=bool(payload.get("archive_processed", False)),
        archive_dir=str(payload.get("archive_dir", "")).strip(),
        create_zip_bundle=bool(payload.get("create_zip_bundle", False)),
        create_report=bool(payload.get("create_report", True)),
        open_mail_draft=bool(payload.get("open_mail_draft", False)),
        skip_existing_on_start=bool(payload.get("skip_existing_on_start", True)),
    )


def fingerprint_file(path: Path) -> str:
    target = Path(path).expanduser().resolve()
    try:
        stat = target.stat()
        return f"{target}|{stat.st_size}|{stat.st_mtime_ns}"
    except OSError:
        return str(target)


def discover_watch_candidates(
    source_dir: Path,
    mode: str,
    recursive: bool,
    seen_fingerprints: Iterable[str],
) -> tuple[list[Path], list[str]]:
    folder = Path(source_dir).expanduser()
    if not folder.exists() or not folder.is_dir():
        return [], []

    supported = supported_extensions_for_mode(mode)
    globber = folder.rglob("*") if recursive else folder.glob("*")
    seen = set(str(item) for item in seen_fingerprints)
    files: list[Path] = []
    fingerprints: list[str] = []
    for candidate in globber:
        if not candidate.is_file():
            continue
        if candidate.suffix.lower() not in supported:
            continue
        fingerprint = fingerprint_file(candidate)
        if fingerprint in seen:
            continue
        files.append(candidate.resolve())
        fingerprints.append(fingerprint)
    files.sort(key=lambda path: str(path).lower())
    fingerprints = [fingerprint_file(path) for path in files]
    return files, fingerprints


def add_fingerprints(existing: Iterable[str], new_items: Iterable[str], limit: int = MAX_SEEN_FINGERPRINTS) -> list[str]:
    merged: list[str] = []
    seen: set[str] = set()
    for item in list(existing) + list(new_items):
        token = str(item).strip()
        if not token or token in seen:
            continue
        seen.add(token)
        merged.append(token)
    if len(merged) > limit:
        merged = merged[-limit:]
    return merged


def move_files_to_archive(files: Iterable[Path], source_root: Path, archive_root: Path) -> list[Path]:
    moved: list[Path] = []
    source_root = Path(source_root).expanduser().resolve()
    archive_root = Path(archive_root).expanduser().resolve()
    archive_root.mkdir(parents=True, exist_ok=True)

    for path in files:
        target = Path(path).expanduser().resolve()
        if not target.exists():
            continue
        try:
            relative = target.relative_to(source_root)
        except ValueError:
            relative = Path(target.name)
        destination = archive_root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        counter = 1
        stem = destination.stem
        suffix = destination.suffix
        while destination.exists():
            destination = destination.with_name(f"{stem}_{counter}{suffix}")
            counter += 1
        shutil.move(str(target), str(destination))
        moved.append(destination)
    return moved


def bundle_paths_as_zip(paths: Iterable[Path], destination_zip: Path) -> Path:
    destination_zip = Path(destination_zip).expanduser()
    destination_zip.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, item in enumerate(paths, start=1):
            path = Path(item).expanduser()
            if not path.exists() or not path.is_file():
                continue
            name = path.name
            if name in archive.namelist():
                name = f"{path.stem}_{index}{path.suffix}"
            archive.write(path, arcname=name)
    return destination_zip


def write_run_report(job: dict[str, object], destination: Path) -> Path:
    destination = Path(destination).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = [
        "Gokul Omni Convert Lite Run Report",
        "=" * 36,
        "",
    ]
    for key in (
        "timestamp",
        "status",
        "mode",
        "job_type",
        "tool",
        "file_count",
        "output_count",
        "output_dir",
        "merge_to_one_pdf",
        "merged_output_name",
        "engine_mode",
    ):
        value = job.get(key, "")
        if value not in ("", None):
            lines.append(f"{key}: {value}")

    inputs_preview = job.get("inputs_preview", []) or []
    outputs_preview = job.get("outputs_preview", []) or []
    error_text = str(job.get("error", "")).strip()

    if inputs_preview:
        lines.extend(["", "Inputs preview:"])
        lines.extend(f"- {item}" for item in inputs_preview)
    if outputs_preview:
        lines.extend(["", "Outputs preview:"])
        lines.extend(f"- {item}" for item in outputs_preview)
    if error_text:
        lines.extend(["", "Error details:", error_text])

    destination.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return destination


def export_presets_to_json(presets: Iterable[dict[str, object]], destination: Path) -> Path:
    normalized = [normalize_preset_record(item).to_dict() for item in presets]
    destination = Path(destination).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "presets": normalized,
    }
    destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return destination


def import_presets_from_json(source: Path) -> list[dict[str, object]]:
    source = Path(source).expanduser()
    data = json.loads(source.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        raw_presets = data.get("presets", [])
    elif isinstance(data, list):
        raw_presets = data
    else:
        raw_presets = []
    presets: list[dict[str, object]] = []
    for item in raw_presets:
        if isinstance(item, dict):
            presets.append(normalize_preset_record(item).to_dict())
    return presets
