"""Zip archive tools for GO Omni Pro.

Pure stdlib ``zipfile`` -- no new dependency. Follows the same conventions as
converter_core: raise ConversionError on failure, accept optional ``log`` and
``progress`` callbacks, return the list of paths produced.

Extraction is guarded against zip-slip (entries whose resolved path escapes the
destination directory), which stdlib's ``extractall`` does not fully protect
against for absolute paths on every platform.
"""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from converter_core import ConversionError, LogFn, ProgressFn, ensure_directory, unique_path

ARCHIVE_EXTENSION = ".zip"

COMPRESSION_LEVELS = {
    "store": (zipfile.ZIP_STORED, None),
    "fast": (zipfile.ZIP_DEFLATED, 1),
    "normal": (zipfile.ZIP_DEFLATED, 6),
    "maximum": (zipfile.ZIP_DEFLATED, 9),
}
DEFAULT_COMPRESSION = "normal"

# Skipped when zipping a folder -- these are noise in every archive.
SKIP_NAMES = {".DS_Store", "Thumbs.db", "desktop.ini", "__pycache__", ".git"}


@dataclass()
class ArchiveEntry:
    """One member of an archive, as shown in the contents preview."""

    name: str
    size: int
    compressed_size: int
    is_dir: bool
    modified: str

    @property
    def ratio(self) -> float:
        if self.size <= 0:
            return 0.0
        return 1.0 - (self.compressed_size / self.size)


@dataclass()
class ArchiveSummary:
    path: Path
    entries: list[ArchiveEntry] = field(default_factory=list)
    total_size: int = 0
    total_compressed: int = 0

    @property
    def file_count(self) -> int:
        return sum(1 for entry in self.entries if not entry.is_dir)

    @property
    def ratio(self) -> float:
        if self.total_size <= 0:
            return 0.0
        return 1.0 - (self.total_compressed / self.total_size)


def _log(log: LogFn | None, message: str) -> None:
    if log:
        log(message)


def _progress(progress: ProgressFn | None, current: int, total: int) -> None:
    if progress:
        progress(current, total)


def format_size(num_bytes: int | float) -> str:
    """Human-readable byte size, e.g. '1.4 MB'."""
    size = float(num_bytes or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def _compression(level_name: str) -> tuple[int, int | None]:
    key = (level_name or DEFAULT_COMPRESSION).strip().lower()
    if key not in COMPRESSION_LEVELS:
        raise ConversionError(
            f"Unknown compression level '{level_name}'. "
            f"Choose one of: {', '.join(sorted(COMPRESSION_LEVELS))}."
        )
    return COMPRESSION_LEVELS[key]


def _open_zip_for_write(path: Path, level_name: str) -> zipfile.ZipFile:
    compression, compresslevel = _compression(level_name)
    kwargs: dict[str, object] = {"mode": "w", "compression": compression}
    if compresslevel is not None:
        kwargs["compresslevel"] = compresslevel
    return zipfile.ZipFile(path, **kwargs)  # type: ignore[arg-type]


def _should_skip(path: Path) -> bool:
    return any(part in SKIP_NAMES for part in path.parts)


def _collect_folder_files(folder: Path) -> list[Path]:
    files = [
        candidate
        for candidate in sorted(folder.rglob("*"))
        if candidate.is_file() and not _should_skip(candidate.relative_to(folder))
    ]
    return files


def create_archive(
    sources: list[Path] | list[str],
    destination: Path | str,
    *,
    compression: str = DEFAULT_COMPRESSION,
    overwrite: bool = False,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> Path:
    """Zip ``sources`` (files and/or folders) into ``destination``.

    Folders keep their internal structure under a top-level folder of the same
    name. Returns the archive path actually written.
    """
    resolved_sources = [Path(item).expanduser() for item in sources]
    if not resolved_sources:
        raise ConversionError("Select at least one file or folder to compress.")

    missing = [item for item in resolved_sources if not item.exists()]
    if missing:
        names = ", ".join(item.name for item in missing[:3])
        raise ConversionError(f"These inputs no longer exist: {names}. Re-add them and try again.")

    target = Path(destination).expanduser()
    if target.suffix.lower() != ARCHIVE_EXTENSION:
        target = target.with_suffix(ARCHIVE_EXTENSION)
    ensure_directory(target.parent)
    if target.exists() and not overwrite:
        target = unique_path(target)

    # (path on disk, name inside the archive)
    members: list[tuple[Path, str]] = []
    for source in resolved_sources:
        if source.is_dir():
            for item in _collect_folder_files(source):
                members.append((item, str(Path(source.name) / item.relative_to(source))))
        else:
            members.append((source, source.name))

    if not members:
        raise ConversionError("Nothing to compress -- the selected folders contain no files.")

    total = len(members)
    _log(log, f"Compressing {total} file(s) into {target.name} ({compression}).")
    _progress(progress, 0, total)

    try:
        with _open_zip_for_write(target, compression) as archive:
            for index, (disk_path, arc_name) in enumerate(members, start=1):
                archive.write(disk_path, arc_name)
                _progress(progress, index, total)
                if index % 25 == 0 or index == total:
                    _log(log, f"  {index}/{total} added")
    except ConversionError:
        raise
    except PermissionError as exc:
        raise ConversionError(
            f"No permission to write {target}. Pick a different output folder."
        ) from exc
    except OSError as exc:
        raise ConversionError(f"Could not write {target.name}: {exc}") from exc

    _log(log, f"Created {target} ({format_size(target.stat().st_size)}).")
    return target


def batch_compress(
    folders: list[Path] | list[str],
    output_dir: Path | str,
    *,
    compression: str = DEFAULT_COMPRESSION,
    overwrite: bool = False,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> list[Path]:
    """Zip each folder in ``folders`` into its own archive under ``output_dir``."""
    resolved = [Path(item).expanduser() for item in folders]
    if not resolved:
        raise ConversionError("Select at least one folder to batch-compress.")

    not_folders = [item for item in resolved if not item.is_dir()]
    if not_folders:
        names = ", ".join(item.name for item in not_folders[:3])
        raise ConversionError(f"Batch compress takes folders, but got: {names}.")

    target_dir = ensure_directory(Path(output_dir).expanduser())
    outputs: list[Path] = []
    total = len(resolved)
    _progress(progress, 0, total)

    for index, folder in enumerate(resolved, start=1):
        _log(log, f"[{index}/{total}] {folder.name}")
        archive_path = target_dir / f"{folder.name}{ARCHIVE_EXTENSION}"
        outputs.append(
            create_archive(
                [folder],
                archive_path,
                compression=compression,
                overwrite=overwrite,
                log=log,
            )
        )
        _progress(progress, index, total)

    _log(log, f"Batch compress finished: {len(outputs)} archive(s).")
    return outputs


def read_archive(path: Path | str) -> ArchiveSummary:
    """List an archive's contents without extracting anything."""
    archive_path = Path(path).expanduser()
    if not archive_path.exists():
        raise ConversionError(f"Archive not found: {archive_path}")
    if not zipfile.is_zipfile(archive_path):
        raise ConversionError(
            f"{archive_path.name} is not a valid .zip archive. "
            "GO Omni Pro reads zip archives only."
        )

    summary = ArchiveSummary(path=archive_path)
    try:
        with zipfile.ZipFile(archive_path) as archive:
            for info in archive.infolist():
                try:
                    modified = datetime(*info.date_time).strftime("%Y-%m-%d %H:%M")
                except (ValueError, TypeError):
                    modified = ""
                summary.entries.append(
                    ArchiveEntry(
                        name=info.filename,
                        size=info.file_size,
                        compressed_size=info.compress_size,
                        is_dir=info.is_dir(),
                        modified=modified,
                    )
                )
                if not info.is_dir():
                    summary.total_size += info.file_size
                    summary.total_compressed += info.compress_size
    except zipfile.BadZipFile as exc:
        raise ConversionError(f"{archive_path.name} is damaged and cannot be read: {exc}") from exc

    return summary


def _safe_destination(destination: Path, member_name: str) -> Path:
    """Resolve an archive member against the destination, refusing path escapes."""
    candidate = (destination / member_name).resolve()
    root = destination.resolve()
    if candidate != root and root not in candidate.parents:
        raise ConversionError(
            f"Refused to extract '{member_name}': it points outside the destination folder. "
            "This archive may be malicious."
        )
    return candidate


def extract_archive(
    archive: Path | str,
    destination: Path | str,
    *,
    members: list[str] | None = None,
    create_subfolder: bool = True,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> Path:
    """Extract ``archive`` into ``destination``, returning the folder written to.

    With ``create_subfolder`` the contents land in a folder named after the
    archive, so extracting never scatters files across an existing directory.
    """
    archive_path = Path(archive).expanduser()
    summary = read_archive(archive_path)

    target = Path(destination).expanduser()
    if create_subfolder:
        target = target / archive_path.stem
    target = ensure_directory(target)

    wanted = set(members) if members else None
    selected = [
        entry for entry in summary.entries if wanted is None or entry.name in wanted
    ]
    if not selected:
        raise ConversionError("No archive entries selected to extract.")

    total = len(selected)
    _log(log, f"Extracting {total} entr(y/ies) from {archive_path.name} into {target}.")
    _progress(progress, 0, total)

    try:
        with zipfile.ZipFile(archive_path) as zf:
            for index, entry in enumerate(selected, start=1):
                out_path = _safe_destination(target, entry.name)
                if entry.is_dir:
                    ensure_directory(out_path)
                else:
                    ensure_directory(out_path.parent)
                    with zf.open(entry.name) as source, open(out_path, "wb") as handle:
                        while True:
                            chunk = source.read(1024 * 256)
                            if not chunk:
                                break
                            handle.write(chunk)
                _progress(progress, index, total)
                if index % 25 == 0 or index == total:
                    _log(log, f"  {index}/{total} extracted")
    except ConversionError:
        raise
    except PermissionError as exc:
        raise ConversionError(
            f"No permission to write into {target}. Pick a different destination folder."
        ) from exc
    except (OSError, zipfile.BadZipFile) as exc:
        raise ConversionError(f"Extract failed: {exc}") from exc

    _log(log, f"Extracted to {target}.")
    return target


def suggest_archive_name(sources: list[Path] | list[str]) -> str:
    """A sensible default archive filename for the current selection."""
    resolved = [Path(item) for item in sources]
    if not resolved:
        return f"archive{ARCHIVE_EXTENSION}"
    if len(resolved) == 1:
        return f"{resolved[0].stem or resolved[0].name}{ARCHIVE_EXTENSION}"
    common = os.path.commonprefix([str(item.parent) for item in resolved])
    label = Path(common).name if common else "archive"
    return f"{label or 'archive'}{ARCHIVE_EXTENSION}"
