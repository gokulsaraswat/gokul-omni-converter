from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from converter_core import ConversionError, IMAGE_EXTS, convert_images_to_single_pdf, ensure_directory, safe_name

IMAGE_FOLDER_SORT_NATURAL = "natural"
IMAGE_FOLDER_SORT_FILENAME = "filename"
IMAGE_FOLDER_SORT_MODIFIED = "modified"
IMAGE_FOLDER_SORT_CREATED = "created"
IMAGE_FOLDER_SORT_OPTIONS = [
    IMAGE_FOLDER_SORT_NATURAL,
    IMAGE_FOLDER_SORT_FILENAME,
    IMAGE_FOLDER_SORT_MODIFIED,
    IMAGE_FOLDER_SORT_CREATED,
]

IMAGE_FOLDER_SCOPE_ALL = "all_images_one_pdf"
IMAGE_FOLDER_SCOPE_PER_FOLDER = "one_pdf_per_folder"
IMAGE_FOLDER_SCOPE_OPTIONS = [IMAGE_FOLDER_SCOPE_ALL, IMAGE_FOLDER_SCOPE_PER_FOLDER]
IMAGE_FOLDER_SCOPE_LABELS = {
    IMAGE_FOLDER_SCOPE_ALL: "All Images -> One PDF",
    IMAGE_FOLDER_SCOPE_PER_FOLDER: "One PDF Per Folder",
}
IMAGE_FOLDER_SORT_LABELS = {
    IMAGE_FOLDER_SORT_NATURAL: "Natural Filename",
    IMAGE_FOLDER_SORT_FILENAME: "Filename A-Z",
    IMAGE_FOLDER_SORT_MODIFIED: "Modified Time",
    IMAGE_FOLDER_SORT_CREATED: "Created Time",
}

ProgressFn = Callable[[int, int], None]
LogFn = Callable[[str], None]


@dataclass(slots=True)
class ImageFolderPdfConfig:
    source_dir: Path
    output_dir: Path
    output_name: str = "images_combined"
    recursive: bool = True
    sort_mode: str = IMAGE_FOLDER_SORT_NATURAL
    scope: str = IMAGE_FOLDER_SCOPE_ALL
    include_hidden: bool = False


def _natural_key(value: Path | str) -> tuple[object, ...]:
    path = Path(value)
    text = str(path.name or value).lower()
    parts = re.split(r"(\d+)", text)
    key: list[object] = []
    for part in parts:
        if part.isdigit():
            key.append(int(part))
        else:
            key.append(part)
    return tuple(key)


def _is_hidden(path: Path, root: Path) -> bool:
    try:
        relative = path.relative_to(root)
    except ValueError:
        relative = path
    return any(part.startswith(".") for part in relative.parts)


def sort_image_paths(paths: Iterable[Path], sort_mode: str = IMAGE_FOLDER_SORT_NATURAL, root: Path | None = None) -> list[Path]:
    mode = (sort_mode or IMAGE_FOLDER_SORT_NATURAL).strip().lower()
    materialized = [Path(path) for path in paths]

    if mode == IMAGE_FOLDER_SORT_FILENAME:
        return sorted(materialized, key=lambda path: str(path).lower())
    if mode == IMAGE_FOLDER_SORT_MODIFIED:
        return sorted(materialized, key=lambda path: (path.stat().st_mtime if path.exists() else 0.0, str(path).lower()))
    if mode == IMAGE_FOLDER_SORT_CREATED:
        return sorted(materialized, key=lambda path: (path.stat().st_ctime if path.exists() else 0.0, str(path).lower()))

    root_path = Path(root).expanduser().resolve() if root is not None else None

    def natural_path_key(path: Path) -> tuple[object, ...]:
        target = Path(path)
        try:
            relative = target.relative_to(root_path) if root_path is not None else target
        except ValueError:
            relative = target
        parts = tuple(_natural_key(part) for part in relative.parts)
        # Root-level files should stay before nested subfolder files.
        return (max(len(relative.parts) - 1, 0), parts)

    return sorted(materialized, key=natural_path_key)


def discover_image_files(
    source_dir: Path,
    *,
    recursive: bool = True,
    sort_mode: str = IMAGE_FOLDER_SORT_NATURAL,
    include_hidden: bool = False,
) -> list[Path]:
    root = Path(source_dir).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        raise ConversionError(f"Image folder not found: {root}")

    iterator = root.rglob("*") if recursive else root.glob("*")
    images: list[Path] = []
    for candidate in iterator:
        if not candidate.is_file():
            continue
        if candidate.suffix.lower() not in IMAGE_EXTS:
            continue
        resolved = candidate.resolve()
        if not include_hidden and _is_hidden(resolved, root):
            continue
        images.append(resolved)
    return sort_image_paths(images, sort_mode=sort_mode, root=root)


def group_images_by_folder(images: Iterable[Path], source_dir: Path, *, sort_mode: str = IMAGE_FOLDER_SORT_NATURAL) -> dict[Path, list[Path]]:
    root = Path(source_dir).expanduser().resolve()
    groups: dict[Path, list[Path]] = {}
    for image_path in images:
        parent = Path(image_path).expanduser().resolve().parent
        try:
            relative_parent = parent.relative_to(root)
        except ValueError:
            relative_parent = Path(".")
        groups.setdefault(relative_parent, []).append(Path(image_path))

    ordered: dict[Path, list[Path]] = {}
    for relative_parent in sorted(groups, key=lambda item: str(item).lower()):
        ordered[relative_parent] = sort_image_paths(groups[relative_parent], sort_mode=sort_mode)
    return ordered


def summarize_image_folder(
    source_dir: Path,
    *,
    recursive: bool = True,
    sort_mode: str = IMAGE_FOLDER_SORT_NATURAL,
    include_hidden: bool = False,
) -> dict[str, object]:
    images = discover_image_files(source_dir, recursive=recursive, sort_mode=sort_mode, include_hidden=include_hidden)
    groups = group_images_by_folder(images, source_dir, sort_mode=sort_mode)
    return {
        "source_dir": str(Path(source_dir).expanduser().resolve()),
        "image_count": len(images),
        "folder_count": len(groups),
        "recursive": bool(recursive),
        "sort_mode": sort_mode,
        "first_image": str(images[0]) if images else "",
        "last_image": str(images[-1]) if images else "",
    }


def _emit_log(log: LogFn | None, message: str) -> None:
    if log:
        log(message)


def _emit_progress(progress: ProgressFn | None, current: int, total: int) -> None:
    if progress:
        progress(int(current), max(int(total), 1))


def build_image_folder_pdfs(
    config: ImageFolderPdfConfig,
    *,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> list[Path]:
    source = Path(config.source_dir).expanduser().resolve()
    output_dir = ensure_directory(Path(config.output_dir).expanduser())
    images = discover_image_files(
        source,
        recursive=config.recursive,
        sort_mode=config.sort_mode,
        include_hidden=config.include_hidden,
    )
    if not images:
        raise ConversionError("No image files were found in the selected folder.")

    scope = (config.scope or IMAGE_FOLDER_SCOPE_ALL).strip().lower()
    output_name = safe_name(config.output_name or source.name or "images_combined")

    if scope == IMAGE_FOLDER_SCOPE_PER_FOLDER:
        groups = group_images_by_folder(images, source, sort_mode=config.sort_mode)
        outputs: list[Path] = []
        total_groups = max(len(groups), 1)
        for index, (relative_folder, folder_images) in enumerate(groups.items(), start=1):
            if not folder_images:
                continue
            if str(relative_folder) in {".", ""}:
                stem = output_name
            else:
                stem = safe_name(str(relative_folder).replace("/", "_").replace("\\", "_"))
            destination = output_dir / f"{stem}.pdf"
            _emit_log(log, f"Building PDF {index}/{total_groups}: {stem} from {len(folder_images)} image(s)")
            outputs.append(convert_images_to_single_pdf(folder_images, destination))
            _emit_progress(progress, index, total_groups)
        return outputs

    _emit_log(log, f"Building one PDF from {len(images)} image(s).")
    output_pdf = output_dir / f"{output_name}.pdf"
    result = convert_images_to_single_pdf(images, output_pdf)
    _emit_progress(progress, 1, 1)
    return [result]
