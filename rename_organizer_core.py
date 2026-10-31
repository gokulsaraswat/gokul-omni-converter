"""Batch rename and folder organisation for GO Omni Pro.

Two operations, both preview-first and both undoable:

* :func:`plan_rename` builds an old -> new mapping from a rule set without
  touching disk, so the UI can show exactly what will happen (including
  collisions) before the user commits.
* :func:`plan_organize` does the same for sorting a folder into subfolders.

:func:`apply_plan` executes a plan and writes a JSON manifest of every move to
the app state directory. :func:`undo_last` reverses the most recent manifest,
which is what makes these destructive operations safe to offer at all.

Conventions match converter_core: ConversionError on failure, optional log and
progress callbacks.
"""

from __future__ import annotations

import json
import re
import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from converter_core import ConversionError, LogFn, ProgressFn, ensure_directory

MANIFEST_SUFFIX = ".rename-manifest.json"
MAX_MANIFESTS = 20

# --- rename rules ------------------------------------------------------------

RENAME_MODES = ("pattern", "find_replace", "regex", "case")
CASE_MODES = ("lower", "upper", "title", "kebab", "snake")

SORT_MODES = ("name", "modified", "created", "size", "extension")

#: Tokens available in a pattern. Kept small and self-explanatory on purpose --
#: this is the string a non-technical user actually has to type.
PATTERN_TOKENS = {
    "{n}": "Sequence number, zero-padded (0001, 0002, ...)",
    "{name}": "Original filename without its extension",
    "{ext}": "Original extension, without the dot",
    "{date}": "File's modified date (YYYY-MM-DD)",
    "{datetime}": "File's modified date and time (YYYY-MM-DD_HHMM)",
    "{parent}": "Name of the containing folder",
}

DEFAULT_PATTERN = "{name}"

# --- organize rules ----------------------------------------------------------

ORGANIZE_MODES = ("type", "extension", "date_modified", "date_created")

#: Extension -> friendly bucket, used by the "type" organise mode.
TYPE_BUCKETS = {
    "Documents": {".pdf", ".doc", ".docx", ".odt", ".rtf", ".txt", ".md", ".tex", ".epub"},
    "Spreadsheets": {".xls", ".xlsx", ".xlsm", ".ods", ".csv", ".tsv"},
    "Presentations": {".ppt", ".pptx", ".odp", ".key"},
    "Images": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".svg", ".ico"},
    "Audio": {".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg"},
    "Video": {".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv"},
    "Archives": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"},
    "Code": {".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".go", ".rs", ".rb", ".sh", ".html", ".css", ".json", ".xml", ".yml", ".yaml"},
}
OTHER_BUCKET = "Other"

# Characters no mainstream filesystem accepts, plus the ones Windows rejects.
_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
# Reserved device names on Windows; renaming a file to one of these breaks it there.
_WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


@dataclass()
class PlannedMove:
    """A single old -> new move, plus why it might not be applicable."""

    source: Path
    target: Path
    problem: str = ""

    @property
    def changed(self) -> bool:
        return self.source != self.target

    @property
    def ok(self) -> bool:
        return not self.problem and self.changed


@dataclass()
class OperationPlan:
    """The full result of a dry run: what would move, and what is wrong with it."""

    kind: str  # "rename" or "organize"
    root: Path
    moves: list[PlannedMove] = field(default_factory=list)
    description: str = ""

    @property
    def applicable(self) -> list[PlannedMove]:
        return [move for move in self.moves if move.ok]

    @property
    def problems(self) -> list[PlannedMove]:
        return [move for move in self.moves if move.problem]

    @property
    def unchanged(self) -> list[PlannedMove]:
        return [move for move in self.moves if not move.changed and not move.problem]


def _log(log: LogFn | None, message: str) -> None:
    if log:
        log(message)


def _progress(progress: ProgressFn | None, current: int, total: int) -> None:
    if progress:
        progress(current, total)


def manifest_dir(state_dir: Path | str) -> Path:
    return ensure_directory(Path(state_dir).expanduser() / "rename_manifests")


# --- name generation ---------------------------------------------------------


def sanitize_name(name: str) -> str:
    """Strip characters that would make the name invalid on Windows/macOS/Linux."""
    cleaned = _ILLEGAL_CHARS.sub("_", str(name or "")).strip().rstrip(".")
    return cleaned


def validate_name(name: str) -> str:
    """Return a problem description for ``name``, or '' when it is usable."""
    if not name:
        return "Resulting name is empty"
    if len(name) > 255:
        return "Resulting name is longer than 255 characters"
    if _ILLEGAL_CHARS.search(name):
        return "Contains characters that are not allowed in filenames"
    if Path(name).stem.upper() in _WINDOWS_RESERVED:
        return f"'{Path(name).stem}' is a reserved name on Windows"
    return ""


def sort_files(files: list[Path], mode: str = "name", reverse: bool = False) -> list[Path]:
    """Order the input list, which decides what {n} counts up in."""
    mode = (mode or "name").strip().lower()

    def key(path: Path):
        try:
            stat = path.stat()
        except OSError:
            stat = None
        if mode == "modified":
            return stat.st_mtime if stat else 0
        if mode == "created":
            # st_birthtime exists on macOS/BSD; st_ctime is the closest elsewhere.
            return getattr(stat, "st_birthtime", stat.st_ctime) if stat else 0
        if mode == "size":
            return stat.st_size if stat else 0
        if mode == "extension":
            return (path.suffix.lower(), path.name.lower())
        return path.name.lower()

    return sorted(files, key=key, reverse=reverse)


def _file_dates(path: Path) -> tuple[str, str]:
    try:
        stamp = datetime.fromtimestamp(path.stat().st_mtime)
    except OSError:
        stamp = datetime.now()
    return stamp.strftime("%Y-%m-%d"), stamp.strftime("%Y-%m-%d_%H%M")


def apply_case(text: str, case_mode: str) -> str:
    mode = (case_mode or "lower").strip().lower()
    if mode == "upper":
        return text.upper()
    if mode == "title":
        return text.title()
    if mode == "kebab":
        return re.sub(r"[\s_]+", "-", text.strip()).lower()
    if mode == "snake":
        return re.sub(r"[\s-]+", "_", text.strip()).lower()
    return text.lower()


def render_pattern(pattern: str, path: Path, index: int, *, pad: int = 4, start: int = 1) -> str:
    """Expand PATTERN_TOKENS for one file. Unknown {tokens} are left as typed."""
    date_str, datetime_str = _file_dates(path)
    number = str(start + index).zfill(max(1, int(pad or 1)))
    replacements = {
        "{n}": number,
        "{name}": path.stem,
        "{ext}": path.suffix.lstrip("."),
        "{date}": date_str,
        "{datetime}": datetime_str,
        "{parent}": path.parent.name,
    }
    result = str(pattern or DEFAULT_PATTERN)
    for token, value in replacements.items():
        result = result.replace(token, value)
    return result


def plan_rename(
    files: list[Path] | list[str],
    *,
    mode: str = "pattern",
    pattern: str = DEFAULT_PATTERN,
    find_text: str = "",
    replace_text: str = "",
    regex_pattern: str = "",
    regex_replacement: str = "",
    case_mode: str = "lower",
    keep_extension: bool = True,
    sort_mode: str = "name",
    sort_reverse: bool = False,
    start_number: int = 1,
    number_padding: int = 4,
) -> OperationPlan:
    """Dry-run a rename. Nothing on disk is touched."""
    resolved = [Path(item).expanduser() for item in files]
    resolved = [item for item in resolved if item.is_file()]
    if not resolved:
        raise ConversionError("Add some files first -- there is nothing to rename.")

    mode = (mode or "pattern").strip().lower()
    if mode not in RENAME_MODES:
        raise ConversionError(f"Unknown rename mode '{mode}'.")
    if mode == "regex" and not regex_pattern:
        raise ConversionError("Enter a regular expression to match before renaming.")
    if mode == "find_replace" and not find_text:
        raise ConversionError("Enter the text to find before renaming.")

    compiled: re.Pattern[str] | None = None
    if mode == "regex":
        try:
            compiled = re.compile(regex_pattern)
        except re.error as exc:
            raise ConversionError(f"That regular expression is not valid: {exc}") from exc

    ordered = sort_files(resolved, sort_mode, sort_reverse)
    root = ordered[0].parent
    plan = OperationPlan(kind="rename", root=root, description=f"Rename ({mode})")

    # Names already claimed by this plan, so two files can't be renamed onto one.
    claimed: dict[str, Path] = {}
    sources = {str(item.resolve()).lower() for item in ordered}

    for index, path in enumerate(ordered):
        stem = path.stem
        suffix = path.suffix

        if mode == "pattern":
            new_stem = render_pattern(pattern, path, index, pad=number_padding, start=start_number)
        elif mode == "find_replace":
            new_stem = stem.replace(find_text, replace_text)
        elif mode == "regex":
            assert compiled is not None
            try:
                new_stem = compiled.sub(regex_replacement, stem)
            except re.error as exc:
                raise ConversionError(f"Replacement failed for {path.name}: {exc}") from exc
        else:  # case
            new_stem = apply_case(stem, case_mode)

        if keep_extension:
            new_name = f"{new_stem}{suffix}"
        else:
            new_name = new_stem or path.name

        target = path.with_name(new_name)
        problem = validate_name(new_name)

        if not problem and target != path:
            key = str(target.resolve()).lower()
            if key in claimed:
                problem = f"Would collide with {claimed[key].name}"
            elif target.exists() and key not in sources:
                problem = "A file with that name already exists"
            else:
                claimed[key] = path

        plan.moves.append(PlannedMove(source=path, target=target, problem=problem))

    return plan


def _bucket_for(path: Path, mode: str) -> str:
    suffix = path.suffix.lower()
    if mode == "extension":
        return suffix.lstrip(".").upper() or OTHER_BUCKET
    if mode in {"date_modified", "date_created"}:
        try:
            stat = path.stat()
            raw = stat.st_mtime if mode == "date_modified" else getattr(stat, "st_birthtime", stat.st_ctime)
            return datetime.fromtimestamp(raw).strftime("%Y-%m")
        except OSError:
            return "Unknown-date"
    for bucket, extensions in TYPE_BUCKETS.items():
        if suffix in extensions:
            return bucket
    return OTHER_BUCKET


def plan_organize(
    folder: Path | str,
    *,
    mode: str = "type",
    recursive: bool = False,
) -> OperationPlan:
    """Dry-run sorting ``folder``'s files into subfolders. Nothing is moved."""
    root = Path(folder).expanduser()
    if not root.is_dir():
        raise ConversionError(f"Not a folder: {root}. Choose the folder you want to organise.")

    mode = (mode or "type").strip().lower()
    if mode not in ORGANIZE_MODES:
        raise ConversionError(f"Unknown organise mode '{mode}'.")

    candidates = sorted(root.rglob("*") if recursive else root.glob("*"))
    files = [item for item in candidates if item.is_file() and not item.name.startswith(".")]
    if not files:
        raise ConversionError(f"{root.name} has no files to organise.")

    plan = OperationPlan(kind="organize", root=root, description=f"Organise by {mode.replace('_', ' ')}")
    existing_buckets = {item.name for item in root.iterdir() if item.is_dir()}

    for path in files:
        bucket = _bucket_for(path, mode)
        # Don't re-file things already sitting in one of our buckets.
        if path.parent != root and path.parent.name in existing_buckets:
            plan.moves.append(PlannedMove(source=path, target=path))
            continue
        target = root / bucket / path.name
        problem = "" if not target.exists() else "A file with that name is already in the target folder"
        plan.moves.append(PlannedMove(source=path, target=target, problem=problem))

    return plan


# --- execution & undo --------------------------------------------------------


def apply_plan(
    plan: OperationPlan,
    state_dir: Path | str,
    *,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> tuple[list[PlannedMove], Path]:
    """Execute ``plan``, writing an undo manifest. Returns (applied, manifest path).

    Moves are done one at a time and recorded as they succeed, so a failure
    part-way through still leaves a manifest that can undo what did happen.
    """
    applicable = plan.applicable
    if not applicable:
        raise ConversionError(
            "Nothing to apply -- every file would keep its current name, "
            "or every change has a problem listed in the preview."
        )

    manifest_path = manifest_dir(state_dir) / f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}{MANIFEST_SUFFIX}"
    records: list[dict[str, str]] = []
    applied: list[PlannedMove] = []
    total = len(applicable)
    _progress(progress, 0, total)
    _log(log, f"{plan.description}: applying {total} change(s).")

    def write_manifest() -> None:
        manifest_path.write_text(
            json.dumps(
                {
                    "kind": plan.kind,
                    "description": plan.description,
                    "root": str(plan.root),
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "moves": records,
                },
                indent=2,
            ),
            encoding="utf-8",
        )

    try:
        for index, move in enumerate(applicable, start=1):
            ensure_directory(move.target.parent)
            shutil.move(str(move.source), str(move.target))
            records.append({"from": str(move.source), "to": str(move.target)})
            applied.append(move)
            _progress(progress, index, total)
            if index % 25 == 0 or index == total:
                _log(log, f"  {index}/{total} done")
    except (OSError, shutil.Error) as exc:
        write_manifest()
        raise ConversionError(
            f"Stopped after {len(applied)} of {total} change(s): {exc}. "
            "Use Undo last operation to roll back what was already applied."
        ) from exc

    write_manifest()
    _prune_manifests(manifest_dir(state_dir))
    _log(log, f"Done. Undo manifest: {manifest_path.name}")
    return applied, manifest_path


def _prune_manifests(directory: Path) -> None:
    manifests = sorted(directory.glob(f"*{MANIFEST_SUFFIX}"))
    for stale in manifests[:-MAX_MANIFESTS]:
        try:
            stale.unlink()
        except OSError:
            continue


def latest_manifest(state_dir: Path | str) -> Path | None:
    manifests = sorted(manifest_dir(state_dir).glob(f"*{MANIFEST_SUFFIX}"))
    return manifests[-1] if manifests else None


def describe_manifest(path: Path | str) -> str:
    """One-line summary of a manifest, for the 'Undo last operation' label."""
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return "Unreadable undo manifest"
    count = len(payload.get("moves", []))
    when = str(payload.get("created_at", "")).replace("T", " ")
    return f"{payload.get('description', 'Operation')} - {count} file(s) - {when}"


def undo_last(
    state_dir: Path | str,
    *,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> int:
    """Reverse the most recent applied operation. Returns the number restored."""
    manifest_path = latest_manifest(state_dir)
    if manifest_path is None:
        raise ConversionError("There is no recent rename or organise operation to undo.")

    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ConversionError(f"Could not read the undo manifest: {exc}") from exc

    moves = list(reversed(payload.get("moves", [])))
    if not moves:
        manifest_path.unlink(missing_ok=True)
        raise ConversionError("That operation recorded no changes, so there is nothing to undo.")

    total = len(moves)
    restored = 0
    skipped: list[str] = []
    _progress(progress, 0, total)
    _log(log, f"Undoing {payload.get('description', 'operation')} ({total} file(s)).")

    for index, record in enumerate(moves, start=1):
        source = Path(record.get("to", ""))
        target = Path(record.get("from", ""))
        if not source.exists():
            skipped.append(f"{source.name} (no longer there)")
        elif target.exists():
            skipped.append(f"{target.name} (something is back in its place)")
        else:
            try:
                ensure_directory(target.parent)
                shutil.move(str(source), str(target))
                restored += 1
            except (OSError, shutil.Error) as exc:
                skipped.append(f"{source.name} ({exc})")
        _progress(progress, index, total)

    manifest_path.unlink(missing_ok=True)
    _log(log, f"Restored {restored} of {total} file(s).")
    if skipped:
        _log(log, f"Skipped {len(skipped)}: {'; '.join(skipped[:5])}")
    if restored == 0:
        raise ConversionError(
            "Nothing could be restored -- the files have been moved or changed since. "
            "No further undo is available."
        )
    return restored
