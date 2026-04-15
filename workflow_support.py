from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path


@dataclass(slots=True)
class DirectoryStats:
    path: str
    file_count: int = 0
    total_bytes: int = 0
    oldest_mtime: float = 0.0
    newest_mtime: float = 0.0

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def format_bytes(value: int | float) -> str:
    size = float(max(value, 0))
    units = ["B", "KB", "MB", "GB", "TB"]
    index = 0
    while size >= 1024 and index < len(units) - 1:
        size /= 1024
        index += 1
    if index == 0:
        return f"{int(size)} {units[index]}"
    return f"{size:.1f} {units[index]}"


def directory_stats(path: Path) -> DirectoryStats:
    root = Path(path).expanduser()
    stats = DirectoryStats(path=str(root))
    if not root.exists():
        return stats
    for item in root.rglob("*"):
        if not item.is_file():
            continue
        try:
            file_stat = item.stat()
        except OSError:
            continue
        stats.file_count += 1
        stats.total_bytes += int(file_stat.st_size)
        mtime = float(file_stat.st_mtime)
        if not stats.oldest_mtime or mtime < stats.oldest_mtime:
            stats.oldest_mtime = mtime
        if mtime > stats.newest_mtime:
            stats.newest_mtime = mtime
    return stats


def summarize_directory(path: Path) -> str:
    stats = directory_stats(path)
    if stats.file_count <= 0:
        return "Cache is empty."
    newest = datetime.fromtimestamp(stats.newest_mtime).strftime("%Y-%m-%d %H:%M") if stats.newest_mtime else "n/a"
    oldest = datetime.fromtimestamp(stats.oldest_mtime).strftime("%Y-%m-%d %H:%M") if stats.oldest_mtime else "n/a"
    return (
        f"{stats.file_count} file(s) • {format_bytes(stats.total_bytes)} total • "
        f"newest {newest} • oldest {oldest}"
    )


def prune_directory(path: Path, *, max_age_days: int = 0, max_total_bytes: int = 0) -> dict[str, object]:
    root = Path(path).expanduser()
    root.mkdir(parents=True, exist_ok=True)

    removed_files: list[str] = []
    removed_bytes = 0
    now = datetime.now()
    cutoff = None
    if max_age_days > 0:
        cutoff = now - timedelta(days=max_age_days)

    files: list[tuple[float, int, Path]] = []
    for item in root.rglob("*"):
        if not item.is_file():
            continue
        try:
            info = item.stat()
        except OSError:
            continue
        mtime = float(info.st_mtime)
        size = int(info.st_size)
        files.append((mtime, size, item))

    # Drop files older than the max age first.
    if cutoff is not None:
        for mtime, size, file_path in list(files):
            if datetime.fromtimestamp(mtime) < cutoff:
                try:
                    file_path.unlink()
                    removed_files.append(str(file_path))
                    removed_bytes += size
                    files.remove((mtime, size, file_path))
                except OSError:
                    continue

    # Then enforce a total-size ceiling by deleting oldest files first.
    if max_total_bytes > 0:
        total_bytes = sum(size for _, size, _ in files)
        if total_bytes > max_total_bytes:
            for mtime, size, file_path in sorted(files, key=lambda item: item[0]):
                if total_bytes <= max_total_bytes:
                    break
                try:
                    file_path.unlink()
                    removed_files.append(str(file_path))
                    removed_bytes += size
                    total_bytes -= size
                except OSError:
                    continue

    return {
        "removed_count": len(removed_files),
        "removed_bytes": removed_bytes,
        "remaining": directory_stats(root).to_dict(),
    }
