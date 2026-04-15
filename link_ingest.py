from __future__ import annotations

import hashlib
import mimetypes
import re
import shutil
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Sequence
from urllib.parse import urlparse, urlsplit, urlunsplit
from urllib.request import Request, urlopen

DEFAULT_TIMEOUT_SECONDS = 25
DEFAULT_MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024  # 150 MB

ProgressFn = Callable[[int, int], None]
StatusFn = Callable[["LinkDownloadResult"], None]
CancelFn = Callable[[], bool]


@dataclass
class LinkDownloadResult:
    url: str
    normalized_url: str
    status: str
    file_path: str = ""
    error: str = ""
    content_type: str = ""
    detail: str = ""
    filename: str = ""
    duplicate_of: str = ""
    attempt: int = 1


def extract_urls(text: str) -> list[str]:
    if not text:
        return []
    matches = re.findall(r"(?i)\bhttps?://[^\s<>'\"`]+", text)
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in matches:
        url = raw.strip().rstrip("),.;")
        normalized = normalize_url(url)
        if normalized and normalized not in seen:
            seen.add(normalized)
            cleaned.append(url)
    return cleaned


def normalize_url(url: str) -> str:
    value = (url or "").strip()
    if not value:
        return ""
    if not re.match(r"(?i)^https?://", value):
        if value.lower().startswith("www."):
            value = "https://" + value
        else:
            return ""
    parsed = urlsplit(value)
    scheme = parsed.scheme.lower()
    if scheme not in {"http", "https"}:
        return ""
    netloc = parsed.netloc.lower()
    path = parsed.path or "/"
    query = parsed.query
    return urlunsplit((scheme, netloc, path, query, ""))


def cache_root_from_setting(value: str | None, app_state_dir: Path) -> Path:
    raw = (value or "").strip()
    if raw:
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = (app_state_dir / raw).resolve()
    else:
        path = (app_state_dir / "link_cache").resolve()
    path.mkdir(parents=True, exist_ok=True)
    return path


def _guess_extension_from_content_type(content_type: str, url: str) -> str:
    content_type = (content_type or "").split(";")[0].strip().lower()
    ext = ""
    if content_type:
        ext = mimetypes.guess_extension(content_type) or ""
    if ext == ".jpe":
        ext = ".jpg"
    if ext:
        return ext
    parsed = urlparse(url)
    suffix = Path(parsed.path).suffix.lower()
    if suffix:
        return suffix
    return ""


def _safe_filename_from_url(url: str, content_type: str = "", fallback_stem: str = "download") -> str:
    parsed = urlparse(url)
    stem = Path(parsed.path).name
    stem = stem.split("?")[0].split("#")[0]
    if not stem or stem in {".", "/"}:
        host = parsed.netloc or fallback_stem
        host = re.sub(r"[^A-Za-z0-9._-]+", "_", host).strip("._") or fallback_stem
        digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]
        ext = _guess_extension_from_content_type(content_type, url) or ".bin"
        return f"{host}_{digest}{ext}"
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")
    if not cleaned:
        cleaned = fallback_stem
    if "." not in cleaned:
        ext = _guess_extension_from_content_type(content_type, url) or ".bin"
        cleaned = cleaned + ext
    return cleaned


def _make_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _status(url: str, normalized_url: str, state: str, **kwargs) -> LinkDownloadResult:
    return LinkDownloadResult(url=url, normalized_url=normalized_url, status=state, **kwargs)


def download_many_urls(
    urls: Sequence[str] | Iterable[str],
    cache_root: Path,
    *,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    max_bytes: int = DEFAULT_MAX_DOWNLOAD_BYTES,
    cancel_requested: CancelFn | None = None,
    pause_requested: CancelFn | None = None,
    status_callback: StatusFn | None = None,
    progress_callback: ProgressFn | None = None,
    keep_downloads: bool = True,
) -> list[LinkDownloadResult]:
    cache_root = Path(cache_root)
    cache_root.mkdir(parents=True, exist_ok=True)

    materialized = [item.strip() for item in urls if str(item).strip()]
    total = len(materialized)
    if total == 0:
        return []

    results: list[LinkDownloadResult] = []
    seen: dict[str, LinkDownloadResult] = {}

    def emit(result: LinkDownloadResult) -> None:
        results.append(result)
        if status_callback:
            status_callback(result)

    def wait_while_paused(original_url: str, normalized_url: str, *, detail: str = "paused by user") -> bool:
        paused_once = False
        while pause_requested and pause_requested():
            if cancel_requested and cancel_requested():
                emit(_status(original_url, normalized_url, "cancelled", detail="cancelled while paused"))
                return False
            if not paused_once:
                emit(_status(original_url, normalized_url, "paused", detail=detail))
                paused_once = True
            time.sleep(0.12)
        if paused_once:
            emit(_status(original_url, normalized_url, "resumed", detail="resumed"))
        return True

    for index, original_url in enumerate(materialized, start=1):
        if progress_callback:
            progress_callback(index - 1, total)

        normalized = normalize_url(original_url)
        if not normalized:
            emit(_status(original_url, "", "invalid", error="Only http:// and https:// links are supported.", detail="invalid URL"))
            continue

        if normalized in seen:
            emit(_status(original_url, normalized, "duplicate", detail="duplicate skipped", duplicate_of=seen[normalized].file_path or seen[normalized].url))
            continue

        if cancel_requested and cancel_requested():
            emit(_status(original_url, normalized, "cancelled", detail="cancelled before download"))
            continue

        if not wait_while_paused(original_url, normalized, detail="paused before request"):
            seen[normalized] = _status(original_url, normalized, "cancelled", detail="cancelled while paused")
            continue

        emit(_status(original_url, normalized, "queued", detail="queued"))
        request = Request(
            normalized,
            headers={
                "User-Agent": "GokulOmniConvertLite/1.3 (+https://example.local)",
                "Accept": "*/*",
            },
        )

        try:
            emit(_status(original_url, normalized, "downloading", detail="starting download"))
            with urlopen(request, timeout=timeout) as response:
                content_type = response.headers.get("Content-Type", "")
                filename = _safe_filename_from_url(normalized, content_type)
                target = _make_unique_path(cache_root / filename)
                bytes_written = 0
                cancelled = False
                with target.open("wb") as handle:
                    while True:
                        if not wait_while_paused(original_url, normalized):
                            cancelled = True
                            handle.close()
                            try:
                                target.unlink(missing_ok=True)  # type: ignore[arg-type]
                            except TypeError:  # pragma: no cover - py<3.8
                                if target.exists():
                                    target.unlink()
                            break
                        chunk = response.read(1024 * 64)
                        if not chunk:
                            break
                        if cancel_requested and cancel_requested():
                            cancelled = True
                            handle.close()
                            try:
                                target.unlink(missing_ok=True)  # type: ignore[arg-type]
                            except TypeError:  # pragma: no cover - py<3.8
                                if target.exists():
                                    target.unlink()
                            emit(_status(original_url, normalized, "cancelled", detail="download cancelled", content_type=content_type, filename=filename))
                            break
                        bytes_written += len(chunk)
                        if bytes_written > max_bytes:
                            handle.close()
                            try:
                                target.unlink(missing_ok=True)  # type: ignore[arg-type]
                            except TypeError:  # pragma: no cover
                                if target.exists():
                                    target.unlink()
                            raise RuntimeError(f"Download exceeded the safety limit of {max_bytes // (1024 * 1024)} MB.")
                        handle.write(chunk)
                if cancelled or (cancel_requested and cancel_requested()):
                    seen[normalized] = _status(original_url, normalized, "cancelled", content_type=content_type, filename=filename)
                    continue

                detail = f"{bytes_written / 1024:.1f} KB"
                result = _status(
                    original_url,
                    normalized,
                    "downloaded",
                    file_path=str(target),
                    content_type=content_type,
                    detail=detail,
                    filename=target.name,
                )
                emit(result)
                seen[normalized] = result
        except Exception as exc:
            emit(_status(original_url, normalized, "failed", error=str(exc), detail="download failed"))
            seen[normalized] = _status(original_url, normalized, "failed", error=str(exc), detail="download failed")

    if progress_callback:
        progress_callback(total, total)

    return results


def clear_cache_dir(cache_root: Path) -> tuple[int, int]:
    cache_root = Path(cache_root)
    if not cache_root.exists():
        return 0, 0
    file_count = 0
    byte_count = 0
    for item in cache_root.rglob("*"):
        if item.is_file():
            try:
                byte_count += item.stat().st_size
            except OSError:
                pass
            file_count += 1
    shutil.rmtree(cache_root, ignore_errors=True)
    cache_root.mkdir(parents=True, exist_ok=True)
    return file_count, byte_count
