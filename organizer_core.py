from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

import fitz  # PyMuPDF
from PIL import Image


class OrganizerError(RuntimeError):
    """Raised when a page organizer action cannot be completed."""


@dataclass(frozen=True)
class OrganizedPage:
    source_index: int
    rotation: int = 0


@dataclass(frozen=True)
class PdfSummary:
    page_count: int
    file_size_bytes: int
    title: str
    author: str
    subject: str
    keywords: str
    producer: str
    encrypted: bool

    @property
    def file_size_label(self) -> str:
        size = float(self.file_size_bytes)
        units = ["B", "KB", "MB", "GB"]
        for unit in units:
            if size < 1024 or unit == units[-1]:
                return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
            size /= 1024
        return f"{self.file_size_bytes} B"


THUMBNAIL_SIZE = (180, 240)
PREVIEW_SIZE = (920, 1200)
EXPORT_IMAGE_SCALE = 2.0


def _normalize_rotation(value: int) -> int:
    return int(value) % 360


def sanitize_positions(positions: Sequence[int], length: int) -> list[int]:
    unique: list[int] = []
    seen: set[int] = set()
    for pos in positions:
        index = int(pos)
        if 0 <= index < length and index not in seen:
            unique.append(index)
            seen.add(index)
    return unique


def build_default_sequence(page_count: int) -> list[OrganizedPage]:
    return [OrganizedPage(source_index=index, rotation=0) for index in range(page_count)]


def pdf_summary(input_pdf: Path) -> PdfSummary:
    pdf_path = Path(input_pdf)
    if not pdf_path.exists():
        raise OrganizerError(f"PDF not found: {pdf_path}")

    with fitz.open(str(pdf_path)) as document:
        metadata = document.metadata or {}
        return PdfSummary(
            page_count=document.page_count,
            file_size_bytes=pdf_path.stat().st_size,
            title=str(metadata.get("title") or "").strip(),
            author=str(metadata.get("author") or "").strip(),
            subject=str(metadata.get("subject") or "").strip(),
            keywords=str(metadata.get("keywords") or "").strip(),
            producer=str(metadata.get("producer") or "").strip(),
            encrypted=bool(getattr(document, "needs_pass", False) or getattr(document, "is_encrypted", False)),
        )


def _render_image_from_document(
    document: fitz.Document,
    source_index: int,
    rotation: int,
    max_size: tuple[int, int],
    base_zoom: float,
) -> Image.Image:
    page = document.load_page(int(source_index))
    matrix = fitz.Matrix(base_zoom, base_zoom)
    pixmap = page.get_pixmap(matrix=matrix, alpha=False)
    image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
    if rotation:
        image = image.rotate(-_normalize_rotation(rotation), expand=True)
    image.thumbnail(max_size, getattr(Image, "Resampling", Image).LANCZOS)
    return image


def render_thumbnail_from_document(document: fitz.Document, source_index: int, rotation: int = 0) -> Image.Image:
    return _render_image_from_document(document, source_index, rotation, THUMBNAIL_SIZE, 0.34)


def render_preview_from_document(document: fitz.Document, source_index: int, rotation: int = 0) -> Image.Image:
    return _render_image_from_document(document, source_index, rotation, PREVIEW_SIZE, 1.25)


def rotate_positions(sequence: Sequence[OrganizedPage], positions: Sequence[int], delta: int) -> list[OrganizedPage]:
    selected = set(sanitize_positions(positions, len(sequence)))
    new_sequence: list[OrganizedPage] = []
    for index, page in enumerate(sequence):
        if index in selected:
            new_sequence.append(OrganizedPage(page.source_index, _normalize_rotation(page.rotation + delta)))
        else:
            new_sequence.append(page)
    return new_sequence


def move_positions_up(sequence: Sequence[OrganizedPage], positions: Sequence[int]) -> tuple[list[OrganizedPage], list[int]]:
    selected = sanitize_positions(positions, len(sequence))
    selected_set = set(selected)
    items = list(sequence)
    for pos in selected:
        if pos > 0 and (pos - 1) not in selected_set:
            items[pos - 1], items[pos] = items[pos], items[pos - 1]
    new_selected = [max(pos - 1, 0) if pos > 0 and (pos - 1) not in selected_set else pos for pos in selected]
    return items, new_selected


def move_positions_down(sequence: Sequence[OrganizedPage], positions: Sequence[int]) -> tuple[list[OrganizedPage], list[int]]:
    selected = sanitize_positions(positions, len(sequence))
    selected_set = set(selected)
    items = list(sequence)
    for pos in reversed(selected):
        if pos < len(items) - 1 and (pos + 1) not in selected_set:
            items[pos + 1], items[pos] = items[pos], items[pos + 1]
    new_selected = [min(pos + 1, len(items) - 1) if pos < len(items) - 1 and (pos + 1) not in selected_set else pos for pos in selected]
    return items, new_selected


def duplicate_positions(sequence: Sequence[OrganizedPage], positions: Sequence[int]) -> tuple[list[OrganizedPage], list[int]]:
    selected = set(sanitize_positions(positions, len(sequence)))
    items: list[OrganizedPage] = []
    duplicates: list[int] = []
    for index, page in enumerate(sequence):
        items.append(page)
        if index in selected:
            items.append(OrganizedPage(page.source_index, page.rotation))
            duplicates.append(len(items) - 1)
    return items, duplicates


def remove_positions(sequence: Sequence[OrganizedPage], positions: Sequence[int]) -> tuple[list[OrganizedPage], list[int]]:
    selected = set(sanitize_positions(positions, len(sequence)))
    items = [page for index, page in enumerate(sequence) if index not in selected]
    return items, []


def reverse_sequence(sequence: Sequence[OrganizedPage], positions: Sequence[int] | None = None) -> tuple[list[OrganizedPage], list[int]]:
    items = list(reversed(sequence))
    selected = sanitize_positions(list(positions or []), len(sequence))
    new_selected = [len(sequence) - 1 - pos for pos in selected]
    new_selected.sort()
    return items, new_selected


def save_sequence_as_pdf(input_pdf: Path, sequence: Sequence[OrganizedPage], output_pdf: Path) -> Path:
    if not sequence:
        raise OrganizerError("The organizer sequence is empty. Add or keep at least one page before saving.")

    source_pdf = Path(input_pdf)
    target_pdf = Path(output_pdf)
    target_pdf.parent.mkdir(parents=True, exist_ok=True)

    source_document = fitz.open(str(source_pdf))
    try:
        if getattr(source_document, "needs_pass", False):
            raise OrganizerError("This PDF is password protected. Unlock it first before using the visual organizer.")

        output_document = fitz.open()
        try:
            for item in sequence:
                output_document.insert_pdf(source_document, from_page=item.source_index, to_page=item.source_index)
                appended_page = output_document[-1]
                if item.rotation:
                    appended_page.set_rotation(_normalize_rotation(item.rotation))

            metadata = source_document.metadata or {}
            clean_metadata = {key: value for key, value in metadata.items() if value}
            if clean_metadata:
                output_document.set_metadata(clean_metadata)
            output_document.save(str(target_pdf))
        finally:
            output_document.close()
    finally:
        source_document.close()
    return target_pdf


def extract_selected_pdf(input_pdf: Path, sequence: Sequence[OrganizedPage], positions: Sequence[int], output_pdf: Path) -> Path:
    selected = sanitize_positions(positions, len(sequence))
    if not selected:
        raise OrganizerError("Select at least one page in the organizer before extracting.")
    subset = [sequence[index] for index in selected]
    return save_sequence_as_pdf(input_pdf, subset, output_pdf)


def export_pages_as_images(
    input_pdf: Path,
    sequence: Sequence[OrganizedPage],
    positions: Sequence[int],
    output_dir: Path,
    image_format: str = "png",
) -> list[Path]:
    selected = sanitize_positions(positions, len(sequence))
    if not selected:
        raise OrganizerError("Select at least one page in the organizer before exporting images.")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    normalized_format = image_format.strip().lower() or "png"
    if normalized_format == "jpg":
        normalized_format = "jpeg"

    outputs: list[Path] = []
    document = fitz.open(str(input_pdf))
    try:
        if getattr(document, "needs_pass", False):
            raise OrganizerError("This PDF is password protected. Unlock it first before exporting organizer pages as images.")

        for order, position in enumerate(selected, start=1):
            item = sequence[position]
            page = document.load_page(item.source_index)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(EXPORT_IMAGE_SCALE, EXPORT_IMAGE_SCALE), alpha=False)
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            if item.rotation:
                image = image.rotate(-_normalize_rotation(item.rotation), expand=True)
            output_path = out_dir / f"page_{order:03d}.{normalized_format}"
            save_format = "JPEG" if normalized_format == "jpeg" else normalized_format.upper()
            image.save(output_path, format=save_format)
            outputs.append(output_path)
    finally:
        document.close()
    return outputs
