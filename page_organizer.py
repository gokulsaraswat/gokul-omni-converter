from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import fitz  # PyMuPDF
from PIL import ImageTk

from organizer_core import (
    OrganizedPage,
    OrganizerError,
    build_default_sequence,
    duplicate_positions,
    export_pages_as_images,
    extract_selected_pdf,
    move_positions_down,
    move_positions_up,
    pdf_summary,
    remove_positions,
    render_preview_from_document,
    render_thumbnail_from_document,
    reverse_sequence,
    rotate_positions,
    save_sequence_as_pdf,
)
from ui_theme import ThemePalette


RecentJobCallback = Callable[[dict[str, object]], None]
StatusCallback = Callable[[str], None]
PathCallback = Callable[[Path], None]


class PreviewWindow(tk.Toplevel):
    def __init__(self, master: tk.Misc, palette: ThemePalette, title: str, image: ImageTk.PhotoImage, details: str) -> None:
        super().__init__(master)
        self.palette = palette
        self.preview_photo = image
        self.title(title)
        self.geometry("900x960")
        self.minsize(620, 520)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self, bg=palette.surface, highlightthickness=1, highlightbackground=palette.border)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        tk.Label(
            header,
            text=title,
            bg=palette.surface,
            fg=palette.text,
            font=("Segoe UI", 14, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 4))
        tk.Label(
            header,
            text=details,
            bg=palette.surface,
            fg=palette.text_muted,
            justify="left",
            anchor="w",
            wraplength=820,
        ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        container = tk.Frame(self, bg=palette.root_bg)
        container.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(container, bg=palette.input_bg, highlightthickness=1, highlightbackground=palette.border)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.canvas.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.inner = tk.Frame(self.canvas, bg=palette.input_bg)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.image_label = tk.Label(self.inner, image=self.preview_photo, bg=palette.input_bg)
        self.image_label.grid(row=0, column=0, padx=20, pady=20)

    def _sync_scroll_region(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfigure(self.window_id, width=max(event.width - 2, 100))

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        self.configure(bg=palette.root_bg)
        self.canvas.configure(bg=palette.input_bg, highlightbackground=palette.border)
        self.inner.configure(bg=palette.input_bg)
        self.image_label.configure(bg=palette.input_bg)


class PageOrganizerPanel(ttk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        palette: ThemePalette,
        set_status: StatusCallback,
        on_recent_job: RecentJobCallback | None = None,
        on_loaded_pdf: PathCallback | None = None,
    ) -> None:
        super().__init__(master, style="Surface.TFrame", padding=0)
        self.palette = palette
        self.set_status = set_status
        self.on_recent_job = on_recent_job
        self.on_loaded_pdf = on_loaded_pdf

        self.source_pdf: Path | None = None
        self.document: fitz.Document | None = None
        self.page_sequence: list[OrganizedPage] = []
        self.selected_positions: list[int] = []
        self.thumbnail_cache: dict[tuple[int, int], ImageTk.PhotoImage] = {}
        self.card_columns = 4
        self.preview_window: PreviewWindow | None = None

        self.file_var = tk.StringVar(value="No PDF loaded yet.")
        self.summary_var = tk.StringVar(value="Load a PDF to inspect metadata, page count, and live page thumbnails.")
        self.selection_var = tk.StringVar(value="Selection: 0 page(s)")
        self.hint_var = tk.StringVar(
            value="Click thumbnails to toggle selection. Then move, rotate, duplicate, remove, extract, or save the organized PDF."
        )

        self._build_ui()
        self.apply_theme(palette)
        self._update_button_states()

    def destroy(self) -> None:
        self._close_document()
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.destroy()
        super().destroy()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        hero = ttk.Frame(self, style="Card.TFrame", padding=22)
        hero.grid(row=0, column=0, sticky="ew")
        hero.grid_columnconfigure(0, weight=1)
        ttk.Label(hero, text="Visual page organizer", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            hero,
            text=(
                "Patch 8 keeps the thumbnail-driven PDF organizer so you can inspect a PDF visually, reorder pages, rotate pages, "
                "duplicate pages, remove pages, extract a subset, and export selected pages as images without leaving the app."
            ),
            style="HeroBody.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        hero_actions = ttk.Frame(hero, style="Surface.TFrame")
        hero_actions.grid(row=0, column=1, rowspan=2, sticky="e")
        self.open_button = ttk.Button(hero_actions, text="Open PDF", command=self.open_pdf_dialog)
        self.open_button.grid(row=0, column=0, padx=(0, 8))
        self.reload_button = ttk.Button(hero_actions, text="Reload", command=self.reload_pdf)
        self.reload_button.grid(row=0, column=1, padx=(0, 8))
        self.save_button = ttk.Button(hero_actions, text="Save organized PDF", style="Primary.TButton", command=self.save_organized_pdf)
        self.save_button.grid(row=0, column=2)

        info = ttk.Frame(self, style="Card.TFrame", padding=18)
        info.grid(row=1, column=0, sticky="ew", pady=(14, 14))
        info.grid_columnconfigure(0, weight=1)
        ttk.Label(info, textvariable=self.file_var, style="CardTitle.TLabel", wraplength=1080, justify="left").grid(row=0, column=0, sticky="w")
        ttk.Label(info, textvariable=self.summary_var, style="CardBody.TLabel", wraplength=1080, justify="left").grid(row=1, column=0, sticky="w", pady=(8, 0))
        info_footer = ttk.Frame(info, style="Surface.TFrame")
        info_footer.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        info_footer.grid_columnconfigure(0, weight=1)
        ttk.Label(info_footer, textvariable=self.selection_var, style="CardBody.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(info_footer, textvariable=self.hint_var, style="CardBody.TLabel", wraplength=780, justify="left").grid(row=0, column=1, sticky="e")

        workspace = ttk.Frame(self, style="Surface.TFrame")
        workspace.grid(row=2, column=0, sticky="nsew")
        workspace.grid_columnconfigure(0, weight=1)
        workspace.grid_rowconfigure(1, weight=1)

        toolbar = ttk.Frame(workspace, style="Card.TFrame", padding=12)
        toolbar.grid(row=0, column=0, sticky="ew")
        for column in range(11):
            toolbar.grid_columnconfigure(column, weight=0)
        ttk.Button(toolbar, text="Select all", command=self.select_all).grid(row=0, column=0, padx=(0, 8), pady=4)
        ttk.Button(toolbar, text="Clear", command=self.clear_selection).grid(row=0, column=1, padx=(0, 8), pady=4)
        self.move_up_button = ttk.Button(toolbar, text="Move up", command=self.move_selected_up)
        self.move_up_button.grid(row=0, column=2, padx=(0, 8), pady=4)
        self.move_down_button = ttk.Button(toolbar, text="Move down", command=self.move_selected_down)
        self.move_down_button.grid(row=0, column=3, padx=(0, 8), pady=4)
        self.rotate_left_button = ttk.Button(toolbar, text="Rotate left", command=lambda: self.rotate_selected(-90))
        self.rotate_left_button.grid(row=0, column=4, padx=(0, 8), pady=4)
        self.rotate_right_button = ttk.Button(toolbar, text="Rotate right", command=lambda: self.rotate_selected(90))
        self.rotate_right_button.grid(row=0, column=5, padx=(0, 8), pady=4)
        self.duplicate_button = ttk.Button(toolbar, text="Duplicate", command=self.duplicate_selected)
        self.duplicate_button.grid(row=0, column=6, padx=(0, 8), pady=4)
        self.remove_button = ttk.Button(toolbar, text="Remove", command=self.remove_selected)
        self.remove_button.grid(row=0, column=7, padx=(0, 8), pady=4)
        self.reverse_button = ttk.Button(toolbar, text="Reverse order", command=self.reverse_pages)
        self.reverse_button.grid(row=0, column=8, padx=(0, 8), pady=4)
        self.extract_button = ttk.Button(toolbar, text="Extract selected PDF", command=self.extract_selected_pages)
        self.extract_button.grid(row=0, column=9, padx=(0, 8), pady=4)
        self.export_button = ttk.Button(toolbar, text="Export selected images", command=self.export_selected_images)
        self.export_button.grid(row=0, column=10, pady=4)

        thumb_shell = ttk.Frame(workspace, style="Card.TFrame", padding=0)
        thumb_shell.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        thumb_shell.grid_columnconfigure(0, weight=1)
        thumb_shell.grid_rowconfigure(0, weight=1)

        self.thumb_canvas = tk.Canvas(thumb_shell, highlightthickness=0, borderwidth=0)
        self.thumb_canvas.grid(row=0, column=0, sticky="nsew")
        self.thumb_scroll = ttk.Scrollbar(thumb_shell, orient="vertical", command=self.thumb_canvas.yview)
        self.thumb_scroll.grid(row=0, column=1, sticky="ns")
        self.thumb_canvas.configure(yscrollcommand=self.thumb_scroll.set)

        self.thumb_inner = tk.Frame(self.thumb_canvas, bd=0, highlightthickness=0)
        self.thumb_window = self.thumb_canvas.create_window((0, 0), window=self.thumb_inner, anchor="nw")
        self.thumb_inner.bind("<Configure>", self._on_thumb_inner_configure)
        self.thumb_canvas.bind("<Configure>", self._on_thumb_canvas_configure)
        self.thumb_canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.thumb_canvas.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.thumb_canvas.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")

        self.empty_label = ttk.Label(
            self.thumb_inner,
            text="Open a PDF to begin organizing pages visually.",
            style="CardBody.TLabel",
            justify="center",
            wraplength=460,
        )
        self.empty_label.grid(row=0, column=0, padx=30, pady=40)

    def apply_theme(self, palette: ThemePalette) -> None:
        self.palette = palette
        self.thumb_canvas.configure(bg=palette.input_bg, highlightbackground=palette.border)
        self.thumb_inner.configure(bg=palette.input_bg)
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.apply_theme(palette)
        self._render_page_cards()

    def _on_thumb_inner_configure(self, _event=None) -> None:
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))

    def _on_thumb_canvas_configure(self, event) -> None:
        self.thumb_canvas.itemconfigure(self.thumb_window, width=max(event.width - 2, 100))
        new_columns = max(1, event.width // 220)
        if new_columns != self.card_columns:
            self.card_columns = new_columns
            self._render_page_cards()

    def _widget_belongs_to_thumb_area(self, widget: tk.Misc | None) -> bool:
        current = widget
        while current is not None:
            if current == self.thumb_canvas or current == self.thumb_inner:
                return True
            try:
                parent_name = current.winfo_parent()
                current = current.nametowidget(parent_name) if parent_name else None
            except Exception:
                current = None
        return False

    def _on_mousewheel(self, event) -> None:
        try:
            if not self.thumb_canvas.winfo_exists():
                return
            widget = self.thumb_canvas.winfo_containing(event.x_root, event.y_root)
            if not self._widget_belongs_to_thumb_area(widget):
                return
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                self.thumb_canvas.yview_scroll(delta, "units")
        except Exception:
            return

    def _on_mousewheel_linux(self, event) -> None:
        try:
            if not self.thumb_canvas.winfo_exists():
                return
            widget = self.thumb_canvas.winfo_containing(event.x_root, event.y_root)
            if not self._widget_belongs_to_thumb_area(widget):
                return
            delta = -1 if getattr(event, "num", None) == 4 else 1
            self.thumb_canvas.yview_scroll(delta, "units")
        except Exception:
            return

    def _close_document(self) -> None:
        if self.document is not None:
            try:
                self.document.close()
            except Exception:
                pass
        self.document = None

    def open_pdf_dialog(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Open PDF for organizer",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if file_path:
            self.load_pdf(Path(file_path))

    def load_pdf(self, input_pdf: Path) -> None:
        pdf_path = Path(input_pdf).expanduser()
        if not pdf_path.exists():
            messagebox.showerror("Organizer", f"PDF not found:\n{pdf_path}")
            return
        try:
            summary = pdf_summary(pdf_path)
            if summary.encrypted:
                raise OrganizerError("This PDF is password protected. Unlock it first before using the visual organizer.")
            self._close_document()
            self.document = fitz.open(str(pdf_path))
        except Exception as exc:
            self._close_document()
            messagebox.showerror("Organizer", f"Could not load the PDF for organization:\n{exc}")
            self.set_status("Organizer load failed.")
            return

        self.source_pdf = pdf_path
        self.page_sequence = build_default_sequence(summary.page_count)
        self.selected_positions = []
        self.thumbnail_cache.clear()
        self.file_var.set(str(pdf_path))
        meta_bits: list[str] = [f"Pages: {summary.page_count}", f"Size: {summary.file_size_label}"]
        if summary.title:
            meta_bits.append(f"Title: {summary.title}")
        if summary.author:
            meta_bits.append(f"Author: {summary.author}")
        if summary.subject:
            meta_bits.append(f"Subject: {summary.subject}")
        if summary.keywords:
            meta_bits.append(f"Keywords: {summary.keywords}")
        if summary.producer:
            meta_bits.append(f"Producer: {summary.producer}")
        self.summary_var.set("  |  ".join(meta_bits))
        self._update_selection_label()
        self._render_page_cards()
        self._update_button_states()
        self.set_status(f"Organizer loaded {summary.page_count} page(s) from {pdf_path.name}.")
        if self.on_loaded_pdf:
            self.on_loaded_pdf(pdf_path)

    def reload_pdf(self) -> None:
        if not self.source_pdf:
            self.open_pdf_dialog()
            return
        self.load_pdf(self.source_pdf)

    def _update_selection_label(self) -> None:
        count = len(self.selected_positions)
        total = len(self.page_sequence)
        self.selection_var.set(f"Selection: {count} page(s) selected  |  Current sequence length: {total} page(s)")

    def _update_button_states(self) -> None:
        has_document = bool(self.source_pdf and self.page_sequence)
        has_selection = bool(self.selected_positions)
        controls = [
            self.reload_button,
            self.save_button,
            self.move_up_button,
            self.move_down_button,
            self.rotate_left_button,
            self.rotate_right_button,
            self.duplicate_button,
            self.remove_button,
            self.reverse_button,
            self.extract_button,
            self.export_button,
        ]
        for button in controls:
            button.state(["!disabled"] if has_document else ["disabled"])
        for button in [
            self.move_up_button,
            self.move_down_button,
            self.rotate_left_button,
            self.rotate_right_button,
            self.duplicate_button,
            self.remove_button,
            self.extract_button,
            self.export_button,
        ]:
            button.state(["!disabled"] if has_selection else ["disabled"])
        if has_document and self.page_sequence:
            self.save_button.state(["!disabled"])
            self.reverse_button.state(["!disabled"])
            self.reload_button.state(["!disabled"])
        else:
            self.save_button.state(["disabled"])
            self.reverse_button.state(["disabled"])
            self.reload_button.state(["disabled"])

    def _render_page_cards(self) -> None:
        for child in self.thumb_inner.winfo_children():
            child.destroy()

        if not self.page_sequence or self.document is None:
            self.empty_label = ttk.Label(
                self.thumb_inner,
                text="Open a PDF to begin organizing pages visually.",
                style="CardBody.TLabel",
                justify="center",
                wraplength=460,
            )
            self.empty_label.grid(row=0, column=0, padx=30, pady=40)
            return

        columns = max(1, self.card_columns)
        for column in range(columns):
            self.thumb_inner.grid_columnconfigure(column, weight=1)

        for position, item in enumerate(self.page_sequence):
            cache_key = (item.source_index, item.rotation)
            photo = self.thumbnail_cache.get(cache_key)
            if photo is None:
                image = render_thumbnail_from_document(self.document, item.source_index, item.rotation)
                photo = ImageTk.PhotoImage(image)
                self.thumbnail_cache[cache_key] = photo

            selected = position in self.selected_positions
            bg = self.palette.surface_alt if selected else self.palette.surface
            border = self.palette.accent if selected else self.palette.border
            title_fg = self.palette.text
            meta_fg = self.palette.text_muted

            card = tk.Frame(
                self.thumb_inner,
                bg=bg,
                highlightthickness=2,
                highlightbackground=border,
                highlightcolor=border,
                bd=0,
                padx=8,
                pady=8,
                cursor="hand2",
            )
            row = position // columns
            column = position % columns
            card.grid(row=row, column=column, sticky="nsew", padx=10, pady=10)

            image_label = tk.Label(card, image=photo, bg=bg, cursor="hand2")
            image_label.image = photo
            image_label.grid(row=0, column=0, sticky="nsew")
            tk.Label(
                card,
                text=f"Slot {position + 1}  •  Source page {item.source_index + 1}",
                bg=bg,
                fg=title_fg,
                anchor="w",
                justify="left",
                font=("Segoe UI", 10, "bold"),
                wraplength=180,
                cursor="hand2",
            ).grid(row=1, column=0, sticky="ew", pady=(10, 2))
            status_parts: list[str] = []
            if item.rotation:
                status_parts.append(f"Rotation {item.rotation % 360}°")
            duplicate_count = sum(1 for other in self.page_sequence if other.source_index == item.source_index)
            if duplicate_count > 1:
                status_parts.append("Duplicated source")
            status_text = "  |  ".join(status_parts) if status_parts else "Original page"
            tk.Label(
                card,
                text=status_text,
                bg=bg,
                fg=meta_fg,
                anchor="w",
                justify="left",
                wraplength=180,
                cursor="hand2",
            ).grid(row=2, column=0, sticky="ew")

            for widget in (card, image_label):
                widget.bind("<Button-1>", lambda _event, pos=position: self.toggle_selection(pos))
                widget.bind("<Double-Button-1>", lambda _event, pos=position: self.open_preview(pos))
            for label in card.winfo_children()[1:]:
                label.bind("<Button-1>", lambda _event, pos=position: self.toggle_selection(pos))
                label.bind("<Double-Button-1>", lambda _event, pos=position: self.open_preview(pos))

        self.thumb_inner.update_idletasks()
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))

    def _ensure_selection(self) -> bool:
        if not self.selected_positions:
            messagebox.showinfo("Organizer", "Select at least one page thumbnail first.")
            return False
        return True

    def toggle_selection(self, position: int) -> None:
        if position in self.selected_positions:
            self.selected_positions = [value for value in self.selected_positions if value != position]
        else:
            self.selected_positions = sorted(self.selected_positions + [position])
        self._update_selection_label()
        self._update_button_states()
        self._render_page_cards()

    def select_all(self) -> None:
        if not self.page_sequence:
            return
        self.selected_positions = list(range(len(self.page_sequence)))
        self._update_selection_label()
        self._update_button_states()
        self._render_page_cards()
        self.set_status(f"Selected all {len(self.page_sequence)} page(s) in the organizer.")

    def clear_selection(self) -> None:
        self.selected_positions = []
        self._update_selection_label()
        self._update_button_states()
        self._render_page_cards()
        self.set_status("Organizer selection cleared.")

    def move_selected_up(self) -> None:
        if not self._ensure_selection():
            return
        self.page_sequence, self.selected_positions = move_positions_up(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Moved selected page(s) up.")

    def move_selected_down(self) -> None:
        if not self._ensure_selection():
            return
        self.page_sequence, self.selected_positions = move_positions_down(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Moved selected page(s) down.")

    def rotate_selected(self, delta: int) -> None:
        if not self._ensure_selection():
            return
        self.page_sequence = rotate_positions(self.page_sequence, self.selected_positions, delta)
        self._post_sequence_change("Updated rotation for selected page(s).")

    def duplicate_selected(self) -> None:
        if not self._ensure_selection():
            return
        self.page_sequence, self.selected_positions = duplicate_positions(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Duplicated the selected page(s).")

    def remove_selected(self) -> None:
        if not self._ensure_selection():
            return
        if len(self.selected_positions) == len(self.page_sequence):
            messagebox.showwarning("Organizer", "You cannot remove every page. Keep at least one page in the sequence.")
            return
        self.page_sequence, self.selected_positions = remove_positions(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Removed the selected page(s) from the sequence.")

    def reverse_pages(self) -> None:
        if not self.page_sequence:
            return
        self.page_sequence, self.selected_positions = reverse_sequence(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Reversed the current page order.")

    def _post_sequence_change(self, status_text: str) -> None:
        self.thumbnail_cache.clear()
        self._update_selection_label()
        self._update_button_states()
        self._render_page_cards()
        self.set_status(status_text)

    def save_organized_pdf(self) -> None:
        if not self.source_pdf or not self.page_sequence:
            messagebox.showinfo("Organizer", "Open a PDF first.")
            return
        default_name = f"{self.source_pdf.stem}_organized.pdf"
        save_path = filedialog.asksaveasfilename(
            title="Save organized PDF",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not save_path:
            return
        try:
            output = save_sequence_as_pdf(self.source_pdf, self.page_sequence, Path(save_path))
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not save the organized PDF:\n{exc}")
            self.set_status("Organizer save failed.")
            return
        self.set_status(f"Saved organized PDF: {output.name}")
        self._record_recent_job("Organizer -> Save PDF", [self.source_pdf], [output], note=f"Pages: {len(self.page_sequence)}")
        messagebox.showinfo("Organizer", f"Saved organized PDF:\n{output}")

    def extract_selected_pages(self) -> None:
        if not self.source_pdf or not self._ensure_selection():
            return
        default_name = f"{self.source_pdf.stem}_selected.pdf"
        save_path = filedialog.asksaveasfilename(
            title="Extract selected pages",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not save_path:
            return
        try:
            output = extract_selected_pdf(self.source_pdf, self.page_sequence, self.selected_positions, Path(save_path))
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not extract the selected pages:\n{exc}")
            self.set_status("Organizer extract failed.")
            return
        self.set_status(f"Extracted selected pages to {output.name}.")
        self._record_recent_job("Organizer -> Extract PDF", [self.source_pdf], [output], note=f"Selected pages: {len(self.selected_positions)}")
        messagebox.showinfo("Organizer", f"Saved extracted PDF:\n{output}")

    def export_selected_images(self) -> None:
        if not self.source_pdf or not self._ensure_selection():
            return
        output_dir = filedialog.askdirectory(title="Export selected pages as images")
        if not output_dir:
            return
        try:
            outputs = export_pages_as_images(self.source_pdf, self.page_sequence, self.selected_positions, Path(output_dir))
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not export selected pages as images:\n{exc}")
            self.set_status("Organizer image export failed.")
            return
        self.set_status(f"Exported {len(outputs)} page image(s) from the organizer.")
        self._record_recent_job("Organizer -> Export Images", [self.source_pdf], outputs, note=f"Selected pages: {len(self.selected_positions)}")
        messagebox.showinfo("Organizer", f"Exported {len(outputs)} page image(s) to:\n{Path(output_dir)}")

    def open_preview(self, position: int) -> None:
        if self.document is None or position < 0 or position >= len(self.page_sequence):
            return
        item = self.page_sequence[position]
        try:
            preview_image = render_preview_from_document(self.document, item.source_index, item.rotation)
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not render the preview:\n{exc}")
            return
        photo = ImageTk.PhotoImage(preview_image)
        title = f"Preview - slot {position + 1} / source page {item.source_index + 1}"
        details = (
            f"Source file: {self.source_pdf.name if self.source_pdf else ''}\n"
            f"Sequence slot: {position + 1}\n"
            f"Source page: {item.source_index + 1}\n"
            f"Rotation: {item.rotation % 360}°"
        )
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.destroy()
        self.preview_window = PreviewWindow(self, self.palette, title, photo, details)
        self.preview_window.preview_photo = photo

    def _record_recent_job(self, mode: str, inputs: list[Path], outputs: list[Path], note: str = "") -> None:
        if not self.on_recent_job:
            return
        preview_inputs = [str(path) for path in inputs[:8]]
        preview_outputs = [str(path) for path in outputs[:8]]
        record: dict[str, object] = {
            "job_type": "organizer",
            "status": "Completed",
            "mode": mode,
            "file_count": len(inputs),
            "output_count": len(outputs),
            "output_dir": str(outputs[0].parent if outputs else Path.cwd()),
            "inputs_preview": preview_inputs,
            "outputs_preview": preview_outputs,
            "note": note,
            "error": "",
        }
        self.on_recent_job(record)
