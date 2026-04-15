from __future__ import annotations

import json
import tkinter as tk
from datetime import datetime
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
    move_positions_to_index,
    move_positions_up,
    pdf_summary,
    remove_positions,
    render_preview_from_document,
    render_thumbnail_from_document,
    reverse_sequence,
    rotate_positions,
    save_sequence_as_pdf,
    sequence_from_payload,
    sequence_to_payload,
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
        self.undo_stack: list[tuple[list[OrganizedPage], list[int]]] = []
        self.redo_stack: list[tuple[list[OrganizedPage], list[int]]] = []
        self.max_history = 40
        self.drag_source_position: int | None = None
        self.drag_target_position: int | None = None
        self.drag_active = False
        self.drag_start_xy: tuple[int, int] | None = None
        self.drag_status_label: tk.Label | None = None

        self.file_var = tk.StringVar(value="No PDF loaded yet.")
        self.summary_var = tk.StringVar(value="Load a PDF to inspect metadata, page count, and live page thumbnails.")
        self.selection_var = tk.StringVar(value="Selection: 0 page(s)")
        self.hint_var = tk.StringVar(
            value="Click thumbnails to toggle selection. Drag a selected card to reorder, or use undo/redo and layout snapshots."
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
                "Patch 19 expands the visual PDF organizer with drag-and-drop reordering, undo/redo history, "
                "layout snapshots, and the same visual rotate, duplicate, remove, extract, and export tools."
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
        for column in range(15):
            toolbar.grid_columnconfigure(column, weight=0)
        ttk.Button(toolbar, text="Select all", command=self.select_all).grid(row=0, column=0, padx=(0, 8), pady=4)
        ttk.Button(toolbar, text="Clear", command=self.clear_selection).grid(row=0, column=1, padx=(0, 8), pady=4)
        self.undo_button = ttk.Button(toolbar, text="Undo", command=self.undo_last_change)
        self.undo_button.grid(row=0, column=2, padx=(0, 8), pady=4)
        self.redo_button = ttk.Button(toolbar, text="Redo", command=self.redo_last_change)
        self.redo_button.grid(row=0, column=3, padx=(0, 8), pady=4)
        self.move_up_button = ttk.Button(toolbar, text="Move up", command=self.move_selected_up)
        self.move_up_button.grid(row=0, column=4, padx=(0, 8), pady=4)
        self.move_down_button = ttk.Button(toolbar, text="Move down", command=self.move_selected_down)
        self.move_down_button.grid(row=0, column=5, padx=(0, 8), pady=4)
        self.rotate_left_button = ttk.Button(toolbar, text="Rotate left", command=lambda: self.rotate_selected(-90))
        self.rotate_left_button.grid(row=0, column=6, padx=(0, 8), pady=4)
        self.rotate_right_button = ttk.Button(toolbar, text="Rotate right", command=lambda: self.rotate_selected(90))
        self.rotate_right_button.grid(row=0, column=7, padx=(0, 8), pady=4)
        self.duplicate_button = ttk.Button(toolbar, text="Duplicate", command=self.duplicate_selected)
        self.duplicate_button.grid(row=0, column=8, padx=(0, 8), pady=4)
        self.remove_button = ttk.Button(toolbar, text="Remove", command=self.remove_selected)
        self.remove_button.grid(row=0, column=9, padx=(0, 8), pady=4)
        self.reverse_button = ttk.Button(toolbar, text="Reverse order", command=self.reverse_pages)
        self.reverse_button.grid(row=0, column=10, padx=(0, 8), pady=4)
        self.layout_save_button = ttk.Button(toolbar, text="Save layout", command=self.save_layout_snapshot)
        self.layout_save_button.grid(row=0, column=11, padx=(0, 8), pady=4)
        self.layout_load_button = ttk.Button(toolbar, text="Load layout", command=self.load_layout_snapshot)
        self.layout_load_button.grid(row=0, column=12, padx=(0, 8), pady=4)
        self.extract_button = ttk.Button(toolbar, text="Extract selected PDF", command=self.extract_selected_pages)
        self.extract_button.grid(row=0, column=13, padx=(0, 8), pady=4)
        self.export_button = ttk.Button(toolbar, text="Export selected images", command=self.export_selected_images)
        self.export_button.grid(row=0, column=14, pady=4)

        thumb_shell = ttk.Frame(workspace, style="Card.TFrame", padding=0)
        thumb_shell.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        thumb_shell.grid_columnconfigure(0, weight=1)
        thumb_shell.grid_rowconfigure(0, weight=1)

        self.thumb_canvas = tk.Canvas(thumb_shell, highlightthickness=0, borderwidth=0, takefocus=1, cursor="arrow")
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
        self.bind_all("<Control-z>", self._on_shortcut_undo, add="+")
        self.bind_all("<Control-y>", self._on_shortcut_redo, add="+")
        self.bind_all("<Control-a>", self._on_shortcut_select_all, add="+")
        self.bind_all("<Delete>", self._on_shortcut_remove, add="+")

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


    def _widget_belongs_to_panel(self, widget: tk.Misc | None) -> bool:
        current = widget
        while current is not None:
            if current == self:
                return True
            try:
                parent_name = current.winfo_parent()
                current = current.nametowidget(parent_name) if parent_name else None
            except Exception:
                current = None
        return False

    def _organizer_shortcuts_enabled(self) -> bool:
        try:
            focused = self.focus_get()
        except Exception:
            focused = None
        if focused and self._widget_belongs_to_panel(focused):
            return True
        try:
            pointer_widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        except Exception:
            pointer_widget = None
        return bool(pointer_widget and self._widget_belongs_to_panel(pointer_widget))

    def _on_shortcut_undo(self, _event=None):
        if not self._organizer_shortcuts_enabled():
            return None
        self.undo_last_change()
        return "break"

    def _on_shortcut_redo(self, _event=None):
        if not self._organizer_shortcuts_enabled():
            return None
        self.redo_last_change()
        return "break"

    def _on_shortcut_select_all(self, _event=None):
        if not self._organizer_shortcuts_enabled():
            return None
        self.select_all()
        return "break"

    def _on_shortcut_remove(self, _event=None):
        if not self._organizer_shortcuts_enabled():
            return None
        if self.selected_positions:
            self.remove_selected()
            return "break"
        return None

    def _reset_drag_state(self) -> None:
        self.drag_source_position = None
        self.drag_target_position = None
        self.drag_start_xy = None
        self.drag_active = False

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
        self.undo_stack.clear()
        self.redo_stack.clear()
        self._reset_drag_state()
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
            self.undo_button,
            self.redo_button,
            self.move_up_button,
            self.move_down_button,
            self.rotate_left_button,
            self.rotate_right_button,
            self.duplicate_button,
            self.remove_button,
            self.reverse_button,
            self.layout_save_button,
            self.layout_load_button,
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
            self.layout_save_button.state(["!disabled"])
            self.layout_load_button.state(["!disabled"])
        else:
            self.save_button.state(["disabled"])
            self.reverse_button.state(["disabled"])
            self.reload_button.state(["disabled"])
            self.layout_save_button.state(["disabled"])
            self.layout_load_button.state(["disabled"])

        self.undo_button.state(["!disabled"] if self.undo_stack else ["disabled"])
        self.redo_button.state(["!disabled"] if self.redo_stack else ["disabled"])

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
            is_drop_target = self.drag_active and self.drag_target_position == position
            bg = self.palette.surface_alt if selected else self.palette.surface
            if is_drop_target and not selected:
                bg = self.palette.input_bg
            border = self.palette.accent if selected else self.palette.border
            if is_drop_target:
                border = self.palette.accent
            title_fg = self.palette.text
            meta_fg = self.palette.text_muted
            card_cursor = "fleur" if selected else "hand2"

            card = tk.Frame(
                self.thumb_inner,
                bg=bg,
                highlightthickness=3 if is_drop_target else 2,
                highlightbackground=border,
                highlightcolor=border,
                bd=0,
                padx=8,
                pady=8,
                cursor=card_cursor,
            )
            row = position // columns
            column = position % columns
            card.grid(row=row, column=column, sticky="nsew", padx=10, pady=10)

            image_label = tk.Label(card, image=photo, bg=bg, cursor=card_cursor)
            image_label.image = photo
            image_label.grid(row=0, column=0, sticky="nsew")
            title_text = f"Slot {position + 1}  •  Source page {item.source_index + 1}"
            if is_drop_target:
                title_text += "  •  Drop here"
            title_label = tk.Label(
                card,
                text=title_text,
                bg=bg,
                fg=title_fg,
                anchor="w",
                justify="left",
                font=("Segoe UI", 10, "bold"),
                wraplength=180,
                cursor=card_cursor,
            )
            title_label.grid(row=1, column=0, sticky="ew", pady=(10, 2))
            status_parts: list[str] = []
            if item.rotation:
                status_parts.append(f"Rotation {item.rotation % 360}°")
            duplicate_count = sum(1 for other in self.page_sequence if other.source_index == item.source_index)
            if duplicate_count > 1:
                status_parts.append("Duplicated source")
            if is_drop_target:
                status_parts.append("Release to place")
            status_text = "  |  ".join(status_parts) if status_parts else "Original page"
            status_label = tk.Label(
                card,
                text=status_text,
                bg=bg,
                fg=meta_fg,
                anchor="w",
                justify="left",
                wraplength=180,
                cursor=card_cursor,
            )
            status_label.grid(row=2, column=0, sticky="ew")

            interactive_widgets = (card, image_label, title_label, status_label)
            for widget in interactive_widgets:
                widget._organizer_position = position  # type: ignore[attr-defined]
                widget.bind("<ButtonPress-1>", lambda event, pos=position: self.on_card_press(event, pos))
                widget.bind("<B1-Motion>", lambda event, pos=position: self.on_card_motion(event, pos))
                widget.bind("<ButtonRelease-1>", lambda event, pos=position: self.on_card_release(event, pos))
                widget.bind("<Double-Button-1>", lambda _event, pos=position: self.open_preview(pos))

        self.thumb_inner.update_idletasks()
        self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))

    def _ensure_selection(self) -> bool:
        if not self.selected_positions:
            messagebox.showinfo("Organizer", "Select at least one page thumbnail first.")
            return False
        return True


    def _snapshot_sequence(self) -> tuple[list[OrganizedPage], list[int]]:
        return (list(self.page_sequence), list(self.selected_positions))

    def _push_undo_snapshot(self) -> None:
        if not self.page_sequence:
            return
        self.undo_stack.append(self._snapshot_sequence())
        if len(self.undo_stack) > self.max_history:
            self.undo_stack = self.undo_stack[-self.max_history :]
        self.redo_stack.clear()

    def _restore_snapshot(self, snapshot: tuple[list[OrganizedPage], list[int]], status_text: str) -> None:
        self.page_sequence = list(snapshot[0])
        self.selected_positions = list(snapshot[1])
        self.thumbnail_cache.clear()
        self._update_selection_label()
        self._update_button_states()
        self._render_page_cards()
        self.set_status(status_text)

    def undo_last_change(self) -> None:
        if not self.undo_stack:
            self.set_status("Organizer undo history is empty.")
            return
        self.redo_stack.append(self._snapshot_sequence())
        snapshot = self.undo_stack.pop()
        self._restore_snapshot(snapshot, "Undid the last organizer change.")

    def redo_last_change(self) -> None:
        if not self.redo_stack:
            self.set_status("Organizer redo history is empty.")
            return
        self.undo_stack.append(self._snapshot_sequence())
        snapshot = self.redo_stack.pop()
        self._restore_snapshot(snapshot, "Re-applied the last organizer change.")

    def _apply_drag_reorder(self, target_position: int) -> None:
        if self.drag_source_position is None or not self.selected_positions:
            return
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = move_positions_to_index(
            self.page_sequence,
            self.selected_positions,
            target_position,
        )
        self._post_sequence_change("Reordered the selected page card(s) with drag-and-drop.")

    def save_layout_snapshot(self) -> None:
        if not self.source_pdf or not self.page_sequence:
            messagebox.showinfo("Organizer", "Open a PDF first.")
            return
        default_name = f"{self.source_pdf.stem}_layout.json"
        save_path = filedialog.asksaveasfilename(
            title="Save organizer layout",
            initialfile=default_name,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
        )
        if not save_path:
            return
        payload = {
            "app": "Gokul Omni Convert Lite",
            "feature": "organizer_layout",
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "source_pdf_name": self.source_pdf.name,
            **sequence_to_payload(
                self.page_sequence,
                source_pdf=str(self.source_pdf),
                page_count=self.document.page_count if self.document is not None else len(self.page_sequence),
                selected_positions=self.selected_positions,
            ),
        }
        target = Path(save_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        self.set_status(f"Saved organizer layout: {target.name}")
        self._record_recent_job("Organizer -> Save Layout", [self.source_pdf], [target], note=f"Pages: {len(self.page_sequence)}")
        messagebox.showinfo("Organizer", f"Saved organizer layout:\n{target}")

    def load_layout_snapshot(self) -> None:
        layout_path = filedialog.askopenfilename(
            title="Load organizer layout",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not layout_path:
            return
        try:
            payload = json.loads(Path(layout_path).read_text(encoding="utf-8"))
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not read the layout file:\n{exc}")
            self.set_status("Organizer layout import failed.")
            return

        if not isinstance(payload, dict):
            messagebox.showerror("Organizer", "The selected layout file is not a valid JSON object.")
            self.set_status("Organizer layout import failed.")
            return

        source_pdf_value = str(payload.get("source_pdf", "")).strip()
        if (self.source_pdf is None or not self.source_pdf.exists()) and source_pdf_value:
            candidate = Path(source_pdf_value).expanduser()
            if candidate.exists():
                self.load_pdf(candidate)
        if self.source_pdf is None:
            messagebox.showinfo("Organizer", "Open the matching PDF first, then load the layout snapshot again.")
            self.set_status("Organizer layout import requires an open PDF.")
            return
        if self.document is None:
            messagebox.showinfo("Organizer", "The organizer PDF is not available yet. Reload the PDF and try again.")
            self.set_status("Organizer layout import requires an active document.")
            return

        expected_page_count = int(payload.get("page_count", self.document.page_count or 0) or 0)
        if expected_page_count and expected_page_count != self.document.page_count:
            messagebox.showerror(
                "Organizer",
                f"This layout expects {expected_page_count} page(s), but the loaded PDF has {self.document.page_count} page(s).",
            )
            self.set_status("Organizer layout import failed because the page counts do not match.")
            return

        try:
            sequence, selected = sequence_from_payload(payload, self.document.page_count)
        except Exception as exc:
            messagebox.showerror("Organizer", f"Could not apply the layout:\n{exc}")
            self.set_status("Organizer layout import failed.")
            return

        self._push_undo_snapshot()
        self.page_sequence = sequence
        self.selected_positions = selected
        self._post_sequence_change("Loaded the organizer layout snapshot.")
        self._record_recent_job("Organizer -> Load Layout", [Path(layout_path)], [self.source_pdf], note=f"Pages: {len(self.page_sequence)}")



    def on_card_press(self, event, position: int) -> None:
        self.thumb_canvas.focus_set()
        self.drag_source_position = int(position)
        self.drag_target_position = int(position)
        self.drag_start_xy = (int(event.x_root), int(event.y_root))
        self.drag_active = False

    def _position_from_pointer(self, x_root: int, y_root: int) -> int | None:
        widget = self.winfo_containing(int(x_root), int(y_root))
        while widget is not None:
            candidate = getattr(widget, "_organizer_position", None)
            if candidate is not None:
                try:
                    return int(candidate)
                except Exception:
                    return None
            try:
                parent_name = widget.winfo_parent()
                widget = widget.nametowidget(parent_name) if parent_name else None
            except Exception:
                widget = None
        return None

    def on_card_motion(self, event, position: int) -> None:
        if self.drag_source_position is None or self.drag_start_xy is None:
            return
        dx = abs(int(event.x_root) - int(self.drag_start_xy[0]))
        dy = abs(int(event.y_root) - int(self.drag_start_xy[1]))
        if not self.drag_active and max(dx, dy) < 8:
            return

        if not self.drag_active:
            if self.drag_source_position not in self.selected_positions:
                self.selected_positions = [self.drag_source_position]
                self._update_selection_label()
                self._update_button_states()
            self.drag_active = True

        target = self._position_from_pointer(int(event.x_root), int(event.y_root))
        if target is None:
            target = int(position)
        if target != self.drag_target_position:
            self.drag_target_position = target
            self._render_page_cards()
        self.set_status("Drag the highlighted cards, then release on the target slot.")

    def on_card_release(self, event, position: int) -> None:
        if self.drag_active:
            target = self._position_from_pointer(int(event.x_root), int(event.y_root))
            if target is None:
                target = self.drag_target_position if self.drag_target_position is not None else int(position)
            target = int(target)
            source = int(self.drag_source_position) if self.drag_source_position is not None else target
            should_reorder = target != source or len(self.selected_positions) > 1
            self._reset_drag_state()
            if should_reorder:
                self._apply_drag_reorder(target)
            else:
                self._render_page_cards()
                self.set_status("Organizer drag cancelled.")
            return

        self._reset_drag_state()
        self.toggle_selection(position)

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
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = move_positions_up(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Moved selected page(s) up.")

    def move_selected_down(self) -> None:
        if not self._ensure_selection():
            return
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = move_positions_down(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Moved selected page(s) down.")

    def rotate_selected(self, delta: int) -> None:
        if not self._ensure_selection():
            return
        self._push_undo_snapshot()
        self.page_sequence = rotate_positions(self.page_sequence, self.selected_positions, delta)
        self._post_sequence_change("Updated rotation for selected page(s).")

    def duplicate_selected(self) -> None:
        if not self._ensure_selection():
            return
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = duplicate_positions(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Duplicated the selected page(s).")

    def remove_selected(self) -> None:
        if not self._ensure_selection():
            return
        if len(self.selected_positions) == len(self.page_sequence):
            messagebox.showwarning("Organizer", "You cannot remove every page. Keep at least one page in the sequence.")
            return
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = remove_positions(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Removed the selected page(s) from the sequence.")

    def reverse_pages(self) -> None:
        if not self.page_sequence:
            return
        self._push_undo_snapshot()
        self.page_sequence, self.selected_positions = reverse_sequence(self.page_sequence, self.selected_positions)
        self._post_sequence_change("Reversed the current page order.")

    def _post_sequence_change(self, status_text: str) -> None:
        self.thumbnail_cache.clear()
        self._reset_drag_state()
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
