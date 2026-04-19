from __future__ import annotations

from pathlib import Path
from tkinter import filedialog, ttk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from typing import Callable, Iterable

from PIL import ImageTk

from preview_support import PreviewResult, render_preview
from ui_theme import ThemePalette, apply_text_widget_theme, apply_ttk_theme


class PreviewCenterWindow(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        palette: ThemePalette,
        *,
        open_path_callback: Callable[[str | Path], None],
        compact: bool = False,
        scale: float = 1.0,
    ) -> None:
        super().__init__(master)
        self.palette = palette
        self.compact = compact
        self.scale = scale
        self.open_path_callback = open_path_callback

        self.title("Gokul Omni Convert Lite - Preview Center")
        self.geometry("1240x820")
        self.minsize(980, 620)

        self.paths: list[Path] = []
        self.current_index = 0
        self.current_page = 0
        self.current_zoom = 1.0
        self.current_result: PreviewResult | None = None
        self._photo: ImageTk.PhotoImage | None = None
        self._resize_after_id: str | None = None

        self.context_var = tk.StringVar(value="Load selected inputs, recent outputs, or add files directly to inspect them here.")
        self.meta_var = tk.StringVar(value="No file loaded.")
        self.page_var = tk.StringVar(value="Page 1 / 1")
        self.zoom_var = tk.StringVar(value="100%")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 10))
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Preview Center", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.context_var, style="CardBody.TLabel", wraplength=940, justify="left").grid(row=1, column=0, sticky="w", pady=(6, 0))

        left = ttk.LabelFrame(self, text="Loaded files")
        left.grid(row=1, column=0, sticky="nsew", padx=(16, 10), pady=(0, 16))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        actions = ttk.Frame(left)
        actions.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 8))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)
        ttk.Button(actions, text="Add files", command=self._add_files).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions, text="Remove", command=self._remove_selected).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(actions, text="Clear", command=self.clear_files).grid(row=0, column=2, sticky="ew", padx=(8, 0))

        list_frame = ttk.Frame(left)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        self.files_listbox = tk.Listbox(list_frame, relief="flat", borderwidth=1, exportselection=False)
        self.files_listbox.grid(row=0, column=0, sticky="nsew")
        self.files_listbox.bind("<<ListboxSelect>>", self._on_select)
        self.files_listbox.bind("<Double-1>", lambda _event: self._refresh_preview())
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.files_listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.files_listbox.configure(yscrollcommand=scroll.set)

        item_row = ttk.Frame(left)
        item_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        for index in range(2):
            item_row.grid_columnconfigure(index, weight=1)
        ttk.Button(item_row, text="Prev file", command=lambda: self._move_item(-1)).grid(row=0, column=0, sticky="ew")
        ttk.Button(item_row, text="Next file", command=lambda: self._move_item(1)).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        right = ttk.LabelFrame(self, text="Preview")
        right.grid(row=1, column=1, sticky="nsew", padx=(0, 16), pady=(0, 16))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=1)
        right.grid_rowconfigure(4, weight=1)

        toolbar = ttk.Frame(right)
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        toolbar.grid_columnconfigure(7, weight=1)
        ttk.Button(toolbar, text="Prev page", command=lambda: self._move_page(-1)).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(toolbar, text="Next page", command=lambda: self._move_page(1)).grid(row=0, column=1, padx=(0, 8))
        ttk.Label(toolbar, textvariable=self.page_var, style="Surface.TLabel").grid(row=0, column=2, padx=(0, 12))
        ttk.Label(toolbar, text="Zoom:", style="Surface.TLabel").grid(row=0, column=3)
        zoom_combo = ttk.Combobox(toolbar, textvariable=self.zoom_var, state="readonly", width=8, values=("75%", "100%", "125%", "150%", "200%"))
        zoom_combo.grid(row=0, column=4, padx=(8, 8))
        zoom_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_zoom_changed())
        ttk.Button(toolbar, text="Refresh", command=self._refresh_preview).grid(row=0, column=5, padx=(0, 8))
        ttk.Button(toolbar, text="Open file", command=self._open_file).grid(row=0, column=6, padx=(0, 8))
        ttk.Button(toolbar, text="Open folder", command=self._open_folder).grid(row=0, column=7, sticky="e")

        ttk.Label(right, textvariable=self.meta_var, style="CardBody.TLabel", wraplength=780, justify="left").grid(
            row=1, column=0, sticky="w", padx=12
        )

        self.preview_frame = ttk.Frame(right)
        self.preview_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(10, 10))
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_label = tk.Label(self.preview_frame, bd=0, relief="flat", anchor="center")
        self.preview_label.grid(row=0, column=0, sticky="nsew")
        self.preview_label.bind("<Configure>", self._schedule_render_from_cache)

        ttk.Label(right, text="Summary", style="CardTitle.TLabel").grid(row=3, column=0, sticky="w", padx=12)
        self.summary_text = ScrolledText(right, wrap=tk.WORD, relief="flat", borderwidth=1, height=10, padx=12, pady=10)
        self.summary_text.grid(row=4, column=0, sticky="nsew", padx=12, pady=(8, 12))
        self.summary_text.configure(state="disabled")

        self.apply_theme(palette, compact=compact, scale=scale)

    def apply_theme(self, palette: ThemePalette, *, compact: bool | None = None, scale: float | None = None) -> None:
        self.palette = palette
        self.compact = self.compact if compact is None else compact
        self.scale = self.scale if scale is None else scale
        apply_ttk_theme(self, palette, compact=self.compact, scale=self.scale)
        self.configure(background=palette.root_bg)
        self.preview_label.configure(background=palette.surface, foreground=palette.text)
        self.files_listbox.configure(
            background=palette.input_bg,
            foreground=palette.text,
            selectbackground=palette.selection,
            selectforeground=palette.accent_text,
            highlightbackground=palette.border,
            highlightcolor=palette.accent,
        )
        apply_text_widget_theme(self.summary_text, palette)

    def set_context_title(self, text: str) -> None:
        if text.strip():
            self.context_var.set(text.strip())

    def set_files(self, paths: Iterable[str | Path], *, replace: bool = True, title: str | None = None) -> None:
        normalized: list[Path] = []
        seen: set[str] = set()
        for raw in paths:
            path = Path(raw).expanduser()
            key = str(path)
            if key not in seen:
                seen.add(key)
                normalized.append(path)
        if replace:
            self.paths = normalized
        else:
            existing = {str(item) for item in self.paths}
            for path in normalized:
                if str(path) not in existing:
                    self.paths.append(path)
        if title:
            self.context_var.set(title)
        self._refresh_listbox()
        if self.paths:
            self.current_index = max(0, min(self.current_index, len(self.paths) - 1))
            self.files_listbox.selection_clear(0, tk.END)
            self.files_listbox.selection_set(self.current_index)
            self.files_listbox.activate(self.current_index)
            self.current_page = 0
            self._refresh_preview()
        else:
            self.clear_preview()

    def clear_files(self) -> None:
        self.paths = []
        self.current_index = 0
        self.current_page = 0
        self._refresh_listbox()
        self.clear_preview()

    def clear_preview(self) -> None:
        self.current_result = None
        self._photo = None
        self.preview_label.configure(image="", text="No file selected")
        self.meta_var.set("No file loaded.")
        self.page_var.set("Page 1 / 1")
        self._set_summary("Load selected inputs, recent outputs, or add files to preview them here.")

    def _set_summary(self, text: str) -> None:
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert("1.0", text)
        self.summary_text.configure(state="disabled")

    def _refresh_listbox(self) -> None:
        self.files_listbox.delete(0, tk.END)
        for path in self.paths:
            self.files_listbox.insert(tk.END, path.name or str(path))

    def _selected_path(self) -> Path | None:
        if not self.paths:
            return None
        selection = self.files_listbox.curselection()
        if selection:
            self.current_index = int(selection[0])
        if 0 <= self.current_index < len(self.paths):
            return self.paths[self.current_index]
        return None

    def _on_select(self, _event=None) -> None:
        self.current_page = 0
        self._refresh_preview()

    def _move_item(self, delta: int) -> None:
        if not self.paths:
            return
        self.current_index = (self.current_index + delta) % len(self.paths)
        self.files_listbox.selection_clear(0, tk.END)
        self.files_listbox.selection_set(self.current_index)
        self.files_listbox.activate(self.current_index)
        self.current_page = 0
        self._refresh_preview()

    def _move_page(self, delta: int) -> None:
        if self.current_result is None or self.current_result.page_count <= 1:
            return
        self.current_page = max(0, min(self.current_page + delta, self.current_result.page_count - 1))
        self._refresh_preview()

    def _zoom_factor(self) -> float:
        raw = self.zoom_var.get().strip().replace("%", "")
        try:
            value = float(raw) / 100.0
        except Exception:
            value = 1.0
        return max(0.5, min(value, 3.0))

    def _on_zoom_changed(self) -> None:
        self.current_page = max(0, self.current_page)
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        path = self._selected_path()
        if path is None:
            self.clear_preview()
            return
        try:
            result = render_preview(path, page=self.current_page, zoom=self._zoom_factor())
        except Exception as exc:  # pragma: no cover - user-visible fallback
            self.meta_var.set(f"Preview error for {path.name}: {exc}")
            self._set_summary(f"Preview could not be rendered for:\n{path}\n\n{exc}")
            self.preview_label.configure(image="", text="Preview unavailable")
            return
        self.current_result = result
        self.current_page = result.current_page
        self.page_var.set(f"Page {result.current_page + 1} / {max(result.page_count, 1)}")
        meta = f"{result.kind.title()} • {path.name}"
        if path.exists():
            meta += f" • {path.suffix.lower() or 'file'} • {path.stat().st_size / 1024:.1f} KB"
        self.meta_var.set(meta)
        self._set_summary(result.summary + f"\n\nPath:\n{path}")
        self._render_current_image()

    def _schedule_render_from_cache(self, _event=None) -> None:
        if self._resize_after_id:
            try:
                self.after_cancel(self._resize_after_id)
            except Exception:
                pass
        self._resize_after_id = self.after(80, self._render_current_image)

    def _render_current_image(self) -> None:
        if self.current_result is None:
            return
        self._resize_after_id = None
        image = self.current_result.image.copy()
        max_width = max(320, self.preview_frame.winfo_width() - 24)
        max_height = max(260, self.preview_frame.winfo_height() - 24)
        image.thumbnail((max_width, max_height))
        self._photo = ImageTk.PhotoImage(image)
        self.preview_label.configure(image=self._photo, text="")

    def _open_file(self) -> None:
        path = self._selected_path()
        if path is None:
            return
        self.open_path_callback(path)

    def _open_folder(self) -> None:
        path = self._selected_path()
        if path is None:
            return
        self.open_path_callback(path.parent if path.parent.exists() else path)

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Add files to Preview Center")
        if not paths:
            return
        self.set_files(paths, replace=False, title="Preview Center")

    def _remove_selected(self) -> None:
        selection = self.files_listbox.curselection()
        if not selection:
            return
        indices = sorted((int(index) for index in selection), reverse=True)
        for index in indices:
            if 0 <= index < len(self.paths):
                self.paths.pop(index)
        self.current_index = max(0, min(self.current_index, len(self.paths) - 1))
        self._refresh_listbox()
        if self.paths:
            self.files_listbox.selection_set(self.current_index)
            self._refresh_preview()
        else:
            self.clear_preview()
