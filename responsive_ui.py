from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageSequence, ImageTk


class ScrollablePage(ttk.Frame):
    """A simple responsive page shell with a vertical scrollbar.

    The inner frame stretches to at least the visible height so existing grid
    weights still behave well, while allowing overflow content to scroll.
    """

    def __init__(
        self,
        master: tk.Misc,
        *,
        style: str = "TFrame",
        inner_style: str = "TFrame",
        padding: int | tuple[int, ...] = 0,
    ) -> None:
        super().__init__(master, style=style)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.content = ttk.Frame(self.canvas, style=inner_style, padding=padding)
        self._window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.content.bind("<Configure>", self._on_content_configure, add="+")
        self.canvas.bind("<Configure>", self._on_canvas_configure, add="+")
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")
        self._last_height = 0

    def _widget_belongs_to_page(self, widget: tk.Misc | None) -> bool:
        current = widget
        while current is not None:
            if current in {self, self.canvas, self.content}:
                return True
            try:
                parent_name = current.winfo_parent()
                current = current.nametowidget(parent_name) if parent_name else None
            except Exception:
                current = None
        return False

    def _exclude_widget(self, widget: tk.Misc | None) -> bool:
        if widget is None:
            return False
        return isinstance(widget, (tk.Text, tk.Listbox, ttk.Treeview, ttk.Combobox, tk.Canvas))

    def _sync_window_bounds(self) -> None:
        try:
            canvas_width = max(self.canvas.winfo_width() - 2, 100)
            canvas_height = max(self.canvas.winfo_height() - 2, 100)
            requested_height = max(self.content.winfo_reqheight(), canvas_height)
            self.canvas.itemconfigure(self._window_id, width=canvas_width, height=requested_height)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self._last_height = requested_height
        except Exception:
            return

    def _on_content_configure(self, _event=None) -> None:
        self._sync_window_bounds()

    def _on_canvas_configure(self, _event=None) -> None:
        self._sync_window_bounds()

    def _on_mousewheel(self, event) -> None:
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if not self._widget_belongs_to_page(widget) or self._exclude_widget(widget):
                return
            delta = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")
        except Exception:
            return

    def _on_mousewheel_linux(self, event) -> None:
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if not self._widget_belongs_to_page(widget) or self._exclude_widget(widget):
                return
            delta = -1 if getattr(event, "num", None) == 4 else 1
            self.canvas.yview_scroll(delta, "units")
        except Exception:
            return

    def scroll_to_top(self) -> None:
        try:
            self.canvas.yview_moveto(0.0)
        except Exception:
            return

    def apply_palette(self, *, background: str, border: str) -> None:
        try:
            self.canvas.configure(background=background, highlightbackground=border)
        except Exception:
            return


def bind_responsive_wrap(
    widget: tk.Misc,
    *,
    padding: int = 32,
    min_wrap: int = 160,
    max_wrap: int | None = None,
    base_wrap: int | None = None,
    relative_to: tk.Misc | None = None,
) -> None:
    """Shrink wraplength when containers narrow, while preserving the original wide-screen cap."""
    if getattr(widget, "_gokul_wrap_bound", False):
        return
    setattr(widget, "_gokul_wrap_bound", True)

    target = relative_to or getattr(widget, "master", None)
    if target is None:
        return

    try:
        original = int(float(base_wrap if base_wrap is not None else widget.cget("wraplength")))
    except Exception:
        original = 0
    cap = max_wrap if max_wrap is not None else (original if original > 0 else None)

    def update(_event=None) -> None:
        if not widget.winfo_exists():
            return
        try:
            available = int(target.winfo_width()) - int(padding)
        except Exception:
            available = 0
        if available <= 0:
            try:
                available = int(widget.winfo_reqwidth())
            except Exception:
                return
        wrap = max(int(min_wrap), int(available))
        if cap is not None and cap > 0:
            wrap = min(wrap, int(cap))
        try:
            widget.configure(wraplength=wrap)
        except Exception:
            return

    try:
        target.bind("<Configure>", update, add="+")
    except Exception:
        pass
    try:
        widget.bind("<Map>", update, add="+")
    except Exception:
        pass
    try:
        widget.after(25, update)
    except Exception:
        pass


def resolve_flow_layout_width(
    current_width: int | float,
    parent_width: int | float,
    requested_width: int | float,
    *,
    min_width: int = 220,
) -> int:
    """Prefer the real allocated width so wrapping still works on narrow screens."""
    try:
        current = int(current_width)
    except Exception:
        current = 0
    try:
        parent = int(parent_width)
    except Exception:
        parent = 0
    try:
        requested = int(requested_width)
    except Exception:
        requested = 0

    for candidate in (current, parent):
        if candidate > 24:
            return max(min_width, candidate)
    if requested > 24:
        return max(min_width, requested)
    return max(min_width, 0)


class FlowButtonBar(ttk.Frame):
    """A small wrapping button bar for responsive toolbars and action rows."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        style: str = "TFrame",
        gap_x: int = 8,
        gap_y: int = 8,
        button_min_width: int = 0,
    ) -> None:
        super().__init__(master, style=style)
        self.gap_x = max(0, int(gap_x))
        self.gap_y = max(0, int(gap_y))
        self.button_min_width = max(0, int(button_min_width))
        self._items: list[tk.Widget] = []
        self.bind("<Configure>", self._schedule_layout, add="+")
        self._layout_after: str | None = None

    def add(self, widget: tk.Widget) -> tk.Widget:
        if widget not in self._items:
            self._items.append(widget)
        self._schedule_layout()
        return widget

    def clear(self) -> None:
        for widget in list(self._items):
            try:
                widget.destroy()
            except Exception:
                pass
        self._items.clear()
        self._schedule_layout()

    def _schedule_layout(self, _event=None) -> None:
        try:
            if self._layout_after:
                self.after_cancel(self._layout_after)
        except Exception:
            pass
        self._layout_after = self.after(30, self.relayout)

    def relayout(self) -> None:
        self._layout_after = None
        if not self.winfo_exists():
            return
        try:
            self.update_idletasks()
        except Exception:
            pass
        width = resolve_flow_layout_width(
            self.winfo_width(),
            getattr(self.master, "winfo_width", lambda: 0)(),
            self.winfo_reqwidth(),
            min_width=max(180, self.button_min_width or 0),
        )
        for widget in self._items:
            widget.grid_forget()

        for column in range(30):
            self.grid_columnconfigure(column, weight=0, uniform="")
        row = 0
        column = 0
        row_used = 0
        max_columns = 0

        for widget in self._items:
            req_width = max(widget.winfo_reqwidth(), self.button_min_width, 86)
            needed = req_width if column == 0 else req_width + self.gap_x
            if column > 0 and row_used + needed > width:
                row += 1
                column = 0
                row_used = 0
            padx = (0, self.gap_x) if column >= 0 else (0, 0)
            widget.grid(row=row, column=column, sticky="ew", padx=padx, pady=(0, self.gap_y))
            self.grid_columnconfigure(column, weight=1)
            row_used += needed
            column += 1
            max_columns = max(max_columns, column)

        if max_columns == 0:
            self.grid_columnconfigure(0, weight=1)
        for column in range(max_columns):
            self.grid_columnconfigure(column, uniform=str(id(self)))


class AnimatedGifLabel(ttk.Label):
    """Lightweight animated GIF label with a safe text fallback."""

    def __init__(
        self,
        master: tk.Misc,
        gif_path: str | Path,
        *,
        style: str = "Logo.TLabel",
        fallback_text: str = "",
        max_size: tuple[int, int] = (180, 54),
        frame_delay_ms: int = 110,
    ) -> None:
        super().__init__(master, style=style, text="")
        self.gif_path = Path(gif_path)
        self.fallback_text = fallback_text
        self.max_size = max_size
        self.frame_delay_ms = max(60, int(frame_delay_ms))
        self.frames: list[ImageTk.PhotoImage] = []
        self._frame_index = 0
        self._after_id: str | None = None
        self._load()
        self.bind("<Destroy>", self._on_destroy, add="+")

    def _load(self) -> None:
        self._stop()
        self.frames = []
        if self.gif_path.exists():
            try:
                image = Image.open(self.gif_path)
                for frame in ImageSequence.Iterator(image):
                    copy = frame.convert("RGBA")
                    copy.thumbnail(self.max_size)
                    self.frames.append(ImageTk.PhotoImage(copy))
                if self.frames:
                    self.configure(image=self.frames[0], text="")
                    if len(self.frames) > 1:
                        self._tick()
                    return
            except Exception:
                self.frames = []
        self.configure(image="", text=self.fallback_text)

    def _tick(self) -> None:
        if not self.winfo_exists() or len(self.frames) <= 1:
            return
        self._frame_index = (self._frame_index + 1) % len(self.frames)
        self.configure(image=self.frames[self._frame_index])
        self._after_id = self.after(self.frame_delay_ms, self._tick)

    def _stop(self) -> None:
        if self._after_id:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _on_destroy(self, _event=None) -> None:
        self._stop()

    def reload(self, gif_path: str | Path | None = None) -> None:
        if gif_path is not None:
            self.gif_path = Path(gif_path)
        self._load()
