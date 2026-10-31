from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from tkinter import ttk
from typing import Callable

from ui_theme import ThemePalette, apply_text_widget_theme, apply_ttk_theme


@dataclass()
class QuickAction:
    label: str
    command: Callable[[], None]
    hint: str = ""
    keywords: str = ""


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str, *, delay_ms: int = 480) -> None:
        self.widget = widget
        self.text = text
        self.delay_ms = max(100, int(delay_ms))
        self._after_id: str | None = None
        self._tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, _event=None) -> None:
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _on_leave(self, _event=None) -> None:
        self._cancel()
        if self._tip and self._tip.winfo_exists():
            self._tip.destroy()
        self._tip = None

    def _cancel(self) -> None:
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self) -> None:
        if self._tip or not self.text.strip():
            return
        if not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        self._tip.attributes("-topmost", True)

        root = self.widget.winfo_toplevel()
        palette = getattr(root, "palette", None)
        background = getattr(palette, "surface_alt", "#1f2430")
        foreground = getattr(palette, "text", "#f5f7fb")
        border = getattr(palette, "border", "#334155")

        try:
            self._tip.configure(background=background, highlightthickness=1, highlightbackground=border)
        except Exception:
            pass

        label = tk.Label(
            self._tip,
            text=self.text,
            justify="left",
            padx=10,
            pady=7,
            relief="solid",
            borderwidth=1,
            background=background,
            foreground=foreground,
            font=("Segoe UI", 9),
            wraplength=320,
        )
        label.pack()
