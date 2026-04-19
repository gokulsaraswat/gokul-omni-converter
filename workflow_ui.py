from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from tkinter import ttk
from typing import Callable

from ui_theme import ThemePalette, apply_text_widget_theme, apply_ttk_theme


@dataclass(slots=True)
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


class CommandPaletteWindow(tk.Toplevel):
    def __init__(self, master: tk.Misc, *, actions: list[QuickAction], palette: ThemePalette, compact: bool = False, scale: float = 1.0) -> None:
        super().__init__(master)
        self.palette = palette
        self.compact = compact
        self.scale = scale
        self.actions = list(actions)
        self.filtered_actions = list(actions)
        self.title("Quick Actions")
        self.geometry("620x430")
        self.minsize(520, 340)
        self.transient(master)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="Card.TFrame", padding=16)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))
        header.grid_columnconfigure(0, weight=1)
        ttk.Label(header, text="Quick Actions", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Search shortcuts, navigation, and common tools. Press Enter to run the selected action.",
            style="CardBody.TLabel",
            wraplength=500,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        search_row = ttk.Frame(self, style="Surface.TFrame", padding=(16, 0, 16, 10))
        search_row.grid(row=1, column=0, sticky="nsew")
        search_row.grid_columnconfigure(0, weight=1)
        search_row.grid_rowconfigure(1, weight=1)

        self.query_var = tk.StringVar()
        entry = ttk.Entry(search_row, textvariable=self.query_var)
        entry.grid(row=0, column=0, sticky="ew")
        entry.bind("<KeyRelease>", self._on_query_changed)
        entry.bind("<Return>", self._run_selected)
        self.entry = entry

        self.listbox = tk.Listbox(search_row, activestyle="none", relief="flat", borderwidth=1)
        self.listbox.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.listbox.bind("<Double-Button-1>", self._run_selected)
        self.listbox.bind("<Return>", self._run_selected)

        footer = ttk.Frame(self, style="Surface.TFrame", padding=(16, 0, 16, 16))
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)
        self.hint_var = tk.StringVar(value="")
        ttk.Label(footer, textvariable=self.hint_var, style="CardBody.TLabel", justify="left", wraplength=500).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(footer, text="Run action", style="Primary.TButton", command=self._run_selected).grid(row=0, column=1, padx=(12, 0))
        ttk.Button(footer, text="Close", command=self.destroy).grid(row=0, column=2, padx=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.apply_theme(palette, compact=compact, scale=scale)
        self._refresh_results()
        self.after(60, self.entry.focus_set)

    def apply_theme(self, palette: ThemePalette, *, compact: bool | None = None, scale: float | None = None) -> None:
        self.palette = palette
        if compact is not None:
            self.compact = compact
        if scale is not None:
            self.scale = scale
        apply_ttk_theme(self, palette, compact=self.compact, scale=self.scale)
        apply_text_widget_theme(self.listbox, palette)

    def _on_query_changed(self, _event=None) -> None:
        self._refresh_results()

    def _refresh_results(self) -> None:
        query = self.query_var.get().strip().lower()
        if not query:
            self.filtered_actions = list(self.actions)
        else:
            self.filtered_actions = [
                item for item in self.actions
                if query in item.label.lower()
                or query in item.hint.lower()
                or query in item.keywords.lower()
            ]
        self.listbox.delete(0, tk.END)
        for item in self.filtered_actions:
            hint = f" — {item.hint}" if item.hint else ""
            self.listbox.insert(tk.END, f"{item.label}{hint}")
        if self.filtered_actions:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            self.hint_var.set(self.filtered_actions[0].hint)
        else:
            self.hint_var.set("No actions matched that search.")

    def _selected_action(self) -> QuickAction | None:
        if not self.filtered_actions:
            return None
        selection = self.listbox.curselection()
        index = int(selection[0]) if selection else 0
        if 0 <= index < len(self.filtered_actions):
            return self.filtered_actions[index]
        return None

    def _run_selected(self, _event=None) -> None:
        action = self._selected_action()
        if not action:
            return
        self.destroy()
        action.command()
