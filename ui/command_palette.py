"""Quick Actions palette.

Replaces the pre-2.4.0 ttk version, which called ``grab_set()``. On macOS that
grab blocks the root window's event loop outright -- the app appears to hang the
moment the palette opens. A command palette does not need to be app-modal, so
this one is a plain non-modal window that closes on Escape or focus loss.
"""

from __future__ import annotations

import tkinter as tk

import customtkinter as ctk


class CommandPalette(ctk.CTkToplevel):
    def __init__(self, app, *, actions: list) -> None:
        super().__init__(app)
        self.app = app
        self.kit = app.kit
        kit = self.kit
        palette = kit.palette

        self.actions = list(actions)
        self.filtered = list(actions)
        self.selected = 0

        self.title("Quick Actions")
        self.geometry("640x480")
        self.minsize(520, 380)
        self.configure(fg_color=palette.gradient_mid)
        # transient keeps it above the app without grabbing input.
        self.transient(app)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = kit.frame(self, variant="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=kit.pad("xl"), pady=(kit.pad("xl"), kit.pad("md")))
        header.grid_columnconfigure(0, weight=1)
        kit.label(header, "Quick Actions", variant="display").grid(row=0, column=0, sticky="w")
        kit.label(
            header,
            "Type to filter. ↑ ↓ to move, Enter to run, Esc to close.",
            variant="small",
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(kit.pad("xs"), 0))

        card = kit.card(self)
        card.grid(row=1, column=0, sticky="nsew", padx=kit.pad("xl"))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        body = kit.frame(card, variant="transparent")
        body.grid(row=0, column=0, sticky="nsew", padx=kit.pad("lg"), pady=kit.pad("lg"))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.query_var = tk.StringVar()
        self.entry = kit.entry(body, self.query_var, placeholder="Search actions…")
        self.entry.grid(row=0, column=0, sticky="ew")

        self.listbox = kit.listbox(body, activestyle="none", height=12)
        self.listbox.grid(row=1, column=0, sticky="nsew", pady=(kit.pad("md"), 0))
        self.listbox.bind("<Double-Button-1>", self._run)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        self.hint_var = tk.StringVar(value="")
        kit.label(body, "", variant="small", textvariable=self.hint_var, wraplength=520,
                  justify="left", anchor="w").grid(row=2, column=0, sticky="ew", pady=(kit.pad("md"), 0))

        actions_row = kit.button_bar(self, columns=2)
        actions_row.grid(row=2, column=0, sticky="e", padx=kit.pad("xl"), pady=kit.pad("xl"))
        actions_row.button("Run", self._run, variant="primary")
        actions_row.button("Close", self.destroy, variant="ghost")

        for sequence, handler in (
            ("<Escape>", lambda _e: self.destroy()),
            ("<Return>", self._run),
            ("<Down>", lambda _e: self._move(1)),
            ("<Up>", lambda _e: self._move(-1)),
        ):
            self.bind(sequence, handler)
        self.query_var.trace_add("write", lambda *_: self._refresh())

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._refresh()
        self.after(80, self._take_focus)

    def _take_focus(self) -> None:
        try:
            self.lift()
            self.focus_force()
            self.entry.focus_set()
        except Exception:
            pass

    # -- filtering ----------------------------------------------------------

    def _refresh(self) -> None:
        query = self.query_var.get().strip().lower()
        if query:
            self.filtered = [
                item
                for item in self.actions
                if query in item.label.lower()
                or query in getattr(item, "keywords", "").lower()
                or query in getattr(item, "hint", "").lower()
            ]
        else:
            self.filtered = list(self.actions)

        self.listbox.delete(0, tk.END)
        for item in self.filtered:
            self.listbox.insert(tk.END, item.label)
        self.selected = 0
        if self.filtered:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            self.hint_var.set(getattr(self.filtered[0], "hint", ""))
        else:
            self.hint_var.set(f"Nothing matches “{query}”. Try a different word.")

    def _on_select(self, _event=None) -> None:
        selection = self.listbox.curselection()
        if selection:
            self.selected = selection[0]
            self.hint_var.set(getattr(self.filtered[self.selected], "hint", ""))

    def _move(self, delta: int) -> str:
        if not self.filtered:
            return "break"
        self.selected = max(0, min(len(self.filtered) - 1, self.selected + delta))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(self.selected)
        self.listbox.activate(self.selected)
        self.listbox.see(self.selected)
        self.hint_var.set(getattr(self.filtered[self.selected], "hint", ""))
        return "break"

    def _run(self, _event=None) -> str:
        if not self.filtered:
            return "break"
        action = self.filtered[min(self.selected, len(self.filtered) - 1)]
        self.destroy()
        # Run after teardown so the action can open its own dialogs cleanly.
        self.app.after(10, action.command)
        return "break"

    def apply_theme(self, palette, *, compact: bool | None = None, scale: float | None = None) -> None:
        try:
            self.configure(fg_color=palette.gradient_mid)
        except Exception:
            pass
