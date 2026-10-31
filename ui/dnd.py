"""Drag-and-drop file input, with a working fallback when tkdnd is absent.

``tkinterdnd2`` ships the native tkdnd binaries and is the only way to accept
OS-level file drops in Tk. It is an *optional* dependency: when it is missing
(or fails to load, which happens on some minimal Linux Tk builds), every drop
zone still works as a click-to-browse target and says so in its own label. No
screen ever becomes unusable because drag-and-drop is unavailable.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from typing import Callable

try:  # pragma: no cover - depends on the host's Tk build
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_IMPORT_OK = True
except Exception:  # pragma: no cover
    DND_FILES = "DND_Files"  # type: ignore[assignment]
    TkinterDnD = None  # type: ignore[assignment]
    _DND_IMPORT_OK = False

_dnd_ready = False


def dnd_available() -> bool:
    """True once tkdnd has actually been loaded into a live Tk interpreter."""
    return _DND_IMPORT_OK and _dnd_ready


def init_root(root: tk.Misc) -> bool:
    """Load tkdnd into ``root``'s interpreter. Safe to call on a CTk window.

    CustomTkinter's CTk subclasses tk.Tk, so rather than swapping the base class
    we pull tkdnd in by hand and graft on the DnDWrapper methods.
    """
    global _dnd_ready
    if not _DND_IMPORT_OK:
        return False
    if _dnd_ready:
        return True
    try:
        root.TkdndVersion = TkinterDnD._require(root)  # type: ignore[union-attr, attr-defined]
        # drop_target_register/dnd_bind live on DnDWrapper; borrow them for any widget.
        for name in ("drop_target_register", "drop_target_unregister", "dnd_bind"):
            setattr(tk.Misc, name, getattr(TkinterDnD.DnDWrapper, name))  # type: ignore[union-attr]
        _dnd_ready = True
    except Exception:
        _dnd_ready = False
    return _dnd_ready


def parse_drop_paths(data: str) -> list[Path]:
    """Split a tkdnd drop payload into paths.

    tkdnd hands over a Tcl list: space-separated, with brace-quoting for any path
    containing spaces. Parsing this by hand is unavoidable.
    """
    text = str(data or "").strip()
    paths: list[Path] = []
    buffer = ""
    depth = 0
    for char in text:
        if char == "{":
            depth += 1
            if depth == 1:
                continue
        elif char == "}":
            depth -= 1
            if depth == 0:
                if buffer:
                    paths.append(Path(buffer))
                buffer = ""
                continue
        if char == " " and depth == 0:
            if buffer:
                paths.append(Path(buffer))
            buffer = ""
            continue
        buffer += char
    if buffer:
        paths.append(Path(buffer))
    return [item for item in paths if str(item).strip()]


def register_drop_target(
    widget: tk.Misc,
    on_drop: Callable[[list[Path]], None],
    *,
    on_enter: Callable[[], None] | None = None,
    on_leave: Callable[[], None] | None = None,
) -> bool:
    """Accept file drops on ``widget``. Returns False when tkdnd is unavailable."""
    if not dnd_available():
        return False
    try:
        widget.drop_target_register(DND_FILES)  # type: ignore[attr-defined]

        def _handle(event) -> str:
            if on_leave:
                on_leave()
            paths = parse_drop_paths(getattr(event, "data", ""))
            if paths:
                on_drop(paths)
            return "break"

        widget.dnd_bind("<<Drop>>", _handle)  # type: ignore[attr-defined]
        if on_enter:
            widget.dnd_bind("<<DropEnter>>", lambda _e: on_enter() or "break")  # type: ignore[attr-defined]
        if on_leave:
            widget.dnd_bind("<<DropLeave>>", lambda _e: on_leave() or "break")  # type: ignore[attr-defined]
    except Exception:
        return False
    return True


class DropZone:
    """A dashed panel that takes dropped files and doubles as a browse button."""

    def __init__(
        self,
        parent: tk.Misc,
        kit,
        *,
        on_files: Callable[[list[Path]], None],
        on_browse: Callable[[], None] | None = None,
        title: str = "Drop files here",
        browse_text: str = "Browse…",
    ) -> None:
        self.kit = kit
        self.on_files = on_files
        self._default_title = title

        self.frame = kit.frame(parent, variant="alt", bordered=True, radius="lg")
        self.frame.grid_columnconfigure(0, weight=1)

        hint = (
            "Drag files or folders in, or use the button."
            if dnd_available()
            else "Drag-and-drop is unavailable on this system — use the button."
        )
        self.title_label = kit.label(self.frame, title, variant="heading")
        self.title_label.grid(row=0, column=0, pady=(kit.pad("lg"), kit.pad("xs")))
        self.hint_label = kit.label(self.frame, hint, variant="small")
        self.hint_label.grid(row=1, column=0)
        if on_browse is not None:
            kit.button(self.frame, browse_text, on_browse, variant="secondary").grid(
                row=2, column=0, pady=(kit.pad("md"), kit.pad("lg"))
            )
        else:
            self.hint_label.grid_configure(pady=(0, kit.pad("lg")))

        for target in (self.frame, self.title_label, self.hint_label):
            register_drop_target(
                target,
                self._handle_drop,
                on_enter=self._highlight,
                on_leave=self._unhighlight,
            )

    def _handle_drop(self, paths: list[Path]) -> None:
        self.on_files(paths)

    def _highlight(self) -> None:
        try:
            self.frame.configure(border_color=self.kit.palette.accent, border_width=2)
            self.title_label.configure(text="Release to add")
        except Exception:
            pass

    def _unhighlight(self) -> None:
        try:
            self.frame.configure(border_color=self.kit.palette.border, border_width=1)
            self.title_label.configure(text=self._default_title)
        except Exception:
            pass

    def grid(self, **kwargs) -> None:
        self.frame.grid(**kwargs)
