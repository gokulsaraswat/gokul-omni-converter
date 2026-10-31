"""CustomTkinter widget kit for GO Omni Pro.

Every screen builds from :class:`UIKit`. The kit owns three things pages should
never do themselves:

1. **Colour.** Widgets are created with explicit colours read from the active
   ``ThemePalette``. Pages ask for a *variant* ("primary", "muted", "card") and
   get whatever that means in the current theme.
2. **Retheming.** Each widget the kit creates is registered with the recipe used
   to build it, so :meth:`UIKit.retheme` can repaint the whole app in place when
   the user switches theme, high-contrast, compact, or UI scale.
3. **The gaps in CustomTkinter.** CTk has no tree, list, or rich-text widget, so
   the kit wraps ``ttk.Treeview``, ``tk.Listbox``, and ``tk.Text`` and keeps them
   painted to match the CTk widgets beside them.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Any, Callable

import customtkinter as ctk

import ui_theme
from ui_theme import (
    RADIUS,
    ThemePalette,
    apply_menu_theme,
    apply_text_widget_theme,
    apply_treeview_tag_colors,
    apply_ttk_theme,
    button_colors,
    card_colors,
    entry_colors,
    font_spec,
    space,
    surface_hover,
)

#: Frame background variants -> ThemePalette attribute.
FRAME_VARIANTS = {
    "surface": "surface",
    "card": "surface",
    "alt": "surface_alt",
    "sidebar": "sidebar",
    "root": "root_bg",
    "input": "input_bg",
}

#: Label variants -> (font role, weight, palette colour attribute).
LABEL_VARIANTS = {
    # Big thin numerals, Apple-Weather style. "thin" picks a lighter font family
    # because Tk only understands normal/bold weights.
    "hero": ("hero", "normal", "text"),
    "display": ("display", "normal", "text"),
    "title": ("title", "bold", "text"),
    "heading": ("heading", "bold", "text"),
    "subheading": ("subheading", "bold", "text_muted"),
    "body": ("body", "normal", "text"),
    "muted": ("body", "normal", "text_muted"),
    "small": ("small", "normal", "text_muted"),
    "tiny": ("tiny", "normal", "text_muted"),
    "success": ("body", "bold", "success"),
    "warning": ("body", "bold", "warning"),
    "danger": ("body", "bold", "danger"),
    "accent": ("body", "bold", "accent"),
}




class SafeDrawMixin:
    """Stops CustomTkinter's redraw from recursing into itself.

    Several CTk widgets call ``update_idletasks()`` from inside ``_draw``. That
    flushes pending <Configure> events, which fire ``_update_dimensions_event``,
    which calls ``_draw`` again -- an unbounded loop that hangs the app whenever a
    theme, font, or scale change resizes the widget. The nested call is always
    redundant (the outer ``_draw`` finishes and paints), so dropping it is safe.
    """

    def _draw(self, *args: Any, **kwargs: Any):
        if getattr(self, "_draw_in_progress", False):
            return None
        self._draw_in_progress = True
        try:
            return super()._draw(*args, **kwargs)
        finally:
            self._draw_in_progress = False


class TtkStateMixin:
    """Adds ttk's ``state([...])`` to a CustomTkinter widget.

    _update_mode_controls and _update_pdf_tool_controls enable and disable roughly
    two dozen widgets using ttk's state-spec API. Supporting it here keeps those
    handlers -- which are pure logic and worth preserving verbatim -- unchanged.
    """

    def state(self, spec: list[str] | None = None) -> tuple[str, ...]:
        if spec is None:
            return (str(self.cget("state")),)
        if any(item == "disabled" for item in spec):
            self.configure(state="disabled")
        elif any(item in {"!disabled", "readonly", "normal"} for item in spec):
            self.configure(state="normal")
        return tuple(spec)


class Button(SafeDrawMixin, TtkStateMixin, ctk.CTkButton):
    pass


class CheckBox(SafeDrawMixin, TtkStateMixin, ctk.CTkCheckBox):
    pass


class Entry(SafeDrawMixin, TtkStateMixin, ctk.CTkEntry):
    pass


class RadioButton(SafeDrawMixin, TtkStateMixin, ctk.CTkRadioButton):
    pass


class UIKit:
    """Widget factory bound to one root window and one palette at a time."""

    def __init__(
        self,
        root: tk.Misc,
        palette: ThemePalette,
        *,
        compact: bool = False,
        scale: float = 1.0,
    ) -> None:
        self.root = root
        self.palette = palette
        self.compact = compact
        self.scale = scale
        # (widget, kind, extra) for every widget we made, so retheme can repaint it.
        self._registry: list[tuple[tk.Misc, str, dict[str, Any]]] = []
        #: Shared backdrop bitmaps, keyed by (width bucket, height bucket, palette).
        self._gradient_cache: dict[tuple[int, int, str], Any] = {}
        self.apply_global_theme()

    # -- theming ------------------------------------------------------------

    def apply_global_theme(self) -> None:
        """Push the palette into CustomTkinter's and ttk's global state."""
        ui_theme.apply_ctk_theme(self.palette, compact=self.compact, scale=self.scale)
        self.style = apply_ttk_theme(self.root, self.palette, compact=self.compact, scale=self.scale)

    def retheme(self, palette: ThemePalette, *, compact: bool | None = None, scale: float | None = None) -> None:
        """Repaint every widget this kit created for the new palette."""
        self.palette = palette
        if compact is not None:
            self.compact = compact
        if scale is not None:
            self.scale = scale
        self._gradient_cache.clear()
        self.apply_global_theme()

        alive: list[tuple[tk.Misc, str, dict[str, Any]]] = []
        for widget, kind, extra in self._registry:
            try:
                if not widget.winfo_exists():
                    continue
            except Exception:
                continue
            try:
                self._paint(widget, kind, extra)
            except Exception:
                # A widget that refuses one property must not abort the whole repaint.
                pass
            alive.append((widget, kind, extra))
        self._registry = alive

    def _register(self, widget: tk.Misc, kind: str, **extra: Any) -> tk.Misc:
        self._registry.append((widget, kind, extra))
        return widget

    def _paint(self, widget: tk.Misc, kind: str, extra: dict[str, Any]) -> None:
        palette = self.palette
        if kind == "frame":
            variant = extra.get("variant", "surface")
            widget.configure(**self._frame_kwargs(variant, extra.get("bordered", False)))
        elif kind == "separator":
            widget.configure(fg_color=palette.border)
        elif kind == "info":
            widget.configure(text_color=palette.text_muted, font=self.font("body"))
        elif kind == "label":
            role, weight, color_attr = LABEL_VARIANTS[extra.get("variant", "body")]
            widget.configure(
                text_color=getattr(palette, color_attr),
                font=self.font(role, weight=weight),
            )
        elif kind == "button":
            widget.configure(
                **button_colors(palette, extra.get("variant", "primary")),
                font=self.font("body", weight="bold" if extra.get("variant") == "primary" else "normal"),
            )
        elif kind == "entry":
            widget.configure(**entry_colors(palette), font=self.font("body"))
        elif kind == "combo":
            # Colours come from the TCombobox ttk style; only the font needs setting.
            widget.configure(font=self.font("body"))
        elif kind == "check":
            widget.configure(
                fg_color=palette.accent,
                hover_color=palette.accent_hover,
                border_color=palette.border,
                text_color=palette.text,
                checkmark_color=palette.accent_text,
                font=self.font("body"),
            )
        elif kind == "switch":
            widget.configure(
                progress_color=palette.accent,
                button_color=palette.text_muted,
                button_hover_color=palette.text,
                fg_color=palette.surface_alt,
                text_color=palette.text,
                font=self.font("body"),
            )
        elif kind == "radio":
            widget.configure(
                fg_color=palette.accent,
                hover_color=palette.accent_hover,
                border_color=palette.border,
                text_color=palette.text,
                font=self.font("body"),
            )
        elif kind == "slider":
            widget.configure(
                progress_color=palette.accent,
                button_color=palette.accent,
                button_hover_color=palette.accent_hover,
                fg_color=palette.surface_alt,
            )
        elif kind == "progress":
            widget.configure(progress_color=palette.accent, fg_color=palette.surface_alt)
        elif kind == "log_box":
            widget.apply_palette(palette, self.font("small", mono=extra.get("mono", True)))
        elif kind == "segmented":
            widget.configure(
                fg_color=palette.surface_alt,
                selected_color=palette.accent,
                selected_hover_color=palette.accent_hover,
                unselected_color=palette.surface_alt,
                unselected_hover_color=surface_hover(palette),
                text_color=palette.text,
                font=self.font("body"),
            )
        elif kind == "scroll_area":
            widget.apply_palette(palette)
        elif kind == "tk_text":
            apply_text_widget_theme(widget, palette)  # type: ignore[arg-type]
            widget.configure(font=self.font("small", mono=True))
        elif kind == "listbox":
            apply_text_widget_theme(widget, palette)  # type: ignore[arg-type]
            widget.configure(font=self.font("body"))
        elif kind == "treeview":
            apply_treeview_tag_colors(widget, palette)  # type: ignore[arg-type]
        elif kind == "canvas":
            widget.configure(background=palette.surface, highlightbackground=palette.border)
        elif kind == "menu":
            apply_menu_theme(widget, palette)  # type: ignore[arg-type]

    # -- tokens -------------------------------------------------------------

    #: Roles rendered in a thin face rather than the regular UI family.
    THIN_ROLES = frozenset({"hero", "display"})

    def gradient_image(self, width: int, height: int):
        """Shared backdrop bitmap for (bucketed) size + palette. Returns (photo, key).

        Rendering is bucketed to 160px steps and memoised so a window resize does
        not rebuild an image per page on every pixel of drag.
        """
        width = max(int(width), 1)
        height = max(int(height), 1)
        bucket = (
            max(320, (width // 160 + 1) * 160),
            max(240, (height // 160 + 1) * 160),
            self.palette.name,
        )
        cache = self._gradient_cache
        if bucket in cache:
            return cache[bucket], bucket
        try:
            from PIL import Image, ImageTk
        except Exception:
            return None, bucket

        palette = self.palette
        stops = (
            (0.0, ui_theme._hex_to_rgb(palette.gradient_top)),
            (0.55, ui_theme._hex_to_rgb(palette.gradient_mid)),
            (1.0, ui_theme._hex_to_rgb(palette.gradient_bottom)),
        )
        small = Image.new("RGB", (32, 32))
        pixels = small.load()
        for y in range(32):
            for x in range(32):
                t = (x / 31 * 0.35) + (y / 31 * 0.65)  # mostly vertical, slight diagonal
                for index in range(len(stops) - 1):
                    left_pos, left = stops[index]
                    right_pos, right = stops[index + 1]
                    if left_pos <= t <= right_pos:
                        span = right_pos - left_pos or 1.0
                        ratio = (t - left_pos) / span
                        pixels[x, y] = tuple(round(left[c] + (right[c] - left[c]) * ratio) for c in range(3))
                        break
                else:
                    pixels[x, y] = stops[-1][1]
        try:
            photo = ImageTk.PhotoImage(small.resize(bucket[:2], Image.BICUBIC))
        except Exception:
            return None, bucket
        # Keep only a couple of buckets so resizing does not leak bitmaps.
        if len(cache) > 3:
            cache.clear()
        cache[bucket] = photo
        return photo, bucket

    def font(self, role: str = "body", *, weight: str = "normal", mono: bool = False):
        return font_spec(
            role,
            weight=weight,
            compact=self.compact,
            scale=self.scale,
            mono=mono,
            thin=role in self.THIN_ROLES and weight != "bold",
            root=self.root,
        )

    def pad(self, name: str = "md") -> int:
        return space(name, self.scale)

    def radius(self, name: str = "md") -> int:
        return RADIUS.get(name, RADIUS["md"])

    # -- containers ---------------------------------------------------------

    def _frame_kwargs(self, variant: str, bordered: bool) -> dict[str, Any]:
        palette = self.palette
        if variant == "transparent":
            kwargs: dict[str, Any] = {"fg_color": "transparent"}
        else:
            kwargs = {"fg_color": getattr(palette, FRAME_VARIANTS.get(variant, "surface"))}
        if bordered:
            kwargs["border_color"] = palette.border
        return kwargs

    def frame(
        self,
        parent: tk.Misc,
        *,
        variant: str = "transparent",
        bordered: bool = False,
        radius: str = "none",
        **kwargs: Any,
    ) -> ctk.CTkFrame:
        widget = ctk.CTkFrame(
            parent,
            corner_radius=self.radius(radius),
            border_width=1 if bordered else 0,
            **self._frame_kwargs(variant, bordered),
            **kwargs,
        )
        self._register(widget, "frame", variant=variant, bordered=bordered)
        return widget

    def card(
        self,
        parent: tk.Misc,
        *,
        title: str = "",
        subtitle: str = "",
        **kwargs: Any,
    ) -> ctk.CTkFrame:
        """A bordered surface panel. When ``title`` is given the card exposes
        ``.body`` -- the frame page content should actually go into."""
        outer = self.frame(parent, variant="card", bordered=True, radius="xl", **kwargs)
        if not title:
            outer.body = outer  # type: ignore[attr-defined]
            return outer

        outer.grid_columnconfigure(0, weight=1)
        header = self.frame(outer, variant="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=self.pad("lg"), pady=(self.pad("lg"), 0))
        header.grid_columnconfigure(0, weight=1)
        self.label(header, title, variant="heading").grid(row=0, column=0, sticky="w")
        if subtitle:
            self.label(header, subtitle, variant="small").grid(row=1, column=0, sticky="w", pady=(2, 0))

        body = self.frame(outer, variant="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=self.pad("lg"), pady=(self.pad("md"), self.pad("lg")))
        body.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)
        outer.body = body  # type: ignore[attr-defined]
        outer.header = header  # type: ignore[attr-defined]
        return outer

    def scrollable(self, parent: tk.Misc, **kwargs: Any) -> "ScrollArea":
        """A vertically scrolling page body.

        Deliberately *not* CTkScrollableFrame: in CustomTkinter 6.0 its scrollbar's
        ``set`` calls ``update_idletasks`` from inside ``_draw``, which recurses
        without end whenever a font or scale change resizes the content. ScrollArea
        pairs a plain canvas with a ttk scrollbar and has no such feedback path.
        """
        widget = ScrollArea(parent, self, **kwargs)
        self._register(widget, "scroll_area")
        return widget

    def separator(self, parent: tk.Misc, **kwargs: Any) -> ctk.CTkFrame:
        widget = ctk.CTkFrame(parent, height=1, fg_color=self.palette.border, corner_radius=0, **kwargs)
        self._register(widget, "separator")
        return widget

    # -- content ------------------------------------------------------------

    def label(self, parent: tk.Misc, text: str = "", *, variant: str = "body", **kwargs: Any) -> ctk.CTkLabel:
        role, weight, color_attr = LABEL_VARIANTS.get(variant, LABEL_VARIANTS["body"])
        widget = ctk.CTkLabel(
            parent,
            text=text,
            text_color=getattr(self.palette, color_attr),
            font=self.font(role, weight=weight),
            fg_color="transparent",
            **kwargs,
        )
        self._register(widget, "label", variant=variant)
        return widget

    def button(
        self,
        parent: tk.Misc,
        text: str,
        command: Callable[[], None] | None = None,
        *,
        variant: str = "secondary",
        width: int = 0,
        tooltip: str = "",
        **kwargs: Any,
    ) -> ctk.CTkButton:
        colors = button_colors(self.palette, variant)
        widget = Button(
            parent,
            text=text,
            command=command,
            corner_radius=self.radius("pill"),
            border_width=1 if variant in {"secondary", "ghost"} else 0,
            height=self._scaled(36),
            font=self.font("body", weight="bold" if variant == "primary" else "normal"),
            **({"width": width} if width else {}),
            **colors,
            **kwargs,
        )
        self._register(widget, "button", variant=variant)
        if tooltip:
            attach_tooltip(widget, tooltip, self)
        return widget

    def entry(
        self,
        parent: tk.Misc,
        textvariable: tk.Variable | None = None,
        *,
        placeholder: str = "",
        **kwargs: Any,
    ) -> ctk.CTkEntry:
        widget = Entry(
            parent,
            textvariable=textvariable,
            placeholder_text=placeholder or None,
            corner_radius=self.radius("lg"),
            border_width=1,
            height=self._scaled(36),
            font=self.font("body"),
            **entry_colors(self.palette),
            **kwargs,
        )
        self._register(widget, "entry")
        return widget

    def combo(
        self,
        parent: tk.Misc,
        values: list[str],
        variable: tk.Variable | None = None,
        *,
        command: Callable[[str], None] | None = None,
        width: int = 0,
        **kwargs: Any,
    ) -> ttk.Combobox:
        """Read-only dropdown.

        Deliberately ttk rather than CTkOptionMenu: in CustomTkinter 6.0 the option
        menu's ``_draw`` calls ``update_idletasks``, and with a screenful of them a
        theme change sets off a redraw cascade that never terminates. ttk.Combobox is
        already fully themed by apply_ttk_theme, and brings native ``state([...])``
        and ``<<ComboboxSelected>>`` that the existing handlers expect.
        """
        widget = ttk.Combobox(
            parent,
            textvariable=variable,
            values=list(values) or [""],
            state="readonly",
            # ttk widths are in characters, not pixels.
            width=max(8, int(width / 8)) if width else 18,
            **kwargs,
        )
        if command is not None:
            widget.bind("<<ComboboxSelected>>", lambda _event: command(widget.get()), add="+")
        self._register(widget, "combo")
        return widget

    def check(self, parent: tk.Misc, text: str, variable: tk.Variable, **kwargs: Any) -> ctk.CTkCheckBox:
        palette = self.palette
        widget = CheckBox(
            parent,
            text=text,
            variable=variable,
            corner_radius=self.radius("sm"),
            checkbox_width=self._scaled(20),
            checkbox_height=self._scaled(20),
            fg_color=palette.accent,
            hover_color=palette.accent_hover,
            border_color=palette.border,
            text_color=palette.text,
            checkmark_color=palette.accent_text,
            font=self.font("body"),
            **kwargs,
        )
        self._register(widget, "check")
        return widget

    def switch(self, parent: tk.Misc, text: str, variable: tk.Variable, **kwargs: Any) -> ctk.CTkSwitch:
        palette = self.palette
        widget = ctk.CTkSwitch(
            parent,
            text=text,
            variable=variable,
            progress_color=palette.accent,
            button_color=palette.text_muted,
            button_hover_color=palette.text,
            fg_color=palette.surface_alt,
            text_color=palette.text,
            font=self.font("body"),
            **kwargs,
        )
        self._register(widget, "switch")
        return widget

    def radio(self, parent: tk.Misc, text: str, variable: tk.Variable, value: Any, **kwargs: Any) -> ctk.CTkRadioButton:
        palette = self.palette
        widget = RadioButton(
            parent,
            text=text,
            variable=variable,
            value=value,
            fg_color=palette.accent,
            hover_color=palette.accent_hover,
            border_color=palette.border,
            text_color=palette.text,
            font=self.font("body"),
            **kwargs,
        )
        self._register(widget, "radio")
        return widget

    def slider(self, parent: tk.Misc, variable: tk.Variable, *, from_: float, to: float, **kwargs: Any) -> ctk.CTkSlider:
        palette = self.palette
        widget = ctk.CTkSlider(
            parent,
            variable=variable,
            from_=from_,
            to=to,
            progress_color=palette.accent,
            button_color=palette.accent,
            button_hover_color=palette.accent_hover,
            fg_color=palette.surface_alt,
            **kwargs,
        )
        self._register(widget, "slider")
        return widget

    def progress(self, parent: tk.Misc, **kwargs: Any) -> "ProgressBar":
        widget = ProgressBar(
            parent,
            animator=getattr(self.root, "animator", None),
            progress_color=self.palette.accent,
            fg_color=self.palette.surface_alt,
            corner_radius=self.radius("pill"),
            height=self._scaled(8),
            **kwargs,
        )
        self._register(widget, "progress")
        return widget

    def button_bar(self, parent: tk.Misc, *, columns: int = 0, **kwargs: Any) -> "ButtonBar":
        """Wrapping action row. Replaces the ttk-era FlowButtonBar."""
        return ButtonBar(parent, self, columns=columns, **kwargs)

    def segmented(
        self,
        parent: tk.Misc,
        values: list[str],
        variable: tk.Variable | None = None,
        *,
        command: Callable[[str], None] | None = None,
        **kwargs: Any,
    ) -> ctk.CTkSegmentedButton:
        palette = self.palette
        widget = ctk.CTkSegmentedButton(
            parent,
            values=list(values),
            variable=variable,
            command=command,
            corner_radius=self.radius("md"),
            fg_color=palette.surface_alt,
            selected_color=palette.accent,
            selected_hover_color=palette.accent_hover,
            unselected_color=palette.surface_alt,
            unselected_hover_color=surface_hover(palette),
            text_color=palette.text,
            font=self.font("body"),
            **kwargs,
        )
        self._register(widget, "segmented")
        return widget

    def textbox(self, parent: tk.Misc, *, mono: bool = True, height: int = 160, **kwargs: Any) -> "LogBox":
        """Scrolling multi-line text area.

        Also avoids CTkTextbox, which embeds the same CTkScrollbar that recurses
        on resize (see :meth:`scrollable`).
        """
        widget = LogBox(parent, self, mono=mono, height=height, **kwargs)
        self._register(widget, "log_box", mono=mono)
        return widget

    # -- widgets CustomTkinter does not provide ------------------------------

    def listbox(self, parent: tk.Misc, **kwargs: Any) -> tk.Listbox:
        """CTk has no list widget, so this is a ``tk.Listbox`` kept in palette."""
        widget = tk.Listbox(
            parent,
            activestyle="none",
            borderwidth=0,
            highlightthickness=1,
            relief="flat",
            font=self.font("body"),
            **kwargs,
        )
        apply_text_widget_theme(widget, self.palette)
        self._register(widget, "listbox")
        return widget

    def treeview(self, parent: tk.Misc, **kwargs: Any) -> ttk.Treeview:
        """CTk has no tree widget; this ``ttk.Treeview`` is styled by apply_ttk_theme."""
        widget = ttk.Treeview(parent, **kwargs)
        apply_treeview_tag_colors(widget, self.palette)
        self._register(widget, "treeview")
        return widget

    def tk_text(self, parent: tk.Misc, **kwargs: Any) -> tk.Text:
        """Plain ``tk.Text`` for rich/tagged content CTkTextbox cannot express."""
        widget = tk.Text(parent, borderwidth=0, highlightthickness=1, relief="flat", wrap="word", **kwargs)
        apply_text_widget_theme(widget, self.palette)
        widget.configure(font=self.font("small", mono=True))
        self._register(widget, "tk_text")
        return widget

    def track_canvas(self, widget: tk.Canvas) -> tk.Canvas:
        self._register(widget, "canvas")
        return widget

    def track_menu(self, widget: tk.Menu) -> tk.Menu:
        apply_menu_theme(widget, self.palette)
        self._register(widget, "menu")
        return widget

    # -- composites ---------------------------------------------------------

    def info(self, parent: tk.Misc, text: str, **kwargs: Any) -> "InfoDot":
        """Small ⓘ carrying help text on hover, instead of a wall of static prose."""
        widget = InfoDot(parent, self, text, **kwargs)
        self._register(widget, "info")
        return widget

    def labelled(self, parent: tk.Misc, text: str, help_text: str = "", *, variant: str = "body") -> ctk.CTkFrame:
        """A label with an optional ⓘ beside it."""
        row = self.frame(parent, variant="transparent")
        self.label(row, text, variant=variant).grid(row=0, column=0, sticky="w")
        if help_text:
            self.info(row, help_text).grid(row=0, column=1, padx=(self.pad("xs"), 0))
        return row

    def hoverable_card(self, parent: tk.Misc, **kwargs: Any) -> ctk.CTkFrame:
        """A card whose border warms toward the accent while the pointer is over it."""
        card = self.card(parent, **kwargs)
        base, hover = self.palette.border, self.palette.accent

        def enter(_event=None) -> None:
            try:
                card.configure(border_color=lerp_color(base, hover, 0.55))
            except Exception:
                pass

        def leave(_event=None) -> None:
            try:
                card.configure(border_color=self.palette.border)
            except Exception:
                pass

        card.bind("<Enter>", enter, add="+")
        card.bind("<Leave>", leave, add="+")
        return card

    def page_header(self, parent: tk.Misc, title: str, subtitle: str = "") -> ctk.CTkFrame:
        header = self.frame(parent, variant="transparent")
        header.grid_columnconfigure(0, weight=1)
        title_row = self.frame(header, variant="transparent")
        title_row.grid(row=0, column=0, sticky="w")
        self.label(title_row, title, variant="display").grid(row=0, column=0, sticky="w")
        if subtitle:
            # Long explanations live on the ⓘ so the page stays uncluttered.
            self.info(title_row, subtitle).grid(row=0, column=1, sticky="w", padx=(self.pad("sm"), 0))
        header.actions = self.frame(header, variant="transparent")  # type: ignore[attr-defined]
        header.actions.grid(row=0, column=1, rowspan=2, sticky="e")  # type: ignore[attr-defined]
        return header

    def field_row(
        self,
        parent: tk.Misc,
        row: int,
        label: str,
        widget: tk.Misc,
        *,
        hint: str = "",
        trailing: tk.Misc | None = None,
    ) -> int:
        """Label on the left, control on the right, optional hint underneath.

        Returns the next free grid row so callers can chain without counting.
        """
        self.label(parent, label, variant="body").grid(
            row=row, column=0, sticky="w", padx=(0, self.pad("md")), pady=self.pad("sm")
        )
        widget.grid(row=row, column=1, sticky="ew", pady=self.pad("sm"))
        if trailing is not None:
            trailing.grid(row=row, column=2, sticky="w", padx=(self.pad("sm"), 0), pady=self.pad("sm"))
        parent.grid_columnconfigure(1, weight=1)
        row += 1
        if hint:
            self.label(parent, hint, variant="tiny").grid(
                row=row, column=1, columnspan=2, sticky="w", pady=(0, self.pad("sm"))
            )
            row += 1
        return row

    def empty_state(
        self,
        parent: tk.Misc,
        *,
        title: str,
        body: str,
        action_text: str = "",
        action: Callable[[], None] | None = None,
    ) -> ctk.CTkFrame:
        """Shown where a list would be, when the list is empty. Always tells the
        user the one thing to do next rather than just saying 'nothing here'."""
        holder = self.frame(parent, variant="transparent")
        holder.grid_columnconfigure(0, weight=1)
        self.label(holder, title, variant="heading").grid(row=0, column=0, pady=(self.pad("xl"), self.pad("xs")))
        self.label(holder, body, variant="muted", wraplength=420, justify="center").grid(row=1, column=0)
        if action_text and action is not None:
            self.button(holder, action_text, action, variant="primary").grid(row=2, column=0, pady=self.pad("lg"))
        return holder

    def button_row(self, parent: tk.Misc, specs: list[tuple[str, Callable[[], None], str]], **grid: Any) -> ctk.CTkFrame:
        """A horizontal run of buttons from (text, command, variant) triples."""
        bar = self.frame(parent, variant="transparent")
        for index, (text, command, variant) in enumerate(specs):
            self.button(bar, text, command, variant=variant).grid(
                row=0, column=index, padx=(0 if index == 0 else self.pad("sm"), 0), sticky="w"
            )
        if grid:
            bar.grid(**grid)
        return bar

    def _scaled(self, value: int) -> int:
        return max(1, int(round(value * min(max(self.scale, 0.85), 1.6))))


# -- tooltips ----------------------------------------------------------------


class Tooltip:
    """Hover hint. Deliberately plain ``tk`` so it can float above CTk widgets."""

    def __init__(self, widget: tk.Misc, text: str, kit: UIKit) -> None:
        self.widget = widget
        self.text = text
        self.kit = kit
        self.window: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _event=None) -> None:
        if self.window is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        except Exception:
            return
        palette = self.kit.palette
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.window,
            text=self.text,
            background=palette.surface_alt,
            foreground=palette.text,
            highlightbackground=palette.border,
            highlightthickness=1,
            padx=8,
            pady=4,
            justify="left",
            wraplength=320,
            font=self.kit.font("small"),
        )
        label.pack()

    def _hide(self, _event=None) -> None:
        if self.window is not None:
            try:
                self.window.destroy()
            except Exception:
                pass
            self.window = None


def attach_tooltip(widget: tk.Misc, text: str, kit: UIKit) -> Tooltip:
    return Tooltip(widget, text, kit)


class ScrollArea(tk.Frame):
    """Vertically scrolling container that hosts CustomTkinter children.

    Mirrors the geometry behaviour of the pre-2.4.0 ScrollablePage: the inner
    frame stretches to at least the visible height so grid weights keep working,
    and the mouse wheel only scrolls when the pointer is actually over this page
    and not over a widget that scrolls itself.
    """

    def __init__(self, parent: tk.Misc, kit: UIKit, **kwargs: Any) -> None:
        super().__init__(parent, borderwidth=0, highlightthickness=0, **kwargs)
        self.kit = kit
        palette = kit.palette

        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=palette.gradient_bottom)
        self.backdrop = GradientBackdrop(self.canvas, kit)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Cards are opaque, so the content frame carries the gradient's mid tone;
        # the gradient itself stays visible in the page's outer padding.
        self.content = ctk.CTkFrame(self.canvas, fg_color=palette.gradient_mid, corner_radius=0)
        self._window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self._sync_after: str | None = None
        self._last_sync: tuple[int, int, int] | None = None
        self.content.bind("<Configure>", self._schedule_sync, add="+")
        self.canvas.bind("<Configure>", self._schedule_sync, add="+")
        self.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.bind_all("<Button-4>", self._on_wheel_x11, add="+")
        self.bind_all("<Button-5>", self._on_wheel_x11, add="+")

    def _schedule_sync(self, _event=None) -> None:
        """Coalesce Configure bursts into one geometry pass.

        Resizing the canvas window and showing/hiding the scrollbar are both layout
        changes, so doing them straight from a <Configure> handler generates more
        Configures. Debouncing keeps a resize to a couple of passes instead of
        hundreds.
        """
        if self._sync_after is not None:
            return
        self._sync_after = self.after_idle(self._sync)

    def _sync(self, _event=None) -> None:
        self._sync_after = None
        try:
            width = max(self.canvas.winfo_width() - 2, 100)
            height = max(self.canvas.winfo_height() - 2, 100)
            needed = self.content.winfo_reqheight()
            signature = (width, height, needed)
            if signature == self._last_sync:
                return
            self._last_sync = signature
            self.canvas.itemconfigure(self._window_id, width=width, height=max(needed, height))
            # Hide the scrollbar when everything already fits.
            if needed <= height:
                self.scrollbar.grid_remove()
            else:
                self.scrollbar.grid()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            return

    def _belongs(self, widget: tk.Misc | None) -> bool:
        while widget is not None:
            if widget in {self, self.canvas, self.content}:
                return True
            try:
                parent = widget.winfo_parent()
                widget = widget.nametowidget(parent) if parent else None
            except Exception:
                return False
        return False

    @staticmethod
    def _self_scrolling(widget: tk.Misc | None) -> bool:
        return isinstance(widget, (tk.Text, tk.Listbox, ttk.Treeview, tk.Canvas))

    def _on_wheel(self, event) -> None:
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if not self._belongs(widget) or self._self_scrolling(widget):
                return
            delta = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")
        except Exception:
            return

    def _on_wheel_x11(self, event) -> None:
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if not self._belongs(widget) or self._self_scrolling(widget):
                return
            self.canvas.yview_scroll(-1 if getattr(event, "num", None) == 4 else 1, "units")
        except Exception:
            return

    def scroll_to_top(self) -> None:
        try:
            self.canvas.yview_moveto(0.0)
        except Exception:
            return

    def apply_palette(self, palette) -> None:
        self._last_sync = None
        try:
            self.canvas.configure(background=palette.gradient_bottom)
            self.content.configure(fg_color=palette.gradient_mid)
            self.backdrop.refresh()
        except Exception:
            return


class LogBox(tk.Frame):
    """A ``tk.Text`` plus a ttk scrollbar, proxying the Text API.

    Used wherever the app shows an activity log. Callers treat it as a Text
    (``insert``, ``see``, ``delete``, ``configure(state=...)``) while ``grid``
    and ``pack`` place the whole composite.
    """

    _TEXT_METHODS = frozenset(
        {"insert", "delete", "see", "get", "index", "mark_set", "tag_add", "tag_config",
         "tag_configure", "yview", "edit_reset", "focus_set", "bind"}
    )

    def __init__(self, parent: tk.Misc, kit: UIKit, *, mono: bool = True, height: int = 160, **kwargs: Any) -> None:
        super().__init__(parent, borderwidth=0, highlightthickness=0, **kwargs)
        self.kit = kit
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.configure(height=height)
        self.grid_propagate(False)

        self.text = tk.Text(self, borderwidth=0, highlightthickness=1, relief="flat", wrap="word")
        self.text.grid(row=0, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=self.scrollbar.set)
        self.apply_palette(kit.palette, kit.font("small", mono=mono))

    def apply_palette(self, palette: ThemePalette, font) -> None:
        apply_text_widget_theme(self.text, palette)
        try:
            self.text.configure(font=font, padx=10, pady=8)
        except Exception:
            pass

    def configure(self, **kwargs: Any):  # type: ignore[override]
        """Text-only options (state, font, wrap) are forwarded to the inner Text."""
        text_options = {key: kwargs.pop(key) for key in ("state", "font", "wrap") if key in kwargs}
        if text_options and getattr(self, "text", None) is not None:
            self.text.configure(**text_options)
        if kwargs:
            return super().configure(**kwargs)
        return None

    config = configure

    def __getattr__(self, name: str):
        if name in LogBox._TEXT_METHODS:
            return getattr(self.__dict__["text"], name)
        raise AttributeError(name)


class ProgressBar(ctk.CTkProgressBar):
    """CTkProgressBar that also speaks the ttk.Progressbar API.

    The conversion/OCR/PDF workers all drive progress with ``bar["value"] = 0..100``
    (ttk's convention). Supporting that here keeps roughly twenty worker call sites
    unchanged; ``set(0..1)`` still works for new code.
    """

    def __init__(self, *args: Any, maximum: float = 100.0, animator: "Animator | None" = None, **kwargs: Any) -> None:
        kwargs.pop("mode", None)  # ttk-ism; CTk has determinate/indeterminate elsewhere
        super().__init__(*args, **kwargs)
        self._maximum = float(maximum) or 100.0
        self._animator = animator
        self.set(0)

    def set_smooth(self, value: float) -> None:
        """Ease toward ``value`` (0..1) so batch progress glides rather than jumps."""
        target = max(0.0, min(1.0, float(value)))
        if self._animator is None:
            self.set(target)
            return
        start = self.get()
        # Jumping backwards (a new run starting) should be instant.
        if target < start:
            self._animator.cancel(f"progress:{self}")
            self.set(target)
            return
        self._animator.tween(
            f"progress:{self}", 180, lambda p: self.set(start + (target - start) * p)
        )

    def __setitem__(self, key: str, value: Any) -> None:
        if key == "value":
            self.set_smooth(float(value) / self._maximum)
        elif key == "maximum":
            self._maximum = float(value) or 100.0
        else:
            self.configure(**{key: value})

    def __getitem__(self, key: str) -> Any:
        if key == "value":
            return self.get() * self._maximum
        if key == "maximum":
            return self._maximum
        return self.cget(key)


class ButtonBar(ctk.CTkFrame):
    """Wrapping row of buttons, API-compatible with the old FlowButtonBar.

    ``add`` returns the widget so call sites can keep doing
    ``self.some_button = bar.add(kit.button(...))``.
    """

    def __init__(self, parent: tk.Misc, kit: "UIKit", *, columns: int = 0, **kwargs: Any) -> None:
        super().__init__(parent, fg_color="transparent", corner_radius=0, **kwargs)
        self.kit = kit
        self.max_columns = max(0, int(columns))
        self._items: list[tk.Widget] = []

    def add(self, widget: tk.Widget) -> tk.Widget:
        if widget not in self._items:
            self._items.append(widget)
            self._place(widget, len(self._items) - 1)
        return widget

    def button(self, text: str, command, *, variant: str = "secondary", **kwargs: Any):
        return self.add(self.kit.button(self, text, command, variant=variant, **kwargs))

    def _place(self, widget: tk.Widget, index: int) -> None:
        gap = self.kit.pad("sm")
        if self.max_columns:
            row, column = divmod(index, self.max_columns)
        else:
            row, column = 0, index
        widget.grid(
            row=row,
            column=column,
            sticky="ew",
            padx=(0 if column == 0 else gap, 0),
            pady=(0 if row == 0 else gap, 0),
        )
        self.grid_columnconfigure(column, weight=1)

    def clear(self) -> None:
        for widget in list(self._items):
            try:
                widget.destroy()
            except Exception:
                pass
        self._items.clear()


class GradientBackdrop:
    """Paints the palette's gradient onto a tk.Canvas and keeps it sized.

    Tk has no compositing, so a real ``backdrop-filter`` blur is impossible. What
    *is* possible -- and pixel-identical for a fixed backdrop -- is to paint the
    gradient here and give the cards on top the exact colour that white-over-
    gradient would produce. ui_theme's ``surface``/``surface_alt``/``border`` are
    precomputed that way.

    The rendered image is cached on the kit and shared by every page, so opening
    eleven pages costs one bitmap rather than eleven.
    """

    def __init__(self, canvas: tk.Canvas, kit: "UIKit") -> None:
        self.canvas = canvas
        self.kit = kit
        self._image_id: int | None = None
        self._applied: tuple[int, int, str] | None = None
        canvas.bind("<Configure>", self._on_resize, add="+")

    def _on_resize(self, event) -> None:
        self.render(event.width, event.height)

    def render(self, width: int, height: int) -> None:
        photo, key = self.kit.gradient_image(width, height)
        if photo is None:
            try:
                self.canvas.configure(background=self.kit.palette.gradient_bottom)
            except Exception:
                pass
            return
        if key == self._applied:
            return
        self._applied = key
        try:
            if self._image_id is None:
                self._image_id = self.canvas.create_image(0, 0, anchor="nw", image=photo)
                self.canvas.tag_lower(self._image_id)
            else:
                self.canvas.itemconfigure(self._image_id, image=photo)
            self.canvas.configure(background=self.kit.palette.gradient_bottom)
        except Exception:
            return

    def refresh(self) -> None:
        """Force a repaint, e.g. after the palette changed."""
        self._applied = None
        try:
            self.render(self.canvas.winfo_width(), self.canvas.winfo_height())
        except Exception:
            return


class InfoDot(ctk.CTkLabel):
    """A small ⓘ that reveals its explanation on hover.

    Used instead of printing a paragraph of help next to every control: the screen
    stays clean for people who already know the app, and the explanation is one
    hover away for people who don't.
    """

    def __init__(self, parent: tk.Misc, kit: "UIKit", text: str, **kwargs: Any) -> None:
        super().__init__(
            parent,
            text="ⓘ",
            text_color=kit.palette.text_muted,
            fg_color="transparent",
            font=kit.font("body"),
            cursor="hand2" if tk.TkVersion >= 8.6 else "",
            **kwargs,
        )
        self.kit = kit
        self.help_text = text
        self._tooltip = Tooltip(self, text, kit)
        self.bind("<Enter>", lambda _e: self.configure(text_color=kit.palette.accent), add="+")
        self.bind("<Leave>", lambda _e: self.configure(text_color=kit.palette.text_muted), add="+")

    def set_help(self, text: str) -> None:
        self.help_text = text
        self._tooltip.text = text


def lerp_color(start: str, end: str, t: float) -> str:
    """Blend two hex colours. Used by the animation helpers."""
    t = max(0.0, min(1.0, t))
    a = ui_theme._hex_to_rgb(start)
    b = ui_theme._hex_to_rgb(end)
    return ui_theme._rgb_to_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


def ease_out(t: float) -> float:
    """cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle."""
    return 1 - pow(1 - max(0.0, min(1.0, t)), 3)


class Animator:
    """Frame-based tweening on Tk's event loop.

    Every animation checks ``reduced_motion`` first and, when set, jumps straight
    to the final value -- motion is decoration, never the only signal.
    """

    FRAME_MS = 16  # ~60fps

    def __init__(self, widget: tk.Misc, *, reduced_motion=None) -> None:
        self.widget = widget
        self._reduced_motion = reduced_motion
        self._jobs: dict[str, str] = {}

    def _skip(self) -> bool:
        try:
            return bool(self._reduced_motion()) if callable(self._reduced_motion) else False
        except Exception:
            return False

    def cancel(self, key: str) -> None:
        job = self._jobs.pop(key, None)
        if job:
            try:
                self.widget.after_cancel(job)
            except Exception:
                pass

    def tween(self, key: str, duration_ms: int, apply, *, on_done=None) -> None:
        """Call ``apply(progress)`` with eased 0→1 until ``duration_ms`` elapses."""
        self.cancel(key)
        if self._skip() or duration_ms <= 0:
            apply(1.0)
            if on_done:
                on_done()
            return

        steps = max(1, duration_ms // self.FRAME_MS)

        def frame(step: int = 0) -> None:
            progress = ease_out(step / steps)
            try:
                apply(progress)
            except Exception:
                self._jobs.pop(key, None)
                return
            if step >= steps:
                self._jobs.pop(key, None)
                if on_done:
                    on_done()
                return
            self._jobs[key] = self.widget.after(self.FRAME_MS, frame, step + 1)

        frame()

    def slide_in(self, target: tk.Misc, *, distance: int = 14, duration_ms: int = 220) -> None:
        """Nudge a page up into place — the app's standard entrance."""
        info = {}
        try:
            info = target.grid_info()
        except Exception:
            return
        if not info:
            return
        base_padx = info.get("padx", 0)

        def apply(progress: float) -> None:
            offset = int(round(distance * (1 - progress)))
            try:
                target.grid_configure(pady=(offset, 0), padx=base_padx)
            except Exception:
                pass

        self.tween(f"slide:{target}", duration_ms, apply)

    def pulse_border(self, widget, *, from_color: str, to_color: str, duration_ms: int = 260) -> None:
        """Briefly flash a widget's border — used to draw the eye to a result."""
        def apply(progress: float) -> None:
            # out and back, so the colour lands where it started
            t = progress * 2 if progress < 0.5 else (1 - progress) * 2
            widget.configure(border_color=lerp_color(from_color, to_color, t))

        self.tween(f"pulse:{widget}", duration_ms, apply)

    def count_to(self, variable: tk.StringVar, end: float, *, duration_ms: int = 400, fmt="{:.0f}") -> None:
        """Roll a number up to its new value, Apple-Weather style."""
        try:
            start = float(str(variable.get()).replace(",", "") or 0)
        except (TypeError, ValueError):
            start = 0.0
        if start == end:
            variable.set(fmt.format(end))
            return
        self.tween(
            f"count:{variable}",
            duration_ms,
            lambda p: variable.set(fmt.format(start + (end - start) * p)),
        )
