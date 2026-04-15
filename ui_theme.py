from __future__ import annotations

import os
import tkinter as tk
from dataclasses import dataclass, replace
from tkinter import font as tkfont, ttk


@dataclass(frozen=True)
class ThemePalette:
    name: str
    root_bg: str
    surface: str
    surface_alt: str
    sidebar: str
    border: str
    text: str
    text_muted: str
    accent: str
    accent_hover: str
    accent_text: str
    selection: str
    input_bg: str
    code_bg: str
    success: str
    warning: str
    danger: str


DARK_THEME = ThemePalette(
    name="dark",
    root_bg="#0f172a",
    surface="#111827",
    surface_alt="#1f2937",
    sidebar="#0b1220",
    border="#263245",
    text="#f8fafc",
    text_muted="#cbd5e1",
    accent="#3b82f6",
    accent_hover="#2563eb",
    accent_text="#ffffff",
    selection="#1d4ed8",
    input_bg="#0b1324",
    code_bg="#0b1324",
    success="#22c55e",
    warning="#f59e0b",
    danger="#ef4444",
)

LIGHT_THEME = ThemePalette(
    name="light",
    root_bg="#f3f4f6",
    surface="#ffffff",
    surface_alt="#e5e7eb",
    sidebar="#e5e7eb",
    border="#d1d5db",
    text="#111827",
    text_muted="#4b5563",
    accent="#2563eb",
    accent_hover="#1d4ed8",
    accent_text="#ffffff",
    selection="#bfdbfe",
    input_bg="#ffffff",
    code_bg="#f8fafc",
    success="#15803d",
    warning="#b45309",
    danger="#b91c1c",
)

HIGH_CONTRAST_DARK_THEME = replace(
    DARK_THEME,
    name="dark-high-contrast",
    root_bg="#000000",
    surface="#04070d",
    surface_alt="#0d1523",
    sidebar="#000000",
    border="#5b728f",
    text="#ffffff",
    text_muted="#dfe9f6",
    accent="#5ea7ff",
    accent_hover="#7db9ff",
    selection="#1f6feb",
    input_bg="#000000",
    code_bg="#000000",
    success="#33dd77",
    warning="#ffb020",
    danger="#ff6363",
)

HIGH_CONTRAST_LIGHT_THEME = replace(
    LIGHT_THEME,
    name="light-high-contrast",
    root_bg="#ffffff",
    surface="#ffffff",
    surface_alt="#eef2f7",
    sidebar="#f5f7fb",
    border="#44556a",
    text="#000000",
    text_muted="#1f2937",
    accent="#0047cc",
    accent_hover="#0039a3",
    accent_text="#ffffff",
    selection="#cfe0ff",
    input_bg="#ffffff",
    code_bg="#ffffff",
    success="#0f7a36",
    warning="#945200",
    danger="#a40f1a",
)


def detect_system_theme() -> str:
    if os.name == "nt":
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if int(value) == 1 else "dark"
        except Exception:
            return "dark"
    return "light"


def resolve_palette(theme_choice: str, high_contrast: bool = False) -> ThemePalette:
    choice = (theme_choice or "dark").strip().lower()
    if choice == "system":
        choice = detect_system_theme()
    if high_contrast:
        return HIGH_CONTRAST_DARK_THEME if choice == "dark" else HIGH_CONTRAST_LIGHT_THEME
    return DARK_THEME if choice == "dark" else LIGHT_THEME


def _scaled(value: int, scale: float) -> int:
    return max(1, int(round(value * scale)))


def _scaled_padding(value: tuple[int, int], scale: float) -> tuple[int, int]:
    return (_scaled(value[0], scale), _scaled(value[1], scale))


def _configure_fonts(compact: bool = False, scale: float = 1.0) -> None:
    scale = min(max(float(scale or 1.0), 0.85), 1.6)
    base_size = _scaled(9 if compact else 10, scale)
    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=base_size)
    text_font = tkfont.nametofont("TkTextFont")
    text_font.configure(size=base_size)
    heading_font = tkfont.nametofont("TkHeadingFont")
    heading_font.configure(size=base_size, weight="bold")
    try:
        fixed_font = tkfont.nametofont("TkFixedFont")
        fixed_font.configure(size=base_size)
    except Exception:
        pass
    try:
        menu_font = tkfont.nametofont("TkMenuFont")
        menu_font.configure(size=base_size)
    except Exception:
        pass


def apply_ttk_theme(root: tk.Misc, palette: ThemePalette, compact: bool = False, scale: float = 1.0) -> ttk.Style:
    scale = min(max(float(scale or 1.0), 0.85), 1.6)
    _configure_fonts(compact=compact, scale=scale)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(background=palette.root_bg)
    root.option_add("*tearOff", False)

    title_size = _scaled(16 if compact else 18, scale)
    subtitle_size = _scaled(9 if compact else 10, scale)
    hero_title_size = _scaled(14 if compact else 16, scale)
    body_size = _scaled(9 if compact else 10, scale)
    button_padding = _scaled_padding((8, 5) if compact else (10, 7), scale)
    primary_padding = _scaled_padding((10, 6) if compact else (12, 8), scale)
    nav_padding = _scaled_padding((12, 8) if compact else (14, 10), scale)
    entry_padding = _scaled_padding((7, 4) if compact else (8, 6), scale)
    combo_padding = _scaled_padding((5, 4) if compact else (6, 6), scale)
    frame_padding = _scaled(10 if compact else 12, scale)
    row_height = _scaled(24 if compact else 28, scale)

    style.configure("TFrame", background=palette.root_bg)
    style.configure("Surface.TFrame", background=palette.surface)
    style.configure("Sidebar.TFrame", background=palette.sidebar)
    style.configure("Card.TFrame", background=palette.surface)
    style.configure("Hero.TFrame", background=palette.surface)
    style.configure("TLabel", background=palette.root_bg, foreground=palette.text)
    style.configure("Surface.TLabel", background=palette.surface, foreground=palette.text)
    style.configure("Sidebar.TLabel", background=palette.sidebar, foreground=palette.text_muted)
    style.configure("Muted.TLabel", background=palette.surface, foreground=palette.text_muted)
    style.configure("SidebarMuted.TLabel", background=palette.sidebar, foreground=palette.text_muted)
    style.configure("Title.TLabel", background=palette.root_bg, foreground=palette.text, font=("Segoe UI", title_size, "bold"))
    style.configure("Subtitle.TLabel", background=palette.root_bg, foreground=palette.text_muted, font=("Segoe UI", subtitle_size))
    style.configure("HeroTitle.TLabel", background=palette.surface, foreground=palette.text, font=("Segoe UI", hero_title_size, "bold"))
    style.configure("HeroBody.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", body_size))
    style.configure("CardTitle.TLabel", background=palette.surface, foreground=palette.text, font=("Segoe UI", _scaled(11 if compact else 12, scale), "bold"))
    style.configure("CardBody.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", body_size))
    style.configure("StatusGood.TLabel", background=palette.surface, foreground=palette.success, font=("Segoe UI", body_size))
    style.configure("StatusWarn.TLabel", background=palette.surface, foreground=palette.warning, font=("Segoe UI", body_size))
    style.configure("StatusBad.TLabel", background=palette.surface, foreground=palette.danger, font=("Segoe UI", body_size))

    style.configure(
        "TButton",
        background=palette.surface_alt,
        foreground=palette.text,
        bordercolor=palette.border,
        lightcolor=palette.surface_alt,
        darkcolor=palette.surface_alt,
        padding=button_padding,
    )
    style.map(
        "TButton",
        background=[("active", palette.surface_alt), ("pressed", palette.surface_alt)],
        foreground=[("disabled", palette.text_muted), ("active", palette.text)],
    )

    style.configure(
        "Primary.TButton",
        background=palette.accent,
        foreground=palette.accent_text,
        bordercolor=palette.accent,
        lightcolor=palette.accent,
        darkcolor=palette.accent,
        padding=primary_padding,
    )
    style.map(
        "Primary.TButton",
        background=[("active", palette.accent_hover), ("pressed", palette.accent_hover), ("disabled", palette.surface_alt)],
        foreground=[("disabled", palette.text_muted)],
    )

    style.configure(
        "Nav.TButton",
        background=palette.sidebar,
        foreground=palette.text_muted,
        bordercolor=palette.sidebar,
        lightcolor=palette.sidebar,
        darkcolor=palette.sidebar,
        anchor="w",
        padding=nav_padding,
    )
    style.map(
        "Nav.TButton",
        background=[("active", palette.surface_alt), ("pressed", palette.surface_alt)],
        foreground=[("active", palette.text)],
    )
    style.configure(
        "NavActive.TButton",
        background=palette.accent,
        foreground=palette.accent_text,
        bordercolor=palette.accent,
        lightcolor=palette.accent,
        darkcolor=palette.accent,
        anchor="w",
        padding=nav_padding,
    )
    style.map(
        "NavActive.TButton",
        background=[("active", palette.accent_hover), ("pressed", palette.accent_hover)],
        foreground=[("active", palette.accent_text)],
    )

    style.configure(
        "TEntry",
        fieldbackground=palette.input_bg,
        foreground=palette.text,
        bordercolor=palette.border,
        insertcolor=palette.text,
        padding=entry_padding,
    )
    style.map("TEntry", bordercolor=[("focus", palette.accent)])
    style.configure(
        "TCombobox",
        fieldbackground=palette.input_bg,
        foreground=palette.text,
        background=palette.surface_alt,
        bordercolor=palette.border,
        arrowsize=_scaled(14, scale),
        padding=combo_padding,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", palette.input_bg)],
        foreground=[("readonly", palette.text)],
        selectbackground=[("readonly", palette.input_bg)],
        selectforeground=[("readonly", palette.text)],
    )

    style.configure("TCheckbutton", background=palette.surface, foreground=palette.text)
    style.map("TCheckbutton", background=[("active", palette.surface)], foreground=[("active", palette.text)])

    style.configure(
        "TLabelframe",
        background=palette.surface,
        foreground=palette.text,
        bordercolor=palette.border,
        relief="solid",
        padding=frame_padding,
    )
    style.configure("TLabelframe.Label", background=palette.surface, foreground=palette.text, font=("Segoe UI", body_size, "bold"))

    style.configure(
        "Horizontal.TProgressbar",
        troughcolor=palette.surface_alt,
        bordercolor=palette.surface_alt,
        background=palette.accent,
        lightcolor=palette.accent,
        darkcolor=palette.accent,
    )

    style.configure(
        "Treeview",
        background=palette.input_bg,
        fieldbackground=palette.input_bg,
        foreground=palette.text,
        bordercolor=palette.border,
        rowheight=row_height,
    )
    style.map("Treeview", background=[("selected", palette.selection)], foreground=[("selected", palette.accent_text)])
    style.configure(
        "Treeview.Heading",
        background=palette.surface_alt,
        foreground=palette.text,
        relief="flat",
        bordercolor=palette.border,
        font=("Segoe UI", body_size, "bold"),
    )
    style.map("Treeview.Heading", background=[("active", palette.surface_alt)])

    style.configure("TScrollbar", background=palette.surface_alt, troughcolor=palette.surface, bordercolor=palette.border)
    style.configure("TSeparator", background=palette.border)

    return style


def apply_menu_theme(menu: tk.Menu, palette: ThemePalette) -> None:
    menu.configure(
        background=palette.surface,
        foreground=palette.text,
        activebackground=palette.accent,
        activeforeground=palette.accent_text,
        selectcolor=palette.accent,
        relief="flat",
        borderwidth=0,
    )


def apply_text_widget_theme(widget: tk.Text | tk.Listbox, palette: ThemePalette) -> None:
    common = {
        "background": palette.code_bg if isinstance(widget, tk.Text) else palette.input_bg,
        "foreground": palette.text,
        "highlightbackground": palette.border,
        "highlightcolor": palette.accent,
        "selectbackground": palette.selection,
        "selectforeground": palette.accent_text,
    }
    widget.configure(**common)
    try:
        widget.configure(insertbackground=palette.text)
    except tk.TclError:
        pass


def apply_treeview_tag_colors(tree: ttk.Treeview, palette: ThemePalette) -> None:
    tree.tag_configure("success", foreground=palette.success)
    tree.tag_configure("warn", foreground=palette.warning)
    tree.tag_configure("error", foreground=palette.danger)
