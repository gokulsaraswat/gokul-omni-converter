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


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    if len(value) != 6:
        return (0, 0, 0)
    return tuple(int(value[index:index + 2], 16) for index in range(0, 6, 2))


def _rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    red, green, blue = (max(0, min(255, int(channel))) for channel in rgb)
    return f"#{red:02x}{green:02x}{blue:02x}"


def _mix(color_a: str, color_b: str, ratio: float) -> str:
    ratio = max(0.0, min(1.0, float(ratio)))
    rgb_a = _hex_to_rgb(color_a)
    rgb_b = _hex_to_rgb(color_b)
    blended = tuple(round(rgb_a[index] * (1.0 - ratio) + rgb_b[index] * ratio) for index in range(3))
    return _rgb_to_hex(blended)


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
    header_bg = _mix(palette.root_bg, palette.surface, 0.55)
    footer_bg = _mix(palette.root_bg, palette.surface, 0.42)
    surface_hover_soft = _mix(palette.surface_alt, palette.accent, 0.10)
    surface_hover = _mix(palette.surface_alt, palette.accent, 0.18)
    surface_pressed = _mix(palette.surface_alt, palette.accent, 0.28)
    sidebar_hover_soft = _mix(palette.sidebar, palette.accent, 0.12)
    sidebar_hover = _mix(palette.sidebar, palette.accent, 0.20)
    sidebar_pressed = _mix(palette.sidebar, palette.accent, 0.32)
    border_hover_soft = _mix(palette.border, palette.accent, 0.32)
    border_hover = _mix(palette.border, palette.accent, 0.52)
    border_pressed = _mix(palette.border, palette.accent, 0.72)
    footer_hover_soft = _mix(footer_bg, palette.accent, 0.12)
    footer_hover = _mix(footer_bg, palette.accent, 0.20)
    footer_pressed = _mix(footer_bg, palette.accent, 0.30)
    accent_soft = _mix(palette.accent, palette.surface_alt, 0.12)
    accent_pressed = _mix(palette.accent_hover, palette.root_bg, 0.14)
    input_hover = _mix(palette.input_bg, palette.accent, 0.12)
    tree_heading_hover = _mix(palette.surface_alt, palette.accent, 0.14)
    scrollbar_hover = _mix(palette.surface_alt, palette.accent, 0.20)

    style.configure("TFrame", background=palette.root_bg)
    style.configure("Surface.TFrame", background=palette.surface)
    style.configure("Sidebar.TFrame", background=palette.sidebar)
    style.configure("Card.TFrame", background=palette.surface)
    style.configure("Hero.TFrame", background=palette.surface)
    style.configure("Header.TFrame", background=header_bg)
    style.configure("Footer.TFrame", background=footer_bg)
    style.configure("TLabel", background=palette.root_bg, foreground=palette.text)
    style.configure("Surface.TLabel", background=palette.surface, foreground=palette.text)
    style.configure("Sidebar.TLabel", background=palette.sidebar, foreground=palette.text_muted)
    style.configure("Muted.TLabel", background=palette.surface, foreground=palette.text_muted)
    style.configure("SidebarMuted.TLabel", background=palette.sidebar, foreground=palette.text_muted)
    style.configure("Title.TLabel", background=palette.root_bg, foreground=palette.text, font=("Segoe UI", title_size, "bold"))
    style.configure("Subtitle.TLabel", background=palette.root_bg, foreground=palette.text_muted, font=("Segoe UI", subtitle_size))
    style.configure("Footer.TLabel", background=footer_bg, foreground=palette.text_muted, font=("Segoe UI", _scaled(8 if compact else 9, scale)))
    style.configure("Logo.TLabel", background=header_bg, foreground=palette.text, font=("Segoe UI", _scaled(9 if compact else 10, scale), "bold"))
    style.configure("HeroTitle.TLabel", background=palette.surface, foreground=palette.text, font=("Segoe UI", hero_title_size, "bold"))
    style.configure("HeroBody.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", body_size))
    style.configure("CardTitle.TLabel", background=palette.surface, foreground=palette.text, font=("Segoe UI", _scaled(11 if compact else 12, scale), "bold"))
    style.configure("CardBody.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", body_size))
    style.configure("Eyebrow.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", _scaled(8 if compact else 9, scale), "bold"))
    style.configure("MetricValue.TLabel", background=palette.surface, foreground=palette.text, font=("Segoe UI", _scaled(13 if compact else 15, scale), "bold"))
    style.configure("MetricHint.TLabel", background=palette.surface, foreground=palette.text_muted, font=("Segoe UI", _scaled(8 if compact else 9, scale)))
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
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        padding=button_padding,
    )
    style.map(
        "TButton",
        background=[("active", surface_hover), ("pressed", surface_pressed)],
        bordercolor=[("active", border_hover), ("pressed", border_pressed)],
        foreground=[("disabled", palette.text_muted), ("active", palette.text)],
    )
    style.configure(
        "SoftHover.TButton",
        background=surface_hover_soft,
        foreground=palette.text,
        bordercolor=border_hover_soft,
        lightcolor=surface_hover_soft,
        darkcolor=surface_hover_soft,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        padding=button_padding,
    )
    style.configure(
        "Hover.TButton",
        background=surface_hover,
        foreground=palette.text,
        bordercolor=border_hover,
        lightcolor=surface_hover,
        darkcolor=surface_hover,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        padding=button_padding,
    )
    style.configure(
        "Pressed.TButton",
        background=surface_pressed,
        foreground=palette.text,
        bordercolor=border_pressed,
        lightcolor=surface_pressed,
        darkcolor=surface_pressed,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        padding=button_padding,
    )

    style.configure(
        "Primary.TButton",
        background=palette.accent,
        foreground=palette.accent_text,
        bordercolor=palette.accent,
        lightcolor=palette.accent,
        darkcolor=palette.accent,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent_hover,
        padding=primary_padding,
    )
    style.map(
        "Primary.TButton",
        background=[("active", palette.accent_hover), ("pressed", accent_pressed), ("disabled", palette.surface_alt)],
        bordercolor=[("active", palette.accent_hover), ("pressed", border_pressed)],
        foreground=[("disabled", palette.text_muted)],
    )
    style.configure(
        "SoftHover.Primary.TButton",
        background=accent_soft,
        foreground=palette.accent_text,
        bordercolor=palette.accent_hover,
        lightcolor=accent_soft,
        darkcolor=accent_soft,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent_hover,
        padding=primary_padding,
    )
    style.configure(
        "Hover.Primary.TButton",
        background=palette.accent_hover,
        foreground=palette.accent_text,
        bordercolor=palette.accent_hover,
        lightcolor=palette.accent_hover,
        darkcolor=palette.accent_hover,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent_hover,
        padding=primary_padding,
    )
    style.configure(
        "Pressed.Primary.TButton",
        background=accent_pressed,
        foreground=palette.accent_text,
        bordercolor=border_pressed,
        lightcolor=accent_pressed,
        darkcolor=accent_pressed,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent_hover,
        padding=primary_padding,
    )

    style.configure(
        "Small.TButton",
        background=footer_bg,
        foreground=palette.text,
        bordercolor=palette.border,
        lightcolor=footer_bg,
        darkcolor=footer_bg,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        padding=_scaled_padding((8, 3) if compact else (9, 4), scale),
        font=("Segoe UI", _scaled(8 if compact else 9, scale)),
    )
    style.map(
        "Small.TButton",
        background=[("active", footer_hover), ("pressed", footer_pressed)],
        bordercolor=[("active", border_hover), ("pressed", border_pressed)],
        foreground=[("disabled", palette.text_muted), ("active", palette.text)],
    )
    small_padding = _scaled_padding((8, 3) if compact else (9, 4), scale)
    small_font = ("Segoe UI", _scaled(8 if compact else 9, scale))
    style.configure("SoftHover.Small.TButton", background=footer_hover_soft, foreground=palette.text, bordercolor=border_hover_soft, lightcolor=footer_hover_soft, darkcolor=footer_hover_soft, relief="flat", focusthickness=1, focuscolor=palette.accent, padding=small_padding, font=small_font)
    style.configure("Hover.Small.TButton", background=footer_hover, foreground=palette.text, bordercolor=border_hover, lightcolor=footer_hover, darkcolor=footer_hover, relief="flat", focusthickness=1, focuscolor=palette.accent, padding=small_padding, font=small_font)
    style.configure("Pressed.Small.TButton", background=footer_pressed, foreground=palette.text, bordercolor=border_pressed, lightcolor=footer_pressed, darkcolor=footer_pressed, relief="flat", focusthickness=1, focuscolor=palette.accent, padding=small_padding, font=small_font)

    style.configure(
        "Nav.TButton",
        background=palette.sidebar,
        foreground=palette.text_muted,
        bordercolor=palette.sidebar,
        lightcolor=palette.sidebar,
        darkcolor=palette.sidebar,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent,
        anchor="w",
        padding=nav_padding,
    )
    style.map(
        "Nav.TButton",
        background=[("active", sidebar_hover), ("pressed", sidebar_pressed)],
        bordercolor=[("active", border_hover), ("pressed", border_pressed)],
        foreground=[("active", palette.text)],
    )
    style.configure("SoftHover.Nav.TButton", background=sidebar_hover_soft, foreground=palette.text, bordercolor=border_hover_soft, lightcolor=sidebar_hover_soft, darkcolor=sidebar_hover_soft, relief="flat", focusthickness=1, focuscolor=palette.accent, anchor="w", padding=nav_padding)
    style.configure("Hover.Nav.TButton", background=sidebar_hover, foreground=palette.text, bordercolor=border_hover, lightcolor=sidebar_hover, darkcolor=sidebar_hover, relief="flat", focusthickness=1, focuscolor=palette.accent, anchor="w", padding=nav_padding)
    style.configure("Pressed.Nav.TButton", background=sidebar_pressed, foreground=palette.text, bordercolor=border_pressed, lightcolor=sidebar_pressed, darkcolor=sidebar_pressed, relief="flat", focusthickness=1, focuscolor=palette.accent, anchor="w", padding=nav_padding)

    style.configure(
        "NavActive.TButton",
        background=palette.accent,
        foreground=palette.accent_text,
        bordercolor=palette.accent,
        lightcolor=palette.accent,
        darkcolor=palette.accent,
        relief="flat",
        focusthickness=1,
        focuscolor=palette.accent_hover,
        anchor="w",
        padding=nav_padding,
    )
    style.map(
        "NavActive.TButton",
        background=[("active", palette.accent_hover), ("pressed", accent_pressed)],
        bordercolor=[("active", palette.accent_hover), ("pressed", border_pressed)],
        foreground=[("active", palette.accent_text)],
    )
    style.configure("SoftHover.NavActive.TButton", background=accent_soft, foreground=palette.accent_text, bordercolor=palette.accent_hover, lightcolor=accent_soft, darkcolor=accent_soft, relief="flat", focusthickness=1, focuscolor=palette.accent_hover, anchor="w", padding=nav_padding)
    style.configure("Hover.NavActive.TButton", background=palette.accent_hover, foreground=palette.accent_text, bordercolor=palette.accent_hover, lightcolor=palette.accent_hover, darkcolor=palette.accent_hover, relief="flat", focusthickness=1, focuscolor=palette.accent_hover, anchor="w", padding=nav_padding)
    style.configure("Pressed.NavActive.TButton", background=accent_pressed, foreground=palette.accent_text, bordercolor=border_pressed, lightcolor=accent_pressed, darkcolor=accent_pressed, relief="flat", focusthickness=1, focuscolor=palette.accent_hover, anchor="w", padding=nav_padding)

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
        fieldbackground=[("readonly", palette.input_bg), ("active", input_hover)],
        foreground=[("readonly", palette.text)],
        selectbackground=[("readonly", palette.input_bg)],
        selectforeground=[("readonly", palette.text)],
        bordercolor=[("focus", palette.accent), ("active", border_hover_soft)],
    )

    style.configure("TCheckbutton", background=palette.surface, foreground=palette.text)
    style.map(
        "TCheckbutton",
        background=[("active", palette.surface)],
        indicatorcolor=[("active", palette.accent)],
        foreground=[("active", palette.text)],
    )

    style.configure(
        "TLabelframe",
        background=palette.surface,
        foreground=palette.text,
        bordercolor=palette.border,
        relief="solid",
        borderwidth=1,
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
    style.map("Treeview.Heading", background=[("active", tree_heading_hover)], bordercolor=[("active", border_hover)])

    style.configure("TScrollbar", background=palette.surface_alt, troughcolor=palette.surface, bordercolor=palette.border, arrowcolor=palette.text_muted)
    style.map("TScrollbar", background=[("active", scrollbar_hover), ("pressed", surface_pressed)], arrowcolor=[("active", palette.text), ("pressed", palette.accent_text)])
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
