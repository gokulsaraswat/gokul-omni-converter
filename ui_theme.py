from __future__ import annotations

import os
import subprocess
import sys
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
    # Backdrop gradient stops, painted behind the shell and page bodies.
    gradient_top: str = "#141a3d"
    gradient_mid: str = "#241a4a"
    gradient_bottom: str = "#0d1030"
    #: Light inner edge along the top of a glass card (CSS: inset 0 1px 0 white/20%).
    glass_highlight: str = "#4e5169"
    #: Tint used for coloured drop shadows instead of plain black.
    shadow_tint: str = "#5b63ff"


DARK_THEME = ThemePalette(
    name="dark",
    # Deep translucent navy/violet field, not flat black.
    root_bg="#161a3a",
    # Card colours are the literal result of compositing white over root_bg at the
    # alphas a CSS glass panel would use -- Tk cannot blur, but on a fixed backdrop
    # the resulting pixels are identical.
    surface="#292c4a",       # white 8%
    surface_alt="#323552",   # white 12%
    sidebar="#0e1126",
    border="#3b3f5a",        # white 16%
    text="#ffffff",
    text_muted="#a6a8b4",    # white 62%
    accent="#6d5efc",
    accent_hover="#8175ff",
    accent_text="#ffffff",
    selection="#3a3480",
    input_bg="#1e2142",
    code_bg="#12152e",
    success="#4ade80",
    warning="#fbbf24",
    danger="#fb7185",
    gradient_top="#1b2350",
    gradient_mid="#2a1f52",
    gradient_bottom="#0b0e26",
    glass_highlight="#4e5169",
    shadow_tint="#6d5efc",
)

LIGHT_THEME = ThemePalette(
    name="light",
    root_bg="#e7ecfb",
    surface="#f8f9fe",       # white 70%
    surface_alt="#eef1fb",
    sidebar="#dde4f7",
    border="#c9d2ee",
    text="#10142c",
    text_muted="#565c7a",
    accent="#5b4bd6",
    accent_hover="#4a3ac2",
    accent_text="#ffffff",
    selection="#c7cdfa",
    input_bg="#ffffff",
    code_bg="#f3f5fd",
    success="#15803d",
    warning="#b45309",
    danger="#be123c",
    gradient_top="#dbe7ff",
    gradient_mid="#e9dcff",
    gradient_bottom="#f6f8ff",
    glass_highlight="#ffffff",
    shadow_tint="#8b93d6",
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
    # The bright accent needs dark text on it; white would only reach 2.5:1.
    accent_text="#00040d",
    selection="#1f6feb",
    input_bg="#000000",
    code_bg="#000000",
    success="#33dd77",
    warning="#ffb020",
    danger="#ff6363",
    gradient_top="#000000",
    gradient_mid="#000000",
    gradient_bottom="#000000",
    glass_highlight="#5b728f",
    shadow_tint="#000000",
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
    gradient_top="#ffffff",
    gradient_mid="#ffffff",
    gradient_bottom="#ffffff",
    glass_highlight="#ffffff",
    shadow_tint="#44556a",
)


def detect_system_theme() -> str:
    if sys.platform == "darwin":
        # AppleInterfaceStyle only exists while dark mode is on; `defaults` exits
        # non-zero in light mode, which is the signal we want.
        try:
            result = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True,
                text=True,
                timeout=2,
            )
            return "dark" if "dark" in (result.stdout or "").strip().lower() else "light"
        except Exception:
            return "light"
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
    style.configure("Logo.TLabel", background=header_bg, foreground=palette.text, font=("Segoe UI", _scaled(10 if compact else 11, scale), "bold"))
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


# ---------------------------------------------------------------------------
# CustomTkinter design system
#
# Single source of truth for the CTk widget layer: spacing, radii, fonts, and
# per-variant widget colours, all derived from the ThemePalette values above so
# ttk-only widgets (Treeview, Listbox, Text) stay visually consistent with the
# CTk widgets around them. Pages must read these tokens rather than hardcoding.
# ---------------------------------------------------------------------------

# 4px base scale.
SPACING = {"xs": 4, "sm": 8, "md": 14, "lg": 20, "xl": 28, "xxl": 40}

RADIUS = {"none": 0, "sm": 8, "md": 12, "lg": 16, "xl": 20, "pill": 100}

# Preferred UI faces per platform, most-preferred first. resolve_font_family()
# picks the first one Tk actually has, so a missing font never renders as a
# fallback serif.
#: Thin display faces for large hero numerals (Tk only knows normal/bold, so a
#: lighter *family* is the only way to get a 200-300 weight).
THIN_FONT_STACK = (
    "SF Pro Display Thin",
    "HelveticaNeue-Thin",
    "Helvetica Neue Thin",
    "Inter Light",
    "Segoe UI Light",
    "Roboto Thin",
)

FONT_STACK = (
    "SF Pro Text",
    "SF Pro Display",
    "Inter",
    "Segoe UI Variable Text",
    "Segoe UI",
    ".AppleSystemUIFont",
    "Helvetica Neue",
    "Ubuntu",
    "Cantarell",
    "Noto Sans",
    "DejaVu Sans",
)

MONO_FONT_STACK = (
    "JetBrains Mono",
    "Cascadia Mono",
    "Consolas",
    "SF Mono",
    "Menlo",
    "Ubuntu Mono",
    "DejaVu Sans Mono",
    "Courier New",
)

# Logical size names -> base point size at 100% scale, before compact/scale adjustment.
FONT_SIZES = {
    "hero": 40,
    "display": 26,
    "title": 17,
    "heading": 13,
    "subheading": 11,
    "body": 11,
    "small": 10,
    "tiny": 9,
}

_FONT_FAMILY_CACHE: dict[str, str] = {}


def resolve_font_family(root: tk.Misc | None = None, mono: bool = False, thin: bool = False) -> str:
    """Return the first font in the stack that this Tk install actually has."""
    cache_key = "mono" if mono else ("thin" if thin else "ui")
    cached = _FONT_FAMILY_CACHE.get(cache_key)
    if cached:
        return cached

    stack = MONO_FONT_STACK if mono else (THIN_FONT_STACK if thin else FONT_STACK)
    available: set[str] = set()
    try:
        available = {name.lower() for name in tkfont.families(root)}
    except Exception:
        available = set()

    chosen = ""
    for candidate in stack:
        if candidate.lower() in available:
            chosen = candidate
            break
    if not chosen and thin:
        # No thin face available; fall back to the regular UI family.
        chosen = resolve_font_family(root, mono=False)
    if not chosen:
        try:
            chosen = tkfont.nametofont("TkFixedFont" if mono else "TkDefaultFont").cget("family")
        except Exception:
            chosen = "Courier" if mono else "Helvetica"

    _FONT_FAMILY_CACHE[cache_key] = chosen
    return chosen


def reset_font_cache() -> None:
    """Drop the resolved-family cache (used by tests and by font-setting changes)."""
    _FONT_FAMILY_CACHE.clear()


def font_spec(
    role: str = "body",
    *,
    weight: str = "normal",
    compact: bool = False,
    scale: float = 1.0,
    mono: bool = False,
    thin: bool = False,
    root: tk.Misc | None = None,
) -> tuple[str, int, str]:
    """A (family, size, weight) tuple usable anywhere Tk accepts a font."""
    scale = min(max(float(scale or 1.0), 0.85), 1.6)
    base = FONT_SIZES.get(role, FONT_SIZES["body"])
    if compact:
        base = max(8, base - 1)
    return (resolve_font_family(root, mono=mono, thin=thin), _scaled(base, scale), weight)


def space(name: str, scale: float = 1.0) -> int:
    """Scaled spacing token. Unknown names fall back to the medium step."""
    return _scaled(SPACING.get(name, SPACING["md"]), min(max(float(scale or 1.0), 0.85), 1.6))


def is_dark(palette: ThemePalette) -> bool:
    return palette.name.startswith("dark")


def appearance_mode(palette: ThemePalette) -> str:
    """CustomTkinter's global light/dark switch for the given palette."""
    return "dark" if is_dark(palette) else "light"


def surface_hover(palette: ThemePalette) -> str:
    """Hover tint for neutral surfaces, lightening on dark and darkening on light."""
    return _mix(palette.surface_alt, palette.text, 0.10 if is_dark(palette) else 0.06)


def surface_pressed(palette: ThemePalette) -> str:
    return _mix(palette.surface_alt, palette.text, 0.18 if is_dark(palette) else 0.12)


def button_colors(palette: ThemePalette, variant: str = "primary") -> dict[str, str]:
    """fg/hover/text/border colours for a CTkButton of the given variant.

    Variants: primary (the one obvious action), secondary (neutral, repeatable),
    danger (destructive), ghost (low-emphasis / inline), success.
    """
    variant = (variant or "primary").strip().lower()
    if variant == "primary":
        return {
            "fg_color": palette.accent,
            "hover_color": palette.accent_hover,
            "text_color": palette.accent_text,
            "border_color": palette.accent,
        }
    if variant == "danger":
        return {
            "fg_color": palette.danger,
            "hover_color": _mix(palette.danger, "#000000", 0.18),
            "text_color": "#ffffff",
            "border_color": palette.danger,
        }
    if variant == "success":
        return {
            "fg_color": palette.success,
            "hover_color": _mix(palette.success, "#000000", 0.18),
            "text_color": "#ffffff",
            "border_color": palette.success,
        }
    if variant == "ghost":
        return {
            "fg_color": "transparent",
            "hover_color": surface_hover(palette),
            "text_color": palette.text_muted,
            "border_color": palette.border,
        }
    # secondary
    return {
        "fg_color": palette.surface_alt,
        "hover_color": surface_hover(palette),
        "text_color": palette.text,
        "border_color": palette.border,
    }


def entry_colors(palette: ThemePalette) -> dict[str, str]:
    return {
        "fg_color": palette.input_bg,
        "border_color": palette.border,
        "text_color": palette.text,
        "placeholder_text_color": palette.text_muted,
    }


def card_colors(palette: ThemePalette) -> dict[str, str]:
    return {"fg_color": palette.surface, "border_color": palette.border}


def apply_ctk_theme(palette: ThemePalette, *, compact: bool = False, scale: float = 1.0) -> None:
    """Point CustomTkinter's globals at the active palette.

    Widget colours are still passed explicitly by ui/kit.py -- this only covers the
    handful of places CTk paints from its own globals (dropdown menus, scaling).
    Import is local so ui_theme stays importable in headless/test contexts.
    """
    try:
        import customtkinter as ctk
    except Exception:
        return

    try:
        ctk.set_appearance_mode(appearance_mode(palette))
    except Exception:
        pass
    # Deliberately NOT calling ctk.set_widget_scaling(): this design system already
    # scales fonts (font_spec) and spacing (space()) by the same factor, so letting
    # CustomTkinter rescale as well would double-apply it -- and its global rescale
    # walks and redraws every widget in the app, which is very slow at this size.
