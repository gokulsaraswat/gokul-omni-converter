from __future__ import annotations

from pathlib import Path

import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageSequence, ImageTk

from engagement_core import SPLASH_AUTO_CLOSE_MS
from ui_theme import ThemePalette, apply_ttk_theme


class FirstLaunchSplashWindow(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        *,
        gif_path: Path,
        palette: ThemePalette,
        on_close,
        duration_ms: int = SPLASH_AUTO_CLOSE_MS,
        reduced_motion: bool = False,
        compact: bool = False,
        scale: float = 1.0,
    ) -> None:
        super().__init__(master)
        self.gif_path = Path(gif_path)
        self.palette = palette
        self.on_close = on_close
        self.duration_ms = max(800, int(duration_ms))
        self.reduced_motion = bool(reduced_motion)
        self.compact = compact
        self.scale = scale
        self.frames: list[ImageTk.PhotoImage] = []
        self.frame_index = 0
        self.fade_value = 0.0
        self._closing = False
        self._gif_after_id: str | None = None
        self._close_after_id: str | None = None

        self.title("Gokul Omni Convert Lite")
        self.geometry("620x360")
        self.resizable(False, False)
        self.transient(master)
        self.attributes("-topmost", True)
        try:
            self.attributes("-alpha", 1.0 if self.reduced_motion else 0.0)
            if self.reduced_motion:
                self.fade_value = 1.0
        except Exception:
            self.fade_value = 1.0
        self.configure(background=palette.root_bg)
        apply_ttk_theme(self, palette, compact=self.compact, scale=self.scale)

        card = ttk.Frame(self, style="Card.TFrame", padding=24)
        card.pack(fill="both", expand=True, padx=10, pady=10)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        ttk.Label(card, text="Welcome to Gokul Omni Convert Lite", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.image_label = ttk.Label(card, anchor="center")
        self.image_label.grid(row=1, column=0, sticky="nsew", pady=(14, 8))
        self.caption_label = ttk.Label(
            card,
            text="Configurable startup splash. Replace the GIF path later from Settings.",
            style="CardBody.TLabel",
            wraplength=540,
            justify="center",
        )
        self.caption_label.grid(row=2, column=0, sticky="ew")

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=3, column=0, sticky="e", pady=(16, 0))
        ttk.Button(actions, text="Skip", command=self.close_now).grid(row=0, column=0)

        self._load_frames()
        self._center_on_screen()
        self.protocol("WM_DELETE_WINDOW", self.close_now)

        if not self.reduced_motion:
            self._animate_fade_in()
            self._schedule_frame_advance()
        self._close_after_id = self.after(self.duration_ms, self.close_now)

    def _load_frames(self) -> None:
        if self.gif_path.exists():
            try:
                image = Image.open(self.gif_path)
                for frame in ImageSequence.Iterator(image):
                    copy = frame.convert("RGBA")
                    copy.thumbnail((540, 220))
                    self.frames.append(ImageTk.PhotoImage(copy))
                    if self.reduced_motion:
                        break
            except Exception:
                self.frames = []
        if self.frames:
            self.image_label.configure(image=self.frames[0], text="")
            self.caption_label.configure(text=f"Showing {self.gif_path.name}. Update the asset path in Settings any time.")
        else:
            self.image_label.configure(
                text="Splash asset missing or unreadable. The app will continue normally.",
                style="HeroBody.TLabel",
                justify="center",
            )

    def _center_on_screen(self) -> None:
        self.update_idletasks()
        width = self.winfo_width() or 620
        height = self.winfo_height() or 360
        x = max((self.winfo_screenwidth() - width) // 2, 0)
        y = max((self.winfo_screenheight() - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _animate_fade_in(self) -> None:
        if self._closing or self.reduced_motion:
            return
        try:
            next_value = min(self.fade_value + 0.12, 1.0)
            self.attributes("-alpha", next_value)
            self.fade_value = next_value
            if next_value < 1.0:
                self.after(28, self._animate_fade_in)
        except Exception:
            pass

    def _schedule_frame_advance(self) -> None:
        if self._closing or self.reduced_motion or len(self.frames) <= 1:
            return
        self.frame_index = (self.frame_index + 1) % len(self.frames)
        self.image_label.configure(image=self.frames[self.frame_index])
        self._gif_after_id = self.after(120, self._schedule_frame_advance)

    def close_now(self) -> None:
        if self._closing:
            return
        self._closing = True
        if self._gif_after_id:
            try:
                self.after_cancel(self._gif_after_id)
            except Exception:
                pass
        if self._close_after_id:
            try:
                self.after_cancel(self._close_after_id)
            except Exception:
                pass
        self.destroy()
        callback = self.on_close
        if callable(callback):
            callback()


class LoginReminderToast(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        *,
        palette: ThemePalette,
        on_dismiss,
        on_complete,
        reduced_motion: bool = False,
        compact: bool = False,
        scale: float = 1.0,
    ) -> None:
        super().__init__(master)
        self.palette = palette
        self.on_dismiss = on_dismiss
        self.on_complete = on_complete
        self.reduced_motion = bool(reduced_motion)
        self.compact = compact
        self.scale = scale
        self.fade_value = 0.0
        self._closing = False

        self.title("Login reminder")
        self.geometry("340x170")
        self.resizable(False, False)
        self.transient(master)
        self.attributes("-topmost", True)
        try:
            self.attributes("-alpha", 0.97 if self.reduced_motion else 0.0)
            if self.reduced_motion:
                self.fade_value = 0.97
        except Exception:
            self.fade_value = 1.0
        self.configure(background=palette.root_bg)
        apply_ttk_theme(self, palette, compact=self.compact, scale=self.scale)

        body = ttk.Frame(self, style="Card.TFrame", padding=18)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1)

        ttk.Label(body, text="Finish your login setup", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            body,
            text="This reminder stays quiet until three days after first run. Close it once to hide it forever.",
            style="CardBody.TLabel",
            wraplength=290,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        actions.grid_columnconfigure(0, weight=1)
        ttk.Button(actions, text="Dismiss forever", command=self._dismiss).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions, text="Mark logged in", style="Primary.TButton", command=self._complete).grid(row=1, column=0, sticky="ew", pady=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self._dismiss)
        self._position_bottom_right()
        if not self.reduced_motion:
            self._animate_fade_in()

    def _position_bottom_right(self) -> None:
        self.update_idletasks()
        width = self.winfo_width() or 340
        height = self.winfo_height() or 170
        margin = 18
        x = max(self.winfo_screenwidth() - width - margin, 0)
        y = max(self.winfo_screenheight() - height - margin - 40, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _animate_fade_in(self) -> None:
        if self._closing or self.reduced_motion:
            return
        try:
            next_value = min(self.fade_value + 0.14, 0.97)
            self.attributes("-alpha", next_value)
            self.fade_value = next_value
            if next_value < 0.97:
                self.after(28, self._animate_fade_in)
        except Exception:
            pass

    def _dismiss(self) -> None:
        if self._closing:
            return
        self._closing = True
        self.destroy()
        callback = self.on_dismiss
        if callable(callback):
            callback()

    def _complete(self) -> None:
        if self._closing:
            return
        self._closing = True
        self.destroy()
        callback = self.on_complete
        if callable(callback):
            callback()
