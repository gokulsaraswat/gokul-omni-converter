"""Headless UI smoke test: build every page, switch every theme, check the invariants.

The window is never mapped. On macOS, ``deiconify()`` and ``focus_force()`` block
whenever the process cannot come to the foreground (CI, a background shell), so
this stays withdrawn and drives Tk with ``update_idletasks()``.

    python tests/test_ui_smoke.py
"""

from __future__ import annotations

import faulthandler
import os
import sys
import threading
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.argv = ["app.py"]

# os._exit() below skips normal flushing; keep stdout unbuffered.
sys.stdout.reconfigure(line_buffering=True)

FAILURES: list[str] = []


def _watchdog(seconds: int = int(os.environ.get("UI_SMOKE_TIMEOUT", "600"))) -> None:
    time.sleep(seconds)
    print(f"\n!!! exceeded {seconds}s — dumping stacks", flush=True)
    faulthandler.dump_traceback()
    os._exit(2)


def main() -> int:
    threading.Thread(target=_watchdog, daemon=True).start()

    import app as app_module
    from ui.kit import InfoDot

    started = time.time()
    app = app_module.GOOmniProApp(skip_startup_overlays=True)
    # CTk binds <Configure> on the root and calls winfo_width() from it, which
    # blocks on an unmapped window. Detach it before any idle processing, give the
    # root a fixed geometry, then run withdrawn.
    app.unbind("<Configure>")
    app.geometry("1440x900")
    app.withdraw()
    app.update_idletasks()
    print(f"built {len(app_module.NAV_PAGES)} pages in {time.time() - started:.2f}s")

    def step(label: str, action) -> None:
        began = time.time()
        try:
            action()
            app.update_idletasks()
            print(f"  ok   {label}  ({time.time() - began:.1f}s)")
        except Exception as exc:  # noqa: BLE001 - reported, not swallowed
            FAILURES.append(f"{label}: {type(exc).__name__}: {exc}")
            print(f"  FAIL {label}: {exc}")
            traceback.print_exc()

    def expect(label: str, condition: bool, detail: str = "") -> None:
        print(f"  {'ok  ' if condition else 'FAIL'} {label}{'  ' + detail if detail else ''}")
        if not condition:
            FAILURES.append(label)

    for name, _ in app_module.NAV_PAGES:
        step(f"page {name}", lambda n=name: app._show_page(n))

    for label, setup in (
        ("theme light", lambda: app.theme_choice_var.set("light")),
        ("theme dark", lambda: app.theme_choice_var.set("dark")),
        ("high contrast on", lambda: app.high_contrast_var.set(True)),
        ("high contrast off", lambda: app.high_contrast_var.set(False)),
        ("compact on", lambda: app.compact_ui_var.set(True)),
        ("compact off", lambda: app.compact_ui_var.set(False)),
    ):
        step(label, lambda s=setup: (s(), app._apply_theme()))

    step("about refresh", app._refresh_about_profile)
    step("history refresh", app._refresh_history_views)
    step("home summary", app._refresh_home_summary)

    # Motion must be optional, never the only signal.
    step("reduced motion on", lambda: app.reduced_motion_var.set(True))
    step("page switch (reduced motion)", lambda: app._show_page("convert"))
    step("reduced motion off", lambda: app.reduced_motion_var.set(False))
    step("page switch (animated)", lambda: app._show_page("ocr"))

    # Workers drive progress with ttk's 0..100 convention; the shim must keep working.
    step(
        "progress 0..100 shim",
        lambda: (
            app.progress.__setitem__("value", 50),
            app.pdf_tool_progress.__setitem__("value", 100),
            app.ocr_progress.__setitem__("value", 0),
        ),
    )
    expect("progress reset is instant", app.ocr_progress.get() == 0)

    def count_info_dots(widget, total: int = 0) -> int:
        if isinstance(widget, InfoDot):
            total += 1
        for child in widget.winfo_children():
            total = count_info_dots(child, total)
        return total

    dots = count_info_dots(app)
    expect("ⓘ hover hints present", dots >= 15, f"({dots})")
    expect("about image hint is a hover hint", isinstance(app.about_image_hint_label, InfoDot))
    expect("no dev-credit line on About", not hasattr(app, "about_credit_label"))
    expect("sidebar status still shown", app.sidebar_status.winfo_exists())

    # Regression guard: a modal grab freezes the root event loop on macOS.
    expect("nothing grabs input", app.tk.call("grab", "current") in ("", None))

    # Pages must actually scroll: the settings page is taller than any window.
    app._show_page("settings")
    app.update_idletasks()
    scroller = app.page_scrollers["settings"]
    scroller._sync()
    region = str(scroller.canvas.cget("scrollregion") or "")
    height = float(region.split()[3]) if len(region.split()) == 4 else 0
    expect("settings page is scrollable", height > 100, f"(scrollregion height {height:.0f})")

    step("save state", app.state_store.save)

    failed = bool(FAILURES)
    if failed:
        print(f"\n{len(FAILURES)} FAILED:")
        for failure in FAILURES:
            print(f" - {failure}")
    else:
        print("\nAll UI smoke checks passed.")
    sys.stdout.flush()
    os._exit(1 if failed else 0)


if __name__ == "__main__":
    main()
