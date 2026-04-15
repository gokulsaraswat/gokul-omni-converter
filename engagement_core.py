from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

LOGIN_POPUP_DELAY_DAYS = 3
LOGIN_POPUP_COOLDOWN_HOURS = 24
SPLASH_AUTO_CLOSE_MS = 2600


def iso_now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def parse_datetime(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def ensure_install_date(state: dict[str, Any], *, now: datetime | None = None) -> str:
    current = parse_datetime(state.get("install_date"))
    if current is not None:
        return current.isoformat(timespec="seconds")
    stamp = (now or datetime.now()).isoformat(timespec="seconds")
    state["install_date"] = stamp
    return stamp


def days_since_install(state: dict[str, Any], *, now: datetime | None = None) -> int:
    current = now or datetime.now()
    installed_at = parse_datetime(state.get("install_date"))
    if installed_at is None:
        ensure_install_date(state, now=current)
        installed_at = parse_datetime(state.get("install_date")) or current
    delta = current.date() - installed_at.date()
    return max(delta.days, 0)


def should_show_login_popup(state: dict[str, Any], *, now: datetime | None = None) -> bool:
    if not bool(state.get("login_popup_enabled", True)):
        return False
    if bool(state.get("login_popup_dismissed")):
        return False
    if bool(state.get("login_popup_completed")):
        return False

    current = now or datetime.now()
    if days_since_install(state, now=current) < LOGIN_POPUP_DELAY_DAYS:
        return False

    last_shown = parse_datetime(state.get("login_popup_last_shown"))
    if last_shown and current - last_shown < timedelta(hours=LOGIN_POPUP_COOLDOWN_HOURS):
        return False

    return True


def should_show_first_launch_splash(state: dict[str, Any]) -> bool:
    if not bool(state.get("splash_enabled", True)):
        return False
    return not bool(state.get("splash_seen", False))


def summarize_login_popup_state(state: dict[str, Any], *, now: datetime | None = None) -> str:
    current = now or datetime.now()
    install_age = days_since_install(state, now=current)
    install_text = f"Installed {install_age} day{'s' if install_age != 1 else ''} ago"
    if not bool(state.get("login_popup_enabled", True)):
        return f"{install_text}. Reminder is disabled."
    if bool(state.get("login_popup_completed")):
        return f"{install_text}. Login reminder completed and will not return."
    if bool(state.get("login_popup_dismissed")):
        return f"{install_text}. Login reminder dismissed permanently."
    if should_show_login_popup(dict(state), now=current):
        return f"{install_text}. Reminder is eligible to show."
    remaining = max(LOGIN_POPUP_DELAY_DAYS - install_age, 0)
    if remaining > 0:
        return f"{install_text}. Reminder unlocks in {remaining} day{'s' if remaining != 1 else ''}."
    last_shown = parse_datetime(state.get("login_popup_last_shown"))
    if last_shown:
        return f"{install_text}. Reminder shown recently and is cooling down."
    return f"{install_text}. Reminder is waiting for the next eligible launch."
