from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ABOUT_PROFILE_PATH = Path(__file__).with_name("about_profile.json")

#: Development credit shown on the About page. Deliberately separate from the
#: profile fields so it is never mistaken for the author's own details.
DEVELOPMENT_CREDIT = "Built with development assistance from Claude (Anthropic)."

#: The only profile fields GO Omni Pro carries: who made it, how to reach them,
#: and where to find them. There is intentionally no company, job title, handle,
#: or bio -- those were placeholder fields and have been removed from the model.
PROFILE_FIELDS = ("name", "email", "image_path", "image_url")

DEFAULT_PROFILE: dict[str, Any] = {
    "name": "Gokul Saraswat",
    "email": "gokulsaraswat07@gmail.com",
    "image_path": "assets/logo/go_omni_pro_256.png",
    "image_url": "",
    "links": [
        {"label": "Website", "url": "https://gokulsaraswat.com"},
        {"label": "GitHub", "url": "https://github.com/gokulsaraswat"},
        {"label": "LinkedIn", "url": "https://linkedin.com/in/gokulsaraswat"},
    ],
}

#: Fields carried by 2.3.0 and earlier that no longer exist. Dropped on load so a
#: stale about_profile.json can never resurrect them in the UI.
_RETIRED_FIELDS = (
    "title",
    "subtitle",
    "company",
    "project",
    "handle",
    "bio",
    "feedback_url",
    "contribute_url",
)


def _normalize_links(value: Any) -> list[dict[str, str]]:
    cleaned_links: list[dict[str, str]] = []
    if not isinstance(value, list):
        return cleaned_links
    for item in value:
        if not isinstance(item, dict):
            continue
        url = str(item.get("url", "")).strip()
        if not url:
            continue
        cleaned_links.append(
            {
                "label": str(item.get("label", "Link")).strip() or "Link",
                "url": url,
            }
        )
    return cleaned_links


def load_about_profile(path: Path | None = None) -> dict[str, Any]:
    """Read the profile, dropping any retired fields left over from older versions."""
    profile_path = path or ABOUT_PROFILE_PATH
    profile_path.parent.mkdir(parents=True, exist_ok=True)
    if not profile_path.exists():
        profile_path.write_text(json.dumps(DEFAULT_PROFILE, indent=2), encoding="utf-8")
        return dict(DEFAULT_PROFILE)

    try:
        data = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        profile_path.write_text(json.dumps(DEFAULT_PROFILE, indent=2), encoding="utf-8")
        return dict(DEFAULT_PROFILE)

    profile = dict(DEFAULT_PROFILE)
    if isinstance(data, dict):
        for field in PROFILE_FIELDS:
            value = str(data.get(field, "")).strip()
            if value:
                profile[field] = value
        links = _normalize_links(data.get("links"))
        if links:
            profile["links"] = links

    for field in _RETIRED_FIELDS:
        profile.pop(field, None)
    return profile


def save_about_profile(profile: dict[str, Any], path: Path | None = None) -> Path:
    """Write only the supported fields back out, so retired keys never reappear."""
    profile_path = path or ABOUT_PROFILE_PATH
    payload: dict[str, Any] = {
        field: str(profile.get(field, DEFAULT_PROFILE.get(field, ""))).strip()
        for field in PROFILE_FIELDS
    }
    payload["links"] = _normalize_links(profile.get("links")) or list(DEFAULT_PROFILE["links"])
    profile_path.parent.mkdir(parents=True, exist_ok=True)
    profile_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return profile_path


def resolve_profile_image(profile: dict[str, Any], base_dir: Path | None = None) -> Path:
    root = base_dir or ABOUT_PROFILE_PATH.parent
    image_value = str(profile.get("image_path") or DEFAULT_PROFILE["image_path"])
    image_path = Path(image_value).expanduser()
    if not image_path.is_absolute():
        image_path = root / image_path
    return image_path
