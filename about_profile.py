from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ABOUT_PROFILE_PATH = Path(__file__).with_name("about_profile.json")

DEFAULT_PROFILE: dict[str, Any] = {
    "name": "Gokul Saraswat",
    "title": "Software Engineer",
    "subtitle": "Creator of Gokul Omni Convert Lite",
    "email": "gokul.saraswat@oracle.com",
    "handle": "@gokul.saraswat",
    "bio": (
        "Local-first desktop conversion toolkit for PDFs, documents, spreadsheets, presentations, "
        "images, HTML, Markdown, and practical batch workflows."
    ),
    "image_path": "assets/gokul_profile_placeholder.png",
    "links": [
        {"label": "Email", "url": "mailto:gokul.saraswat@oracle.com"},
        {"label": "LinkedIn", "url": ""},
        {"label": "GitHub", "url": ""},
        {"label": "X", "url": ""},
        {"label": "Website", "url": ""},
    ],
}


def load_about_profile(path: Path | None = None) -> dict[str, Any]:
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
        profile.update({k: v for k, v in data.items() if k != "links"})
        if isinstance(data.get("links"), list):
            cleaned_links = []
            for item in data["links"]:
                if not isinstance(item, dict):
                    continue
                cleaned_links.append(
                    {
                        "label": str(item.get("label", "Link")).strip() or "Link",
                        "url": str(item.get("url", "")).strip(),
                    }
                )
            if cleaned_links:
                profile["links"] = cleaned_links
    return profile


def resolve_profile_image(profile: dict[str, Any], base_dir: Path | None = None) -> Path:
    root = base_dir or ABOUT_PROFILE_PATH.parent
    image_value = str(profile.get("image_path", DEFAULT_PROFILE["image_path"]))
    image_path = Path(image_value).expanduser()
    if not image_path.is_absolute():
        image_path = root / image_path
    return image_path
