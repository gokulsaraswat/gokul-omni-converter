from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ABOUT_PROFILE_PATH = Path(__file__).with_name("about_profile.json")

DEFAULT_PROFILE: dict[str, Any] = {
    "name": "Gokul Saraswat",
    "title": "Software Engineer",
    "subtitle": "Creator of Gokul Omni Convert Lite",
    "company": "Oracle Corporation",
    "project": "Gokul Omni Convert Lite",
    "email": "gokul.saraswat@oracle.com",
    "handle": "@gokul.saraswat",
    "bio": (
        "Local-first desktop conversion toolkit for PDFs, documents, spreadsheets, presentations, "
        "images, HTML, Markdown, OCR workflows, and practical batch automation."
    ),
    "image_path": "assets/gokul_profile_placeholder.png",
    "feedback_url": "mailto:gokul.saraswat@oracle.com?subject=Gokul%20Omni%20Convert%20Lite%20Feedback",
    "contribute_url": "https://github.com/gokul-saraswat/gokul-omni-convert-lite/issues",
    "links": [
        {"label": "Email", "url": "mailto:gokul.saraswat@oracle.com?subject=Gokul%20Omni%20Convert%20Lite"},
        {"label": "LinkedIn", "url": ""},
        {"label": "GitHub", "url": ""},
        {"label": "X", "url": ""},
        {"label": "Website", "url": ""},
    ],
}


def _normalize_links(value: Any) -> list[dict[str, str]]:
    cleaned_links: list[dict[str, str]] = []
    if not isinstance(value, list):
        return cleaned_links
    for item in value:
        if not isinstance(item, dict):
            continue
        cleaned_links.append(
            {
                "label": str(item.get("label", "Link")).strip() or "Link",
                "url": str(item.get("url", "")).strip(),
            }
        )
    return cleaned_links


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
        migrated = dict(data)

        if not str(migrated.get("company", "")).strip():
            migrated["company"] = DEFAULT_PROFILE["company"]
        if not str(migrated.get("project", "")).strip():
            subtitle = str(migrated.get("subtitle", "")).strip()
            migrated["project"] = subtitle or DEFAULT_PROFILE["project"]

        if not str(migrated.get("feedback_url", "")).strip():
            email = str(migrated.get("email", "")).strip() or DEFAULT_PROFILE["email"]
            migrated["feedback_url"] = f"mailto:{email}?subject=Gokul%20Omni%20Convert%20Lite%20Feedback"
        if "contribute_url" not in migrated:
            migrated["contribute_url"] = DEFAULT_PROFILE["contribute_url"]

        profile.update({k: v for k, v in migrated.items() if k != "links"})
        links = _normalize_links(migrated.get("links"))
        if links:
            profile["links"] = links

    return profile


def resolve_profile_image(profile: dict[str, Any], base_dir: Path | None = None) -> Path:
    root = base_dir or ABOUT_PROFILE_PATH.parent
    image_value = str(profile.get("image_path", DEFAULT_PROFILE["image_path"]))
    image_path = Path(image_value).expanduser()
    if not image_path.is_absolute():
        image_path = root / image_path
    return image_path
