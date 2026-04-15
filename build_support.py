from __future__ import annotations

import json
import platform
import sys
import zipfile
from datetime import datetime
from html import escape
from importlib import metadata
from pathlib import Path
from typing import Any, Iterable


PACKAGE_NAMES = [
    "Pillow",
    "PyMuPDF",
    "pdfplumber",
    "python-docx",
    "openpyxl",
    "pypdf",
    "reportlab",
    "python-pptx",
    "xlrd",
    "pytesseract",
]


def collect_package_versions(package_names: Iterable[str] = PACKAGE_NAMES) -> dict[str, str]:
    versions: dict[str, str] = {}
    for name in package_names:
        try:
            versions[name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            versions[name] = "not installed"
        except Exception:
            versions[name] = "unknown"
    return versions


def collect_installer_assets(installer_dir: Path) -> list[str]:
    root = Path(installer_dir)
    if not root.exists():
        return []
    assets: list[str] = []
    for path in sorted(root.rglob("*")):
        if path.is_file():
            assets.append(str(path.relative_to(root)))
    return assets


def export_json(path: Path, payload: dict[str, Any]) -> Path:
    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return destination


def export_text_file(path: Path, content: str) -> Path:
    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(str(content), encoding="utf-8")
    return destination


def export_diagnostics_report(
    path: Path,
    *,
    app_name: str,
    app_version: str,
    state_path: Path,
    about_profile_path: Path,
    notes_path: Path,
    installer_dir: Path,
    output_dir: Path,
    selected_files: list[str],
    last_outputs: list[str],
    dependency_status: dict[str, Any],
    smtp_summary: dict[str, Any],
    extra: dict[str, Any] | None = None,
) -> Path:
    payload: dict[str, Any] = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "app": {
            "name": app_name,
            "version": app_version,
        },
        "system": {
            "platform": platform.platform(),
            "python": sys.version,
            "python_executable": sys.executable,
            "cwd": str(Path.cwd()),
        },
        "paths": {
            "state_path": str(state_path),
            "about_profile_path": str(about_profile_path),
            "notes_path": str(notes_path),
            "installer_dir": str(installer_dir),
            "output_dir": str(output_dir),
        },
        "packages": collect_package_versions(),
        "installer_assets": collect_installer_assets(installer_dir),
        "dependencies": dependency_status,
        "smtp": smtp_summary,
        "current_context": {
            "selected_files": selected_files,
            "last_outputs": last_outputs,
        },
    }
    if extra:
        payload["extra"] = extra
    return export_json(path, payload)


def export_state_snapshot(path: Path, state: dict[str, Any]) -> Path:
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "state": state,
    }
    return export_json(path, payload)


def import_state_snapshot(path: Path) -> dict[str, Any]:
    source = Path(path).expanduser()
    data = json.loads(source.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Invalid state snapshot format.")
    payload = data.get("state", data)
    if not isinstance(payload, dict):
        raise ValueError("Invalid state snapshot payload.")
    return payload


def _status_class(value: object) -> str:
    text = str(value or "").strip().lower()
    if text in {"success", "completed", "ok"}:
        return "good"
    if text in {"failed", "error"}:
        return "bad"
    return "neutral"


def _value_html(value: object) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return escape(str(value))


def render_activity_report_html(
    path: Path,
    *,
    app_name: str,
    app_version: str,
    recent_jobs: list[dict[str, Any]],
    recent_outputs: list[str],
    failed_jobs: list[dict[str, Any]],
    dependency_status: dict[str, Any] | None = None,
    notes: str = "",
) -> Path:
    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)

    dependency_status = dict(dependency_status or {})
    jobs = list(recent_jobs or [])
    outputs = [str(item) for item in (recent_outputs or [])]
    failures = list(failed_jobs or [])
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    counts = {
        "jobs": len(jobs),
        "success": sum(1 for job in jobs if str(job.get("status", "")).strip().lower() == "success"),
        "failed": sum(1 for job in jobs if str(job.get("status", "")).strip().lower() != "success"),
        "outputs": len(outputs),
        "queued_failures": len(failures),
    }

    rows: list[str] = []
    for job in jobs[:40]:
        status = str(job.get("status", "")).strip() or "unknown"
        mode = str(job.get("mode", "")).strip() or "—"
        timestamp = str(job.get("timestamp", "")).strip() or "—"
        file_count = job.get("file_count", 0)
        output_count = job.get("output_count", 0)
        output_dir = str(job.get("output_dir", "")).strip() or "—"
        rows.append(
            "<tr>"
            f"<td>{escape(timestamp)}</td>"
            f"<td><span class='badge {escape(_status_class(status))}'>{escape(status)}</span></td>"
            f"<td>{escape(mode)}</td>"
            f"<td>{escape(str(file_count))}</td>"
            f"<td>{escape(str(output_count))}</td>"
            f"<td>{escape(output_dir)}</td>"
            "</tr>"
        )

    output_items = "".join(f"<li>{escape(item)}</li>" for item in outputs[:24]) or "<li>No outputs recorded yet.</li>"
    failure_items = "".join(
        "<li>"
        f"<strong>{escape(str(item.get('mode', 'Run')))}</strong> — "
        f"{escape(str(item.get('error', 'No error details')))}"
        "</li>"
        for item in failures[:16]
    ) or "<li>No failed jobs queued.</li>"

    dependency_items = "".join(
        "<li>"
        f"<strong>{escape(str(name))}</strong>: {escape(str(value))}"
        "</li>"
        for name, value in dependency_status.items()
    ) or "<li>Dependency summary unavailable.</li>"

    notes_html = f"<div class='notes'>{escape(notes)}</div>" if str(notes).strip() else ""

    html_text = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>{escape(app_name)} Activity Report</title>
<style>
:root {{
  --bg: #0f172a;
  --panel: #111827;
  --panel2: #1f2937;
  --text: #f8fafc;
  --muted: #cbd5e1;
  --line: #334155;
  --accent: #3b82f6;
  --good: #22c55e;
  --bad: #ef4444;
  --neutral: #64748b;
}}
body {{
  margin: 0;
  font-family: "Segoe UI", Arial, sans-serif;
  background: linear-gradient(135deg, #0b1220, var(--bg));
  color: var(--text);
}}
.container {{
  max-width: 1180px;
  margin: 0 auto;
  padding: 32px 24px 56px;
}}
.header {{
  background: rgba(17,24,39,0.9);
  border: 1px solid var(--line);
  border-radius: 22px;
  padding: 24px 26px;
  box-shadow: 0 18px 40px rgba(0,0,0,0.22);
}}
h1 {{
  margin: 0 0 6px;
  font-size: 28px;
}}
.subtitle {{
  color: var(--muted);
  margin: 0;
  line-height: 1.5;
}}
.grid {{
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 14px;
  margin-top: 18px;
}}
.card {{
  background: rgba(17,24,39,0.92);
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 18px;
}}
.metric {{
  font-size: 28px;
  font-weight: 700;
  margin: 8px 0 0;
}}
.label {{
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.08em;
  font-size: 12px;
}}
.section {{
  margin-top: 18px;
  background: rgba(17,24,39,0.92);
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 18px;
}}
.section h2 {{
  margin: 0 0 12px;
  font-size: 18px;
}}
table {{
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}}
th, td {{
  text-align: left;
  padding: 10px 12px;
  border-top: 1px solid rgba(148,163,184,0.16);
  vertical-align: top;
}}
th {{
  color: var(--muted);
  font-weight: 600;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.05em;
}}
tr:hover td {{
  background: rgba(59,130,246,0.08);
}}
.badge {{
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.04em;
}}
.badge.good {{ background: rgba(34,197,94,0.18); color: #86efac; }}
.badge.bad {{ background: rgba(239,68,68,0.18); color: #fca5a5; }}
.badge.neutral {{ background: rgba(100,116,139,0.22); color: #cbd5e1; }}
.two-col {{
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 18px;
}}
ul {{
  margin: 0;
  padding-left: 18px;
  line-height: 1.6;
}}
.notes {{
  margin-top: 12px;
  padding: 14px 16px;
  border-radius: 16px;
  background: rgba(59,130,246,0.10);
  border: 1px solid rgba(59,130,246,0.28);
  color: var(--muted);
  white-space: pre-wrap;
}}
.footer {{
  margin-top: 20px;
  color: var(--muted);
  font-size: 13px;
}}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>{escape(app_name)} Activity Report</h1>
      <p class="subtitle">Generated {escape(created_at)} · {escape(app_version)}. This report summarizes recent jobs, outputs, queued retries, and dependency signals from the local workspace.</p>
      {notes_html}
      <div class="grid">
        <div class="card"><div class="label">Recent jobs</div><div class="metric">{counts['jobs']}</div></div>
        <div class="card"><div class="label">Successful</div><div class="metric">{counts['success']}</div></div>
        <div class="card"><div class="label">Failed / other</div><div class="metric">{counts['failed']}</div></div>
        <div class="card"><div class="label">Recent outputs</div><div class="metric">{counts['outputs']}</div></div>
        <div class="card"><div class="label">Queued retries</div><div class="metric">{counts['queued_failures']}</div></div>
      </div>
    </div>

    <div class="section">
      <h2>Recent jobs</h2>
      <table>
        <thead>
          <tr>
            <th>Time</th>
            <th>Status</th>
            <th>Mode</th>
            <th>Files</th>
            <th>Outputs</th>
            <th>Output folder</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows) or '<tr><td colspan="6">No recent jobs are stored yet.</td></tr>'}
        </tbody>
      </table>
    </div>

    <div class="two-col">
      <div class="section">
        <h2>Recent outputs</h2>
        <ul>{output_items}</ul>
      </div>
      <div class="section">
        <h2>Failed jobs ready for retry</h2>
        <ul>{failure_items}</ul>
      </div>
    </div>

    <div class="section">
      <h2>Dependency summary</h2>
      <ul>{dependency_items}</ul>
    </div>

    <div class="footer">Keep Pure Python as the default engine and use LibreOffice only when you explicitly configure and select it.</div>
  </div>
</body>
</html>
"""
    destination.write_text(html_text, encoding="utf-8")
    return destination


def export_support_bundle(
    destination_zip: Path,
    *,
    diagnostics_report: Path | None = None,
    state_snapshot: Path | None = None,
    activity_report: Path | None = None,
    logs_path: Path | None = None,
    notes_path: Path | None = None,
    about_profile_path: Path | None = None,
    installer_dir: Path | None = None,
    extra_files: Iterable[Path] | None = None,
) -> Path:
    destination = Path(destination_zip).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(destination, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "manifest.json",
            json.dumps(
                {
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "bundle_type": "support",
                },
                indent=2,
            ),
        )

        def add_file(item: Path | None, arcname: str | None = None) -> None:
            if item is None:
                return
            file_path = Path(item).expanduser()
            if not file_path.exists() or not file_path.is_file():
                return
            archive.write(file_path, arcname=arcname or file_path.name)

        add_file(diagnostics_report, "reports/diagnostics.json")
        add_file(state_snapshot, "reports/state_snapshot.json")
        add_file(activity_report, "reports/activity_report.html")
        add_file(logs_path, "reports/app_logs.txt")
        add_file(notes_path, "profile/footer_notes.md")
        add_file(about_profile_path, "profile/about_profile.json")

        if installer_dir and Path(installer_dir).exists():
            root = Path(installer_dir)
            for path in sorted(root.rglob("*")):
                if path.is_file():
                    archive.write(path, arcname=str(Path("installer") / path.relative_to(root)))

        for item in extra_files or []:
            file_path = Path(item).expanduser()
            if not file_path.exists() or not file_path.is_file():
                continue
            if file_path.name in archive.namelist():
                archive.write(file_path, arcname=str(Path("extras") / file_path.name))
            else:
                archive.write(file_path, arcname=str(Path("extras") / file_path.name))

    return destination
